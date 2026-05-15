import argparse
import json
import math
import os
import re
from pathlib import Path

os.environ["PYTORCH_CUDA_ALLOC_CONF"] = "expandable_segments:True"
os.environ["WANDB_MODE"] = "disabled"
os.environ["WANDB_PROJECT"] = "disabled"
os.environ["TOKENIZERS_PARALLELISM"] = "false"

from PIL import Image

from compact_schema import compact_json_len, compact_schema_reward, compact_to_full
from grader import grade_json, schema_reward


DATA_DIR = Path("output")
OUTPUT_DIR = Path("model_output/slide_json_grpo")
MODEL_NAME = "Qwen/Qwen2.5-VL-3B-Instruct"

MAX_STEPS = 50
MAX_PROMPT_LENGTH = 2048
MAX_COMPLETION_LENGTH = 1024
PER_DEVICE_BATCH_SIZE = 2
GRADIENT_ACCUMULATION_STEPS = 4
NUM_GENERATIONS = 2
LEARNING_RATE = 1e-5
LORA_R = 16
LORA_ALPHA = 32
WARMUP_RATIO = 0.03
WEIGHT_DECAY = 0.01
MAX_GRAD_NORM = 0.3
SAVE_STEPS = 25
SAVE_TOTAL_LIMIT = 3
IMAGE_MIN_PIXELS = 128 * 28 * 28
IMAGE_MAX_PIXELS = 384 * 384

USER_PROMPT = """Infer the PowerPoint-like scene graph from this slide image.

Return only compact JSON in this exact shape:
<json>{"v":1,"s":[13.333,7.5],"bg":["solid",[255,255,255]],"o":[object_rows]}</json>

Each object row is [type,x,y,w,h,props]. Use inches.
Types: tx text, sh shape, tb table, im image, cn connector, ch chart, ff freeform, sv svg, si svg_image.
Common props:
tx {"t":"text","fs":24,"ff":"Arial","c":[0,0,0],"b":1}
sh {"sh":"rect","f":[230,230,230],"l":[40,40,40],"lw":1}
tb {"r":3,"c":3,"cells":[["a","b"]],"hd":[0,0,0],"bd":[255,255,255],"tc":[0,0,0],"fs":10}
im {}
ch {"ct":"column","cat":["A","B"],"ser":[{"name":"S","values":[1,2]}]}
Do not include markdown or explanatory text."""


def require_gpu():
    import torch

    if not torch.cuda.is_available():
        raise RuntimeError("No CUDA GPU detected.")
    props = torch.cuda.get_device_properties(torch.cuda.current_device())
    return {
        "name": props.name,
        "total_gb": round(props.total_memory / 1024**3, 2),
        "bf16": bool(torch.cuda.is_bf16_supported()),
    }


def paired_examples(data_dir):
    examples = []
    for json_path in sorted(Path(data_dir).glob("slide_*.json")):
        image_path = json_path.with_suffix(".jpg")
        if image_path.exists():
            try:
                data = json.loads(json_path.read_text(encoding="utf-8"))
                curriculum = data.get("curriculum", {}) if isinstance(data, dict) else {}
                difficulty = int(curriculum.get("level") or len(data.get("objects", [])))
            except Exception:
                difficulty = 1
            examples.append({"target_json": str(json_path), "image_path": str(image_path), "difficulty": difficulty})
    if not examples:
        raise FileNotFoundError(f"No slide_*.json / slide_*.jpg pairs found in {data_dir}")
    return sorted(examples, key=lambda item: (item["difficulty"], item["target_json"]))


def examples_to_dataset(examples):
    from datasets import Dataset

    rows = []
    for example in examples:
        rows.append(
            {
                "prompt": [
                    {
                        "role": "user",
                        "content": [
                            {"type": "image"},
                            {"type": "text", "text": USER_PROMPT},
                        ],
                    }
                ],
                "image": Image.open(example["image_path"]).convert("RGB"),
                "target_json": example["target_json"],
                "difficulty": example["difficulty"],
            }
        )
    return Dataset.from_list(rows)


def load_dataset(data_dir, limit=None):
    return examples_to_dataset(paired_examples(data_dir)[:limit])


def load_model_and_processor(model_name):
    import torch
    from transformers import AutoProcessor, Qwen2_5_VLForConditionalGeneration

    compute_dtype = torch.bfloat16 if torch.cuda.is_bf16_supported() else torch.float16
    model = Qwen2_5_VLForConditionalGeneration.from_pretrained(
        model_name,
        device_map="auto",
        dtype=compute_dtype,
        attn_implementation="eager",
    )
    model.config.use_cache = False
    model.config.torch_dtype = compute_dtype
    model.gradient_checkpointing_enable()
    processor = AutoProcessor.from_pretrained(
        model_name,
        use_fast=True,
        min_pixels=IMAGE_MIN_PIXELS,
        max_pixels=IMAGE_MAX_PIXELS,
    )
    return model, processor


def completion_to_text(completion):
    if isinstance(completion, str):
        return completion
    if isinstance(completion, dict):
        return completion_to_text(completion.get("content", ""))
    if isinstance(completion, list):
        parts = []
        for item in completion:
            if isinstance(item, dict):
                parts.append(str(item.get("text", completion_to_text(item.get("content", "")))))
            else:
                parts.append(str(item))
        return "".join(parts)
    return str(completion)


def balanced_json_slice(text):
    start = text.find("{")
    if start == -1:
        raise ValueError("No JSON object found")
    depth = 0
    in_string = False
    escape = False
    for index in range(start, len(text)):
        char = text[index]
        if in_string:
            if escape:
                escape = False
            elif char == "\\":
                escape = True
            elif char == '"':
                in_string = False
        else:
            if char == '"':
                in_string = True
            elif char == "{":
                depth += 1
            elif char == "}":
                depth -= 1
                if depth == 0:
                    return text[start : index + 1]
    raise ValueError("Unclosed JSON object")


def extract_json(text):
    match = re.search(r"<json>\s*(.*?)\s*</json>", text, flags=re.DOTALL | re.IGNORECASE)
    return json.loads(match.group(1) if match else balanced_json_slice(text))


def completion_format_reward(text):
    has_json_tag = bool(re.search(r"<json>.*</json>", text, flags=re.DOTALL | re.IGNORECASE))
    clean_envelope = bool(re.fullmatch(r"\s*<json>\s*.*?\s*</json>\s*", text, flags=re.DOTALL | re.IGNORECASE))
    no_markdown = "```" not in text
    brace_present = "{" in text and "}" in text
    return 0.25 * has_json_tag + 0.25 * clean_envelope + 0.15 * no_markdown + 0.35 * brace_present


def completion_bloat_penalty(text, target_len):
    allowed = max(1500, target_len * 1.7)
    if len(text) <= allowed:
        return 0.0
    return min(0.70, (len(text) - allowed) / max(1500, target_len) * 0.35)


def slide_json_reward_func(completions, target_json=None, **kwargs):
    target_json = target_json or kwargs.get("target_json") or []
    rewards = []
    for completion, target_path in zip(completions, target_json):
        text = completion_to_text(completion)
        format_score = completion_format_reward(text)
        try:
            raw_prediction = extract_json(text)
            prediction = compact_to_full(raw_prediction)
            grade = grade_json(target_path, prediction)
            if isinstance(raw_prediction, dict) and "o" in raw_prediction:
                structure_score = (compact_schema_reward(raw_prediction) + 1.0) / 2.0
            else:
                structure_score = (schema_reward(prediction) + 1.0) / 2.0
            task_score = (float(grade["reward"]) + 1.0) / 2.0
            target_len = compact_json_len(target_path)
            length_score = math.exp(-abs(len(text) - target_len) / max(750, target_len * 0.9))
            unit_reward = 0.06 * format_score + 0.10 * structure_score + 0.09 * length_score + 0.75 * task_score
            unit_reward -= completion_bloat_penalty(text, target_len)
            rewards.append(unit_reward * 2.0 - 1.0)
        except Exception:
            rewards.append(-0.90 + 0.15 * format_score)
    return rewards


def make_grpo_config(config_cls, kwargs):
    kwargs = dict(kwargs)
    while True:
        try:
            return config_cls(**kwargs)
        except TypeError as exc:
            match = re.search(r"unexpected keyword argument '([^']+)'", str(exc))
            if not match:
                raise
            removed = match.group(1)
            kwargs.pop(removed, None)
            print(f"GRPOConfig does not support {removed}; continuing without it.")


def grpo_kwargs(args, gpu, max_steps, output_dir):
    return {
        "output_dir": str(output_dir),
        "per_device_train_batch_size": PER_DEVICE_BATCH_SIZE,
        "gradient_accumulation_steps": GRADIENT_ACCUMULATION_STEPS,
        "learning_rate": LEARNING_RATE,
        "max_steps": max_steps,
        "logging_steps": 1,
        "save_steps": SAVE_STEPS,
        "save_total_limit": SAVE_TOTAL_LIMIT,
        "remove_unused_columns": False,
        "max_prompt_length": MAX_PROMPT_LENGTH,
        "max_completion_length": args.max_completion_length,
        "num_generations": NUM_GENERATIONS,
        "report_to": "none",
        "bf16": gpu["bf16"],
        "fp16": not gpu["bf16"],
        "optim": "adamw_torch",
        "warmup_ratio": WARMUP_RATIO,
        "lr_scheduler_type": "cosine",
        "weight_decay": WEIGHT_DECAY,
        "max_grad_norm": MAX_GRAD_NORM,
        "temperature": 0.8,
        "top_p": 0.95,
        "generation_kwargs": {
            "remove_invalid_values": True,
            "renormalize_logits": True,
        },
        "beta": 0.0,
    }


def make_trainer(model, processor, dataset, args, gpu, max_steps, output_dir, peft_config=None):
    from trl import GRPOConfig, GRPOTrainer

    grpo_config = make_grpo_config(GRPOConfig, grpo_kwargs(args, gpu, max_steps, output_dir))
    return GRPOTrainer(
        model=model,
        processing_class=processor,
        reward_funcs=[slide_json_reward_func],
        args=grpo_config,
        train_dataset=dataset,
        peft_config=peft_config,
    )


def generate_validation_completion(model, processor, image_path, max_new_tokens):
    import torch

    image = Image.open(image_path).convert("RGB")
    messages = [
        {
            "role": "user",
            "content": [
                {"type": "image", "image": image},
                {"type": "text", "text": USER_PROMPT},
            ],
        }
    ]
    prompt = processor.apply_chat_template(messages, tokenize=False, add_generation_prompt=True)
    inputs = processor(text=[prompt], images=[image], return_tensors="pt", padding=True, truncation=False)
    device = next(model.parameters()).device
    inputs = {key: value.to(device) if hasattr(value, "to") else value for key, value in inputs.items()}
    with torch.inference_mode():
        output_ids = model.generate(
            **inputs,
            max_new_tokens=max_new_tokens,
            do_sample=False,
            remove_invalid_values=True,
            renormalize_logits=True,
        )
    generated = output_ids[:, inputs["input_ids"].shape[1] :]
    return processor.batch_decode(generated, skip_special_tokens=True)[0]


def evaluate_level(model, processor, examples, args):
    sample = examples[: max(1, args.curriculum_eval_samples)]
    completions = [generate_validation_completion(model, processor, item["image_path"], args.max_completion_length) for item in sample]
    rewards = slide_json_reward_func(completions, target_json=[item["target_json"] for item in sample])
    mean_reward = sum(rewards) / max(1, len(rewards))
    return {"reward": mean_reward, "accuracy": (mean_reward + 1.0) / 2.0, "samples": len(sample)}


def train_curriculum(args, gpu, examples, model, processor, peft_config):
    by_level = {}
    for example in examples:
        by_level.setdefault(example["difficulty"], []).append(example)

    levels = sorted(by_level)
    remaining = args.max_steps
    first_stage = True
    current_index = 0
    repeats = 0

    while remaining > 0 and current_index < len(levels):
        level = levels[current_index]
        stage_steps = min(args.curriculum_stage_steps, remaining)
        stage_examples = by_level[level]
        print(f"curriculum level {level}: {len(stage_examples)} examples, training {stage_steps} steps")

        dataset = examples_to_dataset(stage_examples)
        stage_output = args.output_dir / f"_stage_level_{level:03}"
        trainer = make_trainer(
            model,
            processor,
            dataset,
            args,
            gpu,
            stage_steps,
            stage_output,
            peft_config=peft_config if first_stage else None,
        )
        trainer.train()
        model = trainer.model
        first_stage = False
        remaining -= stage_steps

        metrics = evaluate_level(model, processor, stage_examples, args)
        print(
            f"curriculum level {level}: eval reward={metrics['reward']:.4f} "
            f"accuracy={metrics['accuracy']:.4f} samples={metrics['samples']}"
        )
        if metrics["accuracy"] >= args.curriculum_accuracy_threshold:
            current_index += 1
            repeats = 0
        else:
            repeats += 1
            if repeats >= args.curriculum_max_repeats:
                print(f"curriculum level {level}: advancing after {repeats} repeats")
                current_index += 1
                repeats = 0

    model.save_pretrained(args.output_dir)
    processor.save_pretrained(args.output_dir)


def train(args):
    import torch
    from peft import LoraConfig

    gpu = require_gpu()
    print(f"GPU: {gpu['name']} ({gpu['total_gb']} GB), bf16={gpu['bf16']}")

    examples = paired_examples(args.data_dir)[: args.limit]
    model, processor = load_model_and_processor(args.model_name)

    peft_config = LoraConfig(
        r=LORA_R,
        lora_alpha=LORA_ALPHA,
        lora_dropout=0.0,
        bias="none",
        task_type="CAUSAL_LM",
        target_modules=["q_proj", "k_proj", "v_proj", "o_proj"],
    )

    if args.curriculum:
        train_curriculum(args, gpu, examples, model, processor, peft_config)
        return

    dataset = examples_to_dataset(examples)
    trainer = make_trainer(model, processor, dataset, args, gpu, args.max_steps, args.output_dir, peft_config=peft_config)
    trainer.train()
    trainer.save_model(args.output_dir)
    processor.save_pretrained(args.output_dir)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--data-dir", type=Path, default=DATA_DIR)
    parser.add_argument("--output-dir", type=Path, default=OUTPUT_DIR)
    parser.add_argument("--model-name", default=MODEL_NAME)
    parser.add_argument("--max-steps", type=int, default=MAX_STEPS)
    parser.add_argument("--max-completion-length", type=int, default=MAX_COMPLETION_LENGTH)
    parser.add_argument("--limit", type=int, default=None)
    parser.add_argument("--curriculum", action=argparse.BooleanOptionalAction, default=True)
    parser.add_argument("--curriculum-stage-steps", type=int, default=8)
    parser.add_argument("--curriculum-eval-samples", type=int, default=2)
    parser.add_argument("--curriculum-accuracy-threshold", type=float, default=0.62)
    parser.add_argument("--curriculum-max-repeats", type=int, default=2)
    args = parser.parse_args()

    args.output_dir.mkdir(parents=True, exist_ok=True)
    train(args)


if __name__ == "__main__":
    main()
