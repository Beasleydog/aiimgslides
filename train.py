import argparse
import json
import math
import os
import re
from pathlib import Path

from PIL import Image

from grader import grade_json, schema_reward


DATA_DIR = Path("output")
OUTPUT_DIR = Path("model_output/slide_json_grpo")
MODEL_NAME = "Qwen/Qwen2.5-VL-3B-Instruct"

MAX_STEPS = 150
MAX_PROMPT_LENGTH = 2048
MAX_COMPLETION_LENGTH = 768
PER_DEVICE_BATCH_SIZE = 2
GRADIENT_ACCUMULATION_STEPS = 4
NUM_GENERATIONS = 2
LEARNING_RATE = 1e-5
LORA_R = 32
LORA_ALPHA = 64
WARMUP_RATIO = 0.03
WEIGHT_DECAY = 0.01
MAX_GRAD_NORM = 0.3
SAVE_STEPS = 50
SAVE_TOTAL_LIMIT = 3
IMAGE_MIN_PIXELS = 256 * 28 * 28
IMAGE_MAX_PIXELS = 512 * 512

os.environ["WANDB_MODE"] = "disabled"
os.environ["WANDB_PROJECT"] = "disabled"
os.environ["TOKENIZERS_PARALLELISM"] = "false"


USER_PROMPT = """Infer the PowerPoint-like scene graph from this slide image.

Return exactly:
<json>{
  "version": 1,
  "slide": {"width": 13.333, "height": 7.5, "image_file": null},
  "background": object_or_null,
  "objects": [object, ...]
}</json>

Each object needs: id, type, z_order, bbox {x,y,w,h}, properties.
Use inches. Do not include markdown."""


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
            examples.append({"target_json": str(json_path), "image_path": str(image_path)})
    if not examples:
        raise FileNotFoundError(f"No slide_*.json / slide_*.jpg pairs found in {data_dir}")
    return examples


def load_dataset(data_dir, limit=None):
    from datasets import Dataset

    rows = []
    for example in paired_examples(data_dir)[:limit]:
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
            }
        )
    return Dataset.from_list(rows)


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
            prediction = extract_json(text)
            grade = grade_json(target_path, prediction)
            structure_score = (schema_reward(prediction) + 1.0) / 2.0
            task_score = (float(grade["reward"]) + 1.0) / 2.0
            target_len = len(Path(target_path).read_text(encoding="utf-8"))
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


def train(args):
    import torch
    from peft import LoraConfig
    from trl import GRPOConfig, GRPOTrainer

    gpu = require_gpu()
    print(f"GPU: {gpu['name']} ({gpu['total_gb']} GB), bf16={gpu['bf16']}")

    dataset = load_dataset(args.data_dir, args.limit)
    model, processor = load_model_and_processor(args.model_name)

    peft_config = LoraConfig(
        r=LORA_R,
        lora_alpha=LORA_ALPHA,
        lora_dropout=0.0,
        bias="none",
        task_type="CAUSAL_LM",
        target_modules=["q_proj", "k_proj", "v_proj", "o_proj", "gate_proj", "up_proj", "down_proj"],
    )
    grpo_config = make_grpo_config(
        GRPOConfig,
        {
            "output_dir": str(args.output_dir),
            "per_device_train_batch_size": PER_DEVICE_BATCH_SIZE,
            "gradient_accumulation_steps": GRADIENT_ACCUMULATION_STEPS,
            "learning_rate": LEARNING_RATE,
            "max_steps": args.max_steps,
            "logging_steps": 1,
            "save_steps": SAVE_STEPS,
            "save_total_limit": SAVE_TOTAL_LIMIT,
            "remove_unused_columns": False,
            "max_prompt_length": MAX_PROMPT_LENGTH,
            "max_completion_length": MAX_COMPLETION_LENGTH,
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
        },
    )
    trainer = GRPOTrainer(
        model=model,
        processing_class=processor,
        reward_funcs=[slide_json_reward_func],
        args=grpo_config,
        train_dataset=dataset,
        peft_config=peft_config,
    )
    trainer.train()
    trainer.save_model(args.output_dir)
    processor.save_pretrained(args.output_dir)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--data-dir", type=Path, default=DATA_DIR)
    parser.add_argument("--output-dir", type=Path, default=OUTPUT_DIR)
    parser.add_argument("--model-name", default=MODEL_NAME)
    parser.add_argument("--max-steps", type=int, default=MAX_STEPS)
    parser.add_argument("--limit", type=int, default=None)
    args = parser.parse_args()

    args.output_dir.mkdir(parents=True, exist_ok=True)
    train(args)


if __name__ == "__main__":
    main()
