import argparse
import json
import math
import os
import re
import shutil
from pathlib import Path

# This avoids stale/mismatched generated GRPO wrappers like:
# grpo_accumulated_loss() missing old_logps/ref_logps.
shutil.rmtree("unsloth_compiled_cache", ignore_errors=True)
os.environ["UNSLOTH_COMPILE_DISABLE"] = "1"
os.environ["TORCHDYNAMO_DISABLE"] = "1"
os.environ["TORCH_COMPILE_DISABLE"] = "1"
os.environ["TORCHINDUCTOR_DISABLE"] = "1"
os.environ["WANDB_MODE"] = "disabled"
os.environ["WANDB_PROJECT"] = "disabled"

try:
    import torch

    torch._dynamo.config.suppress_errors = True
    torch._dynamo.disable()
except Exception:
    pass

try:
    import unsloth  # Must be imported before trl/transformers/peft when installed.
    from unsloth import FastVisionModel
except ModuleNotFoundError:
    FastVisionModel = None

from grader import grade_json, schema_reward


DATA_DIR = Path("output")
MODEL_CANDIDATES = [
    "unsloth/gemma-4-E2B-it-unsloth-bnb-4bit",
    "unsloth/gemma-4-E4B-it-unsloth-bnb-4bit",
    "unsloth/gemma-4-E4B-it",
    "google/gemma-4-E4B-it",
    "unsloth/Qwen3-VL-4B-Instruct-unsloth-bnb-4bit",
    "unsloth/Qwen3-VL-2B-Instruct-unsloth-bnb-4bit",
    "unsloth/Qwen2.5-VL-3B-Instruct-unsloth-bnb-4bit",
    "unsloth/gemma-3-4b-it-unsloth-bnb-4bit",
]
OUTPUT_DIR = Path("model_output/slide_json_grpo")
MAX_SEQ_LENGTH = 2048
MAX_STEPS = 200
PER_DEVICE_BATCH_SIZE = 2
GRADIENT_ACCUMULATION_STEPS = 4
LEARNING_RATE = 2e-6
LORA_R = 16
NUM_GENERATIONS = 2
MAX_COMPLETION_LENGTH = 1024

SYSTEM_PROMPT = """<|think|>
You are reconstructing editable PowerPoint slides from screenshots.
Think briefly about the slide structure, then output only a <json>...</json> block."""

USER_PROMPT = """Infer the PowerPoint-like scene graph from this slide image.

Return exactly:
<json>{
  "version": 1,
  "slide": {"width": 13.333, "height": 7.5, "image_file": null},
  "background": object_or_null,
  "objects": [object, ...]
}</json>

Each object needs: id, type, z_order, bbox {x,y,w,h}, properties.
Use inches. The image appears before this instruction. Do not include markdown."""


def require_gpu():
    import torch

    if not torch.cuda.is_available():
        raise RuntimeError("No CUDA GPU detected. Gemma 4 vision GRPO with Unsloth needs an NVIDIA GPU.")
    index = torch.cuda.current_device()
    props = torch.cuda.get_device_properties(index)
    return {
        "name": props.name,
        "total_gb": round(props.total_memory / 1024**3, 2),
        "bf16": bool(torch.cuda.is_bf16_supported()),
    }


def paired_examples(data_dir):
    data_dir = Path(data_dir)
    examples = []
    for json_path in sorted(data_dir.glob("slide_*.json")):
        image_path = json_path.with_suffix(".jpg")
        if image_path.exists():
            examples.append({"target_json": str(json_path), "image_path": str(image_path)})
    if not examples:
        raise FileNotFoundError(f"No slide_*.json / slide_*.jpg pairs found in {data_dir}")
    return examples


def build_prompt(example):
    return [
        {"role": "system", "content": [{"type": "text", "text": SYSTEM_PROMPT}]},
        {
            "role": "user",
            "content": [
                {"type": "image", "image": example["image_path"]},
                {"type": "text", "text": USER_PROMPT},
            ],
        },
    ]


def load_grpo_dataset(data_dir, limit=None):
    from datasets import Dataset

    rows = []
    for example in paired_examples(data_dir)[:limit]:
        rows.append(
            {
                "prompt": build_prompt(example),
                "target_json": example["target_json"],
                "image_path": example["image_path"],
            }
        )
    return Dataset.from_list(rows)


def from_pretrained_compat(**kwargs):
    try:
        return FastVisionModel.from_pretrained(**kwargs, fast_inference=False)
    except TypeError:
        return FastVisionModel.from_pretrained(**kwargs)


def load_model(model_name=None):
    if FastVisionModel is None:
        raise ModuleNotFoundError(
            "Unsloth is not installed. In Colab, install it before training with:\n"
            'pip install --upgrade --no-cache-dir "unsloth[colab-new] @ git+https://github.com/unslothai/unsloth.git"\n'
            'pip install --upgrade --no-cache-dir "git+https://github.com/unslothai/unsloth-zoo.git"'
        )
    candidates = [model_name] if model_name else MODEL_CANDIDATES
    errors = []
    kwargs = dict(
        max_seq_length=MAX_SEQ_LENGTH,
        load_in_4bit=True,
        use_gradient_checkpointing="unsloth",
    )

    model = tokenizer = None
    loaded_name = None
    for candidate in candidates:
        try:
            print(f"Trying model: {candidate}")
            model, tokenizer = from_pretrained_compat(model_name=candidate, **kwargs)
            loaded_name = candidate
            break
        except NotImplementedError as exc:
            errors.append(f"{candidate}: {exc}")
            continue
        except OSError as exc:
            errors.append(f"{candidate}: {exc}")
            continue
        except ValueError as exc:
            errors.append(f"{candidate}: {exc}")
            continue

    if model is None:
        message = "\n\n".join(errors)
        raise RuntimeError(
            "Could not load any configured vision model. If you need Gemma 4 specifically, update Unsloth in Colab with:\n"
            'pip uninstall unsloth unsloth_zoo -y\n'
            'pip install --upgrade --no-cache-dir "unsloth[colab-new] @ git+https://github.com/unslothai/unsloth.git"\n'
            'pip install --upgrade --no-cache-dir "git+https://github.com/unslothai/unsloth-zoo.git"\n\n'
            f"Load errors:\n{message}"
        )
    print(f"Loaded model: {loaded_name}")

    model = FastVisionModel.get_peft_model(
        model,
        finetune_vision_layers=False,
        finetune_language_layers=True,
        finetune_attention_modules=True,
        finetune_mlp_modules=True,
        r=LORA_R,
        lora_alpha=LORA_R,
        lora_dropout=0,
        bias="none",
        target_modules="all-linear",
    )
    if hasattr(FastVisionModel, "for_training"):
        FastVisionModel.for_training(model)
    return model, tokenizer


def completion_to_text(completion):
    if isinstance(completion, str):
        return completion
    if isinstance(completion, dict):
        return completion_to_text(completion.get("content", ""))
    if isinstance(completion, list):
        parts = []
        for item in completion:
            if isinstance(item, dict):
                if "text" in item:
                    parts.append(str(item["text"]))
                else:
                    parts.append(completion_to_text(item.get("content", "")))
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
    candidate = match.group(1) if match else balanced_json_slice(text)
    return json.loads(candidate)


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


def train(args):
    from trl import GRPOConfig, GRPOTrainer
    from unsloth.trainer import UnslothVisionDataCollator

    gpu = require_gpu()
    print(f"GPU: {gpu['name']} ({gpu['total_gb']} GB), bf16={gpu['bf16']}")
    dataset = load_grpo_dataset(args.data_dir, args.limit)
    model, tokenizer = load_model(args.model_name)

    grpo_kwargs = dict(
        output_dir=str(args.output_dir),
        per_device_train_batch_size=PER_DEVICE_BATCH_SIZE,
        gradient_accumulation_steps=GRADIENT_ACCUMULATION_STEPS,
        learning_rate=LEARNING_RATE,
        max_steps=args.max_steps,
        logging_steps=1,
        remove_unused_columns=False,
        max_completion_length=MAX_COMPLETION_LENGTH,
        num_generations=NUM_GENERATIONS,
        report_to="none",
        bf16=gpu["bf16"],
        fp16=not gpu["bf16"],
        optim="adamw_8bit",
        temperature=0.8,
        top_p=0.95,
        max_grad_norm=0.1,
        unsloth_grpo_mini_batch=1,
        unsloth_logit_chunk_multiplier=4,
    )
    config = make_grpo_config(GRPOConfig, grpo_kwargs)

    trainer = GRPOTrainer(
        model=model,
        processing_class=tokenizer,
        reward_funcs=[slide_json_reward_func],
        train_dataset=dataset,
        data_collator=UnslothVisionDataCollator(model, tokenizer),
        args=config,
    )
    trainer.train()
    model.save_pretrained(args.output_dir)
    tokenizer.save_pretrained(args.output_dir)


def make_grpo_config(config_cls, kwargs):
    kwargs = dict(kwargs)
    kwargs["use_vllm"] = False
    while True:
        try:
            return config_cls(**kwargs)
        except TypeError as exc:
            message = str(exc)
            match = re.search(r"unexpected keyword argument '([^']+)'", message)
            if not match:
                raise
            removed = match.group(1)
            kwargs.pop(removed, None)
            print(f"GRPOConfig does not support {removed}; continuing without it.")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--data-dir", type=Path, default=DATA_DIR)
    parser.add_argument("--output-dir", type=Path, default=OUTPUT_DIR)
    parser.add_argument("--model-name", default=None, help="Override model. Defaults to Gemma 4 candidates, then Gemma 3 vision fallback.")
    parser.add_argument("--max-steps", type=int, default=MAX_STEPS)
    parser.add_argument("--limit", type=int, default=None)
    args = parser.parse_args()

    args.output_dir.mkdir(parents=True, exist_ok=True)
    train(args)


if __name__ == "__main__":
    main()
