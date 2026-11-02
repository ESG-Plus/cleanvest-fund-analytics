# -*- coding: utf-8 -*-
"""Minimal in-process chat inference (transformers, no Ollama server needed).

Loads the model once (lazily, on first call) and keeps it in memory. Generation
is blocking, so the calling app (e.g. Streamlit) will freeze until a response
is ready — acceptable for this test setup.
"""

from transformers import AutoModelForCausalLM, AutoTokenizer

MODEL_NAME = "Qwen/Qwen2.5-3B-Instruct"

_model = None
_tokenizer = None


def _load():
    global _model, _tokenizer
    if _model is None:
        _tokenizer = AutoTokenizer.from_pretrained(MODEL_NAME)
        _model = AutoModelForCausalLM.from_pretrained(
            MODEL_NAME, torch_dtype="auto", device_map="auto"
        )
    return _model, _tokenizer


def ask_question(messages: list[dict], max_new_tokens: int = 512) -> str:
    """Run one chat completion. `messages` is a list of {"role", "content"} dicts."""
    model, tokenizer = _load()
    text = tokenizer.apply_chat_template(
        messages, tokenize=False, add_generation_prompt=True
    )
    model_inputs = tokenizer([text], return_tensors="pt").to(model.device)
    generated_ids = model.generate(**model_inputs, max_new_tokens=max_new_tokens)
    generated_ids = [
        output_ids[len(input_ids):]
        for input_ids, output_ids in zip(model_inputs.input_ids, generated_ids)
    ]
    return tokenizer.batch_decode(generated_ids, skip_special_tokens=True)[0]
