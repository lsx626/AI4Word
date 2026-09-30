"""DeepSeek API 客户端：支持真正的 SSE 流式输出。

ai_stream 是生成内容用的流式接口：逐块 yield 内容片段，调用方可以边收边写。
ai_request 是非流式接口，用于排版阶段让 AI 生成代码（一次性返回完整结果）。
"""
import json

import requests

API_URL = "https://api.deepseek.com/chat/completions"
DEFAULT_MODEL = "deepseek-chat"


def ai_request(prompt, api_key, system_prompt, model=DEFAULT_MODEL, timeout=180):
    """非流式请求，返回完整文本；出错返回 None。"""
    headers = {"Content-Type": "application/json", "Authorization": f"Bearer {api_key}"}
    data = {
        "model": model,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": prompt},
        ],
    }
    try:
        resp = requests.post(API_URL, headers=headers, json=data, timeout=timeout)
        resp.raise_for_status()
        return resp.json()["choices"][0]["message"]["content"]
    except Exception as e:
        print(f"调用API时出错: {e}")
        return None


def ai_stream(prompt, api_key, system_prompt, model=DEFAULT_MODEL, connect=10, read=600):
    """流式请求（SSE），逐个 yield 内容片段。

    生成结束自然 return；网络或 HTTP 错误时打印错误并 return。
    """
    headers = {"Content-Type": "application/json", "Authorization": f"Bearer {api_key}"}
    data = {
        "model": model,
        "stream": True,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": prompt},
        ],
    }
    try:
        with requests.post(
            API_URL, headers=headers, json=data, stream=True, timeout=(connect, read)
        ) as resp:
            resp.raise_for_status()
            for raw in resp.iter_lines(decode_unicode=True):
                if not raw:
                    continue
                line = raw.strip()
                if not line.startswith("data:"):
                    continue
                payload = line[len("data:"):].strip()
                if payload == "[DONE]":
                    return
                try:
                    chunk = json.loads(payload)
                except json.JSONDecodeError:
                    continue
                choices = chunk.get("choices") or []
                if not choices:
                    continue
                delta = choices[0].get("delta") or {}
                piece = delta.get("content")
                if piece:
                    yield piece
    except Exception as e:
        print(f"流式请求出错: {e}")
