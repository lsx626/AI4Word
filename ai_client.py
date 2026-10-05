"""Atria API 客户端：支持真正的 SSE 流式输出。

ai_stream 是生成内容用的流式接口：逐块 yield 内容片段，调用方可以边收边写。
ai_request 是非流式接口，用于排版阶段让 AI 生成代码（一次性返回完整结果）。

Atria 是 Intern AI（discovery 平台，https://discovery.intern-ai.org.cn/）提供的
大模型服务，接口为 OpenAI 兼容风格：https://discovery-api.intern-ai.org.cn/v1
默认模型为 Atria-Dawn-Preview，可用环境变量 ATRIA_MODEL 覆盖，密钥用 ATRIA_API_KEY。
"""
import json
import os
import time

import requests

from app import debug

DEFAULT_API_URL = "https://discovery-api.intern-ai.org.cn/v1/chat/completions"
API_URL = DEFAULT_API_URL
DEFAULT_MODEL = "Atria-Dawn-Preview"

_RETRYABLE = (
    requests.exceptions.ConnectionError,
    requests.exceptions.Timeout,
    requests.exceptions.ChunkedEncodingError,
)
_MAX_ATTEMPTS = 3

def _should_retry(exc):
    """???????????????????? 5xx???????

    4xx ????????????/???????????????????
    """
    if isinstance(exc, _RETRYABLE):
        return True
    if isinstance(exc, requests.exceptions.HTTPError):
        code = getattr(getattr(exc, "response", None), "status_code", None)
        return isinstance(code, int) and 500 <= code < 600
    return False


def _model(model):
    """解析实际使用的模型名：参数 > 环境变量 ATRIA_MODEL > 默认值。"""
    return model or os.getenv("ATRIA_MODEL") or DEFAULT_MODEL


def set_base_url(url):
    """覆盖 API 地址（GUI 设置面板里的 base_url 优先于内置默认值）。"""
    global API_URL
    if not url or not str(url).strip():
        return
    base = str(url).strip().rstrip("/")
    if not base:
        return
    if not base.endswith("/chat/completions"):
        base = base + "/chat/completions"
    API_URL = base


def get_base_url():
    """反向取出不含 /chat/completions 的 base，用于设置面板回显。"""
    if API_URL.endswith("/chat/completions"):
        return API_URL[: -len("/chat/completions")]
    return API_URL


def ai_request(prompt, api_key, system_prompt, model=DEFAULT_MODEL, timeout=180,
               attempts=_MAX_ATTEMPTS):
    """?????????????

    ???**????**????? / HTTP 4xx/5xx ???????GUI ?
    CLI??????????????????????? print??????
    ?????????????
    ?????? 5xx ???????????????????????
    3 ???????????????"????"???????????
    ??????????
    """
    headers = {"Content-Type": "application/json", "Authorization": f"Bearer {api_key}"}
    data = {
        "model": _model(model),
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": prompt},
        ],
    }
    t0 = time.time()
    attempt = 0
    while True:
        attempt += 1
        try:
            resp = requests.post(API_URL, headers=headers, json=data, timeout=timeout)
            resp.raise_for_status()
            content = resp.json()["choices"][0]["message"]["content"]
        except Exception as exc:
            debug.exc("api_request_failed", url=API_URL, model=data["model"],
                      prompt_chars=len(prompt), has_key=bool(api_key),
                      elapsed=round(time.time() - t0, 2))
            # ???????"????"??????????????
            # ai_stream ? received==0 ????????????????
            if not (attempt < attempts and _should_retry(exc)):
                raise
            debug.warn("api_request_retry", attempt=attempt, max_attempts=attempts,
                       error=str(exc)[:200], elapsed=round(time.time() - t0, 2))
            time.sleep(min(0.6 * attempt, 2.0))
            continue
        debug.log("api_request_ok", url=API_URL, model=data["model"],
                  prompt_chars=len(prompt), sys_prompt_chars=len(system_prompt),
                  has_key=bool(api_key), elapsed=round(time.time() - t0, 2),
                  chars=len(content), prompt=prompt, code=content, full=True)
        return content


def ai_stream(prompt, api_key, system_prompt, model=DEFAULT_MODEL, connect=10, read=600,
              attempts=_MAX_ATTEMPTS):
    """流式请求（SSE），逐个 yield 内容片段。

    生成结束自然 return；网络或 HTTP 错误时**抛出异常**——吞掉只 print 的
    话，窗口化打包的 GUI 里用户对失败一无所知（引擎会 catch 并提示）。
    连接级故障（TLS 握手中断/连接重置等）在尚未产出任何内容时会自动重试，
    最多 attempts 次；已在流的中途则绝不重试，避免内容重复。
    """
    headers = {"Content-Type": "application/json", "Authorization": f"Bearer {api_key}"}
    data = {
        "model": _model(model),
        "stream": True,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": prompt},
        ],
    }
    t0 = time.time()
    received = 0
    attempt = 0
    while True:
        attempt += 1
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
                        received += len(piece)
                        yield piece
            return
        except Exception as exc:
            debug.exc("api_stream_failed", url=API_URL, model=data["model"],
                      received=received, elapsed=round(time.time() - t0, 2))
            can_retry = (
                attempt < attempts
                and received == 0
                and _should_retry(exc)
            )
            if not can_retry:
                raise
            debug.warn("api_stream_retry", attempt=attempt, max_attempts=attempts,
                       error=str(exc)[:200], received=received,
                       elapsed=round(time.time() - t0, 2))
            time.sleep(min(0.6 * attempt, 2.0))
        finally:
            debug.log("api_stream_done", url=API_URL, model=data["model"],
                      received=received, elapsed=round(time.time() - t0, 2))
