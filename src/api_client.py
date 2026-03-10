# -*- coding: utf-8 -*-
"""
OpenRouter API 客户端

支持文本和多模态（文本+图片）请求，内置重试机制和速率控制。
"""

import base64
import time
import requests


class OpenRouterClient:
    """OpenRouter API 客户端"""

    BASE_URL = "https://openrouter.ai/api/v1/chat/completions"

    def __init__(
        self,
        api_key: str,
        model: str = "anthropic/claude-sonnet-4",
        max_tokens: int = 8192,
        temperature: float = 0.3,
        request_interval: float = 1.0,
        max_retries: int = 3,
    ):
        self.api_key = api_key
        self.model = model
        self.max_tokens = max_tokens
        self.temperature = temperature
        self.request_interval = request_interval
        self.max_retries = max_retries
        self._last_request_time = 0.0

    def _wait_for_rate_limit(self):
        """速率控制：确保请求之间有足够的间隔"""
        elapsed = time.time() - self._last_request_time
        if elapsed < self.request_interval:
            time.sleep(self.request_interval - elapsed)

    def _make_request(self, messages: list[dict]) -> str:
        """
        发送请求到 OpenRouter API，内置重试机制。

        Returns:
            AI 生成的文本内容
        """
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
            "HTTP-Referer": "https://github.com/wordreader",
            "X-Title": "WordReader Academic Polish",
        }

        payload = {
            "model": self.model,
            "max_tokens": self.max_tokens,
            "temperature": self.temperature,
            "messages": messages,
        }

        last_error = None

        for attempt in range(1, self.max_retries + 1):
            self._wait_for_rate_limit()

            try:
                self._last_request_time = time.time()
                response = requests.post(
                    self.BASE_URL,
                    headers=headers,
                    json=payload,
                    timeout=120,
                )

                if response.status_code == 200:
                    data = response.json()
                    return data["choices"][0]["message"]["content"]

                elif response.status_code == 429:
                    # 限流，使用指数退避
                    wait_time = 2 ** attempt * self.request_interval
                    print(f"  ⚠️  限流，等待 {wait_time:.1f}s 后重试 ({attempt}/{self.max_retries})")
                    time.sleep(wait_time)
                    last_error = f"HTTP 429: {response.text}"

                elif response.status_code >= 500:
                    # 服务端错误，重试
                    wait_time = 2 ** attempt
                    print(f"  ⚠️  服务器错误 {response.status_code}，等待 {wait_time}s 后重试 ({attempt}/{self.max_retries})")
                    time.sleep(wait_time)
                    last_error = f"HTTP {response.status_code}: {response.text}"

                else:
                    # 其他错误，不重试
                    raise RuntimeError(
                        f"API 请求失败 (HTTP {response.status_code}): {response.text}"
                    )

            except requests.exceptions.Timeout:
                wait_time = 2 ** attempt
                print(f"  ⚠️  请求超时，等待 {wait_time}s 后重试 ({attempt}/{self.max_retries})")
                time.sleep(wait_time)
                last_error = "请求超时"

            except requests.exceptions.ConnectionError as e:
                wait_time = 2 ** attempt
                print(f"  ⚠️  连接错误，等待 {wait_time}s 后重试 ({attempt}/{self.max_retries})")
                time.sleep(wait_time)
                last_error = f"连接错误: {e}"

        raise RuntimeError(f"在 {self.max_retries} 次重试后仍然失败: {last_error}")

    def polish_text(self, prompt: str) -> str:
        """
        发送纯文本润色请求。

        Args:
            prompt: 完整的润色提示词（包含章节内容）

        Returns:
            润色后的文本
        """
        messages = [
            {"role": "user", "content": prompt}
        ]
        return self._make_request(messages)

    def polish_text_with_images(
        self,
        prompt: str,
        images: list[dict],
    ) -> str:
        """
        发送多模态润色请求（文本+图片）。

        Args:
            prompt: 完整的润色提示词（包含章节内容）
            images: 图片列表，每项包含 "data" (bytes) 和 "content_type" (str)

        Returns:
            润色后的文本
        """
        content = [{"type": "text", "text": prompt}]

        for img in images:
            img_b64 = base64.b64encode(img["data"]).decode("utf-8")
            content_type = img.get("content_type", "image/png")

            # 跳过不支持的图片格式（如 EMF/WMF）
            if content_type in ("image/x-emf", "image/x-wmf"):
                print(f"  ℹ️  跳过不支持的图片格式: {img.get('filename', 'unknown')} ({content_type})")
                continue

            content.append({
                "type": "image_url",
                "image_url": {
                    "url": f"data:{content_type};base64,{img_b64}"
                }
            })

        messages = [
            {"role": "user", "content": content}
        ]
        return self._make_request(messages)
