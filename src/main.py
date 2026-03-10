# -*- coding: utf-8 -*-
"""
WordReader - Word 文档学术润色工具

用法：
    python main.py                          # 使用 config.json 配置
    python main.py --input thesis.docx      # 指定输入文件
    python main.py --split-only             # 仅拆分，不调用 AI
    python main.py --chapter 2              # 仅润色第 2 章
"""

import argparse
import json
import os
import sys
import time

# 确保可以导入同目录下的模块（无论从哪个目录运行）
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_PROJECT_ROOT = os.path.dirname(_SCRIPT_DIR)
if _SCRIPT_DIR not in sys.path:
    sys.path.insert(0, _SCRIPT_DIR)

from splitter import split_document, save_chapters
from api_client import OpenRouterClient
from prompts import ACADEMIC_POLISH_PROMPT, ACADEMIC_POLISH_PROMPT_WITH_IMAGES

# 默认配置文件在项目根目录
_DEFAULT_CONFIG = os.path.join(_PROJECT_ROOT, "config.json")


def load_config(config_path: str = _DEFAULT_CONFIG) -> dict:

    """加载配置文件"""
    if not os.path.exists(config_path):
        print(f"❌ 配置文件不存在: {config_path}")
        print("请复制 config.json 并填写 API key")
        sys.exit(1)

    with open(config_path, 'r', encoding='utf-8') as f:
        config = json.load(f)

    if config.get("openrouter_api_key", "").startswith("sk-or-v1-xxx"):
        print("❌ 请在 config.json 中填写您的 OpenRouter API key")
        sys.exit(1)

    return config


def polish_chapter(client: OpenRouterClient, chapter, chapter_dir: str) -> str:
    """
    对单个章节进行润色。

    Args:
        client: OpenRouter API 客户端
        chapter: Chapter 对象
        chapter_dir: 章节输出目录

    Returns:
        润色后的文本
    """
    chapter_content = f"# {chapter.title}\n\n{chapter.text}"

    if chapter.has_images:
        # 含图片，使用多模态 prompt
        prompt = ACADEMIC_POLISH_PROMPT_WITH_IMAGES.format(chapter_content=chapter_content)
        result = client.polish_text_with_images(prompt, chapter.images)
    else:
        # 纯文本
        prompt = ACADEMIC_POLISH_PROMPT.format(chapter_content=chapter_content)
        result = client.polish_text(prompt)

    # 保存润色结果
    polished_path = os.path.join(chapter_dir, "polished.md")
    with open(polished_path, 'w', encoding='utf-8') as f:
        f.write(result)

    return result


def get_chapter_dir(output_dir: str, chapter) -> str:
    """获取章节目录路径"""
    import re
    safe_title = re.sub(r'[\\/:*?"<>|]', '_', chapter.title)
    safe_title = safe_title[:50]
    return os.path.join(output_dir, f"chapter_{chapter.index}_{safe_title}")


def main():
    parser = argparse.ArgumentParser(
        description="WordReader - Word 文档学术润色工具",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例：
  python main.py                           使用 config.json 中的默认配置
  python main.py --input paper.docx        指定输入文件
  python main.py --split-only              仅拆分文档，不调用 AI
  python main.py --chapter 2               仅润色第 2 章
  python main.py --heading-level 2         按二级标题拆分
  python main.py --model google/gemini-2.0-flash-001  使用其他模型
        """
    )

    parser.add_argument("--input", "-i", help="输入的 .docx 文件路径")
    parser.add_argument("--output", "-o", help="输出目录")
    parser.add_argument("--config", "-c", default=_DEFAULT_CONFIG, help="配置文件路径")
    parser.add_argument("--split-only", action="store_true", help="仅拆分文档，不调用 AI 润色")
    parser.add_argument("--chapter", type=int, help="仅处理指定章节（从 0 开始）")
    parser.add_argument("--heading-level", type=int, help="拆分的标题级别 (1-9)")
    parser.add_argument("--model", help="OpenRouter 模型名称")

    args = parser.parse_args()

    # 加载配置
    config = load_config(args.config)

    # 命令行参数覆盖配置文件
    input_file = args.input or config.get("input_file", "")
    output_dir = args.output or config.get("output_dir", "./output")
    heading_level = args.heading_level or config.get("heading_level", 1)
    model = args.model or config.get("model", "anthropic/claude-sonnet-4")

    if not input_file:
        print("❌ 请指定输入文件：python main.py --input your_file.docx")
        sys.exit(1)

    if not os.path.exists(input_file):
        print(f"❌ 输入文件不存在: {input_file}")
        sys.exit(1)

    # =============================
    # 第一步：拆分文档
    # =============================
    print("=" * 60)
    print(f"📄 WordReader - Word 文档学术润色工具")
    print("=" * 60)
    print(f"\n📂 输入文件: {input_file}")
    print(f"📁 输出目录: {output_dir}")
    print(f"📊 拆分级别: Heading {heading_level}")
    print(f"🤖 AI 模型:  {model}")
    print()

    print("🔍 正在拆分文档...")
    chapters = split_document(input_file, heading_level)

    if not chapters:
        print("❌ 未找到任何章节。请检查文档格式是否包含标题样式。")
        sys.exit(1)

    # 保存拆分结果
    save_chapters(chapters, output_dir)

    # 打印章节概览
    print(f"\n📋 文档共拆分为 {len(chapters)} 个章节：")
    print("-" * 60)
    for ch in chapters:
        img_info = f" 📷 {len(ch.images)} 张图片" if ch.has_images else ""
        text_len = len(ch.text)
        print(f"  [{ch.index}] {ch.title} ({text_len} 字{img_info})")
    print("-" * 60)

    if args.split_only:
        print("\n✅ 拆分完成（--split-only 模式，跳过 AI 润色）")
        return

    # =============================
    # 第二步：AI 润色
    # =============================
    print("\n🤖 开始 AI 学术润色...\n")

    client = OpenRouterClient(
        api_key=config["openrouter_api_key"],
        model=model,
        max_tokens=config.get("max_tokens", 8192),
        temperature=config.get("temperature", 0.3),
        request_interval=config.get("request_interval", 1.0),
        max_retries=config.get("max_retries", 3),
    )

    # 确定要处理的章节
    if args.chapter is not None:
        if 0 <= args.chapter < len(chapters):
            chapters_to_process = [chapters[args.chapter]]
        else:
            print(f"❌ 章节编号 {args.chapter} 不存在（有效范围: 0-{len(chapters)-1}）")
            sys.exit(1)
    else:
        chapters_to_process = chapters

    total = len(chapters_to_process)
    success_count = 0
    fail_count = 0
    start_time = time.time()

    for i, chapter in enumerate(chapters_to_process, 1):
        chapter_dir = get_chapter_dir(output_dir, chapter)

        # 检查是否已有润色结果（支持断点续传）
        polished_path = os.path.join(chapter_dir, "polished.md")
        if os.path.exists(polished_path) and args.chapter is None:
            print(f"  ⏭️  [{i}/{total}] {chapter.title} — 已有润色结果，跳过")
            success_count += 1
            continue

        print(f"  🔄 [{i}/{total}] 正在润色: {chapter.title}...", end="", flush=True)

        try:
            result = polish_chapter(client, chapter, chapter_dir)
            result_len = len(result)
            print(f" ✅ ({result_len} 字)")
            success_count += 1
        except Exception as e:
            print(f" ❌ 失败: {e}")
            fail_count += 1

    # =============================
    # 完成总结
    # =============================
    elapsed = time.time() - start_time
    print()
    print("=" * 60)
    print(f"🎉 润色完成！")
    print(f"   ✅ 成功: {success_count} 章")
    if fail_count > 0:
        print(f"   ❌ 失败: {fail_count} 章")
    print(f"   ⏱️  耗时: {elapsed:.1f}s")
    print(f"   📁 结果: {os.path.abspath(output_dir)}")
    print("=" * 60)


if __name__ == "__main__":
    main()
