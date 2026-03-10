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

from splitter import split_document, save_chapters, split_long_chapters
from api_client import OpenRouterClient
from prompts import (
    ACADEMIC_POLISH_PROMPT, 
    ACADEMIC_POLISH_PROMPT_WITH_IMAGES,
    ABSTRACT_POLISH_PROMPT,
    CONCLUSION_POLISH_PROMPT,
    EXPERIMENT_POLISH_PROMPT
)

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


def polish_chapter(client: OpenRouterClient, chapter, chapter_dir: str, force_prompt_type: str = "auto") -> str:
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

    # 根据章节类型选择合适的提示词
    if chapter.has_images:
        # 含图片章节，使用带图片的通用提示词
        prompt_template = ACADEMIC_POLISH_PROMPT_WITH_IMAGES
        prompt_type = "多模态通用"
        prompt = prompt_template.format(chapter_content=chapter_content)
        result = client.polish_text_with_images(prompt, chapter.images)
    else:
        # 纯文本章节，根据类型选择专用提示词
        if force_prompt_type == "auto":
            prompt_template = select_prompt_by_chapter_type(chapter.title)
        else:
            prompt_template = get_prompt_by_force_type(force_prompt_type)
        
        prompt_type = get_prompt_type_name(prompt_template)
        prompt = prompt_template.format(chapter_content=chapter_content)
        result = client.polish_text(prompt)
    
    print(f"  📝 使用 {prompt_type} 提示词")

    # 验证结果不为空
    if result is None or result.strip() == "":
        raise ValueError("润色结果为空")

    # 保存润色结果
    polished_path = os.path.join(chapter_dir, "polished.md")
    with open(polished_path, 'w', encoding='utf-8') as f:
        f.write(result)

    return result


def select_prompt_by_chapter_type(chapter_title: str) -> str:
    """
    根据章节标题选择最合适的润色提示词
    
    Args:
        chapter_title: 章节标题
        
    Returns:
        对应的提示词模板
    """
    title_lower = chapter_title.lower()
    
    # 摘要相关
    if any(keyword in title_lower for keyword in ['摘要', 'abstract', '摘 要']):
        return ABSTRACT_POLISH_PROMPT
    
    # 结论相关
    if any(keyword in title_lower for keyword in ['结论', 'conclusion', '总结', '结语']):
        return CONCLUSION_POLISH_PROMPT
    
    # 实验/结果相关
    if any(keyword in title_lower for keyword in [
        '实验', 'experiment', '结果', 'result', '分析', 'analysis', 
        '对比', 'comparison', '性能', 'performance', '评估', 'evaluation'
    ]):
        return EXPERIMENT_POLISH_PROMPT
    
    # 默认使用通用学术提示词
    return ACADEMIC_POLISH_PROMPT


def get_prompt_by_force_type(force_type: str) -> str:
    """根据强制指定的类型返回提示词"""
    if force_type == "abstract":
        return ABSTRACT_POLISH_PROMPT
    elif force_type == "conclusion":
        return CONCLUSION_POLISH_PROMPT
    elif force_type == "experiment":
        return EXPERIMENT_POLISH_PROMPT
    else:  # "general"
        return ACADEMIC_POLISH_PROMPT


def get_prompt_type_name(prompt_template: str) -> str:
    """获取提示词类型的中文名称"""
    if prompt_template == ABSTRACT_POLISH_PROMPT:
        return "摘要专用"
    elif prompt_template == CONCLUSION_POLISH_PROMPT:
        return "结论专用"
    elif prompt_template == EXPERIMENT_POLISH_PROMPT:
        return "实验专用"
    else:
        return "通用学术"


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
    parser.add_argument("--test-api", action="store_true", help="仅测试 API key 有效性")
    parser.add_argument("--chapter", type=int, help="仅处理指定章节（从 0 开始）")
    parser.add_argument("--heading-level", type=int, help="拆分的标题级别 (1-9)")
    parser.add_argument("--model", help="OpenRouter 模型名称")
    parser.add_argument("--max-chars", type=int, help="章节最大字符数，超过会按二级标题分割")
    parser.add_argument("--prompt-type", choices=["auto", "abstract", "conclusion", "experiment", "general"], 
                        default="auto", help="强制使用指定类型的润色提示词")

    args = parser.parse_args()

    # 加载配置
    config = load_config(args.config)

    # 如果只是测试 API
    if args.test_api:
        print("🔑 测试 API key...")
        model = args.model or config.get("model", "anthropic/claude-sonnet-4")
        client = OpenRouterClient(
            api_key=config["openrouter_api_key"],
            model=model,
            max_tokens=50,
            temperature=0.01,
            request_interval=1.0,
            max_retries=1,
        )
        
        if client.test_api_key():
            print(f"✅ API key 正常，模型: {model}")
        else:
            print("❌ API key 测试失败")
        return

    # 命令行参数覆盖配置文件
    input_file = args.input or config.get("input_file", "")
    output_dir = args.output or config.get("output_dir", "./output")
    heading_level = args.heading_level or config.get("heading_level", 1)
    model = args.model or config.get("model", "anthropic/claude-sonnet-4")
    max_chars = args.max_chars or config.get("max_chars", 6000)

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

    # 检查并分割过长章节
    print("\n📏 检查章节长度，按需细分...")
    original_count = len(chapters)
    chapters = split_long_chapters(chapters, max_chars=max_chars)
    
    if len(chapters) > original_count:
        print(f"✅ 已将过长章节细分，总章节数：{original_count} → {len(chapters)}")

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

    # 测试 API key 有效性
    print("🔑 正在验证 API key...")
    if not client.test_api_key():
        print("❌ API key 验证失败，请检查配置文件中的 openrouter_api_key")
        print("💡 提示：")
        print("   1. 确认 API key 格式正确 (应以 sk-or-v1- 开头)")
        print("   2. 检查 API key 是否已过期")
        print("   3. 登录 OpenRouter 账户确认余额充足")
        sys.exit(1)
    print("✅ API key 验证成功")

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
            result = polish_chapter(client, chapter, chapter_dir, args.prompt_type)
            if result is None:
                print(f" ❌ 失败: API返回内容为空")
                fail_count += 1
                continue
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
        print()
        print("💡 失败原因可能包括：")
        print("   • 内容过长超出模型限制 (尝试 --max-chars 参数调小分割阈值)")
        print("   • 网络连接问题 (检查网络连接)")
        print("   • API 配额不足 (检查 OpenRouter 余额)")
        print("   • 内容包含敏感信息被拒绝")
        print()
        print("🔄 可以重新运行程序，已成功的章节会自动跳过")
    print(f"   ⏱️  耗时: {elapsed:.1f}s")
    print(f"   📁 结果: {os.path.abspath(output_dir)}")
    print("=" * 60)


if __name__ == "__main__":
    main()
