# -*- coding: utf-8 -*-
"""
Word 文档章节拆分模块

利用 PyWin32 导出 Word HTML 保留完美的公式/图片，再结合 Pandoc 转化为纯净 Markdown，
最后按标题样式（Heading 1/2...）将该 Markdown 分割为独立章节。
"""

import os
import re
import shutil
import win32com.client
import pypandoc
from dataclasses import dataclass, field

@dataclass
class Chapter:
    """表示一个拆分出的章节"""
    index: int
    title: str
    paragraphs: list[str] = field(default_factory=list)
    images: list[dict] = field(default_factory=list)  # [{"filename": ..., "data": bytes}]

    @property
    def text(self) -> str:
        """获取章节纯文本，段落之间用换行分隔，自动过滤参考文献"""
        raw_text = "\n\n".join(p for p in self.paragraphs if p.strip())
        
        # 对前言章节进行特殊处理，只保留中英文摘要
        if "前言" in self.title:
            return filter_preface_content(raw_text)
        
        return filter_references(raw_text)

    @property
    def has_images(self) -> bool:
        return len(self.images) > 0
    
    @property
    def char_count(self) -> int:
        """获取章节字符数"""
        return len(self.text)


def split_long_chapters(chapters: list[Chapter], max_chars: int = 6000) -> list[Chapter]:
    """
    将过长的章节按照二级或三级标题进一步分割
    
    Args:
        chapters: 原始章节列表
        max_chars: 最大字符数，超过此数值的章节会被细分
    
    Returns:
        细分后的章节列表
    """
    result_chapters = []
    chapter_index = 0
    
    for chapter in chapters:
        if chapter.char_count <= max_chars:
            # 章节不长，直接添加但更新索引
            chapter.index = chapter_index
            result_chapters.append(chapter)
            chapter_index += 1
            continue
            
        # 章节过长，尝试按标题分割
        print(f"  📏 章节 '{chapter.title}' 过长 ({chapter.char_count} 字)，尝试按标题分割...")
        
        sub_chapters = split_by_headers(chapter, chapter_index, max_chars)
        
        if len(sub_chapters) <= 1:
            # 无法进一步分割，保留原章节但给出警告
            print(f"  ⚠️  无法进一步分割，保留原章节")
            chapter.index = chapter_index
            result_chapters.append(chapter)
            chapter_index += 1
        else:
            # 成功分割
            print(f"  ✅ 已分割为 {len(sub_chapters)} 个子章节")
            result_chapters.extend(sub_chapters)
            chapter_index += len(sub_chapters)
            
    return result_chapters


def split_by_headers(chapter: Chapter, start_index: int, max_chars: int = 6000) -> list[Chapter]:
    """
    按二级标题(## )或三级标题(### )分割单个章节
    
    Args:
        chapter: 要分割的章节
        start_index: 起始索引号
        max_chars: 最大字符数，决定分割策略
        
    Returns:
        分割后的子章节列表
    """
    # 首先尝试按二级标题分割
    sub_chapters = split_by_secondary_headers(chapter, start_index)
    
    # 如果二级标题分割后仍有过长章节，进一步按三级标题分割
    if len(sub_chapters) > 1:
        final_chapters = []
        current_index = start_index
        
        for sub_chapter in sub_chapters:
            if sub_chapter.char_count <= max_chars:
                sub_chapter.index = current_index
                final_chapters.append(sub_chapter)
                current_index += 1
            else:
                # 子章节仍然过长，尝试按三级标题分割
                tertiary_chapters = split_by_tertiary_headers(sub_chapter, current_index)
                if len(tertiary_chapters) > 1:
                    print(f"    📐 进一步按三级标题分割 '{sub_chapter.title}' 为 {len(tertiary_chapters)} 个子章节")
                    final_chapters.extend(tertiary_chapters)
                    current_index += len(tertiary_chapters)
                else:
                    sub_chapter.index = current_index
                    final_chapters.append(sub_chapter)
                    current_index += 1
        
        return final_chapters
    
    # 如果无法按二级标题分割，尝试直接按三级标题分割
    tertiary_chapters = split_by_tertiary_headers(chapter, start_index)
    return tertiary_chapters


def split_by_secondary_headers(chapter: Chapter, start_index: int) -> list[Chapter]:
    """
    按二级标题(## )分割单个章节
    
    Args:
        chapter: 要分割的章节
        start_index: 起始索引号
        
    Returns:
        分割后的子章节列表
    """
    text = chapter.text
    lines = text.split('\n')
    
    # 寻找二级标题模式
    h2_pattern = re.compile(r'^##\s+(.+)')
    
    sub_chapters = []
    current_lines = []
    current_title = chapter.title
    sub_index = 0
    
    # 为了图片分配，我们需要跟踪每个子章节包含的图片引用
    img_pattern = re.compile(r'!\[[^\]]*\]\([^)]*images/([^)]+)\)')
    
    for line in lines:
        match = h2_pattern.match(line)
        if match:
            # 找到新的二级标题，保存前面的内容
            if current_lines:
                sub_chapter = create_sub_chapter_with_images(
                    start_index + sub_index, current_title, current_lines, chapter.images, img_pattern
                )
                sub_chapters.append(sub_chapter)
                sub_index += 1
            
            # 开始新的子章节
            current_title = f"{chapter.title} - {match.group(1)}"
            current_lines = [line]
        else:
            current_lines.append(line)
    
    # 处理最后一个子章节
    if current_lines:
        sub_chapter = create_sub_chapter_with_images(
            start_index + sub_index, current_title, current_lines, chapter.images, img_pattern
        )
        sub_chapters.append(sub_chapter)
    
    return sub_chapters if len(sub_chapters) > 1 else [chapter]


def split_by_tertiary_headers(chapter: Chapter, start_index: int) -> list[Chapter]:
    """
    按三级标题(### )分割单个章节
    
    Args:
        chapter: 要分割的章节
        start_index: 起始索引号
        
    Returns:
        分割后的子章节列表
    """
    text = chapter.text
    lines = text.split('\n')
    
    # 寻找三级标题模式
    h3_pattern = re.compile(r'^###\s+(.+)')
    
    sub_chapters = []
    current_lines = []
    current_title = chapter.title
    sub_index = 0
    
    # 为了图片分配，我们需要跟踪每个子章节包含的图片引用
    img_pattern = re.compile(r'!\[[^\]]*\]\([^)]*images/([^)]+)\)')
    
    for line in lines:
        match = h3_pattern.match(line)
        if match:
            # 找到新的三级标题，保存前面的内容
            if current_lines:
                sub_chapter = create_sub_chapter_with_images(
                    start_index + sub_index, current_title, current_lines, chapter.images, img_pattern
                )
                sub_chapters.append(sub_chapter)
                sub_index += 1
            
            # 开始新的子章节
            current_title = f"{chapter.title} - {match.group(1)}"
            current_lines = [line]
        else:
            current_lines.append(line)
    
    # 处理最后一个子章节
    if current_lines:
        sub_chapter = create_sub_chapter_with_images(
            start_index + sub_index, current_title, current_lines, chapter.images, img_pattern
        )
        sub_chapters.append(sub_chapter)
    
    return sub_chapters if len(sub_chapters) > 1 else [chapter]


def create_sub_chapter_with_images(index: int, title: str, lines: list[str], all_images: list[dict], img_pattern) -> Chapter:
    """
    创建子章节并分配相关的图片
    
    Args:
        index: 章节索引
        title: 章节标题
        lines: 章节内容行
        all_images: 所有可用的图片
        img_pattern: 图片引用的正则表达式
        
    Returns:
        包含相关图片的章节对象
    """
    content = '\n'.join(lines)
    
    # 查找这个子章节中引用的图片
    referenced_images = []
    image_filenames = set()
    
    for match in img_pattern.finditer(content):
        img_filename = match.group(1).split('{')[0].strip()
        if img_filename not in image_filenames:
            image_filenames.add(img_filename)
            
            # 从所有图片中找到对应的图片数据
            for img in all_images:
                if img["filename"] == img_filename:
                    referenced_images.append(img)
                    break
    
    return Chapter(
        index=index,
        title=title,
        paragraphs=[content],
        images=referenced_images
    )


def split_document(filepath: str, heading_level: int = 1) -> list[Chapter]:
    """
    将 Word 文档按标题级别拆分为章节。

    Args:
        filepath: .docx 文件路径
        heading_level: 拆分的标题级别（1=按一级标题，2=按二级标题，以此类推）

    Returns:
        章节列表
    """
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"文件不存在: {filepath}")

    doc_dir = os.path.dirname(os.path.abspath(filepath))
    temp_dir = os.path.join(doc_dir, "_temp_pandoc")
    os.makedirs(temp_dir, exist_ok=True)
    html_path = os.path.join(temp_dir, "temp_export.htm")

    # 1. Export using Word COM
    print("正在通过 Word 后台导出高质量 HTML (自动剥离清晰图元)...")
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(os.path.abspath(filepath))
        # 10 = wdFormatFilteredHTML, 这会让 Word 导出极简 HTML，且原生渲染所有 MathType 公式为高清 GIF/PNG
        doc.SaveAs2(os.path.abspath(html_path), FileFormat=10)
        doc.Close()
        word.Quit()
    except Exception as e:
        print(f"Word COM 导出失败: {e}")
        return []

    # 2. Fix HTML encoding
    print("转换 HTML 字符编码以适配 Pandoc...")
    try:
        with open(html_path, 'r', encoding='gbk', errors='ignore') as f:
            html_content = f.read()
    except FileNotFoundError:
        with open(html_path, 'r', encoding='utf-8', errors='ignore') as f:
            html_content = f.read()
            
    html_utf8_path = os.path.join(temp_dir, "temp_export_utf8.htm")
    with open(html_utf8_path, 'w', encoding='utf-8') as f:
        f.write(html_content)

    # 3. Use pypandoc to generate clean markdown
    print("使用 Pandoc 将 HTML 转化为标准 Markdown，清理残余格式...")
    md_path = os.path.join(temp_dir, "temp.md")
    to_format = 'markdown-raw_html-bracketed_spans-native_divs-native_spans-fenced_divs-header_attributes-link_attributes'
    pypandoc.convert_file(
        html_utf8_path,
        to_format,
        format='html',
        outputfile=md_path
    )

    # 4. Read Markdown and split
    with open(md_path, 'r', encoding='utf-8') as f:
        full_md = f.read()
        
    lines = full_md.split('\n')
    
    # 使用智能分割策略，能够识别大章节边界  
    chapters = smart_split_chapters(lines, heading_level, temp_dir)
    
    # 清理临时目录
    try:
        shutil.rmtree(temp_dir, ignore_errors=True)
    except:
        pass
        
    return chapters


def smart_split_chapters(lines: list[str], heading_level: int, temp_dir: str) -> list[Chapter]:
    """
    智能分割章节，能够识别章节编号模式和结构
    """
    # 定义章节边界识别模式
    chapter_patterns = [
        r'^##?\s*前言',
        r'^##?\s*引\s*言',
        r'^##?\s*第?一章',
        r'^##?\s*第?二章', 
        r'^##?\s*第?三章',
        r'^##?\s*第?四章',
        r'^##?\s*第?五章',
        r'^##?\s*结\s*论',
        r'^##?\s*谢\s*辞',
        r'^##?\s*1\.',  # 1.
        r'^##?\s*2\.',  # 2.
        r'^##?\s*3\.',  # 3.
        r'^##?\s*4\.',  # 4. 
        r'^##?\s*5\.',  # 5.
    ]
    
    heading_pattern = re.compile(rf'^#{{{1,{heading_level}}}}\s+(.*)')
    chapter_boundary_pattern = re.compile('|'.join(chapter_patterns), re.IGNORECASE)
    
    chapters: list[Chapter] = []
    current_chapter = Chapter(index=0, title="前言")
    
    img_pattern = re.compile(r'!\[([^\]]*)\]\([^)]*temp_export\.files/([^)\s]+)[^)]*\)')
    files_dir = os.path.join(temp_dir, "temp_export.files")
    
    current_lines = []
    
    def finalize_chapter():
        if not current_lines:
            return
            
        md_text = "\n".join(current_lines)
        
        # 提取当前章节引用的图片源文件
        for match in img_pattern.finditer(md_text):
            img_filename = match.group(2).split('{')[0].strip()
            
            img_real_path = os.path.join(files_dir, img_filename)
            if os.path.exists(img_real_path):
                with open(img_real_path, 'rb') as img_f:
                    current_chapter.images.append({
                        "filename": img_filename,
                        "data": img_f.read()
                    })
                    
        # 替换 Markdown 中的原始图片链接为相对本地 links 并脱落多余 span 例如 {width=...}
        md_text = re.sub(
            r'!\[([^\]]*)\]\([^)]*temp_export\.files/([^)\s]+)[^)]*\)(?:\{[^\}]*\})?',
            r'![\1](images/\2)',
            md_text
        )
        
        # 不要添加额外分段，保持原始行
        for par in md_text.split("\n\n"):
            current_chapter.paragraphs.append(par)
            
        if current_chapter.text.strip() or current_chapter.has_images:
            chapters.append(current_chapter)
            
    for line in lines:
        # 检查是否是章节边界
        chapter_match = chapter_boundary_pattern.match(line)
        if chapter_match:
            finalize_chapter()
            
            # 提取标题
            title = re.sub(r'^##?\s*', '', line).strip()
            title = re.sub(r'\{#.*?\}', '', title).strip()
            
            current_chapter = Chapter(
                index=len(chapters),
                title=title
            )
            # 记录并清理标题行的 ID tag
            clean_line = re.sub(r'\{#.*?\}', '', line)
            current_lines = [clean_line]
        elif heading_level == 2 and line.startswith('##') and not line.startswith('###'):
            # 如果是二级标题分割，则按照二级标题分割
            finalize_chapter()
            
            title = line.replace('##', '').strip()
            title = re.sub(r'\{#.*?\}', '', title).strip()
            
            current_chapter = Chapter(
                index=len(chapters),
                title=title
            )
            clean_line = re.sub(r'\{#.*?\}', '', line)
            current_lines = [clean_line]
        else:
            current_lines.append(line)
            
    finalize_chapter()
    
    return chapters


def filter_preface_content(text: str) -> str:
    """
    过滤前言内容，只保留中英文摘要部分
    
    Args:
        text: 前言原始文本
        
    Returns:
        只包含摘要的文本
    """
    lines = text.split('\n')
    result_lines = []
    
    # 寻找摘要开始位置
    abstract_start = -1
    keywords_end = -1
    
    for i, line in enumerate(lines):
        # 找到中文摘要开始
        if re.match(r'^\s*摘\s*要\s*$', line.strip()):
            abstract_start = i
            break
    
    if abstract_start == -1:
        # 如果没找到摘要标题，尝试找含有摘要内容的段落
        for i, line in enumerate(lines):
            if "深度学习" in line and "卷积神经网络" in line:
                abstract_start = i - 1  # 包含前一行可能的标题
                break
    
    if abstract_start >= 0:
        # 找到英文关键词结束位置 - 更宽松的匹配
        for i, line in enumerate(lines[abstract_start:], start=abstract_start):
            # 查找英文关键词行后的下一个空行或目录开始
            if re.match(r'^Keywords?:', line.strip(), re.IGNORECASE):
                # 找到Keywords行，继续寻找该行完整内容的结束
                for j in range(i + 1, len(lines)):
                    next_line = lines[j].strip()
                    # 遇到目录开始或明显的非关键词内容就停止
                    if (next_line and ("目  录" in next_line or "Contents" in next_line or 
                        next_line.startswith("[引") or next_line.startswith("[第") or
                        next_line.startswith("引  言") or len(next_line) > 100)):
                        keywords_end = j - 1
                        break
                    elif j == len(lines) - 1:  # 到达文件末尾
                        keywords_end = j
                        break
                if keywords_end > 0:
                    break
        
        # 如果还没找到结束位置，寻找目录开始
        if keywords_end == -1:
            for i, line in enumerate(lines[abstract_start:], start=abstract_start):
                if ("目  录" in line or "Contents" in line or 
                    line.strip().startswith("[引") or line.strip().startswith("[第") or
                    line.strip().startswith("引  言")):
                    keywords_end = i - 1
                    break
        
        # 如果仍未找到，使用所有剩余内容
        if keywords_end == -1:
            keywords_end = len(lines) - 1
        
        if keywords_end >= abstract_start:
            result_lines = lines[abstract_start:keywords_end + 1]
    
    # 清理结果，移除空白行和格式干扰
    filtered_lines = []
    for line in result_lines:
        line = line.strip()
        # 遇到目录就停止
        if "目  录" in line or "目 录" in line or "Contents" in line:
            break
        # 跳过只包含图片引用、格式标记或明显目录项的行
        if (line and not re.match(r'^[-\s]*$', line) and "![" not in line and 
            not line.startswith('--') and not line.startswith('[') and
            not ("......" in line and "#_Toc" in line)):
            filtered_lines.append(line)
    
    # 后处理：确保以Keywords结尾，移除所有目录相关内容
    result_text = '\n\n'.join(filtered_lines)
    
    # 如果包含目录，在目录之前截断
    if "目  录" in result_text or "目 录" in result_text:
        # 找到Keywords后的位置，确保在目录之前结束
        keywords_pos = result_text.rfind("Keywords:")
        if keywords_pos > 0:
            # 从Keywords行开始向后找到完整的关键词列表结尾
            after_keywords = result_text[keywords_pos:]
            lines_after = after_keywords.split('\n')
            keywords_lines = []
            for line in lines_after:
                if "目" in line and "录" in line:
                    break
                keywords_lines.append(line)
            result_text = result_text[:keywords_pos] + '\n'.join(keywords_lines).strip()
    
    return result_text


def filter_references(text: str) -> str:
    """
    过滤掉文本中的参考文献部分
    
    Args:
        text: 原始文本
        
    Returns:
        移除参考文献后的文本
    """
    # 分行处理
    lines = text.split('\n')
    result_lines = []
    
    # 查找参考文献开始位置
    references_start = -1
    for i, line in enumerate(lines):
        # 匹配参考文献标题
        if re.match(r'^(参考文献|References?|Bibliography)$', line.strip(), re.IGNORECASE):
            references_start = i
            break
        # 匹配第一个引用条目
        if re.match(r'^\s*\[\d+\]\s+[A-Z]', line.strip()):
            references_start = i
            break
    
    if references_start >= 0:
        # 找到参考文献，只保留之前的内容
        result_lines = lines[:references_start]
    else:
        # 没有找到参考文献标题，检查是否有引用条目
        for i, line in enumerate(lines):
            if re.match(r'^\s*\[\d+\]\s+[A-Z]', line.strip()):
                # 找到第一个引用条目，在此处截断
                result_lines = lines[:i]
                break
        else:
            # 没有找到任何引用，保留所有内容
            result_lines = lines
    
    return '\n'.join(result_lines).strip()


def save_chapters(chapters: list[Chapter], output_dir: str) -> None:
    """
    将拆分后的章节保存到输出目录。

    目录结构：
        output_dir/
        ├── chapter_0_前言/
        │   ├── text.md
        │   └── images/
        ├── chapter_1_标题/
        │   ├── text.md
        │   └── images/
        └── ...
    """
    os.makedirs(output_dir, exist_ok=True)

    for chapter in chapters:
        # 清理标题中的非法文件名字符
        safe_title = re.sub(r'[\\/:*?"<>|]', '_', chapter.title)
        safe_title = safe_title[:50]  # 限制长度
        chapter_dir = os.path.join(output_dir, f"chapter_{chapter.index}_{safe_title}")
        os.makedirs(chapter_dir, exist_ok=True)

        # 保存文本
        text_path = os.path.join(chapter_dir, "text.md")
        with open(text_path, 'w', encoding='utf-8') as f:
            # 只有没有以 `# ` 开头时才补充大标题（以防止前言缺失标题）
            if not chapter.text.lstrip().startswith('#'):
                f.write(f"# {chapter.title}\n\n")
            f.write(chapter.text)

        # 保存图片
        if chapter.has_images:
            images_dir = os.path.join(chapter_dir, "images")
            os.makedirs(images_dir, exist_ok=True)
            for img in chapter.images:
                img_path = os.path.join(images_dir, img["filename"])
                with open(img_path, 'wb') as f:
                    f.write(img["data"])

    print(f"✅ 已保存 {len(chapters)} 个章节到 {output_dir}")