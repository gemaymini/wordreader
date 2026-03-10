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
        """获取章节纯文本，段落之间用换行分隔"""
        return "\n\n".join(p for p in self.paragraphs if p.strip())

    @property
    def has_images(self) -> bool:
        return len(self.images) > 0


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
    heading_pattern = re.compile(rf'^#{{1,{heading_level}}}\s+(.*)')
    
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
        match = heading_pattern.match(line)
        if match:
            finalize_chapter()
            
            title = match.group(1)
            # 清理 Pandoc id tags 比如 `{#_id_1}`
            title = re.sub(r'\{#.*?\}', '', title).strip()
            
            current_chapter = Chapter(
                index=len(chapters),
                title=title
            )
            # 记录并清理标题行的 ID tag
            clean_line = re.sub(r'\{#.*?\}', '', line)
            current_lines = [clean_line]
        else:
            current_lines.append(line)
            
    finalize_chapter()
    
    # 清理临时目录
    try:
        shutil.rmtree(temp_dir, ignore_errors=True)
    except:
        pass
        
    return chapters


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
