import os
import sys
import json
import io
import re

from dotenv import load_dotenv

load_dotenv()

# ============================================================
# Windows UTF-8
# ============================================================

os.environ["PYTHONIOENCODING"] = "utf-8"
os.environ["PYTHONLEGACYWINDOWSSTDIO"] = "utf-8"
os.environ["HTTPX_DEFAULT_ENCODING"] = "utf-8"

if sys.stdout.encoding != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")

if sys.stderr.encoding != "utf-8":
    sys.stderr.reconfigure(encoding="utf-8")


# ============================================================
# 第三方库
# ============================================================

import streamlit as st
from tavily import TavilyClient
from openai import OpenAI

import plotly.express as px
import pandas as pd

from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

from fpdf import FPDF


# ============================================================
# Streamlit 页面设置
# ============================================================

st.set_page_config(
    page_title="AI 调研助手",
    page_icon="🤖",
    layout="wide"
)

st.title("🤖 AI 调研报告生成器")

st.caption(
    "输入研究主题，Agent 自动搜索互联网、"
    "整理信息并生成带来源引用的调研报告"
)


# ============================================================
# API Client
# ============================================================

@st.cache_resource
def load_clients():

    tavily_api_key = os.getenv(
        "TAVILY_API_KEY"
    )

    deepseek_api_key = os.getenv(
        "DEEPSEEK_API_KEY"
    )

    if not tavily_api_key:

        st.error(
            "❌ 缺少 TAVILY_API_KEY，请检查 .env 文件。"
        )

        st.stop()

    if not deepseek_api_key:

        st.error(
            "❌ 缺少 DEEPSEEK_API_KEY，请检查 .env 文件。"
        )

        st.stop()

    tavily_client = TavilyClient(
        api_key=tavily_api_key
    )

    deepseek_client = OpenAI(
        api_key=deepseek_api_key,
        base_url="https://api.deepseek.com/v1"
    )

    return (
        tavily_client,
        deepseek_client
    )


tavily_client, deepseek_client = load_clients()


# ============================================================
# Citation / Source Tracking
# ============================================================

def add_sources(
    search_results,
    sources,
    source_url_map
):
    """
    将 Tavily 搜索结果转换成结构化 Source。

    每个 Source：

    {
        "id": 1,
        "title": "...",
        "url": "...",
        "content": "..."
    }

    同一个 URL 不重复保存。
    """

    for result in search_results:

        title = result.get(
            "title",
            "Untitled"
        )

        url = result.get(
            "url",
            ""
        )

        content = result.get(
            "content",
            ""
        )

        if not url:
            continue

        # URL 去重
        if url in source_url_map:
            continue

        source_id = len(sources) + 1

        source = {
            "id": source_id,
            "title": title,
            "url": url,
            "content": content
        }

        sources.append(source)

        source_url_map[url] = source_id

    return sources


# ============================================================
# 格式化 Sources
# ============================================================

def format_sources_for_llm(sources):

    formatted = []

    for source in sources:

        text = (
            f"[Source {source['id']}]\n"
            f"Title: {source['title']}\n"
            f"URL: {source['url']}\n"
            f"Content:\n"
            f"{source['content']}\n"
        )

        formatted.append(text)

    return "\n\n".join(formatted)


# ============================================================
# Citation Validation
# ============================================================

def validate_citations(
    report,
    sources
):
    """
    检查报告里的 [n] 是否对应真实 Source。
    """

    valid_ids = {
        source["id"]
        for source in sources
    }

    citations = re.findall(
        r"\[(\d+)\]",
        report
    )

    invalid_citations = []

    for citation in citations:

        citation_id = int(citation)

        if citation_id not in valid_ids:

            invalid_citations.append(
                citation_id
            )

    invalid_citations = sorted(
        set(invalid_citations)
    )

    cleaned_report = report

    # 删除非法引用
    for citation_id in invalid_citations:

        cleaned_report = cleaned_report.replace(
            f"[{citation_id}]",
            ""
        )

    return (
        cleaned_report,
        invalid_citations
    )


# ============================================================
# 添加参考来源
# ============================================================

def append_sources_to_report(
    report,
    sources
):

    if not sources:

        return report

    result = report

    result += "\n\n## 参考来源\n\n"

    for source in sources:

        result += (
            f"[{source['id']}] "
            f"{source['title']}\n"
            f"{source['url']}\n\n"
        )

    return result


# ============================================================
# Research Agent
# ============================================================

def run_agent(
    user_input,
    history,
    thought_container
):

    # ========================================================
    # Tool Definition
    # ========================================================

    tools = [

        {
            "type": "function",

            "function": {

                "name": "search_web",

                "description": (
                    "搜索互联网获取与研究任务相关的信息。"
                    "适用于事实、数据、新闻、市场信息、"
                    "行业趋势以及最新动态。"
                ),

                "parameters": {

                    "type": "object",

                    "properties": {

                        "query": {
                            "type": "string",
                            "description": "搜索关键词"
                        }

                    },

                    "required": [
                        "query"
                    ]
                }
            }
        }
    ]


    # ========================================================
    # System Prompt
    # ========================================================

    system_prompt = """
你是一个专业的 AI Research Agent。

你的任务是：

1. 理解用户研究问题
2. 使用搜索工具获取资料
3. 综合多个来源
4. 生成结构化中文调研报告
5. 对事实性内容进行来源引用

==============================
报告结构
==============================

第一次研究时，尽量使用：

## 概述

## 最新动态

## 主要观点

## 总结

如果用户的问题适合其他结构，
可以根据问题调整。

==============================
Citation 规则
==============================

所有来自搜索结果的关键事实，
应该在句子末尾添加对应的 Source ID。

例如：

新能源汽车销量继续增长。[1]

中国是全球最大的新能源汽车市场。[2]

如果一个结论由多个来源支持：

新能源汽车市场正在快速增长。[1][2]

==============================
严格要求
==============================

1. 只能使用实际存在的 Source ID。

2. 不允许编造来源。

3. 不允许使用不存在的引用。

4. Citation 格式只能是：

[1]
[2]
[3]

5. 不要使用：

(Source 1)

(来源1)

【1】

6. 如果搜索结果不足以支持一个结论，
   应该说明信息不足。

7. 如果需要更多资料，可以继续调用 search_web。

8. 当信息足够时停止搜索并生成最终报告。

==============================
多轮搜索
==============================

如果问题包含多个方面，
应该分别搜索不同维度。

例如：

用户问：
“比较 OpenAI、Google 和 Anthropic。”

可以分别搜索：

OpenAI
Google
Anthropic

然后进行综合分析。

==============================
最终回答
==============================

不要告诉用户你“准备生成报告”。

直接输出最终报告。

不要输出：

“让我整理一下。”

“我已经收集了资料。”

“下面我准备生成……”

直接输出正式报告。
"""


    # ========================================================
    # Messages
    # ========================================================

    messages = [

        {
            "role": "system",
            "content": system_prompt
        }

    ]


    # 加入历史对话
    for message in history:

        messages.append(
            {
                "role": message["role"],
                "content": message["content"]
            }
        )


    # 当前问题
    messages.append(
        {
            "role": "user",
            "content": user_input
        }
    )


    # ========================================================
    # Sources
    # ========================================================

    sources = []

    source_url_map = {}

    thought_logs = []

    step = 1


    # ========================================================
    # Agent Loop
    # ========================================================

    for _ in range(8):

        try:

            response = (
                deepseek_client
                .chat
                .completions
                .create(

                    model="deepseek-chat",

                    messages=messages,

                    tools=tools,

                    tool_choice="auto"
                )
            )

        except Exception as e:

            return (
                "❌ LLM 调用失败。\n\n"
                f"{str(e)}"
            )


        msg = response.choices[0].message


        # ====================================================
        # Tool Call
        # ====================================================

        if msg.tool_calls:

            messages.append(

                {
                    "role": "assistant",

                    "content": msg.content or "",

                    "tool_calls": [

                        {
                            "id": tc.id,

                            "type": "function",

                            "function": {

                                "name": tc.function.name,

                                "arguments": (
                                    tc.function.arguments
                                )
                            }
                        }

                        for tc in msg.tool_calls
                    ]
                }
            )


            for tool_call in msg.tool_calls:

                # --------------------------------------------
                # 参数解析
                # --------------------------------------------

                try:

                    args = json.loads(
                        tool_call.function.arguments
                    )

                    query = args["query"]

                except Exception:

                    thought_logs.append(
                        f"**第{step}步** ⚠️ "
                        "搜索参数解析失败"
                    )

                    thought_container.markdown(
                        "\n\n".join(
                            thought_logs
                        )
                    )

                    messages.append(

                        {
                            "role": "tool",

                            "tool_call_id": (
                                tool_call.id
                            ),

                            "content": (
                                "搜索参数格式错误，"
                                "请重新生成 query。"
                            )
                        }
                    )

                    step += 1

                    continue


                # --------------------------------------------
                # 显示搜索
                # --------------------------------------------

                thought_logs.append(
                    f"**第{step}步** 🔍 "
                    f"搜索关键词：`{query}`"
                )

                thought_container.markdown(
                    "\n\n".join(
                        thought_logs
                    )
                )


                # --------------------------------------------
                # Tavily Search
                # --------------------------------------------

                try:

                    search_result = (
                        tavily_client.search(

                            query=query,

                            max_results=3
                        )
                    )

                    results = search_result.get(
                        "results",
                        []
                    )

                except Exception as e:

                    thought_logs.append(
                        f"　　❌ 搜索失败：{str(e)}"
                    )

                    thought_container.markdown(
                        "\n\n".join(
                            thought_logs
                        )
                    )

                    messages.append(

                        {
                            "role": "tool",

                            "tool_call_id": (
                                tool_call.id
                            ),

                            "content": (
                                "搜索失败。"
                                "请根据已有资料继续，"
                                "或者尝试其他搜索关键词。"
                            )
                        }
                    )

                    step += 1

                    continue


                # --------------------------------------------
                # Source Tracking
                # --------------------------------------------

                old_count = len(
                    sources
                )

                sources = add_sources(
                    results,
                    sources,
                    source_url_map
                )

                new_count = (
                    len(sources)
                    - old_count
                )


                thought_logs.append(
                    f"　　✅ 获取到 "
                    f"{len(results)} 条搜索结果，"
                    f"新增 {new_count} 个来源"
                )

                thought_container.markdown(
                    "\n\n".join(
                        thought_logs
                    )
                )


                # --------------------------------------------
                # 给 LLM 的 Observation
                # --------------------------------------------

                source_text = (
                    format_sources_for_llm(
                        sources
                    )
                )

                messages.append(

                    {
                        "role": "tool",

                        "tool_call_id": (
                            tool_call.id
                        ),

                        "content": (
                            "以下是目前收集到的全部来源：\n\n"
                            f"{source_text}\n\n"
                            "请判断是否需要继续搜索。"
                        )
                    }
                )

                step += 1


        # ====================================================
        # Final Answer
        # ====================================================

        else:

            thought_logs.append(
                f"**第{step}步** 📝 "
                "生成最终报告..."
            )

            thought_container.markdown(
                "\n\n".join(
                    thought_logs
                )
            )


            report = (
                msg.content
                or ""
            )


            # =================================================
            # Citation Validation
            # =================================================

            (
                report,
                invalid_citations
            ) = validate_citations(
                report,
                sources
            )


            if invalid_citations:

                thought_logs.append(
                    "⚠️ 删除非法 Citation："
                    + ", ".join(
                        f"[{x}]"
                        for x in invalid_citations
                    )
                )

                thought_container.markdown(
                    "\n\n".join(
                        thought_logs
                    )
                )


            # =================================================
            # 添加来源
            # =================================================

            report = append_sources_to_report(
                report,
                sources
            )


            return report


    # ========================================================
    # 超过最大步数
    # ========================================================

    return (
        "⚠️ Agent 搜索次数达到上限，"
        "暂时无法生成完整报告。"
    )


# ============================================================
# Chart Extraction
# ============================================================

def extract_chart_data(report):

    prompt = f"""
请从下面的调研报告中提取可以进行数据可视化的数据。

要求：

1. 只能使用报告中明确存在的数据。
2. 不允许编造数字。
3. 最多提取 3 个图表。
4. 如果报告没有可靠数据，返回空 charts。
5. 只返回 JSON。

格式：

{{
    "charts": [
        {{
            "title": "图表标题",
            "type": "bar",
            "data": [
                {{
                    "label": "名称",
                    "value": 10
                }}
            ]
        }}
    ]
}}

报告：

{report}
"""


    try:

        response = (
            deepseek_client
            .chat
            .completions
            .create(

                model="deepseek-chat",

                messages=[
                    {
                        "role": "user",
                        "content": prompt
                    }
                ],

                temperature=0.2
            )
        )


        text = (
            response
            .choices[0]
            .message
            .content
            or ""
        )


        text = (
            text
            .replace(
                "```json",
                ""
            )
            .replace(
                "```",
                ""
            )
            .strip()
        )


        try:

            data = json.loads(
                text
            )

        except json.JSONDecodeError:

            match = re.search(
                r"\{.*\}",
                text,
                re.DOTALL
            )

            if not match:

                return {
                    "charts": []
                }

            data = json.loads(
                match.group(0)
            )


        if (
            not isinstance(data, dict)
            or "charts" not in data
        ):

            return {
                "charts": []
            }


        return data


    except Exception:

        return {
            "charts": []
        }


# ============================================================
# Markdown 文本解析（用于 Word / PDF 导出）
# ============================================================

def parse_inline_markdown(text):
    """
    将简单 Markdown 转换为可用于 Word / PDF 的文本片段。

    支持：
    - **粗体**
    - `行内代码`
    - 普通文本

    Citation 如 [1][2] 会原样保留。
    """
    parts = []
    pattern = re.compile(r"(\*\*.*?\*\*|`.*?`)")
    last = 0

    for match in pattern.finditer(text):
        if match.start() > last:
            parts.append({
                "text": text[last:match.start()],
                "bold": False,
                "code": False
            })

        token = match.group(0)

        if token.startswith("**") and token.endswith("**"):
            parts.append({
                "text": token[2:-2],
                "bold": True,
                "code": False
            })
        elif token.startswith("`") and token.endswith("`"):
            parts.append({
                "text": token[1:-1],
                "bold": False,
                "code": True
            })

        last = match.end()

    if last < len(text):
        parts.append({
            "text": text[last:],
            "bold": False,
            "code": False
        })

    if not parts:
        parts.append({
            "text": text,
            "bold": False,
            "code": False
        })

    return parts


def clean_markdown_line(line):
    """
    清理导出到文档中的 Markdown 控制符。
    """
    line = line.strip()

    # Markdown 分隔线
    if re.fullmatch(r"[-*_]{3,}", line):
        return "separator", ""

    # 标题
    if line.startswith("### "):
        return "h3", line[4:].strip()

    if line.startswith("## "):
        return "h2", line[3:].strip()

    if line.startswith("# "):
        return "h1", line[2:].strip()

    # 无序列表
    if re.match(r"^[-*]\s+", line):
        return "bullet", re.sub(r"^[-*]\s+", "", line)

    # 有序列表
    if re.match(r"^\d+\.\s+", line):
        return "number", line

    return "body", line


# ============================================================
# Word / PDF Export
# ============================================================

def _clean_report_line(line):
    """清理报告中的 Markdown 标记，同时保留标题层级识别所需的信息。"""
    return line.strip()


def _strip_inline_markdown(text):
    text = re.sub(r"\*\*(.*?)\*\*", r"\1", text)
    text = re.sub(r"__(.*?)__", r"\1", text)
    text = re.sub(r"`([^`]*)`", r"\1", text)
    return text


def _add_word_rich_paragraph(doc, text, style=None, first_line_indent=True):
    """添加支持 **加粗** 和普通文本的 Word 段落。"""
    p = doc.add_paragraph(style=style)

    if first_line_indent:
        p.paragraph_format.first_line_indent = Pt(24)  # 约 2 个中文字符
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.5

    # 简单解析 **粗体**
    parts = re.split(r"(\*\*.*?\*\*)", text)
    for part in parts:
        if not part:
            continue
        run = p.add_run(_strip_inline_markdown(part))
        if part.startswith("**") and part.endswith("**"):
            run.bold = True
        run.font.name = "仿宋_GB2312"
        run.font.size = Pt(16)

    return p


def export_word(topic, report):
    """
    按正式报告格式生成 Word：
    A4；上下 2.54 cm；左右 2.5 cm；
    正文仿宋三号、1.5 倍行距、首行缩进 2 字符；
    一级标题黑体三号；
    二级标题楷体_GB2312三号；
    三级标题仿宋_GB2312三号；
    页码底部居中。
    """
    doc = Document()
    section = doc.sections[0]

    # A4 页面 + 页边距
    section.page_width = Inches(8.27)
    section.page_height = Inches(11.69)
    section.top_margin = Inches(2.54 / 2.54)
    section.bottom_margin = Inches(2.54 / 2.54)
    section.left_margin = Inches(2.5 / 2.54)
    section.right_margin = Inches(2.5 / 2.54)

    # 默认 Normal 样式
    normal = doc.styles["Normal"]
    normal.font.name = "仿宋_GB2312"
    normal.font.size = Pt(16)
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(0)
    normal.paragraph_format.line_spacing = 1.5
    normal.paragraph_format.first_line_indent = Pt(24)

    # 封面
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(120)
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run(f"调研报告：{topic}")
    run.bold = True
    run.font.name = "黑体"
    run.font.size = Pt(22)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(180)
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run("AI 调研报告")
    run.font.name = "黑体"
    run.font.size = Pt(15)

    # 分页进入正文
    doc.add_page_break()

    for raw_line in report.split("\n"):
        line = _clean_report_line(raw_line)

        if not line:
            continue

        # Markdown 标题 → 正式报告标题层级
        if line.startswith("# "):
            title = _strip_inline_markdown(line[2:].strip())
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.5
            run = p.add_run(title)
            run.bold = True
            run.font.name = "黑体"
            run.font.size = Pt(16)
            continue

        if line.startswith("## "):
            title = _strip_inline_markdown(line[3:].strip())
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.5
            run = p.add_run(title)
            run.bold = True
            run.font.name = "黑体"
            run.font.size = Pt(16)
            continue

        if line.startswith("### "):
            title = _strip_inline_markdown(line[4:].strip())
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.5
            run = p.add_run(title)
            run.font.name = "仿宋_GB2312"
            run.font.size = Pt(16)
            continue

        # 正式报告中常见的层级标题
        if re.match(r"^一、|^二、|^三、|^四、|^五、|^六、|^七、|^八、|^九、|^十、", line):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.5
            run = p.add_run(_strip_inline_markdown(line))
            run.bold = True
            run.font.name = "黑体"
            run.font.size = Pt(16)
            continue

        if re.match(r"^（[一二三四五六七八九十]+）", line):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.5
            run = p.add_run(_strip_inline_markdown(line))
            run.font.name = "楷体_GB2312"
            run.font.size = Pt(16)
            continue

        # Markdown 列表
        bullet = re.match(r"^[-*]\s+(.*)", line)
        numbered = re.match(r"^\d+[.)]\s+(.*)", line)
        if bullet:
            p = _add_word_rich_paragraph(
                doc, bullet.group(1), first_line_indent=False
            )
            p.paragraph_format.left_indent = Pt(24)
            p.style = doc.styles["Normal"]
            continue

        if numbered:
            p = _add_word_rich_paragraph(
                doc, numbered.group(1), first_line_indent=False
            )
            p.paragraph_format.left_indent = Pt(24)
            continue

        # 分隔线
        if re.fullmatch(r"[-*_]{3,}", line):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.add_run("────────────────────────")
            continue

        # 普通正文
        _add_word_rich_paragraph(doc, line, first_line_indent=True)

    # 页码：底部居中
    footer = section.footer
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("页码")
    run.font.name = "宋体"
    run.font.size = Pt(9)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


def export_pdf(topic, report):
    """生成中文 PDF，安全处理长 URL 和长连续字符串。"""

    class PDF(FPDF):
        def footer(self):
            self.set_y(-12)
            self.set_font("CJK", "", 9)
            self.cell(self.epw, 6, str(self.page_no()), align="C")

    font_candidates = [
        r"C:\Windows\Fonts\msyh.ttf",
        r"C:\Windows\Fonts\simhei.ttf",
        r"C:\Windows\Fonts\simfang.ttf",
        r"C:\Windows\Fonts\simsun.ttf",
        r"C:\Windows\Fonts\simsun.ttc",
    ]

    regular_font = next(
        (path for path in font_candidates if os.path.exists(path)),
        None
    )

    if regular_font is None:
        st.error("❌ 找不到 Windows 中文字体，PDF 暂时无法生成。")
        return None

    pdf = PDF(orientation="P", unit="mm", format="A4")
    pdf.set_margins(20, 20, 20)
    pdf.set_auto_page_break(auto=True, margin=18)

    try:
        pdf.add_font("CJK", "", regular_font)
        pdf.add_font("CJK", "B", regular_font)
    except Exception as e:
        st.error(f"❌ 中文字体加载失败：{e}")
        return None

    pdf.add_page()

    content_width = pdf.w - pdf.l_margin - pdf.r_margin

    def clean_pdf_text(text):
        text = _strip_inline_markdown(text)
        text = text.replace("**", "")
        text = text.replace("__", "")
        text = text.replace("`", "")
        return text.strip()

    def safe_wrap_long_runs(text, chunk_size=20):
        url_pattern = re.compile(r"https?://[^\s]+")

        def wrap_url(match):
            value = match.group(0)
            return "\n".join(
                value[i:i + chunk_size]
                for i in range(0, len(value), chunk_size)
            )

        text = url_pattern.sub(wrap_url, text)

        def wrap_token(match):
            value = match.group(0)
            if len(value) <= chunk_size:
                return value
            return "\n".join(
                value[i:i + chunk_size]
                for i in range(0, len(value), chunk_size)
            )

        return re.sub(
            r"[A-Za-z0-9@._:/?=&%+\-]+",
            wrap_token,
            text
        )

    report_lines = report.split("\n")
    first_heading = None

    for i, raw in enumerate(report_lines):
        stripped = raw.strip()
        if (
            stripped.startswith("# ")
            or stripped.startswith("## ")
            or stripped.startswith("### ")
            or re.match(r"^[一二三四五六七八九十]+、", stripped)
        ):
            first_heading = i
            break

    if first_heading is not None and first_heading > 0:
        report_lines = report_lines[first_heading:]

    pdf.set_font("CJK", "B", 18)

    title = clean_pdf_text(f"调研报告：{topic}")
    title = safe_wrap_long_runs(title, 16)

    pdf.multi_cell(
        content_width,
        10,
        title,
        align="C",
        wrapmode="CHAR"
    )
    pdf.ln(6)

    for raw_line in report_lines:
        line = raw_line.strip()

        if not line:
            pdf.ln(4)
            continue

        if line.startswith("# ") or line.startswith("## "):
            title = clean_pdf_text(line.split(" ", 1)[1])
            title = safe_wrap_long_runs(title, 18)
            pdf.set_font("CJK", "B", 16)
            pdf.multi_cell(
                content_width,
                9,
                title,
                wrapmode="CHAR"
            )
            pdf.ln(2)
            continue

        if line.startswith("### "):
            title = clean_pdf_text(line[4:])
            title = safe_wrap_long_runs(title, 18)
            pdf.set_font("CJK", "B", 14)
            pdf.multi_cell(
                content_width,
                8,
                title,
                wrapmode="CHAR"
            )
            pdf.ln(1)
            continue

        bullet = re.match(r"^[-*]\s+(.*)", line)
        numbered = re.match(r"^(\d+[.)])\s+(.*)", line)

        if bullet:
            line = "• " + clean_pdf_text(bullet.group(1))
        elif numbered:
            line = numbered.group(1) + " " + clean_pdf_text(numbered.group(2))
        else:
            line = clean_pdf_text(line)

        if re.match(r"^[一二三四五六七八九十]+、", line):
            pdf.set_font("CJK", "B", 15)
        elif re.match(r"^（[一二三四五六七八九十]+）", line):
            pdf.set_font("CJK", "B", 14)
        else:
            pdf.set_font("CJK", "", 11)

        line = safe_wrap_long_runs(line, 20)

        try:
            pdf.multi_cell(
                content_width,
                7,
                line,
                wrapmode="CHAR"
            )
        except Exception:
            chunks = [
                line[i:i + 15]
                for i in range(0, len(line), 15)
            ]
            for chunk in chunks:
                pdf.multi_cell(
                    content_width,
                    7,
                    chunk,
                    wrapmode="CHAR"
                )

    try:
        return bytes(pdf.output())
    except Exception as e:
        st.error(f"❌ PDF 生成失败：{e}")
        return None

# ============================================================
# Session State
# ============================================================

if "messages" not in st.session_state:

    st.session_state.messages = []


if "last_report" not in st.session_state:

    st.session_state.last_report = None


# ============================================================
# 显示历史消息
# ============================================================

for message in st.session_state.messages:

    with st.chat_message(
        message["role"]
    ):

        st.markdown(
            message["content"]
        )


# ============================================================
# 用户输入
# ============================================================

user_input = st.chat_input(
    "请输入你想研究的问题，例如：2025年全球新能源汽车市场发展情况"
)


if user_input:

    # ========================================================
    # 保存用户消息
    # ========================================================

    st.session_state.messages.append(

        {
            "role": "user",
            "content": user_input
        }
    )


    with st.chat_message(
        "user"
    ):

        st.markdown(
            user_input
        )


    # ========================================================
    # Agent
    # ========================================================

    with st.chat_message(
        "assistant"
    ):

        thought_container = st.empty()


        with st.spinner(
            "🤖 Agent 正在进行调研..."
        ):

            history = (
                st.session_state.messages[:-1]
            )

            report = run_agent(

                user_input,

                history,

                thought_container
            )


        st.markdown(
            report
        )


    # ========================================================
    # 保存报告
    # ========================================================

    st.session_state.messages.append(

        {
            "role": "assistant",
            "content": report
        }
    )


    st.session_state.last_report = report


    # ========================================================
    # 数据可视化
    # ========================================================

    with st.spinner(
        "📊 正在分析报告中的数据..."
    ):

        chart_data = extract_chart_data(
            report
        )


    charts = chart_data.get(
        "charts",
        []
    )


    if charts:

        st.subheader(
            "📊 数据可视化"
        )


        for chart in charts:

            data = chart.get(
                "data",
                []
            )

            if not data:

                continue


            df = pd.DataFrame(
                data
            )


            if (
                "label" not in df.columns
                or "value" not in df.columns
            ):

                continue


            chart_type = chart.get(
                "type",
                "bar"
            )


            title = chart.get(
                "title",
                "数据图表"
            )


            if chart_type == "pie":

                fig = px.pie(

                    df,

                    names="label",

                    values="value",

                    title=title
                )

            else:

                fig = px.bar(

                    df,

                    x="label",

                    y="value",

                    title=title
                )


            st.plotly_chart(
                fig,
                use_container_width=True
            )

# ============================================================
# 导出
# ============================================================
# 防止 Streamlit 首次运行时访问未初始化的 session_state
if "last_report" not in st.session_state:
    st.session_state.last_report = ""



if st.session_state.get("last_report", ""):

    st.divider()

    st.subheader(
        "📥 导出报告"
    )


    col1, col2 = st.columns(2)


    # ========================================================
    # Word
    # ========================================================

    with col1:

        try:

            word_file = export_word(

                "AI 调研报告",

                st.session_state.last_report
            )


            st.download_button(

                label="📄 下载 Word",

                data=word_file,

                file_name="AI_调研报告.docx",

                mime=(
                    "application/vnd.openxmlformats-"
                    "officedocument.wordprocessingml.document"
                )
            )

        except Exception as e:

            st.error(
                f"Word 导出失败：{str(e)}"
            )


    # ========================================================
    # PDF
    # ========================================================

    with col2:

        try:

            pdf_file = export_pdf(

                "AI 调研报告",

                st.session_state.last_report
            )


            if pdf_file:

                st.download_button(

                    label="📕 下载 PDF",

                    data=pdf_file,

                    file_name="AI_调研报告.pdf",

                    mime="application/pdf"
                )

        except Exception as e:

            st.error(
                f"PDF 导出失败：{str(e)}"
            )


# ============================================================
# 清空对话
# ============================================================

st.divider()


if st.button(
    "🗑️ 清空对话"
):

    st.session_state.messages = []

    st.session_state.last_report = None

    st.rerun()