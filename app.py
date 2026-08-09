import os
import sys
from dotenv import load_dotenv
load_dotenv()

os.environ["PYTHONIOENCODING"] = "utf-8"
os.environ["PYTHONLEGACYWINDOWSSTDIO"] = "utf-8"
os.environ["HTTPX_DEFAULT_ENCODING"] = "utf-8"
if sys.stdout.encoding != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")
if sys.stderr.encoding != "utf-8":
    sys.stderr.reconfigure(encoding="utf-8")

# ========== 正式导入 ==========
import streamlit as st
from tavily import TavilyClient
from openai import OpenAI
import json
import plotly.express as px
import pandas as pd
import io
from docx import Document
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# ========== 页面设置 ==========
st.set_page_config(page_title="AI 调研助手", page_icon="🤖")
st.title("🤖 AI 调研报告生成器")
st.caption("输入任意主题，Agent 自动搜索并生成结构化报告，支持追问")

# ========== 初始化客户端 ==========
@st.cache_resource
def load_clients():
    tavily = TavilyClient(api_key=os.getenv("TAVILY_API_KEY"))
    deepseek = OpenAI(
        api_key=os.getenv("DEEPSEEK_API_KEY"),
        base_url="https://api.deepseek.com/v1"
    )
    return tavily, deepseek

tavily_client, deepseek_client = load_clients()

# ========== Agent 核心逻辑 ==========
def run_agent(user_input, history, thought_container):
    tools = [
        {
            "type": "function",
            "function": {
                "name": "search_web",
                "description": "搜索网络获取最新信息",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "query": {"type": "string", "description": "搜索关键词"}
                    },
                    "required": ["query"]
                }
            }
        }
    ]

    messages = [
        {"role": "system", "content": """你是一个专业调研助手。
第一次调研时生成包含概述、最新动态、主要观点、总结四个部分的中文报告。
追问时根据对话历史和必要的搜索来回答用户问题。"""}
    ]

    for msg in history:
        messages.append({"role": msg["role"], "content": msg["content"]})
    messages.append({"role": "user", "content": user_input})

    thought_logs = []
    step = 1

    for i in range(8):
        response = deepseek_client.chat.completions.create(
            model="deepseek-chat",
            messages=messages,
            tools=tools,
            tool_choice="auto"
        )

        msg = response.choices[0].message

        if msg.tool_calls:
            messages.append({
                "role": "assistant",
                "content": msg.content or "",
                "tool_calls": [
                    {
                        "id": tc.id,
                        "type": "function",
                        "function": {
                            "name": tc.function.name,
                            "arguments": tc.function.arguments
                        }
                    }
                    for tc in msg.tool_calls
                ]
            })

            for tool_call in msg.tool_calls:
                args = json.loads(tool_call.function.arguments)
                query = args["query"]

                log = f"**第{step}步** 🔍 搜索关键词：`{query}`"
                thought_logs.append(log)
                thought_container.markdown("\n\n".join(thought_logs))
                step += 1

                search_result = tavily_client.search(query=query, max_results=3)
                content = "\n".join([r["content"] for r in search_result["results"]])

                log2 = f"　　✅ 获取到 {len(search_result['results'])} 条相关结果"
                thought_logs.append(log2)
                thought_container.markdown("\n\n".join(thought_logs))

                messages.append({
                    "role": "tool",
                    "tool_call_id": tool_call.id,
                    "content": content
                })
        else:
            if thought_logs:
                thought_logs.append(f"**第{step}步** 📝 生成回答...")
                thought_container.markdown("\n\n".join(thought_logs))
            return msg.content

    return "未能生成完整回答，请重试。"

# ========== 提取图表数据 ==========
def extract_chart_data(report):
    prompt = f"""从以下调研报告中提取可以可视化的数据，返回 JSON 格式。
要求：
1. 提取2-3组数据，每组包含名称和数值
2. 数值必须是数字（可以是估算值）
3. 只返回 JSON，不要其他文字

格式：
{{
  "charts": [
    {{
      "title": "图表标题",
      "type": "pie",
      "data": [{{"label": "名称", "value": 数字}}, ...]
    }},
    {{
      "title": "图表标题",
      "type": "bar",
      "data": [{{"label": "名称", "value": 数字}}, ...]
    }}
  ]
}}

报告内容：
{report}"""

    response = deepseek_client.chat.completions.create(
        model="deepseek-chat",
        messages=[{"role": "user", "content": prompt}],
        temperature=0.3
    )

    text = response.choices[0].message.content
    text = text.replace("```json", "").replace("```", "").strip()
    return json.loads(text)

def render_charts(chart_data):
    for chart in chart_data.get("charts", []):
        df = pd.DataFrame(chart["data"])
        if chart["type"] == "pie":
            fig = px.pie(df, names="label", values="value", title=chart["title"])
        else:
            fig = px.bar(df, x="label", y="value", title=chart["title"])
        st.plotly_chart(fig, use_container_width=True)

# ========== 导出 Word ==========
def export_word(topic, report):
    doc = Document()
    doc.add_heading(f"调研报告：{topic}", 0)
    for line in report.split("\n"):
        if line.startswith("## ") or line.startswith("# "):
            doc.add_heading(line.replace("#", "").strip(), level=1)
        elif line.startswith("### "):
            doc.add_heading(line.replace("#", "").strip(), level=2)
        elif line.strip():
            doc.add_paragraph(line.strip())
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# ========== 导出 PDF ==========
def export_pdf(topic, report):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    width, height = A4

    font_paths = [
        "C:/Windows/Fonts/simhei.ttf",
        "C:/Windows/Fonts/msyh.ttc",
        "C:/Windows/Fonts/simsun.ttc"
    ]
    font_name = "Helvetica"
    for path in font_paths:
        if os.path.exists(path):
            try:
                pdfmetrics.registerFont(TTFont("ChineseFont", path))
                font_name = "ChineseFont"
                break
            except Exception:
                continue

    y = height - 2 * cm
    c.setFont(font_name, 16)
    c.drawString(2 * cm, y, f"调研报告：{topic}")
    y -= 1 * cm

    c.setFont(font_name, 11)
    for line in report.split("\n"):
        if y < 2 * cm:
            c.showPage()
            c.setFont(font_name, 11)
            y = height - 2 * cm
        line = line.replace("#", "").strip()
        if line:
            c.drawString(2 * cm, y, line[:60])
            y -= 0.6 * cm

    c.save()
    buf.seek(0)
    return buf

# ========== Session 初始化 ==========
if "chat_history" not in st.session_state:
    st.session_state.chat_history = []
if "chart_data" not in st.session_state:
    st.session_state.chart_data = None
if "first_report" not in st.session_state:
    st.session_state.first_report = None
if "first_topic" not in st.session_state:
    st.session_state.first_topic = None

# ========== 侧边栏 ==========
with st.sidebar:
    st.header("⚙️ 操作")
    show_charts = st.toggle("📊 生成可视化图表", value=True)

    if st.session_state.first_report:
        st.divider()
        st.header("📥 导出报告")
        word_buf = export_word(st.session_state.first_topic, st.session_state.first_report)
        st.download_button(
            label="📄 下载 Word",
            data=word_buf,
            file_name=f"{st.session_state.first_topic}_报告.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        pdf_buf = export_pdf(st.session_state.first_topic, st.session_state.first_report)
        st.download_button(
            label="📕 下载 PDF",
            data=pdf_buf,
            file_name=f"{st.session_state.first_topic}_报告.pdf",
            mime="application/pdf"
        )

    st.divider()
    if st.button("🗑️ 清空对话"):
        st.session_state.chat_history = []
        st.session_state.chart_data = None
        st.session_state.first_report = None
        st.session_state.first_topic = None
        st.rerun()

# ========== 显示历史对话 + 图表 ==========
for i, msg in enumerate(st.session_state.chat_history):
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])
    if msg["role"] == "assistant" and i == 1 and st.session_state.chart_data:
        st.markdown("### 📊 数据可视化")
        render_charts(st.session_state.chart_data)

# ========== 输入框 ==========
user_input = st.chat_input("输入调研主题或追问...")

if user_input:
    st.session_state.chat_history.append({"role": "user", "content": user_input})
    with st.chat_message("user"):
        st.markdown(user_input)

    with st.expander("🧠 Agent 思考过程", expanded=True):
        thought_container = st.empty()

    with st.chat_message("assistant"):
        with st.spinner("调研中..."):
            result = run_agent(
                user_input,
                st.session_state.chat_history[:-1],
                thought_container
            )
        st.markdown(result)

        if len(st.session_state.chat_history) == 1:
            st.session_state.first_report = result
            st.session_state.first_topic = user_input
            if show_charts:
                with st.spinner("正在生成可视化图表..."):
                    try:
                        chart_data = extract_chart_data(result)
                        st.session_state.chart_data = chart_data
                        st.markdown("### 📊 数据可视化")
                        render_charts(chart_data)
                    except Exception:
                        st.info("本次报告暂无可提取的结构化数据")

    st.session_state.chat_history.append({"role": "assistant", "content": result})
