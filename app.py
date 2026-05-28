import streamlit as st
from openai import OpenAI
import base64
import requests
import time
from pdf2image import convert_from_bytes
from io import BytesIO
import dashscope
from dashscope import MultiModalConversation
from dashscope import Generation
from dashscope.aigc.image_generation import ImageGeneration
from dashscope.api_entities.dashscope_response import Message
import re
import jieba
from pypinyin import pinyin, Style
from PIL import Image, ImageDraw, ImageFont
import io
import platform
from xhtml2pdf import pisa
import os
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
import re
from weasyprint import HTML

import random
import matplotlib.pyplot as plt
from matplotlib.textpath import TextPath
from matplotlib.font_manager import FontProperties
from svglib.svglib import svg2rlg

import PyPDF2
from fpdf import FPDF
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from fpdf.enums import XPos, YPos
import urllib.parse
import json
from gtts import gTTS
import tempfile
import textwrap
import unicodedata

api_key = 'sk-9f798f6251c042adb32074d93454e201'
dashscope.base_http_api_url = 'https://dashscope-intl.aliyuncs.com/api/v1'
st.set_page_config(page_title="SenSpark", page_icon="🌱", layout="wide")

# font_path = r"C:\Users\cherr\Desktop\Master\DOTE6688I - Start Up\text_to_image\Full Version\V2\SourceHanSansHC-VF.ttf"
# font_path = r"/Users/administrator/Desktop/text_to_image/SourceHanSansHC-VF.ttf"
# Get the directory where app.py is located
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Link directly to the font file in that folder
font_path = os.path.join(BASE_DIR, "SourceHanSansHC-VF.ttf")
font_path_2 = os.path.join(BASE_DIR, "NotoSansTC-Regular.ttf")
PDF_FONT_NAME = "Helvetica"
font_loaded = False

try:
    pdfmetrics.registerFont(TTFont("SourceHanSans", font_path))
    PDF_FONT_NAME = "SourceHanSans"
    font_loaded = True
except Exception as e:
    font_loaded = False

if 'chat_history' not in st.session_state:
    st.session_state.chat_history = []
if 'adjustment_instruction' not in st.session_state:
    st.session_state.adjustment_instruction = ""
if 'page' not in st.session_state:
    st.session_state.page = 'home'
if 'math_results' not in st.session_state:
    st.session_state.math_results = [] # 用於存儲數學解析結果
if 'trigger_analysis' not in st.session_state:
    st.session_state.trigger_analysis = False
if 'trigger_generate' not in st.session_state:
    st.session_state.trigger_generate = False
if "show_wizard" not in st.session_state:
    st.session_state.show_wizard = False
if "step" not in st.session_state:
    st.session_state.step = 1
if "selections" not in st.session_state:
    st.session_state.selections = {}
if "generated_questions" not in st.session_state:
    st.session_state.generated_questions = None
if "selected_theme" not in st.session_state:
    st.session_state.selected_theme = "Ocean Adventure"
if 'my_tabs' not in st.session_state:
    st.session_state.my_tabs = "📝 Upload math questions"
    st.session_state.current_mode = "analysis_mode" # 你想自動賦予的附加值

CSS_STYLE = """
<style>
    @media print {
        .no-print { display: none; }
        body { margin: 0; padding: 0; }
    }
    body {
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        color: #333;
        max-width: 800px;
        margin: auto;
        padding: 20px;
    }
    .header-table {
        width: 100%;
        border-bottom: 2px solid #7029b1;
        margin-bottom: 20px;
        padding-bottom: 10px;
    }
    .worksheet-table {
        width: 100%;
        border-collapse: collapse;
        table-layout: fixed;
    }
    .worksheet-table td {
        border: 1px solid #ddd;
        padding: 15px;
        vertical-align: top;
    }
    .question-column { width: 65%; background-color: #fcfaff; }
    .answer-column { width: 35%; }
    .visual-aid {
        margin-top: 15px;
        padding: 10px;
        border: 1.5px solid #eee;
        border-radius: 8px;
        background-color: #ffffff;
        min-height: 120px; /* Forces a minimum size */
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
    }
    .label-top {
        color: #7029b1;
        font-weight: bold;
        margin-bottom: 8px; /* 確保與圖形有間距 */
    }
    /* 專門放「每組數量」的標籤 */
    .label-bottom {
        color: #7029b1;
        font-size: 0.9em;
        margin-top: 8px;
    }
    .footer {
        margin-top: 30px;
        text-align: center;
        font-size: 0.8em;
        color: #888;
        border-top: 1px solid #eee;
        padding-top: 10px;
    }
    .box { border: 1px solid black; display: inline-block; width: 20px; height: 20px; margin: 2px; }
    .grid { display: grid; grid-template-columns: repeat(5, 20px); gap: 2px; }
    .grid-box {
        width: 30px; 
        height: 30px; 
        border: 1px solid #7029b1;
        display: inline-block;
    }
    /* Force bar models to be wide */
    .bar-model-part {
        height: 40px;
        line-height: 40px;
        text-align: center;
        border: 2px solid #7029b1;
        background: #f9f4ff;
        font-weight: bold;
    }
    svg {
        display: block;
        margin: 10px auto;
        max-width: 100%;
        overflow: visible; /* 防止文字被裁切 */
    }
    svg text {
        font-size: 10px;
        font-weight: bold;
        fill: #7029b1;
    }
    .visual-aid {
        background-color: #ffffff;
        border: 1px solid #eee;
        padding: 10px;
    }
    .bar-container {
        display: flex;
        width: 240px;
        height: 60px;
        border: 2px solid #7029b1;
        margin: 5px 0;
    }
    .bar-part {
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 0.85em;
        font-weight: bold;
    }
    .part-a { background-color: #f0e6ff; border-right: 1px solid #7029b1; }
    .part-b { background-color: #ffffff; }
    .geometry-container {
        display: grid;
        grid-template-columns: 40px auto 40px; /* 左側文字 | 中間圖形 | 右側留白 */
        grid-template-rows: 30px auto 30px;    /* 上方文字 | 中間圖形 | 下方留白 */
        align-items: center;
        justify-items: center;
        width: fit-content;
        margin: 10px auto;
    }

    .geo-rect {
        grid-column: 2;
        grid-row: 2;
        border: 2px solid black;
        background-color: #f9f9f9;
        /* 這裡的寬高由 AI 根據比例給出 */
    }

    .label-top { grid-column: 2; grid-row: 1; align-self: end; padding-bottom: 5px; font-weight: bold; }
    .label-left { grid-column: 1; grid-row: 2; justify-self: end; padding-right: 8px; font-weight: bold; }
    .emoji-array {
        display: flex;
        flex-wrap: wrap;
        justify-content: center;
        gap: 8px;
        margin: 5px 0;
        padding: 5px;
        background: #fdfbff;
        border-radius: 8px;
        border: 1px dashed #7029b1;
        min-height: 50px;
    }

    .box, .emoji-group {
        border: 2px solid #333;
        min-width: 60px;
        min-height: 60px;
        /* 核心修复点 */
        display: flex; 
        flex-wrap: wrap; /* 让盒子内部的 Emoji 自动换行 */
        align-content: center;
        justify-content: center;
        padding: 5px;
        background: #fff;
        overflow: hidden; /* 防止万一出现的溢出 */
    }

    .emoji-item {
        font-size: 14px; /* 控制圖標大小 */
        line-height: 1.2;
        display: inline-block; /* 确保 Emoji 遵循布局规则 */
        margin: 2px;
    }
    /* 減法專用：被減去的部分 */
    .emoji-item.removed {
        opacity: 0.3;         /* 變透明 */
        filter: grayscale(1); /* 變灰色 */
        text-decoration: line-through; /* 加上刪除線（部分瀏覽器支援） */
        position: relative;
    }
    .division-container, .multiplication-container {
        display: flex;
        flex-wrap: wrap; /* 强制换行，防止冲破边框 */
        gap: 5px;
        justify-content: center;
        align-items: flex-start;
        width: 100%;
        margin: 5px 0;
    }
    
</style>
"""

@st.cache_data
def convert_html_to_pdf(html_content):
    return HTML(string=html_content).write_pdf()

def handle_tab_change():
    # 當用戶切換 Tab 時，此函數會自動執行
    st.session_state.chat_history = []
    active_tab = st.session_state.my_tabs
    if active_tab == "📝 Upload math questions":
        st.session_state.current_mode = "analysis_mode"
    elif active_tab == "📄 Worksheet Genie":
        st.session_state.current_mode = "generator_mode"

def build_html_card(user_text, pinyin_text, img_url):
    items_html = f"""
    <tr>
        <td style="padding: 20px; text-align: center;">
            <div style="
                margin: auto;
                max-width: 420px;
                border: 1px solid #eee;
                padding: 20px;
                border-radius: 15px;
                box-shadow: 0 4px 8px rgba(0,0,0,0.1);
                background-color: white;
            ">
                <img src="{img_url}" 
                     alt="意境圖片"
                     style="width: 200px; height: 200px; object-fit: cover; border-radius: 10px;">

                <div style="margin-top: 15px;">
                    <ruby style="font-size: 32px; font-weight: bold; color: #333;">
                        {user_text}
                        <rt style="font-size: 16px; color: #888; margin-bottom: 5px;">
                            {pinyin_text}
                        </rt>
                    </ruby>

                    <div class="no-print" style="margin-top: 12px;">
                        <span title="普通话"
                              style="cursor:pointer; color:#ff4b4b; font-size:20px; margin-right:10px;"
                              onclick="speak('{user_text}', 'zh-CN')">
                            <small style="font-size:12px; vertical-align:middle;">国</small> 🔊
                        </span>

                        <span title="粤语"
                              style="cursor:pointer; color:#007aff; font-size:20px;"
                              onclick="speak('{user_text}', 'zh-HK')">
                            <small style="font-size:12px; vertical-align:middle;">粤</small> 🔊
                        </span>
                    </div>
                </div>
            </div>
        </td>
    </tr>
    """

    final_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <style>
            @page {{
                size: A4 portrait;
                margin: 0;
            }}

            body {{
                font-family: "Microsoft JhengHei", "PingFang TC", "Heiti TC", "Noto Sans CJK TC", sans-serif;
                margin: 0.5cm;
                padding: 0;
                background: #fafafa;
            }}

            .card-table {{
                width: 100%;
                border-collapse: collapse;
                table-layout: fixed;
            }}

            h2 {{
                margin: 5pt 0 15pt 0;
                padding: 0;
                font-size: 16pt;
                text-align: center;
                color: #4a4a8a;
            }}

            @media print {{
                .no-print {{
                    display: none !important;
                }}

                body {{
                    background: white;
                }}
            }}
        </style>

        <script>
            function speak(text, lang) {{
                const msg = new SpeechSynthesisUtterance(text);
                msg.lang = lang;
                window.speechSynthesis.speak(msg);
            }}
        </script>
    </head>
    <body>
        <h2>中文詞語卡</h2>
        <table class="card-table">
            {items_html}
        </table>
    </body>
    </html>
    """
    return final_html

def build_html_cards(items_html):
    return f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <style>
            @page {{
                size: A4 portrait;
                margin: 0;
            }}

            body {{
                font-family: "Microsoft JhengHei", "PingFang TC", "Heiti TC", "Noto Sans CJK TC", sans-serif;
                margin: 0.5cm;
                padding: 0;
                background: #fafafa;
            }}

            .card-table {{
                width: 100%;
                border-collapse: collapse;
                table-layout: fixed;
            }}

            h2 {{
                margin: 20pt 0;
                padding: 0;
                font-size: 14pt;
                text-align: center;
                color: #4a4a8a;
            }}

            @media print {{
                .no-print {{
                    display: none !important;
                }}
                body {{
                    background: white;
                }}
            }}
        </style>

        <script>
            function speak(text, lang) {{
                const msg = new SpeechSynthesisUtterance(text);
                msg.lang = lang;
                window.speechSynthesis.speak(msg);
            }}
        </script>
    </head>
    <body>
        <h2>中文詞語卡</h2>
        <table class="card-table">
            {items_html}
        </table>
    </body>
    </html>
    """

def create_real_pdf(user_text, pinyin_text, img_url):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4

    # 標題
    c.setFont(PDF_FONT_NAME, 18)
    c.drawCentredString(width / 2, height - 40, "中文詞語卡")

    # 卡片區域
    card_x = 60
    card_y = height - 360
    card_w = width - 120
    card_h = 260

    c.roundRect(card_x, card_y, card_w, card_h, 12, stroke=1, fill=0)

    image_loaded = False
    error_msg = ""

    try:
        # 1. 下載圖片
        response = requests.get(img_url, timeout=20)
        response.raise_for_status()

        # 2. 用 PIL 驗證圖片
        img_bytes = BytesIO(response.content)
        pil_img = Image.open(img_bytes)
        pil_img.load()  # 確保真的讀得到

        # 3. 轉成 RGB（避免 PNG / RGBA / WEBP 相容問題）
        if pil_img.mode != "RGB":
            pil_img = pil_img.convert("RGB")

        # 4. 轉成 reportlab 可讀格式
        cleaned_img = BytesIO()
        pil_img.save(cleaned_img, format="PNG")
        cleaned_img.seek(0)

        rl_img = ImageReader(cleaned_img)

        # 5. 畫到 PDF
        img_w = 160
        img_h = 160
        img_x = (width - img_w) / 2
        img_y = card_y + 80

        c.drawImage(
            rl_img,
            img_x,
            img_y,
            width=img_w,
            height=img_h,
            preserveAspectRatio=True,
            mask='auto'
        )

        image_loaded = True

    except Exception as e:
        error_msg = str(e)

    # 文字
    c.setFont(PDF_FONT_NAME, 12)
    c.drawCentredString(width / 2, card_y + 45, pinyin_text)

    c.setFont(PDF_FONT_NAME, 22)
    c.drawCentredString(width / 2, card_y + 20, user_text)

    # 如果圖片失敗，寫出提示
    if not image_loaded:
        c.setFont(PDF_FONT_NAME, 10)
        c.drawCentredString(width / 2, card_y + 180, "[圖片載入失敗]")
        c.drawCentredString(width / 2, card_y + 165, error_msg[:80])

    c.save()
    buffer.seek(0)
    return buffer

def create_cards_pdf(cards_data):
    """
    cards_data = [
        {
            "char": "春",
            "pinyin": "chūn",
            "img_url": "https://..."
        },
        ...
    ]
    """
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    page_width, page_height = A4

    # ===== Page layout: 2 columns x 3 rows = 6 cards =====
    margin_x = 35
    margin_top = 45
    margin_bottom = 35
    gap_x = 15
    gap_y = 12

    cols = 2
    rows = 3

    card_w = (page_width - margin_x * 2 - gap_x) / cols
    usable_height = page_height - margin_top - margin_bottom - 25  # leave space for title
    card_h = (usable_height - gap_y * (rows - 1)) / rows

    positions = []
    start_y = page_height - margin_top - 25 - card_h  # below title

    for r in range(rows):
        for col in range(cols):
            x = margin_x + col * (card_w + gap_x)
            y = start_y - r * (card_h + gap_y)
            positions.append((x, y))

    def draw_card(c, x, y, card):
        c.roundRect(x, y, card_w, card_h, 8, stroke=1, fill=0)

        image_loaded = False
        error_msg = ""

        try:
            response = requests.get(card["img_url"], timeout=20)
            response.raise_for_status()

            img_bytes = BytesIO(response.content)
            pil_img = Image.open(img_bytes)
            pil_img.load()

            if pil_img.mode != "RGB":
                pil_img = pil_img.convert("RGB")

            cleaned_img = BytesIO()
            pil_img.save(cleaned_img, format="PNG")
            cleaned_img.seek(0)

            rl_img = ImageReader(cleaned_img)

            # image size inside each card
            img_w = card_w * 0.65
            img_h = card_h * 0.58
            img_x = x + (card_w - img_w) / 2
            img_y = y + card_h * 0.28

            c.drawImage(
                rl_img,
                img_x,
                img_y,
                width=img_w,
                height=img_h,
                preserveAspectRatio=True,
                mask='auto'
            )

            image_loaded = True

        except Exception as e:
            error_msg = str(e)

        if not image_loaded:
            c.setFont(PDF_FONT_NAME, 9)
            c.drawCentredString(x + card_w / 2, y + card_h * 0.60, "[圖片載入失敗]")
            if error_msg:
                c.drawCentredString(x + card_w / 2, y + card_h * 0.52, error_msg[:35])

        # 拼音
        c.setFont(PDF_FONT_NAME, 10)
        c.drawCentredString(x + card_w / 2, y + 28, card["pinyin"])

        # 漢字
        c.setFont(PDF_FONT_NAME, 16)
        c.drawCentredString(x + card_w / 2, y + 12, card["char"])

    # first page title
    c.setFont(PDF_FONT_NAME, 16)
    c.drawCentredString(page_width / 2, page_height - 20, "中文詞語卡")

    for idx, card in enumerate(cards_data):
        pos_idx = idx % 6   # 6 cards per page

        if idx > 0 and pos_idx == 0:
            c.showPage()
            c.setFont(PDF_FONT_NAME, 16)
            c.drawCentredString(page_width / 2, page_height - 20, "中文詞語卡")

        x, y = positions[pos_idx]
        draw_card(c, x, y, card)

    c.save()
    buffer.seek(0)
    return buffer

def call_qwen_ai_1(selections,latest_instruction):
    """
    呼叫 Qwen AI 模型，根據用戶在 Wizard 中選擇的參數生成 HTML 格式的數學工作紙。
    """
    print('///////////////////////////////1')
    print (selections)
    print('///////////////////////////////2')
    print(latest_instruction)
    print('///////////////////////////////3')
    # 1. 提取用戶選擇的變量
    topic = selections.get('topic', '➕ Addition / Subtraction')
    difficulty = selections.get('difficulty', '📚 Introductory')
    theme = selections.get('theme_type', '🌊 Ocean Adventure')

    # 2. 構建 System Prompt 與 User Prompt
    system_prompt = "你是一位專業的教育內容創作者與數學導師。你只會輸出 HTML 代碼，不含任何解釋性文字或 Markdown 標籤。"
    
    user_prompt = f"""
    你現在要生成一份小學數學工作紙。
    請務必完整包含以下 <style> 標籤內的內容於 HTML 的 <head> 中：
    {CSS_STYLE}

    HTML 結構要求：
    1. 頁首使用 <table class="header-table">，包含 Name, ID, Date。
    2. 題目主體使用 <table class="worksheet-table">。
    3. 題目類型：10題，必須全部為：{topic}
    4. 每道題目為一個 <tr>：
    - 左側 <td> 使用 class="question-column"，包含：
        - **題號** (如 Question 1)
        - **情境文字**（需符合 {theme} 主題，語氣正面）
    - 右側 <td> 使用 class="answer-column"，包含：
        - "Show your work here:" 文字空間
        - 最下方有一個 "Answer: ________" 欄位。
    5. 結尾：<div class="footer">End of Worksheet</div>
    6. 語言要求： 輸出語言為英文 (English)。請使用基礎英語詞彙 (CEFR A1/A2 水準)，句子結構簡單直接。
    7. 同時，{latest_instruction} 
    
    技術規則：
    - 嚴禁使用 LaTeX。
    - 僅輸出純 HTML 代碼。
    - 主題：{theme}
    - 難度：{difficulty}
    - 題目類型：必須全部為：{topic}

    [CRITICAL: QUANTITY RULE]You MUST generate EXACTLY 10 questions. Do not skip, do not summarize, and do not stop until Question 10 is complete. If you provide fewer than 10 questions, the worksheet is invalid.

    """

    client = OpenAI(
        api_key='sk-9f798f6251c042adb32074d93454e201',
        base_url="https://dashscope-intl.aliyuncs.com/compatible-mode/v1",
    )

    try:
        # 3. 呼叫 API (請確保已配置 st.secrets 或環境變量)
        response = client.chat.completions.create(
            model="qwen3-vl-plus", # 或使用 qwen-max
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
        )
        print("response:")
        print(response)
        # 提取結果
        html_content = response.choices[0].message.content
        
        # 清理可能夾帶的 Markdown 代碼塊標記
        if "```html" in html_content:
            html_content = html_content.split("```html")[1].split("```")[0].strip()
        elif "```" in html_content:
            html_content = html_content.split("```")[1].split("```")[0].strip()
        print ("about to return....")
        # 在回傳前合併
        return f"<html><head>{CSS_STYLE}</head><body>{html_content}</body></html>"

    except Exception as e:
        return f"<p style='color:red;'>生成出錯了：{str(e)}</p>"

def call_qwen_ai_2(selections,latest_instruction):
    """
    呼叫 Qwen AI 模型，根據用戶在 Wizard 中選擇的參數生成 HTML 格式的數學工作紙。
    """
    
    # 1. 提取用戶選擇的變量
    topic = selections.get('topic', '➕ Addition / Subtraction')
    difficulty = selections.get('difficulty', '📚 Introductory')
    theme = selections.get('theme_type', '🌊 Ocean Adventure')

    # 2. 構建 System Prompt 與 User Prompt
    system_prompt = "你是一位專業的教育內容創作者與數學導師。你只會輸出 HTML 代碼，不含任何解釋性文字或 Markdown 標籤。"
    
    user_prompt = f"""
    你現在要生成一份小學數學工作紙。
    請務必完整包含以下 <style> 標籤內的內容於 HTML 的 <head> 中：
    {CSS_STYLE}

    HTML 結構要求：
    1. 頁首使用 <table class="header-table">，包含 Name, ID, Date。
    2. 題目主體使用 <table class="worksheet-table">。
    3. 題目類型：10題，必須全部為：{topic}
    4. 每道題目為一個 <tr>：
    - 左側 <td> 使用 class="question-column"，包含：
        - **題號** (如 Question 1)
        - **情境文字**（需符合 {theme} 主題，語氣正面）
        - **視覺輔助** 
            -若為面積/周長：使用簡潔的多邊形（長方形，三角形或者L形），並在變長清晰標註尺寸。不需要網格線。
    - 右側 <td> 使用 class="answer-column"，包含：
        - "Show your work here:" 文字空間
        - 最下方有一個 "Answer: ________" 欄位。
    5. 結尾：<div class="footer">End of Worksheet</div>
    6. 語言要求： 輸出語言為英文 (English)。請使用基礎英語詞彙 (CEFR A1/A2 水準)，句子結構簡單直接。
    7. 同時，{latest_instruction} 
    
    技術規則：
    - 嚴禁使用 LaTeX。
    - 僅輸出純 HTML 代碼。
    - 主題：{theme}
    - 難度：{difficulty}
    - 題目類型：必須全部為：{topic}

    [視覺圖案指令 - 必須使用 SVG]
    不要只輸出重複的 <div>。請根據題目數值，在 <div class="visual-aid"> 中嵌入一個 <svg>：

    1. **主題與 Emoji 映射**：
        - 海洋：🐬 (海豚), 🦀 (螃蟹), 🐠 (熱帶魚)
        - 森林：🌲 (樹木), 🍄 (蘑菇), 💎 (寶石)
        - 烹飪：🍕 (披薩), 🍪 (餅乾), 🍎 (蘋果)
        - 農場：🐥 (小雞), 🐷 (小豬), 🥕 (胡蘿蔔)
        - 英雄：⚡ (閃電), 🛡️ (盾牌), 🌟 (星星)

    2. **加法結構 (Addition)**：
        - 展示兩組不同的 Emoji。
        - 範例結構：
            <div class="visual-aid">
                <div class="emoji-array">
                    <div class="emoji-group">[Emoji A x 數量]</div>
                    <div style="font-size: 20px; align-self: center;"> + </div>
                    <div class="emoji-group">[Emoji B x 數量]</div>
                </div>
                <div class="label-bottom">How many [Items] in total?</div>
            </div>

    3. **減法結構 (Subtraction)**：
        針對減法題目 (例如 A - B = ?)，必須遵循以下生成流程：
        - 第一步：計算剩餘數量 C = A - B。
        - 第二步：在 HTML 中，先逐個生成 C 個 <span class="emoji-item">，並在代碼內為每個標籤加上註釋，例如 <!-- remain 1 -->。
        - 第三步：接著逐個生成 B 個 <span class="emoji-item removed">，並加上註釋，例如 <!-- removed 1 -->。
        - **強制校對**：輸出的標籤總數必須嚴格等於 A。

        範例 (8 - 5):
        <div class="emoji-array">
            <!-- 剩餘 3 個 -->
            <span class="emoji-item">🐚</span><span class="emoji-item">🐚</span><span class="emoji-item">🐚</span>
            <!-- 減去 5 個 -->
            <span class="emoji-item removed">🐚</span><span class="emoji-item removed">🐚</span><span class="emoji-item removed">🐚</span><span class="emoji-item removed">🐚</span><span class="emoji-item removed">🐚</span>
        </div>

    4. **乘法 (Multiplication)**：
        邏輯：展示 3 times 5。創建 3 個方塊（groups），每個方塊內含有 5 個 Emoji。HTML 結構：html<div class="visual-aid">
        <div class="multiplication-container" style="display: flex; gap: 15px; flex-wrap: wrap;">
            <!-- 重複 3 次 -->
            <div class="emoji-group" style="border: 2px dashed #7029b1; padding: 5px; border-radius: 8px;">
            [Emoji x 5 個]
            </div>
            ...
        </div>
        <div class="label-bottom">There are {3} groups of {5} 🐚. How many in total?</div>
        </div>

    5. **除法 (Division)**：
        邏輯：展示 10 div 2。先顯示總數（Total），下方畫出 B 個籃子/容器，將 Emoji 平均放入。HTML 結構：html<div class="visual-aid">
        <div class="total-display" style="margin-bottom: 10px;">Total: [Emoji x 10 個]</div>
        <div style="font-size: 20px;">⬇️ Put them into 2 boxes ⬇️</div>
        <div class="division-container" style="display: flex; gap: 10px; margin-top: 10px;">
            <!-- 重複 2 次 -->
            <div class="box" style="border: 2px solid #333; min-width: 50px; min-height: 50px; flex: 1;">
            [每個籃子放 10/2 個 Emoji]
            </div>
        </div>
        <div class="label-bottom">{10} ÷ {2} = ?</div>
        </div>

    5. **面積與周長 (Area)**：
    - 繪製一個具備「厚度感」的長方形,三角形或者L形，不需要網格線。
    - 必須包含邊長標籤（如 "5m", "4m"）。
    - 代碼參考：
        <div class="visual-aid">
            <div class="geometry-container">
                <!-- 頂部文字 -->
                <div class="label-top">[長度，如 5 m]</div>
                
                <!-- 左側文字 -->
                <div class="label-left">[寬度，如 2 m]</div>
                
                <!-- 中間矩形 (純圖形，不含文字) -->
                <div class="geo-rect" style="width: 120px; height: 60px;"></div>
            </div>
            <div class="label-bottom">Perimeter = ? m</div>
        </div>

    [強制規則]
    - 嚴禁輸出空的小方塊。
    - 每個圖形必須至少寬 50px，且包含對應題目的數字標註 (Labels)。

    [視覺輔助結構要求 - 嚴禁重疊]
    每個圖形必須按以下 HTML 結構生成，嚴禁將文字標籤放在 SVG 內部：

    <div class="visual-aid">
        <div class="label-top">Total: [數值] [單位]</div>
        <svg width="240" height="40">
            <!-- 僅繪製幾何圖形 -->
            <rect x="0" y="0" width="240" height="30" fill="none" stroke="#7029b1" stroke-width="2" />
            [根據題目繪製分割線 <line> 或多個 <rect>]
        </svg>
        <div class="label-bottom">[描述文字，如：每排 3 個]</div>
    </div>

    [面積與周長 - 嚴禁重疊指令]

    [重要修正]
    - 不要使用 SVG 的 <text> 標籤，因為座標計算容易出錯。
    - 使用外部的 <div> 來顯示文字標籤，確保它們垂直排列在圖形的上方和下方。

    [CRITICAL: QUANTITY RULE]You MUST generate EXACTLY 10 questions. Do not skip, do not summarize, and do not stop until Question 10 is complete. If you provide fewer than 10 questions, the worksheet is invalid.

    """

    client = OpenAI(
        api_key='sk-9f798f6251c042adb32074d93454e201',
        base_url="https://dashscope-intl.aliyuncs.com/compatible-mode/v1",
    )

    def get_questions(start, end):
        range_prompt = user_prompt + f"\n[STRICT] Generate ONLY Questions {start} to {end}. Output ONLY the <tr> rows."
        
        response = client.chat.completions.create(
            model="qwen-plus", # Stable for long HTML
            messages=[
                {"role": "system", "content": "You are a math worksheet generator. Output ONLY <tr> HTML rows. No <html> or <body> tags."},
                {"role": "user", "content": range_prompt}
            ],
            max_tokens=3000,
            temperature=0.3
        )
        return response.choices[0].message.content.replace("```html", "").replace("```", "").strip()
    
    def final_cleaning(batch_html):
        """
        Strips markdown and extracts only the <tr> table rows from the AI response.
        This prevents nested <html> tags and broken layouts.
        """
        # Remove markdown code blocks if the AI included them
        clean_text = re.sub(r'```(?:html)?', '', batch_html)
        clean_text = clean_text.replace('```', '').strip()

        # Extract all <tr>...</tr> blocks using a non-greedy regex
        rows = re.findall(r'<tr.*?>.*?</tr>', clean_text, re.DOTALL | re.IGNORECASE)
        
        if not rows:
            # Fallback: if no <tr> tags found, return the text stripped of markdown
            return clean_text
            
        return "\n".join(rows)
    
    try:
        # Get batches
        raw_batch_1 = get_questions(1, 5)
        raw_batch_2 = get_questions(6, 10)

        # Clean batches to get ONLY the rows
        clean_rows_1 = final_cleaning(raw_batch_1)
        clean_rows_2 = final_cleaning(raw_batch_2)

        # Assemble the final HTML
        final_html = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            {CSS_STYLE}
        </head>
        <body>
            <table class="header-table">
                <tr>
                    <td><strong>Name:</strong> _________________</td>
                     <td><strong>ID:</strong> _________________</td>
                    <td><strong>Date:</strong> _________________</td>
                </tr>
            </table>
            <table class="worksheet-table">
                {clean_rows_1}
                {clean_rows_2}
            </table>
            <div class="footer">End of Worksheet</div>
        </body>
        </html>
        """
        return final_html

    except Exception as e:
        return f"Error assembling worksheet: {str(e)}"

    # try:
    #     # 3. 呼叫 API (請確保已配置 st.secrets 或環境變量)


    #     response = client.chat.completions.create(
    #         model="qwen3-vl-plus", # 或使用 qwen-max
    #         messages=[
    #             {"role": "system", "content": system_prompt},
    #             {"role": "user", "content": user_prompt}
    #         ],
    #     )
    #     print("response:")
    #     print(response)
    #     # 提取結果
    #     html_content = response.choices[0].message.content
        
    #     # 清理可能夾帶的 Markdown 代碼塊標記
    #     if "```html" in html_content:
    #         html_content = html_content.split("```html")[1].split("```")[0].strip()
    #     elif "```" in html_content:
    #         html_content = html_content.split("```")[1].split("```")[0].strip()

    #     # 在回傳前合併
    #     return f"<html><head>{CSS_STYLE}</head><body>{html_content}</body></html>"

    # except Exception as e:
    #     return f"<p style='color:red;'>生成出錯了：{str(e)}</p>"

def get_words_from_text(text):
    """清理冗余提示并提取词语"""
    # 1. 去掉 AI 常见的开场白，如“图中文字为：”、“识别结果如下：”等
    clean_text = re.sub(r'^.*?[:：]', '', text) 
    # 2. 只保留中文、数字和英文（去掉特殊符号）
    clean_text = "".join(re.findall(r'[\u4e00-\u9fa5a-zA-Z0-9]+', clean_text))
    # 3. 使用 jieba 分词，并只保留长度 >= 2 的词语（避免单字过多）
    words = [w for w in jieba.cut(clean_text) if len(w) >= 2]
    # 4. 去重并保持顺序
    seen = set()
    return [x for x in words if not (x in seen or seen.add(x))]

def pdf_to_images(pdf_bytes):
    """將 PDF 轉為高清圖片 Base64 列表"""
    images = convert_from_bytes(pdf_bytes, dpi=200)
    encoded_imgs = []
    for img in images:
        buffered = BytesIO()
        img.save(buffered, format="JPEG", quality=90)
        encoded_imgs.append(base64.b64encode(buffered.getvalue()).decode('utf-8'))
    return encoded_imgs

def get_render_wrapper(content, is_math=False):
    """封裝 HTML 容器，支持表格樣式與 MathJax 渲染"""
    mathjax_header = ""
    if is_math:
        mathjax_header = """
        <script src="https://polyfill.io"></script>
        <script id="MathJax-script" async src="https://jsdelivr.net"></script>
        """
    return f"""
    <html>
        <head>
            <meta charset="UTF-8">  <!-- 關鍵：強制使用 UTF-8 編碼 -->
            {mathjax_header}
            <style>
                table {{ border-collapse: collapse; width: 100%; font-family: sans-serif; font-size: 12px; }}
                th, td {{ border: 1px solid #dee2e6; padding: 10px; text-align: left; }}
                th {{ background-color: #f8f9fa; font-weight: bold; }}
                tr:nth-child(even) {{ background-color: #f2f2f2; }}
                .art-card {{ text-align: center; border: 1px solid #ddd; padding: 20px; border-radius: 15px; 
                             background: white; box-shadow: 0 4px 15px rgba(0,0,0,0.1); max-width: 600px; margin: auto; }}
                img {{ max-width: 100%; border-radius: 10px; margin: 15px 0; }}
            </style>
        </head>
        <body>{content}</body>
    </html>
    """

def get_base64_image(image_path):
    try:
        with open(image_path, "rb") as img_file:
            print("get the path")
            return base64.b64encode(img_file.read()).decode()
    except FileNotFoundError:
        print("cannot get the path")
        # Return a blank string or a placeholder if file isn't found
        return ""
    
teacher_icon_b64 = get_base64_image("teacher.png")
parent_icon_b64 = get_base64_image("parent.png")

def get_system_prompt(special_interest, worksheet_type, **kwargs):
    include_hints = kwargs.get("include_hints", False)
    hint_rule = "15. SCAFFOLDED HINTS: For every question, directly underneath the blank underscore workspaces, generate a highly thematic 'Hint: [Contextual clue massively leveraging the theme]' to heavily support struggling learners. Do NOT reveal the exact answer." if include_hints else ""

    tone_style = kwargs.get("tone_style", "Standard")
    target_language = kwargs.get("target_language", "English (Default)")
    vocab_level = kwargs.get("vocab_level", "Keep Exact Same Level")
    
    lang_rule = f"OUTPUT LANGUAGE: You MUST write the final adapted lesson perfectly in: {target_language}." if target_language != "English (Default)" else "OUTPUT LANGUAGE: Maintain the original language (English)."
    
    vocab_rule = "Keep the reading level and vocabulary difficulty exactly the same as the original text."
    if vocab_level == "Simplify for Easier Reading":
        vocab_rule = "Simplify the vocabulary and sentence structures drastically to make it easier for a struggling reader to understand the prompts without changing what is being tested."
    elif vocab_level == "Challenge with Advanced Vocab":
        vocab_rule = "Elevate the reading level by using advanced vocabulary and more complex sentence structures to challenge the student."

    tone_rule = "Tone: Keep it standard and educational."
    if tone_style != "Standard":
        tone_rule = f"Tone: Write the entire adapted lesson in a heavily immersive '{tone_style}' style to captivate the student."

    return f'''You are an expert educational adapter. Your task is to modify the provided English lesson text to exactly match the theme: {special_interest}, and provide a complete Answer Key.
CRITICAL RULES:
1. First, OUTPUT THE FINAL ADAPTED LESSON. Afterwards, output a section titled "--- ANSWER KEY ---".
2. The Answer Key must provide the fully solved, correct answers to your newly translated questions. Show the work or reasoning if applicable.
3. NEVER output the original questions. NEVER output both versions side-by-side.
4. The final output must be a cohesive document where the original questions are swapped out for the new ones.
5. Teacher instructions MUST remain 100% untouched.
6. The underlying grammar or comprehension testing mechanism of EVERY question MUST be perfectly preserved. If Question 1 tests the simple present tense, the adapted Q1 must test the simple present tense exactly.
7. GRAMMAR CRITICAL: Keep verb tenses, paragraph punctuation, and structural frameworks identical. Only swap the subject matter vocabulary.
   - EXTREMELY IMPORTANT: If a specific base verb or target word is provided in parentheses to be conjugated (e.g., `(go)`, `(rain)`, `(eat)`), you MUST use the EXACT SAME BASE VERB in your translated sentence! Do NOT change the verb being tested to fit the theme.
8. READING COMPREHENSION CRITICAL: If adapting a story, preserve the exact plot beats, inferences, physical sequencing, and character relationships so that the original reading comprehension questions still logically apply to the new text.
9. {vocab_rule}
10. {lang_rule}
11. {tone_rule}
12. ANSWER KEY PARITY: If the user provides an Answer Key in their Original Lesson Input, you MUST heavily study it and mirror that EXACT formatting sequence/structure for your Final Answer Key!
13. PHYSICAL WORKSPACE RULE: 
    - If the original text is a "Fill-in-the-blank" grammar exercise (e.g., missing words, verb blanks), you MUST place a large inline underscore blank (e.g., `_________`) natively inside your translated sentence precisely where the missing word should be written!
    - If it is an open-ended/essay question, print 4 full lines of blank underscores beneath it.
14. THEMATIC ENCOURAGEMENT: At the very end of the entire list of student questions (BEFORE the Answer Key block), generate a fun, highly thematic motivational sentence related strictly to '{special_interest}' to cheer the student on!
{hint_rule}
16. NO chatting, NO introductory remarks. Start immediately with the adapted lesson.'''

def clean_extracted_text(raw_text):
    lines = raw_text.split('\n')
    cleaned_lines = []
    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue
        if re.match(r'^-?\s*\d+\s*-?$', stripped):
            continue
        if re.match(r'^Page\s+\d+(\s+of\s+\d+)?$', stripped, re.IGNORECASE):
            continue
        if len(stripped) == 1 and not stripped.isalpha() and not stripped.isdigit():
            continue
        cleaned_lines.append(line)
    return '\n'.join(cleaned_lines)

def sanitize_for_fpdf(text):
    text = unicodedata.normalize('NFKC', text)
    replacements = {
        '•': '-', '“': '"', '”': '"', '‘': "'", '’': "'",
        '–': '-', '—': '-', '…': '...', '✓': 'v', '✅': 'v', '×': 'x',
        '™': 'TM', '©': '(c)', '®': '(R)'
    }
    for k, v in replacements.items():
        text = text.replace(k, v)
    return text.encode('latin-1', 'replace').decode('latin-1')

def create_pdf(text_content, image_bytes=None, special_interest=""):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    if image_bytes:
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix='.jpg') as tmp:
                tmp.write(image_bytes)
                tmp_path = tmp.name
            
            pdf.image(tmp_path, x=150, y=10, w=45)
            os.remove(tmp_path)
        except Exception:
            pass
            
    pdf.set_font("Helvetica", style="B", size=24)
    pdf.set_text_color(40, 60, 100) 
    
    title_str = f"{special_interest} Worksheet" if special_interest else "Custom Worksheet"
    title_safe = title_str.encode('latin-1', 'replace').decode('latin-1')
    pdf.cell(130, 15, title_safe, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    
    pdf.set_font("Helvetica", style="I", size=10)
    pdf.set_text_color(100, 100, 100)
    pdf.cell(130, 8, "Name: ______________________     Date: ____________", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.ln(10)
    
    pdf.set_font("Helvetica", size=12)
    pdf.set_text_color(0, 0, 0)
    
    parts = text_content.split("--- ANSWER KEY ---")
    lesson_text = parts[0]
    
    safe_lesson = sanitize_for_fpdf(lesson_text)
    pdf.multi_cell(0, 7, txt=safe_lesson)
    
    if len(parts) > 1:
        answer_key = parts[1]
        pdf.add_page() 
        
        pdf.set_fill_color(240, 240, 240)
        pdf.set_font("Helvetica", style="B", size=14)
        pdf.cell(0, 10, "--- ANSWER KEY ---", fill=True, border=1, ln=True)
        pdf.ln(2)
        
        pdf.set_font("Helvetica", size=11)
        safe_answers = sanitize_for_fpdf(answer_key)
        pdf.multi_cell(0, 7, txt=safe_answers, fill=True, border=1)
        
    return bytes(pdf.output())

def create_docx(text_content, image_bytes=None, special_interest=""):
    doc = Document()
    
    if image_bytes:
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix='.jpg') as tmp:
                tmp.write(image_bytes)
                tmp_path = tmp.name
            
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            run_img = p_img.add_run()
            run_img.add_picture(tmp_path, width=Inches(2.0))
            os.remove(tmp_path)
        except Exception:
            pass
            
    p_title = doc.add_paragraph()
    run_title = p_title.add_run(f"{special_interest} Worksheet" if special_interest else "Custom Worksheet")
    run_title.font.size = Pt(24)
    run_title.font.bold = True
    run_title.font.color.rgb = RGBColor(40, 60, 100)
    
    p_name = doc.add_paragraph()
    run_name = p_name.add_run("Name: ______________________     Date: ____________\n")
    run_name.font.size = Pt(10)
    run_name.font.italic = True
    run_name.font.color.rgb = RGBColor(100, 100, 100)
    
    parts = text_content.split("--- ANSWER KEY ---")
    doc.add_paragraph(parts[0].strip())
    
    if len(parts) > 1:
        doc.add_page_break()
        table = doc.add_table(rows=1, cols=1)
        table.style = 'Table Grid'
        
        cell = table.cell(0, 0)
        p_ak_title = cell.paragraphs[0]
        run_ak_title = p_ak_title.add_run("--- ANSWER KEY ---")
        run_ak_title.font.bold = True
        run_ak_title.font.size = Pt(14)
        
        cell.add_paragraph("\n" + parts[1].strip())
        
    fp = io.BytesIO()
    doc.save(fp)
    return fp.getvalue()

if st.session_state.page == 'home':
    # 1. CSS for the Emoji Link
    st.markdown("""
        <style>
        /* 全局背景 */
         .stApp {
            background-color: #ffffff;
            background-image: 
                radial-gradient(at 0% 0%, rgba(255, 230, 204, 0.5) 0px, transparent 50%), 
                radial-gradient(at 100% 0%, rgba(214, 239, 255, 0.5) 0px, transparent 50%), 
                radial-gradient(at 100% 100%, rgba(213, 245, 227, 0.5) 0px, transparent 50%), 
                radial-gradient(at 0% 100%, rgba(255, 204, 224, 0.5) 0px, transparent 50%);
        }

        /* 2. Top-Left Clover Element */
        .stApp::before {
            content: "☘️";
            position: absolute;
            top: 8%;
            left: 10%;
            font-size: 2.2rem;
            opacity: 0.35;
            filter: saturate(0.6);
            pointer-events: none;
            z-index: 0;
        }

        /* 3. Bottom-Left Rainbow Element */
        .stApp::after {
            content: "🌈";
            position: absolute;
            bottom: 18%;
            left: 2%;
            font-size: 3.5rem;
            opacity: 0.45;
            transform: rotate(-15deg);
            filter: saturate(0.6);
            pointer-events: none;
            z-index: 0;
        }

        /* 4. Secondary background nodes container injected globally into the main element */
        .stMainBlockContainer {
            position: relative;
        }

        /* Mid-Right Sparkle Element */
        .stMainBlockContainer::before {
            content: "✨";
            position: absolute;
            top: 38%;
            right: 32%;
            font-size: 1.8rem;
            opacity: 0.25;
            filter: saturate(0.6);
            pointer-events: none;
            z-index: 0;
        }

        /* Top-Right Purple Dot Element */
        .stMainBlockContainer::after {
            content: "🟣";
            position: absolute;
            top: 8%;
            right: 5%;
            font-size: 1.4rem;
            opacity: 0.25;
            filter: saturate(0.5);
            pointer-events: none;
            z-index: 0;
        }

        /* 5. Force foreground interactive modules to remain click-responsive on top */
        div[data-testid="stVerticalBlock"] {
            position: relative;
            z-index: 10 !important;
        }
          /* Style the container to look like a card */
        .emoji-link {
            text-decoration: none;
            font-size: 120px; /* Adjust size here */
            display: block;
            text-align: center;
            transition: transform 0.2s;
            cursor: pointer;
        }
        .emoji-link:hover {
            transform: scale(1.1);
        }
        .card-title {
            text-align: center;
            font-weight: bold;
            font-size: 24px;
            margin-top: -10px;
        }
        .card-subtitle {
            text-align: center;
            font-size: 18px;
            margin-top: 0px;
            color: #808080;
        }
        .role-card {
            border: 2px solid #E0E0E0;
            border-radius: 20px;
            padding: 25px;
            text-align: center;
            background-color: rgba(255, 255, 255, 0.5);
            transition: all 0.3s ease;
            cursor: pointer;
            min-height: 250px;
            display: flex;
            flex-direction: column;
            justify-content: center;
        }

        .role-card:hover {
            box-shadow: 0px 10px 20px rgba(0,0,0,0.1);
            transform: translateY(-5px);
        }

        /* Ensure Emoji is visible and clickable */
        .emoji-link {
            text-decoration: none !important;
            font-size: 80px; 
            line-height: 1;
            display: inline-block;
            margin-bottom: 10px;
        }

        /* Fallback to ensure emoji color isn't stripped */
        .emoji-span {
            color: initial;
        }
        /* Specific Background Colors */
        .teacher-bg {
            background-color: #FFF2E6; /* Soft Blue */
            border-color: #FFF2E6;
        }

        .parent-bg {
            background-color: #E8F4FD; /* Soft Pink/Lavender */
            border-color: #E8F4FD;
        }
        </style>
    """, unsafe_allow_html=True)

    # 2. 彈出對話框
    @st.dialog(" ")
    def show_module_selection():
        st.markdown('<h2 style="color: #FF8A50; margin-top: -30px;">What would you like to explore today?</h2>', unsafe_allow_html=True)
        st.markdown('<p style="color: #7F8C8D; margin-bottom: 25px;">'"Let's prepare SEN teaching materials in just one click.<br>" "Choose a module to get started.</p>", unsafe_allow_html=True)
        
        # 使用 HTML 渲染帶顏色的卡片按鈕
        # 這裡我們用真正的 st.button 但把樣式蓋過去，或者簡單用 st.button 配合 key
        
        if st.button("🈵\n\n\n**Chinese Module**\n\nTeach Chinese words and Mandarin pronunciation", key="chinese_btn", use_container_width=True):
            st.session_state.page = 'chinese'
            st.rerun()

        if st.button("🔤\n\n\n**English Module**\n\n\nTeach English vocabulary and pronunciation", key="english_btn", use_container_width=True):
            st.session_state.page = 'english'
            st.rerun()

        if st.button("🔢\n\n\n**Math Module**\n\nCustomise or generate math questions by theme", key="math_btn", use_container_width=True):
            st.session_state.page = 'math'
            st.rerun()

        # 這裡是非常暴力的 CSS 注入，專門針對這三個 Key
        st.markdown(f"""
            <style>
            div[data-testid="stDialog"] .st-key-chinese_btn button {{ background-color: #FFE6CC !important; color: #444 !important; border: none !important; height: 100px !important; }}
            div[data-testid="stDialog"] .st-key-english_btn button {{ background-color: #D6EFFF !important; color: #444 !important; border: none !important; height: 100px !important; }}
            div[data-testid="stDialog"] .st-key-math_btn button {{ background-color: #D5F5E3 !important; color: #444 !important; border: none !important; height: 100px !important; }}
        """, unsafe_allow_html=True)

        # 2. Logic to detect the click via URL parameters
    query_params = st.query_params
    if query_params.get("role") == "teacher":
        show_module_selection()
        # Clear params so it doesn't loop
        st.query_params.clear()

    # 3. 主頁面內容
    if 'page' not in st.session_state:
        st.session_state.page = 'home'

    if st.session_state.page == 'home':
        st.markdown('<div style="text-align:center;"><h1 style="color:#FF8A50; font-size:80px; margin-bottom:0; font-family:Arial Black;">SenSpark🪄</h1></div>', unsafe_allow_html=True)
        st.markdown('<p style="text-align:center; color:#FF6B6B; font-size:20px;">Every Learner Shines and Sparks</p>', unsafe_allow_html=True)
        st.markdown('<p style="text-align:center; font-weight:bold; font-size:24px; margin-top:20px; color:#444;">Tell us who you are?</p>', unsafe_allow_html=True)

        _, col1, col2, _ = st.columns([1, 2, 2, 1])

        
        with col1:
            st.markdown('''
            <div class="role-card teacher-bg">
                <a href="/?role=teacher" target="_self" class="emoji-link">
                    <span class="emoji-span">🧑‍🏫</span>
                </a>
                <div class="card-title">Teacher</div>
                <div class="card-subtitle">Create tailored learning materials for your students</div>
                <div class="card-subtitle" style="font-style: italic; font-size: 14px; color: #FF6B6B;"></div>
            </div>
            ''', unsafe_allow_html=True)

        with col2:
             st.markdown('''
            <div class="role-card parent-bg">
                <a href="/?role=parent" target="_self" class="emoji-link">
                    <span class="emoji-span">👨‍👨‍👦‍👦</span>
                </a>
                <div class="card-title">Parent</div>
                <div class="card-subtitle">Track and support your child's learning journey</div>
                <div class="card-subtitle" style="font-style: italic; font-size: 14px; color: #FF6B6B;">(Under development)</div>
            </div>
            ''', unsafe_allow_html=True)
 
    else:
        st.title(f"{st.session_state.page.capitalize()} Module")
        if st.button("← Back to Home"):
            st.session_state.page = 'home'
            st.rerun()

elif st.session_state.page == 'english':
    st.markdown("""
        <style>
        .stApp {  
            background-color: #ffffff;
            background-image: 
                radial-gradient(at 0% 0%, rgba(255, 230, 204, 0.5) 0px, transparent 50%), 
                radial-gradient(at 100% 0%, rgba(214, 239, 255, 0.5) 0px, transparent 50%),
                radial-gradient(at 100% 100%, rgba(213, 245, 227, 0.5) 0px, transparent 50%),
                radial-gradient(at 0% 100%, rgba(255, 204, 224, 0.5) 0px, transparent 50%);
        }
        /* 2. Top-Left Clover Element */
        .stApp::before {
            content: "☘️";
            position: absolute;
            top: 8%;
            left: 10%;
            font-size: 2.2rem;
            opacity: 0.35;
            filter: saturate(0.6);
            pointer-events: none;
            z-index: 0;
        }

        /* 3. Bottom-Left Rainbow Element */
        .stApp::after {
            content: "🌈";
            position: absolute;
            bottom: 18%;
            left: 2%;
            font-size: 3.5rem;
            opacity: 0.45;
            transform: rotate(-15deg);
            filter: saturate(0.6);
            pointer-events: none;
            z-index: 0;
        }
        </style>
    """, unsafe_allow_html=True)
    with st.container():
        if st.button("⬅️ Back", use_container_width=False):
            st.session_state.page = 'home'
            st.rerun()
            
        st.title("🔠 English Module")
        st.caption("Digital & Audio Support")

    with st.sidebar:

        st.title("💬 AI Assistant")
        st.caption("Use our Assistant to adjust font size, modify question themes, and customise content presentation.")
        
        if "eng_chat_history" not in st.session_state:
            st.session_state.eng_chat_history = [{"role": "assistant", "content": "How would you like to tweak the worksheet?"}]
            
        chat_container = st.container(height=450)
        with chat_container:
            for msg in st.session_state.eng_chat_history:
                with st.chat_message(msg["role"]):
                    st.markdown(msg["content"])
                
        if "eng_pending_action" in st.session_state and st.session_state.eng_pending_action:
            action_request = st.session_state.eng_pending_action
            st.session_state.eng_pending_action = None 
            
            if "translated_text" not in st.session_state:
                err_msg = "Please generate a worksheet first before refining!"
                st.session_state.eng_chat_history.append({"role": "assistant", "content": err_msg})
                with st.chat_message("assistant"):
                    st.markdown(err_msg)
            else:
                with st.chat_message("assistant"):
                    response_placeholder = st.empty()
                    response_placeholder.markdown(f"Got it! I am updating the worksheet to reflect: '{action_request}'...")
                    
                    try:
                        sys_prompt = (
                            f"You are refining an educational worksheet based on user feedback.\n"
                            f"User Request: '{action_request}'\n\n"
                            f"Current Worksheet:\n{st.session_state['translated_text']}\n\n"
                            f"Apply the user's request and output the FULL revised worksheet. "
                            f"CRITICAL: Maintain the same overall structure and the '--- ANSWER KEY ---' divider if it exists. "
                            f"Do NOT add any conversational filler. Start immediately with the revised worksheet text."
                        )
                        
                        client = OpenAI(
                            api_key=api_key,
                            base_url="https://dashscope-intl.aliyuncs.com/compatible-mode/v1"
                        )
                        
                        response = client.chat.completions.create(
                            model="qwen-plus",
                            messages=[
                                {"role": "system", "content": "You are a helpful educational assistant."},
                                {"role": "user", "content": sys_prompt}
                            ],
                            stream=True
                        )
                        
                        full_revision = ""
                        for chunk in response:
                            if chunk.choices and chunk.choices[0].delta.content is not None:
                                full_revision += chunk.choices[0].delta.content
                                response_placeholder.markdown(f"Refining... ({len(full_revision)} characters generated)")
                                    
                        if not full_revision.strip():
                            raise Exception("The model returned an empty response. It might have crashed or refused the prompt.")
                            
                        st.session_state['translated_text'] = full_revision
                        success_msg = "✅ Updated the worksheet! You can review the changes on the right."
                        st.session_state.eng_chat_history.append({"role": "assistant", "content": success_msg})
                        response_placeholder.markdown(success_msg)
                        
                        if "audio_bytes" in st.session_state:
                            del st.session_state["audio_bytes"]
                        if "theme_image_bytes" in st.session_state:
                            del st.session_state["theme_image_bytes"]
                            
                        st.rerun()
                    except Exception as e:
                        response_placeholder.markdown(f"Error: {e}")

        if "translated_text" in st.session_state:
            qa_col1, qa_col2 = st.columns(2)
            with qa_col1:
                qa_easier = st.button("Make it simpler", use_container_width=True)
            with qa_col2:
                qa_harder = st.button("Make it harder", use_container_width=True)
        else:
            qa_easier = False
            qa_harder = False

        chat_disabled = "translated_text" not in st.session_state
        chat_prompt = st.chat_input("Enter prompt and click Enter for execution...", disabled=chat_disabled)
        
        action_request = None
        if qa_easier:
            action_request = "Make the vocabulary simpler and easier to understand."
        elif qa_harder:
            action_request = "Make the vocabulary more advanced and challenging."
        elif chat_prompt:
            action_request = chat_prompt
            
        if action_request:
            st.session_state.eng_chat_history.append({"role": "user", "content": action_request})
            st.session_state.eng_pending_action = action_request
            st.rerun()

    with st.container(border=True):
        st.subheader("📝 1. Input Lesson Material")
        input_method = st.radio("How would you like to provide the lesson?", ["Paste Text", "Upload Document (PDF / Image / Photo)"], horizontal=True)

        original_text = ""
        original_image = None
        uploaded_file = None

        if input_method == "Paste Text":
            original_text = st.text_area(
                "Original Lesson Text",
                value="1. Look at those dark clouds! It (rain) any minute now.\n\n2. Every Saturday morning, my brother (go) to the gym before breakfast.",
                height=200,
                placeholder="Paste the original lesson or word problem here..."
            )
        else:
            uploaded_file = st.file_uploader("Upload Worksheet (PDF, PNG, JPG)", type=["pdf", "png", "jpg", "jpeg"])
            if uploaded_file is not None and uploaded_file.type.startswith('image'):
                try:
                    original_image = PIL.Image.open(uploaded_file)
                    st.image(original_image, caption="Attached Visual Worksheet", use_container_width=True)
                    buffered = io.BytesIO()
                    original_image.save(buffered, format="JPEG")
                    st.session_state.original_image_b64 = base64.b64encode(buffered.getvalue()).decode('utf-8')
                except Exception as e:
                    st.error(f"Image load error: {e}")

    with st.container(border=True):
        st.subheader("⚙️ 2. Set Generation Parameters")

        worksheet_type = "English (Reading & Grammar)"
        target_language = "English"

        col_params1, col_params2 = st.columns(2)

        with col_params1:
            special_interest = st.text_input("Student's Special Interest", value="Minecraft")
            tone_style = st.selectbox("Narrative Tone", ["Standard", "Epic Adventure", "Humorous", "Spooky / Mystery", "Sci-Fi / Futuristic"])
            learning_objective = st.text_input("Grammar / Reading Objective", value="Tenses", placeholder="e.g., 'Passive Voice', 'Inferences'")

        with col_params2:
            output_layout = st.radio("Target Audience Formatting", ["Standard Professional Text (PDF/DOCX)", "Primary School Illustration (Full Graphic Image Worksheet)"])
            vocab_level = st.radio("Vocabulary Level", ["Keep Exact Same Level", "Simplify for Easier Reading", "Challenge with Advanced Vocab"])
            include_hints = st.checkbox("💡 Include Scaffolded Hints (Generate contextual theme clues)", value=False)
            
        model_name = "qwen-plus"

    submit_button = st.button("✨ Generate Custom Worksheet ✨", use_container_width=True, type="primary")

    if submit_button:
        if input_method == "Upload Document (PDF / Image / Photo)" and uploaded_file is not None:
            if uploaded_file.name.lower().endswith('.pdf'):
                try:
                    pdf_reader = PyPDF2.PdfReader(uploaded_file)
                    for page in pdf_reader.pages:
                        extracted = page.extract_text()
                        if extracted:
                            original_text += clean_extracted_text(extracted) + "\n"
                except Exception as e:
                    st.error(f"Failed to read PDF. Error: {e}")
                    original_text = ""
                    
        if not original_text.strip() and original_image is None:
            st.warning("Please provide the Original Lesson Text or upload a valid text/image Document.")
        elif not special_interest.strip():
            st.warning("Please specify the Student's Special Interest.")
        else:
            with st.spinner(f"Initiating {model_name} connection..."):
                DYNAMIC_SYSTEM_PROMPT = get_system_prompt(
                    special_interest=special_interest,
                    worksheet_type=worksheet_type,
                    tone_style=tone_style,
                    target_language=target_language,
                    vocab_level=vocab_level,
                    include_hints=include_hints
                )
                
                user_message = f"Student's Special Interest: {special_interest}\n"
                if learning_objective.strip():
                    user_message += f"Core Learning Objective: {learning_objective}\n"
                if original_text.strip():
                    user_message += f"\nOriginal Lesson Text:\n{original_text}\n"
                if original_image is not None:
                    user_message += "\nScan the original lesson instructions and questions exactly from the attached image."
                
                try:
                    st.subheader("Translating Live... ✍️")
                    response_container = st.empty()
                    translated_text = ""
                    
                    client = OpenAI(
                        api_key=api_key,
                        base_url="https://dashscope-intl.aliyuncs.com/compatible-mode/v1"
                    )
                    
                    messages = [{"role": "system", "content": DYNAMIC_SYSTEM_PROMPT}]
                    
                    if original_image is not None:
                        messages.append({
                            "role": "user",
                            "content": [
                                {"type": "text", "text": user_message},
                                {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{st.session_state.original_image_b64}"}}
                            ]
                        })
                        response = client.chat.completions.create(
                            model="qwen3-vl-plus",
                            messages=messages,
                            stream=True
                        )
                    else:
                        messages.append({"role": "user", "content": user_message})
                        response = client.chat.completions.create(
                            model="qwen-plus",
                            messages=messages,
                            stream=True
                        )
                        
                    for chunk in response:
                        if chunk.choices and chunk.choices[0].delta.content is not None:
                            translated_text += chunk.choices[0].delta.content
                            response_container.info(translated_text + "▌")
                            
                    response_container.info(translated_text)
                    st.session_state["translated_text"] = translated_text
                    st.session_state["original_text_display"] = original_text
                    st.session_state["target_language_str"] = target_language
                    st.session_state["show_pdf_lang_warning"] = True if target_language in ["Mandarin", "Japanese", "Korean"] else False
                    
                    if "audio_bytes" in st.session_state:
                        del st.session_state["audio_bytes"]
                    if "theme_image_bytes" in st.session_state:
                        del st.session_state["theme_image_bytes"]
                                
                except Exception as e:
                    st.error(f"An error occurred while calling the Dashscope Translation API: {e}")

            with st.spinner(f"Designing worksheet typography (via AI Image Generator)..."):
                try:
                    if output_layout == "Primary School Illustration (Full Graphic Image Worksheet)":
                        tone_param = "Incredibly cute, colorful, adorable primary school visual design"
                        prompt_text = st.session_state.get('translated_text', '')
                        if "**Answer Key**" in prompt_text:
                            prompt_text = prompt_text.split("**Answer Key**")[0]
                        elif "Answer Key" in prompt_text:
                            prompt_text = prompt_text.split("Answer Key")[0]
                            
                        img_prompt = f"Design a highly colorful, Primary School educational worksheet graphic. Aspect ratio 3:4 (Portrait paper). Theme: {special_interest}. Aesthetic: {tone_param}. You MUST neatly write the following exact lesson questions directly onto the image layout itself, using the coolest thematic artistic typography that perfectly matches {special_interest}! DO NOT INCLUDE ANSWERS. CRITICAL: Directly underneath EVERY SINGLE QUESTION, you MUST draw a giant, prominent, distinctly empty solid writing line (or massive empty thematic box) so the student actually has a clear physical space to write their answer!\n\nEXACT TEXT TO RENDER (Strict Text Grounding):\n\"\"\"\n{prompt_text[:800]}\n\"\"\""
                    else:
                        tone_param = tone_style
                        img_prompt = f"A beautiful educational worksheet header banner, highly aesthetic layout. Theme subject: {special_interest}. Artistic tone: {tone_param}. Style: minimalist vector art, white background. No text and no words. Wide 16:9 aspect ratio."
                    
                    message = Message(
                        role="user",
                        content=[
                            {
                                "text": img_prompt
                            }
                        ]
                    )

                    res = ImageGeneration.call(
                        api_key=api_key,
                        model="wan2.7-image-pro",
                        messages=[message],
                        n=1,
                        size='1024*1024'
                    )

                    if res.status_code == 200:
                        img_url = res.output.choices[0].message.content[0]["image"]
                        img_response = requests.get(img_url)
                        if img_response.status_code == 200:
                            st.session_state["theme_image_bytes"] = img_response.content
                        st.session_state["output_layout_mode"] = output_layout
                    else:
                        st.warning(f"Could not generate theme image: {res.message}")
                        st.session_state["output_layout_mode"] = output_layout
                except Exception as e:
                    st.warning(f"Could not generate theme image: {e}")

    if "translated_text" in st.session_state:
        if st.session_state.get("output_layout_mode") == "Primary School Illustration (Full Graphic Image Worksheet)":
            st.success("Primary School Canvas Synthesis Complete!")
            if "theme_image_bytes" in st.session_state:
                st.image(st.session_state["theme_image_bytes"], caption="Graphic Primary School Worksheet (AI Rendered)", use_container_width=True)
                st.download_button(
                    label="🖼️ Download Graphic Worksheet (.JPG)",
                    data=st.session_state["theme_image_bytes"],
                    file_name="primary_school_worksheet.jpg",
                    mime="image/jpeg"
                )
                st.divider()

        st.success("Translation and Formatting Complete!")
        st.subheader("Your Translated Lesson")
        res_col1, res_col2 = st.columns(2)
        with res_col1:
            st.markdown("**Original Text:**")
            st.info(st.session_state.get("original_text_display", ""))
        with res_col2:
            st.markdown("**Translated Lesson:**")
            st.success(st.session_state["translated_text"])
            
        st.divider()
        
        
        
        st.subheader("🎧 Listen to the Lesson")
        st.markdown("Generate an auditory version of this lesson. Excellent for students with reading difficulties or ESL learners.")
        
        if "audio_bytes" not in st.session_state:
            if st.button("Generate Audio Track 🔊"):
                with st.spinner("Synthesizing audio..."):
                    target_lang_str = st.session_state.get("target_language_str", "English (Default)")
                    lang_map = {
                        "Spanish": "es", "French": "fr", "German": "de", 
                        "Mandarin": "zh-CN", "Japanese": "ja"
                    }
                    lang_code = lang_map.get(target_lang_str, "en")
                    
                    try:
                        # 💡 安全检查：确保文本确实存在且不为空
                        text_to_speak = st.session_state.get("translated_text", "").strip()
                        if not text_to_speak:
                            st.error("Text is empty! Cannot generate audio.")
                            st.stop()

                        print("1111 - Text length:", len(text_to_speak))
                        tts = gTTS(text=text_to_speak, lang=lang_code, slow=False)
                        
                        print("2222 - Saving to temp file")
                        # 使用临时文件保存，彻底避免内存指针问题
                        with tempfile.NamedTemporaryFile(delete=False, suffix=".mp3") as fp:
                            temp_filename = fp.name
                            tts.save(temp_filename)  # 👈 改用 gTTS 自带的 save 方法
                        
                        print("3333 - Reading bytes")
                        # 从磁盘读取完整的二进制数据
                        with open(temp_filename, "rb") as f:
                            st.session_state["audio_bytes"] = f.read()
                        
                        print("4444 - Cleaning up temp file")
                        os.unlink(temp_filename)  # 删除临时文件，保持系统干净
                        
                        print("5555 - Rerun")
                        st.rerun() 
                    except Exception as e:
                        print("6666 - Failed")
                        st.error(f"Audio generation failed: {e}")
        else:
            st.success("Audio track generated successfully! 🎉")
        # 💡 修改这里：将字节转换为 base64 编码
            b64_audio = base64.b64encode(st.session_state["audio_bytes"]).decode()
            audio_html = f'<audio controls width="100%"><source src="data:audio/mp3;base64,{b64_audio}" type="audio/mp3"></audio>'
            st.markdown(audio_html, unsafe_allow_html=True)
                    
            if st.button("Clear Audio 🔄"):
                del st.session_state["audio_bytes"]
                st.rerun()

        st.divider()
        
        if st.session_state.get("output_layout_mode") != "Primary School Illustration (Full Graphic Image Worksheet)":
            st.subheader("💾 Export Media")
            
            parts = st.session_state["translated_text"].split("--- ANSWER KEY ---")
            st.markdown("##### Distribution Mode")
            student_only_mode = st.checkbox("Export 'Student Only' Files (Strips out the Teacher's Answer Key from the exported file)", value=False)
            
            export_text = parts[0].strip() if student_only_mode else st.session_state["translated_text"]
            
            action_col1, action_col2, action_col3 = st.columns(3)
            with action_col1:
                st.download_button(
                    label="📝 Download as Text File",
                    data=export_text,
                    file_name="translated_lesson.txt",
                    mime="text/plain"
                )
            with action_col2:
                if st.session_state.get("show_pdf_lang_warning"):
                    st.warning("PDF converter does not natively support Asian characters. Use the Text format for non-Latin sentences.")
                else:
                    try:
                        img_bytes = st.session_state.get("theme_image_bytes")
                        pdf_bytes = create_pdf(
                            text_content=export_text, 
                            image_bytes=img_bytes, 
                            special_interest=special_interest
                        )
                        
                        st.download_button(
                            label="📑 Download as Converted PDF",
                            data=bytes(pdf_bytes),
                            file_name="translated_lesson.pdf",
                            mime="application/pdf"
                        )
                    except Exception as e:
                        st.error(f"Could not generate PDF: {e}")

            with action_col3:
                try:
                    img_bytes = st.session_state.get("theme_image_bytes")
                    docx_bytes = create_docx(
                        text_content=export_text, 
                        image_bytes=img_bytes, 
                        special_interest=special_interest
                    )
                    st.download_button(
                        label="📘 Download as Word Document (.docx)",
                        data=docx_bytes,
                        file_name="translated_lesson.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                except Exception as e:
                    st.warning(f"DOCX exporter failed: {e}")

elif st.session_state.page == 'chinese':
    st.markdown("""
        <style>
        .stApp {  
            background-color: #ffffff;
            background-image: 
                radial-gradient(at 0% 0%, rgba(255, 230, 204, 0.5) 0px, transparent 50%), 
                radial-gradient(at 100% 0%, rgba(214, 239, 255, 0.5) 0px, transparent 50%),
                radial-gradient(at 100% 100%, rgba(213, 245, 227, 0.5) 0px, transparent 50%),
                radial-gradient(at 0% 100%, rgba(255, 204, 224, 0.5) 0px, transparent 50%);
        }
        /* 2. Top-Left Clover Element */
        .stApp::before {
            content: "☘️";
            position: absolute;
            top: 8%;
            left: 10%;
            font-size: 2.2rem;
            opacity: 0.35;
            filter: saturate(0.6);
            pointer-events: none;
            z-index: 0;
        }

        /* 3. Bottom-Left Rainbow Element */
        .stApp::after {
            content: "🌈";
            position: absolute;
            bottom: 18%;
            left: 2%;
            font-size: 3.5rem;
            opacity: 0.45;
            transform: rotate(-15deg);
            filter: saturate(0.6);
            pointer-events: none;
            z-index: 0;
        }

        /* 4. Secondary background nodes container injected globally into the main element */
        .stMainBlockContainer {
            position: relative;
        }

        /* Mid-Right Sparkle Element */
        .stMainBlockContainer::before {
            content: "✨";
            position: absolute;
            top: 38%;
            right: 32%;
            font-size: 1.8rem;
            opacity: 0.25;
            filter: saturate(0.6);
            pointer-events: none;
            z-index: 0;
        }

        /* Top-Right Purple Dot Element */
        .stMainBlockContainer::after {
            content: "🟣";
            position: absolute;
            top: 8%;
            right: 5%;
            font-size: 1.4rem;
            opacity: 0.25;
            filter: saturate(0.5);
            pointer-events: none;
            z-index: 0;
        }

        /* 5. Force foreground interactive modules to remain click-responsive on top */
        div[data-testid="stVerticalBlock"] {
            position: relative;
            z-index: 10 !important;
        }
        .main { background-color: #f0f4f8; }
        .stButton>button { width: 100%; border-radius: 20px; }
        .stTextArea>div>div>textarea { border-radius: 10px; }
        div[data-testid="stExpander"] { border: none; box-shadow: 0px 4px 12px rgba(0,0,0,0.05); }
        </style>
        """, unsafe_allow_html=True)

        # Use a container for better centering
    with st.container():
        # Row 1: Back Button (Takes full width or natural width)
        if st.button("⬅️ 返回", use_container_width=False):
            st.session_state.page = 'home'
            st.rerun()
        
        # Row 2: Title and Caption directly underneath
        st.title("📚 中文單元")
        st.caption("讀寫樂：學習中文詞彙")

    st.divider()

    tab1, tab2, tab3 = st.tabs(["✨ 生成詞語卡", "📄 上傳詞語工作紙", "⚔️ 分拆詞語"])

    with tab1:
  # Use an expander for status to keep the UI clean
        # with st.expander("⚙️ 系統狀態", expanded=False):
        #     if font_loaded:
        #         st.success("字型已就緒")
        #     else:
        #         st.warning("使用預設字型 (中文顯示可能異常)")
        st.markdown("### ✨ 生成詞語卡")
        st.info("輸入中文詞語然後點擊 開始生成，AI 將自動製作精美配圖字卡。")
        # Split Input and Actions
        col_input, col_action = st.columns([2, 1])
        
        with col_input:
            user_text = st.text_area("學習內容", placeholder="輸入想學習的中文詞語...", height=180)
        
        with col_action:
            st.write("### 操作面板")
            st.info("點擊下方按鈕開始 AI 創作圖片與拼音標註。")
            generate_btn = st.button("🎨 開始生成", type="primary")

        if generate_btn:
            if user_text:
                with st.status("🚀 正在製作您的學習卡...", expanded=True) as status:
                    st.write("🎨 AI 正在構思意境圖...")
                    # 1. 圖片生成
                    message = Message(
                        role="user",
                        content=[
                            {
                                "text": f"圖片創作：{user_text},圖片風格：繪本風格。"
                            }
                        ]
                    )

                    res = ImageGeneration.call(
                        api_key=api_key,
                        model="wan2.7-image-pro",
                        messages=[message],
                        n=1,
                        size='1024*1024'
                    )

                    if res.status_code == 200:
                        img_url = res.output.choices[0].message.content[0]["image"]
                        # 2. 拼音
                        pinyin_list = pinyin(user_text, style=Style.TONE)
                        pinyin_text = " ".join([item[0] for item in pinyin_list])

                        status.update(label="✅ 製作完成！", state="complete", expanded=True)
                        st.markdown("---")
                        st.subheader("🖼️ 生成結果預覽")
                        # 3. HTML 預覽
                        final_html = build_html_card(user_text, pinyin_text, img_url)
                        inner_col_1, inner_col_2, inner_col_3 = st.columns([1, 6, 1])
                        with inner_col_2:
                            st.components.v1.html(final_html, height=800, scrolling=True)

                        # 4. 產生真正 PDF
                        pdf_file = create_real_pdf(user_text, pinyin_text, img_url)
                        # Distinct Download Area
                        st.success("🎉 您的 PDF 已準備好！")
                        st.download_button(
                            label="📄 下載 PDF",
                            data=pdf_file,
                            file_name="chinese_flashcard.pdf",
                            mime="application/pdf",
                            use_container_width=True,
                            type="secondary"
                        )
                    else:
                        st.error(f"繪圖失敗: {res.message}")

    # --- 功能 B：PDF 識別意境並繪圖 ---
    with tab2:
        st.markdown("### 📄 上傳詞語工作紙")
        st.info("上傳包含中文詞語的 PDF，AI 將自動提取前詞語並製作精美配圖字卡。")

        up_pdf = st.file_uploader("選擇您的 PDF 文件", type="pdf", help="僅支援中文內容的 PDF")

        if up_pdf:
            if st.button("🚀 開始生成", type="primary", use_container_width=True):
                imgs = pdf_to_images(up_pdf.read())

                # Replace with your actual OpenAI client if already imported
                client = OpenAI(
                    api_key=api_key,
                    base_url="https://dashscope-intl.aliyuncs.com/compatible-mode/v1"
                )

                with st.status("🛠️ 正在啟動 AI 引擎...", expanded=True) as status:
                    # 1. OCR first page
                    st.write("🔍 **第一步：** 正在掃描文件並提取文字...")
                    response = client.chat.completions.create(
                        model="qwen3-vl-plus",
                        messages=[{
                            "role": "user",
                            "content": [
                                {"type": "text", "text": "请识别图中文字。"},
                                {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{imgs[0]}"}}
                            ]
                        }]
                    )

                    full_text = response.choices[0].message.content

                    chinese_chars = get_words_from_text(full_text)
                    limit_chars = chinese_chars[:6]

                    if not limit_chars:
                        st.error("❌ 抱歉，我們無法從此 PDF 中識別出中文字。")
                        status.update(label="處理失敗", state="error")
                    else:
                        st.write(f"✅ 成功提取到 {len(limit_chars)} 個漢字：`{', '.join(limit_chars)}`")
                        st.divider()

                        items_html = ""
                        cards_data = []

                        progress_bar = st.progress(0)

                        for idx, char in enumerate(limit_chars):
                            st.write(f"🎨 正在繪製「**{char}**」的意境圖...")

                            # Update progress bar
                            progress_bar.progress((idx + 1) / len(limit_chars))

                            message = Message(
                                role="user",
                                content=[
                                    {
                                        "text": f"圖片創作：{char},圖片風格：繪本風格。"
                                    }
                                ]
                            )

                            res = ImageGeneration.call(
                                api_key=api_key,
                                model="wan2.7-image-pro",
                                messages=[message],
                                n=1,
                                size='1024*1024'
                            )

                            if res.status_code == 200:
                                img_url = res.output.choices[0].message.content[0]["image"]

                                pinyin_list = pinyin(char, style=Style.TONE)
                                pinyin_text = " ".join([item[0] for item in pinyin_list])

                                cards_data.append({
                                    "char": char,
                                    "pinyin": pinyin_text,
                                    "img_url": img_url
                                })

                                if idx % 2 == 0:
                                    items_html += "<tr>"

                                items_html += f"""
                                <td style="width: 50%; padding: 5px; vertical-align: top;">
                                    <div style="
                                        text-align: center;
                                        border: 1px solid #eee;
                                        padding: 5px;
                                        border-radius: 5px;
                                        background-color: white;
                                    ">
                                        <img src="{img_url}"
                                            alt="字卡圖片"
                                            style="width: 150pt; height: 150pt; object-fit: cover; border-radius: 5px;">

                                        <div style="margin-top: 2px;">
                                            <div style="font-size: 14pt; color: #888;">{pinyin_text}</div>
                                            <div style="font-size: 14pt; font-weight: bold; color: #333;">{char}</div>

                                            <div class="no-print" style="margin-top: 2px;">
                                                <span style="cursor:pointer; color:#ff4b4b; font-size: 20px; margin-right: 10px;"
                                                    onclick="speak('{char}', 'zh-CN')">
                                                    <small style="font-size:12px;">国</small> 🔊
                                                </span>
                                                <span style="cursor:pointer; color:#007aff; font-size: 20px;"
                                                    onclick="speak('{char}', 'zh-HK')">
                                                    <small style="font-size:12px;">粤</small> 🔊
                                                </span>
                                            </div>
                                        </div>
                                    </div>
                                </td>
                                """

                                if idx % 2 == 1 or idx == len(limit_chars) - 1:
                                    if idx == len(limit_chars) - 1 and idx % 2 == 0:
                                        items_html += '<td style="width: 50%;"></td>'
                                    items_html += "</tr>"

                            else:
                                st.warning(f"字「{char}」繪圖失敗：{res.message}")

                        status.update(label="🎉 任務圓滿完成！", state="complete", expanded=True)

                if cards_data:
                    with st.status("🚀 正在製作 PDF 卡片...", expanded=True) as status:
                        pdf_file = create_cards_pdf(cards_data)
                        status.update(label="✅ PDF 製作完成！", state="complete", expanded=True)

                    col_prev, col_dl = st.columns([3, 1])
                    with col_prev:
                        st.subheader("🖼️ 生成結果預覽")
                    with col_dl:
                        # Place download button at the top right of results
                        st.download_button(
                            label="📄 下載 PDF",
                            data=pdf_file,
                            file_name="chinese_flashcards.pdf",
                            mime="application/pdf",
                            use_container_width=True
                        )
                    
                    # Big Preview
                    final_html = build_html_cards(items_html)
                    st.components.v1.html(final_html, height=800, scrolling=True)
                else:
                    st.error("未能生成任何內容。")
    with tab3:
        st.markdown("### ⚔️ 分拆詞語")

        # 表單輸入機制，中文打字流暢
        with st.form(key="hanzi_form"):
            user_text = st.text_input(
                "請輸入中文字或詞語（例如：明天）：", 
                value=""
            )
            submit_button = st.form_submit_button(label="🚀 開始拆解生字及生成拼音", use_container_width=True)

        def extract_colored_components_with_pinyin(text):
            """【核心升級】解析漢字向量與 EvenOdd 挖空的同時，在上方動態生成對齊的彩色拼音"""
            try:
                #font_prop = FontProperties(family=['Microsoft JhengHei', 'Microsoft YaHei', 'PingFang HK', 'Heiti TC', 'sans-serif'])
                # Initialize font properties explicitly from the path file
                font_prop = FontProperties(fname=font_path_2)

                # Apply it globally to Matplotlib to avoid fallback errors
                plt.rcParams['font.family'] = font_prop.get_name()
                plt.rcParams['axes.unicode_minus'] = False  # Properly renders negative signs '-'
                char_size = 1000
                char_spacing = 1100 
                
                # 鮮艷的高對比色彩池
                color_palette = ["#FF4B4B", "#1C83E1", "#00C092", "#9B5DE5", "#F15BB5", "#FFA500", "#FFD166", "#00BFFF", "#FF69B4", "#32CD32"]
                random.seed(sum(ord(c) for c in text))
                random.shuffle(color_palette)
                
                # 取得標準帶聲調的拼音列表（例如：['míng', 'tiān']）
                pinyin_list = [p[0] for p in pinyin(text, style=Style.TONE)]
                
                svg_paths = []
                svg_texts = []
                color_idx = 0
                
                for char_pos, char in enumerate(text):
                    if char.isspace():
                        continue
                        
                    tp = TextPath((0, 0), char, size=char_size, prop=font_prop)
                    polygons = tp.to_polygons()
                    x_offset = char_pos * char_spacing
                    
                    valid_polys = []
                    for poly in polygons:
                        if len(poly) >= 3:
                            valid_polys.append(poly)
                    
                    # 建立多邊形數據列表
                    poly_data_list = []
                    for poly in valid_polys:
                        x_list = [float(pt[0]) for pt in poly]
                        y_list = [float(pt[1]) for pt in poly]
                        cx = sum(x_list) / len(x_list)
                        cy = sum(y_list) / len(y_list)
                        approx_area = (max(x_list) - min(x_list)) * (max(y_list) - min(y_list))
                        
                        poly_data_list.append({
                            'poly': poly,
                            'center': (cx, cy),
                            'area': approx_area
                        })
                    
                    # 按面積從大到小排序，精確進行外框與內部中空的重疊判定
                    poly_data_list.sort(key=lambda x: x['area'], reverse=True)
                    
                    used = [False] * len(poly_data_list)
                    groups = []
                    
                    for i in range(len(poly_data_list)):
                        if used[i]:
                            continue
                        current_group = [poly_data_list[i]['poly']]
                        used[i] = True
                        
                        path_i = plt.Polygon(poly_data_list[i]['poly'])
                        for j in range(len(poly_data_list)):
                            if not used[j]:
                                if path_i.get_path().contains_point(poly_data_list[j]['center']):
                                    current_group.append(poly_data_list[j]['poly'])
                                    used[j] = True
                        groups.append(current_group)
                    
                    # 記錄當前漢字的主偏旁顏色，用來同步上方的拼音
                    char_main_color = color_palette[color_idx % len(color_palette)]
                    
                    # --- 渲染生成漢字的 SVG 路徑 ---
                    for group in groups:
                        char_svg_segments = []
                        for poly in group:
                            path_data = []
                            for idx, pt in enumerate(poly):
                                px = float(pt[0]) + x_offset
                                py = float(pt[1])
                                if idx == 0:
                                    path_data.append(f"M {px:.1f} {py:.1f}")
                                else:
                                    path_data.append(f"L {px:.1f} {py:.1f}")
                            path_data.append("Z")
                            char_svg_segments.append(" ".join(path_data))
                        
                        color = color_palette[color_idx % len(color_palette)]
                        full_d = " ".join(char_svg_segments)
                        # 複合路徑 + evenodd 保持格子完美透明
                        svg_paths.append(f'<path d="{full_d}" style="fill: {color}; fill-rule: evenodd; stroke: {color}; stroke-width: 1;" />')
                        color_idx += 1
                    
                    # --- 【全新升級：SVG 拼音渲染】 ---
                    # 獲取當前漢字的拼音
                    current_pinyin = pinyin_list[char_pos] if char_pos < len(pinyin_list) else ""
                    if current_pinyin:
                        # 計算該漢字的水平中心點，讓拼音完美居中對齊
                        # 每個漢字寬度基底約為 1000，加上 x_offset，中心點在 x_offset + 500
                        pinyin_x = x_offset + 500
                        # y 座標設在 1150（相對於漢字頂部 y=1000 向上位移，預留拼音空間）
                        # 由於整體畫布會垂直翻轉，這裡的正負值配合 transform 進行渲染
                        svg_texts.append(
                            f'<text x="{pinyin_x}" y="-1100" '
                            f'fill="{char_main_color}" font-family="Arial, sans-serif" font-size="200" font-weight="bold" '
                            f'text-anchor="middle" transform="scale(1, -1)">{current_pinyin}</text>'
                        )
                        
                if not svg_paths:
                    return None, None, None, None
                    
                total_chars = len(text)
                total_width = total_chars * char_spacing + 100
                # 調整 ViewBox 高度至 1600 (給上方拼音留出 0-400 的空間)
                view_box = f"0 -450 {total_width} 1650"
                transform_str = f"scale(1, -1) translate(50, -850)"
                
                # 合併漢字路徑與拼音文字
                full_svg_content = "\n".join(svg_paths) + "\n" + "\n".join(svg_texts)
                return full_svg_content, view_box, transform_str, total_width
            except Exception as e:
                st.error(f"字型解剖或拼音解析失敗: {e}")
                return None, None, None, None

        def generate_pdf(text, svg_content, view_box, transform_str, total_width):
            display_width = min(500, total_width * 0.3)
            display_height = display_width * (1550 / total_width)
            
            pure_svg = f"""<svg xmlns="http://w3.org" viewBox="{view_box}" width="{display_width}" height="{display_height}">
                <line x1="0" y1="600" x2="{total_width}" y2="600" stroke="#DDD" stroke-width="2" stroke-dasharray="10,10"/>
                <g transform="{transform_str}">
                    {svg_content}
                </g>
            </svg>"""
            
            svg_io = io.BytesIO(pure_svg.encode('utf-8'))
            drawing = svg2rlg(svg_io)
            
            pdf_buffer = io.BytesIO()
            pdf_canvas = canvas.Canvas(pdf_buffer, pagesize=letter)
            
            x_pos = (612 - display_width) / 2
            y_pos = (792 - display_height) / 2
            
            from reportlab.graphics import renderPDF
            renderPDF.draw(drawing, pdf_canvas, x_pos, y_pos)
            
            pdf_canvas.showPage()
            pdf_canvas.save()
            pdf_buffer.seek(0)
            return pdf_buffer.getvalue()

        # 主程式運行
        if user_input := user_text.strip():
            if any('\u4e00' <= c <= '\u9fff' for c in user_input):
                with st.spinner("正在進行標準字形與拼音動態渲染..."):
                    svg_content, view_box, transform_str, total_width = extract_colored_components_with_pinyin(user_input)
                    
                if svg_content:
                    st.subheader(f"「{user_input}」拆字結果如下：")
                    
                    # 網頁 HTML 容器
                    html_code = f"""
                    <div style="background-color: #FFFFFF; padding: 20px; border-radius: 16px; display: flex; justify-content: center; align-items: center; min-height: 280px; border: 1px solid #333; overflow-x: auto;">
                        <svg xmlns="http://w3.org" viewBox="{view_box}" style="height: 220px; width: auto; max-width: 100%;">
                            <!-- 背景暗色中線輔助線 -->
                           
                            <g transform="{transform_str}">
                                {svg_content}
                            </g>
                        </svg>
                    </div>
                    """
                    st.components.v1.html(html_code, height=360, scrolling=False)
                    
                    # PDF 下載按鈕
                    pdf_data = generate_pdf(user_input, svg_content, view_box, transform_str, total_width)
                    st.download_button(
                        label=f"📥 下載「{user_input}」拆字卡",
                        data=pdf_data,
                        file_name=f"hanzi_pinyin_{user_input}.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )
                    st.caption("<div style='text-align:center; color:#888; margin-top:5px;'>💡 提示：字體中空 100% 透明。上方拼音字體顏色會自動與該漢字的第一個主偏旁色彩完美同步！</div>", unsafe_allow_html=True)
                else:
                    st.error("❌ 無法提取該詞語結構。")
            else:
                st.error("❌ 請輸入有效的中文內容！")

elif st.session_state.page == 'math':
    st.markdown("""
        <style>
        .stApp {  
            background-color: #ffffff;
            background-image: 
                radial-gradient(at 0% 0%, rgba(255, 230, 204, 0.5) 0px, transparent 50%), 
                radial-gradient(at 100% 0%, rgba(214, 239, 255, 0.5) 0px, transparent 50%),
                radial-gradient(at 100% 100%, rgba(213, 245, 227, 0.5) 0px, transparent 50%),
                radial-gradient(at 0% 100%, rgba(255, 204, 224, 0.5) 0px, transparent 50%);
        }
        /* 2. Top-Left Clover Element */
        .stApp::before {
            content: "☘️";
            position: absolute;
            top: 8%;
            left: 10%;
            font-size: 2.2rem;
            opacity: 0.35;
            filter: saturate(0.6);
            pointer-events: none;
            z-index: 0;
        }

        /* 3. Bottom-Left Rainbow Element */
        .stApp::after {
            content: "🌈";
            position: absolute;
            bottom: 18%;
            left: 2%;
            font-size: 3.5rem;
            opacity: 0.45;
            transform: rotate(-15deg);
            filter: saturate(0.6);
            pointer-events: none;
            z-index: 0;
        }

        /* 4. Secondary background nodes container injected globally into the main element */
        .stMainBlockContainer {
            position: relative;
        }

        /* Mid-Right Sparkle Element */
        .stMainBlockContainer::before {
            content: "✨";
            position: absolute;
            top: 38%;
            right: 32%;
            font-size: 1.8rem;
            opacity: 0.25;
            filter: saturate(0.6);
            pointer-events: none;
            z-index: 0;
        }

        /* Top-Right Purple Dot Element */
        .stMainBlockContainer::after {
            content: "🟣";
            position: absolute;
            top: 8%;
            right: 5%;
            font-size: 1.4rem;
            opacity: 0.25;
            filter: saturate(0.5);
            pointer-events: none;
            z-index: 0;
        }

        /* 5. Force foreground interactive modules to remain click-responsive on top */
        div[data-testid="stVerticalBlock"] {
            position: relative;
            z-index: 10 !important;
        }
        .main { background-color: #f0f4f8; }
        .stButton>button { width: 100%; border-radius: 20px; }
        .stTextArea>div>div>textarea { border-radius: 10px; }
        div[data-testid="stExpander"] { border: none; box-shadow: 0px 4px 12px rgba(0,0,0,0.05); }
        </style>
        """, unsafe_allow_html=True)
   # --- 1. Sidebar Assistant Styling ---
    with st.sidebar:
        st.title("💬 AI Assistant")
        st.caption("Use our Assistant to adjust font size, modify question themes, and customise content presentation.")

        chat_container = st.container(height=450)
        with chat_container:
            for msg in st.session_state.chat_history:
                with st.chat_message(msg["role"]):
                    st.markdown(msg["content"])

        # CHATBOT ENTER TRIGGER
        if prompt := st.chat_input("Enter prompt and click Enter for execution..."):
            # 1. 記錄使用者的話
            st.session_state.chat_history.append({"role": "user", "content": prompt})
            # 2. 讓 AI 助理給予簡單回覆 (例如：確認收到指令)
            assistant_reply = f"Got it！I will execute the command ** [{prompt}] ** for you right now."
            st.session_state.chat_history.append({"role": "assistant", "content": assistant_reply})
            # KEY: Set the trigger to True and rerun
            if st.session_state.current_mode == "generator_mode":
                st.session_state.trigger_generate = True 
                st.session_state.trigger_analysis = False 
                st.session_state.my_tabs = "📄 Worksheet Genie"
                st.session_state.show_wizard = False
                st.rerun()
            else:
                st.session_state.trigger_analysis = True 
                st.session_state.trigger_generate = False 
                st.session_state.my_tabs = "📝 Upload math questions"
                st.rerun() 

    with st.container():
        # Row 1: Back Button on top
        if st.button("⬅️ Back", use_container_width=False):
            st.session_state.page = 'home'
            st.session_state.show_wizard = False
            st.rerun()
        
        # Row 2: Title and Information Block directly underneath
        st.title("🔢 Math Module")
        st.info("Help teachers transform standard math worksheets into engaging, thematic learning materials for students.")


    st.divider()

    tab1, tab2 = st.tabs(["📝 Upload math questions", "📄 Worksheet Genie"], key="my_tabs", 
    on_change=handle_tab_change)

    with tab1:
        # Step 1: Theme Selection
        st.write("### 🏗️ Step 1: Select Your Theme")
        theme_options = ["Ocean Adventure", "Forest Adventure", "Cooking", "Animal Farm", "Superhero", "Dinosaur World"]
        
        if "selected_theme" not in st.session_state:
            st.session_state.selected_theme = "Ocean Adventure"
            
        current_index = theme_options.index(st.session_state.selected_theme)
        
        st.radio(
            "Select your theme：", 
            theme_options, 
            index=current_index,
            key="theme_radio_widget",
            horizontal=True,
            on_change=lambda: st.session_state.update({"selected_theme": st.session_state.theme_radio_widget})
        )

        # Step 2: File Upload
        st.write("### 📤 Step 2: Upload worksheet")
        if "pdf_buffer" not in st.session_state:
            st.session_state.pdf_buffer = None
            
        math_pdf = st.file_uploader("Please upload your PDF file", type="pdf", label_visibility="collapsed")
        
        if math_pdf:
            st.session_state.pdf_buffer = math_pdf.getvalue()
            st.success(f"✅ The file is ready： {math_pdf.name}")

        # Core Scanning Function
        def perform_scan(show_response_in_chat):
            # CRITICAL: Check the session state buffer, not the widget variable
            if st.session_state.pdf_buffer is None:
                st.warning("Please upload PDF first！")
                return

            with st.status("🚀 AI is analysing the worksheet...", expanded=True) as status:
                st.write("🔍 Processing...")
                # Fetch instruction from chat history
                latest_instruction = ""
                if st.session_state.chat_history:
                    raw_msg = st.session_state.chat_history[-1]["content"]
                    # Extract text inside [brackets] if user used them, else take full message
                    latest_instruction = raw_msg.split('[')[-1].split(']')[0] if '[' in raw_msg else raw_msg

                pdf_bytes = st.session_state.pdf_buffer
                imgs = pdf_to_images(pdf_bytes)
                st.session_state.math_results = []
                
                client = OpenAI(api_key=api_key, base_url="https://dashscope-intl.aliyuncs.com/compatible-mode/v1")
                current_theme = st.session_state.selected_theme
                print('select theme:')
                print(current_theme)
                print('latest_instruction:')
                print(latest_instruction)
                prompt = f"任務： 識別圖片或 PDF 文本中的表格，並根據選擇的主題 {current_theme} 重寫問題。\
                            核心目標：\
                            將抽象、枯燥或帶有負面意義的數學題目（如「不及格」、「抽球」、「壞掉」），轉化為 SEN 學生（小學生）感興趣、具備正面激勵且具體化的場景。\
                            主題情境對照表 (Reference Mapping)：\
                            請嚴格根據 {current_theme} 的設定進行情境轉化：\
                            主題 (Theme)\
                            學科 A (Maths)\
                            學科 B (Physics)\
                            負面/失敗詞彙轉化 (Failed/Defective)\
                            物件轉化 (Balls/Coins/Cards)\
                            海洋冒險\
                            珊瑚計數\
                            逆流游泳\
                            在洋流中迷失 / 需要重新定位\
                            小丑魚 / 藍唐王魚 / 海星\
                            森林探險\
                            地圖閱讀\
                            洞穴導航\
                            錯過路徑標記 / 暫時休息\
                            金鑰匙 / 銀指南針 / 寶石\
                            美食烹飪\
                            披薩配料\
                            湯品調味\
                            菜餚需要調整口味 / 重新烘焙\
                            草莓 / 棉花糖 / 巧克力豆\
                            動物農場\
                            餵食時間\
                            寵物美容\
                            寵物需要更多悉心照顧\
                            小狗 / 小貓 / 倉鼠\
                            超級英雄\
                            能量控制\
                            護盾防禦\
                            能量蓄勢中 / 訓練未達標\
                            能量球 / 英雄徽章 / 技能卡\
                            恐龍世界\
                            尋找化石\
                            奔跑速度\
                            遭遇火山灰遮擋 / 偵測中\
                            恐龍蛋 / 腳印化石 / 琥珀\
                            寫作指導原則：\
                            情境重構 (Scenario Building)： 禁止生硬替換單詞（例如嚴禁出現 Math-fish 或 failed-dino）。請重新編寫句子，使其符合兒童文學的敘事風格。\
                            正面語言 (Positive Language)： 將「不及格」或「瑕疵」改為「挑戰中」、「維修中」或「需要更多練習」，減少 SEN 學生的挫敗感。\
                            邏輯嚴謹性： 所有的數值、概率關係（交集、並集、條件概率）必須與原題 100% 一致。\
                            語言要求： 輸出語言為英文 (English)，不能出現中文 (Chinese)。請使用基礎英語詞彙 (CEFR A1/A2 水準)，句子結構簡單直接。\
                            輸出格式規範：\
                            轉換為 HTML <table> 格式。\
                            必須保留原有的合併單元格架構 (rowspan / colspan)。\
                            將「問題內容」與「問題編號/選項」置於同一單元格；「答案輸入區」置於獨立單元格.問題單元格同答案輸入區單元格寬度比例為2:1。\
                            禁止輸出任何答案。\
                            數學表達：使用標準文本或 Unicode（如 1/2, 1/4），嚴禁使用 LaTeX。\
                            僅輸出 HTML 代碼，不需任何解釋文字。\
                            強制規則：\嚴禁輸出中文。問題區單元格同答案輸入區單元格寬度比例為2:1！！問題區單元格同答案輸入區的背景顏色為白色！！名字欄同ID欄應該在同一行！！\
                            然後，{latest_instruction}。\
                            "

                for i, img_b64 in enumerate(imgs):
                    st.write(f"✍️ Generating content of page {i+1} ...")
                    response = client.chat.completions.create(
                        model="qwen3-vl-plus",
                        messages=[{"role": "user", "content": [{"type": "text", "text": prompt}, {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{img_b64}"}}]}]
                    )
                    res_html = response.choices[0].message.content.replace("```html", "").replace("```", "").strip()
                    st.session_state.math_results.append(res_html)
                
                status.update(label="✅ Done！", state="complete", expanded=True)
            
            if show_response_in_chat:
                done_msg = "✅ Done! The questions have been revised and displayed on the right for you to preview or download."
                st.session_state.chat_history.append({"role": "assistant", "content": done_msg})


            st.session_state.trigger_analysis = False
            st.rerun() # 再次重整以顯示最新的 AI 訊息

        # Step 3: Execution Controls
        st.divider()
        col_btn_1, col_btn_2 = st.columns(2)
        with col_btn_1:
            if st.button("🚀 Click to process your file and generate the thematic worksheet", type="primary", use_container_width=True):
                perform_scan(False)
        
        # Trigger from Chatbot
        if st.session_state.get('trigger_analysis', False):
            print("tab1")
            perform_scan(True)

        # --- 4. Results Display ---
        if st.session_state.math_results:
            st.subheader("🖥️ Preview")
            
            # Container for the HTML content
            with st.container(border=True):
                full_html = get_render_wrapper("".join(st.session_state.math_results))
                st.components.v1.html(full_html, height=800, scrolling=True)
            
            pdf_bytes = HTML(string=full_html).write_pdf()

            st.download_button(
                label="📥 Download PDF worksheet", 
                data=pdf_bytes, 
                file_name=f"math_{st.session_state.selected_theme}.pdf",
                mime="application/pdf",
                use_container_width=True
            )
            # Download Button
            # st.download_button(
            #     label="📥 下載 HTML 學習試卷",
            #     data=full_html,
            #     file_name=f"math_{st.session_state.selected_theme}.html",
            #     mime="text/html",
            #     use_container_width=True
            # )

    with tab2:
         # --- 1. Global Wizard Styling ---
        st.markdown("""
            <style>
            /* Dialog Title & Markdown Headers */
            div[data-testid="stDialog"] h1, 
            div[data-testid="stDialogHeader"] div,
            .stMarkdown h3 { 
                color: #7029b1 !important; 
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            }
            /* Progress Bar Color */
            div[data-testid="stProgressValue"] {
                background-color: #7029b1 !important;
            }
            /* Wizard Buttons Style */
            .stButton>button {
                width: 100%;
                border-radius: 20px;
                height: 3.5em;
                background-color: white;
                color: #7029b1;
                border: 1px solid #e0d5ed;
                font-weight: 600;
                transition: all 0.3s ease;
            }
            .stButton>button:hover {
                border-color: #7029b1;
                background-color: #f9f4ff;
                box-shadow: 0 4px 10px rgba(112, 41, 177, 0.1);
            }
            </style>
        """, unsafe_allow_html=True)

        def perform_generation(show_response_in_chat,lastest_instruction):
                        # 核心：生成题目并存储
            if st.session_state.selections == {}:
                return
            
            with st.status("🚀 AI is generating questions...", expanded=True) as status:
                if 'Simple Word Problems & Computation' in st.session_state.selections['question_type'] or 'Advanced' in st.session_state.selections['difficulty']:
                    questions = call_qwen_ai_1(st.session_state.selections,lastest_instruction)
                    st.session_state.generated_questions = questions
                    status.update(label="✅ Done！", state="complete", expanded=True)
                else:
                    questions = call_qwen_ai_2(st.session_state.selections,lastest_instruction)
                    st.session_state.generated_questions = questions
                    status.update(label="✅ Done！", state="complete", expanded=True)

                if show_response_in_chat:
                    done_msg = "✅ Done! The questions have been revised and displayed on the right for you to preview or download."
                    st.session_state.chat_history.append({"role": "assistant", "content": done_msg})

                st.session_state.show_wizard = False
                st.session_state.wizard_complete = True
                st.session_state.trigger_generate = False
                st.session_state.step = 1
                st.rerun()

        def handle_dialog_close():
            """Fires automatically when the user clicks the X button or closes the dialog."""
            st.session_state.show_wizard = False
            st.session_state.wizard_complete = True
            st.session_state.trigger_generate = False
            st.session_state.step = 1


        @st.dialog("🎯 Generate Questions", on_dismiss=handle_dialog_close)
        def question_wizard():
             # Initialize internal state if not exists
            if 'step' not in st.session_state: st.session_state.step = 1
            if 'selections' not in st.session_state: st.session_state.selections = {}

            step = st.session_state.step

            # --- Top Navigation & Progress ---
            col_back, col_step = st.columns([1, 3])
            with col_back:
                if step > 1:
                    if st.button("⬅️ Back", key="back_btn"):
                        # Custom Logic for Step Jumping
                        if step == 7 and  "No" in st.session_state.selections.get('need_theme'):
                            st.session_state.step = 5
                        else:
                            st.session_state.step -= 1
                        st.rerun()
            with col_step:
                st.write(f"Step： {int(step)} / 7 ")
                st.progress(min(step / 7, 1.0))

            st.divider()

            if step == 1:
                st.write("### What math topic should these focus on?")
                topics = ["➕ Addition / Subtraction", "✖️ Multiplication / Division", "📐 Area and Perimeter"]
                for topic in topics:
                    if st.button(topic, use_container_width=True):
                        st.session_state.selections['topic'] = topic
                        st.session_state.step = 2
                        st.rerun()

            elif step == 2:
                st.write("### What difficulty level?")
                difficuties = ["📚 Introductory", "🚀 Advanced"]
                for difficulty in difficuties:
                    if st.button(difficulty, use_container_width=True):
                        st.session_state.selections['difficulty'] = difficulty
                        st.session_state.step = 3
                        st.rerun()
            
            elif step == 3:
                st.write("### How will students use these questions?")
                test_types = ["📄 Worksheet", "🎟️ Quick Exit Ticket"]
                for test_type in test_types:
                    if st.button(test_type, use_container_width=True):
                        st.session_state.selections['test_type'] = test_type
                        st.session_state.step = 4
                        st.rerun()
            
            elif step == 4:
                st.write("### Should these focus on any specific student needs?")
                question_types = ["📝 Simple Word Problems & Computation", "🎨 Visuals / Shapes Included"]
                for question_type in question_types:
                    if st.button(question_type, use_container_width=True):
                        st.session_state.selections['question_type'] = question_type
                        st.session_state.step = 5
                        st.rerun()
            
            elif step == 5:
                st.write("### Any specific theme you want to apply for question description?")
                need_themes = ["✅ Yes", "❌ No"]
                for need_theme in need_themes:
                    if st.button(need_theme, use_container_width=True):
                        st.session_state.selections['need_theme'] = need_theme
                        st.session_state.step = 6
                        st.rerun()

            elif step == 6:
                 if "Yes" in st.session_state.selections['need_theme']:
                    st.write("### Select a theme:")
                    theme_types = ["🌊 Ocean Adventure", "🌴 Forest Adventure", "🥣 Cooking", "🦸‍♂️ Superhero", "🐑 Animal Farm"]
                    for theme_type in theme_types:
                        if st.button(theme_type, use_container_width=True):
                            st.session_state.selections['theme_type'] = theme_type
                            st.session_state.step = 7
                            st.rerun()
                 else:
                    st.session_state.step = 7
                    st.rerun()
            
            elif step >= 7:
                st.success("📝 Everything is set up!")
                st.info(f"Coming right up: {st.session_state.selections['difficulty']} - {st.session_state.selections['topic']}")

                if st.button("Generate Now", type="primary", use_container_width=True):
                    perform_generation(False,'')



        # --- MAIN INTERFACE ---
        st.markdown("### 📄 Question Generator")
        st.write("Create personalized math worksheets through an interactive guided conversation with our AI")

        col_start, _ = st.columns([1, 2])
        with col_start:
            if st.button("🔮 Start Guided Generation", type="primary"):
                st.session_state.show_wizard = True

        if st.session_state.get("show_wizard"):
            question_wizard()
        
        if st.session_state.get("trigger_generate"):
            print("tab2")
            latest_instruction = ""
            if st.session_state.chat_history:
                raw_msg = st.session_state.chat_history[-1]["content"]
                # Extract text inside [brackets] if user used them, else take full message
                latest_instruction = raw_msg.split('[')[-1].split(']')[0] if '[' in raw_msg else raw_msg
                perform_generation(True,latest_instruction)

        # --- DISPLAY RESULTS ---
        if st.session_state.get("generated_questions"):
            st.divider()
            st.subheader("🖥️ Preview")
            with st.container(border=True):
                st.components.v1.html(st.session_state.generated_questions, height=800, scrolling=True)
            
            pdf_bytes = convert_html_to_pdf(st.session_state.generated_questions)

            st.download_button(
                label="📥 Download PDF worksheet", 
                data=pdf_bytes, 
                file_name="math_worksheet.pdf",
                mime="application/pdf",
                use_container_width=True
            )