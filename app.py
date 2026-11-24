import os
import tempfile

import streamlit as st
import whisper
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ============================================
# Streamlit 頁面設定
# ============================================
st.set_page_config(
    page_title="MP3 語音轉文字 → Word",
    page_icon="🎧",
    layout="centered"
)

st.title("🎧 MP3 語音轉文字 → Word")
st.write("上傳 MP3 檔案，自動轉錄為文字，並下載 Word (.docx) 檔案。")

# ============================================
# 快取 Whisper 模型，避免每次互動都重新載入
# ============================================
@st.cache_resource(show_spinner="正在載入 Whisper 模型……")
def load_whisper_model(model_name: str):
    return whisper.load_model(model_name)

# ============================================
# 側邊欄設定
# ============================================
with st.sidebar:
    st.header("轉錄設定")

    model_name = st.selectbox(
        "Whisper 模型",
        options=["base", "small", "medium", "large-v3"],
        index=0,
        help="base 較快；medium / large-v3 較準確，但需要更多時間和資源。"
    )

    language = st.selectbox(
        "語言",
        options=["zh", "en", "auto"],
        index=0,
        help="zh = 中文／粵語；auto = 自動偵測語言。"
    )

    st.info(
        "建議：\n\n"
        "- 一般錄音：base\n"
        "- 更準確：small 或 medium\n"
        "- 最高準確度：large-v3，但速度較慢"
    )

# ============================================
# 上傳 MP3
# ============================================
uploaded_file = st.file_uploader(
    "請選擇 MP3 語音檔案",
    type=["mp3"],
    accept_multiple_files=False
)

if uploaded_file is None:
    st.info("請上傳一個 MP3 檔案以開始轉錄。")
    st.stop()

st.success(f"已選擇檔案：{uploaded_file.name}")

# ============================================
# 開始轉錄
# ============================================
if st.button("開始語音轉文字", type="primary", use_container_width=True):
    model = load_whisper_model(model_name)

    # 將上傳檔案暫存為 MP3
    with tempfile.NamedTemporaryFile(
        delete=False,
        suffix=".mp3"
    ) as tmp_file:
        tmp_file.write(uploaded_file.getvalue())
        mp3_path = tmp_file.name

    try:
        with st.spinner("正在轉換語音為文字，請稍候……"):
            transcribe_language = None if language == "auto" else language

            result = model.transcribe(
                mp3_path,
                language=transcribe_language,
                task="transcribe",
                verbose=False
            )

        transcript = result["text"].strip()

        if not transcript:
            st.warning("未能從音訊中辨識出文字，請確認 MP3 是否包含清晰語音。")
            st.stop()

        # ========================================
        # 顯示轉錄結果
        # ========================================
        st.subheader("轉錄結果")
        st.text_area(
            "語音文字內容",
            value=transcript,
            height=350
        )

        # ========================================
        # 生成 Word 檔
        # ========================================
        doc = Document()

        style = doc.styles["Normal"]
        style.font.name = "Microsoft JhengHei"
        style.font.size = Pt(12)

        title = doc.add_heading("語音轉文字紀錄", level=0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER

        doc.add_paragraph()
        doc.add_paragraph(f"來源音訊檔：{uploaded_file.name}")

        doc.add_paragraph()
        doc.add_paragraph("轉錄內容：")
        doc.add_paragraph()

        paragraphs = [
            paragraph.strip()
            for paragraph in transcript.split("\n")
            if paragraph.strip()
        ]

        if not paragraphs:
            paragraphs = [transcript]

        for paragraph in paragraphs:
            doc.add_paragraph(paragraph)

        word_filename = os.path.splitext(uploaded_file.name)[0] + ".docx"

        with tempfile.NamedTemporaryFile(
            delete=False,
            suffix=".docx"
        ) as tmp_word_file:
            word_path = tmp_word_file.name
            doc.save(word_path)

        # ========================================
        # 下載 Word 檔
        # ========================================
        with open(word_path, "rb") as word_file:
            word_bytes = word_file.read()

        st.download_button(
            label="下載 Word 檔",
            data=word_bytes,
            file_name=word_filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True
        )

        st.success("轉錄完成，請點擊上方按鈕下載 Word 檔。")

    except Exception as e:
        st.error(f"轉錄時發生錯誤：{e}")

    finally:
        # 清理暫存檔案
        if os.path.exists(mp3_path):
            os.remove(mp3_path)
