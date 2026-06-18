import streamlit as st
import anthropic
import json
import base64
import tempfile
import os
from pathlib import Path

st.set_page_config(
    page_title="住宅アフター点検 報告書生成ツール",
    page_icon="🏠",
    layout="wide",
)

SECTIONS = [
    "外壁", "その他外部", "玄関(内装)", "各室(内装)",
    "食堂・台所(設備)", "食堂・台所(内装)", "床下点検口",
    "洗面・脱衣所(設備)", "洗面・脱衣所(内装)", "浴室(設備)",
    "浴室(内装)", "トイレ(設備)", "トイレ(内装)",
    "廊下・階段(内装)", "建具",
]

st.markdown("""
<style>
.header-box {
    background: #1F4E79;
    color: white;
    padding: 18px 24px;
    border-radius: 8px;
    margin-bottom: 24px;
}
.header-box h1 { margin: 0; font-size: 22px; }
.header-box p  { margin: 4px 0 0; font-size: 13px; opacity: 0.85; }
</style>
<div class="header-box">
  <h1>🏠 住宅アフター点検 報告書生成ツール</h1>
  <p>写真・PDF・メモをアップロードしてExcel報告書を自動生成</p>
</div>
""", unsafe_allow_html=True)

with st.form("inspection_form"):

    COMPANY_NAME = "SODESIGN株式会社"
    company_name = COMPANY_NAME
    company_no   = ""

    st.subheader("📋 基本情報")
    c1, c2 = st.columns(2)
    with c1:
        property_name  = st.text_input("登録物件名",       placeholder="例：T邸")
        st.text_input("事業者名", value=COMPANY_NAME, disabled=True)
        delivery_date  = st.text_input("引渡し年月",        placeholder="例：2024年11月")
        inspector      = st.text_input("報告者名",          placeholder="例：山田　太郎")
    with c2:
        property_no     = st.text_input("登録番号（Y番号）", placeholder="例：Y3284262")
        inspection_date = st.date_input("実施日")
        inspection_term = st.selectbox(
            "点検日程",
            ["1年", "2年", "5年", "10年", "15年", "25年", "30年"],
        )

    st.subheader("📷 現場写真")
    photos = st.file_uploader(
        "写真をアップロード（複数可・JPG / PNG / HEIC 対応）",
        type=["jpg", "jpeg", "png", "gif", "webp", "heic", "heif"],
        accept_multiple_files=True,
    )
    if photos:
        cols = st.columns(min(len(photos), 5))
        for i, photo in enumerate(photos):
            with cols[i % 5]:
                st.image(photo, caption=photo.name, use_container_width=True)

    st.subheader("📄 点検報告書 PDF（任意）")
    pdf_file = st.file_uploader("PDFをアップロード", type=["pdf"])

    st.subheader("📝 担当者メモ")
    notes = st.text_area(
        "気になる点・補足情報",
        placeholder="例：2階廊下の壁クロスに浮きあり。玄関ドアの開閉が重い。",
        height=100,
    )

    submitted = st.form_submit_button(
        "📊 Excel報告書を生成する",
        use_container_width=True,
        type="primary",
    )

if submitted:
    api_key = os.getenv("ANTHROPIC_API_KEY", "")
    if not api_key:
        try:
            api_key = st.secrets["ANTHROPIC_API_KEY"]
        except Exception:
            pass
    if not api_key:
        st.error("ANTHROPIC_API_KEY が設定されていません。Streamlit Cloud の Secrets に登録してください。")
        st.stop()

    with st.spinner("AIが写真・PDFを解析中... しばらくお待ちください"):
        try:
            # --- 写真処理 ---
            photo_data = []
            temp_photo_paths = []

            for photo in (photos or []):
                photo_bytes = photo.read()
                b64 = base64.standard_b64encode(photo_bytes).decode("utf-8")
                media_type = photo.type or "image/jpeg"
                photo_data.append({
                    "filename": photo.name,
                    "b64": b64,
                    "media_type": media_type,
                })
                suffix = Path(photo.name).suffix or ".jpg"
                with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
                    tmp.write(photo_bytes)
                    temp_photo_paths.append(Path(tmp.name))

            # --- PDF テキスト抽出 ---
            pdf_text = ""
            if pdf_file:
                import pdfplumber
                with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
                    tmp.write(pdf_file.read())
                    tmp_pdf = tmp.name
                with pdfplumber.open(tmp_pdf) as pdf:
                    for page in pdf.pages:
                        t = page.extract_text()
                        if t:
                            pdf_text += t + "\n"

            # --- Claude プロンプト ---
            sections_str = "、".join(SECTIONS)
            prompt = f"""あなたは住宅アフター点検の専門家です。以下の情報を元に、点検報告書データをJSONで生成してください。

【基本情報】
物件名: {property_name}
点検実施日: {inspection_date}
報告者: {inspector}
担当者メモ: {notes}

【PDF内容】
{pdf_text if pdf_text else "（PDFなし）"}

【指示】
以下のJSON形式のみを返してください（マークダウン・コードブロック不要）：

{{
  "property_name": "{property_name}",
  "property_no": "{property_no}",
  "company_name": "{company_name}",
  "company_no": "{company_no}",
  "delivery_date": "{delivery_date}",
  "inspection_date": "{inspection_date}",
  "inspector": "{inspector}",
  "sections": [
    {{
      "section_no": 1,
      "section_name": "セクション名",
      "items": [
        {{
          "no": "1-1",
          "description": "点検項目の内容",
          "method": "AC",
          "judgment": "○",
          "notes": ""
        }}
      ]
    }}
  ],
  "special_notes": [],
  "photo_descriptions": [
    {{"filename": "ファイル名", "description": "写っている内容・問題点"}}
  ]
}}

sectionは必ず以下を全て含めてください（section_noは1から順）：{sections_str}
各セクションに最低2項目。問題箇所はjudgment「△」でnotesに詳細記載。
methodは「AC」基本。写真がある場合はphoto_descriptionsに説明を記載。"""

            content = []
            for pd_item in photo_data:
                content.append({
                    "type": "image",
                    "source": {
                        "type": "base64",
                        "media_type": pd_item["media_type"],
                        "data": pd_item["b64"],
                    },
                })
            content.append({"type": "text", "text": prompt})

            client = anthropic.Anthropic(api_key=api_key)
            response = client.messages.create(
                model="claude-opus-4-5",
                max_tokens=4096,
                messages=[{"role": "user", "content": content}],
            )

            raw = response.content[0].text.strip()
            for marker in ["```json", "```"]:
                if raw.startswith(marker):
                    raw = raw[len(marker):]
            if raw.endswith("```"):
                raw = raw[:-3]
            raw = raw.strip()

            report_data = json.loads(raw)

            # フォーム入力値で上書き
            overrides = {
                "property_name": property_name,
                "property_no": property_no,
                "company_name": COMPANY_NAME,
                "company_no": "",
                "delivery_date": delivery_date,
                "inspector": inspector,
                "inspection_date": str(inspection_date),
                "inspection_term": inspection_term,
            }
            for k, v in overrides.items():
                if v:
                    report_data[k] = v

            # --- Excel 生成 ---
            from excel_generator import generate_report_excel
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                output_path = Path(tmp.name)

            generate_report_excel(report_data, temp_photo_paths, output_path)
            excel_bytes = output_path.read_bytes()
            filename = f"{property_name or 'report'}_報告書.xlsx"

            st.success("✅ 報告書が生成されました！")
            st.download_button(
                label="📥 Excelをダウンロード",
                data=excel_bytes,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        except json.JSONDecodeError as e:
            st.error(f"AIの出力をJSONとして解析できませんでした: {e}")
        except Exception as e:
            import traceback
            st.error(f"エラー: {e}")
            with st.expander("詳細"):
                st.code(traceback.format_exc())
        finally:
            for p in temp_photo_paths:
                try: p.unlink()
                except Exception: pass
            try: output_path.unlink()
            except Exception: pass
