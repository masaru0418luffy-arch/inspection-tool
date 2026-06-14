import os
import base64
import json
import tempfile
from pathlib import Path
from flask import Flask, request, jsonify, send_from_directory, send_file
from dotenv import load_dotenv
import anthropic

load_dotenv()

app = Flask(__name__, template_folder="templates")
app.config["UPLOAD_FOLDER"] = Path(__file__).parent / "uploads"
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB

ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY")

SECTIONS = [
    "外壁",
    "その他外部",
    "玄関(内装)",
    "各室(内装)",
    "食堂・台所(設備)",
    "食堂・台所(内装)",
    "床下点検口",
    "洗面・脱衣所(設備)",
    "洗面・脱衣所(内装)",
    "浴室(設備)",
    "浴室(内装)",
    "トイレ(設備)",
    "トイレ(内装)",
    "廊下・階段(内装)",
    "建具",
]


@app.route("/")
def index():
    return send_from_directory("templates", "index.html")


@app.route("/generate", methods=["POST"])
def generate():
    try:
        property_name = request.form.get("property_name", "")
        property_no = request.form.get("property_no", "")
        company_name = request.form.get("company_name", "")
        company_no = request.form.get("company_no", "")
        delivery_date = request.form.get("delivery_date", "")
        inspector = request.form.get("inspector", "")
        inspection_date = request.form.get("inspection_date", "")
        notes = request.form.get("notes", "")

        photos = request.files.getlist("photos")
        pdf_file = request.files.get("pdf")

        # Save photos to temp files and encode to base64
        photo_data = []
        saved_photo_paths = []
        upload_dir = app.config["UPLOAD_FOLDER"]
        upload_dir.mkdir(exist_ok=True)

        for photo in photos:
            if photo.filename:
                photo_path = upload_dir / photo.filename
                photo.save(str(photo_path))
                saved_photo_paths.append(photo_path)

                raw = photo_path.read_bytes()
                b64 = base64.standard_b64encode(raw).decode("utf-8")
                media_type = photo.content_type or "image/jpeg"
                photo_data.append({"filename": photo.filename, "b64": b64, "media_type": media_type})

        # Extract PDF text
        pdf_text = ""
        if pdf_file and pdf_file.filename:
            import pdfplumber
            pdf_path = upload_dir / pdf_file.filename
            pdf_file.save(str(pdf_path))
            with pdfplumber.open(str(pdf_path)) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        pdf_text += text + "\n"

        # Build Claude API message
        client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

        sections_str = "、".join(SECTIONS)
        prompt = f"""あなたは住宅アフター点検の専門家です。以下の情報を元に、JIO形式の点検報告書データをJSONで生成してください。

【基本情報】
物件名: {property_name}
点検実施日: {inspection_date}
報告者: {inspector}
担当者メモ: {notes}

【PDF内容】
{pdf_text if pdf_text else "（PDFなし）"}

【指示】
以下のJSON形式で返してください。マークダウンのコードブロックは使わず、JSONのみを返してください。

{{
  "property_name": "物件名",
  "company_name": "事業者名（不明な場合は空文字）",
  "delivery_date": "引渡し年月（不明な場合は空文字）",
  "inspection_date": "実施日",
  "inspector": "報告者名",
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
          "notes": "備考"
        }}
      ]
    }}
  ],
  "special_notes": ["特記事項1", "特記事項2"],
  "photo_descriptions": [
    {{"filename": "ファイル名", "description": "写っている内容・問題点"}}
  ]
}}

sectionは必ず以下を全て含めてください（section_noは1から順に付ける）：
{sections_str}

各セクションには最低2項目の点検項目を含めてください。
問題が見つかった箇所はjudgmentを「△」にし、notesに具体的な内容を記載してください。
問題のない箇所はjudgmentを「○」にしてください。
methodは「AC」（目視確認）を基本としてください。

写真が添付されている場合はphoto_descriptionsに各写真の説明を記載してください。
"""

        # Build message content with images
        content = []
        for pd in photo_data:
            content.append({
                "type": "image",
                "source": {
                    "type": "base64",
                    "media_type": pd["media_type"],
                    "data": pd["b64"],
                },
            })

        content.append({"type": "text", "text": prompt})

        response = client.messages.create(
            model="claude-opus-4-5",
            max_tokens=4096,
            messages=[{"role": "user", "content": content}],
        )

        raw_text = response.content[0].text.strip()

        # Strip JSON code block markers if present
        if raw_text.startswith("```json"):
            raw_text = raw_text[7:]
        if raw_text.startswith("```"):
            raw_text = raw_text[3:]
        if raw_text.endswith("```"):
            raw_text = raw_text[:-3]
        raw_text = raw_text.strip()

        report_data = json.loads(raw_text)

        # Override with form values if provided
        if property_name:
            report_data["property_name"] = property_name
        if property_no:
            report_data["property_no"] = property_no
        if company_name:
            report_data["company_name"] = company_name
        if company_no:
            report_data["company_no"] = company_no
        if delivery_date:
            report_data["delivery_date"] = delivery_date
        if inspector:
            report_data["inspector"] = inspector
        if inspection_date:
            report_data["inspection_date"] = inspection_date

        # Generate Excel
        from excel_generator import generate_report_excel

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            output_path = Path(tmp.name)

        generate_report_excel(report_data, saved_photo_paths, output_path)

        filename = f"{property_name or 'report'}_報告書.xlsx"
        return send_file(
            str(output_path),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=filename,
        )

    except json.JSONDecodeError as e:
        return jsonify({"error": f"AIの出力をJSONとして解析できませんでした: {str(e)}"}), 500
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    app.run(debug=True, port=5000)
