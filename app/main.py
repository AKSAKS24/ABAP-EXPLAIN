from fastapi import FastAPI, Form
from fastapi.responses import StreamingResponse
from app.generate import extract_abap_explanation
from app.docx_writer import create_docx
import io

app = FastAPI()

@app.post("/generate-ts/")
async def generate_ts(abap_code: str = Form(...)):
    ts_text = extract_abap_explanation(abap_code)

    # print(ts_text)
    docx_buffer = io.BytesIO()
    create_docx(ts_text, docx_buffer)
    docx_buffer.seek(0)
    return StreamingResponse(
        docx_buffer,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": "attachment; filename=technical_spec.docx"}
    )

# import sys
# from generate import extract_abap_explanation

# def main():
#     print("Paste your ABAP code (multiline supported). Press Ctrl+D (Linux/Mac) or Ctrl+Z (Windows) then Enter to finish:")
#     abap_code = sys.stdin.read()  # Reads until EOF
#     ts_text = extract_abap_explanation(abap_code)
#     print("\nGenerated Explanation:\n")
#     print(ts_text)

# if __name__ == "__main__":
#     main()