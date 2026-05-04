import mammoth

PREVIEW_CSS = """
<style>
    @page { size: A4; margin: 2.54cm; }
    body {
        font-family: "宋体", "SimSun", "Times New Roman", serif;
        font-size: 12pt;
        line-height: 1.5;
        color: #333;
        max-width: 210mm;
        margin: 0 auto;
        padding: 20px 40px;
        background: white;
    }
    h1 {
        font-family: "黑体", "SimHei", sans-serif;
        font-size: 16pt;
        font-weight: bold;
        text-align: center;
        margin: 24px 0 12px;
    }
    h2 {
        font-family: "黑体", "SimHei", sans-serif;
        font-size: 14pt;
        font-weight: bold;
        margin: 18px 0 8px;
    }
    h3 {
        font-family: "黑体", "SimHei", sans-serif;
        font-size: 12pt;
        font-weight: bold;
        margin: 12px 0 6px;
    }
    p {
        text-indent: 2em;
        margin: 6px 0;
    }
    table {
        border-collapse: collapse;
        width: 100%;
        margin: 12px 0;
        font-size: 10.5pt;
    }
    th, td {
        border: 1px solid #333;
        padding: 6px 10px;
        text-align: center;
    }
    th {
        font-weight: bold;
        background: #f5f5f5;
    }
    img {
        max-width: 100%;
        height: auto;
        display: block;
        margin: 12px auto;
    }
    a { color: #1a73e8; text-decoration: none; }
    a:hover { text-decoration: underline; }
    .page-break-hint {
        border-top: 1px dashed #ccc;
        margin: 40px 0 20px;
        text-align: center;
        color: #999;
        font-size: 10pt;
    }
</style>
"""

PREVIEW_WARNING = """
<div style="background:#fff3cd;border:1px solid #ffc107;border-radius:6px;padding:10px 16px;margin-bottom:16px;font-size:13px;color:#856404;">
    ⚠️ 预览为 HTML 渲染效果，可能与最终 Word 文档存在细微差异（字体、分页、页眉页脚等）。
    如需精确排版效果，请下载 .docx 文件后在 Word 中查看。
</div>
"""


def convert_docx_to_html(docx_path: str) -> str:
    with open(docx_path, "rb") as f:
        result = mammoth.convert_to_html(f)

    body_html = result.value
    messages = result.messages

    msg_html = ""
    if messages:
        msg_items = "".join(f"<li>{m.message}</li>" for m in messages)
        msg_html = f'<details style="margin-bottom:12px;font-size:12px;color:#666;"><summary>转换信息 ({len(messages)})</summary><ul>{msg_items}</ul></details>'

    full_html = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>文档预览</title>
{PREVIEW_CSS}
</head>
<body>
{PREVIEW_WARNING}
{msg_html}
{body_html}
</body>
</html>"""

    return full_html
