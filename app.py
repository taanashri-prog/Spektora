from flask import Flask, render_template_string, request, send_file
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import io

app = Flask(__name__)

# Updated Website Look (More modern colors)
HTML_TEMPLATE = '''
<!DOCTYPE html>
<html>
<head>
    <title>Spektora Design Suite</title>
    <style>
        body { font-family: 'Segoe UI', sans-serif; text-align: center; background: #eef2f3; padding: 50px; }
        .container { background: white; padding: 40px; border-radius: 20px; display: inline-block; box-shadow: 0 10px 25px rgba(0,0,0,0.1); border-top: 8px solid #6a11cb; }
        h1 { color: #2575fc; margin-bottom: 5px; }
        textarea { width: 450px; height: 180px; padding: 15px; border-radius: 10px; border: 1px solid #ddd; margin: 10px 0; font-size: 14px; }
        input { width: 450px; padding: 12px; margin-bottom: 10px; border-radius: 10px; border: 1px solid #ddd; }
        button { background: linear-gradient(to right, #6a11cb, #2575fc); color: white; border: none; padding: 12px 30px; border-radius: 25px; cursor: pointer; font-size: 18px; font-weight: bold; transition: 0.3s; }
        button:hover { transform: scale(1.05); box-shadow: 0 5px 15px rgba(37, 117, 252, 0.4); }
    </style>
</head>
<body>
    <div class="container">
        <h1>🚀 Spektora</h1>
        <p style="color: #666;">Turning raw notes into beautiful slides</p>
        <form action="/generate" method="post">
            <input type="text" name="title" placeholder="Enter Slide Title..." required><br>
            <textarea name="raw_info" placeholder="Paste your raw info here... (e.g. - Feature A - Feature B)" required></textarea><br>
            <button type="submit">Design My Slide ✨</button>
        </form>
    </div>
</body>
</html>
'''

@app.route('/')
def home():
    return render_template_string(HTML_TEMPLATE)

@app.route('/generate', methods=['POST'])
def generate():
    user_title = request.form.get('title', 'Presentation')
    raw_text = request.form.get('raw_info', '')

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6]) # Blank layout

    # --- DECORATION: ADD A BACKGROUND RECTANGLE ---
    # This creates a soft light-blue background for the whole slide
    background = slide.shapes.add_shape(
        1, 0, 0, prs.slide_width, prs.slide_height
    )
    background.fill.solid()
    background.fill.fore_color.rgb = RGBColor(240, 245, 255) # Very light blue
    background.line.fill.background() # Remove border

    # --- DECORATION: ADD A STYLISH SIDEBAR ---
    sidebar = slide.shapes.add_shape(
        1, 0, 0, Inches(0.15), prs.slide_height
    )
    sidebar.fill.solid()
    sidebar.fill.fore_color.rgb = RGBColor(106, 17, 203) # Purple Accent
    sidebar.line.fill.background()

    # --- ADD TITLE ---
    title_box = slide.shapes.add_textbox(Inches(0.6), Inches(0.5), Inches(8), Inches(1))
    tf = title_box.text_frame
    tf.text = user_title
    p = tf.paragraphs[0]
    p.font.bold = True
    p.font.size = Pt(40)
    p.font.color.rgb = RGBColor(40, 40, 40) # Dark Grey

    # --- ADD CONTENT BOX ---
    content_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(8.5), Inches(5))
    cf = content_box.text_frame
    cf.word_wrap = True

    lines = raw_text.split('\n')
    for line in lines:
        if line.strip():
            p = cf.add_paragraph()
            p.text = line.strip()
            p.font.size = Pt(22)
            p.space_after = Pt(12)
            p.font.color.rgb = RGBColor(60, 60, 60)
            p.level = 0 # This adds a bullet point automatically

    # Save and Send
    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name="Spektora_Designed.pptx")

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)

