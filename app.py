from flask import Flask, render_template_string, request, send_file
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import io

app = Flask(__name__)

# --- SAAS INTERFACE ---
HTML_TEMPLATE = '''
<!DOCTYPE html>
<html>
<head>
    <title>Spektora v0.2.1 | Ultra-Polished</title>
    <style>
        :root { --primary: #00d2ff; --secondary: #3a7bd5; --dark: #0f2027; }
        body { font-family: 'Inter', sans-serif; background: #eef2f3; margin: 0; display: flex; }
        .spark-sidebar { width: 320px; background: linear-gradient(to bottom, #2c3e50, #000000); color: white; height: 100vh; padding: 30px; box-sizing: border-box; }
        .spark-bubble { background: rgba(255,255,255,0.1); padding: 15px; border-radius: 15px; font-size: 14px; line-height: 1.6; border-left: 4px solid var(--primary); }
        .main { flex-grow: 1; padding: 60px; display: flex; justify-content: center; }
        .container { background: white; padding: 40px; border-radius: 25px; box-shadow: 0 20px 50px rgba(0,0,0,0.1); width: 100%; max-width: 650px; }
        h1 { background: -webkit-linear-gradient(var(--primary), var(--secondary)); -webkit-background-clip: text; -webkit-text-fill-color: transparent; font-size: 38px; margin-bottom: 10px; }
        select, input, textarea { width: 100%; padding: 15px; margin: 12px 0; border-radius: 12px; border: 1px solid #ddd; font-size: 16px; }
        button { background: linear-gradient(45deg, var(--primary), var(--secondary)); color: white; border: none; padding: 18px; border-radius: 15px; cursor: pointer; font-size: 20px; font-weight: bold; width: 100%; transition: 0.3s; }
        button:hover { transform: translateY(-3px); box-shadow: 0 10px 20px rgba(58, 123, 213, 0.3); }
    </style>
</head>
<body>
    <div class="spark-sidebar">
        <h2 style="letter-spacing: 2px;">SPEKTORA</h2>
        <p style="color: #888;">AI Presentation Studio</p>
        <div class="spark-bubble">
            ✨ <b>Spark Advice:</b><br><br>
            "I've updated the formats! Try <b>Artsy</b> for a neon-on-dark look, or <b>Minimalist</b> to let your ideas breathe."
        </div>
    </div>
    <div class="main">
        <div class="container">
            <h1>Create v0.2.1 🚀</h1>
            <form action="/generate" method="post">
                <input type="text" name="title" placeholder="Project Title (e.g. Q4 Strategy)" required>
                <select name="format">
                    <option value="classic">🏛️ Classic (Corporate Blue)</option>
                    <option value="minimalist">☁️ Minimalist (Clean & Airy)</option>
                    <option value="artsy">🎨 Artsy (Neon & Dark)</option>
                    <option value="academic">🎓 Academic (Elegant Serif)</option>
                </select>
                <textarea name="raw_info" rows="8" placeholder="Enter your notes... Each new paragraph becomes a new slide!" required></textarea>
                <button type="submit">Build Presentation ✨</button>
            </form>
        </div>
    </div>
</body>
</html>
'''

@app.route('/')
def home():
    return render_template_string(HTML_TEMPLATE)

@app.route('/generate', methods=['POST'])
def generate():
    title_text = request.form.get('title')
    raw_info = request.form.get('raw_info')
    style = request.form.get('format')

    prs = Presentation()
    
    # Define Theme Styles
    themes = {
        "classic": {"bg": (255, 255, 255), "text": (0, 32, 96), "font": "Arial", "align": PP_ALIGN.LEFT},
        "minimalist": {"bg": (250, 250, 250), "text": (50, 50, 50), "font": "Helvetica", "align": PP_ALIGN.CENTER},
        "artsy": {"bg": (15, 15, 15), "text": (0, 210, 255), "font": "Impact", "align": PP_ALIGN.LEFT},
        "academic": {"bg": (255, 253, 245), "text": (44, 62, 80), "font": "Georgia", "align": PP_ALIGN.JUSTIFY}
    }
    
    theme = themes.get(style, themes["classic"])
    paragraphs = [p.strip() for p in raw_info.split('\n') if p.strip()]

    # 1. CREATE TITLE SLIDE
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rect = slide.shapes.add_shape(1, 0, 0, prs.slide_width, prs.slide_height)
    rect.fill.solid()
    rect.fill.fore_color.rgb = RGBColor(*theme["bg"])
    rect.line.fill.background()

    t_box = slide.shapes.add_textbox(Inches(1), Inches(3), Inches(8), Inches(2))
    p = t_box.text_frame.paragraphs[0]
    p.text = title_text.upper()
    p.font.size, p.font.bold, p.font.name = Pt(54), True, theme["font"]
    p.font.color.rgb = RGBColor(*theme["text"])
    p.alignment = PP_ALIGN.CENTER

    # 2. CREATE CONTENT SLIDES (Multi-slide logic)
    for para in paragraphs:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        
        # Background
        rect = slide.shapes.add_shape(1, 0, 0, prs.slide_width, prs.slide_height)
        rect.fill.solid()
        rect.fill.fore_color.rgb = RGBColor(*theme["bg"])
        rect.line.fill.background()

        # Header Line (Small)
        h_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(5), Inches(0.5))
        hp = h_box.text_frame.paragraphs[0]
        hp.text = title_text
        hp.font.size, hp.font.name = Pt(12), theme["font"]
        hp.font.color.rgb = RGBColor(*theme["text"])

        # Main Content Body
        c_box = slide.shapes.add_textbox(Inches(1), Inches(1.5), Inches(8), Inches(5))
        cf = c_box.text_frame
        cf.word_wrap = True
        p = cf.paragraphs[0]
        p.text = para
        p.font.size = Pt(32) if len(para) < 100 else Pt(24)
        p.font.name = theme["font"]
        p.font.color.rgb = RGBColor(*theme["text"])
        p.alignment = theme["align"]

    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name=f"Spektora_{style}.pptx")

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)

