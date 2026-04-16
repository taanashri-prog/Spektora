from flask import Flask, render_template_string, request, send_file
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import io

app = Flask(__name__)

# --- NEW SAAS DASHBOARD v0.3 ---
HTML_TEMPLATE = '''
<!DOCTYPE html>
<html>
<head>
    <title>Spektora v0.3 | Intelligence Update</title>
    <style>
        :root { --primary: #7000ff; --secondary: #00d4ff; --bg: #0b0e14; }
        body { font-family: 'Segoe UI', sans-serif; background: var(--bg); color: white; margin: 0; display: flex; }
        
        .sidebar { width: 350px; background: #151921; height: 100vh; padding: 40px 25px; border-right: 1px solid #2d333b; }
        .spark-ai { background: linear-gradient(145deg, #1e2530, #151921); padding: 20px; border-radius: 20px; border: 1px solid #3a424d; margin-top: 30px; }
        .spark-status { display: inline-block; width: 10px; height: 10px; background: #00ff88; border-radius: 50%; margin-right: 10px; }
        
        .content { flex-grow: 1; padding: 60px; display: flex; flex-direction: column; align-items: center; }
        .glass-card { background: rgba(255, 255, 255, 0.05); backdrop-filter: blur(10px); padding: 40px; border-radius: 30px; border: 1px solid rgba(255,255,255,0.1); width: 100%; max-width: 700px; }
        
        h1 { font-size: 42px; background: linear-gradient(to right, #fff, #888); -webkit-background-clip: text; -webkit-text-fill-color: transparent; margin-bottom: 5px; }
        textarea, input, select { background: #1c2128; border: 1px solid #3a424d; color: white; width: 100%; padding: 15px; margin: 12px 0; border-radius: 12px; font-size: 16px; }
        
        .btn-generate { background: linear-gradient(45deg, var(--primary), var(--secondary)); color: white; border: none; padding: 20px; border-radius: 15px; cursor: pointer; font-size: 18px; font-weight: bold; width: 100%; margin-top: 20px; transition: 0.3s; }
        .btn-generate:hover { transform: scale(1.02); box-shadow: 0 0 30px rgba(112, 0, 255, 0.4); }
    </style>
</head>
<body>
    <div class="sidebar">
        <h2 style="letter-spacing: 3px;">SPEKTORA <span style="font-size: 12px; vertical-align: top;">v0.3</span></h2>
        <div class="spark-ai">
            <p><span class="spark-status"></span><b>Spark AI</b> is Active</p>
            <p style="font-size: 14px; color: #a1a1a1;" id="spark-advice">
                "Ready to transform your ideas? Paste your notes, and I'll handle the visual hierarchy."
            </p>
        </div>
        <p style="margin-top: 50px; font-size: 12px; color: #555;">&copy; 2026 Spektora SaaS Studio</p>
    </div>

    <div class="content">
        <div class="glass-card">
            <h1>Engine Room</h1>
            <p style="color: #888; margin-bottom: 30px;">Input your raw data to begin the transformation.</p>
            
            <form action="/generate" method="post">
                <input type="text" name="title" placeholder="Presentation Topic" required>
                <select name="format">
                    <option value="official">🏢 Official (Corporate & Bold)</option>
                    <option value="minimalist">☁️ Minimalist (Clean & Airy)</option>
                    <option value="artsy">🎨 Artsy (Neon Dark Mode)</option>
                    <option value="academic">🎓 Academic (Traditional)</option>
                </select>
                <textarea name="raw_info" rows="8" placeholder="Paste information here... Spark will organize each paragraph into a slide." required></textarea>
                <button type="submit" class="btn-generate">GENERATE PRESENTATION</button>
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
    
    # Theme settings
    themes = {
        "official": {"bg": (255, 255, 255), "accent": (0, 102, 204), "text": (20, 20, 20), "font": "Calibri"},
        "minimalist": {"bg": (255, 255, 255), "accent": (200, 200, 200), "text": (60, 60, 60), "font": "Arial Light"},
        "artsy": {"bg": (10, 10, 15), "accent": (180, 0, 255), "text": (255, 255, 255), "font": "Impact"},
        "academic": {"bg": (255, 254, 250), "accent": (80, 0, 0), "text": (30, 30, 30), "font": "Georgia"}
    }
    
    t = themes.get(style, themes["official"])
    paragraphs = [p.strip() for p in raw_info.split('\n') if p.strip()]

    # Function to apply background
    def apply_bg(slide_obj):
        rect = slide_obj.shapes.add_shape(1, 0, 0, prs.slide_width, prs.slide_height)
        rect.fill.solid()
        rect.fill.fore_color.rgb = RGBColor(*t["bg"])
        rect.line.fill.background()
        if style == "official": # Add a professional side-bar
            sidebar = slide_obj.shapes.add_shape(1, 0, 0, Inches(0.1), prs.slide_height)
            sidebar.fill.solid()
            sidebar.fill.fore_color.rgb = RGBColor(*t["accent"])
            sidebar.line.fill.background()

    # 1. TITLE SLIDE
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    apply_bg(slide)
    
    t_box = slide.shapes.add_textbox(Inches(1), Inches(3.2), Inches(8), Inches(1.5))
    p = t_box.text_frame.paragraphs[0]
    p.text = title_text.upper()
    p.font.size, p.font.bold, p.font.name = Pt(50), True, t["font"]
    p.font.color.rgb = RGBColor(*t["accent"])
    p.alignment = PP_ALIGN.CENTER

    # 2. CONTENT SLIDES
    for para in paragraphs:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        apply_bg(slide)
        
        # Add a subtle "Card" shape for the text (SaaS look)
        if style != "minimalist":
            shape = slide.shapes.add_shape(1, Inches(0.8), Inches(1.2), Inches(8.4), Inches(5))
            shape.fill.solid()
            shape.fill.fore_color.rgb = RGBColor(*t["bg"]) # Slight contrast would go here
            shape.line.color.rgb = RGBColor(*t["accent"])
            shape.line.width = Pt(1.5)

        # Content Text
        c_box = slide.shapes.add_textbox(Inches(1.2), Inches(1.5), Inches(7.6), Inches(4.5))
        cf = c_box.text_frame
        cf.word_wrap = True
        p = cf.paragraphs[0]
        p.text = para
        p.font.size, p.font.name = Pt(28), t["font"]
        p.font.color.rgb = RGBColor(*t["text"])
        p.alignment = PP_ALIGN.LEFT if style != "minimalist" else PP_ALIGN.CENTER

    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name=f"Spektora_v3_{style}.pptx")

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
