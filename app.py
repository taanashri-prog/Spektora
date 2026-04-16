from flask import Flask, render_template_string, request, send_file
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
import io

app = Flask(__name__)

# --- ELITE UI v0.4 ---
# (Using the same HTML_TEMPLATE from our previous chat)
HTML_TEMPLATE = '''
<!DOCTYPE html>
<html>
<head>
    <title>Spektora | Premium Studio</title>
    <style>
        :root { --primary: #7000ff; --secondary: #00d4ff; --bg: #0b0e14; --card: #151921; --text: #ffffff; }
        body { font-family: 'Segoe UI', sans-serif; background: var(--bg); color: var(--text); margin: 0; display: flex; }
        .sidebar { width: 320px; background: var(--card); height: 100vh; padding: 40px 25px; border-right: 1px solid rgba(255,255,255,0.1); }
        .content { flex-grow: 1; padding: 60px; display: flex; justify-content: center; align-items: center; }
        .glass-card { background: var(--card); padding: 40px; border-radius: 30px; border: 1px solid rgba(128,128,128,0.2); width: 100%; max-width: 600px; }
        textarea, input, select { background: rgba(0,0,0,0.2); border: 1px solid rgba(128,128,128,0.3); color: white; width: 100%; padding: 15px; margin: 10px 0; border-radius: 12px; }
        .btn-generate { background: linear-gradient(45deg, var(--primary), var(--secondary)); color: white; border: none; padding: 20px; border-radius: 15px; cursor: pointer; font-size: 18px; font-weight: bold; width: 100%; margin-top: 10px; }
    </style>
</head>
<body>
    <div class="sidebar">
        <h1>SPEKTORA</h1>
        <p>v0.4 Chart Engine</p>
    </div>
    <div class="content">
        <div class="glass-card">
            <form action="/generate" method="post">
                <input type="text" name="title" placeholder="Presentation Topic" required>
                <select name="format">
                    <option value="artsy">🎨 Artsy (Dark Mode)</option>
                    <option value="official">🏢 Official (Clean)</option>
                </select>
                <textarea name="raw_info" rows="8" placeholder="Tip: Use 'Sales: 10, 20, 30' for a chart!" required></textarea>
                <button type="submit" class="btn-generate">GENERATE</button>
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
    
    # 🎨 THEME SETTINGS
    if style == "artsy":
        bg_rgb, txt_rgb, accent_rgb = (15, 15, 20), (255, 255, 255), (112, 0, 255)
    else:
        bg_rgb, txt_rgb, accent_rgb = (255, 255, 255), (30, 30, 30), (0, 102, 204)

    paragraphs = [p.strip() for p in raw_info.split('\n') if p.strip()]

    # --- 1. TITLE SLIDE ---
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # Paint Background
    rect = slide.shapes.add_shape(1, 0, 0, prs.slide_width, prs.slide_height)
    rect.fill.solid()
    rect.fill.fore_color.rgb = RGBColor(*bg_rgb)
    rect.line.fill.background()

    # Add Title Text
    title_box = slide.shapes.add_textbox(Inches(1), Inches(3), Inches(8), Inches(2))
    p = title_box.text_frame.paragraphs[0]
    p.text = title_text.upper()
    p.font.size, p.font.bold = Pt(50), True
    p.font.color.rgb = RGBColor(*accent_rgb)
    p.alignment = PP_ALIGN.CENTER

    # --- 2. CONTENT SLIDES ---
    for para in paragraphs:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        
        # Paint Background again (crucial!)
        rect = slide.shapes.add_shape(1, 0, 0, prs.slide_width, prs.slide_height)
        rect.fill.solid()
        rect.fill.fore_color.rgb = RGBColor(*bg_rgb)
        rect.line.fill.background()

        # Check for Chart Logic: "Label: 10, 20, 30"
        if ":" in para and any(c.isdigit() for c in para):
            try:
                label, vals = para.split(":")
                nums = [float(v.strip()) for v in vals.split(",")]
                
                chart_data = CategoryChartData()
                chart_data.categories = [f"Item {i+1}" for i in range(len(nums))]
                chart_data.add_series(label, tuple(nums))

                chart = slide.shapes.add_chart(
                    XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1), Inches(1.5), Inches(8), Inches(4.5), chart_data
                ).chart
                # Style the chart title
                slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1)).text = label
            except:
                # Fallback to Text if Chart fails
                tb = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(4))
                tb.text_frame.text = para
        else:
            # REGULAR TEXT SLIDE
            tb = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(5))
            tf = tb.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = para
            p.font.size = Pt(28)
            p.font.color.rgb = RGBColor(*txt_rgb)

    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name="Spektora_v04.pptx")

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
