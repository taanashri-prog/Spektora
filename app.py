
from flask import Flask, render_template_string, request, send_file
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import io

app = Flask(__name__)

# --- ELITE UI v0.4 (With Loading & Theme Support) ---
HTML_TEMPLATE = '''
<!DOCTYPE html>
<html>
<head>
    <title>Spektora | Premium Presentation Studio</title>
    <style>
        :root { --primary: #7000ff; --secondary: #00d4ff; --bg: #0b0e14; --card: #151921; --text: #ffffff; }
        .light-theme { --bg: #f5f7fa; --card: #ffffff; --text: #1a1a1a; }
        
        body { font-family: 'Segoe UI', sans-serif; background: var(--bg); color: var(--text); margin: 0; transition: 0.4s; }
        
        /* Splash Screen Logic */
        #loader { position: fixed; width: 100%; height: 100vh; background: var(--bg); display: none; flex-direction: column; justify-content: center; align-items: center; z-index: 1000; }
        .logo-glow { width: 80px; height: 80px; border-radius: 50%; background: var(--primary); box-shadow: 0 0 30px var(--primary); animation: pulse 1.5s infinite; }
        @keyframes pulse { 0% { transform: scale(0.9); opacity: 0.7; } 50% { transform: scale(1.1); opacity: 1; } 100% { transform: scale(0.9); opacity: 0.7; } }

        .sidebar { width: 320px; background: var(--card); height: 100vh; padding: 40px 25px; border-right: 1px solid rgba(255,255,255,0.1); }
        .content { flex-grow: 1; padding: 60px; display: flex; justify-content: center; align-items: center; position: relative; }
        
        .theme-toggle { position: absolute; top: 20px; right: 20px; cursor: pointer; padding: 10px 20px; border-radius: 20px; background: var(--primary); color: white; border: none; font-weight: bold; }
        
        .glass-card { background: var(--card); padding: 40px; border-radius: 30px; border: 1px solid rgba(128,128,128,0.2); width: 100%; max-width: 600px; box-shadow: 0 20px 40px rgba(0,0,0,0.3); }
        h1 { font-size: 36px; margin-bottom: 5px; background: linear-gradient(to right, var(--primary), var(--secondary)); -webkit-background-clip: text; -webkit-text-fill-color: transparent; }
        
        textarea, input, select { background: rgba(0,0,0,0.1); border: 1px solid rgba(128,128,128,0.3); color: var(--text); width: 100%; padding: 15px; margin: 10px 0; border-radius: 12px; }
        .btn-generate { background: linear-gradient(45deg, var(--primary), var(--secondary)); color: white; border: none; padding: 20px; border-radius: 15px; cursor: pointer; font-size: 18px; font-weight: bold; width: 100%; margin-top: 10px; }
    </style>
</head>
<body id="body-tag">

    <div id="loader">
        <div class="logo-glow"></div>
        <h2 style="margin-top: 20px;">Spark is building your deck...</h2>
    </div>

    <div class="sidebar">
        <h1>SPEKTORA</h1>
        <button onclick="toggleTheme()" class="theme-toggle">🌓 Toggle Theme</button>
        <div style="margin-top: 40px; padding: 15px; background: rgba(112,0,255,0.1); border-radius: 15px; font-size: 14px;">
            <b>Spark Suggestion:</b><br>
            <span id="ai-hint">"Paste some data points and I'll suggest the best chart!"</span>
        </div>
    </div>

    <div class="content">
        <div class="glass-card">
            <h1>Engine v0.4</h1>
            <form action="/generate" method="post" onsubmit="showLoading()">
                <input type="text" name="title" placeholder="Presentation Topic" required>
                <select name="format">
                    <option value="official">🏢 Official</option>
                    <option value="minimalist">☁️ Minimalist</option>
                    <option value="artsy">🎨 Artsy</option>
                </select>
                <textarea name="raw_info" rows="8" placeholder="Enter your data or notes..." onkeyup="checkData(this)"></textarea>
                <button type="submit" class="btn-generate">GENERATE</button>
            </form>
        </div>
    </div>

    <script>
        function toggleTheme() {
            document.getElementById('body-tag').classList.toggle('light-theme');
        }
        function showLoading() {
            document.getElementById('loader').style.display = 'flex';
        }
        function checkData(el) {
            if(el.value.includes('%') || el.value.includes('total')) {
                document.getElementById('ai-hint').innerHTML = "📊 I see statistics! Should I generate a Bar Chart slide for this?";
            }
        }
    </script>
</body>
</html>
'''

# [Generate Route remains the same as v0.3 for now]
@app.route('/')
def home():
    return render_template_string(HTML_TEMPLATE)

@app.route('/generate', methods=['POST'])
def generate():
    title_text = request.form.get('title')
    raw_info = request.form.get('raw_info')
    style = request.form.get('format')
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    # [Internal logic for PPTX generation here]
    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name=f"Spektora_v04.pptx")

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
