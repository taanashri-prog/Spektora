from flask import Flask, render_template_string, request, send_file
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
import io

app = Flask(__name__)

# --- UI & UX DESIGN ---

HTML_TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
    <title>Spektora | Raw Ideas. Polished Impact.</title>
    <style>
        :root { 
            --bg: #0D0D0D; --accent: #D4AF37; --card: #1A1A1A; --text: #FFFFFF;
            --font-main: 'Arial', sans-serif;
        }
        .light-theme { --bg: #F5F5F5; --card: #FFFFFF; --text: #1A1A1A; }
        
        body { font-family: var(--font-main); background: var(--bg); color: var(--text); margin: 0; overflow: hidden; transition: 0.5s; }

        /* SPLASH SCREEN & SHRINK LOGO */
        #splash {
            position: fixed; width: 100%; height: 100vh; background: #000;
            display: flex; flex-direction: column; justify-content: center; align-items: center;
            z-index: 9999; transition: 1s ease-in-out;
        }
        .logo-main { width: 150px; transition: 1.2s cubic-bezier(0.7, 0, 0.3, 1); filter: drop-shadow(0 0 15px var(--accent)); }
        
        /* MAIN DASHBOARD */
        #main-app { display: flex; height: 100vh; opacity: 0; transition: 1s; }
        .sidebar { width: 300px; background: var(--card); padding: 30px; border-right: 1px solid rgba(212,175,55,0.2); position: relative; }
        
        .spark-chat { background: rgba(212,175,55,0.1); padding: 20px; border-radius: 20px; border: 1px solid var(--accent); margin-top: 40px; }
        .spark-header { font-weight: bold; color: var(--accent); display: flex; align-items: center; gap: 10px; }
        
        .potted-plant { position: absolute; bottom: 20px; left: 20px; width: 120px; opacity: 0.8; }
        
        .engine-room { flex-grow: 1; display: flex; flex-direction: column; justify-content: center; align-items: center; padding: 40px; }
        .glass-box { background: var(--card); padding: 50px; border-radius: 30px; border: 1px solid rgba(255,255,255,0.1); width: 100%; max-width: 600px; }
        
        input, textarea, select { width: 100%; padding: 15px; margin: 10px 0; border-radius: 12px; background: rgba(0,0,0,0.3); color: white; border: 1px solid #333; }
        .btn-generate { background: var(--accent); color: black; border: none; padding: 20px; border-radius: 15px; font-weight: bold; width: 100%; cursor: pointer; font-size: 18px; }
        
        .theme-toggle { position: absolute; top: 20px; right: 20px; cursor: pointer; background: none; border: 1px solid var(--accent); color: var(--accent); padding: 8px 15px; border-radius: 20px; }
        .logo-small { width: 60px !important; position: absolute; top: 30px; left: 30px; }
    </style>
</head>
<body id="body">

    <div id="splash">
        <div id="logo-container" style="text-align:center;">
             <h1 id="logo" class="logo-main" style="color:var(--accent); font-size: 60px;">S</h1>
        </div>
        <h2 id="slogan" style="color: var(--accent); letter-spacing: 5px; margin-top: 20px;">RAW IDEAS. POLISHED IMPACT.</h2>
        <div style="width: 200px; height: 2px; background: #333; margin-top: 20px;"><div id="bar" style="width:0%; height:100%; background: var(--accent); transition: 2s;"></div></div>
    </div>

    <div id="main-app">
        <div class="sidebar">
            <div class="spark-chat">
                <div class="spark-header">✨ SPARK</div>
                <p style="font-size: 14px;">"I'm here to help! Type your notes and I'll handle the magic."</p>
            </div>
            <div class="potted-plant" style="font-size: 80px;">🌿</div>
        </div>

        <div class="engine-room">
            <button class="theme-toggle" onclick="toggleTheme()">🌓 Light/Dark</button>
            <div class="glass-box">
                <form action="/generate" method="post">
                    <input type="text" name="title" placeholder="Topic Title" required>
                    <textarea name="raw_info" rows="6" placeholder="Paste your data or ideas..."></textarea>
                    <select name="format">
                        <option value="official">🏢 Official (Green Barcode)</option>
                        <option value="classic">📜 Classic (Old English)</option>
                        <option value="artsy">🎨 Artsy (Blue & Teal)</option>
                        <option value="minimalist">☁️ Minimalist (Blue Strips)</option>
                    </select>
                    <button type="submit" class="btn-generate">GENERATE</button>
                </form>
            </div>
        </div>
    </div>

    <script>
        window.onload = () => {
            document.getElementById('bar').style.width = '100%';
            setTimeout(() => {
                const logo = document.getElementById('logo');
                const splash = document.getElementById('splash');
                logo.classList.add('logo-small');
                splash.style.background = 'transparent';
                document.getElementById('slogan').style.opacity = '0';
                setTimeout(() => {
                    splash.style.display = 'none';
                    document.getElementById('main-app').style.opacity = '1';
                }, 1000);
            }, 2500);
        };

        function toggleTheme() {
            document.getElementById('body').classList.toggle('light-theme');
        }
    </script>
</body>
</html>
"""
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
