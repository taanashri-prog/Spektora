import os
import io
from flask import Flask, render_template_string, request, send_file
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE

app = Flask(__name__)

# --- UI & UX DESIGN (HTML/CSS/JS) ---
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Spektora | Your Voice. Our Vision.</title>
    <style>
        :root { 
            --bg: #0D0D0D; 
            --accent: #D4AF37; 
            --card: #1A1A1A; 
            --text: #FFFFFF;
        }
        .light-theme { --bg: #F5F5F5; --card: #FFFFFF; --text: #1A1A1A; }
        
        body { font-family: 'Arial', sans-serif; background: var(--bg); color: var(--text); margin: 0; overflow: hidden; transition: 0.5s; }

        /* SPLASH SCREEN & LOGO ANIMATION */
        #splash {
            position: fixed; width: 100%; height: 100vh; background: #000;
            display: flex; flex-direction: column; justify-content: center; align-items: center;
            z-index: 9999; transition: 1s cubic-bezier(0.7, 0, 0.3, 1);
        }
        
        #logo-container { position: relative; }
        .logo-main { 
            font-size: 80px; font-weight: bold; color: var(--accent); 
            transition: 1.2s cubic-bezier(0.7, 0, 0.3, 1); 
            text-shadow: 0 0 20px rgba(212, 175, 55, 0.6);
        }
        
        /* Glitter particles around the logo */
        .sparkle {
            position: absolute; width: 4px; height: 4px; background: var(--accent);
            border-radius: 50%; pointer-events: none; opacity: 0;
            animation: glitter-move 2s infinite ease-in-out;
        }
        @keyframes glitter-move {
            0% { transform: translate(0,0); opacity: 0; }
            50% { opacity: 1; }
            100% { transform: translate(calc(Math.random()*100px - 50px), -100px); opacity: 0; }
        }

        /* MAIN INTERFACE */
        #main-app { display: flex; height: 100vh; opacity: 0; transition: 1s; }
        
        .sidebar { width: 320px; background: var(--card); padding: 40px; border-right: 1px solid rgba(212,175,55,0.1); position: relative; }
        .spark-box { 
            background: rgba(212,175,55,0.05); border: 1px solid var(--accent); 
            padding: 20px; border-radius: 20px; margin-top: 60px; position: relative;
        }
        .spark-title { font-weight: bold; color: var(--accent); letter-spacing: 1px; margin-bottom: 5px; }
        
        .potted-plant { position: absolute; bottom: 30px; left: 30px; font-size: 80px; opacity: 0.9; }

        .engine-room { flex-grow: 1; display: flex; flex-direction: column; justify-content: center; align-items: center; padding: 40px; }
        .glass-card { 
            background: var(--card); padding: 50px; border-radius: 30px; 
            box-shadow: 0 25px 50px rgba(0,0,0,0.5); width: 100%; max-width: 600px; 
            border: 1px solid rgba(255,255,255,0.05);
        }
        
        input, textarea, select { 
            width: 100%; padding: 18px; margin: 12px 0; border-radius: 12px; 
            background: rgba(0,0,0,0.2); color: white; border: 1px solid #333; box-sizing: border-box;
        }
        
        .btn-gen { 
            background: var(--accent); color: black; border: none; padding: 22px; 
            border-radius: 15px; font-weight: 800; width: 100%; cursor: pointer; 
            font-size: 18px; letter-spacing: 2px; transition: 0.
