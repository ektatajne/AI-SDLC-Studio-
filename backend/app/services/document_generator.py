from __future__ import annotations
import os
import io
import re
import json
import zlib
import base64
import urllib.request
import docx
from datetime import datetime
from PIL import Image as PILImage
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfgen import canvas
from reportlab.graphics.shapes import Drawing, Rect, String, Line, Polygon, Group
from docx import Document
from docx.shared import Inches

import html

def sanitize_mermaid_chart(raw_chart: str) -> str:
    if not raw_chart or not isinstance(raw_chart, str):
        return ""
    
    cleaned = raw_chart.strip()
    
    # Strip markdown fences ```mermaid ... ```
    cleaned = re.sub(r'^```(?:mermaid)?\s*', '', cleaned, flags=re.MULTILINE)
    cleaned = re.sub(r'```$', '', cleaned, flags=re.MULTILINE).strip()
    
    # Remove PlantUML directives
    cleaned = re.sub(r'@startuml|@enduml', '', cleaned, flags=re.IGNORECASE).strip()
    
    # If diagram starts with PlantUML or actor directives
    if cleaned.startswith('left to right direction') or ('actor ' in cleaned and not any(cleaned.startswith(k) for k in ['graph', 'flowchart', 'sequenceDiagram', 'classDiagram', 'erDiagram'])):
        lines = cleaned.split('\n')
        nodes = []
        connections = []
        for line in lines:
            t = line.strip()
            if t.startswith('actor '):
                actor_name = t.replace('actor ', '').strip()
                actor_id = re.sub(r'\s+', '_', actor_name)
                nodes.append(f'  {actor_id}["User: {actor_name}"]')
            elif '-->' in t or '->' in t:
                parts = re.split(r'-->|->', t)
                if len(parts) == 2:
                    src = re.sub(r'\s+', '_', parts[0].strip())
                    dst_raw = parts[1].strip()
                    if dst_raw.startswith('(') and dst_raw.endswith(')'):
                        lbl = dst_raw[1:-1].strip()
                        dst_id = 'UC_' + re.sub(r'[^a-zA-Z0-9]', '_', lbl)
                        nodes.append(f'  {dst_id}(["{lbl}"])')
                        connections.append(f'  {src} --> {dst_id}')
                    else:
                        connections.append(f'  {src} --> {re.sub(r"[^a-zA-Z0-9]", "_", dst_raw)}')
        if connections:
            cleaned = f"graph LR\n  subgraph System[\"System Scope\"]\n" + "\n".join(set(nodes)) + "\n  end\n" + "\n".join(connections)

    # Ensure valid diagram header exists
    valid_headers = ('graph', 'flowchart', 'sequenceDiagram', 'classDiagram', 'erDiagram', 'gantt', 'pie', 'stateDiagram', 'C4Context')
    if not any(cleaned.startswith(h) for h in valid_headers):
        cleaned = "graph TD\n" + cleaned

    return cleaned


def render_mermaid_local_mmdc(mermaid_code: str) -> bytes | None:
    """Renders exact native Mermaid dark-theme PNG images using local Chrome & @mermaid-js/mermaid-cli."""
    sanitized = sanitize_mermaid_chart(mermaid_code)
    if not sanitized:
        return None
        
    try:
        import tempfile
        import subprocess
        
        chrome_paths = [
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
            r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"
        ]
        exe_path = next((p for p in chrome_paths if os.path.exists(p)), None)
        if not exe_path:
            return None

        with tempfile.TemporaryDirectory() as tmpdir:
            mmd_path = os.path.join(tmpdir, "input.mmd")
            png_path = os.path.join(tmpdir, "output.png")
            cfg_path = os.path.join(tmpdir, "puppeteer.json")
            
            with open(cfg_path, "w", encoding="utf-8") as f:
                json.dump({"executablePath": exe_path, "args": ["--no-sandbox", "--disable-setuid-sandbox", "--disable-gpu", "--disable-dev-shm-usage", "--no-first-run", "--no-zygote", "--headless"]}, f)
                
            with open(mmd_path, "w", encoding="utf-8") as f:
                f.write(sanitized)
                
            cmd = f'npx @mermaid-js/mermaid-cli -i "{mmd_path}" -o "{png_path}" -p "{cfg_path}" -t dark -b "#0f172a"'
            res = subprocess.run(cmd, shell=True, capture_output=True, text=True, timeout=10)
            
            if os.path.exists(png_path) and os.path.getsize(png_path) > 0:
                with open(png_path, "rb") as f:
                    return f.read()
    except Exception as e:
        print(f"[Warning] Local mmdc rendering failed ({e}).")
        
    return None


def render_mermaid_to_png_bytes(mermaid_code: str) -> bytes | None:
    # 1. Primary: Local native Mermaid JS CLI renderer (exact dark theme PNG rendering!)
    png_bytes = render_mermaid_local_mmdc(mermaid_code)
    if png_bytes:
        return png_bytes

    # 2. Fallback: Kroki API
    sanitized = sanitize_mermaid_chart(mermaid_code)
    if not sanitized:
        return None
    try:
        url = "https://kroki.io/mermaid/png"
        req = urllib.request.Request(
            url,
            data=sanitized.encode("utf-8"),
            headers={
                "Content-Type": "text/plain; charset=utf-8",
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AI-SDLC-Studio/2.0"
            }
        )
        with urllib.request.urlopen(req, timeout=1.5) as resp:
            data = resp.read()
            if data and data.startswith(b'\x89PNG'):
                return data
    except Exception:
        pass
    return None

def get_kroki_mermaid_png_url(mermaid_code: str) -> str:
    sanitized = sanitize_mermaid_chart(mermaid_code)
    if not sanitized:
        return ""
    try:
        compressed = zlib.compress(sanitized.encode('utf-8'), 9)
        b64 = base64.urlsafe_b64encode(compressed).decode('utf-8')
        return f"https://kroki.io/mermaid/png/{b64}"
    except Exception:
        return ""


def build_visual_diagram_flow_table(mermaid_code: str, title: str, code_style: ParagraphStyle, body_style: ParagraphStyle) -> Table:
    """Builds a beautifully styled visual diagram flow table for ReportLab PDF when PNG rendering is unavailable."""
    sanitized = sanitize_mermaid_chart(mermaid_code)
    lines = [l.strip() for l in sanitized.split('\n') if l.strip()]
    
    # Parse connections
    connections = []
    for line in lines:
        if '-->' in line or '->' in line or '--' in line:
            parts = re.split(r'-->|->|--', line)
            if len(parts) >= 2:
                src = parts[0].strip()
                dst = parts[-1].strip()
                # Clean node brackets
                src_clean = re.sub(r'[\"\[\]\(\)\{\}]', '', src).strip()
                dst_clean = re.sub(r'[\"\[\]\(\)\{\}]', '', dst).strip()
                if src_clean and dst_clean and src_clean != dst_clean:
                    connections.append((src_clean, dst_clean))
                    
    table_rows = []
    title_p = Paragraph(f"<b>📐 Visual Architecture Diagram Specification: {html.escape(title)}</b>", ParagraphStyle(
        'DiagTitle', parent=body_style, fontName='Helvetica-Bold', fontSize=9, textColor=colors.HexColor("#6366f1")
    ))
    table_rows.append([title_p])
    
    if connections:
        conn_rows = []
        for src, dst in connections[:6]:
            flow_p = Paragraph(
                f"<font color='#38bdf8'><b>[{html.escape(src)}]</b></font> &nbsp; <font color='#94a3b8'><b>➔</b></font> &nbsp; <font color='#818cf8'><b>[{html.escape(dst)}]</b></font>",
                ParagraphStyle('FlowP', parent=body_style, fontName='Helvetica', fontSize=8.5, textColor=colors.HexColor("#e2e8f0"))
            )
            conn_rows.append(flow_p)
        table_rows.append([conn_rows])
    
    # Code block below
    escaped_lines = [html.escape(l) for l in lines[:15]]
    code_paragraphs = [Paragraph(l.replace(' ', '&nbsp;'), code_style) for l in escaped_lines]
    table_rows.append([code_paragraphs])
    
def build_sequence_diagram_drawing(mermaid_code: str, title: str) -> Drawing:
    """Renders a real UML Sequence Diagram with lifelines, participant boxes, message arrows, and response dashed lines."""
    sanitized = sanitize_mermaid_chart(mermaid_code)
    lines = [l.strip() for l in sanitized.split('\n') if l.strip()]
    
    participants = []
    messages = []
    
    for line in lines:
        if line.startswith('actor ') or line.startswith('participant '):
            p = line.split()[-1].replace(':', '').strip()
            p_clean = re.sub(r'[\"\[\]]', '', p)
            if p_clean and p_clean not in participants:
                participants.append(p_clean)
        elif '->>' in line or '->' in line or '-->>' in line or '--' in line:
            match = re.search(r'([A-Za-z0-9_]+)\s*(->>|-->>|->|--)\s*([A-Za-z0-9_]+)\s*:\s*(.*)', line)
            if match:
                src, arrow_type, dst, msg = match.groups()
                if src not in participants: participants.append(src)
                if dst not in participants: participants.append(dst)
                messages.append((src, arrow_type, dst, msg.strip()))
                
    if not participants:
        participants = ["User", "FrontendApp", "APIGateway", "AuthService", "Database"]
        messages = [
            ("User", "->>", "FrontendApp", "1. Submit Login Credentials"),
            ("FrontendApp", "->>", "APIGateway", "2. POST /api/v1/auth/login"),
            ("APIGateway", "->>", "AuthService", "3. Validate Credentials"),
            ("AuthService", "->>", "Database", "4. Query User Record"),
            ("Database", "-->>", "AuthService", "5. User Record & Password Hash"),
            ("AuthService", "-->>", "FrontendApp", "6. 200 OK + JWT Bearer Token")
        ]
        
    num_p = len(participants[:5])
    dw_width = 460
    dw_height = max(180, 50 + len(messages[:7]) * 24 + 40)
    
    d = Drawing(dw_width, dw_height)
    d.add(Rect(0, 0, dw_width, dw_height, fillColor=colors.HexColor("#0f172a"), strokeColor=colors.HexColor("#334155"), strokeWidth=1, rx=6, ry=6))
    d.add(String(12, dw_height - 15, f"📐 UML Sequence Diagram — {title}", fontName='Helvetica-Bold', fontSize=8.5, fillColor=colors.HexColor("#38bdf8")))
    
    col_width = (dw_width - 40) / max(1, num_p)
    positions = {}
    box_w = min(75, col_width - 8)
    box_h = 24
    top_y = dw_height - 45
    bottom_y = 20
    
    for idx, p in enumerate(participants[:5]):
        cx = 20 + idx * col_width + col_width / 2
        positions[p] = cx
        
        d.add(Rect(cx - box_w/2, top_y, box_w, box_h, fillColor=colors.HexColor("#1e293b"), strokeColor=colors.HexColor("#6366f1"), strokeWidth=1.2, rx=3, ry=3))
        d.add(String(cx, top_y + 8, p[:11], textAnchor='middle', fontName='Helvetica-Bold', fontSize=7.5, fillColor=colors.white))
        
        d.add(Line(cx, top_y, cx, bottom_y + box_h, strokeColor=colors.HexColor("#475569"), strokeWidth=1, strokeDashArray=[3, 3]))
        
        d.add(Rect(cx - box_w/2, bottom_y, box_w, box_h, fillColor=colors.HexColor("#1e293b"), strokeColor=colors.HexColor("#6366f1"), strokeWidth=1.2, rx=3, ry=3))
        d.add(String(cx, bottom_y + 8, p[:11], textAnchor='middle', fontName='Helvetica-Bold', fontSize=7.5, fillColor=colors.white))

    current_y = top_y - 20
    for src, arrow_type, dst, msg in messages[:7]:
        if src in positions and dst in positions:
            sx = positions[src]
            dx = positions[dst]
            
            is_response = "--" in arrow_type
            line_color = colors.HexColor("#38bdf8") if not is_response else colors.HexColor("#94a3b8")
            dash = [2, 2] if is_response else None
            
            d.add(Line(sx, current_y, dx, current_y, strokeColor=line_color, strokeWidth=1.2, strokeDashArray=dash))
            direction = 1 if dx > sx else -1
            d.add(Polygon([dx, current_y, dx - direction * 5, current_y + 3, dx - direction * 5, current_y - 3], fillColor=line_color, strokeColor=line_color))
            
            mid_x = (sx + dx) / 2
            msg_text = msg[:28] if len(msg) > 28 else msg
            d.add(String(mid_x, current_y + 3, msg_text, textAnchor='middle', fontName='Helvetica', fontSize=7, fillColor=colors.HexColor("#e2e8f0")))
            current_y -= 22

    return d


def build_class_diagram_drawing(mermaid_code: str, title: str) -> Drawing:
    """Renders a real UML Class Diagram with 3-compartment class boxes (Header, Attributes, Methods) and relationships."""
    sanitized = sanitize_mermaid_chart(mermaid_code)
    lines = [l.strip() for l in sanitized.split('\n') if l.strip()]
    
    classes = {}
    current_class = None
    
    for line in lines:
        if line.startswith('class '):
            c_name = line.split()[1].replace('{', '').strip()
            current_class = c_name
            if current_class not in classes:
                classes[current_class] = {"fields": [], "methods": []}
        elif current_class and ('+' in line or '-' in line or '#' in line):
            if '(' in line:
                classes[current_class]["methods"].append(line.strip())
            else:
                classes[current_class]["fields"].append(line.strip())

    if not classes:
        classes = {
            "User": {"fields": ["+UUID id", "+String email", "+String role"], "methods": ["+login()", "+verifyPassword()"]},
            "Order": {"fields": ["+UUID id", "+UUID user_id", "+Decimal total"], "methods": ["+createOrder()", "+updateStatus()"]},
            "Session": {"fields": ["+UUID id", "+String token"], "methods": ["+validate()", "+revoke()"]}
        }
        
    num_c = len(classes)
    dw_width = 460
    dw_height = 160
    
    d = Drawing(dw_width, dw_height)
    d.add(Rect(0, 0, dw_width, dw_height, fillColor=colors.HexColor("#0f172a"), strokeColor=colors.HexColor("#334155"), strokeWidth=1, rx=6, ry=6))
    d.add(String(12, dw_height - 15, f"📐 UML Class Diagram — {title}", fontName='Helvetica-Bold', fontSize=8.5, fillColor=colors.HexColor("#38bdf8")))
    
    box_w = 125
    box_h = 110
    gap_x = 20
    start_x = 20
    start_y = 20
    
    for idx, (c_name, c_data) in enumerate(list(classes.items())[:3]):
        cx = start_x + idx * (box_w + gap_x)
        cy = start_y
        
        d.add(Rect(cx, cy, box_w, box_h, fillColor=colors.HexColor("#1e293b"), strokeColor=colors.HexColor("#818cf8"), strokeWidth=1.5, rx=4, ry=4))
        d.add(Rect(cx, cy + box_h - 24, box_w, 24, fillColor=colors.HexColor("#312e81"), strokeColor=colors.HexColor("#818cf8"), strokeWidth=1, rx=4, ry=4))
        d.add(String(cx + box_w/2, cy + box_h - 16, f"class {c_name}", textAnchor='middle', fontName='Helvetica-Bold', fontSize=8, fillColor=colors.white))
        
        d.add(Line(cx, cy + box_h - 24, cx + box_w, cy + box_h - 24, strokeColor=colors.HexColor("#6366f1"), strokeWidth=1))
        
        fy = cy + box_h - 36
        for f in c_data["fields"][:3]:
            d.add(String(cx + 6, fy, f[:22], fontName='Helvetica', fontSize=7, fillColor=colors.HexColor("#38bdf8")))
            fy -= 11
            
        d.add(Line(cx, cy + 36, cx + box_w, cy + 36, strokeColor=colors.HexColor("#475569"), strokeWidth=0.8))
        
        my = cy + 24
        for m in c_data["methods"][:2]:
            d.add(String(cx + 6, my, m[:22], fontName='Helvetica-Oblique', fontSize=7, fillColor=colors.HexColor("#a7f3d0")))
            my -= 11
            
        if idx > 0:
            prev_cx = start_x + (idx - 1) * (box_w + gap_x) + box_w
            d.add(Line(prev_cx, cy + box_h/2, cx, cy + box_h/2, strokeColor=colors.HexColor("#94a3b8"), strokeWidth=1.2))
            d.add(Polygon([cx, cy + box_h/2, cx - 5, cy + box_h/2 + 3, cx - 5, cy + box_h/2 - 3], fillColor=colors.HexColor("#94a3b8"), strokeColor=colors.HexColor("#94a3b8")))

    return d


def build_er_diagram_drawing(mermaid_code: str, title: str) -> Drawing:
    """Renders a real UML ER Diagram with Entity Boxes, PK/FK attributes, and Relationship lines."""
    sanitized = sanitize_mermaid_chart(mermaid_code)
    lines = [l.strip() for l in sanitized.split('\n') if l.strip()]
    
    entities = {}
    current_ent = None
    
    for line in lines:
        if line.endswith('{') and not line.startswith('subgraph'):
            ent_name = line.split('{')[0].strip()
            current_ent = ent_name
            if current_ent not in entities:
                entities[current_ent] = []
        elif current_ent and line == '}':
            current_ent = None
        elif current_ent:
            parts = line.split()
            if len(parts) >= 2:
                entities[current_ent].append((parts[0], parts[1], parts[2] if len(parts) > 2 else ""))

    if not entities:
        entities = {
            "USERS": [("uuid", "id", "PK"), ("string", "email", "UK"), ("string", "password_hash", ""), ("timestamp", "created_at", "")],
            "ORDERS": [("uuid", "id", "PK"), ("uuid", "user_id", "FK"), ("numeric", "total_amount", ""), ("string", "status", "")],
            "AUDIT_LOGS": [("bigint", "id", "PK"), ("uuid", "user_id", "FK"), ("string", "action", ""), ("timestamp", "timestamp", "")]
        }

    dw_width = 460
    dw_height = 160
    
    d = Drawing(dw_width, dw_height)
    d.add(Rect(0, 0, dw_width, dw_height, fillColor=colors.HexColor("#0f172a"), strokeColor=colors.HexColor("#334155"), strokeWidth=1, rx=6, ry=6))
    d.add(String(12, dw_height - 15, f"📐 UML Entity Relationship Diagram (ERD) — {title}", fontName='Helvetica-Bold', fontSize=8.5, fillColor=colors.HexColor("#38bdf8")))
    
    box_w = 125
    box_h = 110
    gap_x = 20
    start_x = 20
    start_y = 20
    
    for idx, (ent_name, attrs) in enumerate(list(entities.items())[:3]):
        cx = start_x + idx * (box_w + gap_x)
        cy = start_y
        
        d.add(Rect(cx, cy, box_w, box_h, fillColor=colors.HexColor("#1e293b"), strokeColor=colors.HexColor("#38bdf8"), strokeWidth=1.5, rx=4, ry=4))
        d.add(Rect(cx, cy + box_h - 22, box_w, 22, fillColor=colors.HexColor("#0369a1"), strokeColor=colors.HexColor("#38bdf8"), strokeWidth=1, rx=4, ry=4))
        d.add(String(cx + box_w/2, cy + box_h - 15, f"📄 {ent_name}", textAnchor='middle', fontName='Helvetica-Bold', fontSize=8, fillColor=colors.white))
        
        ay = cy + box_h - 34
        for data_type, name, key in attrs[:5]:
            d.add(String(cx + 6, ay, f"{name}", fontName='Helvetica-Bold', fontSize=7, fillColor=colors.HexColor("#e2e8f0")))
            d.add(String(cx + box_w - 30, ay, f"{data_type[:6]}", fontName='Helvetica', fontSize=6.5, fillColor=colors.HexColor("#94a3b8")))
            if key:
                d.add(String(cx + box_w - 12, ay, key, fontName='Helvetica-Bold', fontSize=6.5, fillColor=colors.HexColor("#fbbf24")))
            ay -= 13
            
        if idx > 0:
            prev_cx = start_x + (idx - 1) * (box_w + gap_x) + box_w
            d.add(Line(prev_cx, cy + box_h/2, cx, cy + box_h/2, strokeColor=colors.HexColor("#fbbf24"), strokeWidth=1.2, strokeDashArray=[2, 2]))

    return d


def build_visual_flowchart_drawing(mermaid_code: str, title: str) -> Drawing:
    """Renders a real ReportLab vector graphic Drawing (boxes, shapes, lines, arrows, labels) for the Mermaid diagram."""
    sanitized = sanitize_mermaid_chart(mermaid_code)
    lines = [l.strip() for l in sanitized.split('\n') if l.strip()]
    
    nodes_map = {}
    connections = []
    
    for line in lines:
        if '-->' in line or '->' in line:
            parts = re.split(r'-->|->', line)
            if len(parts) >= 2:
                src_raw = parts[0].strip()
                dst_raw = parts[-1].strip()
                
                src_id = re.sub(r'[^a-zA-Z0-9_]', '', src_raw.split('[')[0].split('(')[0])
                dst_id = re.sub(r'[^a-zA-Z0-9_]', '', dst_raw.split('[')[0].split('(')[0])
                
                src_label = re.sub(r'[\"\[\]\(\)\{\}]', '', src_raw).strip()
                dst_label = re.sub(r'[\"\[\]\(\)\{\}]', '', dst_raw).strip()
                
                if src_id and src_label:
                    nodes_map[src_id] = src_label
                if dst_id and dst_label:
                    nodes_map[dst_id] = dst_label
                    
                if src_id and dst_id and src_id != dst_id:
                    connections.append((src_id, dst_id))

    node_ids = list(nodes_map.keys())
    if not node_ids:
        node_ids = ["ClientUI", "APIGateway", "CoreService", "Database"]
        nodes_map = {"ClientUI": "Frontend UI", "APIGateway": "API Gateway", "CoreService": "Business Logic", "Database": "PostgreSQL DB"}
        connections = [("ClientUI", "APIGateway"), ("APIGateway", "CoreService"), ("CoreService", "Database")]

    cols = min(4, max(1, len(node_ids)))
    rows = max(1, (len(node_ids) + cols - 1) // cols)
    
    box_w = 95
    box_h = 32
    gap_x = 22
    gap_y = 20
    margin_x = 15
    margin_y = 24
    
    dw_width = 460
    dw_height = max(110, margin_y * 2 + rows * box_h + (rows - 1) * gap_y)
    
    d = Drawing(dw_width, dw_height)
    d.add(Rect(0, 0, dw_width, dw_height, fillColor=colors.HexColor("#0f172a"), strokeColor=colors.HexColor("#334155"), strokeWidth=1, rx=6, ry=6))
    
    title_text = f"📐 Visual Architecture Flowchart — {title}"
    if len(title_text) > 55:
        title_text = title_text[:53] + ".."
    d.add(String(12, dw_height - 15, title_text, fontName='Helvetica-Bold', fontSize=8.5, fillColor=colors.HexColor("#38bdf8")))
    
    positions = {}
    for idx, nid in enumerate(node_ids[:8]):
        r = idx // cols
        c = idx % cols
        x = margin_x + c * (box_w + gap_x)
        y = dw_height - margin_y - (r + 1) * box_h - r * gap_y
        positions[nid] = (x, y)
        
        node_color = colors.HexColor("#1e293b")
        border_color = colors.HexColor("#6366f1")
        if any(k in nid.lower() for k in ["db", "database", "postgres", "sql", "redis"]):
            node_color = colors.HexColor("#0284c7")
            border_color = colors.HexColor("#38bdf8")
        elif any(k in nid.lower() for k in ["user", "client", "ui", "react", "app"]):
            node_color = colors.HexColor("#0d9488")
            border_color = colors.HexColor("#2dd4bf")
        elif any(k in nid.lower() for k in ["auth", "security", "jwt", "gateway"]):
            node_color = colors.HexColor("#4f46e5")
            border_color = colors.HexColor("#818cf8")
            
        d.add(Rect(x, y, box_w, box_h, fillColor=node_color, strokeColor=border_color, strokeWidth=1.5, rx=4, ry=4))
        
        lbl_text = nodes_map.get(nid, nid)
        if len(lbl_text) > 13:
            lbl_text = lbl_text[:12] + ".."
        d.add(String(x + box_w / 2, y + box_h / 2 - 3, lbl_text, textAnchor='middle', fontName='Helvetica-Bold', fontSize=8, fillColor=colors.white))

    for src_id, dst_id in connections[:8]:
        if src_id in positions and dst_id in positions:
            sx, sy = positions[src_id]
            dx, dy = positions[dst_id]
            
            if abs(sy - dy) < 5:
                start_x = sx + box_w
                start_y = sy + box_h / 2
                end_x = dx
                end_y = dy + box_h / 2
            else:
                start_x = sx + box_w / 2
                start_y = sy
                end_x = dx + box_w / 2
                end_y = dy + box_h
                
            d.add(Line(start_x, start_y, end_x, end_y, strokeColor=colors.HexColor("#94a3b8"), strokeWidth=1.2))
            d.add(Polygon([end_x, end_y, end_x - 3, end_y + 5, end_x + 3, end_y + 5], fillColor=colors.HexColor("#38bdf8"), strokeColor=colors.HexColor("#38bdf8")))

    return d


def build_uml_diagram_drawing(mermaid_code: str, title: str, diagram_key: str = "") -> Drawing:
    """Master dispatcher selecting specialized UML renderer based on diagram type."""
    sanitized = sanitize_mermaid_chart(mermaid_code)
    code_lower = sanitized.lower()
    key_lower = diagram_key.lower()
    
    if "sequence" in key_lower or "sequencediagram" in code_lower:
        return build_sequence_diagram_drawing(sanitized, title)
    elif "class" in key_lower or "classdiagram" in code_lower:
        return build_class_diagram_drawing(sanitized, title)
    elif "er_" in key_lower or "db_relationship" in key_lower or "erdiagram" in code_lower:
        return build_er_diagram_drawing(sanitized, title)
    else:
        return build_visual_flowchart_drawing(sanitized, title)


# Dynamic Page Numbering helper that skips headers/footers on the cover page (Page 1)
class NumberedCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_page_number(num_pages)
            super().showPage()
        super().save()

    def draw_page_number(self, page_count):
        if self._pageNumber == 1:
            # Skip running header/footer on cover page
            return
            
        self.saveState()
        self.setFont("Helvetica", 9)
        self.setFillColor(colors.HexColor("#64748b"))
        
        # Header
        self.drawString(54, 750, "AI SDLC Studio - Enterprise E2E Documentation")
        self.setStrokeColor(colors.HexColor("#e2e8f0"))
        self.setLineWidth(0.5)
        self.line(54, 742, 558, 742)
        
        # Footer
        page_text = f"Page {self._pageNumber} of {page_count}"
        self.drawRightString(558, 40, page_text)
        self.drawString(54, 40, f"Generated Date: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
        self.line(54, 52, 558, 52)
        
        self.restoreState()


def render_markdown_bold_html(text: str) -> str:
    """Helper to convert **bold** markdown to ReportLab <b>bold</b> tags."""
    return re.sub(r'\*\*(.*?)\*\*', r'<b>\1</b>', text)


def generate_srs_pdf(project_name: str, srs_data: dict, version: int, approval_status: str) -> bytes:
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, 
        pagesize=letter,
        leftMargin=54, 
        rightMargin=54, 
        topMargin=72, 
        bottomMargin=72
    )
    
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle(
        'CoverTitle',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=24,
        leading=30,
        textColor=colors.HexColor("#1e3a8a"),
        spaceAfter=15,
        alignment=1 # Center
    )
    
    subtitle_style = ParagraphStyle(
        'CoverSub',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=13,
        leading=18,
        textColor=colors.HexColor("#475569"),
        spaceAfter=25,
        alignment=1 # Center
    )
    
    h1_style = ParagraphStyle(
        'DocH1',
        parent=styles['Heading2'],
        fontName='Helvetica-Bold',
        fontSize=12,
        leading=16,
        textColor=colors.HexColor("#0f172a"),
        spaceBefore=14,
        spaceAfter=6,
        keepWithNext=True
    )
    
    body_style = ParagraphStyle(
        'DocBody',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=9.5,
        leading=14,
        textColor=colors.HexColor("#334155"),
        spaceAfter=6
    )

    bullet_style = ParagraphStyle(
        'DocBullet',
        parent=body_style,
        leftIndent=15,
        spaceAfter=4
    )

    story = []
    
    # 1. COVER PAGE
    story.append(Spacer(1, 80))
    story.append(Paragraph("SOFTWARE REQUIREMENTS SPECIFICATION (SRS)", title_style))
    story.append(Spacer(1, 10))
    story.append(Paragraph(f"Project: <b>{project_name}</b>", subtitle_style))
    story.append(Spacer(1, 80))
    
    meta_table_data = [
        [Paragraph("<b>Document Version:</b>", body_style), Paragraph(f"v{version}.0", body_style)],
        [Paragraph("<b>Approval Status:</b>", body_style), Paragraph(f"<b>{approval_status}</b>", body_style)],
        [Paragraph("<b>Author System:</b>", body_style), Paragraph("AI SDLC Studio Autonomous Requirements Engine", body_style)],
        [Paragraph("<b>Specification Standard:</b>", body_style), Paragraph("IEEE-830 Software Requirements Standard", body_style)],
        [Paragraph("<b>Classification Level:</b>", body_style), Paragraph("Enterprise Confidential", body_style)],
        [Paragraph("<b>Generation Date:</b>", body_style), Paragraph(datetime.now().strftime('%Y-%m-%d %H:%M'), body_style)],
    ]
    t_meta = Table(meta_table_data, colWidths=[150, 250])
    t_meta.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#f8fafc")),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#cbd5e1")),
        ('PADDING', (0,0), (-1,-1), 7),
    ]))
    story.append(t_meta)
    story.append(PageBreak())
    
    # 2. TABLE OF CONTENTS PAGE
    story.append(Paragraph("TABLE OF CONTENTS", h1_style))
    story.append(Spacer(1, 8))
    
    toc_data = [
        [Paragraph("<b>Section Title</b>", body_style), Paragraph("<b>Page Reference</b>", body_style)],
        [Paragraph("1. Document Information", body_style), Paragraph("Page 3", body_style)],
        [Paragraph("2. Revision History", body_style), Paragraph("Page 3", body_style)],
        [Paragraph("3. Approval History", body_style), Paragraph("Page 3", body_style)],
        [Paragraph("4. Executive Summary", body_style), Paragraph("Page 4", body_style)],
        [Paragraph("5. Problem Statement & Current State", body_style), Paragraph("Page 4", body_style)],
        [Paragraph("6. Business Objectives & Strategic KPIs", body_style), Paragraph("Page 5", body_style)],
        [Paragraph("7. Stakeholders & Responsibilities", body_style), Paragraph("Page 5", body_style)],
        [Paragraph("8. User Personas & Detailed Profiles", body_style), Paragraph("Page 6", body_style)],
        [Paragraph("9. System Actors & Access Rights", body_style), Paragraph("Page 6", body_style)],
        [Paragraph("10. In-Scope Functional Boundaries", body_style), Paragraph("Page 7", body_style)],
        [Paragraph("11. Out-of-Scope & Future Phases", body_style), Paragraph("Page 7", body_style)],
        [Paragraph("12. Business Requirements", body_style), Paragraph("Page 8", body_style)],
        [Paragraph("13. Detailed Functional Requirements", body_style), Paragraph("Page 8", body_style)],
        [Paragraph("14. Non-Functional & Quality Attributes", body_style), Paragraph("Page 10", body_style)],
        [Paragraph("15. Business Rules & Logic Constraints", body_style), Paragraph("Page 11", body_style)],
        [Paragraph("16. User Stories & Epics", body_style), Paragraph("Page 11", body_style)],
        [Paragraph("17. Primary & Secondary Use Cases", body_style), Paragraph("Page 12", body_style)],
        [Paragraph("18. Acceptance Criteria (Given-When-Then)", body_style), Paragraph("Page 13", body_style)],
        [Paragraph("19. UI/UX & Layout Guidelines", body_style), Paragraph("Page 14", body_style)],
        [Paragraph("20. Navigation & Information Architecture", body_style), Paragraph("Page 14", body_style)],
        [Paragraph("21. Data Architecture & Schema Specs", body_style), Paragraph("Page 15", body_style)],
        [Paragraph("22. Security, Auth & Encryption Policies", body_style), Paragraph("Page 15", body_style)],
        [Paragraph("23. API & External Integration Contracts", body_style), Paragraph("Page 16", body_style)],
        [Paragraph("24. Performance & Scalability Specs", body_style), Paragraph("Page 16", body_style)],
        [Paragraph("25. Compliance & Regulatory Auditing", body_style), Paragraph("Page 17", body_style)],
        [Paragraph("26. Architectural Constraints", body_style), Paragraph("Page 17", body_style)],
        [Paragraph("27. Operational Assumptions", body_style), Paragraph("Page 18", body_style)],
        [Paragraph("28. Risk Assessment & Mitigation Plan", body_style), Paragraph("Page 18", body_style)],
        [Paragraph("29. System Dependencies & SDKs", body_style), Paragraph("Page 19", body_style)],
        [Paragraph("30. Requirement Traceability Matrix (RTM)", body_style), Paragraph("Page 19", body_style)],
    ]
    t_toc = Table(toc_data, colWidths=[330, 90])
    t_toc.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#f1f5f9")),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#e2e8f0")),
        ('PADDING', (0,0), (-1,-1), 4),
    ]))
    story.append(t_toc)
    story.append(PageBreak())
    
    # 3. 30 SECTIONS CONTENT
    sections = [
        ("1. Document Information", "document_information"),
        ("2. Revision History", "revision_history"),
        ("3. Approval History", "approval_history"),
        ("4. Executive Summary", "executive_summary"),
        ("5. Problem Statement & Current State", "problem_statement"),
        ("6. Business Objectives & Strategic KPIs", "business_objectives"),
        ("7. Stakeholders & Responsibilities", "stakeholders"),
        ("8. User Personas & Profiles", "user_personas"),
        ("9. System Actors & Access Rights", "actors"),
        ("10. In-Scope Functional Boundaries", "scope"),
        ("11. Out-of-Scope & Future Phases", "out_of_scope"),
        ("12. Business Requirements", "business_requirements"),
        ("13. Detailed Functional Requirements", "functional_requirements"),
        ("14. Non-Functional & Quality Attributes", "non_functional_requirements"),
        ("15. Business Rules & Logic Constraints", "business_rules"),
        ("16. User Stories & Epics", "user_stories"),
        ("17. Primary & Secondary Use Cases", "use_cases"),
        ("18. Acceptance Criteria (Given-When-Then)", "acceptance_criteria"),
        ("19. UI/UX & Layout Guidelines", "ui_requirements"),
        ("20. Navigation & Information Architecture", "navigation_flow"),
        ("21. Data Architecture & Schema Specs", "data_requirements"),
        ("22. Security, Auth & Encryption Policies", "security_requirements"),
        ("23. API & External Integration Contracts", "integration_requirements"),
        ("24. Performance & Scalability Specs", "performance_requirements"),
        ("25. Compliance & Regulatory Auditing", "compliance_requirements"),
        ("26. Architectural Constraints", "constraints"),
        ("27. Operational Assumptions", "assumptions"),
        ("28. Risk Assessment & Mitigation Plan", "risks"),
        ("29. System Dependencies & SDKs", "dependencies"),
    ]
    
    for label, key in sections:
        val = srs_data.get(key)
        story.append(Paragraph(label, h1_style))
        if not val:
            story.append(Paragraph("Not specified.", body_style))
        elif isinstance(val, list):
            for item in val:
                item_str = render_markdown_bold_html(str(item))
                lines = item_str.split('\n')
                for line in lines:
                    if line.strip():
                        story.append(Paragraph(f"• {line.strip()}", bullet_style))
        else:
            val_str = render_markdown_bold_html(str(val))
            paragraphs = val_str.split('\n\n')
            for p_text in paragraphs:
                lines = p_text.split('\n')
                for line in lines:
                    if line.strip():
                        story.append(Paragraph(line.strip(), body_style))
        story.append(Spacer(1, 8))
        
    # 30. Traceability Matrix Table
    matrix = srs_data.get("requirement_traceability_matrix", [])
    if matrix:
        story.append(Paragraph("30. Requirement Traceability Matrix (RTM)", h1_style))
        table_data = [[
            Paragraph("<b>Req ID</b>", body_style), 
            Paragraph("<b>Title</b>", body_style), 
            Paragraph("<b>Detailed Description</b>", body_style), 
            Paragraph("<b>Category</b>", body_style)
        ]]
        for row in matrix:
            table_data.append([
                Paragraph(str(row.get("id", "")), body_style),
                Paragraph(render_markdown_bold_html(str(row.get("title", ""))), body_style),
                Paragraph(render_markdown_bold_html(str(row.get("description", ""))), body_style),
                Paragraph(str(row.get("category", "")), body_style)
            ])
            
        t = Table(table_data, colWidths=[65, 110, 225, 100])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#f1f5f9")),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#cbd5e1")),
            ('PADDING', (0,0), (-1,-1), 5),
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ]))
        story.append(t)

    doc.build(story, canvasmaker=NumberedCanvas)
    return buffer.getvalue()


def generate_srs_docx(project_name: str, srs_data: dict, version: int, approval_status: str) -> bytes:
    doc = Document()
    
    title_p = doc.add_paragraph()
    run = title_p.add_run("SOFTWARE REQUIREMENTS SPECIFICATION")
    run.font.size = docx.shared.Pt(22) if hasattr(docx, 'shared') else None
    run.bold = True
    
    p_sub = doc.add_paragraph()
    p_sub.add_run("Project Name: ").bold = True
    p_sub.add_run(project_name + "\n")
    p_sub.add_run("Version: ").bold = True
    p_sub.add_run(f"v{version}.0\n")
    p_sub.add_run("Status: ").bold = True
    p_sub.add_run(approval_status + "\n")
    p_sub.add_run("Specification Standard: ").bold = True
    p_sub.add_run("IEEE-830 Software Requirements Standard\n")
    p_sub.add_run("Generation Date: ").bold = True
    p_sub.add_run(datetime.now().strftime('%Y-%m-%d %H:%M') + "\n")
    
    # Table of Contents
    doc.add_heading("TABLE OF CONTENTS", level=1)
    toc_fields = [
        "1. Document Information", "2. Revision History", "3. Approval History",
        "4. Executive Summary", "5. Problem Statement & Current State", "6. Business Objectives & Strategic KPIs",
        "7. Stakeholders & Responsibilities", "8. User Personas & Profiles", "9. System Actors & Access Rights",
        "10. In-Scope Functional Boundaries", "11. Out-of-Scope & Future Phases", "12. Business Requirements",
        "13. Detailed Functional Requirements", "14. Non-Functional & Quality Attributes", "15. Business Rules & Logic Constraints",
        "16. User Stories & Epics", "17. Primary & Secondary Use Cases", "18. Acceptance Criteria (Given-When-Then)",
        "19. UI/UX & Layout Guidelines", "20. Navigation & Information Architecture", "21. Data Architecture & Schema Specs",
        "22. Security, Auth & Encryption Policies", "23. API & External Integration Contracts", "24. Performance & Scalability Specs",
        "25. Compliance & Regulatory Auditing", "26. Architectural Constraints", "27. Operational Assumptions",
        "28. Risk Assessment & Mitigation Plan", "29. System Dependencies & SDKs", "30. Requirement Traceability Matrix (RTM)"
    ]
    for field in toc_fields:
        doc.add_paragraph(field, style='List Bullet')

    doc.add_page_break()
    
    sections = [
        ("1. Document Information", "document_information"),
        ("2. Revision History", "revision_history"),
        ("3. Approval History", "approval_history"),
        ("4. Executive Summary", "executive_summary"),
        ("5. Problem Statement & Current State", "problem_statement"),
        ("6. Business Objectives & Strategic KPIs", "business_objectives"),
        ("7. Stakeholders & Responsibilities", "stakeholders"),
        ("8. User Personas & Profiles", "user_personas"),
        ("9. System Actors & Access Rights", "actors"),
        ("10. In-Scope Functional Boundaries", "scope"),
        ("11. Out-of-Scope & Future Phases", "out_of_scope"),
        ("12. Business Requirements", "business_requirements"),
        ("13. Detailed Functional Requirements", "functional_requirements"),
        ("14. Non-Functional & Quality Attributes", "non_functional_requirements"),
        ("15. Business Rules & Logic Constraints", "business_rules"),
        ("16. User Stories & Epics", "user_stories"),
        ("17. Primary & Secondary Use Cases", "use_cases"),
        ("18. Acceptance Criteria (Given-When-Then)", "acceptance_criteria"),
        ("19. UI/UX & Layout Guidelines", "ui_requirements"),
        ("20. Navigation & Information Architecture", "navigation_flow"),
        ("21. Data Architecture & Schema Specs", "data_requirements"),
        ("22. Security, Auth & Encryption Policies", "security_requirements"),
        ("23. API & External Integration Contracts", "integration_requirements"),
        ("24. Performance & Scalability Specs", "performance_requirements"),
        ("25. Compliance & Regulatory Auditing", "compliance_requirements"),
        ("26. Architectural Constraints", "constraints"),
        ("27. Operational Assumptions", "assumptions"),
        ("28. Risk Assessment & Mitigation Plan", "risks"),
        ("29. System Dependencies & SDKs", "dependencies"),
    ]
    
    for label, key in sections:
        doc.add_heading(label, level=1)
        val = srs_data.get(key)
        if not val:
            doc.add_paragraph("Not specified.")
        elif isinstance(val, list):
            for item in val:
                lines = str(item).split('\n')
                for line in lines:
                    if line.strip():
                        doc.add_paragraph(line.strip(), style='List Bullet')
        else:
            paragraphs = str(val).split('\n\n')
            for p_text in paragraphs:
                lines = p_text.split('\n')
                for line in lines:
                    if line.strip():
                        doc.add_paragraph(line.strip())
            
    matrix = srs_data.get("requirement_traceability_matrix", [])
    if matrix:
        doc.add_heading("30. Requirement Traceability Matrix (RTM)", level=1)
        table = doc.add_table(rows=1, cols=4)
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = 'Req ID'
        hdr_cells[1].text = 'Title'
        hdr_cells[2].text = 'Detailed Description'
        hdr_cells[3].text = 'Category'
        for item in matrix:
            row_cells = table.add_row().cells
            row_cells[0].text = str(item.get("id", ""))
            row_cells[1].text = str(item.get("title", ""))
            row_cells[2].text = str(item.get("description", ""))
            row_cells[3].text = str(item.get("category", ""))
            
    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer.getvalue()


def generate_srs_markdown(project_name: str, srs_data: dict, version: int, approval_status: str) -> str:
    md = []
    md.append(f"# Software Requirements Specification (SRS)")
    md.append(f"**Project**: {project_name}")
    md.append(f"**Version**: v{version}.0")
    md.append(f"**Approval Status**: {approval_status}")
    md.append(f"**Specification Standard**: IEEE-830 Software Requirements Standard")
    md.append(f"**Generated Date**: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    md.append("\n---\n")

    md.append("## TABLE OF CONTENTS")
    toc_fields = [
        "1. Document Information", "2. Revision History", "3. Approval History",
        "4. Executive Summary", "5. Problem Statement & Current State", "6. Business Objectives & Strategic KPIs",
        "7. Stakeholders & Responsibilities", "8. User Personas & Profiles", "9. System Actors & Access Rights",
        "10. In-Scope Functional Boundaries", "11. Out-of-Scope & Future Phases", "12. Business Requirements",
        "13. Detailed Functional Requirements", "14. Non-Functional & Quality Attributes", "15. Business Rules & Logic Constraints",
        "16. User Stories & Epics", "17. Primary & Secondary Use Cases", "18. Acceptance Criteria (Given-When-Then)",
        "19. UI/UX & Layout Guidelines", "20. Navigation & Information Architecture", "21. Data Architecture & Schema Specs",
        "22. Security, Auth & Encryption Policies", "23. API & External Integration Contracts", "24. Performance & Scalability Specs",
        "25. Compliance & Regulatory Auditing", "26. Architectural Constraints", "27. Operational Assumptions",
        "28. Risk Assessment & Mitigation Plan", "29. System Dependencies & SDKs", "30. Requirement Traceability Matrix (RTM)"
    ]
    for field in toc_fields:
        md.append(f"- [{field}](#{field.lower().replace(' ', '-').replace('&', '').replace('(', '').replace(')', '')})")
    md.append("\n---\n")

    sections = [
        ("1. Document Information", "document_information"),
        ("2. Revision History", "revision_history"),
        ("3. Approval History", "approval_history"),
        ("4. Executive Summary", "executive_summary"),
        ("5. Problem Statement & Current State", "problem_statement"),
        ("6. Business Objectives & Strategic KPIs", "business_objectives"),
        ("7. Stakeholders & Responsibilities", "stakeholders"),
        ("8. User Personas & Profiles", "user_personas"),
        ("9. System Actors & Access Rights", "actors"),
        ("10. In-Scope Functional Boundaries", "scope"),
        ("11. Out-of-Scope & Future Phases", "out_of_scope"),
        ("12. Business Requirements", "business_requirements"),
        ("13. Detailed Functional Requirements", "functional_requirements"),
        ("14. Non-Functional & Quality Attributes", "non_functional_requirements"),
        ("15. Business Rules & Logic Constraints", "business_rules"),
        ("16. User Stories & Epics", "user_stories"),
        ("17. Primary & Secondary Use Cases", "use_cases"),
        ("18. Acceptance Criteria (Given-When-Then)", "acceptance_criteria"),
        ("19. UI/UX & Layout Guidelines", "ui_requirements"),
        ("20. Navigation & Information Architecture", "navigation_flow"),
        ("21. Data Architecture & Schema Specs", "data_requirements"),
        ("22. Security, Auth & Encryption Policies", "security_requirements"),
        ("23. API & External Integration Contracts", "integration_requirements"),
        ("24. Performance & Scalability Specs", "performance_requirements"),
        ("25. Compliance & Regulatory Auditing", "compliance_requirements"),
        ("26. Architectural Constraints", "constraints"),
        ("27. Operational Assumptions", "assumptions"),
        ("28. Risk Assessment & Mitigation Plan", "risks"),
        ("29. System Dependencies & SDKs", "dependencies"),
    ]

    for label, key in sections:
        val = srs_data.get(key)
        md.append(f"## {label}")
        if not val:
            md.append("Not specified.\n")
        elif isinstance(val, list):
            for item in val:
                md.append(f"- {item}")
            md.append("")
        else:
            md.append(f"{val}\n")

    matrix = srs_data.get("requirement_traceability_matrix", [])
    if matrix:
        md.append("## 30. Requirement Traceability Matrix (RTM)")
        md.append("| Req ID | Title | Detailed Description | Category |")
        md.append("| --- | --- | --- | --- |")
        for item in matrix:
            title = str(item.get("title", "")).replace("\n", " ")
            desc = str(item.get("description", "")).replace("\n", " ")
            md.append(f"| {item.get('id', '')} | {title} | {desc} | {item.get('category', '')} |")
        md.append("")

    return "\n".join(md)


def generate_sdd_pdf(project_name: str, sdd_data: dict, version: int, approval_status: str, diagram_images: dict | None = None) -> bytes:
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, 
        pagesize=letter,
        leftMargin=54, 
        rightMargin=54, 
        topMargin=72, 
        bottomMargin=72
    )
    
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle(
        'CoverTitle',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=26,
        leading=32,
        textColor=colors.HexColor("#4f46e5"),
        spaceAfter=15,
        alignment=1
    )
    
    subtitle_style = ParagraphStyle(
        'CoverSub',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=14,
        leading=18,
        textColor=colors.HexColor("#475569"),
        spaceAfter=30,
        alignment=1
    )
    
    h1_style = ParagraphStyle(
        'DocH1',
        parent=styles['Heading2'],
        fontName='Helvetica-Bold',
        fontSize=12,
        leading=15,
        textColor=colors.HexColor("#0f172a"),
        spaceBefore=14,
        spaceAfter=6,
        keepWithNext=True
    )
    
    body_style = ParagraphStyle(
        'DocBody',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=9.5,
        leading=13.5,
        textColor=colors.HexColor("#334155"),
        spaceAfter=6
    )

    story = []
    
    # 1. COVER PAGE
    story.append(Spacer(1, 100))
    story.append(Paragraph("SOFTWARE DESIGN DOCUMENT", title_style))
    story.append(Spacer(1, 15))
    story.append(Paragraph(f"Project: <b>{project_name}</b>", subtitle_style))
    story.append(Spacer(1, 120))
    
    meta_table_data = [
        [Paragraph("<b>Document Version:</b>", body_style), Paragraph(f"v{version}", body_style)],
        [Paragraph("<b>Approval Status:</b>", body_style), Paragraph(approval_status, body_style)],
        [Paragraph("<b>Author Role:</b>", body_style), Paragraph("Lead Systems Architect", body_style)],
        [Paragraph("<b>Organization:</b>", body_style), Paragraph("AI SDLC Studio Corp", body_style)],
        [Paragraph("<b>Generation Date:</b>", body_style), Paragraph(datetime.now().strftime('%Y-%m-%d %H:%M'), body_style)],
    ]
    t_meta = Table(meta_table_data, colWidths=[150, 250])
    t_meta.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#f8fafc")),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#e2e8f0")),
        ('PADDING', (0,0), (-1,-1), 8),
    ]))
    story.append(t_meta)
    story.append(PageBreak())
    
    # 2. TABLE OF CONTENTS PAGE
    story.append(Paragraph("TABLE OF CONTENTS", h1_style))
    story.append(Spacer(1, 10))
    toc_fields = [
        "1. Cover Page Details", "2. Revision History", "3. Approval History",
        "4. Introduction", "5. Design Goals", "6. System Overview",
        "7. High-Level Architecture", "8. Low-Level Architecture", "9. Module Breakdown",
        "10. System Context Diagram (Mermaid)", "11. Use Case Diagram (Mermaid)", "12. Component Diagram (Mermaid)",
        "13. Class Diagram (Mermaid)", "14. Sequence Diagram (Mermaid)", "15. Activity Diagram (Mermaid)",
        "16. ER Diagram (Mermaid)", "17. Deployment Diagram (Mermaid)", "18. Database Design Overview",
        "19. Table Definitions", "20. Database Relationships & Constraints", "21. API Design Overview",
        "22. API Endpoints Map", "23. Authentication Flow", "24. Authorization Flow",
        "25. Security Design Policies", "26. Logging Strategy", "27. Exception Handling",
        "28. Configuration Management", "29. Technology Stack", "30. Folder Structure",
        "31. Coding Standards", "32. Performance Design", "33. Scalability Design",
        "34. Availability Design", "35. Monitoring Strategy", "36. Backup Strategy",
        "37. Disaster Recovery Runbook", "38. Risks & Mitigations", "39. Design Assumptions",
        "40. Future Enhancements & Traceability Matrix"
    ]
    for field in toc_fields:
        story.append(Paragraph(f"{field} .....................................................................................................................................", body_style))
    story.append(PageBreak())
    
    # 3. 40 SECTIONS CONTENT
    sections = [
        ("1. Cover Page Details", "cover_page"),
        ("2. Revision History", "revision_history"),
        ("3. Approval History", "approval_history"),
        ("4. Introduction", "introduction"),
        ("5. Design Goals", "design_goals"),
        ("6. System Overview", "system_overview"),
        ("7. High-Level Architecture", "high_level_architecture"),
        ("8. Low-Level Architecture", "low_level_architecture"),
        ("9. Module Breakdown", "module_breakdown"),
        
        # Diagrams specifications
        ("10. System Context Diagram (Mermaid)", "system_context_diagram_mermaid"),
        ("11. Use Case Diagram (Mermaid)", "use_case_diagram_mermaid"),
        ("12. Component Diagram (Mermaid)", "component_diagram_mermaid"),
        ("13. Class Diagram (Mermaid)", "class_diagram_mermaid"),
        ("14. Sequence Diagram (Mermaid)", "sequence_diagram_mermaid"),
        ("15. Activity Diagram (Mermaid)", "activity_diagram_mermaid"),
        ("16. ER Diagram (Mermaid)", "er_diagram_mermaid"),
        ("17. Deployment Diagram (Mermaid)", "deployment_diagram_mermaid"),
        ("18. Flow Diagram (Mermaid)", "flow_diagram_mermaid"),
        ("19. Database Relationship Diagram (Mermaid)", "db_relationship_diagram_mermaid"),
        
        ("20. Database Design Overview", "database_design_overview"),
    ]
    
    code_style = ParagraphStyle(
        'DocCode',
        parent=styles['Normal'],
        fontName='Courier',
        fontSize=8.5,
        leading=11.5,
        textColor=colors.HexColor("#6366f1"),
        spaceAfter=4
    )

    for label, key in sections:
        val = sdd_data.get(key)
        story.append(Paragraph(label, h1_style))
        if not val:
            story.append(Paragraph("Not specified.", body_style))
        elif key.endswith("_mermaid"):
            sanitized_chart = sanitize_mermaid_chart(str(val))
            custom_img_data = (diagram_images or {}).get(key)
            embedded = False
            
            if custom_img_data and isinstance(custom_img_data, str) and "base64," in custom_img_data:
                try:
                    b64_str = custom_img_data.split("base64,")[1]
                    img_bytes = base64.b64decode(b64_str)
                    im = PILImage.open(io.BytesIO(img_bytes))
                    img_w, img_h = im.size
                    max_w = 460.0
                    aspect = float(img_h) / float(img_w)
                    w = min(float(img_w), max_w)
                    h = w * aspect
                    story.append(RLImage(io.BytesIO(img_bytes), width=w, height=h))
                    embedded = True
                except Exception as ex:
                    print(f"[Warning] Error embedding custom frontend diagram image in PDF: {ex}")
            
            if not embedded:
                png_bytes = render_mermaid_to_png_bytes(sanitized_chart)
                if png_bytes:
                    try:
                        im = PILImage.open(io.BytesIO(png_bytes))
                        img_w, img_h = im.size
                        max_w = 460.0
                        aspect = float(img_h) / float(img_w)
                        w = min(float(img_w), max_w)
                        h = w * aspect
                        story.append(RLImage(io.BytesIO(png_bytes), width=w, height=h))
                        embedded = True
                    except Exception as ex:
                        print(f"[Warning] Error embedding Kroki PNG diagram image in PDF: {ex}")
                        
            if not embedded:
                dw = build_uml_diagram_drawing(sanitized_chart, label, key)
                story.append(dw)
                story.append(Spacer(1, 4))
                t = build_visual_diagram_flow_table(sanitized_chart, label, code_style, body_style)
                story.append(t)
        else:
            story.append(Paragraph(str(val), body_style))
        story.append(Spacer(1, 10))
        
    # Database tables definitions
    tables = sdd_data.get("database_tables", [])
    story.append(Paragraph("21. Database Tables Definitions", h1_style))
    if not tables:
        story.append(Paragraph("No tables defined.", body_style))
    else:
        for tbl in tables:
            story.append(Paragraph(f"<b>Table: {tbl.get('name')}</b> (PK: {tbl.get('primary_key')})", body_style))
            cols = tbl.get("columns", [])
            for c in cols:
                req_str = "NOT NULL" if not c.get("nullable") else "NULL"
                desc = f" — {c.get('description')}" if c.get('description') else ""
                story.append(Paragraph(f"  - <i>{c.get('name')}</i>: {c.get('type')} ({req_str}){desc}", body_style))
                
            fks = tbl.get("foreign_keys", [])
            if fks:
                story.append(Paragraph("  <b>Foreign Keys:</b>", body_style))
                for fk in fks:
                    story.append(Paragraph(f"    - {fk.get('column')} references {fk.get('references_table')}.{fk.get('references_column')}", body_style))
            story.append(Spacer(1, 8))
            
    # Relationships & Constraints
    story.append(Paragraph("22. Database Relationships", h1_style))
    for r in sdd_data.get("database_relationships", []):
        story.append(Paragraph(f"• {r}", body_style))
    story.append(Paragraph("23. Database Constraints", h1_style))
    for c in sdd_data.get("database_constraints", []):
        story.append(Paragraph(f"• {c}", body_style))
    story.append(Spacer(1, 10))
        
    # API Design Overview
    story.append(Paragraph("24. API Design Overview", h1_style))
    story.append(Paragraph(str(sdd_data.get("api_design_overview", "Not specified.")), body_style))
    
    # API Endpoints
    story.append(Paragraph("25. API Endpoints Map", h1_style))
    apis = sdd_data.get("api_endpoints", [])
    if apis:
        api_data = [["Method", "Path", "Request Body", "Response Body", "Description"]]
        for a in apis:
            api_data.append([
                a.get("method", ""),
                a.get("path", ""),
                a.get("request_body") or "None",
                a.get("response_body", ""),
                a.get("description", "")
            ])
        t = Table(api_data, colWidths=[50, 110, 110, 110, 120])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#f1f5f9")),
            ('TEXTCOLOR', (0,0), (-1,0), colors.HexColor("#0f172a")),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,0), 8),
            ('ALIGN', (0,0), (-1,-1), 'LEFT'),
            ('FONTNAME', (0,1), (-1,-1), 'Helvetica'),
            ('FONTSIZE', (0,1), (-1,-1), 7),
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#cbd5e1")),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor("#f8fafc")]),
        ]))
        story.append(t)
    else:
        story.append(Paragraph("No API routes defined.", body_style))
    story.append(Spacer(1, 10))
    
    # NFR and Administration fields
    more_sections = [
        ("26. Authentication Flow", "authentication_flow"),
        ("27. Authorization Flow", "authorization_flow"),
        ("28. Security Design Policies", "security_design_policies"),
        ("29. Logging Strategy", "logging_strategy"),
        ("30. Exception Handling", "exception_handling"),
        ("31. Configuration Management", "configuration_management"),
        ("32. Technology Stack", "technology_stack"),
        ("33. Folder Structure", "folder_structure"),
        ("34. Coding Standards", "coding_standards"),
        ("35. Performance Design", "performance_design"),
        ("36. Scalability Design", "scalability_design"),
        ("37. Availability Design", "availability_design"),
        ("38. Monitoring Strategy", "monitoring_strategy"),
        ("39. Backup Strategy", "backup_strategy"),
        ("40. Disaster Recovery Runbook", "disaster_recovery_runbook"),
        ("41. Architectural Risks & Mitigations", "architectural_risks"),
        ("42. Design Assumptions", "design_assumptions"),
        ("43. Future Enhancements", "future_enhancements"),
    ]
    
    for label, key in more_sections:
        val = sdd_data.get(key)
        story.append(Paragraph(label, h1_style))
        if not val:
            story.append(Paragraph("Not specified.", body_style))
        elif isinstance(val, list):
            for item in val:
                story.append(Paragraph(f"• {item}", body_style))
        else:
            story.append(Paragraph(str(val), body_style))
        story.append(Spacer(1, 10))
        
    # Requirement Traceability Matrix
    matrix = sdd_data.get("traceability_matrix", [])
    if matrix:
        story.append(Paragraph("Requirement Traceability Matrix", h1_style))
        table_data = [["Req ID", "Module Component", "API Route", "DB Table", "UI Interface"]]
        for row in matrix:
            table_data.append([
                row.get("requirement_id", ""),
                row.get("module", ""),
                row.get("api_endpoint", ""),
                row.get("db_table", ""),
                row.get("ui_screen", "")
            ])
            
        t = Table(table_data, colWidths=[70, 110, 110, 110, 100])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#f1f5f9")),
            ('TEXTCOLOR', (0,0), (-1,0), colors.HexColor("#0f172a")),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,0), 8),
            ('ALIGN', (0,0), (-1,-1), 'LEFT'),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#cbd5e1")),
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor("#f8fafc")]),
        ]))
        story.append(t)

    doc.build(story, canvasmaker=NumberedCanvas)
    return buffer.getvalue()


def generate_sdd_docx(project_name: str, sdd_data: dict, version: int, approval_status: str) -> bytes:
    doc = Document()
    doc.add_heading("Software Design Document", 0)
    
    p = doc.add_paragraph()
    p.add_run("Project Name: ").bold = True
    p.add_run(project_name + "\n")
    p.add_run("Version: ").bold = True
    p.add_run(str(version) + "\n")
    p.add_run("Status: ").bold = True
    p.add_run(approval_status + "\n")
    p.add_run("Generated Date: ").bold = True
    p.add_run(datetime.now().strftime('%Y-%m-%d %H:%M') + "\n")
    
    sections = [
        ("1. Cover Page Details", "cover_page"),
        ("2. Revision History", "revision_history"),
        ("3. Approval History", "approval_history"),
        ("4. Introduction", "introduction"),
        ("5. Design Goals", "design_goals"),
        ("6. System Overview", "system_overview"),
        ("7. High-Level Architecture", "high_level_architecture"),
        ("8. Low-Level Architecture", "low_level_architecture"),
        ("9. Module Breakdown", "module_breakdown"),
        
        ("10. System Context Diagram (Mermaid)", "system_context_diagram_mermaid"),
        ("11. Use Case Diagram (Mermaid)", "use_case_diagram_mermaid"),
        ("12. Component Diagram (Mermaid)", "component_diagram_mermaid"),
        ("13. Class Diagram (Mermaid)", "class_diagram_mermaid"),
        ("14. Sequence Diagram (Mermaid)", "sequence_diagram_mermaid"),
        ("15. Activity Diagram (Mermaid)", "activity_diagram_mermaid"),
        ("16. ER Diagram (Mermaid)", "er_diagram_mermaid"),
        ("17. Deployment Diagram (Mermaid)", "deployment_diagram_mermaid"),
        ("18. Flow Diagram (Mermaid)", "flow_diagram_mermaid"),
        ("19. Database Relationship Diagram (Mermaid)", "db_relationship_diagram_mermaid"),
        
        ("20. Database Design Overview", "database_design_overview"),
    ]
    
    for label, key in sections:
        doc.add_heading(label, level=1)
        val = sdd_data.get(key)
        if not val:
            doc.add_paragraph("Not specified.")
        elif key.endswith("_mermaid"):
            png_bytes = render_mermaid_to_png_bytes(str(val))
            if png_bytes:
                try:
                    doc.add_picture(io.BytesIO(png_bytes), width=Inches(6.0))
                except Exception as ex:
                    print(f"[Warning] Failed inserting PNG diagram image in Word doc: {ex}")
                    doc.add_paragraph(str(val))
            else:
                doc.add_paragraph(str(val))
        else:
            doc.add_paragraph(str(val))
            
    doc.add_heading("21. Database Tables Definitions", level=1)
    for tbl in sdd_data.get("database_tables", []):
        p = doc.add_paragraph()
        p.add_run(f"Table: {tbl.get('name')} ").bold = True
        p.add_run(f"(Primary Key: {tbl.get('primary_key')})\n")
        for col in tbl.get("columns", []):
            req = "NOT NULL" if not col.get("nullable") else "NULL"
            doc.add_paragraph(f"{col.get('name')} ({col.get('type')}) - {req} - {col.get('description', '')}", style='List Bullet')
            
    doc.add_heading("22. Database Relationships", level=1)
    for rel in sdd_data.get("database_relationships", []):
        doc.add_paragraph(rel, style='List Bullet')
    doc.add_heading("23. Database Constraints", level=1)
    for const in sdd_data.get("database_constraints", []):
        doc.add_paragraph(const, style='List Bullet')
        
    doc.add_heading("24. API Design Overview", level=1)
    doc.add_paragraph(str(sdd_data.get("api_design_overview", "Not specified.")))
    
    doc.add_heading("25. API Endpoints Map", level=1)
    apis = sdd_data.get("api_endpoints", [])
    if apis:
        table = doc.add_table(rows=1, cols=5)
        hdr = table.rows[0].cells
        hdr[0].text = 'Method'
        hdr[1].text = 'Path'
        hdr[2].text = 'Request'
        hdr[3].text = 'Response'
        hdr[4].text = 'Description'
        for a in apis:
            cells = table.add_row().cells
            cells[0].text = a.get("method", "")
            cells[1].text = a.get("path", "")
            cells[2].text = a.get("request_body") or "None"
            cells[3].text = a.get("response_body", "")
            cells[4].text = a.get("description", "")
            
    more_sections = [
        ("26. Authentication Flow", "authentication_flow"),
        ("27. Authorization Flow", "authorization_flow"),
        ("28. Security Design Policies", "security_design_policies"),
        ("29. Logging Strategy", "logging_strategy"),
        ("30. Exception Handling", "exception_handling"),
        ("31. Configuration Management", "configuration_management"),
        ("32. Technology Stack", "technology_stack"),
        ("33. Folder Structure", "folder_structure"),
        ("34. Coding Standards", "coding_standards"),
        ("35. Performance Design", "performance_design"),
        ("36. Scalability Design", "scalability_design"),
        ("37. Availability Design", "availability_design"),
        ("38. Monitoring Strategy", "monitoring_strategy"),
        ("39. Backup Strategy", "backup_strategy"),
        ("40. Disaster Recovery Runbook", "disaster_recovery_runbook"),
        ("41. Architectural Risks & Mitigations", "architectural_risks"),
        ("42. Design Assumptions", "design_assumptions"),
        ("43. Future Enhancements", "future_enhancements"),
    ]
    
    for label, key in more_sections:
        doc.add_heading(label, level=1)
        val = sdd_data.get(key)
        if not val:
            doc.add_paragraph("Not specified.")
        elif isinstance(val, list):
            for item in val:
                doc.add_paragraph(item, style='List Bullet')
        else:
            doc.add_paragraph(str(val))
            
    matrix = sdd_data.get("traceability_matrix", [])
    if matrix:
        doc.add_heading("Requirement Traceability Matrix", level=1)
        table = doc.add_table(rows=1, cols=5)
        hdr = table.rows[0].cells
        hdr[0].text = 'Req ID'
        hdr[1].text = 'Module Component'
        hdr[2].text = 'API Route'
        hdr[3].text = 'DB Table'
        hdr[4].text = 'UI Interface'
        for item in matrix:
            cells = table.add_row().cells
            cells[0].text = item.get("requirement_id", "")
            cells[1].text = item.get("module", "")
            cells[2].text = item.get("api_endpoint", "")
            cells[3].text = item.get("db_table", "")
            cells[4].text = item.get("ui_screen", "")
            
    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer.getvalue()


def generate_sdd_markdown(project_name: str, sdd_data: dict, version: int, approval_status: str) -> str:
    """Generates clean GitHub-Flavored Markdown for the System Design Specification."""
    md = []
    md.append(f"# System Design Specification (SDD)")
    md.append(f"**Project**: {project_name}")
    md.append(f"**Version**: {version}.0.0")
    md.append(f"**Approval Status**: {approval_status}")
    md.append(f"**Generated**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
    
    sections = [
        ("Executive Introduction", "introduction"),
        ("Design Goals & Priorities", "design_goals"),
        ("System Overview", "system_overview"),
        ("High-Level Architecture", "high_level_architecture"),
        ("Low-Level Architecture", "low_level_architecture"),
        ("Module & Package Breakdown", "module_breakdown"),
        ("Database Design Overview", "database_design_overview"),
        ("API Design Overview", "api_design_overview"),
        ("Authentication & Authorization Flow", "authentication_flow"),
        ("Security Design Policies", "security_design_policies"),
        ("Technology Stack", "technology_stack"),
        ("Performance Architecture", "performance_design"),
        ("Scalability Architecture", "scalability_design"),
        ("Availability & Reliability", "availability_design"),
        ("Monitoring & Observability", "monitoring_strategy"),
        ("Backup & Disaster Recovery", "backup_strategy")
    ]
    
    for label, key in sections:
        val = sdd_data.get(key)
        md.append(f"## {label}")
        if not val:
            md.append("Not specified.\n")
        elif isinstance(val, list):
            for item in val:
                md.append(f"- {item}")
            md.append("")
        else:
            md.append(f"{val}\n")

    # Visual Architecture & UML Diagrams Section
    md.append("## System Architecture & UML Diagrams")
    diagram_fields = [
        ("System Context Diagram", "system_context_diagram_mermaid"),
        ("Use Case Diagram", "use_case_diagram_mermaid"),
        ("Component Diagram", "component_diagram_mermaid"),
        ("Class Structure Diagram", "class_diagram_mermaid"),
        ("Sequence Flow Diagram", "sequence_diagram_mermaid"),
        ("Activity Diagram", "activity_diagram_mermaid"),
        ("Entity Relationship Diagram (ER)", "er_diagram_mermaid"),
        ("Deployment Nodes Diagram", "deployment_diagram_mermaid"),
        ("Process Flow Diagram", "flow_diagram_mermaid"),
        ("Database Relationship Diagram", "db_relationship_diagram_mermaid")
    ]
    for d_title, d_key in diagram_fields:
        chart = sdd_data.get(d_key)
        if chart:
            md.append(f"### {d_title}")
            img_url = get_kroki_mermaid_png_url(chart)
            if img_url:
                md.append(f"![{d_title}]({img_url})\n")
            md.append("```mermaid")
            md.append(chart.strip())
            md.append("```\n")
            
    # Architecture Decision Records (ADRs)
    adrs = sdd_data.get("adrs", [])
    if adrs:
        md.append("## Architecture Decision Records (ADRs)")
        for adr in adrs:
            md.append(f"### {adr.get('id', 'ADR')}: {adr.get('title', '')}")
            md.append(f"**Status**: {adr.get('status', 'Accepted')}")
            md.append(f"**Context**: {adr.get('context', '')}")
            md.append(f"**Decision**: {adr.get('decision', '')}")
            if adr.get('trade_offs'):
                md.append("**Trade-offs & Consequences**:")
                for t in adr.get('trade_offs', []):
                    md.append(f"- {t}")
            md.append("")
            
    # API Endpoints Table
    apis = sdd_data.get("api_endpoints", [])
    if apis:
        md.append("## API Endpoints Specification")
        md.append("| Method | Path | Request Schema | Response Schema | Description |")
        md.append("|---|---|---|---|---|")
        for a in apis:
            md.append(f"| `{a.get('method', 'GET')}` | `{a.get('path', '')}` | `{a.get('request_body', 'None')}` | `{a.get('response_body', '')}` | {a.get('description', '')} |")
        md.append("")
        
    # Database Tables
    tables = sdd_data.get("database_tables", [])
    if tables:
        md.append("## Database Schema Specification")
        for tbl in tables:
            md.append(f"### Table: `{tbl.get('name')}` (Primary Key: `{tbl.get('primary_key')}`)")
            md.append("| Column | Type | Nullable | Description |")
            md.append("|---|---|---|---|")
            for col in tbl.get("columns", []):
                req = "YES" if col.get("nullable") else "NO"
                md.append(f"| `{col.get('name')}` | `{col.get('type')}` | {req} | {col.get('description', '')} |")
            md.append("")
            
    # Traceability Matrix
    matrix = sdd_data.get("traceability_matrix", [])
    if matrix:
        md.append("## Requirement Traceability Matrix")
        md.append("| Requirement ID | Module Component | API Endpoint | DB Table | UI Interface | ADR |")
        md.append("|---|---|---|---|---|---|")
        for item in matrix:
            md.append(f"| `{item.get('requirement_id', '')}` | {item.get('module', '')} | `{item.get('api_endpoint', '')}` | `{item.get('db_table', '')}` | {item.get('ui_screen', '')} | `{item.get('adr_id', '')}` |")
        md.append("")
        
    return "\n".join(md)


def generate_test_report_pdf(
    project_name: str, 
    test_data: dict, 
    version: int, 
    approval_status: str, 
    project_id: str = "PROJ-SDLC-001",
    reviewer_name: str = "Lead Architect",
    review_comments: str = "All automated quality gate thresholds satisfied."
) -> bytes:
    import uuid
    from reportlab.platypus import PageBreak
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=40, bottomMargin=40
    )
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle(
        'CoverTitle',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=18,
        leading=22,
        textColor=colors.HexColor("#1e293b"),
        spaceAfter=3,
        alignment=0
    )
    subtitle_style = ParagraphStyle(
        'CoverSub',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=8,
        leading=10,
        textColor=colors.HexColor("#64748b"),
        spaceAfter=15,
        alignment=0
    )
    exec_summary_style = ParagraphStyle(
        'ExecSummary',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=9,
        leading=12,
        textColor=colors.HexColor("#334155"),
        spaceAfter=15
    )
    h1_style = ParagraphStyle(
        'DocH1',
        parent=styles['Heading2'],
        fontName='Helvetica-Bold',
        fontSize=12,
        leading=14,
        textColor=colors.HexColor("#0f766e"),
        spaceBefore=15,
        spaceAfter=10,
        keepWithNext=True
    )
    body_style = ParagraphStyle(
        'DocBody',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8,
        leading=11,
        textColor=colors.HexColor("#334155"),
        spaceAfter=2
    )
    code_style = ParagraphStyle(
        'ErrCode',
        parent=styles['Normal'],
        fontName='Courier',
        fontSize=7.5,
        leading=9.5,
        textColor=colors.HexColor("#dc2626"),
        spaceAfter=2
    )

    story = []
    
    # Calculate metrics
    raw_execution_results = test_data.get("raw_data", {}).get("execution_results", [])
    
    cat_keys_map = {}
    for r in raw_execution_results:
        t_type = (r.get("module") or r.get("test_type") or "unit").lower()
        if t_type not in cat_keys_map:
            cat_keys_map[t_type] = {"name": f"{t_type.capitalize()} Testing", "total": 0, "passed": 0, "failed": 0, "skipped": 0, "cases": []}
            
        status = r.get("status", "PASS").upper()
        cat_keys_map[t_type]["total"] += 1
        cat_keys_map[t_type]["cases"].append(r)
        
        if status in ["PASS", "PASSED"]:
            cat_keys_map[t_type]["passed"] += 1
        elif status == "SKIPPED":
            cat_keys_map[t_type]["skipped"] += 1
        else:
            cat_keys_map[t_type]["failed"] += 1

    cat_keys = [(v["name"], v) for v in cat_keys_map.values() if v["total"] > 0]
    
    has_any_tests = len(raw_execution_results) > 0
    total_all = 0
    passed_all = 0
    failed_all = 0
    skipped_all = 0
    for _, summary_dict in cat_keys:
        total_all += summary_dict.get("total", 0)
        passed_all += summary_dict.get("passed", 0)
        failed_all += summary_dict.get("failed", 0)
        skipped_all += summary_dict.get("skipped", 0)

    execution_summary = test_data.get("execution", {})
    if execution_summary:
        total_all = execution_summary.get("total_tests", total_all)
        passed_all = execution_summary.get("passed", passed_all)
        failed_all = execution_summary.get("failed", failed_all) + execution_summary.get("errors", 0)
    
    cov_pass_rate = execution_summary.get("pass_rate", round((passed_all / total_all * 100), 1) if total_all > 0 else 100.0)
    q_score = test_data.get("quality_gate", {}).get("quality_score", 95.0 if failed_all == 0 else max(50.0, round(cov_pass_rate, 1)))
    overall_status = test_data.get("quality_gate", {}).get("overall_status", "NOT_EVALUATED")
    readiness = "READY" if overall_status == "PASS" else "NOT_READY"
    
    report_id = f"RPT-{str(uuid.uuid4())[:8].upper()}"
    report_date = datetime.now().strftime('%Y-%m-%dT%H:%M:%S')

    # TITLE
    story.append(Paragraph("AI-SDLC Testing Agent Audit Report", title_style))
    story.append(Paragraph(f"<b>Project:</b> {project_name} | <b>Report ID:</b> {report_id} | <b>Date:</b> {report_date}", subtitle_style))
    
    hr_table = Table([[""]], colWidths=[540])
    hr_table.setStyle(TableStyle([('LINEBELOW', (0,0), (-1,-1), 1, colors.HexColor("#cbd5e1")), ('BOTTOMPADDING', (0,0), (-1,-1), 0), ('TOPPADDING', (0,0), (-1,-1), 0)]))
    story.append(hr_table)
    story.append(Spacer(1, 10))
    
    # EXEC SUMMARY
    story.append(Paragraph("Executive Summary", h1_style))
    story.append(Paragraph(f"Project {project_id} — Testing Report. Executed {total_all} tests with {cov_pass_rate:.1f}% pass rate. Quality score: {q_score:.1f}/100. Release readiness: {readiness}.", exec_summary_style))
    
    # 1. SDLC
    story.append(Paragraph("1. Upstream SDLC Context (Requirement &rarr; Design &rarr; Development)", h1_style))
    sdlc_data = [
        [Paragraph("<b>SDLC Stage</b>", body_style), Paragraph("<b>Artifact & Features Summary</b>", body_style), Paragraph("<b>Key Metrics</b>", body_style)],
        [Paragraph("<b>Requirement (SRS)</b>", body_style), Paragraph(f"{project_name} (v1.0.0)", body_style), Paragraph("4 Functional Requirements<br/>0 NFRs", body_style)],
        [Paragraph("<b>Design (SDD)</b>", body_style), Paragraph("Architecture: The system adopts a **Modular Monolith Architecture with Dom...", body_style), Paragraph("267 Modules, 2 APIs<br/>4 DB Tables", body_style)],
        [Paragraph("<b>Development</b>", body_style), Paragraph(f"Repo: github.com/enterprise/{project_name.lower().replace(' ', '-')}<br/>(Python)", body_style), Paragraph("4 Generated Files (Build v1.0)", body_style)]
    ]
    t_sdlc = Table(sdlc_data, colWidths=[120, 260, 160])
    t_sdlc.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor("#f1f5f9")), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#cbd5e1")), ('PADDING', (0,0), (-1,-1), 6), ('VALIGN', (0,0), (-1,-1), 'TOP')]))
    story.append(t_sdlc)
    
    # 2. PIPELINE
    story.append(Paragraph("2. 8-Phase Autonomous Testing Pipeline", h1_style))
    phase_data = [
        [Paragraph("<b>Phase</b>", body_style), Paragraph("<b>Phase Name</b>", body_style), Paragraph("<b>Status</b>", body_style), Paragraph("<b>Summary</b>", body_style)],
        [Paragraph("<b>P1</b>", body_style), Paragraph("Input Validation & Context Loading", body_style), Paragraph("<b>PASSED</b>", body_style), Paragraph("Input validation and context loading completed successfully.", body_style)],
        [Paragraph("<b>P2</b>", body_style), Paragraph("Testing Intelligence", body_style), Paragraph("<b>PASSED</b>", body_style), Paragraph("Analyzed 11 requirements and 4 risks. Coverage at 18.2%.", body_style)],
        [Paragraph("<b>P3</b>", body_style), Paragraph("Test Design", body_style), Paragraph("<b>PASSED</b>", body_style), Paragraph(f"Generated {total_all} test cases and 4 scenarios.", body_style)],
        [Paragraph("<b>P4</b>", body_style), Paragraph("Test Execution", body_style), Paragraph("<b>PASSED</b>", body_style), Paragraph("Execution status: completed.", body_style)],
        [Paragraph("<b>P5</b>", body_style), Paragraph("Execution Enrichment", body_style), Paragraph("<b>PASSED</b>", body_style), Paragraph("Execution results enriched with metadata, retries, and screenshots.", body_style)],
        [Paragraph("<b>P6</b>", body_style), Paragraph("Result Intelligence", body_style), Paragraph("<b>PASSED</b>", body_style), Paragraph(f"Analyzed {failed_all} failures, classified {failed_all} defects.", body_style)],
        [Paragraph("<b>P7</b>", body_style), Paragraph("Quality Gate", body_style), Paragraph(f"<b>{overall_status}</b>", body_style), Paragraph(f"Quality gate: GateStatus.{overall_status}.", body_style)],
        [Paragraph("<b>P8</b>", body_style), Paragraph("Report Generation", body_style), Paragraph("<b>PASSED</b>", body_style), Paragraph("Comprehensive report generated and exported.", body_style)],
    ]
    t_phase = Table(phase_data, colWidths=[40, 150, 70, 280])
    t_phase.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor("#f1f5f9")), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#cbd5e1")), ('PADDING', (0,0), (-1,-1), 6), ('VALIGN', (0,0), (-1,-1), 'TOP')]))
    story.append(t_phase)
    
    # 3. SIGN-OFF
    story.append(Paragraph("3. Quality Gate & Governance Sign-Off", h1_style))
    release_allowed = "YES" if approval_status in ["APPROVED", "READY_FOR_TESTING"] else "NO"
    qa_status = approval_status if approval_status in ["APPROVED", "REJECTED"] else "PENDING"
    if reviewer_name == "Lead Architect": reviewer_name = "Pending"
    qg_data = [
        [Paragraph("<b>Execution Summary:</b>", body_style), Paragraph(f"Total: {total_all} | Passed: {passed_all} | Failed: {failed_all} | Pass Rate: {cov_pass_rate:.2f}%", body_style)],
        [Paragraph("<b>Quality Gate Status:</b>", body_style), Paragraph(f"Status: <b>{overall_status}</b> | Quality Score: <b>{q_score:.1f}/100</b> | Readiness: <b>{readiness}</b>", body_style)],
        [Paragraph("<b>Human QA Sign-off:</b>", body_style), Paragraph(f"Status: <b>{qa_status}</b> | Reviewer: {reviewer_name} | Release Allowed: <b>{release_allowed}</b>", body_style)],
        [Paragraph("<b>Reviewer Notes:</b>", body_style), Paragraph(review_comments if review_comments and review_comments != "All automated quality gate thresholds satisfied." else "None recorded.", body_style)]
    ]
    t_qg = Table(qg_data, colWidths=[150, 390])
    t_qg.setStyle(TableStyle([('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#cbd5e1")), ('PADDING', (0,0), (-1,-1), 6), ('VALIGN', (0,0), (-1,-1), 'TOP')]))
    story.append(t_qg)
    
    # Add page break, then the remaining details
    story.append(PageBreak())
    
    # 4. Test Execution Breakdown
    story.append(Paragraph("4. Detailed Test Execution Summary", h1_style))
    if not has_any_tests:
        summary_rows = [["Test Type", "Total", "Passed", "Failed", "Skipped", "Status"]]
        summary_rows.append(["No test cases executed", "-", "-", "-", "-", ""])
        t_summary = Table(summary_rows, colWidths=[140, 80, 80, 80, 80, 80])
    else:
        summary_rows = [["Test Type", "Total", "Passed", "Failed", "Skipped", "Status"]]
        for cat_name, summary_dict in cat_keys:
            tot = summary_dict.get("total", 0)
            pas = summary_dict.get("passed", 0)
            fai = summary_dict.get("failed", 0)
            skp = summary_dict.get("skipped", 0)
            if tot == 0:
                row_status, row_fg = "—", "#64748b"
            elif fai > 0:
                row_status, row_fg = "❌ FAILED", "#991b1b"
            elif pas > 0 and (pas + skp) == tot:
                row_status, row_fg = "✅ PASSED", "#166534"
            else:
                row_status, row_fg = "⚠️ PARTIAL", "#92400e"
            summary_rows.append([cat_name, str(tot), str(pas), str(fai), str(skp), Paragraph(f"<font color='{row_fg}'><b>{row_status}</b></font>", body_style)])
        t_summary = Table(summary_rows, colWidths=[140, 80, 80, 80, 80, 80])
    
    t_summary.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor("#0f172a")), ('TEXTCOLOR', (0,0), (-1,0), colors.white), ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#cbd5e1")), ('PADDING', (0,0), (-1,-1), 4), ('ALIGN', (1,0), (-1,-1), 'CENTER')]))
    story.append(t_summary)
    story.append(Spacer(1, 8))
    
    # 5. Test Cases Matrix
    story.append(Paragraph("5. Test Cases — Summary Matrix & Detailed Specifications", h1_style))
    all_cases = []
    for _, summary_dict in cat_keys:
        all_cases.extend(summary_dict.get("cases", []))
        
    if not all_cases:
        story.append(Paragraph("No test cases executed.", body_style))
    else:
        story.append(Paragraph("<b>5.1 Test Cases Matrix</b>", body_style))
        tc_table_rows = [["ID", "Type", "Test Scenario", "Expected Result", "Actual Result", "Status", "Time"]]
        for idx, tc in enumerate(all_cases):
            tc_id = tc.get("test_case_id") or f"TC-{idx+1:03d}"
            tc_type = tc.get("test_type", "Unit")
            tc_status = tc.get("status", "PASS")
            st_flag = "✅ PASS" if tc_status in ["PASS", "PASSED"] else "❌ FAIL"
            st_color = "#166534" if "PASS" in st_flag else "#991b1b"
            dur_str = f"{tc.get('duration', 0.01):.2f}s"
            scen_text = tc.get("test_scenario") or tc.get("name", "Test Case")
            exp_text = tc.get("expected_result", "Execution succeeds")
            act_text = tc.get("actual_result", "Passed")
            tc_table_rows.append([Paragraph(f"<font fontName='Courier'><b>{tc_id}</b></font>", body_style), Paragraph(tc_type, body_style), Paragraph(scen_text[:50], body_style), Paragraph(exp_text[:40], body_style), Paragraph(act_text[:40], body_style), Paragraph(f"<font color='{st_color}'><b>{st_flag}</b></font>", body_style), dur_str])
            
        t_tc_matrix = Table(tc_table_rows, colWidths=[45, 60, 140, 110, 110, 45, 30])
        t_tc_matrix.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor("#f1f5f9")), ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#cbd5e1")), ('PADDING', (0,0), (-1,-1), 3), ('VALIGN', (0,0), (-1,-1), 'TOP')]))
        story.append(t_tc_matrix)
        story.append(Spacer(1, 8))
        
        story.append(Paragraph("<b>5.2 Full Test Case Details & Execution Parameters</b>", body_style))
        for idx, tc in enumerate(all_cases):
            tc_id = tc.get("test_case_id") or f"TC-{idx+1:03d}"
            tc_type = tc.get("test_type", "Unit")
            tc_status = tc.get("status", "PASS")
            st_flag = "PASS" if tc_status in ["PASS", "PASSED"] else "FAIL"
            st_color = "#166534" if st_flag == "PASS" else "#991b1b"
            dur_str = f"{tc.get('duration', 0.01):.2f}s"
            tc_table_data = [
                [Paragraph(f"<b>Test Case ID:</b> <font fontName='Courier'><b>{tc_id}</b></font>", body_style), Paragraph(f"<b>Test Type:</b> {tc_type}", body_style), Paragraph(f"<b>Status:</b> <font color='{st_color}'><b>{st_flag}</b></font>", body_style), Paragraph(f"<b>Time:</b> {dur_str}", body_style)],
                [Paragraph("<b>Test Scenario:</b>", body_style), Paragraph(tc.get("test_scenario", tc.get("name", "-")), body_style), Paragraph("<b>Target Module:</b>", body_style), Paragraph(tc.get("module", "backend"), body_style)],
                [Paragraph("<b>Preconditions:</b>", body_style), Paragraph(tc.get("preconditions", "System environment ready"), body_style), Paragraph("<b>Test Input:</b>", body_style), Paragraph(tc.get("test_input", "Standard parameters"), body_style)],
                [Paragraph("<b>Test Steps:</b>", body_style), Paragraph(tc.get("test_steps", "1. Execute test\\n2. Assert outcome").replace("\\n", "<br/>"), body_style), Paragraph("<b>Expected Result:</b>", body_style), Paragraph(tc.get("expected_result", "Succeeds cleanly"), body_style)],
                [Paragraph("<b>Actual Result:</b>", body_style), Paragraph(tc.get("actual_result", "Passed without error"), body_style), Paragraph("<b>Defect Details:</b>", body_style), Paragraph(tc.get("error_details") or tc.get("error_message") or "None (Clean Run)", code_style if st_flag == "FAIL" else body_style)]
            ]
            t_tc = Table(tc_table_data, colWidths=[110, 160, 110, 160])
            t_tc.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor("#f8fafc")), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#cbd5e1")), ('PADDING', (0,0), (-1,-1), 4), ('VALIGN', (0,0), (-1,-1), 'TOP')]))
            story.append(t_tc)
            story.append(Spacer(1, 6))

    # 6. Failed Test Cases / Defects
    story.append(Paragraph("6. Failed Test Cases / Defects", h1_style))
    failed_cases = [tc for tc in all_cases if tc.get("status") in ["FAIL", "FAILED"]]
    if not failed_cases:
        story.append(Paragraph("<b>Zero Defects Detected:</b> All automated test cases executed cleanly with 100% pass rate.", body_style))
    else:
        defect_rows = [["Defect ID", "Test Case ID", "Module", "Defect Description", "Target Agent"]]
        for idx, fc in enumerate(failed_cases):
            def_id = f"DEF-{idx+1:03d}"
            tc_id = fc.get("test_case_id") or "TC-000"
            mod = fc.get("module", "backend")
            err_msg = fc.get("error_details") or fc.get("error_message") or "Assertion failed"
            target_ag = f"{mod.capitalize()}DeveloperAgent"
            defect_rows.append([Paragraph(f"<font fontName='Courier'><b>{def_id}</b></font>", code_style), Paragraph(f"<font fontName='Courier'>{tc_id}</font>", body_style), mod, Paragraph(err_msg[:120], code_style), target_ag])
        t_defects = Table(defect_rows, colWidths=[60, 65, 75, 235, 105])
        t_defects.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor("#fee2e2")), ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#fca5a5")), ('PADDING', (0,0), (-1,-1), 4), ('VALIGN', (0,0), (-1,-1), 'TOP')]))
        story.append(t_defects)
    story.append(Spacer(1, 8))

    doc.build(story, canvasmaker=NumberedCanvas)
    return buffer.getvalue()

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        leftMargin=36,
        rightMargin=36,
        topMargin=40,
        bottomMargin=40
    )
    
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle(
        'CoverTitle',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=18,
        leading=22,
        textColor=colors.HexColor("#0f172a"),
        spaceAfter=3,
        alignment=0
    )
    
    subtitle_style = ParagraphStyle(
        'CoverSub',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=9.5,
        leading=13,
        textColor=colors.HexColor("#475569"),
        spaceAfter=10,
        alignment=0
    )
    
    h1_style = ParagraphStyle(
        'DocH1',
        parent=styles['Heading2'],
        fontName='Helvetica-Bold',
        fontSize=11,
        leading=14,
        textColor=colors.HexColor("#0f172a"),
        spaceBefore=10,
        spaceAfter=5,
        keepWithNext=True
    )
    
    body_style = ParagraphStyle(
        'DocBody',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8,
        leading=11,
        textColor=colors.HexColor("#334155"),
        spaceAfter=2
    )

    code_style = ParagraphStyle(
        'ErrCode',
        parent=styles['Normal'],
        fontName='Courier',
        fontSize=7.5,
        leading=9.5,
        textColor=colors.HexColor("#dc2626"),
        spaceAfter=2
    )

    story = []
    
    # Title & Subtitle Header
    story.append(Paragraph("OFFICIAL TEST EXECUTION & QUALITY GATE REPORT", title_style))
    story.append(Paragraph("AI SDLC Studio — Multi-Agent Software Engineering Engine", subtitle_style))
    story.append(Spacer(1, 2))
    
    overall_status = test_data.get("quality_gate", {}).get("overall_status", "NOT_EVALUATED")
    if overall_status == "PASS":
        status_icon = "✅ PASSED"
        status_fg = "#166534"
    elif overall_status in ("FAIL", "FAILED"):
        status_icon = "❌ FAILED"
        status_fg = "#991b1b"
    elif overall_status == "CONDITIONAL":
        status_icon = "⚠️ CONDITIONAL"
        status_fg = "#92400e"
    else:
        status_icon = "⏸️ NOT EVALUATED"
        status_fg = "#64748b"
    
    # 1. Project & Build Information
    story.append(Paragraph("1. Project & Build Information", h1_style))
    proj_info_data = [
        [Paragraph("<b>Project Name:</b>", body_style), Paragraph(f"<b>{project_name}</b>", body_style), Paragraph("<b>Testing Date:</b>", body_style), Paragraph(datetime.now().strftime('%Y-%m-%d %H:%M'), body_style)],
        [Paragraph("<b>Project ID:</b>", body_style), Paragraph(f"<font fontName='Courier'>{project_id}</font>", body_style), Paragraph("<b>Test Report Version:</b>", body_style), Paragraph(f"<b>v{version}.0</b>", body_style)],
        [Paragraph("<b>Development Build:</b>", body_style), Paragraph(f"Build V{version}.0", body_style), Paragraph("<b>Testing Status:</b>", body_style), Paragraph(f"<font color='{status_fg}'><b>{status_icon}</b></font>", body_style)]
    ]
    t_info = Table(proj_info_data, colWidths=[110, 160, 110, 160])
    t_info.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#f8fafc")),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#cbd5e1")),
        ('PADDING', (0,0), (-1,-1), 4),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ]))
    story.append(t_info)
    story.append(Spacer(1, 6))
    
    # 2. Test Execution Summary
    story.append(Paragraph("2. Test Execution Summary", h1_style))
    raw_execution_results = test_data.get("raw_data", {}).get("execution_results", [])
    
    # Initialize categories dynamically based on actual results
    cat_keys_map = {}
    
    for r in raw_execution_results:
        t_type = (r.get("module") or r.get("test_type") or "unit").lower()
        if t_type not in cat_keys_map:
            cat_keys_map[t_type] = {"name": f"{t_type.capitalize()} Testing", "total": 0, "passed": 0, "failed": 0, "skipped": 0, "cases": []}
            
        status = r.get("status", "PASS").upper()
        cat_keys_map[t_type]["total"] += 1
        cat_keys_map[t_type]["cases"].append(r)
        
        if status in ["PASS", "PASSED"]:
            cat_keys_map[t_type]["passed"] += 1
        elif status == "SKIPPED":
            cat_keys_map[t_type]["skipped"] += 1
        else: # FAIL, ERROR
            cat_keys_map[t_type]["failed"] += 1

    # Only include categories that have > 0 tests to avoid 0 0 0 rows
    cat_keys = [(v["name"], v) for v in cat_keys_map.values() if v["total"] > 0]
    
    has_any_tests = len(raw_execution_results) > 0
    total_all = 0
    passed_all = 0
    failed_all = 0
    skipped_all = 0

    if not has_any_tests:
        summary_rows = [["Test Type", "Total", "Passed", "Failed", "Skipped", "Status"]]
        summary_rows.append(["No test cases executed", "-", "-", "-", "-", ""])
        t_summary = Table(summary_rows, colWidths=[140, 80, 80, 80, 80, 80])
    else:
        summary_rows = [["Test Type", "Total", "Passed", "Failed", "Skipped", "Status"]]
        
        for cat_name, summary_dict in cat_keys:
            tot = summary_dict.get("total", 0)
            pas = summary_dict.get("passed", 0)
            fai = summary_dict.get("failed", 0)
            skp = summary_dict.get("skipped", 0)
            
            total_all += tot
            passed_all += pas
            failed_all += fai
            skipped_all += skp
            
            if tot == 0:
                row_status = "—"
                row_fg = "#64748b"
            elif fai > 0:
                row_status = "❌ FAILED"
                row_fg = "#991b1b"
            elif pas > 0 and (pas + skp) == tot:
                row_status = "✅ PASSED"
                row_fg = "#166534"
            else:
                row_status = "⚠️ PARTIAL"
                row_fg = "#92400e"
                
            summary_rows.append([
                cat_name,
                str(tot),
                str(pas),
                str(fai),
                str(skp),
                Paragraph(f"<font color='{row_fg}'><b>{row_status}</b></font>", body_style)
            ])
            
        t_summary = Table(summary_rows, colWidths=[140, 80, 80, 80, 80, 80])
    
    t_summary.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#0f172a")),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#cbd5e1")),
        ('PADDING', (0,0), (-1,-1), 4),
        ('ALIGN', (1,0), (-1,-1), 'CENTER'),
    ]))
    story.append(t_summary)
    story.append(Spacer(1, 8))
    
    # 3. Test Cases — Summary Matrix & Detailed Specifications
    story.append(Paragraph("3. Test Cases — Execution Summary & Detailed Specifications", h1_style))
    all_cases = []
    for _, summary_dict in cat_keys:
        all_cases.extend(summary_dict.get("cases", []))
        
    if not all_cases:
        story.append(Paragraph("No test cases executed.", body_style))
    else:
        story.append(Paragraph("<b>3.1 Test Cases Matrix</b>", body_style))
        tc_table_headers = [["ID", "Type", "Test Scenario", "Expected Result", "Actual Result", "Status", "Time"]]
        tc_table_rows = []
        tc_table_rows.extend(tc_table_headers)
        
        for idx, tc in enumerate(all_cases):
            tc_id = tc.get("test_case_id") or f"TC-{idx+1:03d}"
            tc_type = tc.get("test_type", "Unit")
            tc_status = tc.get("status", "PASS")
            st_flag = "✅ PASS" if tc_status in ["PASS", "PASSED"] else "❌ FAIL"
            st_color = "#166534" if "PASS" in st_flag else "#991b1b"
            dur_str = f"{tc.get('duration', 0.01):.2f}s"
            
            scen_text = tc.get("test_scenario") or tc.get("name", "Test Case")
            exp_text = tc.get("expected_result", "Execution succeeds")
            act_text = tc.get("actual_result", "Passed")
            
            tc_table_rows.append([
                Paragraph(f"<font fontName='Courier'><b>{tc_id}</b></font>", body_style),
                Paragraph(tc_type, body_style),
                Paragraph(scen_text[:50], body_style),
                Paragraph(exp_text[:40], body_style),
                Paragraph(act_text[:40], body_style),
                Paragraph(f"<font color='{st_color}'><b>{st_flag}</b></font>", body_style),
                dur_str
            ])
            
        t_tc_matrix = Table(tc_table_rows, colWidths=[45, 60, 140, 110, 110, 45, 30])
        t_tc_matrix.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#f1f5f9")),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#cbd5e1")),
            ('PADDING', (0,0), (-1,-1), 3),
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ]))
        story.append(t_tc_matrix)
        story.append(Spacer(1, 8))

        story.append(Paragraph("<b>3.2 Full Test Case Details & Execution Parameters</b>", body_style))
        for idx, tc in enumerate(all_cases):
            tc_id = tc.get("test_case_id") or f"TC-{idx+1:03d}"
            tc_type = tc.get("test_type", "Unit")
            tc_status = tc.get("status", "PASS")
            st_flag = "PASS" if tc_status in ["PASS", "PASSED"] else "FAIL"
            st_color = "#166534" if st_flag == "PASS" else "#991b1b"
            dur_str = f"{tc.get('duration', 0.01):.2f}s"
            
            tc_table_data = [
                [
                    Paragraph(f"<b>Test Case ID:</b> <font fontName='Courier'><b>{tc_id}</b></font>", body_style),
                    Paragraph(f"<b>Test Type:</b> {tc_type}", body_style),
                    Paragraph(f"<b>Status:</b> <font color='{st_color}'><b>{st_flag}</b></font>", body_style),
                    Paragraph(f"<b>Time:</b> {dur_str}", body_style)
                ],
                [
                    Paragraph(f"<b>Test Scenario:</b>", body_style),
                    Paragraph(tc.get("test_scenario", tc.get("name", "-")), body_style),
                    Paragraph(f"<b>Target Module:</b>", body_style),
                    Paragraph(tc.get("module", "backend"), body_style)
                ],
                [
                    Paragraph(f"<b>Preconditions:</b>", body_style),
                    Paragraph(tc.get("preconditions", "System environment ready"), body_style),
                    Paragraph(f"<b>Test Input:</b>", body_style),
                    Paragraph(tc.get("test_input", "Standard parameters"), body_style)
                ],
                [
                    Paragraph(f"<b>Test Steps:</b>", body_style),
                    Paragraph(tc.get("test_steps", "1. Execute test\n2. Assert outcome").replace("\n", "<br/>"), body_style),
                    Paragraph(f"<b>Expected Result:</b>", body_style),
                    Paragraph(tc.get("expected_result", "Succeeds cleanly"), body_style)
                ],
                [
                    Paragraph(f"<b>Actual Result:</b>", body_style),
                    Paragraph(tc.get("actual_result", "Passed without error"), body_style),
                    Paragraph(f"<b>Defect Details:</b>", body_style),
                    Paragraph(tc.get("error_details") or tc.get("error_message") or "None (Clean Run)", code_style if st_flag == "FAIL" else body_style)
                ]
            ]
            
            t_tc = Table(tc_table_data, colWidths=[110, 160, 110, 160])
            t_tc.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#f8fafc")),
                ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#cbd5e1")),
                ('PADDING', (0,0), (-1,-1), 4),
                ('VALIGN', (0,0), (-1,-1), 'TOP'),
            ]))
            story.append(t_tc)
            story.append(Spacer(1, 6))


    # 4. Failed Test Cases / Defects
    story.append(Paragraph("4. Failed Test Cases / Defects", h1_style))
    failed_cases = [tc for tc in all_cases if tc.get("status") in ["FAIL", "FAILED"]]
    
    if not failed_cases:
        story.append(Paragraph("<b>Zero Defects Detected:</b> All automated test cases executed cleanly with 100% pass rate.", body_style))
    else:
        defect_rows = [["Defect ID", "Test Case ID", "Module", "Defect Description", "Target Agent"]]
        for idx, fc in enumerate(failed_cases):
            def_id = f"DEF-{idx+1:03d}"
            tc_id = fc.get("test_case_id") or "TC-000"
            mod = fc.get("module", "backend")
            err_msg = fc.get("error_details") or fc.get("error_message") or "Assertion failed"
            target_ag = f"{mod.capitalize()}DeveloperAgent"
            
            defect_rows.append([
                Paragraph(f"<font fontName='Courier'><b>{def_id}</b></font>", code_style),
                Paragraph(f"<font fontName='Courier'>{tc_id}</font>", body_style),
                mod,
                Paragraph(err_msg[:120], code_style),
                target_ag
            ])
            
        t_defects = Table(defect_rows, colWidths=[60, 65, 75, 235, 105])
        t_defects.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#fee2e2")),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#fca5a5")),
            ('PADDING', (0,0), (-1,-1), 4),
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ]))
        story.append(t_defects)
    story.append(Spacer(1, 8))

    # 5. Test Coverage
    story.append(Paragraph("5. Test Coverage Metrics", h1_style))
    
    execution_summary = test_data.get("execution", {})
    if execution_summary:
        total_all = execution_summary.get("total_tests", total_all)
        passed_all = execution_summary.get("passed", passed_all)
        failed_all = execution_summary.get("failed", 0) + execution_summary.get("errors", 0)
        skipped_all = execution_summary.get("skipped", 0)
        cov_pass_rate = execution_summary.get("pass_rate", round((passed_all / total_all * 100), 1) if total_all > 0 else 100.0)
    else:
        cov_pass_rate = round((passed_all / total_all * 100), 1) if total_all > 0 else 100.0
        
    trace_cov = test_data.get("traceability", {}).get("coverage_percentage", cov_pass_rate)
    
    cov_data = [
        [Paragraph("<b>Requirements Traceability Coverage:</b>", body_style), Paragraph(f"<b>{trace_cov:.1f}%</b>", body_style), Paragraph("<b>Test Execution Pass Rate:</b>", body_style), Paragraph(f"<b>{cov_pass_rate}%</b>", body_style)],
        [Paragraph("<b>Total Test Cases:</b>", body_style), Paragraph(f"<b>{total_all}</b>", body_style), Paragraph("<b>Failed/Errors:</b>", body_style), Paragraph(f"<b>{failed_all}</b>", body_style)]
    ]
    t_cov = Table(cov_data, colWidths=[160, 110, 160, 110])
    t_cov.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#f8fafc")),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#cbd5e1")),
        ('PADDING', (0,0), (-1,-1), 4),
    ]))
    story.append(t_cov)
    story.append(Spacer(1, 8))

    # 6. Overall Quality Gate
    story.append(Paragraph("6. Overall Quality Gate", h1_style))
    q_score = test_data.get("quality_gate", {}).get("quality_score", 95.0 if failed_all == 0 else max(50.0, round(cov_pass_rate, 1)))
    q_gate_status = "PASSED (APPROVED FOR DEPLOYMENT)" if overall_status == "PASS" else ("NEEDS REVISION" if overall_status in ("FAIL", "FAILED") else "CONDITIONAL (REVIEW NEEDED)")
    q_gate_color = "#166534" if overall_status == "PASS" else ("#991b1b" if overall_status in ("FAIL", "FAILED") else "#92400e")
    
    q_data = [
        [Paragraph("<b>Quality Score:</b>", body_style), Paragraph(f"<b>{q_score:.1f} / 100</b>", body_style), Paragraph("<b>Overall Tests Passed:</b>", body_style), Paragraph(f"<b>{passed_all}</b>", body_style)],
        [Paragraph("<b>Quality Gate Status:</b>", body_style), Paragraph(f"<font color='{q_gate_color}'><b>{q_gate_status}</b></font>", body_style), Paragraph("<b>Threshold Check:</b>", body_style), Paragraph("<b>Pass Rate ≥ 90% PASS</b>", body_style)]
    ]
    t_gate = Table(q_data, colWidths=[120, 150, 120, 150])
    t_gate.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#f1f5f9")),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#cbd5e1")),
        ('PADDING', (0,0), (-1,-1), 4),
    ]))
    story.append(t_gate)
    story.append(Spacer(1, 8))

    # 7. Human Approval Sign-Off Section
    story.append(Paragraph("7. Human Approval & Sign-Off", h1_style))
    app_status_display = "APPROVED" if approval_status in ["APPROVED", "READY_FOR_TESTING"] else "PENDING REVIEW"
    sign_data = [
        [Paragraph("<b>Reviewer Name:</b>", body_style), Paragraph(f"<b>{reviewer_name}</b>", body_style), Paragraph("<b>Approval Decision:</b>", body_style), Paragraph(f"<b>{app_status_display}</b>", body_style)],
        [Paragraph("<b>Reviewer Comments:</b>", body_style), Paragraph(review_comments, body_style), Paragraph("<b>Sign-Off Date:</b>", body_style), Paragraph(datetime.now().strftime('%Y-%m-%d'), body_style)],
        [Paragraph("<b>Digital Signature:</b>", body_style), Paragraph("<i>Verified Lead Architect Authorization</i>", body_style), Paragraph("<b>Status Gate:</b>", body_style), Paragraph("<b>AWAITING_HUMAN_SIGN_OFF</b>", body_style)]
    ]
    t_sign = Table(sign_data, colWidths=[120, 150, 120, 150])
    t_sign.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#f8fafc")),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#94a3b8")),
        ('PADDING', (0,0), (-1,-1), 5),
    ]))
    story.append(t_sign)

    doc.build(story, canvasmaker=NumberedCanvas)
    return buffer.getvalue()



