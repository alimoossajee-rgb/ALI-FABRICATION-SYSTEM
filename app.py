
from __future__ import annotations

import base64
from io import BytesIO
import json
from pathlib import Path
import re
from datetime import datetime

import pandas as pd
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas

from engine import (
    TypologyCatalog,
    build_summary,
    default_aluminium_offcuts,
    export_project_workbook,
    expand_window_rows,
    load_default_glass_offcuts,
    load_glass_specs,
    optimise_aluminium,
    optimise_glass,
)

BASE_DIR = Path(__file__).resolve().parent
PREVIEW_DIR = BASE_DIR / "typology_previews"
DEFAULT_LOGO = BASE_DIR / "ali_fab_logo.png"
PROJECTS_DIR = BASE_DIR / "projects_data"
PROJECTS_DIR.mkdir(exist_ok=True)
ALL_INPUT_FIELDS = [
    "OVERALL WIDTH",
    "OVERALL HEIGHT",
    "VENT WIDTH",
    "BOTTOM FIXED HEIGHT",
    "BOTTOM CLEARANCE REQUIRED",
    "MAIN VENT WIDTH",
]
FIELD_ORDER = {
    "OVERALL WIDTH": 1,
    "OVERALL HEIGHT": 2,
    "VENT WIDTH": 3,
    "MAIN VENT WIDTH": 4,
    "BOTTOM FIXED HEIGHT": 5,
    "BOTTOM CLEARANCE REQUIRED": 6,
}
FIELD_HELP = {
    "OVERALL WIDTH": "Full system width.",
    "OVERALL HEIGHT": "Full system height.",
    "VENT WIDTH": "Opening vent width for this typology.",
    "MAIN VENT WIDTH": "Main active leaf / vent width where required.",
    "BOTTOM FIXED HEIGHT": "Bottom fixed light height for the selected orientation.",
    "BOTTOM CLEARANCE REQUIRED": "Required clearance below the door leaf.",
}

st.set_page_config(page_title="Ali Fabrication System", page_icon="🪟", layout="wide")


@st.cache_resource
def get_catalog():
    return TypologyCatalog()


@st.cache_data
def get_glass_specs():
    return load_glass_specs()


@st.cache_data
def get_default_glass_offcuts():
    return load_default_glass_offcuts()


def safe_name(s: str) -> str:
    return re.sub(r"[^A-Za-z0-9._-]+", "_", str(s))


def mm_to_m(mm_value: float) -> float:
    return round(float(mm_value or 0) / 1000.0, 2)


def file_as_base64(path: Path) -> str:
    if not path.exists():
        return ""
    return base64.b64encode(path.read_bytes()).decode("utf-8")


def uploaded_or_default_logo(uploaded_file) -> str:
    if uploaded_file:
        return base64.b64encode(uploaded_file.read()).decode("utf-8")
    return file_as_base64(DEFAULT_LOGO)


def preview_base64(variant_key: str) -> str:
    base = safe_name(variant_key)
    for ext in [".png", ".jpg", ".jpeg"]:
        p = PREVIEW_DIR / f"{base}{ext}"
        if p.exists():
            return base64.b64encode(p.read_bytes()).decode("utf-8")
    return ""


def system_code_from_label(label: str) -> str:
    label = str(label or "")
    code = label.split("·")[0].strip()
    if " - " in code:
        return code.split(" - ")[0].strip()
    return code


def variant_short_name(label: str) -> str:
    text = str(label or "")
    if "·" in text:
        return text.split("·", 1)[1].strip()
    return text


def slugify_project_name(name: str) -> str:
    slug = re.sub(r"[^A-Za-z0-9_-]+", "_", str(name).strip()).strip("_")
    return slug or "project"


def project_file(name: str) -> Path:
    return PROJECTS_DIR / f"{slugify_project_name(name)}.json"


def list_saved_projects() -> list[str]:
    names = []
    for p in sorted(PROJECTS_DIR.glob("*.json")):
        try:
            payload = json.loads(p.read_text(encoding="utf-8"))
            names.append(payload.get("project_name", p.stem))
        except Exception:
            names.append(p.stem)
    return names


def load_project_data(name: str):
    p = project_file(name)
    if not p.exists():
        return None
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return None


def save_project_data(name: str, payload: dict):
    p = project_file(name)
    payload = dict(payload)
    payload["project_name"] = name
    payload["saved_at"] = datetime.utcnow().isoformat() + "Z"
    p.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def delete_project_data(name: str):
    p = project_file(name)
    if p.exists():
        p.unlink()


def apply_project_payload(payload: dict, default_variant: str, default_glass_spec: str, variant_map: dict[str, str]):
    windows = payload.get("windows") or [blank_window(default_variant, default_glass_spec, variant_map, 1, "W1")]
    st.session_state.windows = windows
    st.session_state.next_window_id = int(payload.get("next_window_id", max([w.get("id", 0) for w in windows] + [0]) + 1))
    st.session_state.al_offcuts = payload.get("al_offcuts", default_aluminium_offcuts())
    st.session_state.glass_offcuts = payload.get("glass_offcuts", get_default_glass_offcuts()[:60])
    st.session_state.project_name_value = payload.get("project_name", "Ali Fabrication Project")
    st.session_state.client_name_value = payload.get("client_name", "")
    st.session_state.finish_value = payload.get("finish", "Powder Coated")
    st.session_state.stock_length_mm_value = float(payload.get("stock_length_mm", 6400.0))
    st.session_state.glass_sheet_width_mm_value = float(payload.get("glass_sheet_width_mm", 3660.0))
    st.session_state.glass_sheet_height_mm_value = float(payload.get("glass_sheet_height_mm", 2440.0))
    st.session_state.kerf_mm_value = float(payload.get("kerf_mm", 3.0))
    st.session_state.default_row_glass_value = payload.get("default_row_glass", default_glass_spec)


def project_payload_from_state(project_name: str, client_name: str, finish: str, stock_length_mm: float, glass_sheet_width_mm: float, glass_sheet_height_mm: float, kerf_mm: float, default_row_glass: str) -> dict:
    return {
        "project_name": project_name,
        "client_name": client_name,
        "finish": finish,
        "stock_length_mm": stock_length_mm,
        "glass_sheet_width_mm": glass_sheet_width_mm,
        "glass_sheet_height_mm": glass_sheet_height_mm,
        "kerf_mm": kerf_mm,
        "default_row_glass": default_row_glass,
        "windows": st.session_state.windows,
        "next_window_id": st.session_state.next_window_id,
        "al_offcuts": st.session_state.al_offcuts,
        "glass_offcuts": st.session_state.glass_offcuts,
    }


def inject_brand_css(primary: str, accent: str):
    st.markdown(
        f"""
        <style>
        .stApp {{
            background:
                radial-gradient(circle at top right, rgba(22,163,74,0.07), transparent 28%),
                linear-gradient(180deg, #f8fafc 0%, #eef2ff 100%);
        }}
        .brand-hero {{
            background: linear-gradient(135deg, {primary} 0%, #0b2539 50%, {accent} 100%);
            color: white;
            border-radius: 28px;
            padding: 26px 30px;
            box-shadow: 0 18px 44px rgba(15, 23, 42, 0.16);
            margin-bottom: 1rem;
            position: relative;
            overflow: hidden;
        }}
        .brand-hero:before {{
            content: "";
            position: absolute;
            right: -60px;
            top: -60px;
            width: 220px;
            height: 220px;
            background: rgba(255,255,255,0.08);
            border-radius: 999px;
        }}
        .hero-wrap {{
            display: flex;
            align-items: center;
            gap: 20px;
            position: relative;
            z-index: 2;
        }}
        .hero-logo {{
            width: 94px;
            height: 94px;
            object-fit: contain;
            border-radius: 18px;
            background: rgba(255,255,255,0.08);
            padding: 10px;
        }}
        .hero-title {{
            margin: 0;
            font-size: 2rem;
            line-height: 1.05;
            font-weight: 800;
        }}
        .hero-sub {{
            margin-top: 0.45rem;
            opacity: 0.95;
            font-size: 1rem;
        }}
        .metric-card {{
            background: white;
            border-radius: 18px;
            padding: 16px 18px;
            border: 1px solid rgba(148,163,184,0.15);
            box-shadow: 0 8px 24px rgba(15, 23, 42, 0.06);
        }}
        .metric-label {{
            color: #475569;
            font-size: 0.92rem;
            margin-bottom: 0.35rem;
        }}
        .metric-value {{
            color: #0f172a;
            font-size: 1.7rem;
            font-weight: 700;
        }}
        .soft-card {{
            background: rgba(255,255,255,0.9);
            border: 1px solid rgba(148,163,184,0.18);
            border-radius: 22px;
            padding: 18px 18px 10px 18px;
            box-shadow: 0 10px 30px rgba(15, 23, 42, 0.06);
            backdrop-filter: blur(6px);
            margin-bottom: 1rem;
        }}
        .section-title {{
            font-size: 1.14rem;
            font-weight: 800;
            color: #0f172a;
            margin-bottom: 0.65rem;
        }}
        .preview-shell {{
            background: linear-gradient(180deg, rgba(248,250,252,0.95), rgba(255,255,255,0.98));
            border: 1px solid rgba(148,163,184,0.18);
            border-radius: 18px;
            padding: 14px;
            text-align: center;
        }}
        .preview-code {{
            display: inline-block;
            background: rgba(15,76,129,0.1);
            color: {primary};
            border: 1px solid rgba(15,76,129,0.12);
            border-radius: 999px;
            padding: 5px 10px;
            font-size: 0.82rem;
            font-weight: 700;
            margin-bottom: 10px;
        }}
        .preview-variant {{
            color: #334155;
            font-size: 0.88rem;
            margin-bottom: 8px;
            min-height: 38px;
        }}
        .preview-img {{
            width: 100%;
            max-height: 230px;
            object-fit: contain;
            background: white;
            border-radius: 16px;
            padding: 10px;
        }}
        .mini-tag {{
            display:inline-block;
            background:#eff6ff;
            color:#1d4ed8;
            padding:4px 8px;
            border-radius:999px;
            margin: 0 6px 6px 0;
            font-size:0.78rem;
            font-weight:600;
        }}
        div[data-testid="stExpander"] {{
            border: 1px solid rgba(148,163,184,0.18);
            border-radius: 18px;
            background: rgba(255,255,255,0.88);
        }}
        .stButton>button {{
            border-radius: 12px;
            border: none;
            background: {primary};
            color: white;
            font-weight: 700;
        }}
        .stDownloadButton>button {{
            border-radius: 12px;
            border: none;
            background: {accent};
            color: white;
            font-weight: 700;
        }}
        @media (max-width: 900px) {{
            .brand-hero {{
                padding: 18px 16px;
                border-radius: 20px;
            }}
            .hero-wrap {{
                gap: 12px;
                align-items: flex-start;
            }}
            .hero-logo {{
                width: 62px;
                height: 62px;
                padding: 6px;
            }}
            .hero-title {{
                font-size: 1.35rem;
            }}
            .hero-sub {{
                font-size: 0.92rem;
            }}
            .soft-card {{
                padding: 14px 12px 8px 12px;
                border-radius: 16px;
            }}
            .preview-img {{
                max-height: 160px;
            }}
            .metric-value {{
                font-size: 1.3rem;
            }}
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_metric(label: str, value: str):
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-label">{label}</div>
            <div class="metric-value">{value}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_aluminium_bar_layouts(aluminium: dict):
    bars = aluminium.get("bars", []) or []
    if not bars:
        st.info("No new aluminium bars are required.")
        return

    st.markdown('<div class="section-title">Visual aluminium cutting layouts</div>', unsafe_allow_html=True)
    for bar in bars:
        stock = float(bar.get("stock_length_mm", 0) or 0)
        used = float(bar.get("used_mm", 0) or 0)
        waste = float(bar.get("waste_mm", 0) or 0)
        cuts = bar.get("cuts", []) or []
        if stock <= 0:
            continue

        segs = []
        line_markers = []
        cut_list_html = []
        running = 0.0

        def fmt_degree(value):
            try:
                deg = float(value)
                return f"{int(deg) if deg.is_integer() else round(deg,1)}°"
            except Exception:
                txt = str(value).strip()
                return txt if "°" in txt else f"{txt}°"

        for idx, cut in enumerate(cuts, start=1):
            length = float(cut.get("length_mm", 0) or 0)
            degree_text = fmt_degree(cut.get("cut_degree", "90"))
            win = cut.get("window_label", "")
            label = f"{win}<br>{int(length)} mm<br>{degree_text}"
            width_pct = max((length / stock) * 100.0, 2.0)
            segs.append(
                f'<div style="width:{width_pct:.3f}%;min-width:60px;height:86px;background:#dbeafe;border-right:1px solid white;display:flex;align-items:center;justify-content:center;text-align:center;font-size:11px;font-weight:700;color:#1e3a8a;padding:4px;overflow:hidden;line-height:1.12;">{label}</div>'
            )
            running += length

            if idx < len(cuts):
                left_pct = max(min((running / stock) * 100.0, 100.0), 0.0)
                line_markers.append(
                    f'<div style="position:absolute;left:calc({left_pct:.4f}% - 1px);top:0;bottom:0;width:2px;background:#0f172a;"></div>'
                    f'<div style="position:absolute;left:calc({left_pct:.4f}% - 28px);top:-26px;background:#0f172a;color:white;border-radius:999px;padding:3px 8px;font-size:11px;font-weight:700;">{degree_text}</div>'
                )

            cut_list_html.append(
                f'<div style="padding:7px 10px;border:1px solid rgba(148,163,184,0.18);border-radius:10px;background:white;font-size:12px;">'
                f'<b>{idx}.</b> {win} · {int(length)} mm · Cut angle <b>{degree_text}</b></div>'
            )

        if waste > 0:
            waste_pct = max((waste / stock) * 100.0, 2.0)
            segs.append(
                f'<div style="width:{waste_pct:.3f}%;min-width:52px;height:86px;background:#fee2e2;display:flex;align-items:center;justify-content:center;text-align:center;font-size:11px;font-weight:700;color:#991b1b;padding:4px;line-height:1.1;">Waste<br>{int(waste)} mm</div>'
            )

        st.markdown(
            f"""
            <div class="soft-card">
                <div style="display:flex;justify-content:space-between;gap:12px;align-items:center;flex-wrap:wrap;margin-bottom:10px;">
                    <div style="font-weight:800;color:#0f172a;">Bar {bar.get('bar_no','')}</div>
                    <div style="color:#475569;font-size:0.9rem;">Profile: <b>{bar.get('profile','')}</b> &nbsp;|&nbsp; Stock: <b>{int(stock)} mm</b> &nbsp;|&nbsp; Used: <b>{int(used)} mm</b> &nbsp;|&nbsp; Waste: <b>{int(waste)} mm</b></div>
                </div>
                <div style="position:relative;padding-top:30px;margin-bottom:10px;">
                    <div style="display:flex;width:100%;border-radius:14px;overflow:hidden;border:1px solid rgba(148,163,184,0.25);background:white;">
                        {''.join(segs)}
                    </div>
                    {''.join(line_markers)}
                </div>
                <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:8px;">{''.join(cut_list_html) if cut_list_html else '<div class="mini-tag">Single cut piece</div>'}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )


def render_glass_sheet_layouts(glass: dict, sheet_w: float, sheet_h: float):
    jobs = glass.get("optimiser_jobs", []) or []
    unplaced = glass.get("unplaced_jobs", []) or []
    if unplaced:
        st.error(f"{len(unplaced)} glass piece(s) could not fit on the selected sheet size. Increase the sheet size or review those piece dimensions in Results.")
    if not jobs:
        if not unplaced:
            st.info("All glass pieces were covered by offcuts.")
        return

    st.markdown('<div class="section-title">Visual glass sheet layouts</div>', unsafe_allow_html=True)

    by_sheet = {}
    for item in jobs:
        by_sheet.setdefault(item.get("sheet_no", 1), []).append(item)

    for sheet_no, items in sorted(by_sheet.items(), key=lambda x: x[0]):
        scale = min(700.0 / max(sheet_w, 1), 420.0 / max(sheet_h, 1))
        canvas_w = max(int(sheet_w * scale), 240)
        canvas_h = max(int(sheet_h * scale), 180)

        pieces_html = []
        for i, item in enumerate(items):
            x = float(item.get("x_mm", 0) or 0)
            y = float(item.get("y_mm", 0) or 0)
            w = float(item.get("placed_width_mm", item.get("width_mm", 0)) or 0)
            h = float(item.get("placed_height_mm", item.get("height_mm", 0)) or 0)
            left = int(x * scale)
            top = int(y * scale)
            width = max(int(w * scale), 56)
            height = max(int(h * scale), 40)
            label = item.get("window_label", f"W{i+1}")
            subtitle = f"{int(w)} x {int(h)}"
            piece = item.get("piece_id", "")
            pieces_html.append(
                f'<div style="position:absolute;left:{left}px;top:{top}px;width:{width}px;height:{height}px;background:rgba(15,118,110,0.18);border:2px solid #0f766e;border-radius:8px;box-sizing:border-box;display:flex;align-items:center;justify-content:center;text-align:center;font-size:10px;font-weight:700;color:#134e4a;padding:3px;overflow:hidden;line-height:1.05;">{label}<br>{subtitle}<br>{piece}</div>'
            )

        st.markdown(
            f"""
            <div class="soft-card">
                <div style="display:flex;justify-content:space-between;gap:12px;align-items:center;flex-wrap:wrap;margin-bottom:10px;">
                    <div style="font-weight:800;color:#0f172a;">Sheet {sheet_no}</div>
                    <div style="color:#475569;font-size:0.9rem;">Sheet size: <b>{int(sheet_w)} x {int(sheet_h)} mm</b> &nbsp;|&nbsp; Pieces: <b>{len(items)}</b></div>
                </div>
                <div style="position:relative;width:{canvas_w}px;height:{canvas_h}px;max-width:100%;border:2px solid #0f172a;border-radius:16px;background:linear-gradient(180deg,#ffffff,#f8fafc);overflow:hidden;">
                    {''.join(pieces_html)}
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )


def build_layout_pdf(project_name: str, client_name: str, finish: str, stock_length_mm: float, glass_sheet_width_mm: float, glass_sheet_height_mm: float, summary: dict, aluminium: dict, glass: dict) -> bytes:
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=landscape(A4))
    page_w, page_h = landscape(A4)

    def new_page(title: str, subtitle: str = ""):
        c.setFont("Helvetica-Bold", 20)
        c.drawString(30, page_h - 35, title)
        if subtitle:
            c.setFont("Helvetica", 10)
            c.setFillColor(colors.HexColor("#475569"))
            c.drawString(30, page_h - 52, subtitle)
            c.setFillColor(colors.black)

    def footer():
        c.setFont("Helvetica", 8)
        c.setFillColor(colors.HexColor("#64748b"))
        c.drawRightString(page_w - 24, 16, "Ali Fabrication Layout Pack")
        c.setFillColor(colors.black)

    new_page("Cutting Layout Summary", f"Project: {project_name}   Client: {client_name or '-'}   Finish: {finish}")
    c.setFont("Helvetica-Bold", 12)
    y = page_h - 90
    rows = [
        ("Window lines", str(summary.get("window_lines", 0))),
        ("Profile cuts", str(summary.get("profile_cuts", 0))),
        ("New aluminium bars", str(summary.get("aluminium_new_bars", 0))),
        ("Glass offcut hits", str(summary.get("glass_offcut_hits", 0))),
        ("New glass sheets", str(summary.get("glass_new_sheets", 0))),
        ("Total profile length", f"{round(float(summary.get('total_profile_length_mm', 0))/1000.0, 2)} m"),
        ("Total glass area", f"{round(float(summary.get('total_glass_area_m2', 0)), 2)} m²"),
        ("Aluminium stock length", f"{int(stock_length_mm)} mm"),
        ("Glass sheet size", f"{int(glass_sheet_width_mm)} x {int(glass_sheet_height_mm)} mm"),
    ]
    for label, value in rows:
        c.setFillColor(colors.HexColor("#0f172a"))
        c.drawString(40, y, label)
        c.setFillColor(colors.HexColor("#1d4ed8"))
        c.drawRightString(310, y, value)
        c.setFillColor(colors.black)
        y -= 20

    unplaced = glass.get("unplaced_jobs", []) or []
    c.setFillColor(colors.HexColor("#991b1b"))
    c.setFont("Helvetica-Bold", 11)
    if unplaced:
        c.drawString(360, page_h - 90, "Unplaced glass pieces")
        c.setFont("Helvetica", 9)
        yy = page_h - 110
        for item in unplaced[:15]:
            c.drawString(360, yy, f"{item.get('piece_id', item.get('window_label', 'Piece'))}: {int(item.get('width_mm',0))} x {int(item.get('height_mm',0))} mm")
            yy -= 14
    footer()
    c.showPage()

    bars = aluminium.get("bars", []) or []
    for bar in bars:
        new_page(f"Aluminium Bar {bar.get('bar_no', '')}", f"Profile: {bar.get('profile', '')}   Stock: {int(bar.get('stock_length_mm', 0))} mm")
        x0 = 40
        y0 = page_h / 2
        total_w = page_w - 80
        bar_h = 60
        c.setStrokeColor(colors.HexColor("#0f172a"))
        c.rect(x0, y0, total_w, bar_h, stroke=1, fill=0)
        stock = float(bar.get("stock_length_mm", 0) or 1)
        cursor = x0
        list_y = y0 - 24
        for idx, cut in enumerate(bar.get("cuts", []) or [], start=1):
            seg_w = max((float(cut.get("length_mm", 0))/stock) * total_w, 18)
            c.setFillColor(colors.HexColor("#dbeafe"))
            c.rect(cursor, y0, seg_w, bar_h, stroke=1, fill=1)
            c.setFillColor(colors.HexColor("#1e3a8a"))
            c.setFont("Helvetica-Bold", 7)
            txt1 = f"{cut.get('window_label', '')}"
            txt2 = f"{int(cut.get('length_mm', 0))} mm / {cut.get('cut_degree', '90')}°"
            c.drawCentredString(cursor + seg_w/2, y0 + (bar_h/2) + 6, txt1[:20])
            c.drawCentredString(cursor + seg_w/2, y0 + (bar_h/2) - 6, txt2[:24])
            c.setFillColor(colors.black)
            c.setFont("Helvetica", 9)
            c.drawString(40, list_y, f"{idx}. {cut.get('window_label', '')}  {int(cut.get('length_mm', 0))} mm  {cut.get('cut_degree', '90')}°")
            list_y -= 12
            cursor += seg_w
        waste = float(bar.get("waste_mm", 0) or 0)
        if waste > 0 and cursor < x0 + total_w:
            c.setFillColor(colors.HexColor("#fee2e2"))
            c.rect(cursor, y0, x0 + total_w - cursor, bar_h, stroke=1, fill=1)
            c.setFillColor(colors.HexColor("#991b1b"))
            c.drawCentredString(cursor + (x0 + total_w - cursor)/2, y0 + bar_h/2, f"Waste {int(waste)} mm")
        footer()
        c.showPage()

    jobs = glass.get("optimiser_jobs", []) or []
    by_sheet = {}
    for item in jobs:
        by_sheet.setdefault(item.get("sheet_no", 1), []).append(item)

    for sheet_no, items in sorted(by_sheet.items(), key=lambda x: x[0]):
        new_page(f"Glass Sheet {sheet_no}", f"Sheet size: {int(glass_sheet_width_mm)} x {int(glass_sheet_height_mm)} mm")
        x0, y0 = 40, 80
        draw_w, draw_h = page_w - 80, page_h - 150
        scale = min(draw_w / max(glass_sheet_width_mm, 1), draw_h / max(glass_sheet_height_mm, 1))
        sheet_w = glass_sheet_width_mm * scale
        sheet_h = glass_sheet_height_mm * scale
        c.setStrokeColor(colors.black)
        c.rect(x0, y0, sheet_w, sheet_h, stroke=1, fill=0)
        palette = ["#ccfbf1", "#bfdbfe", "#fde68a", "#fecaca", "#ddd6fe", "#fed7aa"]
        for i, item in enumerate(items):
            x = x0 + float(item.get("x_mm", 0) or 0) * scale
            y = y0 + float(item.get("y_mm", 0) or 0) * scale
            w = float(item.get("placed_width_mm", item.get("width_mm", 0)) or 0) * scale
            h = float(item.get("placed_height_mm", item.get("height_mm", 0)) or 0) * scale
            c.setFillColor(colors.HexColor(palette[i % len(palette)]))
            c.rect(x, y, w, h, stroke=1, fill=1)
            c.setFillColor(colors.HexColor("#134e4a"))
            c.setFont("Helvetica-Bold", 7)
            c.drawCentredString(x + w/2, y + h/2 + 8, f"{item.get('window_label', 'W')}"[:18])
            c.setFont("Helvetica", 6)
            c.drawCentredString(x + w/2, y + h/2, f"{int(item.get('placed_width_mm', item.get('width_mm', 0)))} x {int(item.get('placed_height_mm', item.get('height_mm', 0)))}")
            c.drawCentredString(x + w/2, y + h/2 - 8, f"{item.get('piece_id', 'P')}"[:18])
        footer()
        c.showPage()

    c.save()
    return buffer.getvalue()


def blank_window(default_variant: str, default_glass_spec: str, variant_map: dict[str, str], next_id: int, label: str):
    base = {
        "id": next_id,
        "label": label,
        "variant_key": default_variant,
        "variant_label": variant_map[default_variant],
        "window_qty": 1,
        "glass_spec": default_glass_spec,
    }
    for field in ALL_INPUT_FIELDS:
        base[field] = 0.0
    base["OVERALL WIDTH"] = 1200.0
    base["OVERALL HEIGHT"] = 1500.0
    return base


def set_default_windows(default_variant: str, default_glass_spec: str, variant_map: dict[str, str]):
    if "windows" not in st.session_state:
        st.session_state.windows = [blank_window(default_variant, default_glass_spec, variant_map, 1, "W1")]
    if "next_window_id" not in st.session_state:
        st.session_state.next_window_id = 2


def ensure_supporting_state():
    if "al_offcuts" not in st.session_state:
        st.session_state.al_offcuts = default_aluminium_offcuts()
    if "glass_offcuts" not in st.session_state:
        st.session_state.glass_offcuts = get_default_glass_offcuts()[:60]


def add_window(default_variant: str, default_glass_spec: str, variant_map: dict[str, str]):
    next_id = st.session_state.next_window_id
    st.session_state.windows.append(blank_window(default_variant, default_glass_spec, variant_map, next_id, f"W{len(st.session_state.windows)+1}"))
    st.session_state.next_window_id += 1


def duplicate_window(window_id: int):
    windows = st.session_state.windows
    for i, w in enumerate(windows):
        if w["id"] == window_id:
            next_id = st.session_state.next_window_id
            clone = dict(w)
            clone["id"] = next_id
            clone["label"] = f"W{len(windows)+1}"
            windows.insert(i + 1, clone)
            st.session_state.next_window_id += 1
            break


def remove_window(window_id: int):
    st.session_state.windows = [w for w in st.session_state.windows if w["id"] != window_id]


def update_window_field(index: int, field: str, value):
    st.session_state.windows[index][field] = value


catalog = get_catalog()
variant_options = catalog.list_variant_options()
variant_lookup = catalog.variant_lookup()
variant_keys = [k for k, _ in variant_options]
variant_map = dict(variant_options)
default_variant = variant_options[0][0]
glass_specs = get_glass_specs()
default_glass_spec = "6.38MM GREY TINTED LAMINATED GLASS" if "6.38MM GREY TINTED LAMINATED GLASS" in glass_specs else glass_specs[0]

set_default_windows(default_variant, default_glass_spec, variant_map)
ensure_supporting_state()

if "active_project_name" not in st.session_state:
    saved = list_saved_projects()
    if saved:
        latest_name = saved[-1]
        payload = load_project_data(latest_name)
        if payload:
            apply_project_payload(payload, default_variant, default_glass_spec, variant_map)
            st.session_state.active_project_name = latest_name
        else:
            st.session_state.active_project_name = "Ali Fabrication Project"
            st.session_state.project_name_value = "Ali Fabrication Project"
    else:
        st.session_state.active_project_name = "Ali Fabrication Project"
        st.session_state.project_name_value = "Ali Fabrication Project"
        st.session_state.client_name_value = ""
        st.session_state.finish_value = "Powder Coated"
        st.session_state.stock_length_mm_value = 6400.0
        st.session_state.glass_sheet_width_mm_value = 3660.0
        st.session_state.glass_sheet_height_mm_value = 2440.0
        st.session_state.kerf_mm_value = 3.0
        st.session_state.default_row_glass_value = default_glass_spec

with st.sidebar:
    st.header("Projects")
    saved_projects = list_saved_projects()
    selectable_projects = saved_projects if saved_projects else [st.session_state.active_project_name]
    current_idx = selectable_projects.index(st.session_state.active_project_name) if st.session_state.active_project_name in selectable_projects else 0
    selected_project = st.selectbox("Open saved project", selectable_projects, index=current_idx)
    if selected_project != st.session_state.active_project_name:
        payload = load_project_data(selected_project)
        if payload:
            apply_project_payload(payload, default_variant, default_glass_spec, variant_map)
            st.session_state.active_project_name = selected_project
            st.rerun()

    new_project_name = st.text_input("Create new project", value="", placeholder="e.g. ABC Apartments")
    p1, p2 = st.columns(2)
    with p1:
        if st.button("New project", use_container_width=True):
            name = (new_project_name or "New Project").strip()
            st.session_state.active_project_name = name
            st.session_state.windows = [blank_window(default_variant, default_glass_spec, variant_map, 1, "W1")]
            st.session_state.next_window_id = 2
            st.session_state.al_offcuts = default_aluminium_offcuts()
            st.session_state.glass_offcuts = get_default_glass_offcuts()[:60]
            st.session_state.project_name_value = name
            st.session_state.client_name_value = ""
            st.session_state.finish_value = "Powder Coated"
            st.session_state.stock_length_mm_value = 6400.0
            st.session_state.glass_sheet_width_mm_value = 3660.0
            st.session_state.glass_sheet_height_mm_value = 2440.0
            st.session_state.kerf_mm_value = 3.0
            st.session_state.default_row_glass_value = default_glass_spec
            save_project_data(name, project_payload_from_state(name, "", "Powder Coated", 6400.0, 3660.0, 2440.0, 3.0, default_glass_spec))
            st.rerun()
    with p2:
        if st.button("Delete project", use_container_width=True):
            delete_project_data(st.session_state.active_project_name)
            remaining = list_saved_projects()
            if remaining:
                payload = load_project_data(remaining[-1])
                if payload:
                    apply_project_payload(payload, default_variant, default_glass_spec, variant_map)
                    st.session_state.active_project_name = remaining[-1]
            else:
                st.session_state.active_project_name = "Ali Fabrication Project"
                st.session_state.windows = [blank_window(default_variant, default_glass_spec, variant_map, 1, "W1")]
                st.session_state.next_window_id = 2
                st.session_state.al_offcuts = default_aluminium_offcuts()
                st.session_state.glass_offcuts = get_default_glass_offcuts()[:60]
                st.session_state.project_name_value = "Ali Fabrication Project"
                st.session_state.client_name_value = ""
                st.session_state.finish_value = "Powder Coated"
                st.session_state.stock_length_mm_value = 6400.0
                st.session_state.glass_sheet_width_mm_value = 3660.0
                st.session_state.glass_sheet_height_mm_value = 2440.0
                st.session_state.kerf_mm_value = 3.0
                st.session_state.default_row_glass_value = default_glass_spec
            st.rerun()

    st.caption(f"Active project: {st.session_state.active_project_name}")
    st.caption("Changes autosave as you work and should remain after refresh.")

    st.header("Project Details")
    project_name = st.text_input("Project name", key="project_name_value")
    client_name = st.text_input("Client name", key="client_name_value")
    finish = st.text_input("Aluminium finish", key="finish_value")
    primary = "#0F4C81"
    accent = "#F28C36"
    logo_file = None

    st.header("Material Controls")
    stock_length_mm = st.number_input("Aluminium stock length (mm)", min_value=1000.0, step=100.0, key="stock_length_mm_value")
    glass_sheet_width_mm = st.number_input("Glass sheet width (mm)", min_value=500.0, step=10.0, key="glass_sheet_width_mm_value")
    glass_sheet_height_mm = st.number_input("Glass sheet height (mm)", min_value=500.0, step=10.0, key="glass_sheet_height_mm_value")
    kerf_mm = st.number_input("Saw / cut kerf (mm)", min_value=0.0, step=0.5, key="kerf_mm_value")
    default_glass_index = glass_specs.index(st.session_state.default_row_glass_value) if st.session_state.default_row_glass_value in glass_specs else 0
    default_row_glass = st.selectbox("Default glass specification", glass_specs, index=default_glass_index, key="default_row_glass_value")

inject_brand_css(primary, accent)

logo_b64 = uploaded_or_default_logo(logo_file)
logo_html = f'<img class="hero-logo" src="data:image/png;base64,{logo_b64}" />' if logo_b64 else ""

st.markdown(
    f"""
    <div class="brand-hero">
        <div class="hero-wrap">
            {logo_html}
            <div>
                <h1 class="hero-title">{project_name}</h1>
                <div class="hero-sub">Master craftsmanship planning for aluminium profiles, glass optimisation, offcuts, and export-ready jobcards.</div>
            </div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    f"""
    <div class="soft-card" style="padding:14px 18px 12px 18px;">
        <div style="display:flex;justify-content:space-between;align-items:center;gap:12px;flex-wrap:wrap;">
            <div>
                <div style="font-size:1.05rem;font-weight:800;color:#0f172a;">Client: {client_name or '-'}</div>
                <div style="font-size:0.92rem;color:#475569;margin-top:4px;">Finish: {finish or '-'} &nbsp;|&nbsp; Autosaved project: <b>{st.session_state.active_project_name}</b></div>
            </div>
            <div style="background:#eff6ff;color:#1d4ed8;border:1px solid rgba(29,78,216,0.12);padding:6px 12px;border-radius:999px;font-size:0.8rem;font-weight:700;">
                PROJECT ACTIVE
            </div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

tab1, tab2, tab3, tab4 = st.tabs(["Window Entry", "Offcuts", "Results", "Visual Layouts"])

with tab1:
    st.markdown('<div class="section-title">Window Calculator</div>', unsafe_allow_html=True)
    st.caption("Each selected system and orientation now shows only the input fields required for that exact variant.")

    for idx, window in enumerate(st.session_state.windows):
        current_variant = window.get("variant_key", default_variant)
        current_variant_meta = variant_lookup[current_variant]
        current_required = sorted(current_variant_meta.get("input_labels", []), key=lambda x: FIELD_ORDER.get(x, 999))
        title = f"{window.get('label', f'W{idx+1}')} — {variant_map.get(current_variant, 'Typology')}"
        with st.expander(title, expanded=True):
            form_col, preview_col = st.columns([2.4, 1])

            with form_col:
                c1, c2, c3, c4 = st.columns([1, 2, 1, 1])
                with c1:
                    label = st.text_input("Label", value=window.get("label", f"W{idx+1}"), key=f"label_{window['id']}")
                with c2:
                    variant_key = st.selectbox(
                        "Typology / Orientation",
                        options=variant_keys,
                        index=variant_keys.index(current_variant),
                        format_func=lambda x: variant_map.get(x, x),
                        key=f"variant_{window['id']}",
                    )
                with c3:
                    qty = st.number_input("Qty", min_value=1, value=int(window.get("window_qty", 1)), step=1, key=f"qty_{window['id']}")
                with c4:
                    row_glass = st.selectbox(
                        "Glass Spec",
                        options=glass_specs,
                        index=glass_specs.index(window.get("glass_spec", default_row_glass)) if window.get("glass_spec", default_row_glass) in glass_specs else 0,
                        key=f"glass_{window['id']}",
                    )

                selected_variant_meta = variant_lookup[variant_key]
                required_fields = sorted(selected_variant_meta.get("input_labels", []), key=lambda x: FIELD_ORDER.get(x, 999))

                st.markdown("**Required inputs for this exact system:**", unsafe_allow_html=False)
                st.markdown("".join([f'<span class="mini-tag">{f}</span>' for f in required_fields]), unsafe_allow_html=True)

                if required_fields:
                    cols = st.columns(3)
                    for pos, field in enumerate(required_fields):
                        with cols[pos % 3]:
                            value = st.number_input(
                                f"{field} (mm)",
                                min_value=0.0,
                                value=float(window.get(field, 0.0)),
                                step=10.0,
                                key=f"{safe_name(field)}_{window['id']}",
                                help=FIELD_HELP.get(field, ""),
                            )
                            update_window_field(idx, field, value)

                for field in ALL_INPUT_FIELDS:
                    if field not in required_fields:
                        # keep hidden fields but do not force any value
                        update_window_field(idx, field, float(window.get(field, 0.0)))

                update_window_field(idx, "label", label)
                update_window_field(idx, "variant_key", variant_key)
                update_window_field(idx, "variant_label", variant_map.get(variant_key, variant_key))
                update_window_field(idx, "window_qty", qty)
                update_window_field(idx, "glass_spec", row_glass)

                action1, action2, action_spacer = st.columns([1, 1, 3])
                with action1:
                    if st.button("Duplicate window", key=f"dup_{window['id']}", use_container_width=True):
                        duplicate_window(window["id"])
                        st.rerun()
                with action2:
                    if len(st.session_state.windows) > 1 and st.button("Remove this window", key=f"remove_{window['id']}", use_container_width=True):
                        remove_window(window["id"])
                        st.rerun()

            with preview_col:
                selected_variant_label = variant_map.get(variant_key, variant_key)
                selected_code = system_code_from_label(selected_variant_label)
                preview_b64 = preview_base64(variant_key)
                variant_short = variant_short_name(selected_variant_label)
                img_html = f'<img class="preview-img" src="data:image/png;base64,{preview_b64}" />' if preview_b64 else '<div style="padding:40px 0;color:#64748b;">Preview not available</div>'
                st.markdown(
                    f"""
                    <div class="preview-shell">
                        <div class="preview-code">System Code: {selected_code}</div>
                        <div class="preview-variant">{variant_short}</div>
                        {img_html}
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

    st.markdown("<div style='height:8px;'></div>", unsafe_allow_html=True)
    add_col, spacer_col = st.columns([1.2, 4])
    with add_col:
        if st.button("Add window", use_container_width=True):
            add_window(default_variant, default_row_glass, variant_map)
            st.rerun()
    with spacer_col:
        st.caption("Add the next window from the bottom of the page so you do not need to scroll back up.")

with tab2:
    st.markdown('<div class="section-title">Stock & Offcuts</div>', unsafe_allow_html=True)
    a1, a2 = st.columns(2)
    with a1:
        st.caption("Aluminium offcuts")
        al_df = pd.DataFrame(st.session_state.al_offcuts or [{"profile": "", "length_mm": 0.0, "qty": 1}])
        al_edited = st.data_editor(
            al_df,
            use_container_width=True,
            num_rows="dynamic",
            key="al_offcuts_editor_variant",
            column_config={
                "profile": st.column_config.TextColumn("Profile"),
                "length_mm": st.column_config.NumberColumn("Length (mm)", min_value=0.0),
                "qty": st.column_config.NumberColumn("Qty", min_value=1, step=1),
            },
        )
        st.session_state.al_offcuts = al_edited.fillna("").to_dict("records")
    with a2:
        st.caption("Glass offcuts")
        glass_df = pd.DataFrame(st.session_state.glass_offcuts or [{"spec": default_row_glass, "width_mm": 0.0, "height_mm": 0.0, "qty": 1}])
        glass_edited = st.data_editor(
            glass_df,
            use_container_width=True,
            num_rows="dynamic",
            key="glass_offcuts_editor_variant",
            column_config={
                "spec": st.column_config.SelectboxColumn("Specification", options=glass_specs),
                "width_mm": st.column_config.NumberColumn("Width (mm)", min_value=0.0),
                "height_mm": st.column_config.NumberColumn("Height (mm)", min_value=0.0),
                "qty": st.column_config.NumberColumn("Qty", min_value=1, step=1),
            },
        )
        st.session_state.glass_offcuts = glass_edited.fillna("").to_dict("records")

profile_rows, glass_rows, warnings = expand_window_rows(st.session_state.windows, catalog, default_row_glass)
aluminium = optimise_aluminium(profile_rows, stock_length_mm, kerf_mm, st.session_state.al_offcuts)
glass = optimise_glass(glass_rows, glass_sheet_width_mm, glass_sheet_height_mm, kerf_mm, st.session_state.glass_offcuts)
summary = build_summary(st.session_state.windows, profile_rows, glass_rows, aluminium, glass, catalog.weights)

active_name = (project_name or st.session_state.active_project_name or "Ali Fabrication Project").strip()
st.session_state.active_project_name = active_name
save_project_data(
    active_name,
    project_payload_from_state(
        active_name,
        client_name,
        finish,
        stock_length_mm,
        glass_sheet_width_mm,
        glass_sheet_height_mm,
        kerf_mm,
        default_row_glass,
    ),
)

bar_df = pd.DataFrame(aluminium["bars"])
if not bar_df.empty:
    order_breakdown = (
        bar_df.groupby("profile", dropna=False)
        .agg(
            bars_to_order=("bar_no", "count"),
            ordered_length_mm=("stock_length_mm", "sum"),
            used_length_mm=("used_mm", "sum"),
            waste_mm=("waste_mm", "sum"),
        )
        .reset_index()
    )
    order_breakdown["ordered_length_m"] = order_breakdown["ordered_length_mm"].apply(mm_to_m)
    order_breakdown["used_length_m"] = order_breakdown["used_length_mm"].apply(mm_to_m)
    order_breakdown["waste_m"] = order_breakdown["waste_mm"].apply(mm_to_m)
else:
    order_breakdown = pd.DataFrame(columns=["profile", "bars_to_order", "ordered_length_mm", "ordered_length_m", "used_length_mm", "used_length_m", "waste_mm", "waste_m"])

profile_piece_df = pd.DataFrame(profile_rows)
if not profile_piece_df.empty:
    profile_totals = (
        profile_piece_df.groupby("profile", dropna=False)
        .agg(total_cut_qty=("qty", "sum"), total_cut_length_mm=("length_mm", lambda s: float((s * profile_piece_df.loc[s.index, "qty"]).sum())))
        .reset_index()
    )
    profile_totals["total_cut_length_m"] = profile_totals["total_cut_length_mm"].apply(mm_to_m)
else:
    profile_totals = pd.DataFrame(columns=["profile", "total_cut_qty", "total_cut_length_mm", "total_cut_length_m"])

with tab3:
    for msg in warnings:
        st.warning(msg)
    if glass.get("unplaced_jobs"):
        st.error(f"{len(glass.get('unplaced_jobs', []))} glass piece(s) do not fit on the selected sheet size and were excluded from placement.")

    m1, m2, m3, m4 = st.columns(4)
    with m1:
        render_metric("Window lines", str(summary["window_lines"]))
    with m2:
        render_metric("Profile cuts", str(summary["profile_cuts"]))
    with m3:
        render_metric("New aluminium bars", str(summary["aluminium_new_bars"]))
    with m4:
        render_metric("New glass sheets", str(summary["glass_new_sheets"]))

    r1, r2 = st.columns([1.2, 1])
    with r1:
        st.markdown('<div class="soft-card">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">Profile breakdown</div>', unsafe_allow_html=True)
        if not profile_totals.empty:
            st.dataframe(profile_totals, use_container_width=True, hide_index=True)
        else:
            st.info("No profile cuts generated yet.")
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('<div class="soft-card">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">Lengths to order after offcuts</div>', unsafe_allow_html=True)
        st.caption("Actual stock to order after available offcuts have been deducted.")
        if not order_breakdown.empty:
            st.dataframe(
                order_breakdown[["profile", "bars_to_order", "ordered_length_mm", "ordered_length_m", "used_length_mm", "used_length_m", "waste_mm", "waste_m"]],
                use_container_width=True,
                hide_index=True,
            )
        else:
            st.success("All profile cuts were covered by offcuts. No new stock bars are required.")
        st.markdown('</div>', unsafe_allow_html=True)

    with r2:
        st.markdown('<div class="soft-card">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">Planning summary</div>', unsafe_allow_html=True)
        st.write(f"**Client:** {client_name or '-'}")
        st.write(f"**Finish:** {finish}")
        st.write(f"**Total profile length:** {mm_to_m(summary['total_profile_length_mm'])} m")
        st.write(f"**Estimated aluminium weight:** {round(summary['estimated_weight_kg'], 2)} kg")
        st.write(f"**Aluminium offcut hits:** {summary['aluminium_offcut_hits']}")
        st.write(f"**Glass offcut hits:** {summary['glass_offcut_hits']}")
        st.write(f"**Total glass area:** {round(summary['total_glass_area_m2'], 2)} m²")
        st.write(f"**Total aluminium waste from new bars:** {mm_to_m(aluminium['total_waste_mm'])} m")
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('<div class="soft-card">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">Export project workbook</div>', unsafe_allow_html=True)
        payload = export_project_workbook(
            {
                "project_name": project_name,
                "client_name": client_name,
                "finish": finish,
                "stock_length_mm": stock_length_mm,
                "glass_sheet_width_mm": glass_sheet_width_mm,
                "glass_sheet_height_mm": glass_sheet_height_mm,
            },
            st.session_state.windows,
            profile_rows,
            glass_rows,
            aluminium,
            glass,
            summary,
        )
        st.download_button(
            "Download Excel project file",
            data=payload,
            file_name=f"{project_name.strip().replace(' ', '_') or 'project'}_fabrication_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.markdown('</div>', unsafe_allow_html=True)

    g1, g2 = st.columns(2)
    with g1:
        st.markdown('<div class="soft-card">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">Aluminium offcut jobs</div>', unsafe_allow_html=True)
        offcut_df = pd.DataFrame(aluminium["offcut_jobs"])
        if not offcut_df.empty:
            st.dataframe(
                offcut_df[["window_label", "profile", "length_mm", "source_offcut_id", "source_length_mm", "remaining_after_mm"]],
                use_container_width=True,
                hide_index=True,
            )
        else:
            st.info("No aluminium offcut matches yet.")
        st.markdown('</div>', unsafe_allow_html=True)

    with g2:
        st.markdown('<div class="soft-card">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">Glass optimiser placements</div>', unsafe_allow_html=True)
        glass_df = pd.DataFrame(glass["optimiser_jobs"])
        if not glass_df.empty:
            view_cols = [c for c in ["piece_id", "window_label", "spec", "sheet_no", "x_mm", "y_mm", "placed_width_mm", "placed_height_mm", "rotated"] if c in glass_df.columns]
            st.dataframe(glass_df[view_cols], use_container_width=True, hide_index=True)
        else:
            st.info("All glass pieces were covered by offcuts.")
        unplaced_df = pd.DataFrame(glass.get("unplaced_jobs", []) or [])
        if not unplaced_df.empty:
            st.warning("Some glass pieces do not fit on the selected sheet size.")
            show_cols = [c for c in ["piece_id", "window_label", "spec", "width_mm", "height_mm", "reason"] if c in unplaced_df.columns]
            st.dataframe(unplaced_df[show_cols], use_container_width=True, hide_index=True)
        st.markdown('</div>', unsafe_allow_html=True)


with tab4:
    st.caption("This tab is tuned for workshop and mobile viewing. Scroll vertically on smaller screens.")
    d1, d2, d3 = st.columns([1, 1, 1])
    with d1:
        pdf_bytes = build_layout_pdf(project_name, client_name, finish, stock_length_mm, glass_sheet_width_mm, glass_sheet_height_mm, summary, aluminium, glass)
        st.download_button(
            "Download cutting layouts PDF",
            data=pdf_bytes,
            file_name=f"{project_name.strip().replace(' ', '_') or 'project'}_cutting_layouts.pdf",
            mime="application/pdf",
            use_container_width=True,
        )
    with d2:
        backup_json = json.dumps(project_payload_from_state(st.session_state.active_project_name, client_name, finish, stock_length_mm, glass_sheet_width_mm, glass_sheet_height_mm, kerf_mm, default_row_glass), indent=2)
        st.download_button(
            "Download project backup",
            data=backup_json,
            file_name=f"{st.session_state.active_project_name.strip().replace(' ', '_') or 'project'}_backup.json",
            mime="application/json",
            use_container_width=True,
        )
    with d3:
        st.info("The PDF includes a summary, aluminium bar layouts, glass sheet layouts, and any unplaced glass pieces.")
    c1, c2 = st.columns(2)
    with c1:
        render_aluminium_bar_layouts(aluminium)
    with c2:
        render_glass_sheet_layouts(glass, glass_sheet_width_mm, glass_sheet_height_mm)
