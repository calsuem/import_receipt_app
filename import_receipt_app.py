# import_receipt_app.py
# ------------------------------------------------------------
# 📄 수입신고필증 PDF → 엑셀 자동화 (전 항목 ROI 템플릿 · 영구 저장/불러오기 · 시각 오버레이)
# ✨ 한글 폰트: "맑은 고딕" 우선 + 폴백(Noto/Nanum) + PIL 라벨도 동일 적용
# ✨ 템플릿: 최초 저장 → 이후 자동 사용(Last Used) · 필요시 전환/삭제/가져오기/내보내기
# ✨ 국내도착항: '항 코드(KRPTK 등) 무시' + '한글만 추출'로 변경
# ------------------------------------------------------------

import io
import json
import re
import os
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Tuple
from urllib.request import urlretrieve

import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
from PIL import Image, ImageDraw, ImageFont

try:
    from streamlit_image_coordinates import streamlit_image_coordinates
except ImportError:
    st.error("📦 streamlit-image-coordinates가 필요합니다.\n\npip install streamlit-image-coordinates")
    st.stop()

# =========================
# 전역 설정 및 상수
# =========================
st.set_page_config(page_title="수입신고필증 PDF → 엑셀 (ROI 템플릿)", page_icon="📄", layout="wide")

TEMPLATES_FILE = "receipt_templates.json"   # 앱 폴더 내 JSON로 영구 저장
DPI_DEFAULT = 144

# 출력 컬럼(=필드) 순서
FIELDS: List[str] = [
    "b/l(awb)번호",
    "국내도착항",
    "신고일",
    "환율",
    "세율(구분)",
    "부가가치세 과표",
    "관세",
    "부가가치세",
    "신고번호",
]

# 필드별 색상 (오버레이용)
FIELD_COLORS: Dict[str, str] = {
    "b/l(awb)번호": "#E74C3C",
    "국내도착항":   "#3498DB",
    "신고일":       "#2ECC71",
    "환율":         "#9B59B6",
    "세율(구분)":    "#F1C40F",
    "부가가치세 과표": "#1ABC9C",
    "관세":         "#E67E22",
    "부가가치세":     "#34495E",
    "신고번호":      "#D35400",
}

_AMOUNT_PATTERN = r'([0-9]{1,3}(?:,[0-9]{3})+|[0-9]{4,})\s*원?'

# =========================
# 한글 폰트 설정 (UI + PIL 공통)
# =========================
def ensure_korean_fonts():
    """
    Streamlit UI 전역 폰트와 PIL 라벨 폰트를 설정한다.
    - UI: CSS로 '맑은 고딕' 우선, 폴백 스택 지정
    - PIL: 시스템/동봉/다운로드 순으로 폰트 탐색하여 적용
    """
    css_font_stack = "'Malgun Gothic','Apple SD Gothic Neo','Nanum Gothic','Noto Sans KR',sans-serif"
    st.markdown(
        f"""
        <style>
        html, body, [class*="css"]  {{
            font-family: {css_font_stack} !important;
        }}
        .stButton>button, .stDownloadButton>button, .stTextInput>div>div>input,
        .stSelectbox>div>div>select, .stMarkdown, .stDataFrame, .stCaption {{
            font-family: {css_font_stack} !important;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )

    # PIL 라벨 폰트
    candidates = [
        "C:/Windows/Fonts/malgun.ttf",  # Windows - 맑은 고딕
        "/System/Library/Fonts/AppleSDGothicNeo.ttc",  # macOS
        "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",  # Ubuntu/Nanum
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/truetype/noto/NotoSansKR-Regular.otf",
    ]
    font_path = None
    for p in candidates:
        if os.path.exists(p):
            font_path = p
            break
    if font_path is None:
        font_dir = Path("fonts"); font_dir.mkdir(exist_ok=True)
        font_path = str(font_dir / "NotoSansKR-Regular.otf")
        if not os.path.exists(font_path):
            try:
                urlretrieve(
                    "https://github.com/googlefonts/noto-cjk/raw/main/Sans/OTF/Korean/NotoSansKR-Regular.otf",
                    font_path
                )
            except Exception:
                font_path = None

    pil_font = None
    try:
        if font_path and os.path.exists(font_path):
            pil_font = ImageFont.truetype(font_path, size=16)
    except Exception:
        pil_font = None
    return pil_font

PIL_LABEL_FONT = ensure_korean_fonts()

# =========================
# 간단 유틸
# =========================
def fmt_date_uniform(s: str) -> str:
    """여러 날짜 표기를 YYYY/MM/DD로 표준화 (연도 없으면 당해년도, 실패시 YYYY/00/00)."""
    cur_year = datetime.now().year
    if not s:
        return f"{cur_year:04d}/00/00"
    s = " ".join(str(s).split())
    m = re.search(r'(\d{4})[./-](\d{1,2})[./-](\d{1,2})', s)
    if m:
        y, mo, d = m.groups()
        return f"{int(y):04d}/{int(mo):02d}/{int(d):02d}"
    m = re.search(r'(\d{4})\s*년\s*(\d{1,2})\s*월\s*(\d{1,2})\s*일', s)
    if m:
        y, mo, d = m.groups()
        return f"{int(y):04d}/{int(mo):02d}/{int(d):02d}"
    m = re.search(r'(?<!\d)(\d{1,2})[./-](\d{1,2})(?!\d)', s)
    if m:
        mo, d = m.groups()
        return f"{cur_year:04d}/{int(mo):02d}/{int(d):02d}"
    m = re.search(r'(?<!\d)(\d{8})(?!\d)', s)  # YYYYMMDD
    if m:
        val = m.group(1)
        return f"{int(val[:4]):04d}/{int(val[4:6]):02d}/{int(val[6:]):02d}"
    return f"{cur_year:04d}/00/00"


def clean_number(s: str):
    """천단위 콤마 제거 → int/float 변환."""
    if s is None:
        return None
    t = str(s).replace(",", "").strip()
    if t == "":
        return None
    try:
        return float(t) if re.search(r'\d+\.\d+', t) else int(t)
    except Exception:
        return None

# =========================
# 템플릿 저장/불러오기
# 구조:
# {
#   "__meta": {"last_used": "템플릿명"},
#   "템플릿명": { "created_at": "...", "dpi": 144, "norm_rects": {"필드":[x1n,y1n,x2n,y2n], ...} }
# }
# =========================
def load_all_templates() -> Dict[str, dict]:
    path = Path(TEMPLATES_FILE)
    if path.exists():
        try:
            return json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            return {"__meta": {}}
    return {"__meta": {}}

def save_all_templates(all_tmpls: Dict[str, dict]) -> None:
    if "__meta" not in all_tmpls:
        all_tmpls["__meta"] = {}
    Path(TEMPLATES_FILE).write_text(json.dumps(all_tmpls, ensure_ascii=False, indent=2), encoding="utf-8")

def set_last_used(all_tmpls: Dict[str, dict], name: str):
    if "__meta" not in all_tmpls:
        all_tmpls["__meta"] = {}
    all_tmpls["__meta"]["last_used"] = name
    save_all_templates(all_tmpls)

def get_last_used(all_tmpls: Dict[str, dict]) -> str | None:
    return all_tmpls.get("__meta", {}).get("last_used")

# =========================
# PDF 관련: 렌더/텍스트 클립
# =========================
def pdf_first_page_pix(file_bytes: bytes, dpi: int = DPI_DEFAULT) -> tuple[Image.Image, int, int, fitz.Rect]:
    with fitz.open(stream=file_bytes, filetype="pdf") as doc:
        page = doc.load_page(0)
        mat = fitz.Matrix(dpi / 72.0, dpi / 72.0)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        return img, pix.width, pix.height, page.rect

def clip_text_by_norm_rect(file_bytes: bytes, norm_rect: List[float], page_rect: fitz.Rect) -> str:
    x1n, y1n, x2n, y2n = norm_rect
    x1 = page_rect.x0 + page_rect.width * x1n
    y1 = page_rect.y0 + page_rect.height * y1n
    x2 = page_rect.x0 + page_rect.width * x2n
    y2 = page_rect.y0 + page_rect.height * y2n
    clip = fitz.Rect(x1, y1, x2, y2)
    with fitz.open(stream=file_bytes, filetype="pdf") as doc:
        page = doc.load_page(0)
        txt = page.get_text("text", clip=clip) or ""
        return " ".join(txt.split())

# =========================
# 국내도착항(한글만) 추출 도우미
# =========================
def extract_korean_port(text: str) -> str:
    """
    ROI에서 읽어온 문자열에서 'KRPTK' 같은 영문/코드 제거하고,
    '평택항/인천항/부산항/김포공항' 등 '한글+항/공항/항만/항구' 패턴 우선 추출.
    없으면 한글 단어 중 가장 길게 보이는 것을 반환.
    """
    if not text:
        return ""
    t = " ".join(text.split())
    # 괄호 속 코드 제거: (KRPTK), (CODE) 등
    t = re.sub(r'\([^)]+\)', ' ', t)
    # 단독 대문자 코드 제거
    t = re.sub(r'\b[A-Z]{3,}\b', ' ', t)
    # 혼합 코드(숫자 포함) 제거
    t = re.sub(r'\b[A-Z0-9\-]{3,}\b', ' ', t)

    # 1) '한글 + (공항|항만|항구|항)' 패턴 우선
    m = re.search(r'([가-힣]+(?:공항|항만|항구|항))', t)
    if m:
        return m.group(1)

    # 2) 그 외 한글 단어 후보 중 가장 긴 것
    cand = re.findall(r'[가-힣]{2,}', t)
    if cand:
        cand.sort(key=len, reverse=True)
        return cand[0]

    # 3) 실패 시 원본 정리본
    return t.strip()

# =========================
# 필드 후처리 규칙 (ROI → 정제)
# =========================
def postprocess_field(name: str, raw: str):
    text = " ".join((raw or "").split())

    if name == "b/l(awb)번호":
        m = re.search(r'([A-Za-z0-9\-]+)', text)
        return m.group(1) if m else ""

    if name == "국내도착항":
        return extract_korean_port(text)

    if name == "신고일":
        return fmt_date_uniform(text)

    if name == "환율":
        m = re.search(r'([\d,]+\.\d+|\d+\.\d+|\d+)', text)
        return clean_number(m.group(1)) if m else clean_number(text)

    if name == "세율(구분)":
        m = re.search(r'관\s*([0-9.]+)', text)
        if not m:
            m = re.search(r'([0-9.]+)', text)
        return m.group(1) if m else ""

    if name in ("부가가치세 과표", "관세", "부가가치세"):
        m = re.search(_AMOUNT_PATTERN, text)
        if m:
            return clean_number(m.group(1))
        return clean_number(text)

    if name == "신고번호":
        m = re.search(r'\b(\d{5}-\d{2}-\d{6}M)\b', text)
        return m.group(1) if m else text

    return text

# =========================
# 상태 초기화 + 자동 템플릿 로드
# =========================
def ensure_state():
    if "all_templates" not in st.session_state:
        st.session_state.all_templates = load_all_templates()

    # 자동 로드: 마지막 사용 템플릿
    if "auto_loaded" not in st.session_state:
        st.session_state.auto_loaded = True
        last = get_last_used(st.session_state.all_templates)
        if last and last in st.session_state.all_templates:
            data = st.session_state.all_templates[last]
            st.session_state.template_name = last
            st.session_state.tmpl_dpi = data.get("dpi", DPI_DEFAULT)
            st.session_state.norm_rects = data.get("norm_rects", {})
        else:
            st.session_state.template_name = ""
            st.session_state.tmpl_dpi = DPI_DEFAULT
            st.session_state.norm_rects = {}

    if "display_width" not in st.session_state:
        st.session_state.display_width = 1000
    if "click_phase" not in st.session_state:
        st.session_state.click_phase = 0
    if "temp_points" not in st.session_state:
        st.session_state.temp_points = []
    if "current_field_idx" not in st.session_state:
        st.session_state.current_field_idx = 0
    if "lock_template" not in st.session_state:
        st.session_state.lock_template = True  # 기본: 마지막 템플릿 고정 사용

# =========================
# 오버레이 렌더링
# =========================
def render_with_overlays(img_resized: Image.Image, w_orig: int, h_orig: int, ratio: float,
                         norm_rects: Dict[str, List[float]], temp_points: List[Tuple[float, float]],
                         current_field: str | None) -> Image.Image:
    img = img_resized.copy()
    draw = ImageDraw.Draw(img)

    def to_disp(x, y):
        return (x * ratio, y * ratio)

    for name, rect in norm_rects.items():
        color = FIELD_COLORS.get(name, "#FF00FF")
        x1n, y1n, x2n, y2n = rect
        x1, y1 = x1n * w_orig, y1n * h_orig
        x2, y2 = x2n * w_orig, y2n * h_orig
        dx1, dy1 = to_disp(x1, y1); dx2, dy2 = to_disp(x2, y2)
        for off in range(2):
            draw.rectangle([dx1 - off, dy1 - off, dx2 + off, dy2 + off], outline=color, width=2)
        label = name
        pad = 6
        tw = max(60, len(label) * 10)
        th = 18
        bx1, by1 = dx1, max(0, dy1 - (th + pad*2 + 2))
        bx2, by2 = dx1 + tw + pad*2, by1 + th + pad*2
        draw.rectangle([bx1, by1, bx2, by2], fill=color)
        if PIL_LABEL_FONT:
            draw.text((bx1 + pad, by1 + pad), label, fill="white", font=PIL_LABEL_FONT)
        else:
            draw.text((bx1 + pad, by1 + pad), label, fill="white")

    if temp_points:
        color = FIELD_COLORS.get(current_field or "", "#FF00FF")
        if len(temp_points) == 1:
            x, y = temp_points[0]
            dx, dy = to_disp(x, y)
            draw.line([dx - 12, dy, dx + 12, dy], fill=color, width=3)
            draw.line([dx, dy - 12, dx, dy + 12], fill=color, width=3)
            draw.ellipse([dx - 6, dy - 6, dx + 6, dy + 6], outline=color, width=3)
        elif len(temp_points) == 2:
            (x1, y1), (x2, y2) = temp_points
            x1, x2 = sorted([x1, x2]); y1, y2 = sorted([y1, y2])
            dx1, dy1 = to_disp(x1, y1); dx2, dy2 = to_disp(x2, y2)
            for off in range(2):
                draw.rectangle([dx1 - off, dy1 - off, dx2 + off, dy2 + off], outline=color, width=2)
            for px, py in [(dx1, dy1), (dx2, dy2)]:
                draw.ellipse([px - 5, py - 5, px + 5, py + 5], fill=color)

    return img

# =========================
# 메인 UI
# =========================
def main():
    ensure_state()

    st.markdown("""
    <h1 style="text-align:center;margin-bottom:0.5rem;">📄 수입신고필증 PDF → 엑셀 (전 항목 ROI 템플릿)</h1>
    <p style="text-align:center;color:#555;">한 번 좌표 지정 → 계속 사용(자동 로드). 필요하면 언제든 템플릿 전환.</p>
    """, unsafe_allow_html=True)

    # 템플릿 관리
    with st.expander("🧩 템플릿 관리", expanded=True):
        c0, c1, c2, c3, c4 = st.columns([1.2, 2, 1, 1, 1])
        with c0:
            st.checkbox("현재 템플릿 고정 사용", key="lock_template",
                        help="체크 시 ROI 재지정 섹션을 건너뛰고 바로 변환에 사용")
        with c1:
            st.text_input("템플릿 이름", key="template_name", placeholder="예) UNIPASS_2025_v1")
        with c2:
            if st.button("💾 현재 좌표 저장", use_container_width=True, type="primary"):
                name = st.session_state.template_name.strip()
                if not name:
                    st.warning("템플릿 이름을 입력하세요.")
                elif len(st.session_state.norm_rects) < len(FIELDS):
                    st.warning("모든 필드 좌표를 먼저 지정하세요.")
                else:
                    tmpls = st.session_state.all_templates
                    tmpls[name] = {
                        "created_at": datetime.now().isoformat(),
                        "dpi": st.session_state.tmpl_dpi,
                        "norm_rects": st.session_state.norm_rects,
                    }
                    set_last_used(tmpls, name)
                    st.success(f"저장 & 마지막 사용 지정: {name}")

        with c3:
            up = st.file_uploader("가져오기(JSON)", type=["json"], key="tmpl_upload")
            if up is not None:
                try:
                    data = json.loads(up.read().decode("utf-8"))
                    st.session_state.tmpl_dpi = data.get("dpi", DPI_DEFAULT)
                    st.session_state.norm_rects = data.get("norm_rects", {})
                    if st.session_state.template_name.strip():
                        name = st.session_state.template_name.strip()
                        tmpls = st.session_state.all_templates
                        tmpls[name] = {
                            "created_at": datetime.now().isoformat(),
                            "dpi": st.session_state.tmpl_dpi,
                            "norm_rects": st.session_state.norm_rects,
                        }
                        set_last_used(tmpls, name)
                        st.success(f"가져온 좌표를 '{name}' 이름으로 저장 & 사용")
                    else:
                        st.info("좌측에 템플릿 이름을 입력하면 가져온 좌표를 바로 저장할 수 있어요.")
                except Exception as e:
                    st.error(f"가져오기 실패: {e}")

        with c4:
            if st.session_state.norm_rects:
                export_data = {
                    "created_at": datetime.now().isoformat(),
                    "dpi": st.session_state.tmpl_dpi,
                    "norm_rects": st.session_state.norm_rects,
                }
                st.download_button(
                    "⬇️ 내보내기",
                    data=json.dumps(export_data, ensure_ascii=False, indent=2).encode("utf-8"),
                    file_name=f"receipt_template_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
                    mime="application/json",
                    use_container_width=True
                )

        tmpls = st.session_state.all_templates
        names = sorted([n for n in tmpls.keys() if n != "__meta"])
        last_used = get_last_used(tmpls)

        cL, cM, cR, cD = st.columns([2, 1, 1, 1])
        with cL:
            sel = st.selectbox("불러올 템플릿", options=["(선택 없음)"] + names,
                               index=(names.index(last_used) + 1) if last_used in names else 0)
        with cM:
            if st.button("📂 불러오기", use_container_width=True):
                if sel != "(선택 없음)":
                    data = tmpls.get(sel)
                    st.session_state.template_name = sel
                    st.session_state.tmpl_dpi = data.get("dpi", DPI_DEFAULT)
                    st.session_state.norm_rects = data.get("norm_rects", {})
                    st.session_state.current_field_idx = 0
                    st.session_state.click_phase = 0
                    st.session_state.temp_points = []
                    set_last_used(tmpls, sel)
                    st.success(f"불러오기 & 마지막 사용 지정: {sel}")
                else:
                    st.info("불러올 템플릿을 선택하세요.")
        with cR:
            if st.button("🗑️ 삭제", use_container_width=True):
                if sel != "(선택 없음)" and sel in tmpls:
                    del tmpls[sel]
                    save_all_templates(tmpls)
                    if st.session_state.template_name == sel:
                        st.session_state.template_name = ""
                        st.session_state.norm_rects = {}
                    st.success(f"삭제 완료: {sel}")
                else:
                    st.info("삭제할 템플릿을 선택하세요.")
        with cD:
            if st.button("⭐ 마지막 사용으로 지정", use_container_width=True):
                if sel != "(선택 없음)":
                    set_last_used(tmpls, sel)
                    st.success(f"이 템플릿을 다음에도 자동 사용: {sel}")
                else:
                    st.info("지정할 템플릿을 선택하세요.")

        st.caption("※ '마지막 사용'으로 지정된 템플릿은 앱을 다시 켜도 자동 적용됩니다.")

    st.markdown("---")

    # 파일 업로드
    files = st.file_uploader("📎 PDF 업로드 (대표 1개 + 배치 여러 개 가능)", type=["pdf"], accept_multiple_files=True)

    # 좌표 지정 (템플릿 고정 사용 해제 시에만)
    if not st.session_state.lock_template:
        st.markdown("### 🎯 좌표 지정 (대표 PDF 1페이지 기준)")
        st.caption("필드 순서: " + " → ".join(FIELDS))

        if files:
            rep = files[0]
            rep_bytes = rep.getvalue() if hasattr(rep, "getvalue") else rep.read()
            img, w, h, page_rect = pdf_first_page_pix(rep_bytes, dpi=st.session_state.tmpl_dpi)

            disp_w = st.slider("표시 너비(px)", min_value=600, max_value=1600,
                               value=st.session_state.display_width, step=50)
            st.session_state.display_width = disp_w
            ratio = disp_w / w
            img_resized = img.resize((disp_w, int(h * ratio)))

            done_cnt = sum(1 for f in FIELDS if f in st.session_state.norm_rects)
            st.progress(done_cnt / len(FIELDS))
            st.write(f"완료 {done_cnt}/{len(FIELDS)}")

            current_field = FIELDS[st.session_state.current_field_idx] if st.session_state.current_field_idx < len(FIELDS) else None
            if current_field:
                st.info(f"🖱️ {current_field} 영역을 지정하세요 — 먼저 **좌상단**, 다음 **우하단**")
            else:
                st.success("✅ 모든 필드 좌표 지정 완료!")

            overlay = render_with_overlays(
                img_resized, w_orig=w, h_orig=h, ratio=ratio,
                norm_rects=st.session_state.norm_rects,
                temp_points=st.session_state.temp_points[:],
                current_field=current_field
            )

            clicked = streamlit_image_coordinates(
                overlay,
                key=f"coord_{st.session_state.current_field_idx}_{st.session_state.click_phase}"
            )

            if clicked and current_field:
                ox = clicked["x"] / ratio
                oy = clicked["y"] / ratio
                if st.session_state.click_phase == 0:
                    st.session_state.temp_points = [(ox, oy)]
                    st.session_state.click_phase = 1
                    st.toast("좌상단 기록!")
                else:
                    (x1, y1) = st.session_state.temp_points[0]
                    x2, y2 = ox, oy
                    x1, x2 = sorted([x1, x2]); y1, y2 = sorted([y1, y2])
                    xn1, yn1 = x1 / w, y1 / h
                    xn2, yn2 = x2 / w, y2 / h
                    st.session_state.norm_rects[current_field] = [xn1, yn1, xn2, yn2]
                    st.session_state.temp_points = []
                    st.session_state.click_phase = 0
                    st.session_state.current_field_idx += 1
                    st.toast(f"{current_field} 좌표 저장!")

            colA, colB, colC, colD = st.columns(4)
            with colA:
                if st.button("⏮ 이전 필드", use_container_width=True):
                    st.session_state.current_field_idx = max(0, st.session_state.current_field_idx - 1)
                    st.session_state.click_phase = 0
                    st.session_state.temp_points = []
            with colB:
                if st.button("⏭ 다음 필드", use_container_width=True):
                    st.session_state.current_field_idx = min(len(FIELDS) - 1, st.session_state.current_field_idx + 1)
                    st.session_state.click_phase = 0
                    st.session_state.temp_points = []
            with colC:
                if st.button("🧹 현재 필드 좌표 삭제", use_container_width=True):
                    if current_field and current_field in st.session_state.norm_rects:
                        del st.session_state.norm_rects[current_field]
                        st.session_state.temp_points = []
                        st.session_state.click_phase = 0
                        st.toast(f"{current_field} 좌표 삭제")
            with colD:
                if st.button("🔁 전체 좌표 초기화", use_container_width=True):
                    st.session_state.norm_rects = {}
                    st.session_state.current_field_idx = 0
                    st.session_state.click_phase = 0
                    st.session_state.temp_points = []
                    st.success("전체 좌표 초기화 완료")

            if st.session_state.norm_rects:
                st.markdown("#### 📋 저장된 좌표(정규화)")
                rows = []
                for k in FIELDS:
                    rect = st.session_state.norm_rects.get(k)
                    if rect:
                        xn1, yn1, xn2, yn2 = rect
                        rows.append({"필드": k, "x1": round(xn1, 4), "y1": round(yn1, 4),
                                     "x2": round(xn2, 4), "y2": round(yn2, 4)})
                st.dataframe(pd.DataFrame(rows), use_container_width=True)
        else:
            st.info("대표 PDF를 포함해 파일을 업로드하세요.")
    else:
        st.info("현재 설정: '템플릿 고정 사용' — 좌표 지정 섹션을 생략하고 바로 변환합니다.")

    st.markdown("---")

    # 변환 실행
    if files and st.button("🚀 변환 시작", type="primary", use_container_width=True):
        if not st.session_state.norm_rects or any(f not in st.session_state.norm_rects for f in FIELDS):
            st.error("모든 필드의 좌표가 지정되지 않았어요. '템플릿 고정 사용'을 끄고 ROI를 먼저 지정/저장하세요.")
            st.stop()

        rows, issues = [], []
        for f in files:
            try:
                f_bytes = f.getvalue() if hasattr(f, "getvalue") else f.read()
                _, _, _, page_rect = pdf_first_page_pix(f_bytes, dpi=st.session_state.tmpl_dpi)

                data = {}
                for name in FIELDS:
                    rect = st.session_state.norm_rects[name]
                    raw = clip_text_by_norm_rect(f_bytes, rect, page_rect)
                    data[name] = postprocess_field(name, raw)

                if data["신고일"] and not re.match(r'^\d{4}/\d{2}/\d{2}$', data["신고일"]):
                    issues.append(f"⚠️ {getattr(f,'name','파일')} : 신고일 형식 확인 → {data['신고일']}")
                for k in ["환율", "부가가치세 과표", "관세", "부가가치세"]:
                    if data[k] is None or data[k] == "":
                        issues.append(f"⚠️ {getattr(f,'name','파일')} : {k} 인식 실패/형식 오류")

                rows.append(data)

            except Exception as e:
                issues.append(f"❌ {getattr(f,'name','파일')} 처리 오류: {e}")

        if not rows:
            st.error("변환 가능한 결과가 없습니다.")
            if issues:
                st.warning("\n".join(issues))
            st.stop()

        df = pd.DataFrame(rows, columns=FIELDS)

        dup_mask = df["b/l(awb)번호"].duplicated(keep=False)
        if dup_mask.any():
            dupped = df.loc[dup_mask, "b/l(awb)번호"].unique().tolist()
            st.warning(f"⚠️ 동일 B/L 번호 중복: {', '.join(dupped)}")

        def to_date(s):
            try:
                return datetime.strptime(s, "%Y/%m/%d")
            except Exception:
                return datetime.max
        df = df.sort_values(by="신고일", key=lambda col: col.map(to_date))

        st.markdown("### ✅ 변환 결과")
        view = df.copy()
        view["환율"] = view["환율"].map(lambda x: f"{x:.4f}" if isinstance(x, (int, float)) else "")
        for k in ["부가가치세 과표", "관세", "부가가치세"]:
            view[k] = view[k].map(lambda x: f"{int(x):,}" if pd.notnull(x) and x != "" else "")
        st.dataframe(view, use_container_width=True)

        if issues:
            st.markdown("### 🔎 점검 결과")
            for line in issues:
                st.write(line)

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            df.to_excel(writer, sheet_name="results", index=False)
            wb = writer.book
            ws = writer.sheets["results"]
            money_fmt = wb.add_format({'num_format': '#,##0'})
            fx_fmt = wb.add_format({'num_format': '0.0000'})
            date_fmt = wb.add_format({'num_format': 'yyyy/mm/dd'})
            col_idx = {c: i for i, c in enumerate(df.columns)}
            if "환율" in col_idx:
                ws.set_column(col_idx["환율"], col_idx["환율"], 12, fx_fmt)
            for k in ["부가가치세 과표", "관세", "부가가치세"]:
                if k in col_idx:
                    ws.set_column(col_idx[k], col_idx[k], 14, money_fmt)
            if "신고일" in col_idx:
                ws.set_column(col_idx["신고일"], col_idx["신고일"], 12, date_fmt)
            ws.set_column(0, len(df.columns) - 1, 16)

        st.download_button(
            "⬇️ 엑셀(.xlsx) 다운로드",
            data=buffer.getvalue(),
            file_name=f"수입신고필증_추출_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    st.caption("ⓘ 국내도착항은 항 코드(KRPTK 등)를 무시하고 한글 지명(예: 평택항/인천공항 등)만 추출합니다. 템플릿은 '마지막 사용' 지정 시 자동 적용됩니다.")

if __name__ == "__main__":
    main()
