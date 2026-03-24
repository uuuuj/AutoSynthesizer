"""
masking_gui.py
Excel Key-Value 데이터 마스킹 프로그램

[설명]
    A열 = Key, B열 = Value 구조의 Excel 파일에서
    Value를 랜덤 값으로 마스킹하여 새 Excel 파일로 저장한다.

[실행 전 설치]
    pip install pandas numpy openpyxl
    (선택) pip install xlwings    ← DRM 해제 필요 시

[사용법]
    python masking_gui.py

[GUI 구성]
    ① 엑셀 파일 선택 (xlwings 우선, openpyxl 자동 폴백)
    ② 데이터 미리보기 + 마스킹 대상 확인
    ③ 저장 경로 / 파일명 설정
    ④ 마스킹 실행 → 진행 로그 실시간 표시
"""

import os
import sys
import json
import string
import warnings
import threading
import re
from datetime import datetime, timedelta

import numpy as np
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

# PIL — 로고 이미지 리사이즈용 (없으면 로고 생략)
HAS_PIL = False
try:
    from PIL import Image, ImageTk
    HAS_PIL = True
except ImportError:
    pass

warnings.filterwarnings('ignore')


def _resource_path(filename: str) -> str:
    """PyInstaller EXE에서도 동작하는 리소스 경로 반환.

    Args:
        filename (str): 리소스 파일명

    Returns:
        str: 절대 경로
    """
    if getattr(sys, 'frozen', False):
        base = sys._MEIPASS
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, filename)


# xlwings는 선택사항 — 없거나 Excel 미설치면 openpyxl로 폴백
HAS_XLWINGS = False
try:
    import xlwings as xw
    HAS_XLWINGS = True
except (ImportError, OSError, Exception):
    pass

try:
    import openpyxl
except ImportError:
    if not HAS_XLWINGS:
        raise ImportError("xlwings 또는 openpyxl 중 하나가 필요합니다.\n"
                          "  pip install openpyxl   또는   pip install xlwings")


# ══════════════════════════════════════════════════════════════
# Excel 로드 — xlwings 우선, 실패 시 openpyxl/pandas 폴백
# ══════════════════════════════════════════════════════════════

def _get_sheet_names_xlwings(excel_path: str) -> list:
    """xlwings로 시트 이름 목록을 읽는다.

    Args:
        excel_path (str): Excel 파일 경로

    Returns:
        list: 시트 이름 리스트
    """
    app, wb = None, None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        wb = app.books.open(os.path.abspath(excel_path))
        return [s.name for s in wb.sheets]
    finally:
        if wb:
            wb.close()
        if app:
            app.quit()


def _load_excel_xlwings(excel_path: str, sheet_name: str = None) -> tuple:
    """xlwings로 Excel 파일을 로드한다.

    Args:
        excel_path (str): Excel 파일 경로
        sheet_name (str): 시트 이름 (None이면 첫 번째 시트)

    Returns:
        tuple: (DataFrame, info dict)
    """
    app, wb, close_after = None, None, False
    try:
        abs_path = os.path.abspath(excel_path)
        if not os.path.exists(abs_path):
            raise FileNotFoundError(f"파일 없음: {abs_path}")
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        wb = app.books.open(abs_path)
        close_after = True

        ws = wb.sheets[sheet_name] if sheet_name else wb.sheets[0]
        raw_data = ws.used_range.value
        if raw_data is None:
            raise ValueError(f"시트 '{ws.name}' 데이터 없음")
        if not isinstance(raw_data[0], list):
            raw_data = [raw_data]

        headers = [h if h is not None else f'col_{i}' for i, h in enumerate(raw_data[0])]
        df = pd.DataFrame(raw_data[1:], columns=headers)

        info = {
            'file_name': wb.name,
            'file_path': excel_path or wb.fullname,
            'sheet_name': ws.name,
            'all_sheets': [s.name for s in wb.sheets],
            'rows': len(df),
            'cols': len(df.columns),
        }
        return df, info
    finally:
        if close_after:
            if wb:
                wb.close()
            if app:
                app.quit()


def _get_sheet_names_openpyxl(excel_path: str) -> list:
    """openpyxl로 시트 이름 목록을 읽는다.

    Args:
        excel_path (str): Excel 파일 경로

    Returns:
        list: 시트 이름 리스트
    """
    abs_path = os.path.abspath(excel_path)
    if not os.path.exists(abs_path):
        raise FileNotFoundError(f"파일 없음: {abs_path}")
    wb = openpyxl.load_workbook(abs_path, read_only=True, data_only=True)
    names = wb.sheetnames
    wb.close()
    return names


def _load_excel_openpyxl(excel_path: str, sheet_name: str = None) -> tuple:
    """openpyxl + pandas로 Excel 파일을 로드한다.

    Args:
        excel_path (str): Excel 파일 경로
        sheet_name (str): 시트 이름 (None이면 첫 번째 시트)

    Returns:
        tuple: (DataFrame, info dict)
    """
    abs_path = os.path.abspath(excel_path)
    if not os.path.exists(abs_path):
        raise FileNotFoundError(f"파일 없음: {abs_path}")

    file_name = os.path.basename(abs_path)
    xls = pd.ExcelFile(abs_path, engine='openpyxl')
    all_sheets = xls.sheet_names

    target_sheet = sheet_name if sheet_name else all_sheets[0]
    df = pd.read_excel(xls, sheet_name=target_sheet, engine='openpyxl')
    xls.close()

    info = {
        'file_name': file_name,
        'file_path': abs_path,
        'sheet_name': target_sheet,
        'all_sheets': all_sheets,
        'rows': len(df),
        'cols': len(df.columns),
    }
    return df, info


def get_sheet_names(excel_path: str) -> tuple:
    """시트 이름 목록 반환. xlwings → openpyxl 순서로 시도.

    Args:
        excel_path (str): Excel 파일 경로

    Returns:
        tuple: (시트 이름 리스트, 사용된 엔진 이름)
    """
    errors = []
    if HAS_XLWINGS:
        try:
            return _get_sheet_names_xlwings(excel_path), 'xlwings'
        except Exception as e:
            errors.append(f"xlwings: {e}")
    try:
        return _get_sheet_names_openpyxl(excel_path), 'openpyxl'
    except Exception as e:
        errors.append(f"openpyxl: {e}")
    raise RuntimeError("시트 목록 로드 실패:\n" + "\n".join(errors))


def load_excel(excel_path: str, sheet_name: str = None) -> tuple:
    """Excel 로드. xlwings → openpyxl 순서로 시도.

    Args:
        excel_path (str): Excel 파일 경로
        sheet_name (str): 시트 이름 (None이면 첫 번째 시트)

    Returns:
        tuple: ((DataFrame, info dict), 사용된 엔진 이름)
    """
    errors = []
    if HAS_XLWINGS:
        try:
            return _load_excel_xlwings(excel_path, sheet_name), 'xlwings'
        except Exception as e:
            errors.append(f"xlwings: {e}")
    try:
        return _load_excel_openpyxl(excel_path, sheet_name), 'openpyxl'
    except Exception as e:
        errors.append(f"openpyxl: {e}")
    raise RuntimeError("파일 로드 실패:\n" + "\n".join(errors))


# ══════════════════════════════════════════════════════════════
# 마스킹 로직 — 타입별 랜덤 값 생성
# ══════════════════════════════════════════════════════════════

def _detect_value_type(value: str) -> str:
    """값의 타입을 감지하여 적절한 마스킹 전략을 결정한다.

    Args:
        value (str): 검사할 값 (문자열)

    Returns:
        str: 'number', 'date', 'email', 'phone', 'text' 중 하나
    """
    s = str(value).strip()

    # 숫자
    try:
        float(s.replace(',', ''))
        return 'number'
    except (ValueError, AttributeError):
        pass

    # 날짜 패턴
    date_patterns = [
        r'^\d{4}[-/\.]\d{1,2}[-/\.]\d{1,2}',  # 2024-03-15
        r'^\d{4}\s*년\s*\d{1,2}\s*월',  # 2024년 3월
    ]
    for pat in date_patterns:
        if re.match(pat, s):
            return 'date'

    # 이메일
    if re.match(r'^[^@\s]+@[^@\s]+\.[^@\s]+$', s):
        return 'email'

    # 전화번호
    if re.match(r'^[\d\-\+\(\)\s]{8,}$', s):
        return 'phone'

    return 'text'


def _mask_number(value: str, rng: np.random.Generator) -> str:
    """숫자 값을 랜덤 마스킹한다.

    Args:
        value (str): 원본 숫자 문자열
        rng (np.random.Generator): 난수 생성기

    Returns:
        str: 마스킹된 숫자 문자열
    """
    s = str(value).strip()
    has_comma = ',' in s
    clean = s.replace(',', '')

    try:
        num = float(clean)
    except ValueError:
        return s

    is_int = '.' not in clean
    if num == 0:
        result = rng.integers(1, 100)
    else:
        # 원본의 10%~200% 범위에서 랜덤
        low = abs(num) * 0.1
        high = abs(num) * 2.0
        result = rng.uniform(low, high)
        if num < 0:
            result = -result

    if is_int:
        result = int(round(result))
        result_str = str(result)
        if has_comma:
            result_str = f"{result:,}"
    else:
        # 소수점 자릿수 보존
        dec_places = len(clean.split('.')[1]) if '.' in clean else 2
        result = round(result, dec_places)
        result_str = f"{result:.{dec_places}f}"
        if has_comma:
            parts = result_str.split('.')
            parts[0] = f"{int(parts[0]):,}"
            result_str = '.'.join(parts)

    return result_str


def _mask_date(value: str, rng: np.random.Generator) -> str:
    """날짜 값을 랜덤 마스킹한다.

    Args:
        value (str): 원본 날짜 문자열
        rng (np.random.Generator): 난수 생성기

    Returns:
        str: 마스킹된 날짜 문자열
    """
    s = str(value).strip()
    # 연-월-일 패턴
    m = re.match(r'(\d{4})([-/\.])(\d{1,2})\2(\d{1,2})', s)
    if m:
        sep = m.group(2)
        y = rng.integers(2020, 2026)
        mo = rng.integers(1, 13)
        d = rng.integers(1, 29)
        return f"{y}{sep}{mo:02d}{sep}{d:02d}"

    # 한글 날짜
    m = re.match(r'(\d{4})\s*년\s*(\d{1,2})\s*월\s*(\d{1,2})\s*일', s)
    if m:
        y = rng.integers(2020, 2026)
        mo = rng.integers(1, 13)
        d = rng.integers(1, 29)
        return f"{y}년 {mo}월 {d}일"

    m = re.match(r'(\d{4})\s*년\s*(\d{1,2})\s*월', s)
    if m:
        y = rng.integers(2020, 2026)
        mo = rng.integers(1, 13)
        return f"{y}년 {mo}월"

    return s


_MASK_DOMAINS = ['example.com', 'test.org', 'masked.net', 'dummy.co.kr', 'sample.io']

_MASK_LAST_NAMES = ['김', '이', '박', '최', '정', '강', '윤', '임', '한', '오',
                    '신', '홍', '문', '류', '배', '전', '조', '남', '서', '권']
_MASK_FIRST_PARTS = ['민', '서', '지', '수', '도', '하', '준', '소', '예', '태',
                     '재', '채', '나', '기', '윤', '성', '우', '세', '진', '아']
_MASK_LAST_PARTS = ['준', '연', '호', '아', '서', '윤', '혁', '율', '훈', '린',
                    '민', '원', '영', '현', '양', '나', '재', '진', '수', '은']

# 일반 한글 텍스트 마스킹용 단어 풀
_MASK_NOUNS = [
    '사과', '바다', '하늘', '구름', '산호', '별빛', '은하', '초원', '강물', '숲길',
    '노을', '안개', '무지개', '태양', '달빛', '소나무', '대나무', '매화', '진달래', '장미',
    '호수', '계곡', '폭포', '섬마을', '항구', '등대', '해변', '갯벌', '모래밭', '언덕',
    '기린', '코끼리', '사슴', '토끼', '고래', '돌고래', '독수리', '참새', '까치', '두루미',
]
_MASK_PLACES = [
    '서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종',
    '수원', '창원', '포항', '천안', '전주', '청주', '제주', '춘천',
    '여수', '순천', '김해', '양산', '진주', '통영', '거창', '함양',
]
_MASK_DISTRICTS = [
    '중앙동', '태평동', '신정동', '도화동', '용산동', '남산동', '북문동', '서문동',
    '장평동', '명동', '수정동', '옥포동', '신현동', '아주동', '상문동', '하청면',
]
_MASK_POSITIONS = [
    '수석연구원', '선임연구원', '책임연구원', '주임연구원', '전임연구원',
    '부장', '차장', '과장', '대리', '사원', '팀장', '실장', '센터장',
]
_MASK_DEPTS = [
    '기술연구소', '품질관리팀', '생산관리팀', '설계기술팀', '공정기술팀',
    '전략기획팀', '인사관리팀', '재무회계팀', '영업지원팀', '안전환경팀',
]


def _mask_email(value: str, rng: np.random.Generator) -> str:
    """이메일 값을 랜덤 마스킹한다.

    Args:
        value (str): 원본 이메일 문자열
        rng (np.random.Generator): 난수 생성기

    Returns:
        str: 마스킹된 이메일 문자열
    """
    user_len = rng.integers(5, 12)
    chars = string.ascii_lowercase + string.digits
    user = ''.join(rng.choice(list(chars)) for _ in range(user_len))
    domain = _MASK_DOMAINS[rng.integers(len(_MASK_DOMAINS))]
    return f"{user}@{domain}"


def _mask_phone(value: str, rng: np.random.Generator) -> str:
    """전화번호 값을 랜덤 마스킹한다.

    Args:
        value (str): 원본 전화번호 문자열
        rng (np.random.Generator): 난수 생성기

    Returns:
        str: 마스킹된 전화번호 문자열
    """
    s = str(value).strip()
    # 한국 휴대폰 패턴
    if re.match(r'^01[016789][-\s]?\d{3,4}[-\s]?\d{4}$', s):
        mid = rng.integers(1000, 10000)
        end = rng.integers(1000, 10000)
        return f"010-{mid}-{end}"
    # 일반 전화
    if '-' in s:
        parts = s.split('-')
        masked_parts = [parts[0]]
        for p in parts[1:]:
            n_digits = len(p)
            masked_parts.append(str(rng.integers(10**(n_digits-1), 10**n_digits)))
        return '-'.join(masked_parts)
    # 그 외 숫자 전화번호
    digits = re.sub(r'\D', '', s)
    new_digits = ''.join(str(rng.integers(0, 10)) for _ in digits)
    return new_digits


def _mask_text(value: str, rng: np.random.Generator) -> str:
    """텍스트 값을 랜덤 마스킹한다.

    Args:
        value (str): 원본 텍스트 문자열
        rng (np.random.Generator): 난수 생성기

    Returns:
        str: 마스킹된 텍스트 문자열
    """
    s = str(value).strip()

    # 한글 이름 패턴 (2~4글자 한글)
    if re.match(r'^[가-힣]{2,4}$', s):
        name = (_MASK_LAST_NAMES[rng.integers(len(_MASK_LAST_NAMES))]
                + _MASK_FIRST_PARTS[rng.integers(len(_MASK_FIRST_PARTS))]
                + _MASK_LAST_PARTS[rng.integers(len(_MASK_LAST_PARTS))])
        return name

    # ID 패턴 (영문+숫자 또는 영문+구분자+숫자)
    m = re.match(r'^([A-Za-z]+)([\-_]?)(\d+)$', s)
    if m:
        prefix = m.group(1)
        sep = m.group(2)
        num_len = len(m.group(3))
        new_num = rng.integers(10**(num_len-1), 10**num_len)
        new_prefix = ''.join(rng.choice(list(string.ascii_uppercase)) for _ in prefix)
        return f"{new_prefix}{sep}{new_num}"

    # 주소 패턴 (시/도 + 구/군/시 + 동/면/리 + 번지)
    if re.search(r'[시도군구읍면동리로길]\s', s) or re.search(r'\d+[-−]\d+', s):
        city = _MASK_PLACES[rng.integers(len(_MASK_PLACES))]
        district = _MASK_DISTRICTS[rng.integers(len(_MASK_DISTRICTS))]
        num1 = rng.integers(1, 500)
        num2 = rng.integers(1, 30)
        return f"{city}시 {district} {num1}-{num2}"

    # 직급/직위 패턴
    for pos in _MASK_POSITIONS:
        if pos in s or s in _MASK_POSITIONS:
            return _MASK_POSITIONS[rng.integers(len(_MASK_POSITIONS))]

    # 부서 패턴
    if re.search(r'(센터|팀|부|실|소|과)$', s):
        return _MASK_DEPTS[rng.integers(len(_MASK_DEPTS))]

    # 일반 한글 텍스트 — 단어 풀에서 조합
    if re.search(r'[가-힣]', s):
        words = s.split()
        if len(words) <= 1:
            # 한 단어 → 단어 풀에서 랜덤 선택
            return _MASK_NOUNS[rng.integers(len(_MASK_NOUNS))]
        else:
            # 여러 단어 → 같은 개수만큼 단어 풀에서 조합
            result_words = []
            for _ in words:
                result_words.append(_MASK_NOUNS[rng.integers(len(_MASK_NOUNS))])
            return ' '.join(result_words)

    # 영문 텍스트 — 같은 길이의 랜덤 영문
    result = []
    for ch in s:
        if ch.isupper():
            result.append(chr(rng.integers(ord('A'), ord('Z') + 1)))
        elif ch.islower():
            result.append(chr(rng.integers(ord('a'), ord('z') + 1)))
        elif ch.isdigit():
            result.append(str(rng.integers(0, 10)))
        else:
            result.append(ch)
    return ''.join(result)


def mask_value(value: str, rng: np.random.Generator) -> str:
    """값의 타입을 감지하고 적절한 마스킹을 적용한다.

    Args:
        value (str): 원본 값
        rng (np.random.Generator): 난수 생성기

    Returns:
        str: 마스킹된 값
    """
    vtype = _detect_value_type(value)
    if vtype == 'number':
        return _mask_number(value, rng)
    elif vtype == 'date':
        return _mask_date(value, rng)
    elif vtype == 'email':
        return _mask_email(value, rng)
    elif vtype == 'phone':
        return _mask_phone(value, rng)
    else:
        return _mask_text(value, rng)


# ══════════════════════════════════════════════════════════════
# GUI 클래스
# ══════════════════════════════════════════════════════════════

class MaskingApp:
    """Excel Key-Value 데이터 마스킹 GUI 애플리케이션."""

    def __init__(self, root: tk.Tk) -> None:
        """MaskingApp을 초기화한다.

        Args:
            root (tk.Tk): Tkinter 루트 윈도우

        Returns:
            None
        """
        self.root = root
        self.root.title("데이터 마스킹 도구  —  삼성중공업 생산 DT센터")
        self.root.geometry("1000x850")
        self.root.minsize(900, 750)

        # 창 아이콘 설정
        try:
            ico_path = _resource_path("synth_ico.ico")
            if os.path.exists(ico_path):
                self.root.iconbitmap(ico_path)
        except Exception:
            pass

        self.df = None
        self.info = None
        self._running = False

        self._current_step = 0
        self._blink_state = True
        self._blink_id = None
        self.step_labels = []
        self.step_arrows = []

        self.excel_path = tk.StringVar()
        self.sheet_var = tk.StringVar()
        self.save_dir = tk.StringVar()
        self.save_name = tk.StringVar(value="마스킹결과")
        self.key_col_var = tk.StringVar(value="A열 (첫 번째 컬럼)")
        self.val_col_var = tk.StringVar(value="B열 (두 번째 컬럼)")
        self.seed_var = tk.StringVar(value="42")

        self._build_ui()
        self._set_step(0)

    def _build_ui(self) -> None:
        """GUI 전체 레이아웃을 구성한다.

        Args:
            None

        Returns:
            None
        """
        style = ttk.Style()
        style.configure("Sub.TLabel", font=("맑은 고딕", 9), foreground="#666")
        style.configure("ColHeader.TLabel", font=("맑은 고딕", 10, "bold"), foreground="#333")
        style.configure("Run.TButton", font=("맑은 고딕", 11, "bold"))

        # ── 로고 배너 ──
        self._logo_image = None
        try:
            logo_path = _resource_path("logo.png")
            if os.path.exists(logo_path) and HAS_PIL:
                img = Image.open(logo_path)
                target_w = 990
                ratio = target_w / img.width
                target_h = int(img.height * ratio)
                img = img.resize((target_w, target_h), Image.LANCZOS)
                self._logo_image = ImageTk.PhotoImage(img)
                logo_lbl = tk.Label(self.root, image=self._logo_image, bg="#2d4a7a")
                logo_lbl.pack(fill=tk.X, padx=5, pady=(5, 0))
        except Exception:
            pass

        # ── 단계 안내 표시 바 ──
        self._build_step_indicator(self.root)

        # ══ ① 엑셀 파일 선택 ══
        sec1 = ttk.LabelFrame(self.root,
                               text="  ① 엑셀 파일 선택  (xlwings 우선 / openpyxl 폴백)  ",
                               padding=8)
        sec1.pack(fill=tk.X, padx=5, pady=(0, 5))

        r1 = ttk.Frame(sec1)
        r1.pack(fill=tk.X)
        ttk.Label(r1, text="파일 경로:", width=10).pack(side=tk.LEFT)
        ttk.Entry(r1, textvariable=self.excel_path, font=("Consolas", 9)).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        ttk.Button(r1, text="찾아보기...", command=self._browse_file).pack(side=tk.LEFT)

        r2 = ttk.Frame(sec1)
        r2.pack(fill=tk.X, pady=(5, 0))
        ttk.Label(r2, text="시트 선택:", width=10).pack(side=tk.LEFT)
        self.sheet_combo = ttk.Combobox(r2, textvariable=self.sheet_var,
                                         state='readonly', width=25)
        self.sheet_combo.pack(side=tk.LEFT, padx=(0, 20))
        ttk.Label(r2, text="시드값:").pack(side=tk.LEFT)
        ttk.Entry(r2, textvariable=self.seed_var, width=8).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Label(r2, text="(같은 시드 → 같은 마스킹 결과)", style="Sub.TLabel").pack(side=tk.LEFT)

        ttk.Button(sec1, text="📊  파일 분석", command=self._analyze_file).pack(pady=(7, 0))

        # ══ ② 데이터 미리보기 ══
        sec2 = ttk.LabelFrame(self.root, text="  ② 데이터 미리보기  ", padding=8)
        sec2.pack(fill=tk.BOTH, expand=True, padx=5, pady=(0, 5))

        # 분석 요약
        self.analysis_text = tk.Text(sec2, height=3, font=("Consolas", 9),
                                      bg="#f5f5f0", state=tk.DISABLED, wrap=tk.WORD)
        self.analysis_text.pack(fill=tk.X, pady=(0, 5))

        # 미리보기 테이블 (Treeview)
        tree_frame = ttk.Frame(sec2)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        self.tree = ttk.Treeview(tree_frame, columns=("key", "value"), show="headings",
                                  height=12)
        self.tree.heading("key", text="Key (A열)")
        self.tree.heading("value", text="Value (B열)")
        self.tree.column("key", width=300, anchor="w")
        self.tree.column("value", width=400, anchor="w")

        tree_sb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=tree_sb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_sb.pack(side=tk.RIGHT, fill=tk.Y)

        # ══ ③ 저장 설정 ══
        sec3 = ttk.LabelFrame(self.root, text="  ③ 저장 설정  ", padding=8)
        sec3.pack(fill=tk.X, padx=5, pady=(0, 5))

        rs1 = ttk.Frame(sec3)
        rs1.pack(fill=tk.X)
        ttk.Label(rs1, text="저장 경로:", width=10).pack(side=tk.LEFT)
        ttk.Entry(rs1, textvariable=self.save_dir, font=("Consolas", 9)).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        ttk.Button(rs1, text="폴더 선택...", command=self._browse_dir).pack(side=tk.LEFT)

        rs2 = ttk.Frame(sec3)
        rs2.pack(fill=tk.X, pady=(5, 0))
        ttk.Label(rs2, text="파일 이름:", width=10).pack(side=tk.LEFT)
        ttk.Entry(rs2, textvariable=self.save_name, width=35, font=("Consolas", 9)).pack(
            side=tk.LEFT, padx=(0, 6))
        ttk.Label(rs2, text=".xlsx / _마스킹키.json 자동 생성", style="Sub.TLabel").pack(side=tk.LEFT)

        # ══ ④ 실행 로그 ══
        sec4 = ttk.LabelFrame(self.root, text="  ④ 실행 로그  ", padding=8)
        sec4.pack(fill=tk.BOTH, expand=False, padx=5, pady=(0, 5))

        bf = ttk.Frame(sec4)
        bf.pack(fill=tk.X, pady=(0, 5))
        self.run_btn = ttk.Button(bf, text="▶  마스킹 실행",
                                   command=self._run, style="Run.TButton",
                                   state='disabled')
        self.run_btn.pack(side=tk.LEFT)
        self.progress = ttk.Progressbar(bf, mode='determinate', maximum=100, length=200)
        self.progress.pack(side=tk.LEFT, padx=(12, 0))
        self.status_lbl = ttk.Label(bf, text="", style="Sub.TLabel")
        self.status_lbl.pack(side=tk.LEFT, padx=(12, 0))

        self.log_text = scrolledtext.ScrolledText(sec4, height=6, font=("Consolas", 9),
                                                    state=tk.DISABLED, wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True)

    # ══════════════════════════════════════════════════════════
    # 단계 안내 표시 바
    # ══════════════════════════════════════════════════════════

    def _build_step_indicator(self, parent: tk.Widget) -> None:
        """단계 진행 바를 생성한다.

        Args:
            parent (tk.Widget): 부모 위젯

        Returns:
            None
        """
        step_frame = tk.Frame(parent, bg="#f0f4f8", relief="ridge", bd=1)
        step_frame.pack(fill=tk.X, padx=5, pady=(5, 5), ipady=6)

        inner = tk.Frame(step_frame, bg="#f0f4f8")
        inner.pack(anchor='center')

        steps = [
            "❶ 파일 선택",
            "❷ 파일 분석",
            "❸ 저장 설정",
            "❹ 실행",
        ]

        self.step_labels = []
        self.step_arrows = []

        for i, text in enumerate(steps):
            if i > 0:
                arrow = tk.Label(inner, text="  ➤  ", font=("맑은 고딕", 13, "bold"),
                                 bg="#f0f4f8", fg="#bbb")
                arrow.pack(side=tk.LEFT)
                self.step_arrows.append(arrow)

            lbl = tk.Label(inner, text=f"  {text}  ",
                           font=("맑은 고딕", 10, "bold"),
                           bg="#ddd", fg="#999", padx=12, pady=4,
                           relief="groove", bd=1)
            lbl.pack(side=tk.LEFT, padx=2)
            self.step_labels.append(lbl)

        self.step_hint = tk.Label(inner, text="", font=("맑은 고딕", 9, "bold"),
                                  bg="#f0f4f8", fg="#e74c3c")
        self.step_hint.pack(side=tk.LEFT, padx=(15, 0))

    def _set_step(self, step_num: int) -> None:
        """현재 단계를 설정하고 UI를 업데이트한다.

        Args:
            step_num (int): 단계 번호 (0-based)

        Returns:
            None
        """
        self._current_step = step_num
        hints = [
            "◀ 엑셀 파일을 선택하세요",
            "◀ '📊 파일 분석' 버튼을 클릭하세요",
            "◀ 저장 경로 확인 후 '▶ 마스킹 실행' 클릭",
            "◀ 마스킹 실행 중...",
        ]

        for i, lbl in enumerate(self.step_labels):
            if i < step_num:
                lbl.config(bg="#27ae60", fg="white")
            elif i == step_num:
                lbl.config(bg="#3498db", fg="white")
            else:
                lbl.config(bg="#ddd", fg="#999")

        for i, arrow in enumerate(self.step_arrows):
            if i < step_num:
                arrow.config(fg="#27ae60")
            else:
                arrow.config(fg="#bbb")

        self.step_hint.config(text=hints[step_num] if step_num < len(hints) else "")

        # 깜빡임
        if self._blink_id:
            self.root.after_cancel(self._blink_id)
            self._blink_id = None
        self._blink_state = True
        self._do_blink()

    def _do_blink(self) -> None:
        """현재 단계 라벨을 깜빡이게 한다.

        Args:
            None

        Returns:
            None
        """
        idx = self._current_step
        if idx >= len(self.step_labels):
            return
        lbl = self.step_labels[idx]
        if self._blink_state:
            lbl.config(bg="#3498db", fg="white")
        else:
            lbl.config(bg="#85c1e9", fg="white")
        self._blink_state = not self._blink_state
        self._blink_id = self.root.after(600, self._do_blink)

    # ══════════════════════════════════════════════════════════
    # 파일 선택 / 분석
    # ══════════════════════════════════════════════════════════

    def _browse_file(self) -> None:
        """파일 탐색기를 열어 Excel 파일을 선택한다.

        Args:
            None

        Returns:
            None
        """
        path = filedialog.askopenfilename(
            title="마스킹할 엑셀 파일 선택",
            filetypes=[("Excel", "*.xlsx *.xls *.xlsm"), ("모든 파일", "*.*")])
        if path:
            self.excel_path.set(path)
            self._set_step(1)
            # 시트 목록 로드
            try:
                sheets, engine = get_sheet_names(path)
                self.sheet_combo['values'] = sheets
                if sheets:
                    self.sheet_var.set(sheets[0])
                self._log(f"[시트 목록 로드 완료] 엔진: {engine}, 시트: {sheets}")
            except Exception as e:
                self._log(f"[오류] 시트 목록 로드 실패: {e}")

    def _browse_dir(self) -> None:
        """저장 폴더를 선택한다.

        Args:
            None

        Returns:
            None
        """
        d = filedialog.askdirectory(title="마스킹 결과 저장 폴더 선택")
        if d:
            self.save_dir.set(d)

    def _analyze_file(self) -> None:
        """Excel 파일을 분석하고 미리보기를 표시한다.

        Args:
            None

        Returns:
            None
        """
        path = self.excel_path.get().strip()
        if not path:
            messagebox.showwarning("경고", "먼저 엑셀 파일을 선택하세요.")
            return

        try:
            (df, info), engine = load_excel(path, self.sheet_var.get() or None)
        except Exception as e:
            messagebox.showerror("로드 실패", str(e))
            return

        if len(df.columns) < 2:
            messagebox.showerror("오류", "최소 2개 컬럼(A: Key, B: Value)이 필요합니다.")
            return

        self.df = df
        self.info = info

        # 분석 요약
        key_col = df.columns[0]
        val_col = df.columns[1]
        total_rows = len(df)
        null_count = df[val_col].isna().sum()

        summary = (f"엔진: {engine}  |  시트: {info['sheet_name']}  |  "
                   f"전체 행: {total_rows}  |  Key 컬럼: '{key_col}'  |  Value 컬럼: '{val_col}'\n"
                   f"Value 중 빈 값: {null_count}개  |  마스킹 대상: {total_rows - null_count}개")

        self.analysis_text.config(state=tk.NORMAL)
        self.analysis_text.delete("1.0", tk.END)
        self.analysis_text.insert(tk.END, summary)
        self.analysis_text.config(state=tk.DISABLED)

        # 미리보기 테이블 갱신
        self.tree.delete(*self.tree.get_children())
        self.tree.heading("key", text=f"Key ({key_col})")
        self.tree.heading("value", text=f"Value ({val_col})")

        preview_n = min(100, total_rows)
        for i in range(preview_n):
            k = str(df.iloc[i, 0]) if pd.notna(df.iloc[i, 0]) else "(빈 값)"
            v = str(df.iloc[i, 1]) if pd.notna(df.iloc[i, 1]) else "(빈 값)"
            self.tree.insert("", tk.END, values=(k, v))

        if total_rows > preview_n:
            self.tree.insert("", tk.END, values=(f"... 외 {total_rows - preview_n}건 ...", ""))

        # 저장 경로 기본값
        if not self.save_dir.get():
            self.save_dir.set(os.path.dirname(path))

        # 실행 버튼 활성화
        self.run_btn.config(state='normal')
        self._set_step(2)
        self._log(f"[분석 완료] {total_rows}건 로드. 마스킹 대상: {total_rows - null_count}건")

    # ══════════════════════════════════════════════════════════
    # 로그 유틸
    # ══════════════════════════════════════════════════════════

    def _log(self, msg: str) -> None:
        """로그 메시지를 출력한다.

        Args:
            msg (str): 로그 메시지

        Returns:
            None
        """
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)

    def _update_progress(self, value: float, text: str = "") -> None:
        """프로그레스 바와 상태 라벨을 업데이트한다.

        Args:
            value (float): 진행률 (0~100)
            text (str): 상태 텍스트

        Returns:
            None
        """
        self.progress['value'] = value
        if text:
            self.status_lbl.config(text=text)

    def _set_ui_state(self, state: str) -> None:
        """실행 중 UI를 비활성화/활성화한다.

        Args:
            state (str): 'disabled' 또는 'normal'

        Returns:
            None
        """
        for widget in self.root.winfo_children():
            self._recursive_state(widget, state)
        # 로그는 항상 보이게
        self.log_text.config(state=tk.DISABLED)

    def _recursive_state(self, widget: tk.Widget, state: str) -> None:
        """위젯 트리를 재귀적으로 순회하며 상태를 변경한다.

        Args:
            widget (tk.Widget): 대상 위젯
            state (str): 'disabled' 또는 'normal'

        Returns:
            None
        """
        try:
            if isinstance(widget, (ttk.Button, ttk.Entry, ttk.Combobox, tk.Button)):
                widget.config(state=state)
        except Exception:
            pass
        for child in widget.winfo_children():
            self._recursive_state(child, state)

    # ══════════════════════════════════════════════════════════
    # 마스킹 실행
    # ══════════════════════════════════════════════════════════

    def _run(self) -> None:
        """마스킹을 실행한다 (별도 스레드).

        Args:
            None

        Returns:
            None
        """
        if self._running:
            return
        if self.df is None:
            messagebox.showwarning("경고", "먼저 파일을 분석하세요.")
            return

        save_dir = self.save_dir.get().strip()
        save_name = self.save_name.get().strip()
        if not save_dir or not save_name:
            messagebox.showwarning("경고", "저장 경로와 파일 이름을 입력하세요.")
            return

        try:
            seed = int(self.seed_var.get().strip())
        except ValueError:
            messagebox.showwarning("경고", "시드값은 정수를 입력하세요.")
            return

        self._running = True
        self._set_ui_state('disabled')
        self._set_step(3)
        threading.Thread(target=self._run_worker, args=(save_dir, save_name, seed),
                         daemon=True).start()

    def _run_worker(self, save_dir: str, save_name: str, seed: int) -> None:
        """마스킹 작업을 수행하는 워커 스레드.

        Args:
            save_dir (str): 저장 디렉토리
            save_name (str): 저장 파일명 (확장자 제외)
            seed (int): 랜덤 시드

        Returns:
            None
        """
        try:
            rng = np.random.default_rng(seed)
            df = self.df.copy()
            key_col = df.columns[0]
            val_col = df.columns[1]
            total = len(df)

            self.root.after(0, self._log, f"[시작] 마스킹 시작 — 전체 {total}행")
            self.root.after(0, self._update_progress, 0, "마스킹 중...")

            # 마스킹 수행 + 매핑 기록
            key_masking_map = {}  # {원본키: 마스킹키}
            val_masking_map = {}  # {원본값: 마스킹값}
            masked_keys = []
            masked_values = []

            for i in range(total):
                # Key 마스킹
                key = df.iloc[i, 0]
                if pd.isna(key) or key is None:
                    masked_keys.append(key)
                else:
                    key_str = str(key)
                    if key_str not in key_masking_map:
                        key_masking_map[key_str] = mask_value(key_str, rng)
                    masked_keys.append(key_masking_map[key_str])

                # Value 마스킹
                val = df.iloc[i, 1]
                if pd.isna(val) or val is None:
                    masked_values.append(val)
                else:
                    val_str = str(val)
                    if val_str not in val_masking_map:
                        val_masking_map[val_str] = mask_value(val_str, rng)
                    masked_values.append(val_masking_map[val_str])

                # 진행률 업데이트 (10% 단위)
                if (i + 1) % max(1, total // 10) == 0:
                    pct = int((i + 1) / total * 80)
                    self.root.after(0, self._update_progress, pct,
                                   f"마스킹 중... {i+1}/{total}")

            df[key_col] = masked_keys
            df[val_col] = masked_values
            self.root.after(0, self._update_progress, 80, "파일 저장 중...")
            self.root.after(0, self._log,
                            f"[마스킹 완료] Key 고유값 {len(key_masking_map)}개, "
                            f"Value 고유값 {len(val_masking_map)}개 변환")

            # ── 결과 저장 ──
            os.makedirs(save_dir, exist_ok=True)
            xlsx_path = os.path.join(save_dir, f"{save_name}.xlsx")
            key_path = os.path.join(save_dir, f"{save_name}_마스킹키.json")

            # Excel 저장
            df.to_excel(xlsx_path, index=False, engine='openpyxl')
            self.root.after(0, self._log, f"[저장] {xlsx_path}")

            # 마스킹 키 저장 (복원용)
            key_data = {
                'created_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'source_file': self.info.get('file_name', ''),
                'sheet_name': self.info.get('sheet_name', ''),
                'seed': seed,
                'key_column': key_col,
                'value_column': val_col,
                'total_rows': total,
                'unique_keys': len(key_masking_map),
                'unique_values': len(val_masking_map),
                'key_mapping': {v: k for k, v in key_masking_map.items()},  # 역방향: 마스킹키→원본키
                'value_mapping': {v: k for k, v in val_masking_map.items()},  # 역방향: 마스킹값→원본값
            }
            with open(key_path, 'w', encoding='utf-8') as f:
                json.dump(key_data, f, ensure_ascii=False, indent=2)
            self.root.after(0, self._log, f"[저장] {key_path}")

            self.root.after(0, self._update_progress, 100, "완료!")
            self.root.after(0, self._log,
                            f"\n✅ 마스킹 완료!\n"
                            f"   결과: {xlsx_path}\n"
                            f"   복원키: {key_path}\n"
                            f"   (복원키를 분실하면 원본 복원이 불가합니다)")
            self.root.after(0, lambda: messagebox.showinfo(
                "완료", f"마스킹 완료!\n\n결과: {xlsx_path}\n복원키: {key_path}"))

        except Exception as e:
            self.root.after(0, self._log, f"\n❌ 오류 발생: {e}")
            self.root.after(0, lambda: messagebox.showerror("오류", str(e)))
        finally:
            self.root.after(0, self._finish)

    def _finish(self) -> None:
        """마스킹 완료 후 UI를 복원한다.

        Args:
            None

        Returns:
            None
        """
        self._running = False
        self._set_ui_state('normal')


# ══════════════════════════════════════════════════════════════
# 메인 실행
# ══════════════════════════════════════════════════════════════

def main() -> None:
    """프로그램 진입점.

    Args:
        None

    Returns:
        None
    """
    root = tk.Tk()
    MaskingApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
