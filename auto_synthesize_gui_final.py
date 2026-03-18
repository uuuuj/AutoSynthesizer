"""
auto_synthesize_gui.py
Excel → 합성 데이터 자동 생성 파이프라인  (GUI 버전)

[실행 전 설치]
    pip install pandas numpy scipy openpyxl
    (선택) pip install xlwings    ← DRM 해제 필요 시

[사용법]
    python auto_synthesize_gui.py

[GUI 구성]
    ① 엑셀 파일 선택 (xlwings 우선, openpyxl 자동 폴백)
    ② 컬럼 분석 결과 확인 + 컬럼명 변경
    ③ 문자열 컬럼별 가짜 데이터 1:1 입력 (원본값 → 가짜값)
    ④ 저장 경로 / 파일명 설정
    ⑤ 합성 실행 → 진행 로그 실시간 표시
    ⑥ 원래 데이터로 복원 (변환키 파일 기반)
"""

import os, sys, json, warnings, threading, re
import numpy as np
import pandas as pd
from scipy import stats
from datetime import datetime
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

def _resource_path(filename):
    """PyInstaller EXE에서도 동작하는 리소스 경로 반환"""
    if getattr(sys, 'frozen', False):
        # EXE 실행 중 → PyInstaller 임시 폴더
        base = sys._MEIPASS
    else:
        # 일반 Python 실행 → 스크립트 위치
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, filename)

# xlwings는 선택사항 — 없거나 Excel 미설치면 openpyxl로 폴백
HAS_XLWINGS = False
try:
    import xlwings as xw
    # xlwings가 import되더라도 Excel COM 연결이 안 되면 사용 불가
    # 실제 사용 시점에서 에러 처리하므로 여기서는 플래그만 세팅
    HAS_XLWINGS = True
except (ImportError, OSError, Exception):
    # ImportError: xlwings 미설치
    # OSError: COM 서버 등록 실패 (Excel 미설치)
    # Exception: 기타 초기화 에러
    pass

try:
    import openpyxl
except ImportError:
    if not HAS_XLWINGS:
        raise ImportError("xlwings 또는 openpyxl 중 하나가 필요합니다.\n"
                          "  pip install openpyxl   또는   pip install xlwings")

# ══════════════════════════════════════════════════════════════
# 문자열 동적 생성기 (자동 채우기용)
# ══════════════════════════════════════════════════════════════

_LAST_NAMES  = ['김','이','박','최','정','강','윤','임','한','오',
                '신','홍','문','류','배','전','조','남','서','권']
_FIRST_PARTS = ['민','서','지','수','도','하','준','소','예','태',
                '재','채','나','기','윤','성','우','세','진','아']
_LAST_PARTS  = ['준','연','호','아','서','윤','혁','율','훈','린',
                '민','원','영','현','양','나','재','진','수','은']
_CO_PREFIX = ['Alpha','Beta','Gamma','Delta','Epsilon','Zeta','Eta',
              'Theta','Iota','Kappa','Lambda','Mu','Nu','Xi','Omicron',
              'Pi','Rho','Sigma','Tau','Upsilon','Nova','Apex','Nexus',
              'Prime','Vertex','Zenith','Orion','Titan','Vega','Polaris']
_CO_SUFFIX = ['Corp','Group','Co','Inc','Ltd','Partners','Holdings',
              'Solutions','Systems','Global','Industries','Ventures',
              'Marine','Logistics','Energy','Shipping','Tech','Works']

def generate_fake_persons(n, seed=42):
    rng, names, result, attempt = np.random.default_rng(seed), set(), [], 0
    while len(result) < n:
        name = (_LAST_NAMES[rng.integers(len(_LAST_NAMES))]
              + _FIRST_PARTS[rng.integers(len(_FIRST_PARTS))]
              + _LAST_PARTS[rng.integers(len(_LAST_PARTS))])
        if name not in names:
            names.add(name); result.append(name)
        attempt += 1
        if attempt > n * 200:
            result.append(f'직원_{len(result):04d}')
    return result

def generate_fake_companies(n, seed=42):
    rng, names, result, attempt = np.random.default_rng(seed), set(), [], 0
    while len(result) < n:
        name = f'{_CO_PREFIX[rng.integers(len(_CO_PREFIX))]}-{_CO_SUFFIX[rng.integers(len(_CO_SUFFIX))]}'
        if name not in names:
            names.add(name); result.append(name)
        attempt += 1
        if attempt > n * 200:
            result.append(f'Company-{len(result):04d}')
    return result

def generate_auto_codes(col_name, n, seed=42):
    """컬럼명 기반 자동 코드 생성"""
    cs = str(col_name)[:4]
    alpha = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    return [f'{cs}_{alpha[i] if i<26 else str(i)}' for i in range(n)]

# ══════════════════════════════════════════════════════════════
# 한글 포함 날짜 파싱 유틸리티
# ══════════════════════════════════════════════════════════════

def _parse_korean_date(val):
    """한글이 포함된 날짜 문자열을 파싱한다.
    예: '2024년 3월 15일', '2024년03월15일', '2024. 03. 15.' 등"""
    if not isinstance(val, str):
        return val
    s = val.strip()

    # ── 오전/오후 (Korean AM/PM) 처리: '2025-11-17 오전 12:00:00' ──
    m_ampm = re.match(
        r'(\d{4}[-/\.]\d{1,2}[-/\.]\d{1,2})\s*(오전|오후)\s*(\d{1,2})[:\s](\d{2})(?:[:\s](\d{2}))?',
        s
    )
    if m_ampm:
        date_part = m_ampm.group(1)
        ampm = m_ampm.group(2)
        h, mi = int(m_ampm.group(3)), int(m_ampm.group(4))
        sc = int(m_ampm.group(5)) if m_ampm.group(5) else 0
        if ampm == '오후' and h < 12:
            h += 12
        elif ampm == '오전' and h == 12:
            h = 0
        try:
            base = pd.to_datetime(date_part)
            return base.replace(hour=h, minute=mi, second=sc)
        except Exception:
            pass

    # ── 일반 datetime+시간 혼합 ('2025-11-17 13:00:00' 등) ──
    m_dt = re.match(
        r'(\d{4}[-/\.]\d{1,2}[-/\.]\d{1,2})\s+(\d{1,2}:\d{2}(?::\d{2})?)', s
    )
    if m_dt:
        try:
            return pd.to_datetime(s)
        except Exception:
            pass

    # 한글 날짜 패턴: 2024년 3월 15일 (시, 분, 초 포함 가능)
    m = re.match(
        r'(\d{4})\s*년\s*(\d{1,2})\s*월\s*(\d{1,2})\s*일'
        r'(?:\s*(\d{1,2})\s*시\s*(\d{1,2})\s*분(?:\s*(\d{1,2})\s*초)?)?',
        s
    )
    if m:
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
        h = int(m.group(4)) if m.group(4) else 0
        mi = int(m.group(5)) if m.group(5) else 0
        sc = int(m.group(6)) if m.group(6) else 0
        try:
            return datetime(y, mo, d, h, mi, sc)
        except ValueError:
            return pd.NaT
    # 한글만 제거하고 다시 시도 (예: '24년3월' 같은 부분 패턴)
    cleaned = re.sub(r'[년월일시분초]', ' ', s).strip()
    cleaned = re.sub(r'\s+', ' ', cleaned)
    if cleaned != s:
        try:
            return pd.to_datetime(cleaned)
        except Exception:
            pass
    return None  # 한글 날짜가 아님 → None 반환하여 기본 파서에 위임


def _clean_datetime_column(series):
    """datetime 컬럼을 순수 Date(YYYY-MM-DD) 또는 Time(HH:MM:SS)으로 정리.
    - 모든 시간이 00:00:00이면 → date only
    - 모든 날짜가 동일하면 → time only
    - 그 외 → date only (시간 제거)
    """
    if not pd.api.types.is_datetime64_any_dtype(series):
        return series
    valid = series.dropna()
    if len(valid) == 0:
        return series

    times = valid.dt.time
    dates = valid.dt.date
    all_midnight = (times == pd.Timestamp('00:00:00').time()).all()
    unique_dates = dates.nunique()

    if unique_dates <= 1 and not all_midnight:
        # 날짜가 모두 같고 시간이 다양 → Time only
        return series.dt.strftime('%H:%M:%S').where(series.notna(), other=np.nan)
    else:
        # Date only (시간 제거)
        return series.dt.strftime('%Y-%m-%d').where(series.notna(), other=np.nan)


def _safe_to_datetime(series):
    """한글 날짜를 포함할 수 있는 시리즈를 안전하게 datetime으로 변환한다.
    null 값은 그대로 유지한다."""
    result = series.copy()
    for idx, val in series.items():
        if pd.isna(val) or val is None:
            result.at[idx] = pd.NaT
            continue
        if isinstance(val, datetime):
            continue
        # 한글 날짜 파싱 시도
        parsed = _parse_korean_date(val)
        if parsed is not None:
            result.at[idx] = parsed
        else:
            # 일반 파서에 위임
            try:
                result.at[idx] = pd.to_datetime(val)
            except Exception:
                result.at[idx] = pd.NaT
    return pd.to_datetime(result, errors='coerce')


# ══════════════════════════════════════════════════════════════
# Excel 로드 — xlwings 우선, 실패 시 openpyxl/pandas 폴백
# ══════════════════════════════════════════════════════════════

def _postprocess_dataframe(df):
    """로드된 DataFrame에 대해 컬럼별 타입 추론 (한글 날짜, 숫자 등)
    null 값은 보존한다."""
    for col in df.columns:
        s = df[col]

        # ── datetime 감지 (python datetime 객체) ──
        if s.dtype == object:
            smp = s.dropna().head(10)
            if len(smp) and smp.apply(lambda x: isinstance(x, datetime)).all():
                df[col] = pd.to_datetime(s, errors='coerce')
                continue

        # ── 오전/오후 포함 날짜 패턴 감지 ──
        if s.dtype == object:
            smp_str = s.dropna().head(20).astype(str)
            has_ampm = smp_str.str.contains(r'오전|오후', na=False).mean() > 0.3
            if has_ampm:
                converted = _safe_to_datetime(s)
                if converted.notna().mean() > 0.5:
                    df[col] = converted
                    continue

        # ── 한글 날짜 패턴 감지 ('년', '월', '일' 포함) ──
        if s.dtype == object:
            smp_str = s.dropna().head(20).astype(str)
            has_korean_date = smp_str.str.contains(r'\d+\s*년', na=False).mean() > 0.5
            if has_korean_date:
                converted = _safe_to_datetime(s)
                if converted.notna().mean() > 0.5:
                    df[col] = converted
                    continue

        # ── 일반 날짜 문자열 파싱 ──
        if s.dtype == object:
            try:
                conv = pd.to_datetime(s, infer_datetime_format=True, errors='coerce')
                if conv.notna().mean() > 0.8:
                    df[col] = conv
                    continue
            except Exception:
                pass

        # ── 숫자 변환 (콤마 제거 후) ──
        if s.dtype == object:
            try:
                num_conv = pd.to_numeric(
                    s.astype(str).str.replace(',', '', regex=False).str.strip(),
                    errors='coerce'
                )
                if num_conv.notna().mean() >= 0.5:
                    df[col] = num_conv
                else:
                    pass  # 의미없으면 원본 유지
            except Exception:
                pass

    # ── null 값 보존: 전체 행/열이 비어있을 때만 제거 ──
    df.dropna(how='all', inplace=True)
    df.dropna(axis=1, how='all', inplace=True)
    df.reset_index(drop=True, inplace=True)
    return df

# ── xlwings 전용 함수 ──────────────────────────────────

def _get_sheet_names_xlwings(excel_path):
    app, wb = None, None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        wb = app.books.open(os.path.abspath(excel_path))
        return [s.name for s in wb.sheets]
    finally:
        if wb:  wb.close()
        if app: app.quit()

def _load_excel_xlwings(excel_path, sheet_name=None):
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
        df = _postprocess_dataframe(df)

        info = {
            'file_name': wb.name, 'file_path': excel_path or wb.fullname,
            'sheet_name': ws.name, 'all_sheets': [s.name for s in wb.sheets],
            'rows': len(df), 'cols': len(df.columns),
        }
        return df, info
    finally:
        if close_after:
            if wb:  wb.close()
            if app: app.quit()

# ── openpyxl / pandas 폴백 함수 ──────────────────────

def _get_sheet_names_openpyxl(excel_path):
    """openpyxl로 시트 이름 읽기 (read_only 모드 — 빠름)"""
    abs_path = os.path.abspath(excel_path)
    if not os.path.exists(abs_path):
        raise FileNotFoundError(f"파일 없음: {abs_path}")
    wb = openpyxl.load_workbook(abs_path, read_only=True, data_only=True)
    names = wb.sheetnames
    wb.close()
    return names

def _load_excel_openpyxl(excel_path, sheet_name=None):
    """openpyxl + pandas로 Excel 로드 (DRM 걸린 파일은 읽지 못할 수 있음)"""
    abs_path = os.path.abspath(excel_path)
    if not os.path.exists(abs_path):
        raise FileNotFoundError(f"파일 없음: {abs_path}")

    file_name = os.path.basename(abs_path)

    # pandas read_excel (openpyxl 엔진)
    xls = pd.ExcelFile(abs_path, engine='openpyxl')
    all_sheets = xls.sheet_names

    target_sheet = sheet_name if sheet_name else all_sheets[0]
    df = pd.read_excel(xls, sheet_name=target_sheet, engine='openpyxl')
    xls.close()

    df = _postprocess_dataframe(df)

    info = {
        'file_name': file_name, 'file_path': abs_path,
        'sheet_name': target_sheet, 'all_sheets': all_sheets,
        'rows': len(df), 'cols': len(df.columns),
    }
    return df, info

# ── 통합 래퍼 — xlwings 우선, 실패 시 openpyxl ─────

def get_sheet_names(excel_path):
    """시트 이름 목록 반환. xlwings → openpyxl 순서로 시도."""
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

def load_excel(excel_path, sheet_name=None):
    """Excel 로드. xlwings → openpyxl 순서로 시도."""
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
# 컬럼 타입 감지
# ══════════════════════════════════════════════════════════════

def auto_detect_column_type(series):
    sample = series.dropna()
    if len(sample) == 0: return 'categorical'
    if pd.api.types.is_datetime64_any_dtype(series): return 'datetime'
    if series.dtype == object:
        smp_str = sample.astype(str)
        # 한글 날짜 패턴도 datetime으로 인식
        if smp_str.str.contains(r'\d+\s*년\s*\d+\s*월', na=False).mean() > 0.5:
            return 'datetime'
        if smp_str.str.match(r'\d{4}[-/]\d{2}[-/]\d{2}').mean() > 0.8:
            return 'datetime'
    if pd.api.types.is_numeric_dtype(series):
        return 'categorical' if series.nunique()/len(series) < 0.05 else 'numerical'
    if series.dtype == object:
        if series.nunique()/len(series) < 0.15: return 'categorical'
    return 'categorical'

def _is_id_col(series):
    return series.dropna().astype(str).str.match(
        r'^[A-Za-z가-힣]{1,3}[\-_]?\d{3,}$').mean() > 0.7

# ══════════════════════════════════════════════════════════════
# 문자열 합성 (매핑 기반) — null 보존
# ══════════════════════════════════════════════════════════════

def synthesize_text_columns(df, col_types, mapping_dict):
    """
    mapping_dict: {컬럼명: {원본값: 가짜값, ...}}
    null 값은 그대로 유지한다.
    """
    syn, desc_map = df.copy(), {}
    str_cols = [c for c in df.columns
                if (df[c].dtype == object or pd.api.types.is_string_dtype(df[c]))
                and col_types.get(c) == 'categorical']
    for col in str_cols:
        if col not in mapping_dict:
            continue
        mapping = mapping_dict[col]
        if not mapping:
            continue
        # null 마스크를 먼저 저장
        null_mask = df[col].isna()
        mapped = df[col].astype(str).map(mapping)
        # 매핑에 없는 값은 원본 유지, null은 NaN 유지
        mapped = mapped.where(mapped.notna(), df[col])
        mapped = mapped.where(~null_mask, other=np.nan)
        syn[col] = mapped
        desc_map[col] = {'method': '사용자 매핑', 'mapping': mapping}
    return syn, desc_map

# ══════════════════════════════════════════════════════════════
# 상관관계 / 제약조건 / 수치합성 / 품질검증
# ══════════════════════════════════════════════════════════════

def analyze_correlations(df, col_types):
    num_cols = [c for c, t in col_types.items() if t == 'numerical']
    if len(num_cols) < 2:
        return {}
    corr = df[num_cols].corr()
    pairs = []
    for i, c1 in enumerate(num_cols):
        for c2 in num_cols[i+1:]:
            r = corr.loc[c1, c2]
            if abs(r) >= 0.5:
                pairs.append({'col1': c1, 'col2': c2, 'r': round(r, 3)})
    return {'correlation_matrix': corr.round(3).to_dict(), 'strong_pairs': pairs}

def auto_detect_constraints(df, col_types):
    constraints = []
    num_cols  = [c for c, t in col_types.items() if t == 'numerical']
    date_cols = [c for c, t in col_types.items() if t == 'datetime']
    for col in num_cols:
        col_data = df[col].dropna()
        if len(col_data) == 0:
            continue
        if col_data.min() >= 0:
            constraints.append({'type': 'positive', 'column': col})
        if col_data.min() >= 0 and col_data.max() <= 100:
            constraints.append({'type': 'range_0_100', 'column': col})
    start_kw = ['시작', '착수', '계획', 'start', 'begin', 'from']
    end_kw   = ['종료', '완료', '끝', 'end', 'finish', 'to']
    s_cols = [c for c in date_cols if any(k in str(c) for k in start_kw)]
    e_cols = [c for c in date_cols if any(k in str(c) for k in end_kw)]
    for s in s_cols:
        for e in e_cols:
            try:
                mask = df[s].notna() & df[e].notna()
                if mask.sum() > 0 and (df.loc[mask, s] < df.loc[mask, e]).mean() > 0.9:
                    constraints.append({'type': 'inequality', 'low': s, 'high': e})
            except Exception:
                pass
    return constraints

def generate_numeric_datetime(df, col_types, constraints, num_rows=None):
    """수치/날짜 합성 — 상관관계 보존 + 제약조건 + null 비율 유지

    [Issue 2] 제약조건 적용 시 abs()/clip() 대신 유효 범위 내 재분배
              → 상관관계 구조 보존
    [Issue 6] null 삽입은 모든 제약조건 적용 후 마지막에 수행
    [Issue 7] 날짜 기간 분포에서 0일도 포함 (원본 그대로)
    """
    n = num_rows or len(df)
    syn = pd.DataFrame(index=range(n))
    num_cols  = [c for c, t in col_types.items() if t == 'numerical']
    date_cols = [c for c, t in col_types.items() if t == 'datetime']

    # ── Step 1: 각 컬럼의 null 비율 기록 ──
    null_ratios = {}
    for col in num_cols + date_cols:
        total = len(df[col])
        null_count = df[col].isna().sum()
        null_ratios[col] = null_count / total if total > 0 else 0

    # ── Step 2: 수치 컬럼 — 제약조건 반영 Copula ──
    # 제약조건 사전 수집 (Copula 단계에서 참조)
    pos_cols = {c['column'] for c in constraints if c['type'] == 'positive'}
    range_cols = {c['column'] for c in constraints if c['type'] == 'range_0_100'}

    if num_cols:
        # 컬럼별 독립 dropna (listwise deletion 대신) → 더 많은 데이터 활용
        col_data_map = {}
        for col in num_cols:
            vals = df[col].dropna().values.astype(float)
            if len(vals) > 0:
                col_data_map[col] = vals

        # 상관관계 추정 — pairwise (listwise보다 데이터 손실 적음)
        valid_num_cols = [c for c in num_cols if c in col_data_map]
        if len(valid_num_cols) > 1:
            sub = df[valid_num_cols].dropna()
            if len(sub) > 1:
                normals = np.zeros((len(sub), len(valid_num_cols)))
                params = {}
                for i, col in enumerate(valid_num_cols):
                    vals = sub[col].values.astype(float)
                    normals[:, i] = stats.norm.ppf(
                        stats.rankdata(vals) / (len(vals) + 1))
                    # 제약조건이 있는 컬럼은 제약 범위 내 값만 경험 분포에 포함
                    sorted_vals = np.sort(col_data_map[col])
                    if col in pos_cols:
                        sorted_vals = sorted_vals[sorted_vals >= 0]
                    if col in range_cols:
                        sorted_vals = sorted_vals[(sorted_vals >= 0) & (sorted_vals <= 100)]
                    if len(sorted_vals) == 0:
                        sorted_vals = np.sort(col_data_map[col])
                    params[col] = sorted_vals

                cov = np.cov(normals.T) + np.eye(len(valid_num_cols)) * 1e-6
                sample = np.random.multivariate_normal(
                    np.zeros(len(valid_num_cols)), cov, n)
                unif = stats.norm.cdf(sample)

                for i, col in enumerate(valid_num_cols):
                    idx = np.clip(
                        (unif[:, i] * (len(params[col]) - 1)).astype(int),
                        0, len(params[col]) - 1
                    )
                    syn[col] = params[col][idx]
                    if df[col].dtype in [np.int32, np.int64]:
                        syn[col] = syn[col].round().astype(int)
            elif len(valid_num_cols) == 1:
                col = valid_num_cols[0]
                sorted_vals = np.sort(col_data_map[col])
                idx = np.random.randint(0, len(sorted_vals), n)
                syn[col] = sorted_vals[idx]
        elif len(valid_num_cols) == 1:
            col = valid_num_cols[0]
            sorted_vals = np.sort(col_data_map[col])
            idx = np.random.randint(0, len(sorted_vals), n)
            syn[col] = sorted_vals[idx]

    # ── Step 3: 날짜 컬럼 ──
    ineq = {c['low']: c['high'] for c in constraints if c['type'] == 'inequality'}
    for col in date_cols:
        if col in ineq.values():
            continue
        ts = df[col].dropna()
        if len(ts) == 0:
            continue
        ts_vals = ts.values.astype('datetime64[s]').astype(np.int64)
        sampled = np.random.choice(ts_vals, size=n)
        if col in ineq:
            end_col = ineq[col]
            mask = df[col].notna() & df[end_col].notna()
            dur = (df.loc[mask, end_col] - df.loc[mask, col]).dt.days.values
            # [Issue 7] 0일 포함, 음수만 abs 보정 (원본 분포 그대로)
            if len(dur) == 0:
                dur = np.array([0])
            dur = np.abs(dur)  # 데이터 오류로 인한 음수만 보정
            syn[col] = pd.to_datetime(sampled, unit='s')
            syn[end_col] = syn[col] + pd.to_timedelta(
                np.random.choice(dur, n), unit='D'
            )
        else:
            syn[col] = pd.to_datetime(sampled, unit='s')

    # ── Step 4: 제약조건 — 범위 내 재분배 (상관관계 보존) ──
    for c in constraints:
        col = c.get('column')
        if col is None or col not in syn.columns:
            continue
        col_vals = syn[col].values.astype(float)
        non_null = ~np.isnan(col_vals)

        if c['type'] == 'positive':
            violations = non_null & (col_vals < 0)
            if violations.any():
                valid = col_vals[non_null & (col_vals >= 0)]
                if len(valid) > 0:
                    # 위반값을 유효 범위 하한 근처로 재분배
                    lo = valid.min()
                    hi = np.percentile(valid, 25)
                    if lo == hi:
                        hi = lo + 1
                    col_vals[violations] = np.random.uniform(lo, hi, violations.sum())
                else:
                    col_vals[violations] = np.abs(col_vals[violations])
                syn[col] = col_vals
                if df[col].dtype in [np.int32, np.int64]:
                    syn[col] = syn[col].round().astype(int)

        elif c['type'] == 'range_0_100':
            violations = non_null & ((col_vals < 0) | (col_vals > 100))
            if violations.any():
                valid = col_vals[non_null & (col_vals >= 0) & (col_vals <= 100)]
                if len(valid) > 0:
                    col_vals[violations] = np.random.choice(valid, violations.sum())
                else:
                    col_vals[violations] = np.clip(col_vals[violations], 0, 100)
                syn[col] = col_vals
                if df[col].dtype in [np.int32, np.int64]:
                    syn[col] = syn[col].round().astype(int)

    # ── Step 5 (최후): null 비율 복원 — 모든 제약 적용 완료 후 ──
    for col in syn.columns:
        if col in null_ratios and null_ratios[col] > 0:
            null_count = int(round(n * null_ratios[col]))
            if null_count > 0:
                null_indices = np.random.choice(n, size=null_count, replace=False)
                syn.loc[null_indices, col] = np.nan

    return syn

def detect_functional_dependencies(df, col_types):
    """[Issue 3] 범주형 컬럼 간 함수 종속성 감지.
    A → B: A의 각 고유값이 항상 같은 B값에 매핑되면 A가 B를 결정.
    """
    cat_cols = [c for c, t in col_types.items() if t == 'categorical']
    deps = []
    for a in cat_cols:
        for b in cat_cols:
            if a == b:
                continue
            sub = df[[a, b]].dropna()
            if len(sub) < 2:
                continue
            grouped = sub.groupby(a)[b].nunique()
            if (grouped == 1).all():
                deps.append({'from': a, 'to': b})
    return deps


def validate_quality(real, synthetic, col_types):
    scores = {}
    for col, ctype in col_types.items():
        if col not in real.columns or col not in synthetic.columns:
            continue
        try:
            if ctype == 'numerical':
                r_clean = real[col].dropna()
                s_clean = synthetic[col].dropna()
                if len(r_clean) == 0 or len(s_clean) == 0:
                    scores[col] = 0.5
                    continue
                ks, _ = stats.ks_2samp(r_clean, s_clean)
                scores[col] = round(1 - ks, 3)
            elif ctype == 'categorical':
                r = real[col].value_counts(normalize=True)
                s = synthetic[col].value_counts(normalize=True)
                r_s = r.sort_values(ascending=False).values
                s_s = s.sort_values(ascending=False).values
                mn = min(len(r_s), len(s_s))
                scores[col] = round(1 - np.abs(r_s[:mn] - s_s[:mn]).sum() / 2, 3)
            elif ctype == 'datetime':
                r_clean = real[col].dropna()
                s_clean = synthetic[col].dropna()
                if len(r_clean) == 0 or len(s_clean) == 0:
                    scores[col] = 0.5
                    continue
                r_ts = r_clean.values.astype('datetime64[s]').astype(np.int64)
                s_ts = s_clean.values.astype('datetime64[s]').astype(np.int64)
                ks, _ = stats.ks_2samp(r_ts, s_ts)
                scores[col] = round(1 - ks, 3)
        except Exception:
            scores[col] = 0.5
    overall = round(float(np.mean(list(scores.values()))), 3) if scores else 0.0
    return overall, scores

# ══════════════════════════════════════════════════════════════
# GUI
# ══════════════════════════════════════════════════════════════

class SynthesizeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("합성 데이터 자동 생성기  —  삼성중공업 생산 DT센터")
        self.root.geometry("1150x1100")
        self.root.minsize(1050, 950)

        # 창 아이콘 설정 (ico 파일이 있으면)
        try:
            ico_path = _resource_path("synth_ico.ico")
            if os.path.exists(ico_path):
                self.root.iconbitmap(ico_path)
        except Exception:
            pass

        self.df = None
        self.info = None
        self.col_types = {}
        self.col_entry_map = {}       # {col_name: [(orig_val, entry_widget, fixed_val), ...]}
        self.col_rename_entries = {}   # {orig_col_name: Entry widget}
        self.original_columns = []    # 원본 컬럼명 보존

        self._current_step = 0
        self._blink_state  = True
        self._blink_id     = None
        self._data_confirmed = False    # 데이터 확정 여부
        self.step_labels   = []
        self.step_arrows   = []

        self.excel_path   = tk.StringVar()
        self.sheet_var    = tk.StringVar()
        self.num_rows_var = tk.StringVar(value="")
        self.save_dir     = tk.StringVar()
        self.save_name    = tk.StringVar(value="합성데이터")
        self._build_ui()
        self._set_step(0)  # 초기 단계: 파일 선택

    def _build_ui(self):
        style = ttk.Style()
        style.configure("Sub.TLabel", font=("맑은 고딕", 9), foreground="#666")
        style.configure("ColHeader.TLabel", font=("맑은 고딕", 10, "bold"), foreground="#333")
        style.configure("Run.TButton", font=("맑은 고딕", 11, "bold"))
        style.configure("Auto.TButton", font=("맑은 고딕", 9))

        # ── 로고 배너 ──
        self._logo_image = None  # 참조 유지 (GC 방지)
        try:
            logo_path = _resource_path("logo.png")
            if os.path.exists(logo_path) and HAS_PIL:
                img = Image.open(logo_path)
                # 창 너비에 맞게 리사이즈 (비율 유지)
                target_w = 1140
                ratio = target_w / img.width
                target_h = int(img.height * ratio)
                img = img.resize((target_w, target_h), Image.LANCZOS)
                self._logo_image = ImageTk.PhotoImage(img)
                logo_lbl = tk.Label(self.root, image=self._logo_image, bg="#2d4a7a")
                logo_lbl.pack(fill=tk.X, padx=5, pady=(5, 0))
        except Exception:
            pass  # 로고 로드 실패 시 무시 — 기능에 영향 없음

        # ── 노트북(탭) 구성 ──
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # ═══ 탭1: 합성 데이터 생성 ═══
        tab_synth = ttk.Frame(self.notebook, padding=5)
        self.notebook.add(tab_synth, text="  📊 합성 데이터 생성  ")

        # ── 단계 안내 표시 바 ──
        self._build_step_indicator(tab_synth)

        # ┌─────────────────────────────────────────────┐
        # │  레이아웃: 상단(파일선택) 고정               │
        # │           중간(분석+컬럼명+가짜데이터) 스크롤 │
        # │           하단(저장+실행) 고정                │
        # └─────────────────────────────────────────────┘

        # ══ 하단 고정: ④ 저장 + ⑤ 실행 (먼저 pack → 항상 보임) ══
        bottom_frame = ttk.Frame(tab_synth)
        bottom_frame.pack(side=tk.BOTTOM, fill=tk.X)

        sec4 = ttk.LabelFrame(bottom_frame, text="  ④ 저장 설정  ", padding=8)
        sec4.pack(fill=tk.X, pady=(2, 2))

        rs1 = ttk.Frame(sec4); rs1.pack(fill=tk.X)
        ttk.Label(rs1, text="저장 경로:", width=10).pack(side=tk.LEFT)
        ttk.Entry(rs1, textvariable=self.save_dir, font=("Consolas", 9)).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        ttk.Button(rs1, text="폴더 선택...", command=self._browse_dir).pack(side=tk.LEFT)

        rs2 = ttk.Frame(sec4); rs2.pack(fill=tk.X, pady=(5, 0))
        ttk.Label(rs2, text="파일 이름:", width=10).pack(side=tk.LEFT)
        ttk.Entry(rs2, textvariable=self.save_name, width=35, font=("Consolas", 9)).pack(
            side=tk.LEFT, padx=(0, 6))
        ttk.Label(rs2, text=".xlsx / _변환키.json / _description.json / _품질리포트.json 자동 생성",
                  style="Sub.TLabel").pack(side=tk.LEFT)

        sec5 = ttk.LabelFrame(bottom_frame, text="  ⑤ 실행 로그  ", padding=8)
        sec5.pack(fill=tk.BOTH, expand=True, pady=(0, 2))

        bf = ttk.Frame(sec5); bf.pack(fill=tk.X, pady=(0, 5))
        self.run_btn = ttk.Button(bf, text="▶  합성 데이터 생성 실행",
                                  command=self._run, style="Run.TButton",
                                  state='disabled')
        self.run_btn.pack(side=tk.LEFT)
        self.progress = ttk.Progressbar(bf, mode='determinate', maximum=100, length=200)
        self.progress.pack(side=tk.LEFT, padx=(12, 0))
        self.status_lbl = ttk.Label(bf, text="", style="Sub.TLabel")
        self.status_lbl.pack(side=tk.LEFT, padx=(12, 0))

        self.log_text = scrolledtext.ScrolledText(sec5, height=6, font=("Consolas", 9),
                                                   state=tk.DISABLED, wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # ══ 상단 고정: ① 엑셀 파일 선택 ══
        top_frame = ttk.Frame(tab_synth)
        top_frame.pack(side=tk.TOP, fill=tk.X)

        sec1 = ttk.LabelFrame(top_frame, text="  ① 엑셀 파일 선택  (xlwings 우선 / openpyxl 폴백)  ", padding=8)
        sec1.pack(fill=tk.X, pady=(0, 5))

        r1 = ttk.Frame(sec1); r1.pack(fill=tk.X)
        ttk.Label(r1, text="파일 경로:", width=10).pack(side=tk.LEFT)
        ttk.Entry(r1, textvariable=self.excel_path, font=("Consolas", 9)).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        ttk.Button(r1, text="찾아보기...", command=self._browse_file).pack(side=tk.LEFT)
        ttk.Button(r1, text="🔄 새 파일", command=self._refresh_all,
                   style="Auto.TButton").pack(side=tk.LEFT, padx=(8, 0))

        r2 = ttk.Frame(sec1); r2.pack(fill=tk.X, pady=(5, 0))
        ttk.Label(r2, text="시트 선택:", width=10).pack(side=tk.LEFT)
        self.sheet_combo = ttk.Combobox(r2, textvariable=self.sheet_var, state='readonly', width=25)
        self.sheet_combo.pack(side=tk.LEFT, padx=(0, 20))
        ttk.Label(r2, text="생성 행 수:").pack(side=tk.LEFT)
        ttk.Entry(r2, textvariable=self.num_rows_var, width=10).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Label(r2, text="(비우면 원본과 동일)", style="Sub.TLabel").pack(side=tk.LEFT)

        ttk.Button(sec1, text="📊  파일 분석", command=self._analyze_file).pack(pady=(7, 0))

        # ── ② 컬럼 분석 요약 ────────────────────────────
        self.analysis_text = tk.Text(top_frame, height=3, font=("Consolas", 9),
                                     bg="#f5f5f0", state=tk.DISABLED, wrap=tk.WORD)
        self.analysis_text.pack(fill=tk.X, pady=(0, 5))

        # ══ 중간 스크롤 영역: ②-1 컬럼명 변경 + ③ 가짜데이터 ══
        mid_frame = ttk.Frame(tab_synth)
        mid_frame.pack(fill=tk.BOTH, expand=True)

        self.main_canvas = tk.Canvas(mid_frame, highlightthickness=0)
        main_sb = ttk.Scrollbar(mid_frame, orient=tk.VERTICAL, command=self.main_canvas.yview)
        self.main_inner = ttk.Frame(self.main_canvas)
        self.main_inner.bind("<Configure>",
            lambda e: self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all")))
        self._main_cw = self.main_canvas.create_window((0, 0), window=self.main_inner, anchor="nw")
        self.main_canvas.configure(yscrollcommand=main_sb.set)
        self.main_canvas.bind("<Configure>",
            lambda e: self.main_canvas.itemconfig(self._main_cw, width=e.width))
        self.main_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        main_sb.pack(side=tk.RIGHT, fill=tk.Y)

        # ── ②-1 컬럼명 변경 ─────────────────────────────
        sec_rename = ttk.LabelFrame(self.main_inner, text="  ②-1 컬럼명 변경  (변경할 컬럼만 새 이름 입력)  ", padding=8)
        sec_rename.pack(fill=tk.X, pady=(0, 5), padx=2)

        rename_top = ttk.Frame(sec_rename)
        rename_top.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(rename_top,
                  text="각 컬럼의 '새 이름'을 입력하세요. 비워두면 원본 이름 유지.",
                  style="Sub.TLabel").pack(side=tk.LEFT)
        ttk.Button(rename_top, text="✅ 컬럼명 적용",
                   command=self._apply_column_rename, style="Auto.TButton").pack(side=tk.RIGHT)

        # 컬럼명 변경 내부 프레임 (스크롤 없이 직접 배치 → 전체 스크롤에 포함)
        self.rename_inner = ttk.Frame(sec_rename)
        self.rename_inner.pack(fill=tk.X)

        # ── ③ 가짜 데이터 1:1 입력 ──────────────────────
        sec3 = ttk.LabelFrame(self.main_inner, text="  ③ 가짜 데이터 입력  (원본값 → 가짜값)  ", padding=8)
        sec3.pack(fill=tk.X, pady=(0, 5), padx=2)

        top_bar = ttk.Frame(sec3)
        top_bar.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(top_bar,
                  text="각 원본값 옆에 가짜값을 입력하세요.  비운 항목은 자동 채우기로 한번에 채울 수 있습니다.",
                  style="Sub.TLabel").pack(side=tk.LEFT)
        self.auto_all_btn = ttk.Button(top_bar, text="🔄 전체 자동 채우기",
                                        command=self._auto_fill_all, style="Auto.TButton")
        self.auto_all_btn.pack(side=tk.RIGHT)

        # 가짜 데이터 내부 프레임 (스크롤 없이 직접 배치 → 전체 스크롤에 포함)
        self.input_inner = ttk.Frame(sec3)
        self.input_inner.pack(fill=tk.X)

        # ── ③-1 데이터 확정 버튼 + 안내 ──
        self.confirm_frame = ttk.Frame(self.main_inner)
        self.confirm_frame.pack(fill=tk.X, pady=(8, 5), padx=2)

        self.confirm_btn = tk.Button(
            self.confirm_frame, text="  ✅  변환 계획 확정  ",
            font=("맑은 고딕", 12, "bold"), bg="#27ae60", fg="white",
            activebackground="#219a52", activeforeground="white",
            relief="raised", bd=2, padx=20, pady=8,
            command=self._confirm_data)
        self.confirm_btn.pack(pady=(5, 3))

        self.confirm_guide = tk.Label(
            self.confirm_frame,
            text="▲ 컬럼명과 가짜 데이터를 확인한 뒤, 위 버튼을 눌러 변환 계획을 확정하세요.",
            font=("맑은 고딕", 9, "bold"), fg="#e67e22")
        self.confirm_guide.pack()

        # 확정 후 안내 (처음엔 숨김)
        self.after_confirm_lbl = tk.Label(
            self.confirm_frame,
            text="",
            font=("맑은 고딕", 10, "bold"), fg="#27ae60")
        self.after_confirm_lbl.pack(pady=(3, 0))

        # ── 마우스 휠: 중간 스크롤 영역에서만 동작 ──
        def _on_mousewheel(event):
            self.main_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        def _bind_mousewheel(event):
            self.main_canvas.bind_all("<MouseWheel>", _on_mousewheel)

        def _unbind_mousewheel(event):
            self.main_canvas.unbind_all("<MouseWheel>")

        self.main_canvas.bind("<Enter>", _bind_mousewheel)
        self.main_canvas.bind("<Leave>", _unbind_mousewheel)

    # ══════════════════════════════════════════════════════════
    # Refresh (새 파일) / 데이터 확정
    # ══════════════════════════════════════════════════════════

    def _refresh_all(self):
        """모든 상태를 초기화하여 새 파일을 시작할 수 있게 한다."""
        # 데이터 초기화
        self.df = None
        self.info = None
        self.col_types = {}
        self.col_entry_map = {}
        self.col_rename_entries = {}
        self.original_columns = []
        self._data_confirmed = False

        # 입력 필드 초기화
        self.excel_path.set("")
        self.sheet_var.set("")
        self.num_rows_var.set("")
        self.save_dir.set("")
        self.save_name.set("합성데이터")
        self.sheet_combo['values'] = []

        # 분석 텍스트 초기화
        self.analysis_text.config(state=tk.NORMAL)
        self.analysis_text.delete("1.0", tk.END)
        self.analysis_text.config(state=tk.DISABLED)

        # 컬럼명/가짜데이터 위젯 초기화
        for w in self.rename_inner.winfo_children():
            w.destroy()
        for w in self.input_inner.winfo_children():
            w.destroy()

        # 확정 버튼 복원
        self.confirm_btn.config(state='normal', bg="#27ae60", text="  ✅  변환 계획 확정  ")
        self.confirm_guide.config(
            text="▲ 컬럼명과 가짜 데이터를 확인한 뒤, 위 버튼을 눌러 변환 계획을 확정하세요.")
        self.after_confirm_lbl.config(text="")

        # 실행 버튼 잠금
        self.run_btn.config(state='disabled')

        # 로그 초기화
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete("1.0", tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.status_lbl.config(text="")

        # 단계 초기화
        self._set_step(0)

    def _validate_mappings(self):
        """가짜값 중복 검사 — 같은 가짜값이 서로 다른 원본값에 할당되면 복원 불가."""
        issues = []
        for col, entries in self.col_entry_map.items():
            fake_to_orig = {}
            for val, entry, fixed in entries:
                fake = fixed if fixed is not None else (entry.get().strip() if entry else None)
                if not fake:
                    continue
                if fake in fake_to_orig:
                    issues.append(
                        f"  ⚠ '{col}' 컬럼: 가짜값 '{fake}'이(가) "
                        f"'{fake_to_orig[fake]}'와 '{val}'에 중복 할당됨"
                    )
                else:
                    fake_to_orig[fake] = val
        return issues

    def _confirm_data(self):
        """변환 계획을 확정하고 실행 단계로 진행."""
        if self.df is None:
            messagebox.showwarning("경고", "먼저 파일을 분석해 주세요.")
            return

        # ── Issue 1: 매핑 중복 검증 ──
        dup_issues = self._validate_mappings()
        if dup_issues:
            msg = ("가짜값이 중복된 항목이 있습니다.\n"
                   "이대로 진행하면 복원 시 구분이 불가능합니다.\n\n"
                   + "\n".join(dup_issues)
                   + "\n\n그래도 진행하시겠습니까?")
            if not messagebox.askyesno("매핑 중복 경고", msg, icon='warning'):
                return

        # 컬럼명 자동 적용
        if self.col_rename_entries:
            has_rename = any(e.get().strip() for e in self.col_rename_entries.values())
            if has_rename:
                self._apply_column_rename_silent()

        self._data_confirmed = True

        # 확정 버튼 비활성화 + 안내 메시지 변경
        self.confirm_btn.config(state='disabled', bg="#95a5a6",
                                text="  ✅  변환 계획 확정 완료  ")
        self.confirm_guide.config(text="")
        self.after_confirm_lbl.config(
            text="✅ 변환 계획이 확정되었습니다.\n"
                 "→  아래 ④ 저장 설정에서 저장 경로·파일명을 확인한 뒤,\n"
                 "→  '▶ 합성 데이터 생성 실행' 버튼을 클릭하여 데이터를 변환 및 저장하십시오.")

        # 실행 버튼 활성화
        self.run_btn.config(state='normal')

        # 단계 진행 → ❺ 실행
        self._set_step(4)
        self.status_lbl.config(text="변환 계획 확정됨 — '▶ 합성 데이터 생성 실행' 버튼을 클릭하세요")

    # ══════════════════════════════════════════════════════════
    # 단계 안내 표시 바 (깜빡이는 화살표)
    # ══════════════════════════════════════════════════════════

    def _build_step_indicator(self, parent):
        step_frame = tk.Frame(parent, bg="#f0f4f8", relief="ridge", bd=1)
        step_frame.pack(side=tk.TOP, fill=tk.X, pady=(0, 5), ipady=6)

        inner = tk.Frame(step_frame, bg="#f0f4f8")
        inner.pack(anchor='center')

        steps = [
            "❶ 파일 선택",
            "❷ 파일 분석",
            "❸ 확인/수정",
            "❹ 변환 확정",
            "❺ 실행",
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

        # 힌트 라벨
        self.step_hint = tk.Label(inner, text="", font=("맑은 고딕", 9, "bold"),
                                  bg="#f0f4f8", fg="#e74c3c")
        self.step_hint.pack(side=tk.LEFT, padx=(15, 0))

    def _set_step(self, step_num):
        """현재 단계 설정 (0-based). 완료된 단계는 초록, 현재 단계는 파란색 깜빡임."""
        self._current_step = step_num
        hints = [
            "◀ 엑셀 파일을 선택하세요",
            "◀ '📊 파일 분석' 버튼을 클릭하세요",
            "◀ 자동 입력된 컬럼명·데이터를 확인 후 '✅ 변환 계획 확정' 클릭",
            "◀ '✅ 변환 계획 확정' 버튼을 클릭하세요",
            "◀ 저장 경로 확인 후 '▶ 합성 데이터 생성 실행' 클릭",
        ]

        for i, lbl in enumerate(self.step_labels):
            if i < step_num:
                lbl.config(bg="#27ae60", fg="white")  # 완료
            elif i == step_num:
                lbl.config(bg="#3498db", fg="white")  # 현재
            else:
                lbl.config(bg="#ddd", fg="#999")       # 미래

        # 화살표 색상 업데이트
        for i, arrow in enumerate(self.step_arrows):
            if i < step_num:
                arrow.config(fg="#27ae60")
            elif i == step_num:
                arrow.config(fg="#3498db")
            else:
                arrow.config(fg="#bbb")

        if step_num < len(hints):
            self.step_hint.config(text=hints[step_num])
        else:
            self.step_hint.config(text="✅ 완료!")

        self._start_blink()

    def _start_blink(self):
        if self._blink_id:
            self.root.after_cancel(self._blink_id)
        self._blink_state = True
        self._do_blink()

    def _do_blink(self):
        step = self._current_step
        if step < len(self.step_labels):
            lbl = self.step_labels[step]
            if self._blink_state:
                lbl.config(bg="#3498db", fg="white")
            else:
                lbl.config(bg="#1a6fb5", fg="#b8daef")
            self._blink_state = not self._blink_state

        # 힌트도 깜빡임
        if self.step_hint.cget('text'):
            cur_fg = self.step_hint.cget('fg')
            self.step_hint.config(fg="#e74c3c" if cur_fg != "#e74c3c" else "#f5b7b1")

        # 확정 버튼 깜빡임 (확인/수정 단계일 때)
        if step == 2 and hasattr(self, 'confirm_btn'):
            try:
                if self.confirm_btn.cget('state') != 'disabled':
                    if self._blink_state:
                        self.confirm_btn.config(bg="#27ae60")
                    else:
                        self.confirm_btn.config(bg="#1e8449")
            except Exception:
                pass

        # 실행 버튼 안내 강조 (실행 단계일 때)
        if step == 4 and hasattr(self, 'run_btn'):
            try:
                if self._blink_state:
                    self.status_lbl.config(foreground="#e74c3c")
                else:
                    self.status_lbl.config(foreground="#666")
            except Exception:
                pass

        self._blink_id = self.root.after(650, self._do_blink)

    # ══════════════════════════════════════════════════════════
    # 이벤트
    # ══════════════════════════════════════════════════════════

    def _browse_file(self):
        p = filedialog.askopenfilename(title="엑셀 파일 선택",
            filetypes=[("Excel", "*.xlsx *.xls *.xlsm"), ("All", "*.*")])
        if not p:
            return
        self.excel_path.set(p)
        self.save_dir.set(os.path.dirname(p))
        self.save_name.set(os.path.splitext(os.path.basename(p))[0] + "_합성데이터")
        self.status_lbl.config(text="시트 목록 로드 중...")
        self.root.update_idletasks()
        try:
            sheets, engine = get_sheet_names(p)
            self.sheet_combo['values'] = sheets
            if sheets:
                self.sheet_var.set(sheets[0])
            self.status_lbl.config(text=f"시트 {len(sheets)}개 감지 ({engine})")
            self._set_step(1)  # → 파일 분석 단계로
        except Exception as e:
            self.status_lbl.config(text="")
            messagebox.showerror("오류", f"시트 목록 로드 실패:\n{e}")

    def _browse_dir(self):
        d = filedialog.askdirectory(title="저장 폴더 선택")
        if d:
            self.save_dir.set(d)

    def _log(self, msg):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.root.update_idletasks()

    def _set_analysis(self, text):
        self.analysis_text.config(state=tk.NORMAL)
        self.analysis_text.delete("1.0", tk.END)
        self.analysis_text.insert(tk.END, text)
        self.analysis_text.config(state=tk.DISABLED)

    # ── 파일 분석 ────────────────────────────────────────

    def _analyze_file(self):
        path = self.excel_path.get().strip()
        if not path:
            messagebox.showwarning("경고", "엑셀 파일을 먼저 선택해 주세요.")
            return

        self.status_lbl.config(text="파일 로드 중...")
        self.root.update_idletasks()
        try:
            (self.df, self.info), engine = load_excel(path, self.sheet_var.get() or None)
        except Exception as e:
            self.status_lbl.config(text="")
            messagebox.showerror("로드 오류", str(e))
            return

        self.original_columns = list(self.df.columns)

        # ── 500행 초과 시 잘라내기 ──
        MAX_ROWS = 500
        if len(self.df) > MAX_ROWS:
            original_rows = len(self.df)
            self.df = self.df.iloc[:MAX_ROWS].reset_index(drop=True)
            self.info['rows'] = len(self.df)
            messagebox.showinfo(
                "행 수 제한 안내",
                f"원본 데이터가 {original_rows}행으로 제한({MAX_ROWS}행)을 초과합니다.\n\n"
                f"상위 {MAX_ROWS}행만 사용하고 나머지 {original_rows - MAX_ROWS}행은 제외됩니다.")

        self.col_types = {}
        type_counts = {'numerical': 0, 'datetime': 0, 'categorical': 0}
        for col in self.df.columns:
            t = auto_detect_column_type(self.df[col])
            self.col_types[col] = t
            type_counts[t] = type_counts.get(t, 0) + 1

        icons = {'numerical': '📐수치', 'datetime': '📅날짜', 'categorical': '🏷️범주'}
        col_list = "  |  ".join([f"{icons[t]} {str(c)}" for c, t in self.col_types.items()])
        summary = (f"파일: {self.info['file_name']}  |  시트: {self.info['sheet_name']}  |  "
                   f"{self.info['rows']}행 × {self.info['cols']}열  "
                   f"(수치 {type_counts['numerical']}, 날짜 {type_counts['datetime']}, "
                   f"범주 {type_counts['categorical']})\n"
                   f"{col_list}")
        self._set_analysis(summary)
        self._build_rename_widgets()
        self._auto_fill_column_names()   # 자동 컬럼명 입력
        self._build_input_widgets()
        self._auto_fill_all()            # 자동 가짜 데이터 입력
        self._set_step(2)                # → 확인/수정 단계
        self.status_lbl.config(text=f"분석 완료 ({engine}) — 컬럼명·데이터가 자동 입력되었습니다. 확인 후 실행하세요.")

    # ── 컬럼명 변경 위젯 동적 생성 ───────────────────────

    def _build_rename_widgets(self):
        """컬럼 갯수를 읽어 기존 컬럼명을 표시하고 새 이름 입력 Entry를 생성"""
        for w in self.rename_inner.winfo_children():
            w.destroy()
        self.col_rename_entries = {}

        if self.df is None:
            return

        # 헤더 행
        header = ttk.Frame(self.rename_inner)
        header.pack(fill=tk.X, padx=5, pady=(2, 4))
        ttk.Label(header, text="No.", width=5, font=("맑은 고딕", 9, "bold")).pack(side=tk.LEFT)
        ttk.Label(header, text="기존 컬럼명", width=25, font=("맑은 고딕", 9, "bold"),
                  anchor='w').pack(side=tk.LEFT, padx=(5, 0))
        ttk.Label(header, text="→", width=3, font=("맑은 고딕", 9, "bold")).pack(side=tk.LEFT)
        ttk.Label(header, text="새 컬럼명 (비우면 유지)", width=30,
                  font=("맑은 고딕", 9, "bold"), anchor='w').pack(side=tk.LEFT)
        ttk.Label(header, text="타입", width=8, font=("맑은 고딕", 9, "bold")).pack(side=tk.LEFT)

        ttk.Separator(self.rename_inner, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=5)

        icons = {'numerical': '📐', 'datetime': '📅', 'categorical': '🏷️'}
        for idx, col in enumerate(self.df.columns):
            row = ttk.Frame(self.rename_inner)
            row.pack(fill=tk.X, padx=5, pady=1)

            ttk.Label(row, text=f"{idx+1}", width=5, font=("Consolas", 9),
                      foreground="#999").pack(side=tk.LEFT)
            ttk.Label(row, text=str(col)[:30], width=25, anchor='w',
                      font=("Consolas", 9)).pack(side=tk.LEFT, padx=(5, 0))
            ttk.Label(row, text="→", width=3).pack(side=tk.LEFT)

            entry = ttk.Entry(row, font=("Consolas", 9), width=30)
            entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
            self.col_rename_entries[col] = entry

            col_type = self.col_types.get(col, 'categorical')
            icon = icons.get(col_type, '🏷️')
            ttk.Label(row, text=f"{icon}{col_type}", width=12,
                      font=("맑은 고딕", 8), foreground="#888").pack(side=tk.LEFT, padx=(5, 0))

    def _apply_column_rename(self):
        """사용자가 입력한 컬럼명을 적용"""
        if self.df is None:
            messagebox.showwarning("경고", "먼저 파일을 분석해 주세요.")
            return

        rename_map = {}
        for orig_col, entry in self.col_rename_entries.items():
            new_name = entry.get().strip()
            if new_name and new_name != orig_col:
                rename_map[orig_col] = new_name

        if not rename_map:
            messagebox.showinfo("알림", "변경할 컬럼이 없습니다.")
            return

        # 중복 검사
        new_names = list(rename_map.values())
        existing = [c for c in self.df.columns if c not in rename_map]
        all_names = existing + new_names
        if len(all_names) != len(set(all_names)):
            messagebox.showerror("오류", "컬럼명이 중복됩니다. 다시 확인해 주세요.")
            return

        # 적용
        self.df.rename(columns=rename_map, inplace=True)

        # col_types도 업데이트
        new_col_types = {}
        for old_col, ctype in self.col_types.items():
            new_col = rename_map.get(old_col, old_col)
            new_col_types[new_col] = ctype
        self.col_types = new_col_types

        # [Issue 5] col_entry_map 키도 동기화
        new_entry_map = {}
        for old_col, entries in self.col_entry_map.items():
            new_col = rename_map.get(old_col, old_col)
            new_entry_map[new_col] = entries
        self.col_entry_map = new_entry_map

        # UI 갱신
        self._build_rename_widgets()
        self._build_input_widgets()

        # 분석 텍스트 갱신
        icons = {'numerical': '📐수치', 'datetime': '📅날짜', 'categorical': '🏷️범주'}
        col_list = "  |  ".join([f"{icons[t]} {str(c)}" for c, t in self.col_types.items()])
        type_counts = {'numerical': 0, 'datetime': 0, 'categorical': 0}
        for t in self.col_types.values():
            type_counts[t] = type_counts.get(t, 0) + 1
        summary = (f"파일: {self.info['file_name']}  |  시트: {self.info['sheet_name']}  |  "
                   f"{self.info['rows']}행 × {self.info['cols']}열  "
                   f"(수치 {type_counts['numerical']}, 날짜 {type_counts['datetime']}, "
                   f"범주 {type_counts['categorical']})\n"
                   f"{col_list}")
        self._set_analysis(summary)

        changed_str = ", ".join([f"'{k}'→'{v}'" for k, v in rename_map.items()])
        self.status_lbl.config(text=f"컬럼명 변경 완료: {changed_str}")
        messagebox.showinfo("완료", f"{len(rename_map)}개 컬럼명이 변경되었습니다.\n\n" +
                            "\n".join([f"  {k} → {v}" for k, v in rename_map.items()]))

    def _apply_column_rename_silent(self):
        """컬럼명 적용 (메시지박스 없이 — 실행 시 자동 호출)"""
        if self.df is None:
            return
        rename_map = {}
        for orig_col, entry in self.col_rename_entries.items():
            new_name = entry.get().strip()
            if new_name and new_name != orig_col:
                rename_map[orig_col] = new_name
        if not rename_map:
            return
        # 중복 검사
        new_names = list(rename_map.values())
        existing = [c for c in self.df.columns if c not in rename_map]
        all_names = existing + new_names
        if len(all_names) != len(set(all_names)):
            return  # 중복 시 조용히 무시
        self.df.rename(columns=rename_map, inplace=True)
        new_col_types = {}
        for old_col, ctype in self.col_types.items():
            new_col = rename_map.get(old_col, old_col)
            new_col_types[new_col] = ctype
        self.col_types = new_col_types
        # [Issue 5] col_entry_map 키도 동기화
        new_entry_map = {}
        for old_col, entries in self.col_entry_map.items():
            new_col = rename_map.get(old_col, old_col)
            new_entry_map[new_col] = entries
        self.col_entry_map = new_entry_map

    # ── 컬럼명 자동 채우기 ──────────────────────────────

    def _auto_fill_column_names(self):
        """데이터 타입 기반 자동 컬럼명 생성: Text1, Num1, Date1, ..."""
        if self.df is None or not self.col_rename_entries:
            return
        counters = {'categorical': 0, 'numerical': 0, 'datetime': 0}
        prefix_map = {'categorical': 'Text', 'numerical': 'Num', 'datetime': 'Date'}
        for col, entry in self.col_rename_entries.items():
            col_type = self.col_types.get(col, 'categorical')
            counters[col_type] = counters.get(col_type, 0) + 1
            prefix = prefix_map.get(col_type, 'Col')
            new_name = f"{prefix}{counters[col_type]}"
            entry.delete(0, tk.END)
            entry.insert(0, new_name)

    # ── 1:1 입력 위젯 동적 생성 ──────────────────────────

    def _build_input_widgets(self):
        for w in self.input_inner.winfo_children():
            w.destroy()
        self.col_entry_map = {}

        str_cols = [c for c in self.df.columns
                    if (self.df[c].dtype == object or pd.api.types.is_string_dtype(self.df[c]))
                    and self.col_types.get(c) == 'categorical']

        if not str_cols:
            ttk.Label(self.input_inner, text="  입력이 필요한 문자열 컬럼이 없습니다.",
                      style="Sub.TLabel").pack(anchor='w', pady=10)
            return

        for col_idx, col in enumerate(str_cols):
            series = self.df[col].dropna()
            is_id  = _is_id_col(series)
            unique_vals = sorted([str(v) for v in series.unique() if v is not None])
            n = len(unique_vals)

            col_frame = ttk.Frame(self.input_inner)
            col_frame.pack(fill=tk.X, pady=(8 if col_idx > 0 else 0, 2))

            ttk.Label(col_frame,
                      text=f"📋 {col}  ({n}개)",
                      style="ColHeader.TLabel").pack(side=tk.LEFT)

            if is_id:
                ttk.Label(col_frame, text="  (ID형 — 자동 재생성)",
                          style="Sub.TLabel").pack(side=tk.LEFT, padx=(8, 0))
            else:
                def _auto_col(c=col):
                    self._auto_fill_column(c)
                ttk.Button(col_frame, text="🔄 자동 채우기",
                           command=_auto_col, style="Auto.TButton").pack(side=tk.RIGHT)

            ttk.Separator(self.input_inner, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=1)

            entries = []

            if is_id:
                for i, val in enumerate(unique_vals):
                    row = ttk.Frame(self.input_inner)
                    row.pack(fill=tk.X, padx=(20, 0), pady=1)
                    ttk.Label(row, text=f"{val}", width=25, anchor='w',
                              font=("Consolas", 9)).pack(side=tk.LEFT)
                    ttk.Label(row, text="→", width=3).pack(side=tk.LEFT)
                    fake_val = 'S' + str(i + 1).zfill(4)
                    lbl = ttk.Label(row, text=fake_val, font=("Consolas", 9), foreground="#2266aa")
                    lbl.pack(side=tk.LEFT)
                    entries.append((val, None, fake_val))
            else:
                for val in unique_vals:
                    row = ttk.Frame(self.input_inner)
                    row.pack(fill=tk.X, padx=(20, 0), pady=1)

                    ttk.Label(row, text=f"{str(val)[:30]}", width=25, anchor='w',
                              font=("Consolas", 9)).pack(side=tk.LEFT)
                    ttk.Label(row, text="→", width=3).pack(side=tk.LEFT)

                    entry = ttk.Entry(row, font=("Consolas", 9), width=30)
                    entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
                    entries.append((val, entry, None))

            self.col_entry_map[col] = entries

    # ── 자동 채우기 ──────────────────────────────────────

    def _auto_fill_column(self, col):
        if col not in self.col_entry_map:
            return
        entries = self.col_entry_map[col]
        n = len(entries)
        seed = abs(hash(col)) % 9999

        series = self.df[col].dropna().astype(str)
        if series.str.match(r'^[가-힣]{2,4}$').mean() > 0.7:
            pool = generate_fake_persons(n, seed=seed)
        elif str(col).lower() in ['client', '고객', '발주', '선주', 'company', '업체']:
            pool = generate_fake_companies(n, seed=seed)
        else:
            pool = generate_auto_codes(col, n, seed=seed)

        for i, (val, entry, fixed) in enumerate(entries):
            if entry is not None and not entry.get().strip():
                entry.delete(0, tk.END)
                entry.insert(0, pool[i] if i < len(pool) else f"{str(col)[:4]}_{i}")

    def _auto_fill_all(self):
        for col in self.col_entry_map:
            self._auto_fill_column(col)
        self.status_lbl.config(text="전체 자동 채우기 완료")

    # ── 매핑 수집 ────────────────────────────────────────

    def _collect_mappings(self):
        mapping_dict = {}
        for col, entries in self.col_entry_map.items():
            mapping = {}
            for val, entry, fixed in entries:
                if fixed is not None:
                    mapping[val] = fixed
                elif entry is not None:
                    fake = entry.get().strip()
                    if not fake:
                        idx = len(mapping)
                        alpha = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
                        fake = f"{str(col)[:4]}_{alpha[idx] if idx < 26 else str(idx)}"
                    mapping[val] = fake
            if mapping:
                mapping_dict[col] = mapping
        return mapping_dict

    # ── 합성 실행 ────────────────────────────────────────

    def _run(self):
        if self.df is None:
            messagebox.showwarning("경고", "먼저 '📊 파일 분석'을 실행해 주세요.")
            return
        if not self._data_confirmed:
            messagebox.showwarning("경고",
                "먼저 '✅ 변환 계획 확정' 버튼을 클릭하여\n변환 계획을 확정해 주세요.")
            return
        if not self.save_dir.get().strip():
            messagebox.showwarning("경고", "저장 경로를 선택해 주세요.")
            return
        if not self.save_name.get().strip():
            messagebox.showwarning("경고", "저장 파일명을 입력해 주세요.")
            return
        self.run_btn.config(state='disabled')
        self.progress['value'] = 0
        self.status_lbl.config(text="합성 중...")
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete("1.0", tk.END)
        self.log_text.config(state=tk.DISABLED)
        threading.Thread(target=self._worker, daemon=True).start()

    def _worker(self):
        try:
            self._do_synth()
        except Exception as e:
            self.root.after(0, lambda: self._log(f"\n❌ 오류: {e}"))
            self.root.after(0, lambda: messagebox.showerror("오류", str(e)))
        finally:
            self.root.after(0, lambda: self.run_btn.config(state='normal'))
            self.root.after(0, lambda: self.progress.configure(value=100))

    def _do_synth(self):
        L = lambda m: self.root.after(0, lambda msg=m: self._log(msg))
        S = lambda m: self.root.after(0, lambda msg=m: self.status_lbl.config(text=msg))
        P = lambda v: self.root.after(0, lambda val=v: self.progress.configure(value=val))

        df, ct = self.df, self.col_types

        L("=" * 55); L("  합성 데이터 생성 시작"); L("=" * 55)

        # 1. 매핑 수집
        P(0); S("[1/8] 매핑 수집..."); L("\n[1/8] 매핑 수집...")
        mapping_dict = self._collect_mappings()
        for col, mp in mapping_dict.items():
            L(f"  📋 {col}: {len(mp)}개 매핑")
            sample = list(mp.items())[:2]
            for o, f in sample:
                L(f"     {o[:15]} → {f[:15]}")

        # 2. 문자열 합성 (null 보존)
        P(12); S("[2/8] 문자열 합성..."); L("\n[2/8] 문자열 합성...")
        df_text, desc_map = synthesize_text_columns(df, ct, mapping_dict)
        L(f"  완료: {len(desc_map)}개 컬럼")

        # 3. 함수 종속성 감지 [Issue 3]
        P(25); S("[3/8] 종속성 분석..."); L("\n[3/8] 컬럼 간 함수 종속성 분석...")
        func_deps = detect_functional_dependencies(df, ct)
        if func_deps:
            for dep in func_deps:
                L(f"  🔗 {dep['from']} → {dep['to']}")
        else:
            L("  종속 관계 없음")

        # 4. 상관관계
        P(37); S("[4/8] 상관관계..."); L("\n[4/8] 상관관계 분석...")
        cr = analyze_correlations(df, ct)
        if cr.get('strong_pairs'):
            for p in cr['strong_pairs']:
                L(f"  {p['col1']} ↔ {p['col2']}  r={p['r']}")
        else:
            L("  강한 상관관계 없음")

        # 5. 제약
        P(50); S("[5/8] 제약조건..."); L("\n[5/8] 제약 조건...")
        cons = auto_detect_constraints(df, ct)
        for c in cons:
            if c['type'] == 'positive':      L(f"  [양수]  {c['column']}")
            elif c['type'] == 'range_0_100': L(f"  [0~100] {c['column']}")
            elif c['type'] == 'inequality':  L(f"  [순서]  {c['low']} < {c['high']}")
        if not cons:
            L("  없음")

        # 6. 수치/날짜 (null 비율 유지)
        P(62); S("[6/8] 수치 합성..."); L("\n[6/8] Gaussian Copula 합성...")
        nr_str = self.num_rows_var.get().strip()
        nr = int(nr_str) if nr_str.isdigit() else None
        syn_num = generate_numeric_datetime(df, ct, cons, nr)
        n = len(syn_num) if len(syn_num) > 0 else (nr or len(df))
        L(f"  {n}행 생성")

        # ── [Issue 8] 가중 샘플링: 원본 분포 비율 보존 ──
        P(75); S("[7/8] 행 조합..."); L("\n[7/8] 행 조합 (분포 보존)...")

        # 행별 가중치 계산 — 범주형 컬럼의 빈도 기반
        cat_cols = [c for c, t in ct.items() if t == 'categorical' and c in df_text.columns]
        if cat_cols and len(df_text) > 0:
            # 주요 범주 컬럼(고유값이 적은 것)의 빈도로 가중치 산출
            key_cat = [c for c in cat_cols
                       if df_text[c].nunique() <= max(20, len(df_text) * 0.15)]
            if key_cat:
                freq_col = key_cat[0]  # 가장 대표적인 범주 컬럼
                freqs = df_text[freq_col].value_counts(normalize=True)
                weights = df_text[freq_col].map(freqs).fillna(1.0 / len(df_text)).values
                weights = weights / weights.sum()
                idx = np.random.choice(len(df_text), size=n, replace=True, p=weights)
                L(f"  가중 샘플링 기준: '{freq_col}' (분포 보존)")
            else:
                idx = np.random.choice(len(df_text), size=n,
                                       replace=(n > len(df_text)))
                L("  균등 샘플링 (범주 컬럼 없음)")
        else:
            idx = np.random.choice(len(df_text), size=n,
                                   replace=(n > len(df_text)))

        final = df_text.iloc[idx].reset_index(drop=True)
        for c in syn_num.columns:
            final[c] = syn_num[c].values
        final = final[[c for c in df.columns if c in final.columns]]

        # ── [Issue 3] 함수 종속성 복원 ──
        if func_deps:
            L("\n  [종속성 복원]")
            for dep in func_deps:
                from_col, to_col = dep['from'], dep['to']
                if from_col not in final.columns or to_col not in final.columns:
                    continue
                # 원본 데이터의 종속 매핑 테이블
                dep_map = df.dropna(subset=[from_col, to_col]).drop_duplicates(
                    subset=[from_col]).set_index(from_col)[to_col].to_dict()
                # 합성 데이터에 매핑 적용 — from_col 값 기준으로 to_col 복원
                null_mask = final[to_col].isna()
                mapped = final[from_col].map(dep_map)
                # 매핑키에 해당 매핑이 있는 행만 적용 (text 치환 후이므로 fake 매핑도 확인)
                if from_col in mapping_dict:
                    fake_dep_map = {}
                    from_map = mapping_dict[from_col]
                    to_map = mapping_dict.get(to_col, {})
                    for orig_from, orig_to in dep_map.items():
                        fake_from = from_map.get(str(orig_from), str(orig_from))
                        fake_to = to_map.get(str(orig_to), str(orig_to))
                        fake_dep_map[fake_from] = fake_to
                    mapped = final[from_col].map(fake_dep_map)
                final[to_col] = mapped.where(mapped.notna(), final[to_col])
                final[to_col] = final[to_col].where(~null_mask, other=np.nan)
                L(f"  🔗 {from_col} → {to_col}: 종속 관계 복원 완료")

        # ── [Issue 4] ID 컬럼 유니크성 보장 ──
        for col in final.columns:
            if col in ct and ct[col] != 'categorical':
                continue
            if col not in df.columns:
                continue
            orig_series = df[col].dropna()
            if len(orig_series) == 0:
                continue
            if not _is_id_col(orig_series):
                continue
            # 원본에서 ID가 행마다 고유했는지 확인
            orig_unique_ratio = orig_series.nunique() / len(orig_series)
            if orig_unique_ratio < 0.9:
                continue  # 원본도 중복 ID가 있으면 그대로 유지
            # 가짜 ID 재생성 — 행 수만큼 고유 ID 생성
            null_mask = final[col].isna()
            # 기존 매핑에서 prefix 추출
            existing_fakes = [v for v in final[col].dropna().unique()]
            prefix = 'S'
            if existing_fakes:
                m = re.match(r'^([A-Za-z]+)', str(existing_fakes[0]))
                if m:
                    prefix = m.group(1)
            new_ids = [f'{prefix}{i+1:05d}' for i in range(len(final))]
            final[col] = new_ids
            final.loc[null_mask, col] = np.nan
            L(f"  🆔 {col}: {len(final)}개 고유 ID 재생성 ({prefix}00001~)")

        # 8. 품질
        P(88); S("[8/8] 품질 검증..."); L("\n[8/8] 품질 검증...")
        ov, cs = validate_quality(df, final, ct)
        for c, sc in cs.items():
            bar = '█' * int(sc * 20) + '░' * (20 - int(sc * 20))
            g = '✅' if sc >= 0.8 else ('⚠️' if sc >= 0.6 else '❌')
            L(f"  {g} {str(c):<20} {bar} {sc:.1%}")
        L(f"\n  종합 품질: {ov:.1%}")

        # ── 날짜 컬럼 정리: 순수 Date 또는 Time 형식으로 변환 ──
        S("날짜 정리..."); L("\n[날짜 형식 정리]")
        date_cols = [c for c, t in ct.items() if t == 'datetime']
        for col in date_cols:
            if col in final.columns:
                try:
                    cleaned = _clean_datetime_column(final[col])
                    final[col] = cleaned
                    L(f"  📅 {col}: 정리 완료")
                except Exception as e:
                    L(f"  ⚠️ {col}: 정리 실패 ({e})")

        # 저장
        P(95); S("저장 중..."); L("\n" + "─" * 55 + "\n  파일 저장...")
        bp = os.path.join(self.save_dir.get().strip(), self.save_name.get().strip())

        out_xl = bp + '.xlsx'
        final.to_excel(out_xl, index=False)
        L(f"  ✅ Excel       : {out_xl}")

        # ── 변환키 파일 저장 (컬럼명 변경 + 값 매핑) ──
        col_rename_map = {}
        for orig_col in self.original_columns:
            for current_col in self.df.columns:
                if orig_col != current_col:
                    # original_columns에 있지만 현재 컬럼에 없는 것 = 변경됨
                    pass
            # 더 정확하게: original_columns[i] vs df.columns[i]
        # original_columns와 현재 columns 비교
        col_rename_map = {}
        for i, orig_col in enumerate(self.original_columns):
            if i < len(self.df.columns):
                current_col = self.df.columns[i]
                if str(orig_col) != str(current_col):
                    col_rename_map[str(orig_col)] = str(current_col)

        key_data = {
            'generated_at': datetime.now().isoformat(),
            'source_file': self.info['file_path'],
            'sheet_name': self.info['sheet_name'],
            'original_columns': [str(c) for c in self.original_columns],
            'current_columns': [str(c) for c in self.df.columns],
            'column_rename': col_rename_map,
            'value_mapping': {col: mapping for col, mapping in mapping_dict.items()},
        }
        out_key = bp + '_변환키.json'
        with open(out_key, 'w', encoding='utf-8') as f:
            json.dump(key_data, f, ensure_ascii=False, indent=2, default=str)
        L(f"  ✅ 변환키      : {out_key}")

        out_desc = bp + '_description.json'
        with open(out_desc, 'w', encoding='utf-8') as f:
            json.dump(desc_map, f, ensure_ascii=False, indent=2, default=str)
        L(f"  ✅ Description : {out_desc}")

        out_rpt = bp + '_품질리포트.json'
        rpt = {
            'generated_at': datetime.now().isoformat(),
            'source_file': self.info['file_path'],
            'sheet_name': self.info['sheet_name'],
            'original_rows': len(df),
            'synthetic_rows': len(final),
            'column_types': ct,
            'constraints': cons,
            'quality_overall': ov,
            'quality_by_col': cs,
            'strong_correlations': cr.get('strong_pairs', []),
        }
        with open(out_rpt, 'w', encoding='utf-8') as f:
            json.dump(rpt, f, ensure_ascii=False, indent=2, default=str)
        L(f"  ✅ 품질리포트  : {out_rpt}")

        try:
            out_pq = bp + '.parquet'
            final.to_parquet(out_pq, index=False)
            L(f"  ✅ Parquet     : {out_pq}")
        except Exception:
            out_csv = bp + '.csv'
            final.to_csv(out_csv, index=False, encoding='utf-8-sig')
            L(f"  ✅ CSV         : {out_csv}")

        P(100)
        L("\n" + "=" * 55)
        L(f"  🎉 완료!  {len(final)}행  |  품질 {ov:.1%}")
        L("=" * 55)
        S(f"완료 — {len(final)}행, 품질 {ov:.1%}")
        self.root.after(0, lambda: self._set_step(5))  # → 완료 단계

        self.root.after(0, lambda: messagebox.showinfo("완료",
            f"합성 데이터 생성 완료!\n\n행 수: {len(final)}\n품질: {ov:.1%}\n\n"
            f"저장:\n{out_xl}\n{out_key}"))


if __name__ == "__main__":
    # PyInstaller exe 호환
    if getattr(sys, 'frozen', False):
        os.chdir(os.path.dirname(sys.executable))

    root = tk.Tk()
    SynthesizeApp(root)
    root.mainloop()
