import streamlit as st
import pandas as pd
import numpy as np
import math
import re
import datetime
import io

# 1. 페이지 설정
st.set_page_config(page_title="MTA 기종 운영수요량 자동 분석 프로그램", layout="wide")

# 스타일 설정 (expander 및 테이블 커스텀 스타일 정의)
st.markdown("""
<style>
    [data-testid="stDataFrame"] table {
        font-size: 16px;
    }
</style>
""", unsafe_allow_html=True)

st.title("🛠️ MTA 기종 운영수요량 자동 분석 프로그램")
st.markdown("**AI 파라미터(수주/견적 급등, 계약율, 계절성 퍼센트, CAPA 관리, 발주 총량)**를 반영하여 최적의 수요와 순발주량을 도출하는 프로그램")

# 2. 파일 업로드 및 글로벌 설정 영역
uploaded_file = st.file_uploader("엑셀 파일 업로드 (.xlsx)", type=['xlsx'])

col1, col2, col3 = st.columns(3)
with col1:
    global_lead_time = st.number_input("NC 리드타임:", value=6.7, step=0.1)
with col2:
    model_choice = st.selectbox(
        "수요 분석 모델:",
        options=[
            ('elastic_capa', '양방향 탄력 밴드 완충 모델 (상하방 유연 연동)'),
            ('trend_short', '단기 수주추이 중심'),
            ('trend_long', '중장기 안정세 중심')
        ],
        format_func=lambda x: x[1]
    )[0]
with col3:
    st.markdown("#### 🎯 총 발주량 제어 옵션")
    use_total_cap = st.checkbox("전체 총 발주량 제한 활성화", value=False)
    max_total_order_limit = st.number_input("MTA 기종 최대 발주량 (대):", value=20.0, step=1.0, disabled=not use_total_cap)

# 3. 유틸리티 함수 정의 (안전장치 강화)
def safe_float(val):
    if pd.isna(val): 
        return None
    if isinstance(val, (int, float)): 
        return float(val)
    val_str = str(val).strip()
    if val_str == '' or val_str in ['-', '↓', '↑', 'N/A', 'nan', 'None']: 
        return None
    try:
        if val_str.endswith('%'):
            return float(val_str[:-1].strip()) / 100.0
        return float(val_str)
    except ValueError:
        cleaned = ''.join(c for c in val_str if c.isdigit() or c in ['.', '-'])
        try:
            return float(cleaned) if cleaned else None
        except:
            return None

def extract_op_demand_from_a23(file_bytes, sheet_name):
    try:
        df_raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=None)
        if len(df_raw) > 22 and len(df_raw.columns) > 0:
            val = df_raw.iloc[22, 0]
            parsed = safe_float(val)
            if parsed is not None:
                return parsed
    except Exception:
        pass
    return 5.0

def extract_contract_rate_final(df):
    try:
        for r_idx in range(len(df)):
            for c_idx in range(len(df.columns)):
                cell_val = str(df.iloc[r_idx, c_idx])
                if '계약율' in cell_val or '계약률' in cell_val:
                    if '※' not in cell_val:
                        match = re.search(r'([\d.]+)\s*%', cell_val)
                        if match:
                            val = float(match.group(1))
                            return val / 100.0 if val > 1.0 else val
                        for offset in range(1, 4):
                            if c_idx + offset < len(df.columns):
                                right_val = df.iloc[r_idx, c_idx + offset]
                                if pd.notna(right_val):
                                    parsed = safe_float(right_val)
                                    if parsed is not None:
                                        return parsed / 100.0 if parsed > 1.0 else parsed
    except:
        pass
    
    try:
        for r in range(12, min(16, len(df))):
            for c in range(len(df.columns)):
                val = df.iloc[r, c]
                if pd.notna(val):
                    val_str = str(val)
                    if '%' in val_str and '※' not in val_str:
                        match = re.search(r'([\d.]+)\s*%', val_str)
                        if match:
                            num = float(match.group(1))
                            if num not in [90.0, 110.0]:
                                return num / 100.0 if num > 1.0 else num
    except:
        pass
    return 0.46

# 4. 파일이 업로드된 경우 메인 대시보드 실행
if uploaded_file is not None:
    file_bytes = uploaded_file.getvalue()
    try:
        excel_file = pd.ExcelFile(io.BytesIO(file_bytes))
        sheet_names = excel_file.sheet_names
    except Exception as e:
        st.error(f"엑셀 파일을 읽는 중 오류가 발생했습니다: {e}")
        sheet_names = []
    
    if sheet_names:
        st.markdown("---")
        st.markdown("### ⚙️ 기종별 개별 설정 (최대 CAPA · 현재고 · 프로모션 여부)")
        
        machine_configs = {}
        for s in sheet_names:
            cols = st.columns([2, 2, 2, 2])
            cols[0].markdown(f"**[{s}]**")
            capa = cols[1].number_input(f"최대 CAPA ({s})", value=10.0, step=1.0, key=f"capa_{s}")
            inv = cols[2].number_input(f"현재고 ({s})", value=19, step=1, key=f"inv_{s}")
            promo = cols[3].checkbox(f"프로모션 여부 ({s})", value=False, key=f"promo_{s}")
            machine_configs[s] = {'capa': capa, 'inventory': inv, 'promo': promo}
        
        st.markdown("---")
        
        if st.button("🚀 운영수요량 및 발주량 분석 및 비교 실행", type="primary"):
            raw_analysis_results = []
            
            with st.spinner("AI 분석 모델 및 파라미터 연산 진행 중..."):
                # 1단계: 각 기종별 독립 수치 선행 계산
                for sheet_name, cfg in machine_configs.items():
                    try:
                        max_capa = cfg['capa']
                        current_inventory = cfg['inventory']
                        is_promo = cfg['promo']
                        
                        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=None)
                        prev_op_demand = extract_op_demand_from_a23(file_bytes, sheet_name)
                        
                        # 안전한 행 데이터 추출 함수
                        def get_safe_row(row_idx):
                            try:
                                if row_idx < len(df):
                                    return df.iloc[row_idx, 1:].values
                            except:
                                pass
                            return []

                        years_raw = get_safe_row(34)
                        months_raw = get_safe_row(35)
                        
                        estimates_raw = get_safe_row(36)
                        estimates = [safe_float(x) for x in estimates_raw] if len(estimates_raw) > 0 else []
                        valid_estimates = [x for x in estimates if x is not None]
                        
                        orders_raw = get_safe_row(38)
                        orders = [safe_float(x) for x in orders_raw] if len(orders_raw) > 0 else []
                        
                        trend_raw = get_safe_row(39)
                        trend_orders = [safe_float(x) for x in trend_raw] if len(trend_raw) > 0 else []
                        valid_trend = [x for x in trend_orders if x is not None and x > 0]
                        
                        excel_contract_rate = extract_contract_rate_final(df)
                        
                        current_year = 2024
                        parsed_years = []
                        for y in years_raw:
                            if pd.notna(y):
                                digits = ''.join(filter(str.isdigit, str(y)))
                                if digits:
                                    try:
                                        current_year = int(digits)
                                    except:
                                        pass
                            parsed_years.append(current_year)
                        
                        valid_data_pairs = []
                        for yr, m_val, ord_val in zip(parsed_years, months_raw, orders):
                            if pd.notna(m_val) and ord_val is not None:
                                m_str = str(m_val).strip()
                                try:
                                    m_num = int(m_str[:-1].strip()) if m_str.endswith('월') else int(float(m_str))
                                    if 1 <= m_num <= 12:
                                        valid_data_pairs.append((datetime.datetime(yr, m_num, 1), ord_val))
                                except:
                                    pass
                        
                        valid_data_pairs.sort(key=lambda x: x[0])
                        valid_orders = [x for _, x in valid_data_pairs] if valid_data_pairs else [x for x in orders if x is not None]
                        
                        avg_3m
