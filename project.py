import streamlit as st
import pandas as pd
import numpy as np
import math
import re
import datetime
import io

# 1. 페이지 설정
st.set_page_config(page_title="MTA 기종 수요량 AI 분석 및 발주 최적화 대시보드", layout="wide")

# 스타일 설정
st.markdown("""
<style>
    [data-testid="stDataFrame"] table {
        font-size: 16px;
    }
</style>
""", unsafe_allow_html=True)

st.title("🛠️ MTA 기종 운영 수요량 AI 기반 분석 및 발주 최적화 대시보드")
st.markdown("엑셀 기반 수동 산출의 한계를 극복하고, **AI 파라미터(수주/견적 급등, 계약율, 계절성, CAPA, 총량 Capping)**를 반영하여 최적의 수요와 순발주량을 도출하는 프로그램입니다.")

# 2. 파일 업로드 및 글로벌 설정 영역
uploaded_file = st.file_uploader("수요 조정 엑셀 파일 업로드 (.xlsx)", type=['xlsx'])

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
    use_total_cap = st.checkbox("전체 총 발주량 제한(Capping) 활성화", value=False)
    max_total_order_limit = st.number_input("전체 최대 발주 허용 총량 (대):", value=20, step=1, disabled=not use_total_cap)

# 3. 유틸리티 함수 정의
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
        val = df_raw.iloc[22, 0]
        parsed = safe_float(val)
        if parsed is not None:
            return parsed
    except Exception:
        pass
    return 5.0

def extract_contract_rate_final(df):
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
    excel_file = pd.ExcelFile(io.BytesIO(file_bytes))
    sheet_names = excel_file.sheet_names
    
    st.markdown("---")
    st.markdown("### ⚙️ 기종별 개별 설정 (최대 CAPA · 현재고 · 프로모션 방어 여부)")
    
    machine_configs = {}
    for s in sheet_names:
        cols = st.columns([2, 2, 2, 2])
        cols[0].markdown(f"**[{s}]**")
        capa = cols[1].number_input(f"최대 CAPA ({s})", value=10.0, step=0.5, key=f"capa_{s}")
        inv = cols[2].number_input(f"현재고 ({s})", value=19, step=1, key=f"inv_{s}")
        promo = cols[3].checkbox(f"프로모션 방어 ({s})", value=False, key=f"promo_{s}")
        machine_configs[s] = {'capa': capa, 'inventory': inv, 'promo': promo}
    
    st.markdown("---")
    
    if st.button("🚀 기종별 맞춤 일괄 분석 및 비교 실행", type="primary"):
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
                    
                    years_raw = df.iloc[34, 1:].values
                    months_raw = df.iloc[35, 1:].values
                    
                    estimates_raw = df.iloc[36, 1:len(months_raw)+1].values
                    estimates = [safe_float(x) for x in estimates_raw]
                    valid_estimates = [x for x in estimates if x is not None]
                    
                    orders_raw = df.iloc[38, 1:len(months_raw)+1].values
                    orders = [safe_float(x) for x in orders_raw]
                    
                    trend_raw = df.iloc[39, 1:len(months_raw)+1].values
                    trend_orders = [safe_float(x) for x in trend_raw]
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
                            m_num = int(m_str[:-1].strip()) if m_str.endswith('월') else int(float(m_str))
                            if 1 <= m_num <= 12:
                                valid_data_pairs.append((datetime.datetime(yr, m_num, 1), ord_val))
                    
                    valid_data_pairs.sort(key=lambda x: x[0])
                    valid_orders = [x for _, x in valid_data_pairs] if valid_data_pairs else [x for x in orders if x is not None]
                    
                    avg_3m = np.mean(valid_orders[-3:]) if len(valid_orders) >= 3 else (np.mean(valid_orders) if valid_orders else 0.0)
                    avg_6m = np.mean(valid_orders[-6:]) if len(valid_orders) >= 6 else (np.mean(valid_orders) if valid_orders else 0.0)
                    latest_trend = valid_trend[-1] if valid_trend else avg_3m
                    
                    # [기본 AI 모델 산출] (수주·견적·계약율 조합)
                    if model_choice == 'elastic_capa':
                        raw_demand = (avg_3m * 0.4) + (avg_6m * 0.3) + (latest_trend * 0.3)
                    elif model_choice == 'trend_short':
                        raw_demand = latest_trend
                    else:
                        raw_demand = avg_6m
                    
                    base_pct = (raw_demand - prev_op_demand) / prev_op_demand if prev_op_demand > 0 else 0.0
                    rate_score = 0.0 if 0.20 <= excel_contract_rate <= 0.30 else (0.03 if excel_contract_rate > 0.30 else -0.05)
                    trend_score_basic = 0.05 if avg_3m > avg_6m else (-0.06 if avg_3m < avg_6m else 0.0)
                    
                    basic_ai_score = base_pct + rate_score + trend_score_basic
                    step = 0.05
                    basic_snapped_ratio = round(basic_ai_score / step) * step
                    basic_snapped_ratio = max(-0.15, min(0.15, basic_snapped_ratio))
                    
                    basic_demand_calc = prev_op_demand * (1.0 + basic_snapped_ratio)
                    basic_demand = min(basic_demand_calc, max_capa)
                    basic_target_inv = math.ceil(basic_demand * global_lead_time)
                    basic_net_order = max(0, basic_target_inv - current_inventory)

                    # [확장 파라미터 모델 산출] (상승 모멘텀 25% 기준 + 성수기 + CAPA 제어)
                    prev_order_val = valid_orders[-2] if len(valid_orders) >= 2 else valid_orders[-1]
                    latest_order_val = valid_orders[-1] if valid_orders else 0.0
                    order_mom_pct = ((latest_order_val - prev_order_val) / prev_order_val * 100) if prev_order_val > 0 else 0.0
                    is_boom_by_order = order_mom_pct >= 25.0

                    prev_est_val = valid_estimates[-2] if len(valid_estimates) >= 2 else (valid_estimates[-1] if valid_estimates else 0.0)
                    latest_est_val = valid_estimates[-1] if valid_estimates else 0.0
                    est_mom_pct = ((latest_est_val - prev_est_val) / prev_est_val * 100) if prev_est_val > 0 else 0.0
                    is_boom_by_est = est_mom_pct >= 25.0

                    is_auto_boom = is_boom_by_order or is_boom_by_est
                    
                    momentum_reasons = []
                    if is_promo:
                        momentum_reasons.append("프로모션 방어 설정 적용")
                    if is_boom_by_order:
                        momentum_reasons.append(f"단기 수주 급등 (+{order_mom_pct:.1f}%)")
                    if is_boom_by_est:
                        momentum_reasons.append(f"견적 추이 급등 (+{est_mom_pct:.1f}%)")
                    momentum_desc = " / ".join(momentum_reasons) if momentum_reasons else "조건 미충족 (변동 < 25%)"

                    seasonal_desc = "YoY 계절성 보통"
                    if valid_data_pairs:
                        latest_dt = valid_data_pairs[-1][0]
                        future_target_month_numbers = []
                        for i in range(round(global_lead_time) - 1, round(global_lead_time) + 2):
                            future_dt = latest_dt + datetime.timedelta(days=int(i * 30.44))
                            future_target_month_numbers.append(future_dt.month)
                        
                        historical_matching_orders = [ord_val for dt, ord_val in valid_data_pairs if dt.month in future_target_month_numbers]
                        overall_historical_avg = np.mean([ord_val for _, ord_val in valid_data_pairs]) if valid_data_pairs else avg_6m
                        
                        if historical_matching_orders and overall_historical_avg > 0:
                            matching_avg = np.mean(historical_matching_orders)
                            if matching_avg > overall_historical_avg * 1.10:
                                seasonal_desc = "YoY 계절성 성수기 (상향 요인)"
                            elif matching_avg < overall_historical_avg * 0.90:
                                seasonal_desc = "YoY 계절성 비수기 (하향 요인)"

                    trend_score_advanced = trend_score_basic + (0.05 if is_auto_boom else 0.0)
                    seasonal_score = 0.05 if "성수기" in seasonal_desc else (-0.05 if "비수기" in seasonal_desc else 0.0)
                    
                    advanced_ai_score = base_pct + rate_score + trend_score_advanced + seasonal_score
                    advanced_snapped_ratio = round(advanced_ai_score / step) * step
                    advanced_snapped_ratio = max(-0.15, min(0.15, advanced_snapped_ratio))

                    if (is_promo or is_auto_boom) and advanced_snapped_ratio < -0.05:
                        advanced_snapped_ratio = -0.05

                    advanced_demand_calc = prev_op_demand * (1.0 + advanced_snapped_ratio)
                    advanced_demand = min(advanced_demand_calc, max_capa)
                    advanced_target_inv = math.ceil(advanced_demand * global_lead_time)
                    raw_net_order = max(0, advanced_target_inv - current_inventory)

                    raw_analysis_results.append({
                        'sheet_name': sheet_name,
                        'max_capa': max_capa,
                        'current_inventory': current_inventory,
                        'basic_demand': basic_demand,
                        'basic_snapped_ratio': basic_snapped_ratio,
                        'basic_net_order': basic_net_order,
                        'advanced_demand': advanced_demand,
                        'advanced_snapped_ratio': advanced_snapped_ratio,
                        'raw_net_order': raw_net_order,
                        'momentum_desc': momentum_desc,
                        'seasonal_desc': seasonal_desc
                    })
                except Exception as e:
                    st.error(f"시트 '{sheet_name}' 분석 중 오류 발생: {e}")

            # 2단계: 전체 총 발주량 제한(Capping) 및 비율 안분 계산
            total_raw_orders = sum([r['raw_net_order'] for r in raw_analysis_results])
            
            summary_results = []
            for r in raw_analysis_results:
                sheet_name = r['sheet_name']
                raw_ord = r['raw_net_order']
                
                if use_total_cap and total_raw_orders > max_total_order_limit and total_raw_orders > 0:
                    proportional_share = (raw_ord / total_raw_orders) * max_total_order_limit
                    final_capped_order = round(proportional_share)
                else:
                    final_capped_order = raw_ord

                summary_results.append({
                    '기종(시트)': sheet_name,
                    '월 CAPA': r['max_capa'],
                    '현재고': r['current_inventory'],
                    '기본 모델 수요': round(r['basic_demand'], 2),
                    '기본 변동률': f"{r['basic_snapped_ratio']*100:+.0f}%",
                    '확장 맞춤 수요': round(r['advanced_demand'], 2),
                    '확장 변동률': f"{r['advanced_snapped_ratio']*100:+.0f}%",
                    '원래 산출 발주량': f"{raw_ord} 대",
                    '🚨 최종 조정 발주량': f"👉 {final_capped_order} 대"
                })

                with st.expander(f"📌 [{sheet_name}] 기본 AI ({r['basic_snapped_ratio']*100:+.0f}%) ➡️ 확장 파라미터 반영 ({r['advanced_snapped_ratio']*100:+.0f}%) 상세 진단 리포트"):
                    report_table_data = [
                        {"분석 비교 항목": "1. 기본 AI 모델 (수주·견적·계약율)", "운영수요량": f"{r['basic_demand']:.2f} 대", "변동률": f"{r['basic_snapped_ratio']*100:+.0f}%", "산출 발주량": f"{r['basic_net_order']} 대"},
                        {"분석 비교 항목": "2. 확장 맞춤 모델 (+모멘텀·성수기·CAPA)", "운영수요량": f"{r['advanced_demand']:.2f} 대", "변동률": f"{r['advanced_snapped_ratio']*100:+.0f}%", "산출 발주량": f"{raw_ord} 대"}
                    ]
                    st.dataframe(pd.DataFrame(report_table_data), use_container_width=True, hide_index=True)
                    st.markdown(f"- **상승 모멘텀 진단**: {r['momentum_desc']}")
                    st.markdown(f"- **계절성 진단**: {r['seasonal_desc']}")
                    st.markdown(f"- **최종 발주 확정**: 총량 Capping 반영 결과 **{final_capped_order}대**로 조정되었습니다.")

            st.markdown(f"### 📊 [비교 분석 종합 표] 기본 AI 모델 vs 확장 파라미터 반영 모델 (총 산출: {total_raw_orders}대)")
            if use_total_cap:
                st.info(f"💡 **총량 제한 적용중**: 전체 목표 발주 한도 **{max_total_order_limit}대**에 맞춰 각 기종별 발주량이 비례 배분(Capping)되었습니다.")
            
            df_summary = pd.DataFrame(summary_results)
            st.dataframe(df_summary, use_container_width=True, hide_index=True)
