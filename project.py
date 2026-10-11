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
                    
                    # [기본 모델 산출]
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

                    # [확장 파라미터 모델 산출]
                    prev_order_val = valid_orders[-2] if len(valid_orders) >= 2 else valid_orders[-1]
                    latest_order_val = valid_orders[-1] if valid_orders else 0.0
                    order_mom_pct = ((latest_order_val - prev_order_val) / prev_order_val * 100) if prev_order_val > 0 else 0.0
                    is_boom_by_order = order_mom_pct >= 20.0

                    prev_est_val = valid_estimates[-2] if len(valid_estimates) >= 2 else (valid_estimates[-1] if valid_estimates else 0.0)
                    latest_est_val = valid_estimates[-1] if valid_estimates else 0.0
                    est_mom_pct = ((latest_est_val - prev_est_val) / prev_est_val * 100) if prev_est_val > 0 else 0.0
                    is_boom_by_est = est_mom_pct >= 20.0

                    is_auto_boom = is_boom_by_order or is_boom_by_est
                    
                    momentum_reasons = []
                    if is_promo:
                        momentum_reasons.append("프로모션 방어 설정 적용")
                    if is_boom_by_order:
                        momentum_reasons.append(f"단기 수주 추이 급등 ({order_mom_pct:+.1f}%)")
                    if is_boom_by_est:
                        momentum_reasons.append(f"단기 견적 추이 급등 ({est_mom_pct:+.1f}%)")
                    
                    if not momentum_reasons:
                        momentum_desc = f"일반 변동 (수주 {order_mom_pct:+.1f}%, 견적 {est_mom_pct:+.1f}%)"
                    else:
                        momentum_desc = " / ".join(momentum_reasons)

                    # 🌟 모멘텀 방향성 판단 (상승: red, 하락: blue, 보합/기타: inherit)
                    max_mom_val = max(order_mom_pct, est_mom_pct)
                    min_mom_val = min(order_mom_pct, est_mom_pct)
                    if max_mom_val > 0:
                        mom_color = "#ff4b4b" # 빨간색 계열
                        mom_bg = "#ffeef0"
                    elif min_mom_val < 0:
                        mom_color = "#1c83e1" # 파란색 계열
                        mom_bg = "#eef6ff"
                    else:
                        mom_color = "#6c757d"
                        mom_bg = "#f8f9fa"

                    # 계절성 성수기/비수기 퍼센트 정밀 진단
                    seasonal_desc = "YoY 계절성 보통"
                    seasonal_diff_pct = 0.0
                    seasonal_type = "normal" 
                    
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
                            seasonal_diff_pct = ((matching_avg - overall_historical_avg) / overall_historical_avg) * 100.0
                            
                            if matching_avg > overall_historical_avg * 1.10:
                                seasonal_type = "boom"
                                seasonal_desc = f"YoY 계절성 성수기 (전체 평균 대비 {seasonal_diff_pct:+.1f}% 상향)"
                            elif matching_avg < overall_historical_avg * 0.90:
                                seasonal_type = "slump"
                                seasonal_desc = f"YoY 계절성 비수기 (전체 평균 대비 {seasonal_diff_pct:+.1f}% 하향)"
                            else:
                                seasonal_desc = f"YoY 계절성 보통 (전체 평균 대비 {seasonal_diff_pct:+.1f}% 편차)"

                    trend_score_advanced = trend_score_basic + (0.05 if is_auto_boom else 0.0)
                    seasonal_score = 0.05 if seasonal_type == "boom" else (-0.05 if seasonal_type == "slump" else 0.0)
                    
                    advanced_ai_score = base_pct + rate_score + trend_score_advanced + seasonal_score
                    advanced_snapped_ratio = round(advanced_ai_score / step) * step
                    advanced_snapped_ratio = max(-0.15, min(0.15, advanced_snapped_ratio))

                    if (is_promo or is_auto_boom) and advanced_snapped_ratio < -0.05:
                        advanced_snapped_ratio = -0.05

                    # CAPA 유연 수렴 로직
                    advanced_demand_calc = prev_op_demand * (1.0 + advanced_snapped_ratio)
                    is_exceeding_capa = advanced_demand_calc > max_capa
                    
                    if is_exceeding_capa:
                        advanced_demand = max_capa + (advanced_demand_calc - max_capa) * 0.4
                    else:
                        advanced_demand = advanced_demand_calc

                    advanced_target_inv = math.ceil(advanced_demand * global_lead_time)
                    raw_net_order = max(0, advanced_target_inv - current_inventory)

                    raw_analysis_results.append({
                        'sheet_name': sheet_name,
                        'max_capa': max_capa,
                        'current_inventory': current_inventory,
                        'prev_op_demand': prev_op_demand,
                        'basic_demand': basic_demand,
                        'basic_snapped_ratio': basic_snapped_ratio,
                        'basic_net_order': basic_net_order,
                        'advanced_demand': advanced_demand,
                        'advanced_snapped_ratio': advanced_snapped_ratio,
                        'raw_net_order': raw_net_order,
                        'momentum_desc': momentum_desc,
                        'mom_color': mom_color,
                        'mom_bg': mom_bg,
                        'seasonal_desc': seasonal_desc,
                        'seasonal_type': seasonal_type,
                        'is_exceeding_capa': is_exceeding_capa,
                        'original_calc_demand': advanced_demand_calc
                    })
                except Exception as e:
                    st.error(f"시트 '{sheet_name}' 분석 중 오류 발생: {e}")

            # 2단계: 전체 총 발주량 제한 계산
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
                    '기존 운영수요량': round(r['prev_op_demand'], 2),
                    'AI 기반 운영수요량': round(r['advanced_demand'], 2),
                    'AI 기반 변동률': f"{r['advanced_snapped_ratio']*100:+.0f}%",
                    '상태 진단': "🔥 활성" if r['seasonal_type']=='boom' else ("❄️ 위축" if r['seasonal_type']=='slump' else "정상"),
                    '원래 산출 발주량': f"{raw_ord} 대",
                    '🚨 최종 조정 발주량': f"👉 **{final_capped_order} 대**"
                })

                # 🌟 [상세 리포트 탭(Expander) 색상 동적 반영 기술]
                expander_border = r['mom_color']
                expander_bg = r['mom_bg']
                
                st.markdown(f"""
                <style>
                div.streamlit-expanderHeader {{
                    background-color: {expander_bg} !important;
                    border-left: 6px solid {expander_border} !important;
                    border-radius: 4px;
                }}
                </style>
                """, unsafe_allow_html=True)

                with st.expander(f"📌 [{sheet_name}] 기본 프로그램 ({r['basic_snapped_ratio']*100:+.0f}%) ➡️ 확장 파라미터 반영 ({r['advanced_snapped_ratio']*100:+.0f}%) 상세 진단 리포트"):
                    report_table_data = [
                        {"분석 비교 항목": "1. 기본 AI 모델 (수주·견적·계약율)", "운영수요량": f"{r['basic_demand']:.2f} 대", "변동률": f"{r['basic_snapped_ratio']*100:+.0f}%", "산출 발주량": f"{r['basic_net_order']} 대"},
                        {"분석 비교 항목": "2. 확장 파라메터 모델 (+모멘텀·성수기·CAPA 유연화)", "운영수요량": f"{r['advanced_demand']:.2f} 대", "변동률": f"{r['advanced_snapped_ratio']*100:+.0f}%", "산출 발주량": f"{raw_ord} 대"}
                    ]
                    st.dataframe(pd.DataFrame(report_table_data), use_container_width=True, hide_index=True)
                    
                    # 상승 모멘텀 텍스트 및 탭 컬러 강조
                    st.markdown(f"- **상승 모멘텀 진단**: <span style='color:{r['mom_color']}; font-weight:bold;'>{r['momentum_desc']}</span>", unsafe_allow_html=True)
                    
                    # 계절성 결과에 따른 컬러 박스 적용
                    if r['seasonal_type'] == 'boom':
                        st.success(f"📈 **계절성 진단**: {r['seasonal_desc']}")
                    elif r['seasonal_type'] == 'slump':
                        st.warning(f"📉 **계절성 진단**: {r['seasonal_desc']}")
                    else:
                        st.info(f"📊 **계절성 진단**: {r['seasonal_desc']}")
                    
                    # CAPA 경고 박스
                    if r['is_exceeding_capa']:
                        st.error(f"⚠️ **CAPA 초과 경고**: 설정된 최대 CAPA({r['max_capa']}대)를 초과하여 운영수요량({r['original_calc_demand']:.2f}대)이 산출되었습니다. 무조건적인 강제 컷오프 대신 CAPA 근처에서 유연하게 수렴하도록 조정되었으므로, 신중하게 고려하여 발주하시기 바랍니다.")
                    else:
                        st.success("✅ 최대 CAPA 내에서 운영수요량 운영 중입니다.")
                        
                    st.markdown(f"- **최종 발주 확정**: 최대 발주량 제한에 의해 **<span style='color:#ff4b4b; font-size:18px;'>{final_capped_order}대</span>**로 조정되었습니다.", unsafe_allow_html=True)

            st.markdown(f"### 📊 [비교 분석 종합 표] 기본 모델 vs 확장 파라미터 반영 모델 (총 산출: {total_raw_orders}대)")
            if use_total_cap:
                st.info(f"💡 **총량 제한 적용중**: 최대 발주량 **{max_total_order_limit}대**에 맞춰 각 기종별 발주량 산정됨.")
            
            df_summary = pd.DataFrame(summary_results)
            st.dataframe(df_summary, use_container_width=True, hide_index=True)
