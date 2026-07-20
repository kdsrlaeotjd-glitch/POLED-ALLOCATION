import io
import zipfile
import datetime
import numpy as np
import pandas as pd
import streamlit as st
import warnings

# 엑셀 형식 경고창 출력 방지
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# ==========================================================
# 0. Web UI 구성 및 기본 세팅 (무적 배정 엔진 v2.0 🍶)
# ==========================================================
st.set_page_config(page_title="폴레드 주문분배 시스템", page_icon="🍶", layout="wide")

SIDEBAR_LOGO_URL = "https://cdn-pro-web-223-233.cdn-nhncommerce.com/poled0304_godomall_com/data/skin/front/db_poled_C/img/dimg/about_logo02.png"

st.title("🍶 MADE BY DS ")
st.caption("Seosan & Yongma Multi-Warehouse Allocation Engine (v2.0 - Total Code Wash & Smart Naming)")
st.markdown("---")

# VIP 정상 8자리 특수코드 명부
ALLOWED_8DIGIT_CODES = [
    '10101101', '10101102', '10101105', '10101106', '10101108',
    '10101109', '10101110', '10101111', '10101112', '10101113',
    '10101103', '10101104', '10101107', '10101114', '10101115',
    '10101116', '10111108', '10111110', '10111112', '10111106',
    '10102102', '10102101'
]

# 제품코드 초강력 세척 함수 (에러 완벽 방어)
def clean_product_code(series):
    s = series.fillna("").astype(str).str.strip()
    s = s.str.replace(r'\.0$', '', regex=True)
    
    def remove_fake_zero(val):
        val_str = str(val).strip()
        if val_str.endswith('.0'):
            val_str = val_str[:-2]
            
        if val_str == "" or val_str.lower() == "nan":
            return ""
            
        if len(val_str) == 8 and val_str not in ALLOWED_8DIGIT_CODES:
            if val_str[-1] == '0':
                return val_str[:-1]
        elif len(val_str) == 6:
            if val_str[-1] == '0':
                return val_str[:-1]
                
        return val_str
        
    return s.apply(remove_fake_zero)

# 단포 / 단수합포 / 이종합포 자동 감지 함수
def get_pack_stats(df):
    if df is None or df.empty or '주문번호' not in df.columns:
        return {'단포': 0, '단수합포': 0, '이종합포': 0}
    
    needed = ['주문번호', '제품코드', '수량']
    for col in needed:
        if col not in df.columns:
            return {'단포': 0, '단수합포': 0, '이종합포': 0}

    grouped = df.groupby('주문번호')
    stats = {'단포': 0, '단수합포': 0, '이종합포': 0}
    for _, group in grouped:
        sku_cnt = group['제품코드'].nunique()
        total_qty = group['수량'].sum()
        if sku_cnt > 1: stats['이종합포'] += 1
        elif total_qty > 1: stats['단수합포'] += 1
        else: stats['단포'] += 1
    return stats

# ==========================================================
# 1. 세션 금고(Memory Vault)
# ==========================================================
if 'inventory_loaded' not in st.session_state:
    st.session_state['inventory_loaded'] = False
if 'stock_seosan' not in st.session_state:
    st.session_state['stock_seosan'] = {}
if 'stock_yongma' not in st.session_state:
    st.session_state['stock_yongma'] = {}
if 'order_count' not in st.session_state:
    st.session_state['order_count'] = 0
if 'history' not in st.session_state:
    st.session_state['history'] = []

# ==========================================================
# 2. 사이드바
# ==========================================================
with st.sidebar:
    st.image(SIDEBAR_LOGO_URL, width="stretch")
    st.markdown("---")
    st.header("🏢 1단계: 창고 재고 업로드")
    is_disabled = st.session_state['inventory_loaded']
    file_seosan = st.file_uploader("📂 서산창고 (B, L열)", type=['xlsx', 'xls'], disabled=is_disabled)
    file_yongma = st.file_uploader("📂 용마창고 (B, H열)", type=['xlsx', 'xls'], disabled=is_disabled)
    
    if st.button("📥 재고 확정", type="primary", disabled=is_disabled):
        if file_seosan and file_yongma:
            try:
                df_s = pd.read_excel(file_seosan, usecols="B,L", engine='xlrd' if file_seosan.name.endswith('.xls') else None)
                df_s.columns = ['제품코드', '재고수량']
                df_s['제품코드'] = clean_product_code(df_s['제품코드'])
                df_s['재고수량'] = pd.to_numeric(df_s['재고수량'], errors='coerce').fillna(0)
                df_s = df_s[df_s['제품코드'] != ""]
                st.session_state['stock_seosan'] = df_s.groupby('제품코드')['재고수량'].sum().to_dict()
                
                df_y = pd.read_excel(file_yongma, usecols="B,H", engine='xlrd' if file_yongma.name.endswith('.xls') else None)
                df_y.columns = ['제품코드', '재고수량']
                df_y['제품코드'] = clean_product_code(df_y['제품코드'])
                df_y['재고수량'] = pd.to_numeric(df_y['재고수량'], errors='coerce').fillna(0)
                df_y = df_y[df_y['제품코드'] != ""]
                st.session_state['stock_yongma'] = df_y.groupby('제품코드')['재고수량'].sum().to_dict()
                
                st.session_state['inventory_loaded'] = True
                st.success("✅ 재고 등록 완료!")
                st.rerun()
            except Exception as e:
                st.error(f"⚠️ 재고 로딩 에러: {e}")
    
    st.markdown("---")
    if st.button("🚨 당일 마감 & 초기화", type="secondary"):
        st.session_state.clear()
        st.success("🔄 초기화되었습니다.")
        st.rerun()

# ==========================================================
# 3. 메인 화면
# ==========================================================
c1, c2 = st.columns(2)
c1.info(f"🍶 **서산 잔여 품목:** {len(st.session_state['stock_seosan'])}개")
c2.info(f"🍶 **용마 잔여 품목:** {len(st.session_state['stock_yongma'])}개")

if not st.session_state['inventory_loaded']:
    st.warning("👈 좌측에서 재고를 먼저 등록해주세요.")
    st.stop()

st.header("📋 2단계: 발주서 분배 (연속 차감)")
priority_choice = st.radio("🍶 **우선 순위:**", ('서산창고 우선', '용마창고 우선'), horizontal=True)
priority_str = '서산' if '서산' in priority_choice else '용마'

file_order = st.file_uploader(f"📑 발주서 ({st.session_state['order_count']+1}차)", type=['xlsx', 'xls'])

if file_order and st.button("🚀 자동 분배 실행", type="primary"):
    try:
        with st.spinner("배정 로직 가동 중..."):
            try:
                orders_df = pd.read_excel(file_order, engine='xlrd' if file_order.name.endswith('.xls') else None)
            except:
                orders_df = pd.read_excel(file_order)

            orders_df.columns = orders_df.columns.str.strip()
            orig_columns = orders_df.columns.tolist()
            qty_col_name = orig_columns[18]
            
            col_A_str = orders_df.iloc[:, 0].astype(str).str.strip()
            col_B_str = orders_df.iloc[:, 1].astype(str).str.strip()
            pattern = r'\d{6}[a-zA-Z]{2}\d{3}'
            is_type1 = col_A_str.str.contains(pattern, na=False, regex=True)
            orders_df['주문번호'] = np.where(is_type1, col_A_str, col_B_str)
            
            # 💡 [핵심 해결] 원본 상품코드 열(10번째 열) 자체를 세척된 코드로 완벽 덮어쓰기!
            orig_pcode_col_name = orig_columns[9]
            orders_df[orig_pcode_col_name] = clean_product_code(orders_df.iloc[:, 9])
            orders_df['제품코드'] = orders_df[orig_pcode_col_name]
            
            orders_df['수량'] = pd.to_numeric(orders_df.iloc[:, 18], errors='coerce').fillna(0)
            
            gift_mask = orders_df['주문번호'].astype(str).str.contains('_사은품', na=False)
            orders_df = orders_df[~gift_mask].reset_index(drop=True)
            orders_df = orders_df[orders_df['제품코드'] != ""].reset_index(drop=True)
            orders_df['_orig_idx'] = orders_df.index
            
            total_stats = get_pack_stats(orders_df)
            
            temp_s = st.session_state['stock_seosan'].copy()
            temp_y = st.session_state['stock_yongma'].copy()
            results_map = {}
            
            pri_name, sec_name = (('서산', '용마') if priority_str == '서산' else ('용마', '서산'))
            pri_stock, sec_stock = ((temp_s, temp_y) if priority_str == '서산' else (temp_y, temp_s))
                
            grouped = list(orders_df.groupby('주문번호', sort=False))
            hold_pass1 = []
            for oid, group in grouped:
                items = group.to_dict('records')
                if all(pri_stock.get(it['제품코드'], 0) >= it['수량'] for it in items):
                    for it in items:
                        pc, q, idx = it['제품코드'], it['수량'], it['_orig_idx']
                        pri_stock[pc] = pri_stock.get(pc, 0) - q
                        res = {'주문번호': oid, '제품코드': pc, '수량': q, '서산배정': 0, '용마배정': 0, '상태': f'{pri_name} 완배'}
                        res[f'{pri_name}배정'] = q
                        results_map[idx] = res
                else: hold_pass1.append((oid, items))
                    
            hold_pass2 = []
            for oid, items in hold_pass1:
                if all(sec_stock.get(it['제품코드'], 0) >= it['수량'] for it in items):
                    for it in items:
                        pc, q, idx = it['제품코드'], it['수량'], it['_orig_idx']
                        sec_stock[pc] = sec_stock.get(pc, 0) - q
                        res = {'주문번호': oid, '제품코드': pc, '수량': q, '서산배정': 0, '용마배정': 0, '상태': f'{sec_name} 완배'}
                        res[f'{sec_name}배정'] = q
                        results_map[idx] = res
                else: hold_pass2.append((oid, items))
                    
            for oid, items in hold_pass2:
                reqs = {}
                for it in items: reqs[it['제품코드']] = reqs.get(it['제품코드'], 0) + it['수량']
                if all(tr <= (temp_s.get(pc, 0) + temp_y.get(pc, 0)) for pc, tr in reqs.items()):
                    for it in items:
                        pc, q, idx = it['제품코드'], it['수량'], it['_orig_idx']
                        av_s, av_y = temp_s.get(pc, 0), temp_y.get(pc, 0)
                        if priority_str == '서산':
                            t_s = min(q, av_s); temp_s[pc] = av_s - t_s; t_y = q - t_s; temp_y[pc] = av_y - t_y
                        else:
                            t_y = min(q, av_y); temp_y[pc] = av_y - t_y; t_s = q - t_y; temp_s[pc] = av_s - t_s
                        results_map[idx] = {'주문번호': oid, '제품코드': pc, '수량': q, '서산배정': t_s, '용마배정': t_y, '상태': '분할배정'}
                else:
                    for it in items:
                        idx = it['_orig_idx']
                        pc = it['제품코드']
                        total_required = reqs[pc]
                        total_avail = temp_s.get(pc, 0) + temp_y.get(pc, 0)
                        
                        if total_required > total_avail:
                            reason_str = '실재고부족'
                        else:
                            reason_str = '합배송품절'
                            
                        results_map[idx] = {'주문번호': oid, '제품코드': it['제품코드'], '수량': it['수량'], '서산배정': 0, '용마배정': 0, '상태': reason_str}
            
            results_list = [results_map[i] for i in range(len(orders_df))]
            st.session_state['stock_seosan'] = temp_s
            st.session_state['stock_yongma'] = temp_y
            st.session_state['order_count'] += 1
            
            list_s, list_y, list_un = [], [], []
            for i, row in enumerate(results_list):
                orig_row = orders_df.iloc[i].to_dict()
                if row['서산배정'] > 0:
                    r = orig_row.copy(); r[qty_col_name] = row['서산배정']; r['수량'] = row['서산배정']; list_s.append(r)
                if row['용마배정'] > 0:
                    r = orig_row.copy(); r[qty_col_name] = row['용마배정']; r['수량'] = row['용마배정']; list_y.append(r)
                if row['서산배정'] == 0 and row['용마배정'] == 0:
                    r = orig_row.copy(); r['[사유]'] = row['상태']; list_un.append(r)
            
            df_s = pd.DataFrame(list_s) if list_s else pd.DataFrame(columns=orders_df.columns)
            df_y = pd.DataFrame(list_y) if list_y else pd.DataFrame(columns=orders_df.columns)
            df_un = pd.DataFrame(list_un) if list_un else pd.DataFrame(columns=orders_df.columns.tolist() + ['[사유]'])
            
            s_stats = get_pack_stats(df_s); y_stats = get_pack_stats(df_y)
            
            st.session_state['history'].append({
                '차수': f"{st.session_state['order_count']}차",
                '서산 단포': s_stats['단포'], '서산 단수합포': s_stats['단수합포'], '서산 이종합포': s_stats['이종합포'],
                '용마 단포': y_stats['단포'], '용마 단수합포': y_stats['단수합포'], '용마 이종합포': y_stats['이종합포'],
                '미배정': df_un['주문번호'].nunique() if not df_un.empty else 0
            })
            
            today_str = datetime.datetime.now().strftime("%m%d")
            order_cnt = st.session_state['order_count']
            
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                for fn_label, dfd in [
                    (f"{today_str}_{order_cnt}차 서산.xlsx", df_s[orig_columns] if not df_s.empty else df_s), 
                    (f"{today_str}_{order_cnt}차 용마.xlsx", df_y[orig_columns] if not df_y.empty else df_y), 
                    (f"{today_str}_{order_cnt}차 미배정.xlsx", df_un), 
                    (f"{today_str}_{order_cnt}차 모니터링.xlsx", pd.DataFrame(results_list))
                ]:
                    eb = io.BytesIO()
                    dfd.to_excel(eb, index=False)
                    zf.writestr(fn_label, eb.getvalue())
            
            st.success(f"🎉 {st.session_state['order_count']}차 배정 완료!")
            
            st.subheader(f"📊 {st.session_state['order_count']}차 포장 유형 분석")
            rc1, rc2, rc3 = st.columns(3)
            with rc1: st.write("**📑 전체**"); st.write(f"단포: `{total_stats['단포']}` / 단수: `{total_stats['단수합포']}` / 이종: `{total_stats['이종합포']}`")
            with rc2: st.write("**🏢 서산**"); st.write(f"단포: `{s_stats['단포']}` / 단수: `{s_stats['단수합포']}` / 이종: `{s_stats['이종합포']}`")
            with rc3: st.write("**🏢 용마**"); st.write(f"단포: `{y_stats['단포']}` / 단수: `{y_stats['단수합포']}` / 이종: `{y_stats['이종합포']}`")
            
            zip_filename = f"{today_str}_{order_cnt}차.zip"
            st.download_button("💾 통합 다운로드", zip_buffer.getvalue(), zip_filename, "application/zip", width="stretch")
    except Exception as e:
        st.error(f"🚨 배정 중 중단됨: {e}")

# 누적 히스토리
if st.session_state['history']:
    st.markdown("---")
    st.subheader("📈 누적 배정 히스토리")
    st.dataframe(pd.DataFrame(st.session_state['history']), hide_index=True, width="stretch")
