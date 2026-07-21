import streamlit as st
import pandas as pd
import sqlite3
import datetime
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# 1. 웹 페이지 기본 레이아웃 및 타이틀
st.set_page_config(
    page_title="(주)진테크 스마트 부자재 통합 재고 관리 솔루션",
    page_icon="📦",
    layout="wide"
)

DB_NAME = "materials_advanced.db"

# 2. 데이터베이스 초기화 함수
def init_db():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS items (
            name TEXT PRIMARY KEY,
            specification TEXT,
            safety_stock INTEGER DEFAULT 0,
            price INTEGER DEFAULT 0
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS workers (
            name TEXT PRIMARY KEY
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT,
            name TEXT,
            division TEXT,
            qty INTEGER,
            manager TEXT,
            note TEXT
        )
    """)
    conn.commit()
    conn.close()

init_db()

# -------------------------------------------------------------------
# [사이드바 메뉴 설정]
# -------------------------------------------------------------------
with st.sidebar:
    try:
        st.image("logo.webp", use_container_width=True)
    except:
        st.markdown("### 🏬 JINTECH\nSmart System")
    
    st.markdown("---")
    menu = st.radio(
        "메뉴 선택",
        ["🔄 입출고 관리", "📊 실시간 재고 모니터링", "👥 시스템 마스터 등록 관리"]
    )
    st.markdown("<br><br><br><br><br><p style='font-size:12px; color:gray;'>Designer. 수진</p>", unsafe_allow_html=True)

# Helper DB 함수들
def get_items():
    conn = sqlite3.connect(DB_NAME)
    df = pd.read_sql_query("SELECT name, specification, safety_stock, price FROM items", conn)
    conn.close()
    return df

def get_workers():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("SELECT name FROM workers ORDER BY name ASC")
    workers = [r[0] for r in cursor.fetchall()]
    conn.close()
    return workers

def get_history(limit=50):
    conn = sqlite3.connect(DB_NAME)
    df = pd.read_sql_query(f"SELECT date AS '날짜', name AS '품목명', division AS '구분', qty AS '수량', manager AS '작업자/입고처', note AS '비고' FROM history ORDER BY id DESC LIMIT {limit}", conn)
    conn.close()
    return df

# -------------------------------------------------------------------
# 1. 🔄 입출고 관리 메뉴
# -------------------------------------------------------------------
if menu == "🔄 입출고 관리":
    st.title("🔄 입출고 등록 및 로그 관리 대시보드")
    
    col_input, col_right = st.columns([1, 1.3])
    
    items_df = get_items()
    item_list = items_df["name"].tolist() if not items_df.empty else ["등록된 품목 없음"]
    worker_list = get_workers() if get_workers() else ["등록된 작업자 없음"]

    with col_input:
        st.subheader("📝 데이터 장부 입력")
        with st.form("history_form", clear_on_submit=True):
            ent_date = st.date_input("날짜", datetime.date.today())
            
            selected_item = st.selectbox("품목 선택", item_list)
            
            # 단가 정보 실시간 안내
            current_price = 0
            if not items_df.empty and selected_item in items_df["name"].values:
                current_price = items_df.loc[items_df["name"] == selected_item, "price"].values[0]
            st.caption(f"💡 현재 선택 품목 단가: **{current_price:,} 원**")

            division = st.selectbox("구분", ["출고", "입고"])
            qty = st.number_input("수량", min_value=1, step=1, value=1)
            
            if division == "출고":
                manager = st.selectbox("작업자 선택 (가져가는 사람)", worker_list)
            else:
                manager = st.text_input("입고처 (주문/발주 업체명)", placeholder="예: 대성자재 (생략 시 '주문업체 수령')")
            
            note = st.text_input("비고 메모", placeholder="특이사항 기재")
            
            btn_submit = st.form_submit_button("📥 내역 등록 및 재고 반영", use_container_width=True, type="primary")
            
            if btn_submit:
                if selected_item == "등록된 품목 없음":
                    st.error("품목을 먼저 마스터에 등록해 주세요.")
                elif division == "출고" and manager == "등록된 작업자 없음":
                    st.error("작업자를 먼저 마스터에 등록해 주세요.")
                else:
                    final_manager = manager if manager else ("주문업체 수령" if division == "입고" else "")
                    conn = sqlite3.connect(DB_NAME)
                    cursor = conn.cursor()
                    cursor.execute("""
                        INSERT INTO history (date, name, division, qty, manager, note)
                        VALUES (?, ?, ?, ?, ?, ?)
                    """, (str(ent_date), selected_item, division, qty, final_manager, note))
                    conn.commit()
                    conn.close()
                    st.success(f"[{selected_item}] {qty}개 {division} 등록 완료!")
                    st.rerun()

    with col_right:
        st.subheader("📊 요약 품목별 현재고 실시간 현황")
        # 현재고 계산
        conn = sqlite3.connect(DB_NAME)
        history_all = pd.read_sql_query("SELECT name, division, qty FROM history", conn)
        conn.close()
        
        stock_summary = []
        for _, row in items_df.iterrows():
            p_name = row["name"]
            in_qty = history_all[(history_all["name"] == p_name) & (history_all["division"] == "입고")]["qty"].sum()
            out_qty = history_all[(history_all["name"] == p_name) & (history_all["division"] == "출고")]["qty"].sum()
            curr_stock = in_qty - out_qty
            stock_summary.append({"품목명": p_name, "현재 재고량": f"{curr_stock:,}"})
            
        st.dataframe(pd.DataFrame(stock_summary), use_container_width=True, hide_index=True)

        st.subheader("📜 입출고 추적 로그 내역 (최근 50건)")
        st.dataframe(get_history(50), use_container_width=True, hide_index=True)

# -------------------------------------------------------------------
# 2. 📊 실시간 재고 모니터링 메뉴
# -------------------------------------------------------------------
elif menu == "📊 실시간 재고 모니터링":
    st.title("📊 창고 내 스마트 통합 관제실")
    
    monitor_tab1, monitor_tab2, monitor_tab3 = st.tabs(["📦 현재고 종합 현황판", "📥 입고 내역 추적 장부", "📤 출고 내역 추적 장부"])
    
    conn = sqlite3.connect(DB_NAME)
    items_df = get_items()
    history_df = pd.read_sql_query("SELECT * FROM history", conn)
    conn.close()

    with monitor_tab1:
        st.subheader("📦 창고 내 부자재 현재고 및 상태 현황판")
        full_stock_data = []
        for _, row in items_df.iterrows():
            p_name = row["name"]
            in_q = history_df[(history_df["name"] == p_name) & (history_df["division"] == "입고")]["qty"].sum()
            out_q = history_df[(history_df["name"] == p_name) & (history_df["division"] == "출고")]["qty"].sum()
            c_stock = in_q - out_q
            s_stock = row["safety_stock"]
            status = "정상" if c_stock >= s_stock else "🚨 재고 부족"
            full_stock_data.append({
                "품목명": p_name,
                "규격 / 단위": row["specification"],
                "현재고 수량": f"{c_stock:,}",
                "안전 재고 기준": f"{s_stock:,}",
                "등록 단가(원)": f"{row['price']:,}",
                "상태 정보": status
            })
        st.dataframe(pd.DataFrame(full_stock_data), use_container_width=True, hide_index=True)

    with monitor_tab2:
        st.subheader("📥 입고 내역 추적 장부")
        in_df = history_df[history_df["division"] == "입고"].copy()
        if not in_df.empty:
            in_df = in_df.merge(items_df[["name", "price"]], on="name", how="left")
            in_df["입고금액(원)"] = in_df["qty"] * in_df["price"]
            in_df = in_df[["date", "name", "qty", "price", "입고금액(원)", "manager", "note"]]
            in_df.columns = ["입고 날짜", "품목명", "수량(개)", "단가(원)", "입고금액(원)", "발주/입고처", "비고 메모"]
            st.dataframe(in_df, use_container_width=True, hide_index=True)
        else:
            st.info("누적된 입고 내역이 없습니다.")

    with monitor_tab3:
        st.subheader("📤 출고 내역 추적 장부 (작업자)")
        out_df = history_df[history_df["division"] == "출고"].copy()
        if not out_df.empty:
            out_df = out_df.merge(items_df[["name", "price"]], on="name", how="left")
            out_df["출고금액(원)"] = out_df["qty"] * out_df["price"]
            out_df = out_df[["date", "name", "qty", "price", "출고금액(원)", "manager", "note"]]
            out_df.columns = ["출고 일자", "품목명", "수량(개)", "단가(원)", "출고금액(원)", "작업자 명단", "현장 특이사항"]
            st.dataframe(out_df, use_container_width=True, hide_index=True)
        else:
            st.info("누적된 출고 내역이 없습니다.")

# -------------------------------------------------------------------
# 3. 👥 시스템 마스터 등록 관리 메뉴 (✨ 완벽 삭제 기능 추가!)
# -------------------------------------------------------------------
elif menu == "👥 시스템 마스터 등록 관리":
    st.title("👥 진테크 마스터 인프라 등록 관리 통제실")
    
    col_master_item, col_master_worker = st.columns([1.2, 1])
    
    # --- [왼쪽: 부자재 품목 마스터 등록 및 삭제] ---
    with col_master_item:
        st.subheader("📦 1. 부자재 품목 마스터 정보 등록")
        with st.form("master_item_form", clear_on_submit=True):
            m_name = st.text_input("품목명 (중복불가)", placeholder="예: 가스켓")
            m_spec = st.text_input("규격 / 단위 (예: 40*40)", placeholder="1ea")
            m_safety = st.number_input("안전재고 기준량 수치", min_value=0, value=0)
            m_price = st.number_input("개당 부자재 단가 입력 (원)", min_value=0, value=0)
            
            btn_add_item = st.form_submit_button("➕ 새 부자재 등록", use_container_width=True, type="primary")
            
            if btn_add_item:
                if not m_name.strip():
                    st.warning("품목명을 입력해 주세요.")
                else:
                    conn = sqlite3.connect(DB_NAME)
                    cursor = conn.cursor()
                    try:
                        cursor.execute("INSERT INTO items (name, specification, safety_stock, price) VALUES (?, ?, ?, ?)",
                                       (m_name.strip(), m_spec.strip(), m_safety, m_price))
                        conn.commit()
                        st.success(f"[{m_name}] 품목 등록 완료!")
                        st.rerun()
                    except sqlite3.IntegrityError:
                        st.error("이미 존재하는 품목명입니다.")
                    finally:
                        conn.close()

        st.markdown("---")
        st.subheader("📋 등록된 부자재 목록 및 삭제")
        items_df = get_items()
        st.dataframe(items_df, use_container_width=True, hide_index=True)
        
        # ✨ [버그 수정 1] 부자재 품목 삭제 기능 구현
        if not items_df.empty:
            del_item_target = st.selectbox("삭제할 부자재 선택", items_df["name"].tolist(), key="del_item_select")
            if st.button("❌ 선택 품목 삭제", type="primary", use_container_width=True):
                conn = sqlite3.connect(DB_NAME)
                cursor = conn.cursor()
                cursor.execute("DELETE FROM items WHERE name = ?", (del_item_target,))
                conn.commit()
                conn.close()
                st.success(f"[{del_item_target}] 품목이 안전하게 삭제되었습니다.")
                st.rerun()

    # --- [오른쪽: 작업자 마스터 등록 및 삭제] ---
    with col_master_worker:
        st.subheader("👥 2. 작업자(직원) 마스터 등록")
        with st.form("master_worker_form", clear_on_submit=True):
            w_name = st.text_input("작업자 직원 이름 (중복불가)", placeholder="이름 입력")
            btn_add_worker = st.form_submit_button("👥 새 작업자 등록", use_container_width=True, type="primary")
            
            if btn_add_worker:
                if not w_name.strip():
                    st.warning("작업자 이름을 입력해 주세요.")
                else:
                    conn = sqlite3.connect(DB_NAME)
                    cursor = conn.cursor()
                    try:
                        cursor.execute("INSERT INTO workers (name) VALUES (?)", (w_name.strip(),))
                        conn.commit()
                        st.success(f"[{w_name}] 작업자 등록 완료!")
                        st.rerun()
                    except sqlite3.IntegrityError:
                        st.error("이미 존재하는 작업자 이름입니다.")
                    finally:
                        conn.close()

        st.markdown("---")
        st.subheader("📋 등록된 작업자 명단 및 삭제")
        workers_list = get_workers()
        st.dataframe(pd.DataFrame(workers_list, columns=["등록된 작업자 명단"]), use_container_width=True, hide_index=True)
        
        # ✨ [버그 수정 2] 작업자 삭제 기능 구현
        if workers_list:
            del_worker_target = st.selectbox("삭제할 작업자 선택", workers_list, key="del_worker_select")
            if st.button("❌ 선택 작업자 삭제", type="primary", use_container_width=True):
                conn = sqlite3.connect(DB_NAME)
                cursor = conn.cursor()
                cursor.execute("DELETE FROM workers WHERE name = ?", (del_worker_target,))
                conn.commit()
                conn.close()
                st.success(f"[{del_worker_target}] 작업자가 완전히 삭제되었습니다.")
                st.rerun()
