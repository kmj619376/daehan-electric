import io
import json
import sqlite3
import hashlib
import uuid
import random
from datetime import datetime
import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from PIL import Image

# ------------------------------------------------------------------------------
# 🔑 API Key 안전 연동 (Secrets에서 불러오기)
# ------------------------------------------------------------------------------
GEMINI_API_KEY = st.secrets.get("GEMINI_API_KEY", "")

MARGIN_RATE = 0.10          # 자재 마진율 10%
LABOR_RATE_MAIN = 260000    # 메인 차단기 노무비
LABOR_RATE_BRANCH = 15000   # 분기 차단기 노무비
EXTRA_SHIPPING = 100000     # 현장 운반비 및 양중비

DEFAULT_COLUMNS = ["분전반명", "구분", "종류", "극수", "용량", "부하명", "수량", "단가"]

# ------------------------------------------------------------------------------
# 🌐 접속자 IP 주소 추출 함수
# ------------------------------------------------------------------------------
def get_remote_ip():
    try:
        headers = st.context.headers
        if "X-Forwarded-For" in headers:
            return headers["X-Forwarded-For"].split(",")[0].strip()
        elif "x-forwarded-for" in headers:
            return headers["x-forwarded-for"].split(",")[0].strip()
        return "알수없음"
    except Exception:
        return "127.0.0.1"

# ------------------------------------------------------------------------------
# 🗄️ Database (SQLite) 설정 및 초기화
# ------------------------------------------------------------------------------
DB_FILE = "users.db"

def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    
    c.execute('''
        CREATE TABLE IF NOT EXISTS users (
            username TEXT PRIMARY KEY,
            password TEXT NOT NULL,
            name TEXT NOT NULL,
            role TEXT NOT NULL,
            status TEXT NOT NULL,
            ip_address TEXT,
            session_id TEXT,
            pin_code TEXT,
            created_at TEXT NOT NULL
        )
    ''')
    
    for col in ["ip_address", "session_id", "pin_code"]:
        try:
            c.execute(f"ALTER TABLE users ADD COLUMN {col} TEXT")
        except Exception:
            pass

    c.execute('''
        CREATE TABLE IF NOT EXISTS usage_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL,
            file_name TEXT NOT NULL,
            item_count INTEGER NOT NULL,
            ip_address TEXT,
            analyzed_at TEXT NOT NULL
        )
    ''')
    
    try:
        c.execute("ALTER TABLE usage_logs ADD COLUMN ip_address TEXT")
    except Exception:
        pass
    
    admin_id = "syd1007"
    admin_pass = hashlib.sha256("kmj851007".encode()).hexdigest()
    
    c.execute("SELECT * FROM users WHERE username=?", (admin_id,))
    if not c.fetchone():
        c.execute('''
            INSERT INTO users (username, password, name, role, status, ip_address, session_id, pin_code, created_at)
            VALUES (?, ?, '최고관리자', 'admin', 'approved', '관리자PC', '', '000000', ?)
        ''', (admin_id, admin_pass, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    else:
        c.execute('''
            UPDATE users SET password=?, role='admin', status='approved' WHERE username=?
        ''', (admin_pass, admin_id))
    
    conn.commit()
    conn.close()

def hash_pw(pw):
    return hashlib.sha256(pw.encode()).hexdigest()

init_db()

# ------------------------------------------------------------------------------
# 1. 페이지 설정
# ------------------------------------------------------------------------------
st.set_page_config(
    page_title="대한일렉트릭 견적 프로그램",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed"
)

user_ip = get_remote_ip()

# ------------------------------------------------------------------------------
# 2. URL 세션 기반 자동 로그인 복원 & 동시 접속 통제 (F5 새로고침 방지)
# ------------------------------------------------------------------------------
url_session = st.query_params.get("session", None)

# F5 새로고침 시 URL의 세션 토큰으로 로그인 복원
if url_session and not st.session_state.get('logged_in', False):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT username, name, role, status, session_id FROM users WHERE session_id=?", (url_session,))
    user_db = c.fetchone()
    conn.close()
    
    if user_db and user_db[3] == "approved":
        st.session_state['logged_in'] = True
        st.session_state['user_info'] = {"username": user_db[0], "name": user_db[1], "role": user_db[2], "ip": user_ip, "session": url_session}

# 다른 기기에서 새로 로그인하여 DB 세션 ID가 변경된 경우 즉시 강제 로그아웃
if st.session_state.get('logged_in', False):
    current_user_id = st.session_state['user_info']['username']
    current_session_id = st.session_state['user_info'].get('session', '')
    
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT session_id FROM users WHERE username=?", (current_user_id,))
    db_session = c.fetchone()
    conn.close()
    
    if db_session and db_session[0] != current_session_id:
        st.session_state['logged_in'] = False
        st.session_state['user_info'] = None
        st.query_params.clear()
        st.error("🚨 다른 기기(PC/모바일)에서 동일한 계정으로 로그인되어 현재 접속이 강제 종료되었습니다.")
        st.stop()

# ------------------------------------------------------------------------------
# 3. 로그인 / 회원가입 UI
# ------------------------------------------------------------------------------
if not st.session_state.get('logged_in', False):
    st.title("⚡ 대한일렉트릭 견적 프로그램")
    st.subheader("🔒 사용자 인증 및 승인이 되어야 접속 가능")
    st.caption(f"🖥️ 현재 접속 IP: **{user_ip}**")
    
    tab1, tab2 = st.tabs(["🔑 로그인", "📝 회원가입 신청"])
    
    with tab1:
        st.markdown("### 로그인")
        
        with st.form(key="login_form"):
            login_id = st.text_input("아이디 (ID)", key="login_id")
            login_pw = st.text_input("비밀번호", type="password", key="login_pw")
            submit_login = st.form_submit_button("로그인하기", type="primary", use_container_width=True)
            
        if submit_login:
            conn = sqlite3.connect(DB_FILE)
            c = conn.cursor()
            c.execute("SELECT username, name, role, status FROM users WHERE username=? AND password=?", 
                      (login_id, hash_pw(login_pw)))
            user = c.fetchone()
            conn.close()
            
            if user:
                username, name, role, status = user
                if status == "approved":
                    new_session = str(uuid.uuid4())
                    conn = sqlite3.connect(DB_FILE)
                    c = conn.cursor()
                    c.execute("UPDATE users SET ip_address=?, session_id=? WHERE username=?", (user_ip, new_session, username))
                    conn.commit()
                    conn.close()
                    
                    st.session_state['logged_in'] = True
                    st.session_state['user_info'] = {"username": username, "name": name, "role": role, "ip": user_ip, "session": new_session}
                    
                    # URL 파라미터에 세션 등록하여 F5 누적 방어
                    st.query_params["session"] = new_session
                    st.success(f"{name}님, 환영합니다!")
                    st.rerun()
                elif status == "pending":
                    st.warning("⏳ 아직 관리자 승인 대기 중인 계정입니다. 관리자가 가입을 승인해야 이용할 수 있습니다.")
                else:
                    st.error("🚫 사용이 차단되거나 비활성화된 계정입니다.")
            else:
                st.error("❌ 아이디 또는 비밀번호가 일치하지 않습니다.")
                
    with tab2:
        st.markdown("### 회원가입 신청")
        st.info("💡 회원가입 시 대표님(관리자)에게 사전 전달받은 **2차 승인 핀코드**를 입력하셔야 신청이 완료됩니다.")
        
        with st.form(key="register_form"):
            reg_id = st.text_input("사용할 아이디 (ID)", key="reg_id")
            reg_name = st.text_input("이름 / 회사명", key="reg_name")
            reg_pw = st.text_input("비밀번호", type="password", key="reg_pw")
            reg_pw_confirm = st.text_input("비밀번호 확인", type="password", key="reg_pw_confirm")
            reg_pin = st.text_input("2차 승인 핀코드 (관리자에게 전달받은 코드)", type="password", key="reg_pin")
            submit_reg = st.form_submit_button("가입 신청하기", type="primary", use_container_width=True)
            
        if submit_reg:
            if not reg_id or not reg_name or not reg_pw or not reg_pin:
                st.error("모든 항목과 2차 승인 핀 코드를 입력해 주세요.")
            elif reg_pw != reg_pw_confirm:
                st.error("비밀번호가 일치하지 않습니다.")
            else:
                conn = sqlite3.connect(DB_FILE)
                c = conn.cursor()
                
                c.execute("SELECT * FROM users WHERE username=?", (reg_id,))
                if c.fetchone():
                    st.error("이미 존재하는 아이디입니다.")
                    conn.close()
                else:
                    c.execute("SELECT username FROM users WHERE pin_code=?", (reg_pin,))
                    matched_admin_pin = c.fetchone()
                    
                    if matched_admin_pin or reg_pin == "000000":
                        c.execute('''
                            INSERT INTO users (username, password, name, role, status, ip_address, pin_code, created_at)
                            VALUES (?, ?, ?, 'user', 'pending', ?, ?, ?)
                        ''', (reg_id, hash_pw(reg_pw), reg_name, user_ip, reg_pin, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
                        conn.commit()
                        conn.close()
                        st.success("🎉 가입 신청이 성공적으로 완료되었습니다! 대표님(관리자)이 승인해 주시면 로그인이 가능합니다.")
                    else:
                        conn.close()
                        st.error("❌ 유효하지 않은 2차 승인 핀코드입니다. 대표님(관리자)에게 전달받은 코드를 확인해 주세요.")
    st.stop()

# ------------------------------------------------------------------------------
# 4. 메인 프로그램 화면
# ------------------------------------------------------------------------------
user = st.session_state['user_info']

col_h1, col_h2 = st.columns([8, 2])
with col_h1:
    st.title("⚡ 대한일렉트릭 견적 프로그램")
    st.caption(f"접속 계정: **{user['name']} ({user['username']})** [{user['role'].upper()}] | 접속 IP: **{user.get('ip', user_ip)}**")
with col_h2:
    st.write("")
    if st.button("🚪 로그아웃", type="secondary"):
        st.session_state['logged_in'] = False
        st.session_state['user_info'] = None
        st.query_params.clear()
        st.rerun()

# ------------------------------------------------------------------------------
# 👑 관리자 전용 메뉴
# ------------------------------------------------------------------------------
if user['role'] == 'admin':
    with st.expander("👑 [관리자 전용] 회원 승인 및 가입용 2차 PIN 관리", expanded=True):
        admin_tab1, admin_tab2 = st.tabs(["👥 회원 승인 관리", "📜 이용 이력(로그) 보기"])
        
        with admin_tab1:
            st.subheader("회원가입 신청 승인 및 상태 관리")
            conn = sqlite3.connect(DB_FILE)
            df_users = pd.read_sql_query("SELECT username AS 아이디, name AS 이름, role AS 권한, status AS 상태, pin_code AS 가입시사용한PIN, ip_address AS 접속IP, created_at AS 가입일시 FROM users", conn)
            st.dataframe(df_users, use_container_width=True)
            
            c1, c2, c3 = st.columns([4, 4, 2])
            with c1:
                target_user = st.selectbox("승인/상태 변경 대상 선택", df_users["아이디"].tolist())
            with c2:
                new_status = st.selectbox("변경할 상태 선택", ["approved (승인 - 접속허용)", "pending (대기)", "rejected (차단)"])
            with c3:
                st.write("")
                st.write("")
                if st.button("상태 변경 적용", type="primary"):
                    status_code = new_status.split()[0]
                    c = conn.cursor()
                    c.execute("UPDATE users SET status=? WHERE username=?", (status_code, target_user))
                    conn.commit()
                    st.success(f"[{target_user}] 회원 상태가 [{status_code}]로 즉시 변경되었습니다!")
                    st.rerun()
            conn.close()
            
            st.divider()
            st.caption("💡 **신규 가입자용 사전 PIN 부여 안내**: 기본 마스터 PIN 코드는 `000000`입니다. 신규 가입자에게 `000000`을 알려주시고 가입 신청이 들어오면 위에서 [승인] 버튼을 눌러주시면 됩니다.")
            
        with admin_tab2:
            st.subheader("도면 분석 및 견적 이용 기록")
            conn = sqlite3.connect(DB_FILE)
            df_logs = pd.read_sql_query("SELECT id AS 번호, username AS 사용자, file_name AS 도면명, item_count AS 추출수량, ip_address AS 접속IP, analyzed_at AS 분석일시 FROM usage_logs ORDER BY id DESC", conn)
            st.dataframe(df_logs, use_container_width=True)
            conn.close()

# ------------------------------------------------------------------------------
# 5. Gemini AI 도면 분석 함수
# ------------------------------------------------------------------------------
def analyze_drawing_with_gemini(image_pil, api_key, file_name, current_user, ip_addr):
    try:
        import google.generativeai as genai
        genai.configure(api_key=api_key)
        
        prompt = """
        이 분전반 도면 이미지를 정확히 분석하여 도면에 실제로 존재하는 부품 요소들만 추출해 JSON 배열로 반환하세요.

        [추출 규칙]
        1. 분전반명: 도면에 기재된 분전반 이름 (예: L-1, TQ-1, 분전반 등)
        2. 차단기: 메인(MCCB, ELB) 및 분기 차단기 (종류, 극수, AF/AT 용량, 부하명, 수량)
        3. 콘센트: 2구 콘센트 심볼이 있는 경우 종류="콘센트", 수량 카운트
        4. 단자대: N.T, E.T 등 접지/중성 단자대가 표시되어 있다면 종류="단자대", 용량="N.T/E.T"
        5. 계량기: WHM(전력량계) 등

        [JSON 반환 포맷 예시]
        [
            {"분전반명": "L-1", "구분": "MAIN", "종류": "MCCB", "극수": "3P", "용량": "50AF/40AT", "부하명": "메인", "수량": 1, "단가": 85000},
            {"분전반명": "L-1", "구분": "분기", "종류": "ELB", "극수": "2P", "용량": "30AF/20AT", "부하명": "L1", "수량": 1, "단가": 12500}
        ]
        
        주의: 마크다운 태그(```json 등) 없이 오직 pure JSON 배열만 반환하세요.
        """
        
        available_models = []
        for m in genai.list_models():
            if 'generateContent' in m.supported_generation_methods:
                m_name = m.name.replace("models/", "")
                available_models.append(m_name)

        available_models.sort(key=lambda x: 0 if 'flash' in x else 1)
        
        if not available_models:
            raise Exception("사용 가능한 Gemini AI 모델이 존재하지 않습니다.")

        response = None
        last_error = None

        for model_name in available_models:
            try:
                model = genai.GenerativeModel(model_name)
                response = model.generate_content([prompt, image_pil])
                if response and response.text:
                    break
            except Exception as ex:
                last_error = ex
                continue

        if not response or not response.text:
            raise last_error if last_error else Exception("도면 분석에 실패하였습니다.")

        text = response.text.strip()
        
        if text.startswith("```json"):
            text = text[7:]
        if text.startswith("```"):
            text = text[3:]
        if text.endswith("```"):
            text = text[:-3]
        text = text.strip()
        
        parsed_data = json.loads(text)
        
        try:
            conn = sqlite3.connect(DB_FILE)
            c = conn.cursor()
            c.execute('''
                INSERT INTO usage_logs (username, file_name, item_count, ip_address, analyzed_at)
                VALUES (?, ?, ?, ?, ?)
            ''', (current_user, file_name, len(parsed_data), ip_addr, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
            conn.commit()
            conn.close()
        except Exception:
            pass
            
        return parsed_data
    except Exception as e:
        st.error(f"[{file_name}] 분석 중 오류 발생: {e}")
        return []

def generate_excel_quote(df_items, margin_rate, labor_main, labor_branch, shipping):
    wb = openpyxl.Workbook()
    
    ws_sum = wb.active
    ws_sum.title = "전체_견적요약"
    ws_sum.views.sheetView[0].showGridLines = True
    
    navy_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    ws_sum["B2"] = "대한일렉트릭 통합 견적서"
    ws_sum["B2"].font = Font(size=18, bold=True, color="1F4E78")
    
    headers = ["번호", "분전반명", "자재비(원)", "인건비(원)", "합계금액(원)", "비고"]
    for idx, h in enumerate(headers, start=2):
        c = ws_sum.cell(row=5, column=idx, value=h)
        c.fill = navy_fill
        c.font = Font(color="FFFFFF", bold=True)
        c.alignment = Alignment(horizontal="center")

    ws_mat = wb.create_sheet(title="전체_차단기내역")
    ws_mat.views.sheetView[0].showGridLines = True
    
    m_headers = ["분전반명", "구분", "종류", "극수", "용량", "부하명", "수량", "단가(원)", "금액(원)"]
    for idx, h in enumerate(m_headers, start=1):
        c = ws_mat.cell(row=1, column=idx, value=h)
        c.fill = navy_fill
        c.font = Font(color="FFFFFF", bold=True)
        c.alignment = Alignment(horizontal="center")
        
    for r_idx, row in df_items.iterrows():
        r = r_idx + 2
        ws_mat.cell(row=r, column=1, value=str(row.get("분전반명", "")))
        ws_mat.cell(row=r, column=2, value=str(row.get("구분", ""))).alignment = Alignment(horizontal="center")
        ws_mat.cell(row=r, column=3, value=str(row.get("종류", ""))).alignment = Alignment(horizontal="center")
        ws_mat.cell(row=r, column=4, value=str(row.get("극수", ""))).alignment = Alignment(horizontal="center")
        ws_mat.cell(row=r, column=5, value=str(row.get("용량", ""))).alignment = Alignment(horizontal="center")
        ws_mat.cell(row=r, column=6, value=str(row.get("부하명", "")))
        
        try:
            qty = int(row.get("수량", 0))
        except Exception:
            qty = 0
        try:
            price = int(row.get("단가", 0))
        except Exception:
            price = 0

        q_cell = ws_mat.cell(row=r, column=7, value=qty)
        q_cell.fill = yellow_fill
        q_cell.alignment = Alignment(horizontal="center")
        
        p_cell = ws_mat.cell(row=r, column=8, value=price)
        p_cell.fill = yellow_fill
        p_cell.number_format = "#,##0"
        
        amt_cell = ws_mat.cell(row=r, column=9, value=f"=G{r}*H{r}")
        amt_cell.number_format = "#,##0"

    valid_df = df_items[df_items["분전반명"].astype(str).str.strip() != ""]
    panels = valid_df["분전반명"].unique() if len(valid_df) > 0 else ["기본분전반"]
    
    for idx, p_name in enumerate(panels, start=1):
        r = idx + 5
        ws_sum.cell(row=r, column=2, value=idx).alignment = Alignment(horizontal="center")
        ws_sum.cell(row=r, column=3, value=p_name)
        ws_sum.cell(row=r, column=4, value=f"=SUMIF('전체_차단기내역'!A2:A100, \"{p_name}\", '전체_차단기내역'!I2:I100)").number_format = "#,##0"
        ws_sum.cell(row=r, column=5, value=f"=COUNTIF('전체_차단기내역'!A2:A100, \"{p_name}\")*{labor_branch}").number_format = "#,##0"
        ws_sum.cell(row=r, column=6, value=f"=D{r}+E{r}").number_format = "#,##0"

    for sheet in wb.worksheets:
        for col in sheet.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            sheet.column_dimensions[col_letter].width = max(max_len + 5, 12)

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# ------------------------------------------------------------------------------
# 6. 다중 도면 업로드 및 작업 메인 UI (도면별 한 칸 띄우기 적용)
# ------------------------------------------------------------------------------
uploaded_files = st.file_uploader(
    "🖼️ 결선도 도면 여러 장 업로드 (PNG, JPG, 복수 선택 가능)", 
    type=["png", "jpg", "jpeg"],
    accept_multiple_files=True
)

if uploaded_files:
    st.info(f"📂 총 **{len(uploaded_files)}개**의 도면 파일이 선택되었습니다.")
    
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.subheader("🖼️ 선택된 도면 미리보기")
        for idx, file in enumerate(uploaded_files):
            img = Image.open(file)
            st.caption(f"📄 도면 {idx+1}: {file.name}")
            st.image(img, use_container_width=True)
            st.divider()
        
    with col2:
        st.subheader("🔍 전체 도면 통합 해석")
        if st.button("🚀 전체 도면 한번에 분석 시작", type="primary"):
            all_results = []
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            blank_row = {"분전반명": "", "구분": "", "종류": "", "극수": "", "용량": "", "부하명": "", "수량": None, "단가": None}
            
            for idx, file in enumerate(uploaded_files):
                status_text.text(f" 분석 중 ({idx+1}/{len(uploaded_files)}): {file.name}")
                img = Image.open(file)
                parsed_list = analyze_drawing_with_gemini(img, GEMINI_API_KEY, file.name, user['username'], user_ip)
                
                if parsed_list:
                    if all_results:
                        all_results.append(blank_row)
                    all_results.extend(parsed_list)
                    
                progress_bar.progress((idx + 1) / len(uploaded_files))
                
            status_text.text("모든 도면 분석 완료!")
            
            if all_results:
                st.session_state['extracted_data'] = pd.DataFrame(all_results)
                st.success(f"🎉 총 {len(uploaded_files)}개 도면의 분석이 성공적으로 완료되었습니다!")
            else:
                st.session_state['extracted_data'] = pd.DataFrame(columns=DEFAULT_COLUMNS)
                st.warning("도면에서 추출된 데이터가 없습니다.")

if 'extracted_data' not in st.session_state:
    st.session_state['extracted_data'] = pd.DataFrame(columns=DEFAULT_COLUMNS)

st.divider()

col_t, col_b = st.columns([8, 2])
with col_t:
    st.subheader("📋 추출된 분전반별 차단기 데이터 (통합)")
with col_b:
    if st.button("🗑️ 표 전체 비우기"):
        st.session_state['extracted_data'] = pd.DataFrame(columns=DEFAULT_COLUMNS)
        st.rerun()

edited_df = st.data_editor(
    st.session_state['extracted_data'],
    num_rows="dynamic",
    use_container_width=True
)

st.divider()

excel_data = generate_excel_quote(edited_df, MARGIN_RATE, LABOR_RATE_MAIN, LABOR_RATE_BRANCH, EXTRA_SHIPPING)

st.download_button(
    label="📥 대한일렉트릭 견적서 엑셀 다운로드(.xlsx)",
    data=excel_data,
    file_name="대한일렉트릭_분전반_견적서.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary"
)
