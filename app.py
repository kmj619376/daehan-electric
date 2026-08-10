import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io
import json
import sqlite3
import hashlib
from datetime import datetime
from PIL import Image

# ------------------------------------------------------------------------------
# 🔑 API Key 고정 연동
# ------------------------------------------------------------------------------
GEMINI_API_KEY = "AIzaSyAgIpFSVnGSTIMCMBaWtrUhkjNIF-CdqQU"

MARGIN_RATE = 0.10          # 자재 마진율 10%
LABOR_RATE_MAIN = 260000    # 메인 차단기 노무비
LABOR_RATE_BRANCH = 15000   # 분기 차단기 노무비
EXTRA_SHIPPING = 100000     # 현장 운반비 및 양중비

# ------------------------------------------------------------------------------
# 🌐 접속자 IP 주소 추출 함수
# ------------------------------------------------------------------------------
def get_remote_ip():
    try:
        # Streamlit 클라이언트 헤더에서 IP 가져오기
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
    
    # 사용자 테이블 (ip_address 컬럼 포함)
    c.execute('''
        CREATE TABLE IF NOT EXISTS users (
            username TEXT PRIMARY KEY,
            password TEXT NOT NULL,
            name TEXT NOT NULL,
            role TEXT NOT NULL,
            status TEXT NOT NULL,
            ip_address TEXT,
            created_at TEXT NOT NULL
        )
    ''')
    
    # 기존 DB에 ip_address 컬럼이 없는 경우 자동 추가
    try:
        c.execute("ALTER TABLE users ADD COLUMN ip_address TEXT")
    except Exception:
        pass

    # 분석 이력 로그 테이블 (ip_address 컬럼 포함)
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
    
    # 지정하신 최고 관리자 계정 생성 (syd1007 / kmj851007)
    admin_id = "syd1007"
    admin_pass = hashlib.sha256("kmj851007".encode()).hexdigest()
    
    c.execute("SELECT * FROM users WHERE username=?", (admin_id,))
    if not c.fetchone():
        c.execute('''
            INSERT INTO users (username, password, name, role, status, ip_address, created_at)
            VALUES (?, ?, '최고관리자', 'admin', 'approved', '관리자PC', ?)
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

# ------------------------------------------------------------------------------
# 2. 로그인 / 회원가입 UI 및 인증 로직
# ------------------------------------------------------------------------------
if 'logged_in' not in st.session_state:
    st.session_state['logged_in'] = False
if 'user_info' not in st.session_state:
    st.session_state['user_info'] = None

user_ip = get_remote_ip()

if not st.session_state['logged_in']:
    st.title("⚡ 대한일렉트릭 견적 프로그램")
    st.subheader("🔒 사용자 인증 및 승인 관리")
    st.caption(f"🖥️ 현재 접속 PC IP 주소: **{user_ip}**")
    
    tab1, tab2 = st.tabs(["🔑 로그인", "📝 회원가입 신청"])
    
    with tab1:
        st.markdown("### 로그인")
        login_id = st.text_input("아이디 (ID)", key="login_id")
        login_pw = st.text_input("비밀번호", type="password", key="login_pw")
        
        if st.button("로그인하기", type="primary"):
            conn = sqlite3.connect(DB_FILE)
            c = conn.cursor()
            c.execute("SELECT username, name, role, status FROM users WHERE username=? AND password=?", 
                      (login_id, hash_pw(login_pw)))
            user = c.fetchone()
            
            if user:
                username, name, role, status = user
                if status == "approved":
                    # 로그인 성공 시 IP 업데이트
                    c.execute("UPDATE users SET ip_address=? WHERE username=?", (user_ip, username))
                    conn.commit()
                    conn.close()
                    
                    st.session_state['logged_in'] = True
                    st.session_state['user_info'] = {"username": username, "name": name, "role": role, "ip": user_ip}
                    st.success(f"{name}님, 환영합니다!")
                    st.rerun()
                elif status == "pending":
                    conn.close()
                    st.warning("⏳ 아직 관리자 승인 대기 중인 계정입니다. 관리자 승인 후 이용 가능합니다.")
                else:
                    conn.close()
                    st.error("🚫 사용이 차단되거나 비활성화된 계정입니다.")
            else:
                conn.close()
                st.error("❌ 아이디 또는 비밀번호가 일치하지 않습니다.")
                
    with tab2:
        st.markdown("### 회원가입 신청")
        st.info("회원가입 후 관리자(대한일렉트릭)의 승인을 받아야 프로그램 사용이 가능합니다.")
        reg_id = st.text_input("사용할 아이디 (ID)", key="reg_id")
        reg_name = st.text_input("이름 / 회사명", key="reg_name")
        reg_pw = st.text_input("비밀번호", type="password", key="reg_pw")
        reg_pw_confirm = st.text_input("비밀번호 확인", type="password", key="reg_pw_confirm")
        
        if st.button("가입 신청하기"):
            if not reg_id or not reg_name or not reg_pw:
                st.error("모든 항목을 입력해 주세요.")
            elif reg_pw != reg_pw_confirm:
                st.error("비밀번호가 일치하지 않습니다.")
            else:
                conn = sqlite3.connect(DB_FILE)
                c = conn.cursor()
                c.execute("SELECT * FROM users WHERE username=?", (reg_id,))
                if c.fetchone():
                    st.error("이미 존재하는 아이디입니다.")
                else:
                    c.execute('''
                        INSERT INTO users (username, password, name, role, status, ip_address, created_at)
                        VALUES (?, ?, ?, 'user', 'pending', ?, ?)
                    ''', (reg_id, hash_pw(reg_pw), reg_name, user_ip, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
                    conn.commit()
                    st.success("🎉 가입 신청이 완료되었습니다! 관리자 승인 후 이용 가능합니다.")
                conn.close()
    st.stop()

# ------------------------------------------------------------------------------
# 3. 메인 프로그램 화면 (로그인 성공 시)
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
        st.rerun()

# ------------------------------------------------------------------------------
# 👑 관리자(Admin) 전용 통제 메뉴
# ------------------------------------------------------------------------------
if user['role'] == 'admin':
    with st.expander("👑 [관리자 전용] 회원 승인 및 견적 이용 이력 관리 패널", expanded=True):
        admin_tab1, admin_tab2 = st.tabs(["👥 회원 승인 관리", "📜 이용 이력(로그) 보기"])
        
        with admin_tab1:
            st.subheader("사용자 승인 및 상태 변경")
            conn = sqlite3.connect(DB_FILE)
            df_users = pd.read_sql_query("SELECT username AS 아이디, name AS 이름, role AS 권한, status AS 상태, ip_address AS 접속IP, created_at AS 가입일시 FROM users", conn)
            st.dataframe(df_users, use_container_width=True)
            
            c1, c2, c3 = st.columns([3, 3, 2])
            with c1:
                target_user = st.selectbox("상태 변경 대상 아이디 선택", df_users["아이디"].tolist())
            with c2:
                new_status = st.selectbox("새 상태 선택", ["approved (승인)", "pending (대기)", "rejected (차단)"])
            with c3:
                st.write("")
                st.write("")
                if st.button("상태 업데이트 적용"):
                    status_code = new_status.split()[0]
                    c = conn.cursor()
                    c.execute("UPDATE users SET status=? WHERE username=?", (status_code, target_user))
                    conn.commit()
                    st.success(f"{target_user} 계정이 [{status_code}] 상태로 변경되었습니다!")
                    st.rerun()
            conn.close()
            
        with admin_tab2:
            st.subheader("도면 분석 및 견적 이용 기록")
            conn = sqlite3.connect(DB_FILE)
            df_logs = pd.read_sql_query("SELECT id AS 번호, username AS 사용자, file_name AS 도면명, item_count AS 추출수량, ip_address AS 접속IP, analyzed_at AS 분석일시 FROM usage_logs ORDER BY id DESC", conn)
            st.dataframe(df_logs, use_container_width=True)
            conn.close()

# ------------------------------------------------------------------------------
# 4. Gemini AI 분석 및 엑셀 생성 함수
# ------------------------------------------------------------------------------
def get_multi_panel_sample():
    return [
        {"분전반명": "TQ-1 ~ TQ-3 [3면]", "구분": "MAIN", "종류": "MCCB", "극수": "4P", "용량": "100AF/100AT", "부하명": "메인 전원", "수량": 3, "단가": 85000},
        {"분전반명": "TQ-1 ~ TQ-3 [3면]", "구분": "분기", "종류": "ELB", "극수": "2P", "용량": "30AF/20AT", "부하명": "전등/전열", "수량": 18, "단가": 12500},
        {"분전반명": "계단분전반 [13면]", "구분": "MAIN", "종류": "MCCB", "극수": "4P", "용량": "100AF/75AT", "부하명": "계단 메인", "수량": 13, "단가": 75000},
        {"분전반명": "계단분전반 [13면]", "구분": "분기", "종류": "ELB", "극수": "2P", "용량": "30AF/20AT", "부하명": "계단 전등", "수량": 26, "단가": 12500},
        {"분전반명": "동력반(P1~P3) [3면]", "구분": "MAIN", "종류": "MCCB", "극수": "4P", "용량": "225AF/150AT", "부하명": "동력 메인", "수량": 3, "단가": 150000},
        {"분전반명": "동력반(P1~P3) [3면]", "구분": "분기", "종류": "MCCB", "극수": "3P", "용량": "50AF/30AT", "부하명": "펌프/팬 동력", "수량": 12, "단가": 32000},
        {"분전반명": "동력반(P1~P3) [3면]", "구분": "부속", "종류": "콘센트", "극수": "2P", "용량": "16A", "부하명": "분전반 내 2구 콘센트", "수량": 4, "단가": 8000},
        {"분전반명": "동력반(P1~P3) [3면]", "구분": "부속", "종류": "단자대", "극수": "-", "용량": "N.T / E.T", "부하명": "중성선/접지 단자대", "수량": 2, "단가": 15000},
    ]

def analyze_drawing_with_gemini(image_pil, api_key, file_name, current_user, ip_addr):
    try:
        import google.generativeai as genai
        genai.configure(api_key=api_key)
        
        prompt = """
        이 분전반 도면 이미지에는 차단기 표 외에도 여러 가지 전기 부품 심볼이 포함되어 있습니다.
        다음 요소들을 철저히 분석하여 JSON 배열 형태로만 반환하세요:

        1. 차단기: 메인(MCCB, ELB) 및 분기 차단기 (종류, 극수, AF/AT 용량, 부하명, 수량)
        2. 콘센트 심볼: 도면 하단이나 옆에 배치된 2구 콘센트 기호(동그라미 2개 모양)가 있다면 종류="콘센트", 부하명="분전반 내 콘센트", 수량 정밀 카운트
        3. 단자대 심볼: N.T(Neutral Terminal), E.T(Earth Terminal) 등 접지/중성 단자대가 표시되어 있다면 종류="단자대", 용량="N.T/E.T" 항목 추가
        4. 계량기/부속: WHM(전력량계), 지상/벽부형 외함 등

        반환 JSON 형식 예시:
        [
            {"분전반명": "분전반 이름(예: TQ-1)", "구분": "MAIN 또는 분기 또는 부속", "종류": "MCCB / ELB / 콘센트 / 단자대 / WHM 등", "극수": "2P/3P/4P/-", "용량": "100AF/75AT 또는 16A 등", "부하명": "부하 이름", "수량": 1, "단가": 85000}
        ]
        
        주의사항:
        1. ```json 과 같은 마크다운 태그나 설명글을 절대 붙이지 말고 오직 [ ... ] 형식의 pure JSON 배열 문자열만 출력하세요.
        2. 수량과 단가는 숫자 형태로 출력하세요.
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
        
        # 📜 이용 이력(로그) DB 기록 (IP 포함)
        try:
            conn = sqlite3.connect(DB_FILE)
            c = conn.cursor()
            c.execute('''
                INSERT INTO usage_logs (username, file_name, item_count, ip_address, analyzed_at)
                VALUES (?, ?, ?, ?, ?)
            ''', (current_user, file_name, len(parsed_data), ip_addr, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
            conn.commit()
            conn.close()
        except Exception as log_err:
            pass
            
        return parsed_data
    except Exception as e:
        st.error(f"도면 분석 중 오류 발생: {e}")
        return get_multi_panel_sample()

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

    panels = df_items["분전반명"].unique() if "분전반명" in df_items.columns else ["기본분전반"]
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
# 5. 도면 업로드 및 작업 메인 UI
# ------------------------------------------------------------------------------
uploaded_file = st.file_uploader("🖼️ 결선도 도면 업로드 (PNG, JPG)", type=["png", "jpg", "jpeg"])

if uploaded_file:
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.subheader("🖼️ 업로드된 도면")
        image = Image.open(uploaded_file)
        st.image(image, use_container_width=True)
        
    with col2:
        st.subheader("🔍 도면 해석")
        if st.button("🚀 도면 분석 시작"):
            with st.spinner("도면의 차단기, 콘센트, 단자대 등 모든 부품 데이터를 분석 중입니다..."):
                raw_data = analyze_drawing_with_gemini(image, GEMINI_API_KEY, uploaded_file.name, user['username'], user_ip)
                st.session_state['extracted_data'] = pd.DataFrame(raw_data)
                st.success("도면 분석이 완료되었습니다!")

if 'extracted_data' not in st.session_state:
    st.session_state['extracted_data'] = pd.DataFrame(get_multi_panel_sample())

st.divider()

col_t, col_b1, col_b2 = st.columns([6, 2, 2])
with col_t:
    st.subheader("📋 추출된 분전반별 차단기 데이터")
with col_b1:
    if st.button("🔄 샘플 데이터로 리셋"):
        st.session_state['extracted_data'] = pd.DataFrame(get_multi_panel_sample())
        st.rerun()
with col_b2:
    if st.button("🗑️ 전체 삭제"):
        st.session_state['extracted_data'] = pd.DataFrame(columns=["분전반명", "구분", "종류", "극수", "용량", "부하명", "수량", "단가"])
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
