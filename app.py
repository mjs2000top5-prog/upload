import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# --- 구글 시트 설정 ---
SPREADSHEET_ID = "1aEZgrhStJ09lFkdOJjUEuDXJb-cmxQvsfheKpXPgwP0"
JSON_FILE = 'credentials.json'

def get_worksheet(sheet_name):
    """구글 API 인증 및 시트 가져오기"""
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(JSON_FILE, scope)
    client = gspread.authorize(creds)
    return client.open_by_key(SPREADSHEET_ID).worksheet(sheet_name)

def upload_logic(uploaded_file, sheet_name):
    """파일 처리 및 업로드 로직"""
    if uploaded_file:
        try:
            # 엑셀 읽기
            df = pd.read_excel(uploaded_file, index_col=None)
            
            # 열 밀림 방지 (첫 번째 열이 Unnamed인 경우 제거)
            if df.columns[0].startswith('Unnamed'):
                df = df.iloc[:, 1:]
                
            df = df.fillna('') # 빈 칸 처리
            data = df.values.tolist()
            
            if not data:
                st.warning(f"⚠️ {sheet_name}: 업로드할 데이터가 없습니다.")
                return

            # 시트에 추가
            sheet = get_worksheet(sheet_name)
            sheet.append_rows(data, value_input_option='USER_ENTERED')
            st.success(f"✅ {sheet_name} 시트에 {len(data)}건 업로드 성공!")
            
        except Exception as e:
            st.error(f"❌ {sheet_name} 업로드 중 오류 발생: {e}")
    else:
        st.info(f"📂 {sheet_name}에 올릴 엑셀 파일을 먼저 선택해주세요.")

# --- 스트림릿 UI 레이아웃 ---
st.set_page_config(page_title="위멤버스 실적 통합 업로드", layout="centered")

st.title("📊 주간 실적 통합 관리 시스템")
st.markdown("---")

# 3개의 탭 생성
tab1, tab2, tab3 = st.tabs(["🏢 위멤버스", "💰 경리나라T", "📁 경리나라"])

with tab1:
    st.subheader("위멤버스 실적 업로드")
    file1 = st.file_uploader("위멤버스 엑셀 파일 선택", type=['xlsx', 'xls'], key="u1")
    if st.button("위멤버스 데이터 전송", key="b1"):
        upload_logic(file1, "위멤버스")

with tab2:
    st.subheader("경리나라T 실적 업로드")
    file2 = st.file_uploader("경리나라T 엑셀 파일 선택", type=['xlsx', 'xls'], key="u2")
    if st.button("경리나라T 데이터 전송", key="b2"):
        upload_logic(file2, "경리나라T")

with tab3:
    st.subheader("경리나라 실적 업로드")
    file3 = st.file_uploader("경리나라 엑셀 파일 선택", type=['xlsx', 'xls'], key="u3")
    if st.button("경리나라 데이터 전송", key="b3"):
        upload_logic(file3, "경리나라")

st.markdown("---")
st.caption(f"연결된 시트 ID: {SPREADSHEET_ID}")