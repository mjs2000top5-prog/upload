import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# --- 1. 구글 스프레드시트 설정 ---
# 본인의 스프레드시트 ID (URL에서 추출)
SPREADSHEET_ID = "1aEZgrhStJ09lFkdOJjUEuDXJb-cmxQvsfheKpXPgwP0"

def get_worksheet(sheet_name):
    """
    Streamlit Secrets에서 인증 정보를 안전하게 읽어옵니다.
    Incorrect padding 오류를 방지하기 위해 dict 방식을 사용합니다.
    """
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    
    try:
        # Streamlit Cloud의 Settings -> Secrets에 저장된 [gcp_service_account]를 가져옴
        creds_info = st.secrets["gcp_service_account"]
        
        # 딕셔너리 데이터를 사용하여 인증 객체 생성
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_info, scope)
        client = gspread.authorize(creds)
        return client.open_by_key(SPREADSHEET_ID).worksheet(sheet_name)
    except Exception as e:
        st.error(f"❌ 인증 설정 오류: {e}")
        st.info("💡 Streamlit Secrets 설정을 다시 확인해주세요.")
        return None

def upload_logic(uploaded_file, sheet_name):
    """엑셀 파일 처리 및 구글 시트 업로드 공통 로직"""
    if uploaded_file:
        try:
            # 엑셀 읽기
            df = pd.read_excel(uploaded_file, engine='openpyxl')
            
            # 열 밀림 방지: 첫 번째 열이 제목 없는 빈 열(Unnamed)인 경우 삭제
            if df.columns[0].startswith('Unnamed'):
                df = df.iloc[:, 1:]
                
            df = df.fillna('') # 빈 셀 처리
            data = df.values.tolist()
            
            if not data:
                st.warning(f"⚠️ {sheet_name}: 업로드할 데이터가 없습니다.")
                return

            # 시트 연결 및 데이터 추가
            sheet = get_worksheet(sheet_name)
            if sheet:
                # 데이터 행 추가
                sheet.append_rows(data, value_input_option='USER_ENTERED')
                st.success(f"✅ {sheet_name} 시트에 {len(data)}건 업로드 완료!")
            
        except Exception as e:
            st.error(f"❌ {sheet_name} 업로드 중 오류 발생: {e}")
    else:
        st.info(f"📂 {sheet_name}에 올릴 엑셀 파일을 선택해주세요.")

# --- 2. Streamlit UI 레이아웃 ---
st.set_page_config(page_title="위멤버스 실적 통합 관리", layout="centered")

st.title("📊 주간 실적 통합 관리 시스템")
st.write("각 탭을 선택하여 해당 시트에 엑셀 데이터를 업로드하세요.")
st.markdown("---")

# 3개의 탭 생성
tab1, tab2, tab3 = st.tabs(["🏢 위멤버스", "💰 경리나라T", "📁 경리나라"])

with tab1:
    st.subheader("위멤버스 실적 업로드")
    file1 = st.file_uploader("엑셀 파일 선택 (위멤버스)", type=['xlsx', 'xls'], key="u1")
    if st.button("위멤버스 데이터 전송", key="b1"):
        upload_logic(file1, "위멤버스")

with tab2:
    st.subheader("경리나라T 실적 업로드")
    file2 = st.file_uploader("엑셀 파일 선택 (경리나라T)", type=['xlsx', 'xls'], key="u2")
    if st.button("경리나라T 데이터 전송", key="b2"):
        upload_logic(file2, "경리나라T")

with tab3:
    st.subheader("경리나라 실적 업로드")
    file3 = st.file_uploader("엑셀 파일 선택 (경리나라)", type=['xlsx', 'xls'], key="u3")
    if st.button("경리나라 데이터 전송", key="b3"):
        upload_logic(file3, "경리나라")

st.markdown("---")
st.caption(f"연동된 스프레드시트 ID: {SPREADSHEET_ID}")