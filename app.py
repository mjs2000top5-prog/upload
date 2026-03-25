import os
import pandas as pd
import gspread
import traceback
from flask import Flask, request
from oauth2client.service_account import ServiceAccountCredentials

app = Flask(__name__)

# --- 설정 구간 ---
SPREADSHEET_ID = "1aEZgrhStJ09lFkdOJjUEuDXJb-cmxQvsfheKpXPgwP0"
JSON_FILE = 'credentials.json'
# ----------------

def get_worksheet(sheet_name):
    """지정한 이름의 시트 가져오기"""
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(JSON_FILE, scope)
    client = gspread.authorize(creds)
    return client.open_by_key(SPREADSHEET_ID).worksheet(sheet_name)

def process_and_upload(file, sheet_name):
    """엑셀 처리 및 업로드 공통 로직"""
    df = pd.read_excel(file, index_col=None)
    # 열 밀림 방지: 첫 번째 열이 Unnamed인 경우 삭제
    if df.columns[0].startswith('Unnamed'):
        df = df.iloc[:, 1:]
    df = df.fillna('')
    data = df.values.tolist()
    
    if data:
        sheet = get_worksheet(sheet_name)
        sheet.append_rows(data, value_input_option='USER_ENTERED')
    return len(data)

@app.route('/')
def index():
    return f'''
    <!doctype html>
    <html>
        <head>
            <title>주간 실적 통합 업로드</title>
            <style>
                body {{ font-family: 'Malgun Gothic', sans-serif; margin: 40px; background-color: #f0f2f5; text-align: center; }}
                h1 {{ color: #202124; margin-bottom: 10px; }}
                .main-p {{ color: #5f6368; margin-bottom: 40px; }}
                .container {{ max-width: 1100px; margin: auto; display: flex; gap: 20px; justify-content: center; flex-wrap: wrap; }}
                .card {{ background: white; padding: 25px; border-radius: 12px; box-shadow: 0 1px 3px rgba(0,0,0,0.12), 0 1px 2px rgba(0,0,0,0.24); width: 300px; transition: all 0.3s cubic-bezier(.25,.8,.25,1); }}
                .card:hover {{ box-shadow: 0 14px 28px rgba(0,0,0,0.25), 0 10px 10px rgba(0,0,0,0.22); }}
                h2 {{ font-size: 1.25em; margin-bottom: 15px; padding-bottom: 10px; border-bottom: 2px solid #eee; }}
                .card.wemembers h2 {{ color: #1a73e8; border-color: #1a73e8; }}
                .card.gyeongrit h2 {{ color: #34a853; border-color: #34a853; }}
                .card.gyeongri h2 {{ color: #f9ab00; border-color: #f9ab00; }}
                input[type=file] {{ width: 100%; margin: 15px 0; font-size: 0.8em; }}
                .btn {{ color: white; border: none; padding: 12px; cursor: pointer; border-radius: 6px; width: 100%; font-size: 15px; font-weight: bold; }}
                .btn-blue {{ background: #1a73e8; }}
                .btn-green {{ background: #34a853; }}
                .btn-yellow {{ background: #f9ab00; }}
                .btn:hover {{ opacity: 0.8; }}
            </style>
        </head>
        <body>
            <h1>📊 주간 실적 통합 관리 시스템</h1>
            <p class="main-p">해당하는 시트의 카드를 선택해 파일을 업로드하세요.</p>
            
            <div class="container">
                <div class="card wemembers">
                    <h2>🏢 위멤버스</h2>
                    <form method="post" enctype="multipart/form-data" action="/upload/위멤버스">
                        <input type="file" name="file" accept=".xlsx, .xls" required>
                        <input type="submit" class="btn btn-blue" value="위멤버스 업로드">
                    </form>
                </div>

                <div class="card gyeongrit">
                    <h2>💰 경리나라T</h2>
                    <form method="post" enctype="multipart/form-data" action="/upload/경리나라T">
                        <input type="file" name="file" accept=".xlsx, .xls" required>
                        <input type="submit" class="btn btn-green" value="경리나라T 업로드">
                    </form>
                </div>

                <div class="card gyeongri">
                    <h2>📁 경리나라</h2>
                    <form method="post" enctype="multipart/form-data" action="/upload/경리나라">
                        <input type="file" name="file" accept=".xlsx, .xls" required>
                        <input type="submit" class="btn btn-yellow" value="경리나라 업로드">
                    </form>
                </div>
            </div>
        </body>
    </html>
    '''

@app.route('/upload/<sheet_name>', methods=['POST'])
def upload_file(sheet_name):
    if 'file' not in request.files: return "파일이 없습니다."
    file = request.files['file']
    
    try:
        count = process_and_upload(file, sheet_name)
        return f'''
            <div style="text-align:center; margin-top:100px; font-family: 'Malgun Gothic';">
                <h1 style="color: #202124;">✅ 업로드 완료</h1>
                <p style="font-size: 1.2em;"><b>[{sheet_name}]</b> 시트에 <b>{count}</b>건의 데이터가 추가되었습니다.</p>
                <br>
                <a href="/" style="display:inline-block; padding:10px 20px; background:#eee; color:#333; text-decoration:none; border-radius:5px;">돌아가기</a>
                <a href="https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}" target="_blank" style="display:inline-block; padding:10px 20px; background:#1a73e8; color:white; text-decoration:none; border-radius:5px; margin-left:10px;">시트 전체 확인</a>
            </div>
        '''
    except Exception as e:
        error_msg = traceback.format_exc()
        return f"<div style='padding:30px;'><h2 style='color:red;'>오류 발생 ({sheet_name})</h2><pre>{error_msg}</pre><a href='/'>다시 시도</a></div>"

if __name__ == '__main__':
    app.run(debug=True, port=5000)