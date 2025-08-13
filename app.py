# save to csv file
from flask import Flask, render_template, request, jsonify, redirect, url_for
import smtplib
from email.mime.text import MIMEText
import csv
import os
from datetime import datetime
import json

app = Flask(__name__)

# 저장할 폴더 경로 
SAVE_FOLDER = "./sw_requests"

# 폴더가 없으면 생성
if not os.path.exists(SAVE_FOLDER):
    os.makedirs(SAVE_FOLDER)

# config.json 파일 읽기 >>> 경로 수정 후 테스트하기!! 
try:
    with open('.\.gitignore\config.json', 'r') as f:
        config = json.load(f)
except FileNotFoundError:
    print("No config.json file exists.. Plz create file")
    config = {}


@app.route('/')
def index():
    return render_template('survey.html')

# 이메일 발송 함수 
def send_email(subject, body):
    # 환경 변수에서 이메일 정보 가져오기
    sender_email = config.get('SENDER_EMAIL')
    sender_password = config.get('SENDER_PASSWORD')
    recipient_email = config.get('RECIPIENT_EMAIL')
    
    if not sender_email or not sender_password or not recipient_email:
        print("설정 파일에 이메일 정보가 부족합니다.")
        return False

    smtp_server = "smtp.gmail.com"
    port = 587 # TLS port

    msg = MIMEText(body, 'html')
    msg['Subject'] = subject
    msg['From'] = sender_email
    msg['To'] = recipient_email

    print(f"HELLO~~~~~~~~~~~~~")

    try:
        # with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        #     smtp.login(sender_email, sender_password)
        #     smtp.send_message(msg)
        server = smtplib.SMTP(smtp_server, port)
        server.starttls()
        server.login(sender_email, sender_password)
        print("이메일 발송 성공!")
        return True
    except Exception as e:
        print(f"이메일 발송 오류: {e}")
        return False
    finally:
        server.quit()


@app.route('/submit', methods=['POST'])
def submit_survey():
    try:
        # 폼 데이터 받기
        data = request.get_json()
        
        # 현재 날짜 및 시간
        request_time = datetime.now().strftime('%H:%M:%S')
        today = datetime.now().strftime('%Y%m%d')
        
        # CSV 파일 경로 - 파일명 변경
        csv_filename = f"request_{today}.csv"
        csv_path = os.path.join(SAVE_FOLDER, csv_filename)
        
        # 헤더 확인 및 파일 생성
        file_exists = os.path.exists(csv_path)
        
        with open(csv_path, 'a', newline='', encoding='utf-8-sig') as csvfile:
            # 헤더 변경 - 한 row에 모든 정보 저장
            fieldnames = ['req_date','req_time', 'department', 'name', 'employee_id', 'software_nm', 'work_type']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            
            # 파일이 새로 생성된 경우 헤더 추가
            if not file_exists:
                writer.writeheader()
            
            # 한 row에 모든 설문 데이터 저장
            survey_row = {
                'req_date': today,
                'req_time': request_time,
                'department': data['department'],
                'name': data['name'],
                'employee_id': data['employee_id'],
                'software_nm': data['software'],
                'work_type': data['work_type']
            }
            
            writer.writerow(survey_row)
        
        print(f"설문이 성공적으로 저장되었습니다! at {csv_path}")
        print(f"SW담당자에게 이메일을 발송합니다...on progress")
        
        # SW 담당자 이메일 발송
        recipient_email = "sw.admin@lghnh.com"

        sw_type = survey_row.get('software_nm', 'N/A')
        subject = f"새로운 SW 신청서 접수: {sw_type}"

        # HTML 테이블 생성
        html_table = "<h2>새로운 소프트웨어 신청서가 접수되었습니다.</h2>"
        html_table += "<table border='1' cellpadding='5' cellspacing='0'>"
        html_table += "<tr><th>항목</th><th>내용</th></tr>"

        # 딕셔너리의 키와 값을 순회하며 HTML 테이블 행으로 만듭니다.
        for key, value in survey_row.items():
            html_table += f"<tr><td><strong>{key}</strong></td><td>{value}</td></tr>"

        html_table += "</table>"
        
        # 전체 이메일 본문
        email_body = f"""
        <html>
            <body>
                {html_table}
                <hr>
                <p>이 이메일은 자동 발송된 것입니다.</p>
            </body>
        </html>
        """

        send_email(subject, email_body)

        return jsonify({'status': 'success', 'message': '설문이 성공적으로 저장되었습니다!'})
    
    except Exception as e:
        return jsonify({'status': 'error', 'message': f'저장 중 오류가 발생했습니다: {str(e)}'})


# @app.route('/success')
# def success():
#     return "설문지가 성공적으로 제출되었습니다."
    
if __name__ == '__main__':
    app.run(debug=True, port=5000)