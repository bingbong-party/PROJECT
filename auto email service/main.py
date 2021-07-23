import smtplib
from email.mime.text import MIMEText
import openpyxl as xl
from datetime import datetime
import time

now = time.localtime()

# print "%04d/%02d/%02d %02d:%02d:%02d" % (now.tm_year, now.tm_mon, now.tm_mday, now.tm_hour, now.tm_min, now.tm_sec)

def get_excel_data():
    wb = xl.load_workbook(input('파일 경로를 입력해 주세요.\n'), data_only=True)
    ws1 = wb['학생정보']
    ws2 = wb['성적']

    # 학생 정보 시트 데이터
    student_list = []
    for student in ws1.rows:
        if student[0].value is None:
            continue
        student_list.append([])
        for c in student:
            student_list[-1].append(c.value)
    student_list.pop(0)

    # 성적 시트 데이터
    score_list = []
    for score in ws2.rows:
        if score[0].value is None:
            continue
        score_list.append([])
        for c in score:
            score_list[-1].append(c.value)
    score_list.pop(0)

    # 학생 정보 및 성적 시트
    student_dict = {}
    for student_info in student_list:
        student_dict[student_info[0]] = {
            'name': student_info[0],
            'school': student_info[1],
            'grade': student_info[2],
            'email': student_info[3],
            'address': student_info[4],
            'phone': student_info[5],
            'par_phone': student_info[6],
            'fee': student_info[7],
            'fee_date': student_info[8]
        }

    for score in score_list:
        student_dict[score[2]]['midterm'] = score[3]
        student_dict[score[2]]['final_exam'] = score[4]
        student_dict[score[2]]['exam_avg'] = round(score[5], 2)
        student_dict[score[2]]['mock_exam1'] = score[6]
        student_dict[score[2]]['mock_exam2'] = score[7]
        student_dict[score[2]]['mock_exam3'] = score[8]
        student_dict[score[2]]['mock_exam4'] = score[9]
        student_dict[score[2]]['mock_exam5'] = score[10]
        student_dict[score[2]]['mock_exam6'] = score[11]
        student_dict[score[2]]['mock_exam7'] = score[12]
        student_dict[score[2]]['mock_exam8'] = score[13]
        student_dict[score[2]]['mock_exam9'] = score[14]
        student_dict[score[2]]['mock_exam10'] = score[15]
        student_dict[score[2]]['mock_exam11'] = score[16]
        student_dict[score[2]]['mock_exam12'] = score[17]
        student_dict[score[2]]['mock_exam_avg'] = round(score[18], 2)

    return student_dict

# 이메일 전송
def send_email(select, *args):
    print('네이버 또는 구글 메일만 사용 가능합니다.')
    sender_id = input('메일을 보낼 계정을 입력해 주세요: ')
    sender_pw = input('계정 비밀번호를 입력해 주세요: ')
    if 'naver' in sender_id:
        smtp_server = "smtp.naver.com"
        print('naver')
    elif 'google' in sender_id:
        print('google')
        smtp_server = "smtp.google.com"
    else:
        print('네이버 또는 구글 메일만 사용 가능합니다.\n메일 주소를 확인해 주세요')
        raise Exception('네이버 또는 구글 메일만 사용 가능합니다.')

    smtp_info = {
        "smtp_server": smtp_server,  # SMTP 서버 주소
        "smtp_user_id": sender_id,
        "smtp_user_pw": sender_pw,
        "smtp_port": 587  # SMTP 서버 포트
    }

    # 엑셀 데이터 메일 전송
    if select == 'excel':
        smtp = smtplib.SMTP(smtp_info['smtp_server'], smtp_info['smtp_port'])
        smtp.ehlo
        smtp.starttls()  # TLS 보안 처리
        smtp.login(sender_id, sender_pw)  # 로그인

        for value in args[0].values():
            title = f"{value['school']} {value['grade']} {value['name']}"
            content = f"{(datetime.today()).month}월 원비는 {value['fee']}원 입니다. {value['fee_date']}까지 납입바랍니다.\n\
중간고사: {value['midterm']}\n\
기말고사: {value['final_exam']}\n\
평균: {value['exam_avg']}\n\
모의고사 성적\n\
1월: {value['mock_exam1']}"

            msg = MIMEText(content)  # , _charset="utf8")
            msg['Subject'] = title  # 메일 제목
            msg['From'] = smtp_info['smtp_user_id']  # 송신자
            msg['To'] = value['email']

            if msg['To']:
                # smtp.sendmail(sender_id, msg['To'], msg.as_string())  # 메일 전송, 문자열로 변환하여 전송
                with open("logfilexl.txt", 'a') as logfilexl :
                    logfilexl.write("수신인 : "+value['email']+'\n')
                    logfilexl.write("메일 전송시간 : "+"%04d/%02d/%02d %02d:%02d:%02d" % (now.tm_year, now.tm_mon, now.tm_mday, now.tm_hour, now.tm_min, now.tm_sec)+"\n")
                    logfilexl.write("내용 : \n"+content+'\n\n')

    # 일반 메일 전송
    if select == 'normal':
        to = input('받는 분 메일 주소를 입력해 주세요\n여러명일경우 ,로 구분됩니다.\nex)test@test.com, test2@test.com\n')
        title = input('메일 제목을 입력해 주세요\n')

        # 메일 내용이 몇줄이 들어갈지 모르기 때문에 무한반복으로 데이터를 인풋받아 리스트에 넣어줌
        content = []
        print('메일 내용을 작성해 주세요\n 작성 완료시 숫자 0을 입력해주세요.')
        while (True):
            data = input()
            if data == '0':
                break
            else:
                content.append(data)

        logcontent=",".join(content)

        # 리스트로 받은 content를 \n로 조인하여 줄바꿈
        msg = MIMEText('\n'.join(content), _charset="utf8")

        msg['Subject'] = title  # 메일 제목
        msg['From'] = smtp_info['smtp_user_id']  # 송신자
        msg['To'] = to

        smtp = smtplib.SMTP(smtp_info['smtp_server'], smtp_info['smtp_port'])
        smtp.ehlo
        smtp.starttls()  # TLS 보안 처리
        smtp.login(sender_id, sender_pw)  # 로그인

        smtp.sendmail(msg['From'], msg['To'].split(','), msg.as_string())

        logfile=open("logfile.txt", 'w')
        logfile.write("수신인 : "+msg['To']+"\n")
        logfile.write("메일 전송시간 : " + "%04d/%02d/%02d %02d:%02d:%02d" % (now.tm_year, now.tm_mon, now.tm_mday, now.tm_hour, now.tm_min, now.tm_sec) + "\n")
        logfile.write("내용 : \n" + logcontent + '\n\n')
        logfile.close()

    smtp.close()
    print('메일을 성공적으로 보냈습니다.')

# 일반 메일 전송
# send_email('normal')

# 엑셀 데이터 메일 전송
send_email('excel', get_excel_data()) 


