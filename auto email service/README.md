## 2021.07.20 : 메일전송 로그파일 생성

메일 전송 후, 누구에게 어떤 내용의 메일을 보냈는지(+발신 일시) 기록되는 txt file 생성 코드 추가

- 엑셀메일링

    아래의 if블럭을 아래와 같이 수정한다.

    ```python
    if msg['To']:
    		smtp.sendmail(sender_id, msg['To'], msg.as_string())  # 메일 전송, 문자열로 변환하여 전송
        with open("logfilexl.txt",'a') as logfilexl :
    		    logfilexl.write("수신인 : "+value['email']+'\n')
            logfilexl.write("메일 전송시간 : "+"%04d/%02d/%02d %02d:%02d:%02d" % (now.tm_year, now.tm_mon, now.tm_mday, now.tm_hour, now.tm_min, now.tm_sec)+"\n")
            logfilexl.write("내용 : \n"+content+'\n\n')
    ```

    write 메서드의 인자에는 반드시 문자열의 형태만 들어간다는 사실!

    그래서 내용을 추가하는 세 번째 logfilexl.write에서 msg['To']가 아니라, content를 넣은 것.

    *살펴볼 코드*

    ```python
     with open("logfilexl.txt",'a') as logfilexl :
    ```

    open 사용 시, close로 꼭 닫아줘야 하는데 with as 구문 사용 시 닫아줄 필요가 없다.

- 일반 메일링

    아래의 if블럭을 아래와 같이 수정한다.

    ```python
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
    	
    	    logfile=open("logfile.txt",'w')
    	    logfile.write("수신인 : "+msg['To']+"\n")
    	    logfile.write("메일 전송시간 : " + "%04d/%02d/%02d %02d:%02d:%02d" % (now.tm_year, now.tm_mon, now.tm_mday, now.tm_hour, now.tm_min, now.tm_sec) + "\n")
    	    logfile.write("내용 : \n" + logcontent + '\n\n')
    	    logfile.close()
    ```

    *살펴볼 코드*

    ```python
     logcontent=",".join(content)
    ```

    위에서 언급했듯, write메서드는 인자에 문자열만 들어갈 수 있기에, join을 통해 리스트(content)를 문자열로 변형시켜주었다. ","로 리스트의 원소들을 연결(?)하여 문자열화!