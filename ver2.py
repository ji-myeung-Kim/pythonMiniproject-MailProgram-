import smtplib
from email.mime.text import MIMEText
from imap_tools import MailBox 
import requests
import openpyxl as xl
from datetime import datetime
import requests
import cx_Oracle

def connect_db(query):
    try:
        cx_Oracle.init_oracle_client(lib_dir=r"D:\Oracle\instantclient_19_11")
    except:
        pass
    
    result = ''
    connection = cx_Oracle.connect(user='ora01', password='oracle_4U2021', dsn='edudb1_high')

    cursor = connection.cursor()
    
    cursor.execute(query)
    if 'select' in query:
        result = cursor.fetchall()
    connection.commit()
    connection.close()
    
    return result

# 엑셀 데이터 가져오기
# DB 데이터 insert
def get_excel_data():
    wb = xl.load_workbook(input('파일 경로를 입력해 주세요.\n'), data_only=True)
    ws1 = wb['학생정보']
    ws2 = wb['성적']
    
    # 학생 정보 시트 데이터
    student_list = []
    for student in ws1.rows:
        if student[0].value is not None:
            student_list.append([])
            for c in student:
                student_list[-1].append(c.value)
    student_list.pop(0)

    # 성적 시트 데이터
    score_list = []
    for score in ws2.rows:
        if score[0].value is not None:
            score_list.append([])
            for c in score:
                if not c.value:
                    c.value = -1
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
    cx_Oracle.init_oracle_client(lib_dir=r"D:\Oracle\instantclient_19_11")

    connection = cx_Oracle.connect(user='ora01', password='oracle_4U2021', dsn='edudb1_high')
    cursor = connection.cursor()
    for student_val in student_dict.values():
        cursor.execute(f"""
        insert into student(id, name, school, grade, phone, par_phone, email, address, fee, fee_date) 
        values (student_id_seq.nextval, '{student_val['name']}', '{student_val['school']}', '{student_val['grade']}', '{student_val['phone']}', '{student_val['par_phone']}', '{student_val['email']}', '{student_val['address']}', '{student_val['fee']}', '{student_val['fee_date']}')
        """)

        cursor.execute(f"""
        insert into exam(id, name, midterm, final_exam, mock_exam1, mock_exam2, mock_exam3, mock_exam4, mock_exam5, mock_exam6, mock_exam7, mock_exam8, mock_exam9, mock_exam10, mock_exam11, mock_exam12)
        values (exam_id_seq.nextval, '{student_val['name']}', '{student_val['midterm']}', '{student_val['final_exam']}', '{student_val['mock_exam1']}', '{student_val['mock_exam2']}', '{student_val['mock_exam3']}', '{student_val['mock_exam4']}', '{student_val['mock_exam5']}', '{student_val['mock_exam6']}', '{student_val['mock_exam7']}', '{student_val['mock_exam8']}', '{student_val['mock_exam9']}', '{student_val['mock_exam10']}', '{student_val['mock_exam11']}', '{student_val['mock_exam12']}')
        """)
        
    connection.commit()
    connection.close()
    start()

# 이메일 전송
# DB 데이터 메일링
def send_email(select):
    print('네이버 또는 구글 메일만 사용 가능합니다.')
    sender_id = input('메일을 보낼 계정을 입력해 주세요: ')
    print('='*50)
    sender_pw = input('계정 비밀번호를 입력해 주세요: ')
    print('='*50)
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
    "smtp_port": 587 # SMTP 서버 포트
    }
    
    # DB 데이터 메일 전송
    if select == 'excel':
        mail_choice = int(input('원비납부 요청 메일 = 1, 학생 성적 메일 = 2\n'))
        smtp = smtplib.SMTP(smtp_info['smtp_server'], smtp_info['smtp_port'])
        smtp.ehlo
        smtp.starttls()  # TLS 보안 처리
        smtp.login(sender_id , sender_pw)  # 로그인
        query = f"""
        select student.id, student.school, student.grade, student.phone, student.par_phone, student.email, student.address, student.fee, student.fee_date,
        TRANSLATE(exam.midterm, -1, 'X'), TRANSLATE(exam.final_exam, -1, 'X'), TRANSLATE(exam.mock_exam1, -1, 'X'), TRANSLATE(exam.mock_exam2, -1, 'X'), TRANSLATE(exam.mock_exam3, -1, 'X'), TRANSLATE(exam.mock_exam4, -1, 'X'), TRANSLATE(exam.mock_exam5, -1, 'X'), TRANSLATE(exam.mock_exam6, -1, 'X'), TRANSLATE(exam.mock_exam7, -1, 'X'), TRANSLATE(exam.mock_exam8, -1, 'X'), TRANSLATE(exam.mock_exam9, -1, 'X'), TRANSLATE(exam.mock_exam10, -1, 'X'), TRANSLATE(exam.mock_exam11, -1, 'X'), TRANSLATE(exam.mock_exam12, -1, 'X')
        from student, exam
        where student.name = exam.name
        """
        
        data_list = connect_db(query)

        for data in data_list:
            if mail_choice == 1:
                title = f'{data[1]}학생 {datetime.now().month}월 원비 납부 요청드립니다.'
                content = f'{data[1]}학생 {datetime.now().month}월 원비는 {format(data[8], ",")}원 이며, {data[9]}일까지 결제 바랍니다.'

                msg = MIMEText(content)
                msg['Subject'] = title # 메일 제목
                msg['From'] = smtp_info['smtp_user_id'] # 송신자
                msg['To'] = data[5]

            elif mail_choice == 2:
                title = f'{data[1]}학생 성적 발송드립니다.'
                content = f'{data[1]}학생\n중간고사{data[10]}점 기말고사 {data[11]}점이며 평균은 {(data[10]+data[11])/2}점 입니다.'
                
                msg = MIMEText(content)
                msg['Subject'] = title # 메일 제목
                msg['From'] = smtp_info['smtp_user_id'] # 송신자
                msg['To'] = data[5]
                
            if msg['To']:
                smtp.sendmail(sender_id , msg['To'], msg.as_string())  # 메일 전송, 문자열로 변환하여 전송
                query = f"insert into maillog(id, sender, receiver, subject) values (maillog_id_seq.nextval, '{msg['From']}', '{msg['To']}', '{msg['Subject']}')"
                connect_db(query)
        
        start()

    # 일반 메일 전송
    if select == 'nomal':
        to = input('받는 분 메일 주소를 입력해 주세요\n여러명일경우 ,로 구분됩니다.\nex)test@test.com, test2@test.com\n')
        print('='*50)
        title = input('메일 제목을 입력해 주세요\n')
        print('='*50)
        
        # 메일 내용이 몇줄이 들어갈지 모르기 때문에 무한반복으로 데이터를 인풋받아 리스트에 넣어줌
        content = []
        print('메일 내용을 작성해 주세요\n 작성 완료시 숫자 0을 입력해주세요.')
        i = 0
        while(True):
            i += 1
            data = input(f'{i}번째 라인: ')
            if data == '0':
                break
            else:
                content.append(data)
        print('='*50)
        # 리스트로 받은 content를 \n로 조인하여 줄바꿈
        msg = MIMEText('\n'.join(content),_charset="utf8")
        
        msg['Subject'] = title  # 메일 제목
        msg['From'] = smtp_info['smtp_user_id']  # 송신자
        msg['To'] = to
        
        smtp = smtplib.SMTP(smtp_info['smtp_server'], smtp_info['smtp_port'])
        smtp.ehlo
        smtp.starttls()  # TLS 보안 처리
        smtp.login(sender_id , sender_pw)  # 로그인
        
        smtp.sendmail(msg['From'], msg['To'].split(','), msg.as_string())

        query = f"insert into maillog(id, sender, receiver, subject) values (maillog_id_seq.nextval, '{msg['From']}', '{msg['To']}', '{msg['Subject']}')"
        connect_db(query)

    smtp.close()
    print('메일을 성공적으로 보냈습니다.')
    start()

# 메일 수신
def receive():
    print('네이버 또는 구글 메일만 사용 가능합니다.')
    sender_id = input('메일계정을 입력해 주세요: ')
    print('='*50)
    sender_pw = input('계정 비밀번호를 입력해 주세요: ')
    print('='*50)

    if 'naver' in sender_id:
        imap_server = "imap.naver.com"
        print('naver')    
    elif 'google' in sender_id:
        print('google')
        imap_server = "imap.google.com"
    else:
        print('네이버 또는 구글 메일만 사용 가능합니다.\n메일 주소를 확인해 주세요')
        raise Exception('네이버 또는 구글 메일만 사용 가능합니다.')
    
    imap_info = {
        "imap_server": imap_server,  # Imap 서버 주소
        "imap_user_id": sender_id,
        "imap_user_pw": sender_pw,
        "imap_port": 993 # Imap 서버 포트
        }

    #로그인하여 메일박스를 열어 검색할 수 있도록 해줌
    mailbox = MailBox(imap_info["imap_server"], imap_info["imap_port"])
    mailbox.login(sender_id , sender_pw, initial_folder="INBOX")

    #limit: 최대 메일 갯수
    #reverse: True일 경우 최근 메일부터, False일 경우 과거 메일부터
    #끝에서 한가지 메일 추출, limit을 원하는 만큼 추출 가능
    count = int(input("몇 개의 메일을 가져오실건가요?\nex)5개의 메일을 가져올 경우 5 입력\n"))
    print("="*50)
    for msg in mailbox.fetch(limit=count, reverse=True):
        print("제목:", msg.subject)
        print("발신자", msg.from_)
        print("수신자:", msg.to)
        print("본문", msg.text)
        print("="*50)
        
        trans_choice = int(input('메일 내용을 번역하시겠습니까?\n예 = 1, 아니요 = 2\n'))
        if trans_choice == 1:

            client_id = "M2TCtkuAE0f1_4eJTDeS" # 본인의 client_id 작성
            client_secret = "QJ9CRWHp8i" # 본인의 client_secret 작성
            
            lang = ''
            trans_lang = ''
            
            print("어떤 언어로 번역할지 선택해주세요: ")
            choice = int(input("한국어 =1, 영어 = 2 \n"))
            if choice == 1:   #1번 선택시 영어를 한글로
                lang = 'en'
                trans_lang = 'ko'
            elif choice == 2:   #2번 선택시 한글을 영어로
                lang = 'ko'
                trans_lang = 'en'
            else:
                print('잘못 입력했습니다.')
                raise Exception('1, 2 중 선택해주세요.')  #에러가 뜰시 종료가 안되도록 예외처리 해줌
        
            data = {'text' : msg.text,
                    'source' : lang,
                '   target': trans_lang}

            url = "https://openapi.naver.com/v1/papago/n2mt"
            header = {"X-Naver-Client-Id":client_id,
                    "X-Naver-Client-Secret":client_secret}
            response = requests.post(url, headers=header, data= data)
            rescode = response.status_code

            if(rescode==200):
                print(f'번역: {response.json()["message"]["result"]["translatedText"]}') #우리 원하는 번역값을 출력
            else:    
                print("Error Code:" , rescode)        
        elif trans_choice == 2:
            continue
        else:
            print('잘못 입력하셨습니다.')
            raise Exception('1, 2 중 선택해주세요.')
    mailbox.logout()

    start()

# 번역기
def translate():
    lang = ''
    trans_lang = ''
    
    print('번역할 언어를 선택해주세요.')
    choice = int(input('한국어 = 1, 영어 = 2, 일본어 = 3\n'))
    print('='*50)
    if choice == 1:
        lang = 'ko'
    elif choice == 2:
        lang = 'en'
    elif choice == 3:
        lang = 'ja'
    else:
        print('잘못 입력했습니다.')
        raise Exception('1, 2, 3 중 선택해주세요.')

    print('어떤 언어로 번역할지 선택해 주세요')
    choice2 = int(input('한국어 = 1, 영어 = 2, 일본어 = 3\n'))
    print('='*50)
    if choice2 == 1:
        trans_lang = 'ko'
    elif choice2 == 2: 
        trans_lang = 'en'
    elif choice2 == 3:
        trans_lang = 'ja'
    else:
        print('잘못 입력했습니다.')
        raise Exception('1, 2, 3 중 선택해주세요.')
    
    text = input('\n번역할 내용을 작성해 주세요\n')
    
    data = {'text' : text,
            'source' : lang,
            'target': trans_lang}

    url = "https://openapi.naver.com/v1/papago/n2mt"

    header = {"X-Naver-Client-Id":'M2TCtkuAE0f1_4eJTDeS',
              "X-Naver-Client-Secret":'QJ9CRWHp8i'}

    response = requests.post(url, headers=header, data= data)
    rescode = response.status_code

    if(rescode==200):
        print('='*50)
        print(f'원본: {text}')
        print(f'번역: {response.json()["message"]["result"]["translatedText"]}')
        print('='*50)
    else:
        print("Error Code:" , rescode)
    
    start()

def start():
    print('='*50)
    print('사용하실 기능을 선택해주세요.')
    function_choice = int(input('메일발송 = 1, 메일 수신 = 2, 번역기 = 3, 엑셀데이터 DB삽입 = 4, 종료 = 0\n'))
    print('='*50)

    if function_choice == 1:
        # 일반 메일을 전송할 것인지, 정해진 양식 메일 전송할지 선택
        send_choice = int(input('일반 메일 전송 = 1, 정해진 양식 메일 전송 = 2\n'))
        print('='*50)
        if send_choice == 1:
            print('='*50)
            send_email('nomal')
        elif send_choice == 2:
            send_email('excel')
        else:
            print('잘못 입력하셨습니다.\n프로그램을 종료합니다.')    
    elif function_choice == 2:
        receive()
    elif function_choice == 3:
        translate()
    elif function_choice == 4:
        get_excel_data()
    elif function_choice == 0:
        print('프로그램을 종료합니다.')    
    else:
        print('잘못 입력하셨습니다.\n프로그램을 종료합니다.')


start()