# 메일 발송 프로그램
> ## 1.프로젝트 구성을 위한 필수 확인 항목
>-라이브러리 설치: pip install -r requirements.txt
>- smtplib
>- email
>- datetime
>- openpyxl
>- imap_tools
>- requests
</br>
 
 > ## 2. 주요 기능
 > - 일반메일, 첨부파일 전송
 > - 일반메일, 첨부파일 수신
 > - 첨부파일 다운로드
 > - 원하는 양의 메일의 수를 설정해 원하는 언어로 번역
 </br>
 
 
 
 > ## 3. 서비스 구성도
![serviceFrom](https://user-images.githubusercontent.com/87023534/129568837-bcabdb2a-d66f-4a1e-a0ac-5582889ca954.png)
</br>

> ##	4.Smtp 주요 메서드

> - #### < smtplib.SMTP('stmp server url', port) >
SMTP 또는 ESMTP 서버에 대한 연결을 관리한다.
> - #### < stmp.ehlo() >
SMTP개체를 얻고나면 ehlo() 메소드를 호출하여 SMTP 이메일 서버에 Hello 메세지를 보낸다. Hello 메세지는 SMTP에서 첫번째 단계이자 서버 연결을 설정하는 중요한 단계이다. 핸드셰이킹(handshaking) 시도
> - #### < stmp.starttls() >
SMTP 서버의 포트 587에 연결할 때(TLS 암호화 사용)에는 starttls() 메소드를 호출해야한다. 위 필수단계를 통해 연결을 암호화할 수 있다.
> - #### < smtp.login('id', 'password') >
SMTP서버로 암호화된 연결을 설정하고 나면 login() 메소드를 호출하여 사용자이름(보통 이메일 주소) 및 이메일 비밀번호를 사용하여 로그인할 수 있다.
> - #### < smtp.sendmail(from_addr, to_addrs, msg, mail_options) >
로그인한 SMTP 서버에 메일 전송
> - #### < smtp.quit() >
연결종료
