# gui1.py (시간표 생성 함수, 시간표 캡쳐 함수, 이메일 전송 함수, 시간표 셀, 입력창, 버튼 생성 및 위치)
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from PIL import ImageGrab
import new1
import pyautogui
import smtplib
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os
import socket


### 시간표 생성 함수### 
def table_make(*event):# 사용자가 '만들기' 버튼을 누르면 실행됨.
    try:
        #사용자가 선택한 값을 new1.py의 listing함수를 이용하여 리스트로 저장함.
        select=new1.listing(str(choice_major.get()),int(choice_level.get()), int(choice_num.get()),int(choice_a.get()),int(choice_score.get()), int(choice_lunch.get()), str(choice_lib1.get()),str(choice_lib2.get()))

        if select[4]==0:#수강학점을 선택하지 않았을 때 .
            messagebox.showwarning('Warning', '수강학점을 선택해주세요')#경고문구를 띄움.
        else: #수강학점을 선택하면
            data=new1.convert_excel('data')#new1.py의 convert_excel로 과목데이터를 읽어와서 튜플로 변환함. 
            major,score=new1.major(select,data)#new1.py의 전공선택 함수를 실행하고 반환값(과목명을 저장하는 2차원 리스트, 남은 학점)을 저장함.
            table,score=new1.lib(select,major,data,score)#new1.py의 영역별교양 함수를 실행하고 반환값(과목명을 저장하는 2차원 리스트, 남은 학점)을 저장함.

            ###최종적으로 선택된 과목명이 저장된 2차원 리스트의 값들을 출력되는 시간표의 'text'변수에 저장함.###
            mon = table[0]
            tue = table[1]
            wed = table[2]
            thu = table[3]
            fri = table[4]

            # 월요일
            label_11['text'] = mon[0]
            label_12['text'] = mon[1]
            label_13['text'] = mon[2]
            label_14['text'] = mon[3]
            label_15['text'] = mon[4]
            label_16['text'] = mon[5]
            label_17['text'] = mon[6]
            label_18['text'] = mon[7]
            label_19['text'] = mon[8]
            label_110['text'] = mon[9]

            # 화요일
            label_21['text'] = tue[0]
            label_22['text'] = tue[1]
            label_23['text'] = tue[2]
            label_24['text'] = tue[3]
            label_25['text'] = tue[4]
            label_26['text'] = tue[5]
            label_27['text'] = tue[6]
            label_28['text'] = tue[7]
            label_29['text'] = tue[8]
            label_210['text'] = tue[9]

            # 수요일
            label_31['text'] = wed[0]
            label_32['text'] = wed[1]
            label_33['text'] = wed[2]
            label_34['text'] = wed[3]
            label_35['text'] = wed[4]
            label_36['text'] = wed[5]
            label_37['text'] = wed[6]
            label_38['text'] = wed[7]
            label_39['text'] = wed[8]
            label_310['text'] = wed[9]

            # 목요일
            label_41['text'] = thu[0]
            label_42['text'] = thu[1]
            label_43['text'] = thu[2]
            label_44['text'] = thu[3]
            label_45['text'] = thu[4]
            label_46['text'] = thu[5]
            label_47['text'] = thu[6]
            label_48['text'] = thu[7]
            label_49['text'] = thu[8]
            label_410['text'] = thu[9]

            # 금요일
            label_51['text'] = fri[0]
            label_52['text'] = fri[1]
            label_53['text'] = fri[2]
            label_54['text'] = fri[3]
            label_55['text'] = fri[4]
            label_56['text'] = fri[5]
            label_57['text'] = fri[6]
            label_58['text'] = fri[7]
            label_59['text'] = fri[8]
            label_510['text'] = fri[9]   
                
            #시간표가 완성되었음을 알려주는 문구와 생성된 시간표의 총 학점을 출력함.
            messagebox.showinfo('Complete!', '총 %d학점의 시간표가 완성되었습니다.'%(select[4]-score))
    except: #오류: 사용자가 값을 아무값도 입력하지 않았을 경우
        messagebox.showwarning('Warning', '학과, 학년, 전공과목 수, 영역별 교양1, 영역별 교양2에 \n값을 입력했는지 확인해주세요.')#경고문구를 출력


### 생성된 시간표 캡쳐###
def capture(*event):#캡쳐버튼을 누르면 실행됨.
    x, y = pyautogui.position()# 현재 화살표의 위치
    icon = ImageGrab.grab(bbox=(x - 1120, y - 670, x - 350, y + 120)) #캡쳐되는 범위
    s = 'table.png'#캡쳐된 시간표의 파일명
    icon.save(s)  # 캡쳐된 시간표 저장
    messagebox.showinfo('Complete!', '저장이 완료되었습니다.') #저장이 완료되었음을 출력함.


### 캡쳐된 시간표를 이메일로 전송하는 함수-Gmail####
def send_email(*event):#'보내기'버튼을 누르면 실행됨
    email_send= str(email_entry.get())#사용자로부터 입력받은 이메일 주소
    email_user='1234@gmail.com' #(수정)보내는 메일 주소
    email_password='1234*' #(수정)보내는 메일 계정의 비밀번호
    subject='시간표' #보내는 메일의 제목
    file = 'table.png' #보내는 메일의 파일명
    body = '시간표사진' #보내는 메일의 내용

    try:
        msg=MIMEMultipart()
        msg['From']=email_user
        msg['To']=email_send
        msg['Subject']=subject
        
        msg.attach(MIMEText(body,'plain'))
        part = MIMEBase('application','octet-stream')
        part.set_payload(open(file,'rb').read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition','attachment; filename="%s"'% os.path.basename(file))
        msg.attach(part)


        text = msg.as_string()
        server = smtplib.SMTP('smtp.gmail.com',587)
        server.starttls()
        server.login(email_user,email_password)

        server.sendmail(email_user,email_send,text)
        server.quit()
        messagebox.showinfo('Complete!', '전송이 완료되었습니다.')#전송이 완료되었음을 출력함.
        

    ##이메일 전송의 오류처리##
        
    # 캡쳐를 안한 상태에서 이메일 전송을 시도하면 에러가 생김
    except FileNotFoundError:
        messagebox.showwarning('Warning', '파일의 생성유무와 경로를 확인해 주십시오.')


    # 이메일 주소가 잘못되었을 경우 생기는 에러
    except smtplib.SMTPRecipientsRefused:
        email_entry.delete(0, END)
        messagebox.showwarning('Warning', '이메일 주소를 잘못 입력하셨습니다.')


    # 인터넷 연결이 안되어 있을 경우 생기는 에러
    except socket.gaierror:
        messagebox.showwarning('Warning', '인터넷연결을 확인해 주십시오.')


if __name__ == '__main__':
    win = Tk()
    win.title("시간표만들기")#위젯의 제목
    choice_major = StringVar()  # 전공 선택 값을 입력받는 문자열 변수
    choice_level = IntVar()  # 학년 선택 값을 입력받는 정수 변수
    choice_num = IntVar()  # 들을 전공 수의 선택 값을 입력 받는 정수 변수
    choice_a = IntVar() #A교시 포함여부
    choice_score = IntVar()#수강학점 선택
    choice_lunch = IntVar(value=3)#점심시간 선택
    choice_lib1 = StringVar()#영역별 교양1 선택 값을 입력받는 문자열 변수
    choice_lib2 = StringVar()#영역별 교양2 선택 값을 입력받는 문자열 변수


    #########label로 표 셀 만들기###########

    ##요일명이 들어갈 셀##
    empty=Label(win, font=('맑은 고딕', 10),width=3)
    row_title_label1 = Label(win, text='월', font=('맑은 고딕', 10), bg='#b2e1ff',
                             relief='ridge', width=13)
    row_title_label2 = Label(win, text='화', font=('맑은 고딕', 10), bg='#b2e1ff',
                             relief='ridge', width=13)
    row_title_label3 = Label(win, text='수', font=('맑은 고딕', 10), bg='#b2e1ff',
                             relief='ridge', width=13)
    row_title_label4 = Label(win, text='목', font=('맑은 고딕', 10), bg='#b2e1ff',
                             relief='ridge', width=13)
    row_title_label5 = Label(win, text='금', font=('맑은 고딕', 10), bg='#b2e1ff',
                             relief='ridge', width=13)
    ##시간명이 들어갈 셀##
    col_title_label1 = Label(win, text='A교시', font=('맑은 고딕', 10), bg='#b2e1ff',
                             relief='ridge', width=6, height=3)
    col_title_label2 = Label(win, text='B교시', font=('맑은 고딕', 10), bg='#b2e1ff',
                             relief='ridge', width=6, height=3)
    col_title_label3 = Label(win, text='C교시', font=('맑은 고딕', 10), bg='#b2e1ff',
                             relief='ridge', width=6, height=3)
    col_title_label4 = Label(win, text='D교시', font=('맑은 고딕', 10), bg='#b2e1ff',
                             relief='ridge', width=6, height=3)
    col_title_label5 = Label(win, text='E교시', font=('맑은 고딕', 10), bg='#b2e1ff',
                             relief='ridge', width=6, height=3)
    col_title_label6 = Label(win, text='F교시', font=('맑은 고딕', 10), bg='#b2e1ff',
                             relief='ridge', width=6, height=3)
    col_title_label7 = Label(win, text='6교시', font=('맑은 고딕', 10), bg='#b2e1ff',
                             relief='ridge', width=6, height=3)
    col_title_label8 = Label(win, text='7교시', font=('맑은 고딕', 10), bg='#b2e1ff',
                             relief='ridge', width=6, height=3)
    col_title_label9 = Label(win, text='8교시', font=('맑은 고딕', 10), bg='#b2e1ff',
                             relief='ridge', width=6, height=3)
    col_title_label10 = Label(win, text='9교시', font=('맑은 고딕', 10), bg='#b2e1ff',
                             relief='ridge', width=6, height=3)

    ##과목명이 들어갈 셀##
    
    # 월요일
    label_11 = Label(win, text='', font=('맑은 고딕', 10), relief='ridge', width=13, height=3, wraplength=120)
    label_12 = Label(win, text='', font=('맑은 고딕', 10), relief='ridge', width=13, height=3, wraplength=120)
    label_13 = Label(win, text='', font=('맑은 고딕', 10), relief='ridge', width=13, height=3, wraplength=120)
    label_14 = Label(win, text='', font=('맑은 고딕', 10), relief='ridge', width=13, height=3, wraplength=120)
    label_15 = Label(win, text='', font=('맑은 고딕', 10), relief='ridge', width=13, height=3, wraplength=120)
    label_16 = Label(win, text='', font=('맑은 고딕', 10), relief='ridge', width=13, height=3, wraplength=120)
    label_17 = Label(win, text='', font=('맑은 고딕', 10), relief='ridge', width=13, height=3, wraplength=120)
    label_18 = Label(win, text='', font=('맑은 고딕', 10), relief='ridge', width=13, height=3, wraplength=120)
    label_19 = Label(win, text='', font=('맑은 고딕', 10), relief='ridge', width=13, height=3, wraplength=120)
    label_110 = Label(win, text='', font=('맑은 고딕', 10), relief='ridge', width=13, height=3, wraplength=120)

    # 화요일
    label_21 = Label(win, text='', font=('맑은 고딕', 10), relief='ridge', width=13, height=3, wraplength=120)
    label_22 = Label(win, text='', font=('맑은 고딕', 10), relief='ridge', width=13, height=3, wraplength=120)
    label_23 = Label(win, text='', font=('맑은 고딕', 10), relief='ridge', width=13, height=3, wraplength=120)
    label_24 = Label(win, text='', font=('맑은 고딕', 10), relief='ridge', width=13, height=3, wraplength=120)
    label_25 = Label(win, text='', font=('맑은 고딕', 10), relief='ridge', width=13, height=3, wraplength=120)
    label_26 = Label(win, text='', font=('맑은 고딕', 10), relief='ridge', width=13, height=3, wraplength=120)
    label_27 = Label(win, text='', font=('맑은 고딕', 10), relief='ridge', width=13, height=3, wraplength=120)
    label_28 = Label(win, text='', font=('맑은 고딕', 10), relief='ridge', width=13, height=3, wraplength=120)
    label_29 = Label(win, text='', font=('맑은 고딕', 10), relief='ridge', width=13, height=3, wraplength=120)
    label_210 = Label(win, text='', font=('맑은 고딕', 10), relief='ridge', width=13, height=3, wraplength=120)

    # 수요일
    label_31 = Label(win, text='', font=('맑은 고딕', 10), relief='ridge', width=13, height=3, wraplength=120)
    label_32 = Label(win, text='', font=('맑은 고딕', 10), relief='ridge', width=13, height=3, wraplength=120)
    label_33 = Label(win, text='', font=('맑은 고딕', 10), relief='ridge', width=13, height=3, wraplength=120)
    label_34 = Label(win, text='', font=('맑은 고딕', 10), relief='ridge', width=13, height=3, wraplength=120)
    label_35 = Label(win, text='', font=('맑은 고딕', 10), relief='ridge', width=13, height=3, wraplength=120)
    label_36 = Label(win, text='', font=('맑은 고딕', 10), relief='ridge', width=13, height=3, wraplength=120)
    label_37 = Label(win, text='', font=('맑은 고딕', 10), relief='ridge', width=13, height=3, wraplength=120)
    label_38 = Label(win, text='', font=('맑은 고딕', 10), relief='ridge', width=13, height=3, wraplength=120)
    label_39 = Label(win, text='', font=('맑은 고딕', 10), relief='ridge', width=13, height=3, wraplength=120)
    label_310 = Label(win, text='', font=('맑은 고딕', 10), relief='ridge', width=13, height=3, wraplength=120)

    # 목요일
    label_41 = Label(win, text='', font=('맑은 고딕', 10), relief='ridge', width=13, height=3, wraplength=120)
    label_42 = Label(win, text='', font=('맑은 고딕', 10), relief='ridge', width=13, height=3, wraplength=120)
    label_43 = Label(win, text='', font=('맑은 고딕', 10), relief='ridge', width=13, height=3, wraplength=120)
    label_44 = Label(win, text='', font=('맑은 고딕', 10), relief='ridge', width=13, height=3, wraplength=120)
    label_45 = Label(win, text='', font=('맑은 고딕', 10), relief='ridge', width=13, height=3, wraplength=120)
    label_46 = Label(win, text='', font=('맑은 고딕', 10), relief='ridge', width=13, height=3, wraplength=120)
    label_47 = Label(win, text='', font=('맑은 고딕', 10), relief='ridge', width=13, height=3, wraplength=120)
    label_48 = Label(win, text='', font=('맑은 고딕', 10), relief='ridge', width=13, height=3, wraplength=120)
    label_49 = Label(win, text='', font=('맑은 고딕', 10), relief='ridge', width=13, height=3, wraplength=120)
    label_410 = Label(win, text='', font=('맑은 고딕', 10), relief='ridge', width=13, height=3, wraplength=120)

    # 금요일
    label_51 = Label(win, text='', font=('맑은 고딕', 10), relief='ridge', width=13, height=3, wraplength=120)
    label_52 = Label(win, text='', font=('맑은 고딕', 10), relief='ridge', width=13, height=3, wraplength=120)
    label_53 = Label(win, text='', font=('맑은 고딕', 10), relief='ridge', width=13, height=3, wraplength=120)
    label_54 = Label(win, text='', font=('맑은 고딕', 10), relief='ridge', width=13, height=3, wraplength=120)
    label_55 = Label(win, text='', font=('맑은 고딕', 10), relief='ridge', width=13, height=3, wraplength=120)
    label_56 = Label(win, text='', font=('맑은 고딕', 10), relief='ridge', width=13, height=3, wraplength=120)
    label_57 = Label(win, text='', font=('맑은 고딕', 10), relief='ridge', width=13, height=3, wraplength=120)
    label_58 = Label(win, text='', font=('맑은 고딕', 10), relief='ridge', width=13, height=3, wraplength=120)
    label_59 = Label(win, text='', font=('맑은 고딕', 10), relief='ridge', width=13, height=3, wraplength=120)
    label_510 = Label(win, text='', font=('맑은 고딕', 10), relief='ridge', width=13, height=3, wraplength=120)


    ########버튼, 입력창 생성########

    #만들기 버튼
    build_but = Button(win, text='만들기', command=table_make)

    #전공
    major_label = Label(win, text='학과', font=('맑은 고딕', 11),anchor= 'ne', width=15)
    combo_major = ttk.Combobox(win, textvariable=choice_major)
    combo_major['value'] = ('경영학과', '금융공학과','생명공학과', '문화콘텐츠학과','소프트웨어학과')

    #학년
    level_label = Label(win, text='학년', font=('맑은 고딕', 11),anchor= 'ne', width=15)
    combo_level = ttk.Combobox(win, textvariable=choice_level,width=3)
    combo_level['value']=(1,2,3,4)

    #전공과목 수
    num_label = Label(win, text='전공과목 수', font=('맑은 고딕', 11),anchor= 'ne', width=15)
    combo_num = ttk.Combobox(win, textvariable=choice_num,width=3)
    combo_num['value']=(0,1,2,3,4,5,6,7)

    #A교시 포함여부 
    a_label = Label(win, text='A교시', font=('맑은 고딕', 11),anchor= 'ne', width=15)

    a_but1 = Radiobutton(win, text='포함', variable=choice_a, value=1)
    a_but2 = Radiobutton(win, text='제외', variable=choice_a, value=0)

    #희망 수강학점
    score_label = Label(win, text='     수강학점', font=('맑은 고딕', 11),anchor= 'ne', width=15)

    combo_score = ttk.Combobox(win, textvariable=choice_score,width=3)
    combo_score['value']=(0,3,6,9,10,11,12,13,14,15,16,17,18,19,21,22,23,24)
    mes1 = Label(win, text='※선택한 학점이 초과되지 않는 범위에서 시간표가 생성됩니다.', font=('맑은 고딕', 11))
    
    #점심시간 선택
    lunch_label = Label(win, text='점심시간', font=('맑은 고딕', 11),anchor= 'ne', width=15)

    lunch_box1 = Radiobutton(win, text='C교시 12:00-13:15', variable=choice_lunch,value=1)
    lunch_box2 = Radiobutton(win, text='D교시 13:30-14:45', variable=choice_lunch,value=2)
    lunch_box3 = Radiobutton(win, text='고려하지 않음', variable=choice_lunch,value=3)

    #영역별교양
    lib1_label = Label(win, text='영역별교양1', font=('맑은 고딕', 11),anchor= 'ne', width=15)

    lib1 = ttk.Combobox(win, textvariable=choice_lib1)
    lib1['value'] = ('선택 안 함','역사와 철학', '문학과 예술', '인간과 사회', '자연과 과학')
    lib2_label = Label(win, text='영역별교양2', font=('맑은 고딕', 11),anchor= 'ne', width=15)

    lib2 = ttk.Combobox(win, textvariable=choice_lib2)
    lib2['value'] = ('선택 안 함','역사와 철학', '문학과 예술', '인간과 사회', '자연과 과학')
    
    #캡쳐 버튼
    capture = Button(win, text='캡쳐', command=capture)

    #이메일
    email_label = Label(win, text='E-mail', font=('맑은 고딕', 11))
    email_entry = Entry(win,width=23)
    email_but = Button(win, text='보내기', command=send_email)
    mes = Label(win, text='※이 정보는 이메일을 보내는 용도로만 사용됩니다.', font=('맑은 고딕', 11))

    
    

    ######## 시간표 셀 배치###########
    empty.grid(column=0, row=0)
    row_title_label1.grid(column=2, row=0)
    row_title_label2.grid(column=3, row=0)
    row_title_label3.grid(column=4, row=0)
    row_title_label4.grid(column=5, row=0)
    row_title_label5.grid(column=6, row=0)

    col_title_label1.grid(column=1, row=1)
    col_title_label2.grid(column=1, row=2)
    col_title_label3.grid(column=1, row=3)
    col_title_label4.grid(column=1, row=4)
    col_title_label5.grid(column=1, row=5)
    col_title_label6.grid(column=1, row=6)
    col_title_label7.grid(column=1, row=7)
    col_title_label8.grid(column=1, row=8)
    col_title_label9.grid(column=1, row=9)
    col_title_label10.grid(column=1, row=10)
    
    
    # 월요일
    label_11.grid(column=2, row=1)
    label_12.grid(column=2, row=2)
    label_13.grid(column=2, row=3)
    label_14.grid(column=2, row=4)
    label_15.grid(column=2, row=5)
    label_16.grid(column=2, row=6)
    label_17.grid(column=2, row=7)
    label_18.grid(column=2, row=8)
    label_19.grid(column=2, row=9)
    label_110.grid(column=2, row=10)

    # 화요일
    label_21.grid(column=3, row=1)
    label_22.grid(column=3, row=2)
    label_23.grid(column=3, row=3)
    label_24.grid(column=3, row=4)
    label_25.grid(column=3, row=5)
    label_26.grid(column=3, row=6)
    label_27.grid(column=3, row=7)
    label_28.grid(column=3, row=8)
    label_29.grid(column=3, row=9)
    label_210.grid(column=3, row=10)

    # 수요일
    label_31.grid(column=4, row=1)
    label_32.grid(column=4, row=2)
    label_33.grid(column=4, row=3)
    label_34.grid(column=4, row=4)
    label_35.grid(column=4, row=5)
    label_36.grid(column=4, row=6)
    label_37.grid(column=4, row=7)
    label_38.grid(column=4, row=8)
    label_39.grid(column=4, row=9)
    label_310.grid(column=4, row=10)

    # 목요일
    label_41.grid(column=5, row=1)
    label_42.grid(column=5, row=2)
    label_43.grid(column=5, row=3)
    label_44.grid(column=5, row=4)
    label_45.grid(column=5, row=5)
    label_46.grid(column=5, row=6)
    label_47.grid(column=5, row=7)
    label_48.grid(column=5, row=8)
    label_49.grid(column=5, row=9)
    label_410.grid(column=5, row=10)

    # 금요일
    label_51.grid(column=6, row=1)
    label_52.grid(column=6, row=2)
    label_53.grid(column=6, row=3)
    label_54.grid(column=6, row=4)
    label_55.grid(column=6, row=5)
    label_56.grid(column=6, row=6)
    label_57.grid(column=6, row=7)
    label_58.grid(column=6, row=8)
    label_59.grid(column=6, row=9)
    label_510.grid(column=6, row=10)

   
    #########입력창,버튼 위치##########
    #전공
    major_label.grid(column=7,row=0)
    combo_major.grid(column=8,row=0,sticky=W)

    #학년
    level_label.grid(column=7,row=1)
    combo_level.grid(column=8,row=1,sticky=W)
    
    #들을 전공 수
    num_label.grid(column=7,row=2)
    combo_num.grid(column=8,row=2,sticky=W)

    #A교시 포함여부 
    a_label.grid(column=7,row=3)
    a_but1.grid(column=8,row=3,sticky=W)
    a_but2.grid(column=8,row=3)

    #희망 수강학점 
    score_label.grid(column=7,row=4)
    combo_score.grid(column=8,row=4,sticky=W)
    mes1.grid(column=9,row=4)
    #점심시간 선택
    lunch_label.grid(column=7,row=5)
    lunch_box1.grid(column=8,row=5,sticky=E)
    lunch_box2.grid(column=9,row=5,sticky=W)
    lunch_box3.grid(column=9,row=5)

    #영역별교양
    lib1_label.grid(column=7,row=6)
    lib1.grid(column=8,row=6)
    lib2_label.grid(column=7,row=7)
    lib2.grid(column=8,row=7)

    # 만들기 버튼 위치
    build_but.grid(column=8, row=9,sticky=E+W)
    capture.grid(column=9,row=9,sticky=W)
    
    #이메일 위치
    email_label.grid(column=7, row=10, sticky=E)
    email_entry.grid(column=8, row=10, columnspan=3, sticky=W)
    email_but.grid(column=9, row=10, sticky=W)
    mes.grid(column=8, row=11, columnspan=7, sticky=N)


    win.mainloop()
    
