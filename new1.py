# new1.py (엑셀파일 튜플로 전환함수, 전공선택함수, 영역별 교양선택함수, 그 외)
from pandas import DataFrame
import pandas as pd
import random
from tkinter import *
from tkinter import messagebox

##### 엑셀파일을 튜플로 전환하는 함수 #####
def convert_excel(name):#파일의 이름을 매개변수로 함.
    file_name=str(name)#엑셀파일의 이름, 파일이름은 data이어야 함.
    full = 'C:/Users/USER/AppData/Local/Programs/Python/Python38-32/금소웨 기말 프로젝트/'+file_name + '.xlsx' #사용자 컴퓨터의 엑셀파일 저장위치를 입력할 것
    data = []

    df = pd.read_excel(full)
    frame = DataFrame(df)

    index = frame.index

    for i in index:
        a = df.loc[i].values  # 배열로 값이 생성된다.
        line = tuple(a)  # 튜플로 바꾼다.
        data.append(line)
    return data #과목 데이터를 2차원 튜플로 반환


##### 사용자에게 입력받은 값들을 리스트로 변환할 함수.######
def listing(*variable):#사용자로부터 입력받은 값을 매개변수로 함.
    lis=list(variable)#입력받은 값 리스트로 변환
    return lis#변환한 리스트 반환


#####시간표의 과목명을 저장하는 리스트가 채워져 있는지 확인하는 함수#####
def Filled(n):#확인할 리스트를 매개변수로 함.
    if n==[]: return False #비어있으면 거짓
    return True # 채워져 있으면 참


#####선택할 수 있는 과목들 중 이미 저장된 과목들과 시간이 겹치는 것을 삭제하는 함수#####
def del_filled(table,lis):
#과목명을 저장하는 2차원 리스트와 선택할 수 있는 과목들을 모아 놓은 리스트를 매개변수로 함. 
    i=0
    while True:#무한 루프
        if Filled(table[lis[i][2][0]][lis[i][2][1]]) or Filled(table[lis[i][2][2]][lis[i][2][3]]):#위에서 정의한 Filled 함수 사용
            #선택할 수 있는 과목들의 시간에 해당하는 table(시간표의 과목명을 저장하는 리스트)의 칸이 채워져 있으면
                
            lis.pop(i)#해당 값을 리스트에서 삭제함.
            i-=1
        i+=1
        if i==len(lis):break #i가 리스트의 크기와 같아지면 무한 루프를 빠져나옴.
    return lis # 시간이 겹치지 않는 과목들만 있는 리스트를 반환함.


##### 문자로 입력된 요일과 시간을 숫자로 변환하는 함수#####
def time_trans(lis):#변환해야하는 값이 저장된 리스트를 매개변수로 함.
    for i in range(len(lis)):
        #행의 인덱스
        if lis[i]=='월': lis[i]=0
        if lis[i]=='화': lis[i]=1
        if lis[i]=='수': lis[i]=2
        if lis[i]=='목': lis[i]=3
        if lis[i]=='금': lis[i]=4
        #열의 인덱스
        if lis[i]=='A': lis[i]=0
        if lis[i]=='B': lis[i]=1
        if lis[i]=='C': lis[i]=2
        if lis[i]=='D': lis[i]=3
        if lis[i]=='E': lis[i]=4
        if lis[i]=='F': lis[i]=5
    return lis#숫자로 변환된 값이 저장된 리스트를 반환함.


#####시간표의 과목명을 저장하는 리스트 생성하는 함수#####
def time_table():#월,화,수,목,금 /A~F교시: 2차원 리스트
    monday = [[], [], [], [], [], [], [], [], [], []]
    tuesday = [[], [], [], [], [], [], [], [], [], []]
    wednesday = [[], [], [], [], [], [], [], [], [], []]
    thursday = [[], [], [], [], [], [], [], [], [], []]
    friday = [[], [], [], [], [], [], [], [], [], []]
    return [monday, tuesday, wednesday, thursday, friday]#생성한 2차원 리스트 반환


#####사용자의 A교시 포함여부와 선택한 점심시간에 따라 칸을 채우는 함수#####
def A_lunch(select,table):#사용자가 선택한 값이 저장된 리스트와 시간표의 과목명을 저장하는 2차원 리스트를 매개변수로 함.
    if select[3]==0: #A교시 제외
        for i in range(len(table)):
            table[i][0]='\t' #띄어쓰기으로 칸을 채움
    if select[5]==1:#점심시간 C교시선택
        for i in range(len(table)): 
            table[i][2]='\t' #띄어쓰기으로 칸을 채움
    elif select[5]==2: #점심시간 D교시 선택
        for i in range(len(table)): 
            table[i][3]='\t' #띄어쓰기로 칸을 채움
    return table#사용자가 선택한 대로 칸을 채운 2차원(행(요일),열(시간)) 리스트 반환

   
#####전공과목을 선택하는 함수#####
def major(select,data):#사용자가 선택한 값을 저장하는 리스트와, 과목데이터를 매개변수로 함.
    table=A_lunch(select,time_table())# A_lunch함수를 실행하여 전달받은 값을 변수 table에 저장함.
    score=select[4]#사용자가 선택한 수강학점
    major=select[0]#사용자가 선택한 학과
    level=select[1]#사용자가 선택한 학년
    num=select[2]#사용자가 선택한 전공과목 수
    major_lis=[]# 조건에 맞는 과목들을 모을 리스트
    
    while num>0 and score>0:#선택한 전공과목 수가 0이되거나 학점이 0이 될때까지 반복
        if num==0: break #사용자가 전공과목 수 0을 선택했을 경우 반복을 빠져나옴

        else: #선택한 전공과목 수가 0이 아닐 경우
            for i in range(len(data)):#과목 데이터의 길이만큼 반복
                if data[i][2]==major and data[i][1]==level:#과목 데이터에 저장된 학년,학과명이 사용자가 선택한 값과 같으면
                    major_lis.append([data[i][0],data[i][3],time_trans(list(data[i][4:8])),data[i][8]])
                    #리스트에 해당 과목의 [과목 고유번호,과목명,과목시간(함수 time_trans로 숫자로변환->(행1(요일1),열1(시간1),행2(요일2),열(시간1)),학점]을 추가함.       
            major_lis=del_filled(table,major_lis)# 함수 del_filled를 사용하여 시간이 겹치는 과목들을 제거함.
            try:
                while True:
                    choose=random.choice(major_lis)#남은 과목들 중에서 무작위로 하나를 추출함.
                     
                    if choose[3]<= score:#남은 학점보다 선택된 과목의 학점이 작거나 같은 경우
                        #추출된 과목의 시간대로 과목명을 채움.
                        table[choose[2][0]][choose[2][1]]= choose[1]
                        table[choose[2][2]][choose[2][3]]= choose[1]  
                        score-=choose[3]#과목의 학점만큼 학점을 뺌
                        num-=1#전공과목 수에서 1 뺌.
                        break; # 무한 루프에서 빠져나옴.
                    else: major_lis.remove(choose) #남은 학점보다 선택된 과목의 학점이 큰 경우 선택된 과목을 리스트에서 삭제함.
                
            except:# 오류가 날 경우: 사용자가 선택한 전공과목 수가 추가할 수 있는 과목 수보다 클 경우 
                messagebox.showinfo('Information', ' 선택한 전공 과목 수가 추가 할 수 있는 전공과목 수 보다 큽니다.\n추가 가능한 전공과목만 시간표에 포함됩니다.')
                #경고 문구를 출력하고
                return table, score #함수를 종료하고 선택할 수 있는 과목명들만 채운 2차원 리스트와 남은 학점을 저장한 변수를 반환함. 

    return table,score #선택된 전공과목명을 채워넣은 2차원리스트와 선택하고 남은 학점을 반환함.




#####영역별 교양을 선택하는 함수#####
def lib(select,major,data,score):
    #사용자가 선택한 값을 저장하는 리스트, 전공선택함수에서 전달받은 과목명을 저장하는 2차원 리스트, 과목데이터, 남은 학점을 저장한 변수를 매개변수로 함.
    
    af_major=major#전공 함수에서 전달받은 과목명을 저장하는 리스트
    lib1=select[6] #사용자가 선택한 영역별 교양1의 분야
    lib2=select[7]#사용자가 선택한 영역별 교양2의 분야

    # 조건에 맞는 과목들을 모을 리스트
    lib1_lis=[] #영역별 교양 1
    lib2_lis=[] #영역별 교양 2
    
    while score>=3:# 남은 학점이 3이상이면 계속 반복함.
        if lib1=='선택 안 함' and lib2=='선택 안 함': break #사용자가 둘 다 '선택 안 함'을 선택했으면 반복을 빠져 나옴. 

        if lib1!='선택 안 함':#영역별 교양1이 '선택 안 함'이 아닐 경우
            for i in range(len(data)):#과목 데이터의 길이만큼 반복
                if data[i][2]==lib1:#과목 데이터에 저장된 영역별 교양의 분야와 사용자가 선택한 값이 같으면
                    lib1_lis.append([data[i][0],data[i][3],time_trans(list(data[i][4:8])),data[i][8]])
                    #리스트에 해당 과목의 [과목 고유번호,과목명,과목시간(함수 time_trans로 숫자로변환->(행1(요일1),열1(시간1),행2(요일2),열(시간1)),학점]을 추가함.
                     
            lib1_lis=del_filled(af_major,lib1_lis)# 함수 del_filled를 사용하여 시간이 겹치는 과목들을 제거함.
            try:
                choose=random.choice(lib1_lis)#남은 과목들 중에서 무작위로 하나를 추출함.

                #추출된 과목의 시간대로 과목명을 채움.
                af_major[choose[2][0]][choose[2][1]]= choose[1] #요일1
                af_major[choose[2][2]][choose[2][3]]= choose[1] #요일2                      
                score-=choose[3]#과목의 학점만큼 학점을 뺌
                
            except:# 오류가 날 경우: 남은 학점만큼 더 추가 할 수 있는 과목 수 보다 조건에 맞는 과목 수가 작을 경우.
                messagebox.showinfo('Information', '남은 학점만큼 추가할 수 있는 영역별 교양이 없습니다.\n영역별 교양추가를 더 원하시면 "A교시 포함" 또는 "점심시간 고려하지 않음" 또는 영역별 교양1, 2를  모두 선택해주세요.')
                #경고 문구를 출력하고
                return af_major, score #함수를 종료하고 선택할 수 있는 과목명들만 채운 2차원 리스트와 남은 학점을 저장한 변수를 반환함. 

        #영역별 교양1과 같음.    
        if lib2!='선택 안 함' and score>=3:
            for i in range(len(data)):
                if data[i][2]==lib2:
                    lib2_lis.append([data[i][0],data[i][3],time_trans(list(data[i][4:8])),data[i][8]])#과목 고유번호,과목명,과목시간,학점(숫자로 변환->(행1(요일1),열1(시간1),행2(요일2),열(시간1))
            lib2_lis=del_filled(af_major,lib2_lis)
            try:
                choose=random.choice(lib2_lis)
                af_major[choose[2][0]][choose[2][1]]= choose[1]
                af_major[choose[2][2]][choose[2][3]]= choose[1]                       
                score-=choose[3]
            except: 
                messagebox.showinfo('Information', '남은 학점만큼 추가할 수 있는 영역별 교양이 없습니다.\n영역별 교양추가를 더 원하시면 "A교시 포함" 또는 "점심시간 고려하지 않음"또는 영역별 교양1,2를  모두 선택해주세요.')
                return af_major, score
    return af_major, score#선택된 교양과목명을 채워넣은 2차원리스트와 선택하고 남은 학점을 반환함.
