import pandas as pd
import re

df = pd.read_excel('종합강의시간표내역(학부).xlsx')
data_list = df.values.tolist()  #엑셀 입력



def get_info_from_list(target_course_num): #강의실 정보 가져오기
    
    COURSE_NUM_IDX = 4   # '과목번호' 열 위치
    ROOM_INFO_IDX = 13   # '강의요시/강의실' 열 위치

    if str(target_course_num) == "0000":
        return ["모시래 본관", 0]

    elif str(target_course_num) == "0001":
        return ["모시래 여동신관", 0]

    elif str(target_course_num) == "0002":
        return ["해오름", 0]
    else: 
        for row in data_list:
            if str(row[COURSE_NUM_IDX]) == str(target_course_num):
                raw_room = str(row[ROOM_INFO_IDX])
            
                match = re.search(r'\((.*?)\s*(\d{3,4})',raw_room)

                if match:
                    building = match.group(1).strip()
                    room_num = match.group(2)
                    floor = room_num[:-2]  # 뒤 2자리를 뺀 앞자리가 층수
                    return building, floor




                
        return "해당 과목 없음", "해당 과목 없음"

def set_2timetable(): # 과목번호 입력, 시간표 만들기?
    timetable2 = []

    while len(timetable2) != 2:
        subject = input('''과목번호를 입력하시오. 
모시래 본관: 0000
모시래 여동신관: 0001
해오름: 0002 
입력 : ''')

        if len(subject) != 4:
            print("유효하지 않은 형식입니다")

        elif subject in timetable2:
            print("중복된 값입니다.")

        else:
            timetable2.append(subject)

    return timetable2

def set_timetable(): # 과목번호 입력, 시간표 만들기?
    timetable = []

    while True:
        subject = input('''과목번호를 입력하시오. 
        모시래 본관: 0000
        모시래 여동신관: 0001
        해오름: 0002 
        다 입력하면 '0'혹은 '종료를 누르시오''')

        if subject == "종료" or subject == "0":
            break

        elif len(subject) != 4:
            print("유효하지 않은 형식입니다")

        elif subject in timetable:
            print("중복된 값입니다.")

        else:
            timetable.append(subject)


    return timetable

def personsl_information(): #이동시간 측정을 위한 개인정보 수집
    personsl_information = {'키': None, '속도': None, '성별': None}

    while True:
        subject = input("속도를 선택하시오(걷기: 1, 빠르게 걷기: 2, 뛰기: 3) : ")

        if subject == "1":
            personsl_information["속도"] = "1"
            break
        elif subject == "2":
            personsl_information["속도"] = "2"
            break
        elif subject == "3":
            personsl_information["속도"] = "3"
            break
        else:
            print("1,2,3 중 하나를 입력하세요.")

    while True:
        subject = input("키를 입력하시오(단위 : cm) : ")

        if subject.isdigit():
            personsl_information["키"]=subject
            break
        else:
            print("입력 형식 오류")

    gender=None

    while True:
        subject = input("성별을 입력하시오 (남/여) : ")

        if subject in ("남","남자"):
            personsl_information["성별"]="남"
            break
        if subject in ("여","여자"):
            personsl_information["성별"]="여"
            break
        else:
            print("입력 형식 오류")
    return personsl_information

def speed(mod, hight, gender): #분당 이동 거리(m)
    hight=float(hight)
    if mod == "1":
        if gender == "남":
            speed = (hight *0.415)*100
            return speed/100
        if gender == "여":
            speed = (hight * 0.413)*100
            return speed/100
    if mod == "2":
        speed = (hight *0.45)*120
        return speed/100
    if mod == "3":
        speed = hight*0.65*150
        return speed/100

def turn_alpadet(info1, info2): #과목을 알파벳으로 바꾸기
    subject1 = None
    subject2 = None
    if info1 == "글로컬이음관":
        subject1 = "a"
    elif info1 == "건국체육관":
        subject1 = "b"
    elif info1 == "자연과학관":
        subject1 = "c"
    elif info1 == "국제교육관":
        subject1 = "d"
    elif info1 == "인문사회관":
        subject1 = "e"
    elif info1 == "KU-Station":
        subject1 = "f"
    elif info1 == "상허연구동":
        subject1 = "g"
    elif info1 == "창의예술관":
        subject1 = "h"
    elif info1 == "교수연구동":
        subject1 = "i"
    elif info1 == "복합실습동":
        subject1 = "j"
    elif info1 == "생명과학관":
        subject1 = "k"
    elif info1 == "실시간 온라인":
        subject1 = "l"
    elif info1 == "골프연습장":
        subject1 = "m"
    elif info1 == "모시래 본관":
        subject1 = "n"
    elif info1 == "모시래 여동신관":
        subject1 = "o"
    elif info1 == "해오름":
        subject1 = "p"



    if info2 == "글로컬이음관":
        subject2 = "a"
    elif info2 == "건국체육관":
        subject2 = "b"
    elif info2 == "자연과학관":
        subject2 = "c"
    elif info2 == "국제교육관":
        subject2 = "d"
    elif info2 == "인문사회관":
        subject2 = "e"
    elif info2 == "KU-Station":
        subject2 = "f"
    elif info2 == "상허연구동":
        subject2 = "g"
    elif info2 == "창의예술관":
        subject2 = "h"
    elif info2 == "교수연구동":
        subject2 = "i"
    elif info2 == "복합실습동":
        subject2 = "j"
    elif info2 == "생명과학관":
        subject2 = "k"
    elif info2 == "실시간 온라인":
        subject2 = "l"
    elif info2 == "골프연습장":
        subject2 = "m"
    elif info2 == "모시래 본관":
        subject2 = "n"
    elif info2 == "모시래 여동신관":
        subject2 = "o"
    elif info2 == "해오름":
        subject2 = "p"

    resuert=[subject1, subject2]
    return resuert

def longth_info(info1, info2): #알파벳으로 거리정보 추출
    if info1 == info2:
        return 0
    if info1 in ("a","b") and info2 in ("a","b") and info1 != info2:
        return 112
    if info1 in ("a","c") and info2 in ("a","c") and info1 != info2:
        return 230
    if info1 in ("a","d") and info2 in ("a","d") and info1 != info2:
        return 260
    if info1 in ("a","e") and info2 in ("a","e") and info1 != info2:
        return 230
    if info1 in ("a","f") and info2 in ("a","f") and info1 != info2:
        return 430
    if info1 in ("a","g") and info2 in ("a","g") and info1 != info2:
        return 90
    if info1 in ("a","h") and info2 in ("a","h") and info1 != info2:
        return 150
    if info1 in ("a","i") and info2 in ("a","i") and info1 != info2:
        return 300
    if info1 in ("a","j") and info2 in ("a","j") and info1 != info2:
        return 200
    if info1 in ("a","k") and info2 in ("a","k") and info1 != info2:
        return 325
    if info1 in ("a","m") and info2 in ("a","m") and info1 != info2:
        return 120

    if info1 in ("b","c") and info2 in ("b","c") and info1 != info2:
        return 300
    if info1 in ("b","d") and info2 in ("b","d") and info1 != info2:
        return 300
    if info1 in ("b","e") and info2 in ("b","e") and info1 != info2:
        return 70
    if info1 in ("b","f") and info2 in ("b","f") and info1 != info2:
        return 270
    if info1 in ("b","g") and info2 in ("b","g") and info1 != info2:
        return 170
    if info1 in ("b","h") and info2 in ("b","h") and info1 != info2:
        return 200
    if info1 in ("b","i") and info2 in ("b","i") and info1 != info2:
        return 230
    if info1 in ("b","j") and info2 in ("b","j") and info1 != info2:
        return 270
    if info1 in ("b","k") and info2 in ("b","k") and info1 != info2:
        return 400
    if info1 in ("b","m") and info2 in ("b","m") and info1 != info2:
        return 220


    if info1 in ("c","d") and info2 in ("c","d") and info1 != info2:
        return 120
    if info1 in ("c","e") and info2 in ("c","e") and info1 != info2:
        return 190
    if info1 in ("c","f") and info2 in ("c","f") and info1 != info2:
        return 177
    if info1 in ("c","g") and info2 in ("c","g") and info1 != info2:
        return 180
    if info1 in ("c","h") and info2 in ("c","h") and info1 != info2:
        return 90
    if info1 in ("c","i") and info2 in ("c","i") and info1 != info2:
        return 200
    if info1 in ("c","j") and info2 in ("c","j") and info1 != info2:
        return 170
    if info1 in ("c","k") and info2 in ("c","k") and info1 != info2:
        return 70
    if info1 in ("c","m") and info2 in ("c","m") and info1 != info2:
        return 290

    if info1 in ("d","e") and info2 in ("d","e") and info1 != info2:
        return 220
    if info1 in ("d","f") and info2 in ("d","f") and info1 != info2:
        return 300
    if info1 in ("d","g") and info2 in ("d","g") and info1 != info2:
        return 150
    if info1 in ("d","h") and info2 in ("d","h") and info1 != info2:
        return 60
    if info1 in ("d","i") and info2 in ("d","i") and info1 != info2:
        return 288
    if info1 in ("d","j") and info2 in ("d","j") and info1 != info2:
        return 80
    if info1 in ("d","k") and info2 in ("d","k") and info1 != info2:
        return 125
    if info1 in ("d","m") and info2 in ("d","m") and info1 != info2:
        return 250


    if info1 in ("e","f") and info2 in ("e","f") and info1 != info2:
        return 180
    if info1 in ("e","g") and info2 in ("e","g") and info1 != info2:
        return 160
    if info1 in ("e","h") and info2 in ("e","h") and info1 != info2:
        return 170
    if info1 in ("e","i") and info2 in ("e","i") and info1 != info2:
        return 70
    if info1 in ("e","j") and info2 in ("e","j") and info1 != info2:
        return 300
    if info1 in ("e","k") and info2 in ("e","k") and info1 != info2:
        return 270
    if info1 in ("e","m") and info2 in ("e","m") and info1 != info2:
        return 240


    if info1 in ("f","g") and info2 in ("f","g") and info1 != info2:
        return 310
    if info1 in ("f","h") and info2 in ("f","h") and info1 != info2:
        return 240
    if info1 in ("f","i") and info2 in ("f","i") and info1 != info2:
        return 50
    if info1 in ("f","j") and info2 in ("f","j") and info1 != info2:
        return 350
    if info1 in ("f","k") and info2 in ("f","k") and info1 != info2:
        return 200
    if info1 in ("f","m") and info2 in ("f","m") and info1 != info2:
        return 420


    if info1 in ("g","h") and info2 in ("g","h") and info1 != info2:
        return 50
    if info1 in ("g","i") and info2 in ("g","i") and info1 != info2:
        return 300
    if info1 in ("g","j") and info2 in ("g","j") and info1 != info2:
        return 80
    if info1 in ("g","k") and info2 in ("g","k") and info1 != info2:
        return 250
    if info1 in ("g","m") and info2 in ("g","m") and info1 != info2:
        return 100


    if info1 in ("h","i") and info2 in ("h","i") and info1 != info2:
        return 240
    if info1 in ("h","j") and info2 in ("h","j") and info1 != info2:
        return 90
    if info1 in ("h","k") and info2 in ("h","k") and info1 != info2:
        return 150
    if info1 in ("h","m") and info2 in ("h","m") and info1 != info2:
        return 210

    if info1 in ("i","j") and info2 in ("i","j") and info1 != info2:
        return 350
    if info1 in ("i","k") and info2 in ("i","k") and info1 != info2:
        return 230
    if info1 in ("i","m") and info2 in ("i","m") and info1 != info2:
        return 400


    if info1 in ("j","k") and info2 in ("j","k") and info1 != info2:
        return 200
    if info1 in ("j","m") and info2 in ("j","m") and info1 != info2:
        return 190

    if info1 in ("k","m") and info2 in ("k","m") and info1 != info2:
        return 370

    if info1 in ("n","a") and info2 in ("n","a") and info1 != info2:
        return 380    
    if info1 in ("n","b") and info2 in ("n","b") and info1 != info2:
        return 280
    if info1 in ("n","c") and info2 in ("n","c") and info1 != info2:
        return 230
    if info1 in ("n","d") and info2 in ("n","d") and info1 != info2:
        return 260
    if info1 in ("n","e") and info2 in ("n","e") and info1 != info2:
        return 230
    if info1 in ("n","f") and info2 in ("n","f") and info1 != info2:
        return 430
    if info1 in ("n","g") and info2 in ("n","g") and info1 != info2:
        return 190
    if info1 in ("n","h") and info2 in ("n","h") and info1 != info2:
        return 250
    if info1 in ("n","i") and info2 in ("n","i") and info1 != info2:
        return 300
    if info1 in ("n","j") and info2 in ("n","j") and info1 != info2:
        return 200
    if info1 in ("n","k") and info2 in ("n","k") and info1 != info2:
        return 325
    if info1 in ("n","m") and info2 in ("n","m") and info1 != info2:
        return 320

    if info1 in ("o","a") and info2 in ("o","a") and info1 != info2:
        return 220    
    if info1 in ("o","b") and info2 in ("o","b") and info1 != info2:
        return 212
    if info1 in ("o","c") and info2 in ("o","c") and info1 != info2:
        return 230
    if info1 in ("o","d") and info2 in ("o","d") and info1 != info2:
        return 260
    if info1 in ("o","e") and info2 in ("o","e") and info1 != info2:
        return 200
    if info1 in ("o","f") and info2 in ("o","f") and info1 != info2:
        return 440
    if info1 in ("o","g") and info2 in ("o","g") and info1 != info2:
        return 190
    if info1 in ("o","h") and info2 in ("o","h") and info1 != info2:
        return 190
    if info1 in ("o","i") and info2 in ("o","i") and info1 != info2:
        return 300
    if info1 in ("o","j") and info2 in ("o","j") and info1 != info2:
        return 200
    if info1 in ("o","k") and info2 in ("o","k") and info1 != info2:
        return 325
    if info1 in ("o","m") and info2 in ("o","m") and info1 != info2:
        return 180

    if info1 in ("p","a") and info2 in ("p","a") and info1 != info2:
        return 320    
    if info1 in ("p","b") and info2 in ("p","b") and info1 != info2:
        return 312
    if info1 in ("p","c") and info2 in ("p","c") and info1 != info2:
        return 330
    if info1 in ("p","d") and info2 in ("p","d") and info1 != info2:
        return 290
    if info1 in ("p","e") and info2 in ("p","e") and info1 != info2:
        return 280
    if info1 in ("p","f") and info2 in ("p","f") and info1 != info2:
        return 530
    if info1 in ("p","g") and info2 in ("p","g") and info1 != info2:
        return 390
    if info1 in ("p","h") and info2 in ("p","h") and info1 != info2:
        return 350
    if info1 in ("p","i") and info2 in ("p","i") and info1 != info2:
        return 300
    if info1 in ("p","j") and info2 in ("p","j") and info1 != info2:
        return 200
    if info1 in ("p","k") and info2 in ("p","k") and info1 != info2:
        return 325
    if info1 in ("p","m") and info2 in ("p","m") and info1 != info2:
        return 320

    if info1 in ("n","o") and info2 in ("n","o") and info1 != info2:
        return 30
    if info1 in ("n","p") and info2 in ("n","p") and info1 != info2:
        return 200
     

def stair_time(info1, info2): #계단 오르고 내리는 시간
    if info1[0] == info2[0]:
        return abs((int(info1[1])-int(info2[1]))*21/60)
    else:
        return (int(info1[1])+int(info2[1]))*21/60

alias={}

def set_alias(): #과목변호와 과목이름 매칭
    while True:
        number=input('''과목번호를 입력하시오. 
모시래 본관: 0000
모시래 여동신관: 0001
해오름: 0002 
입력 : ''')

        if len(number) != 4:
            print("유효하지 않은 형식입니다")

        else: 
            set_name=input("부를 이름을 입력하시오: ")
            if set_name in alias.keys():
                print("중복된 값입니다.")

            alias[set_name]=number
            print(f"이제 6번 옵션에서 {set_name}을 사용할 수 있습니다")
            stopornot=input("입력 멈추기: '0'입력, 계속 입력하려면 아무거나 입력하시오 : ")
            if stopornot == "0":
                break

def set_alias_timetable(): # 과목번호 입력, 시간표 만들기?
    alias_timetable = []

    while len(alias_timetable) != 2:
        subject = input("과목이름을 입력하시오. : ")

        if subject in alias_timetable:
            print("중복된 값입니다.")

        if not subject in alias.keys():
            print("과목번호와 매칭되지 않은 이름입니다.")

        else:
            alias_timetable.append(alias[subject])

    return alias_timetable


realspeed=None
all_timtable=None

while True:
    tesk =input('''
원하시는 작업을 입력하시오. 
(0: 종료)
(1: 과목 2개 사이에 딘순 이동시간 측정)
(2: 개인 정보와 속도 정보 초기화 및 입력)
(3: 개인 정보 입력 생략하고 이동속도 입력하기)
(4: 수업명과 과목번호 매칭하기)
(5: 수업명과 과목번호 매칭 초기화)
(6: 매칭된 수업명으로 이동거리 측정하기)
입력 : ''')
    if tesk == "0": #종료
        break
    elif tesk == "1": #과목 2개 사이에 딘순 이동시간 측정
        if realspeed == None:
            personsl_inforo=personsl_information()
            realspeed=speed(personsl_inforo["속도"], personsl_inforo["키"], personsl_inforo["성별"])
        
        timtable=set_2timetable()
        room_info1 = get_info_from_list(timtable[0])
        room_info2 = get_info_from_list(timtable[1])
        if room_info2 == None or room_info1 == None:
            print('정보가 존재하지 않음')
        else:
            room_alpabet=turn_alpadet(room_info1[0], room_info2[0])
            longth=longth_info(room_alpabet[0], room_alpabet[1])
            realstair_time=stair_time(room_info1, room_info2)
            movingtime=float(longth)/float(realspeed)+realstair_time
            print(f"이동 시간은{int(movingtime)}분 입니다.")

    elif tesk == "2":
        personsl_inforo=personsl_information()
        realspeed=speed(personsl_inforo["속도"], personsl_inforo["키"], personsl_inforo["성별"])

    elif tesk == "3": #개인 정보와 속도 정보 초기화 및 입력
        speedinfo=input("이동할 속도를 입력하시오. (단위는 분당 이동거리(m) : m/min)")
        if not speedinfo.isdigit():
            print("유효하지 않은 형식입니다.")
        else:
            realspeed=speedinfo

    elif tesk == "4": #개인 정보 입력 생략하고 이동속도 입력하기
        set_alias()

    elif tesk == "5": #수업명과 과목번호 매칭 초기화
        alias = {}

    elif tesk == "6": #매칭된 수업명으로 이동거리 측정하기
        if alias == {}:
            print("매칭된 수업명이 없습니다.")
        elif len(alias.keys()) < 2:
            print("2개 이상의 수업을 매칭해야합니다.")

        else: 
            if realspeed == None:
                personsl_inforo=personsl_information()
                realspeed=speed(personsl_inforo["속도"], personsl_inforo["키"], personsl_inforo["성별"])
        
            timtable=set_alias_timetable()
            room_info1 = get_info_from_list(timtable[0])
            room_info2 = get_info_from_list(timtable[1])
            if room_info2 == None or room_info1 == None:
                print('정보가 존재하지 않음')
            else:
                room_alpabet=turn_alpadet(room_info1[0], room_info2[0])
                longth=longth_info(room_alpabet[0], room_alpabet[1])
                realstair_time=stair_time(room_info1, room_info2)
                movingtime=float(longth)/float(realspeed)+realstair_time
                print(f"이동 시간은{int(movingtime)}분 입니다.")
        
        

    else:
        print("존재하지 않는 선택지입니다.")