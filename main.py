import win32gui
from pypresence import Presence
import time
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.oauth2.credentials import Credentials
import win32process, psutil

runtime = time.time()
classroom_id = '814092799326421034'
zoom_id = '883221488440864788'
neis_id = '883221355837919292'
gmeet_id = '890564140073107477'
ebs_id = '902544734629814312'
e_id = '908363416203296830'
RPC_Classroom = Presence(classroom_id,pipe=0)
RPC_Neis = Presence(neis_id,pipe=0)
RPC_Zoom = Presence(zoom_id,pipe=0)
RPC_Gmeet = Presence(gmeet_id,pipe=0)
RPC_EBSoc = Presence(ebs_id,pipe=0)
RPC_e = Presence(e_id,pipe=0)
RPC_Classroom.connect()
RPC_Neis.connect()
RPC_Zoom.connect()
RPC_Gmeet.connect()
RPC_EBSoc.connect()
RPC_e.connect()
rpcConnected = ''
rpcConnecteded = ''

usernum = "2"
darkmod = "dark"
browserName = "Chrome"
meetid_open = 1 #Meet 회의의 ID를 RPC에 표시하는 지 여부를 표시. 추후 json 파일에 분리 예정
classlist = ["미술","영어","기가","기술가정","역사","국어","수학","과학","과A","과B","사회","체육"] #추후 Json 파일로 분리 시 함께 분리 예정

coursenum = ''
classnum = ''
courseWork = ''
materialnum = ''

courses = ''

focuswindow_checkingchange = '' #focuswindow 상태가 변했는지 용도를 확인하기 위함
classroom_link = '' #현재 내가 들어가 있는 클래스룸의 링크를 기록
bchange_link = '0' #링크의 변화를 감지하기 위함.
classid = '' #현재 클래스룸 교실의 ID. 세부 수업이나 과제 등을 확인하기 위해 필요
classid_save = '' #RPC를 변경할 때 클래스룸 교실의 ID의 변경 여부를 감지하기 위해 필요
closeRPC = 0
changeid = 0
workStatus = ""
classroom_info = '' #교과의 이름
classroom_change = '' #classroom_info 변수가 바뀔 시 RPC 정보를 새로고침 할 수 있게 하기 위함.
etcClassName = '' #기타교과의 이름름 기록
etcClassnameBackup = '' #etcClassName 변수가 바뀔 시 RPC 정보를 새로고침 할 수 있게 하기 위함.
e_menu = '' #e학습터 지역을 구분하는 용도

neisinfo = '' #현재 나이스에서 접속한 교육청의 지역과 서비스 종류(학생, 학부모 등)을 기록
neis_status = '' #현재 나이스에서 접속한 서비스의 이름을 기록
neisServiceSave = ''
neis_link = ''

meet_id = '' #Meet 회의의 ID를 기록, 설정에서 표시 여부 설정가능
meet_idSave = ''

def classinfo_find() :
    global classroom_info
    global imagefile
    if "미술" in website[0] and classroom_info != "미술" :
        classroom_info = "미술"
        imagefile = "art_icon_"+darkmod
    if "영어" in website[0] and classroom_info != "영어" :
        classroom_info = "영어"
        imagefile = "english_icon_"+darkmod
    if "기가" in website[0] or "기술가정" in website[0] and classroom_info != "기술가정" :
        classroom_info = "기술가정"
        imagefile = "giga_icon_"+darkmod
    if "역사" in website[0] and classroom_info != "역사" :
        classroom_info = "역사"
        imagefile = "history_icon_"+darkmod
    if "국어" in website[0] :
        if "중국어" in website[0] and classroom_info != "중국어" :
            classroom_info = "중국어"
            imagefile = "chinese_icon_"+darkmod
        elif "국어" in website[0] and "중국어" not in website[0] and classroom_info != "국어" :
            classroom_info = "국어"
            imagefile = "korean_icon_"+darkmod
    if "수학" in website[0] and classroom_info != "수학" :
        classroom_info = "수학"
        imagefile = "mathematic_icon_"+darkmod
    if "체육" in website[0] and classroom_info != "체육" :
        classroom_info = "체육"
        imagefile = "pe_icon_"+darkmod
    if "과A" in website[0] or "과학A" in website[0] and classroom_info != "과학A" :    
        classroom_info = "과학A"
        #classid = courses[classnum]['id']
        #classid_save = classid
        imagefile = "science_icon_"+darkmod
    if "과B" in website[0] or "과학B" in website[0] and classroom_info != "과학B" :
        classroom_info = "과학B"
        #classid = courses[classnum]['id']
        #classid_save = classid
        imagefile = "science_icon_"+darkmod
    if "과학" in website[0] and "과학A" not in website[0] and "과학B" not in website[0] and "과A" not in website[0] and "과B" not in website[0] :
        classroom_info = "과학"
        #classid = courses[classnum]['id']
        #classid_save = classid
        imagefile = "science_icon_"+darkmod
    if "사회" in website[0] and classroom_info != "사회" :
        classroom_info = "사회"
        imagefile = "social_icon_"+darkmod

def RPCclear() :
    global rpcConnected
    global RPC_Classroom
    global RPC_Neis
    global RPC_Gmeet
    global RPC_Zoom
    global runtime
    global RPC_EBSoc
    global RPC_e
    print(rpcConnected)
    if rpcConnected == 'Classroom' :
        RPC_Classroom.clear()
        print("Google Classroom Rich Presence has cleared.")
    if rpcConnected == 'EBSoc' :
        RPC_EBSoc.clear()
        print("EBS Online Class Rich Presence has cleared.")    
    if rpcConnected == 'Neis' :
        RPC_Neis.clear()
        print("Neis Rich Presence has cleared.")
    if rpcConnected == 'Zoom' :
        RPC_Neis.clear()
        print("Zoom Rich Presence has cleared.")
    if rpcConnected == 'Gmeet' :
        RPC_Gmeet.clear()
        print("Google Meet Rich Presence has cleared.")
    if rpcConnected == 'e학습터' :
        RPC_e.clear()
        print("e학습터 Rich Presence has cleared.")
    runtime = time.time()

while courses == '' :
    SCOPES=[u"https://www.googleapis.com/auth/classroom.student-submissions.students.readonly", u"https://www.googleapis.com/auth/classroom.student-submissions.me.readonly", u"https://www.googleapis.com/auth/classroom.courses.readonly", u"https://www.googleapis.com/auth/classroom.courseworkmaterials.readonly"]
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    try :
        if os.path.exists('./credentials/token.json'):
            creds = Credentials.from_authorized_user_file('./credentials/token.json', SCOPES)
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                os.remove("./credentials/token.json")
                print("Token has expired or needed to create it again. Please re-authorize.")
                flow = InstalledAppFlow.from_client_secrets_file(
                    './credentials/credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    './credentials/credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open('./credentials/token.json', 'w') as token:
                token.write(creds.to_json())
    except :
        print("Error occured while creating, checking json file. Please restart to authorize.")
        try : 
            os.remove('./credentials/token.json')
        except :
            exit()
        exit()

    service = build('classroom', 'v1', credentials=creds)
    courseresults = service.courses().list().execute()
    courses = courseresults.get('courses', [])

if not courses:
    print('No courses found.')
else:
    print("Classroom courses loaded successfully!")
    for course in courses:
        for i in range(256) :
            try :
                globals()['Courses{}'.format(i)] = courses[i]
            except IndexError:
                globals()['Courses{}'.format(i)] = 'null'
                continue

while True:  
    time.sleep(0.1)
    try :
        handle = win32gui.GetForegroundWindow()
        tid, pid = win32process.GetWindowThreadProcessId(handle)
        process_name = psutil.Process(pid).name()
    except :
        print ("An Unknown Exception had occured.")
    focuswindow = win32gui.GetWindowText(win32gui.GetForegroundWindow())
    if len(focuswindow) >= 128:
        focuswindow = focuswindow[:len(focuswindow) - 3] + "..."
    website = focuswindow.split(' - ')
    neisService = focuswindow.replace(f" - {browserName}","")
    neisService = neisService.split('<')
    classcoursework = focuswindow.replace(f" - {browserName}","")

    if "https://classroom.google.com" in website[0]:
        if classroom_link != website[0] :
            classroom_link = website[0]
            neis_link = ''
            if bchange_link != classroom_link :
                print(classroom_link)
    if "https://neis.go.kr" in website[0]:
        if neis_link != website[0] :
            neis_link = website[0]
            classroom_link = ''
            if bchange_link != neis_link :
                print(neis_link)
    if "https://stu." in website[0] and ".go.kr" in website[0]:
        if neis_link != website[0] :
            neis_link = website[0]
            classroom_link = ''
            if bchange_link != neis_link :
                print(neis_link)

    for i in range(256) :
        try : 
            if globals()['Courses{}'.format(i)] != 'null' and courses[i]['alternateLink'].replace("/c/","/u/"+usernum+"/c/") == classroom_link :
                classnum = i
        except IndexError :
            i = 'null'
        if i == 'null' :
            continue

    if classnum != '' and courses[classnum]['alternateLink'].replace("/c/","/u/"+usernum+"/c/") == classroom_link :
        classid = courses[classnum]['id']
        classid_save = classid
        if classroom_change == classroom_info :
            i = -1
            for i in range(256) :
                try :
                    if courses[i]['alternateLink'].replace("/c/","/u/"+usernum+"/c/") == website[0] :
                        a = 0
                        for a in range(256) :
                            try :
                                if classlist[a] not in courses[i]['name'] :
                                    etcClassName = 'InProgress'
                                if classlist[a] in courses[i]['name'] :
                                    etcClassName = ''
                                    break
                            except :
                                if etcClassName == 'InProgress' :
                                    etcClassName = courses[i]['name']
                                    classroom_info = "기타"
                                    imagefile = f"etc_icon_{darkmod}"
                                continue
                    i = i+1
                except :
                    continue

        time.sleep(1.5)
        focuswindow = win32gui.GetWindowText(win32gui.GetForegroundWindow())
        if len(focuswindow) >= 128:
            focuswindow = focuswindow[:len(focuswindow) - 3] + "..."
        website = focuswindow.split(' - ')
        classinfo_find()

        if etcClassName != etcClassnameBackup or classroom_change != classroom_info or classroom_link != courses[classnum]['alternateLink'].replace("/c/","/u/"+usernum+"/c/") and 'https://classroom.google.com' in classroom_link :
            workStatus = ''
            bchange_link = classroom_link
            classroom_link = ''
            classroom_change = classroom_info
            if rpcConnected != 'Classroom' :
                RPCclear()
                rpcConnected = 'Classroom'
                closeRPC = 0
                print(rpcConnected)
            if rpcConnected == 'Classroom' :
                if classroom_info != "기타" :
                    RPC_Classroom.update(details="교과 : "+classroom_info,
                        state="수업을 선택하는 중",
                        large_image=imagefile,
                        small_image="classroom_icon_circle",
                        large_text="Google Classroom",
                        start=runtime)
                if classroom_info == "기타" :
                    if etcClassName != '' :
                        etcClassnameBackup = etcClassName
                        RPC_Classroom.update(details="교과 : "+etcClassName,
                            state="수업을 선택하는 중",
                            large_image=imagefile,
                            small_image="classroom_icon_circle",
                            large_text="Google Classroom",
                            start=runtime)

    if focuswindow_checkingchange != focuswindow:
        focuswindow_checkingchange = focuswindow
        print(classcoursework)

    if  classroom_link == "https://classroom.google.com/u/"+usernum+"/h" :
        classroom_info = '홈'
        workStatus = ''
        bchange_link = classroom_link
        classroom_link = ''
        classroom_change = classroom_info
        if rpcConnected != 'Classroom' :
            rpcConnected = 'Classroom'
            RPCclear()
            print(rpcConnected)
            closeRPC = 0
        if rpcConnected == 'Classroom' :
            RPC_Classroom.update(details="메인 화면",
                state="교실을 선택하는 중",
                large_image="classroom_icon",
                large_text="Google Classroom",
                buttons=[{"label":"사이트에 접속하기", "url":"https://classroom.google.com"}],
                start=runtime)

    if classid != '' or classid != classid_save :
        if changeid != classid :
            changeid = classid
            print(classid)

            materialresults = service.courses().courseWorkMaterials().list(courseId=classid).execute()
            courseWorkMaterials = materialresults.get('courseWorkMaterial', [])
            if not courseWorkMaterials :
                print ("No Materials found.")
            else :
                for courseWorkMaterial in courseWorkMaterials :
                    for i in range(256) :
                        try :
                            globals()['courseWorkMaterial{}'.format(i)] = courseWorkMaterials[i]
                        except IndexError :
                            globals()['courseWorkMaterial{}'.format(i)] = 'null'
                            continue

            workresults = service.courses().courseWork().list(courseId=classid).execute()
            courseWorks = workresults.get('courseWork', [])
            print("CourseWork imported in this Program!")
            if not courseWorks:
                print('No courseWorks found.')
            else :
                for CourseWork in courseWorks:
                    for i in range(256) :
                        try :
                            globals()['courseWork{}'.format(i)] = courseWorks[i]
                        except IndexError :
                            globals()['courseWork{}'.format(i)] = 'null'
                            continue

        for i in range(256) :
            try :
                if globals()['courseWork{}'.format(i)] != 'null' and classroom_link == courseWorks[i]['alternateLink'].replace("/c/","/u/"+usernum+"/c/") :
                    coursenum = i
            except:
                i = 'null'
                continue
        
        for i in range(256) :
            try :
                if globals()['courseWorkMaterial{}'.format(i)] != 'null' and classroom_link == courseWorkMaterials[i]['alternateLink'].replace("/c/","/u/"+usernum+"/c/") :
                    materialnum = i
            except :
                i = 'null'
                continue
        try : 
            if coursenum != '' and classroom_link == courseWorks[coursenum]['alternateLink'].replace("/c/","/u/"+usernum+"/c/"):
                bchange_link = classroom_link
                classroom_link = ''
                if courseWorks[coursenum]['workType'] == "MULTIPLE_CHOICE_QUESTION" :
                    workStatus = " (객관식 질문)"
                elif courseWorks[coursenum]['workType'] == "SHORT_ANSWER_QUESTION" :
                    workStatus = " (단답형 질문)"
                elif courseWorks[coursenum]['workType'] == "ASSIGNMENT" :
                    workStatus = " (과제)"
                #classroom_info = courseWorks[coursenum]['title']
                if rpcConnected != 'Classroom' :
                    RPCclear()
                    rpcConnected = 'Classroom'
                    closeRPC = 0
                if rpcConnected == 'Classroom' :
                    if classroom_info != "기타" :
                        RPC_Classroom.update(details="교과 : "+classroom_info,
                            state="수업 : "+courseWorks[coursenum]['title']+workStatus,
                            #state="수업을 선택하는 중",
                            large_image=imagefile,
                            small_image="classroom_icon_circle",
                            start=runtime)
                    if classroom_info == "기타" :
                        RPC_Classroom.update(details="교과 : "+etcClassName,
                            state="수업 : "+courseWorks[coursenum]['title']+workStatus,
                            #state="수업을 선택하는 중",
                            large_image=imagefile,
                            small_image="classroom_icon_circle",
                            start=runtime)
        
            if materialnum != '' and classroom_link == courseWorkMaterials[materialnum]['alternateLink'].replace("/c/","/u/"+usernum+"/c/") :
                workStatus = " (자료)"
                bchange_link = classroom_link
                classroom_link = ''
                #classroom_info = courseWorkMaterials[materialnum]['title']
                if rpcConnected != 'Classroom' :
                    rpcConnected = 'Classroom'
                if rpcConnected == 'Classroom' :
                    if classroom_info != "기타" :
                        RPC_Classroom.update(details="교과 : "+classroom_info,
                            state="수업 : "+courseWorkMaterials[materialnum]['title']+workStatus,
                            large_image=imagefile,
                            small_image="classroom_icon_circle",
                            start=runtime)
                    if classroom_info == "기타" :
                        RPC_Classroom.update(details="교과 : "+etcClassName,
                            state="수업 : "+courseWorkMaterials[materialnum]['title']+workStatus,
                            large_image=imagefile,
                            small_image="classroom_icon_circle",
                            start=runtime)
        except IndexError :
            continue
    
    if "교육청 나이스" in website[0] and "서비스" in website[0] :
        neisinfo = website[0].replace(" 나이스","")
        if rpcConnected != 'Neis' :
            RPCclear()
            rpcConnected = 'Neis'
            closeRPC = 0
        if rpcConnected == 'Neis' :
            RPC_Neis.update(details=neisinfo,
                state="서비스를 선택하는 중",
                #large_image=imagefile,
                #small_image="classroom_icon_circle",
                large_image=f"neis_icon_{darkmod}",
                small_image=f"logo_icon_{darkmod}",
                start=runtime)
    if "나이스" in website[0] and "서비스 로그인" in website[0] :
        neis_status = "서비스를 선택하는 중"
    if neisService != neisServiceSave :
        print(neisService)
        neisServiceSave = neisService
        try :
            neis_status = neisService[1]+" | "+neisService[0]
            print(neis_status)
            if rpcConnected != 'Neis' :
                RPCclear()
                rpcConnected = 'Neis'
                closeRPC = 0
        except :
            pass
        if rpcConnected == 'Neis' :
                try :
                    RPC_Neis.update(details=neisinfo,
                        state=neis_status,
                        large_image=f"neis_icon_{darkmod}",
                        small_image=f"logo_icon_{darkmod}",
                        start=runtime)
                except :
                    pass
    
    if website[0] == "나이스 대국민서비스" :
        if neisinfo != '나이스 대국민서비스' :
            neisinfo = '나이스 대국민서비스'
            neis_status = "서비스를 선택하는 중"
            print("나이스 대국민서비스 작동 중")
        if rpcConnected != 'Neis' :
            RPCclear()
            rpcConnected = 'Neis'
            closeRPC = 0
        if rpcConnected == 'Neis' :
            RPC_Neis.update(details="나이스 대국민서비스",
                state="서비스를 선택하는 중",
                large_image=f"neis_icon_{darkmod}",
                small_image=f"logo_icon_{darkmod}",
                start=runtime)

    if "Meet - " in focuswindow :
        if website[0] == 'Google Meet' :
            meet_id = ''
            if meet_id != meet_idSave :
                meet_idSave = meet_id
            if rpcConnected != 'Gmeet' :
                RPCclear()
                rpcConnected = 'Gmeet'
                closeRPC = 0
            if rpcConnected == 'Gmeet' :
                RPC_Gmeet.update(details="메인 화면",
                        state="회의를 선택하는 중",
                        large_image=f"gmeet_icon_{darkmod}",
                        small_image=f"logo_icon_{darkmod}",
                        start=runtime)
        else :
            meet_id = website[1]
            if meet_id != meet_idSave :
                meet_idSave = meet_id
                if rpcConnected != 'Gmeet' :
                    RPCclear()
                    rpcConnected = 'Gmeet'
                    closeRPC = 0
                if rpcConnected == 'Gmeet' :
                    if meetid_open == 1 :
                        RPC_Gmeet.update(details="Google Meet 회의 중",
                        state=f"회의 ID : {meet_id}",
                        large_image=f"gmeet_icon_{darkmod}",
                        small_image=f"logo_icon_{darkmod}",
                        start=runtime)
                    if classroom_info != '' :
                        RPC_Gmeet.update(details="Google Meet 회의 중",
                        state=f"교과 : {classroom_info}",
                        large_image=f"gmeet_icon_{darkmod}",
                        small_image=f"logo_icon_{darkmod}",
                        start=runtime)
                    if meetid_open == 0 :
                        RPC_Gmeet.update(details="Google Meet 회의 중",
                        state=f"회의 ID : [비공개]",
                        large_image=f"gmeet_icon_{darkmod}",
                        small_image=f"logo_icon_{darkmod}",
                        start=runtime)

    if "Zoom.exe" == process_name :
        if "Zoom" == focuswindow : 
            if rpcConnected != 'Zoom' :
                RPCclear()
                rpcConnected = 'Zoom'
                closeRPC = 0
            if rpcConnected == 'Zoom' :
                RPC_Zoom.update(details="메인 화면",
                state="회의에 접속하는 중",
                large_image="zoom_icon",
                small_image=f"logo_icon_{darkmod}",
                start=runtime)
        if "Zoom 회의" == focuswindow or "Zoom Meetings" == focuswindow : 
            if rpcConnected != 'Zoom' :
                RPCclear()
                rpcConnected = 'Zoom'
                closeRPC = 0
            if rpcConnected == 'Zoom' :
                RPC_Zoom.update(details="회의에 참여하는 중",
                large_image="zoom_icon",
                small_image=f"logo_icon_{darkmod}",
                start=runtime)

    if "EBS 온라인클래스" == website[0] :
        if rpcConnected != 'EBSoc' :
            RPCclear()
            rpcConnected = 'EBSoc'
            closeRPC = 0
        if rpcConnected == 'EBSoc' :
            RPC_EBSoc.update(details="온라인 수업에 참여하는 중",
            large_image=f"ebsoc_icon_{darkmod}",
            small_image=f"logo_icon_{darkmod}",
            start=runtime)
    
    if "e학습터" in website[0] :
        if "e학습터" == website[0] :
            e_menu = '메인 화면'
        if "서울 e학습터" == website[0] :
            e_menu = '서울 지역 초/중학교'
        if "부산, 울산 e학습터" == website[0] :
            e_menu = '부산, 울산 지역 초/중학교'
        if "대구, 강원 e학습터" == website[0] :
            e_menu = '대구, 강원도 지역 초/중학교'
        if "인천, 충북 e학습터" == website[0] :
            e_menu = '인천, 충청북도 지역 초/중학교'
        if "광주, 대전 e학습터" == website[0] :
            e_menu = '광주, 대전 지역 초/중학교'
        if "세종, 충남 e학습터" == website[0] :
            e_menu = '세종, 충청남도 지역 초/중학교'
        if "경기 e학습터" == website[0] :
            e_menu = '경기도 지역 초/중학교'
        if "전북, 전남 e학습터" == website[0] :
            e_menu = '전라북도, 전라남도 지역 초/중학교'
        if "경북, 제주 e학습터" == website[0] :
            e_menu = '경상북도, 제주 지역 초/중학교'
        if rpcConnected != 'e학습터' :
            RPCclear()
            rpcConnected = 'e학습터'
            closeRPC = 0
        if rpcConnected == 'e학습터' :
            RPC_e.update(details=e_menu,
            state="온라인 수업에 참여하는 중",
            large_image=f"e_icon_{darkmod}",
            small_image=f"logo_icon_{darkmod}",
            start=runtime)

    else :
        closeRPC = closeRPC+1
        if closeRPC == 3000 :
            RPCclear()
            classroom_info = ''
            etcClassName = ''
            neisinfo = ''
            neis_status = ''
            print("Presence has disabled due to inactivity.")