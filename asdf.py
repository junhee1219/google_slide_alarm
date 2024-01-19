from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

# 서비스 계정 키 파일 경로
SERVICE_ACCOUNT_FILE = './asdf.json'
SCOPES = ['https://www.googleapis.com/auth/presentations']

# 서비스 계정 인증 정보 설정
creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
service = build('slides', 'v1', credentials=creds)


def getdict():
    presentation_id = '1Zvt9zjRaN3CK37D25uv9yxWMqQgxAqIfZHpu-PX9EcQ'
    presentation = service.presentations().get(presentationId=presentation_id).execute()

    slides = presentation.get('slides')
    dict = {}

    for i, slide in enumerate(slides):
        text_content = ""
        pageNo = i+1
        page_elements = slide.get('pageElements')
        if page_elements:
            for element in page_elements:    
                try:
                    rgb_color = element['shape']['shapeProperties']['shapeBackgroundFill']['solidFill']['color']['rgbColor']
                    if rgb_color.get('green', 0) == 1 and rgb_color.get('red', 0) == 0 and rgb_color.get('blue', 0) == 0:
                        try :
                            text_content = element['shape']['text']['textElements'][1]['textRun']['content'].strip()
                            if text_content == "배포대기":
                                break
                            elif text_content not in ["성주한","박정우","정강욱","김경희", "박지안", "박상률", "이재성", "이준희", "김민정", "유수민", "강산아", "강지윤", "류성현", "강민아", "정강욱", "김혜경"]:
                                text_content = "확인필요"
                        except :
                            text_content = "미배정"
                        dict[pageNo] = text_content
                    elif rgb_color.get('green', 0) == 0 and rgb_color.get('red', 0) == 0.6 and rgb_color.get('blue', 0) == 1:
                        dict[pageNo] = False
                        try :
                            text_content = element['shape']['text']['textElements'][1]['textRun']['content'].strip()
                        except :
                            text_content = "확인필요"
                        dict[pageNo] = text_content
                except KeyError:
                    continue
        if text_content == "배포대기":
            break
    return dict



import tkinter as tk
from tkinter import messagebox
import threading
import queue
import time
import webbrowser

last_status = 'red'
topmost = False
init = False
def toggle_topmost():
    root.attributes('-topmost', topmost_var.get() == 1)

def change_color_mode(type) :
    if type == "1":
        print(1)
    elif type == "2":
        print(2)

## 색깔 바꾸고 
        

def update_status():
    global last_status  # 이전 상태를 추적하기 위한 전역 변수
    global init
    try:
        # 데이터 큐에서 데이터를 가져옵니다. 논블로킹 방식.
        data = data_queue.get_nowait()
        unassigned_count = list(data.values()).count('미배정')  # 미배정 작업의 수를 계산
        unassigned_label.config(text=f'미배정 작업 건수: {unassigned_count}')  # 미배정 라벨 업데이트
        unassigned_pages = [k for k, v in data.items() if v == '미배정']  # 미배정 슬라이드 번호를 리스트로 가져옴
        unassigned_pages_label.config(text='미배정 슬라이드: ' + ', '.join(map(str, unassigned_pages)))  # 미배정 슬라이드 번호 업데이트

        if '미배정' in data.values():
            status_label.config(bg='red')
            # 이전 상태가 빨강이 아니었다면
            if last_status != 'red':  
                # alert 창 표시
                messagebox.showwarning("경고", "미배정 작업이 있습니다!")  
            last_status = 'red'
        else:
            status_label.config(bg='green')
            last_status = 'green'
        if not init:
            init = True
            status_label.config(text='CS시트 바로가기', width=40, height=2, fg="white")
    except queue.Empty:
        pass
    
    root.after(5000, update_status)  # 5초마다 후에 다시 update_status 함수를 호출
    

def start_update_thread():
    update_thread = threading.Thread(target=update_data_thread)
    update_thread.daemon = True  # 프로그램 종료 시 스레드도 종료되도록 설정
    update_thread.start()

def update_data_thread():
    while True:
        try:
            time.sleep(3)
            data = getdict() 
            data_queue.put(data)  # 데이터 큐에 넣음
        except Exception as e:
            print(e)
def open_website(event):
    webbrowser.open_new("http://qpqp.site")
    
root = tk.Tk()
root.title("대시보드")

data_queue = queue.Queue()  # 데이터를 저장할 큐

status_label = tk.Label(root, text='로딩중입니다...', width=40, height=2, fg="black")
status_label.pack(pady=5)
status_label.bind("<Button-1>", open_website)  # status_label 클릭 이벤트 바인딩

unassigned_label = tk.Label(root, text='', width=50, height=2)  # 미배정 작업 수를 표시할 라벨
unassigned_label.pack(pady=5)

unassigned_pages_label = tk.Label(root, text='', width=50, height=2)  # 미배정 슬라이드 번호를 표시할 라벨
unassigned_pages_label.pack(pady=5)

topmost_var = tk.IntVar()  # 최상단 상태를 추적하는 변수
topmost_check = tk.Checkbutton(root, text="창을 항상 최상단에 놓기", variable=topmost_var, command=toggle_topmost)
topmost_check.pack(pady=5)

update_status()  # 상태 업데이트 시작
start_update_thread()  # 업데이트 스레드 시작

root.mainloop()
