
import tkinter as tk
import serial.tools.list_ports
from tkinter import scrolledtext, ttk, font, Label, StringVar,simpledialog, messagebox,Tk
import serial
import threading
import time
import datetime
import winsound
import os
import sys
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

# 시리얼 포트 및 보레이트 설정
SERIAL_PORT = "COM3"  # 기본 시리얼 포트
DEFAULT_BAUDRATE = "1200"  # 기본 보레이트

# 데이터 전송 간격 (초)
SEND_INTERVAL = 1

# 리튬배터리 [V]
LITUM37 = 37

# 소켓 및 스레드 초기화
ser = None
send_thread = None
receive_thread = None
stop_signal = threading.Event()
send_count = 0
receive_count = 0
global rx_status, status_message_label
rx_status = "READY"
rx_status_lock = threading.Lock()  # 스레드 간 동기화용 Lock




# 구글 시트 ID와 범위 정의
SPREADSHEET_ID = '1ayeNf-dKmB8CmY2Tqzyc4sVx5hPzdsCkioKaOnxH8EU'  # 구글 시트 ID
RANGE_NAME = 'LicenseList!A2:A'  # A열의 범위

# JSON 파일 경로 설정
json_file_path = r"C:\제품등록\gcp9304-4410543fedf2.json"

# 인증 정보 설정
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
creds = Credentials.from_service_account_file(json_file_path, scopes=SCOPES)

# 구글 시트 서비스 객체 생성
service = build('sheets', 'v4', credentials=creds)

# 시트 데이터를 가져오는 함수
def get_emails_from_sheet():
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME).execute()
    values = result.get('values', [])

    # 이메일 리스트 생성
    emails = [row[0] for row in values if row]  # 빈 행 제외
    return emails


class PacketDefinition:
    def __init__(self, start_bytes, end_bytes, length):
        self.start_bytes = start_bytes
        self.end_bytes = end_bytes
        self.length = length

# Global packet definitions
REQ_METER_PACKET = PacketDefinition(start_bytes=bytes([0x10]), end_bytes=bytes([0x16]), length=5)  # 5바이트
ACK_METER_PACKET = PacketDefinition(start_bytes=bytes([0x68]), end_bytes=bytes([0x16]), length=21)  # 21바이트
TCP_PACKET = PacketDefinition(start_bytes=bytes([]), end_bytes=bytes([]), length=0)  # 길이 미정의




# 이메일 주소 확인 함수
def check_email():
    # Tkinter 루트 창 생성
    root = Tk()
    root.withdraw()  # 루트 창 숨기기

    # 사용자 홈 디렉토리에 이메일 인증 여부 파일 저장 경로 설정
    home_dir = os.path.expanduser("~")
    file_path = os.path.join(home_dir, "email_verified.txt")

    # 파일 경로 출력 (파일이 어디에 저장되는지 확인)
    print(f"파일이 저장될 경로: {file_path}")

    # 이메일 인증이 이미 완료되었는지 확인
    if os.path.exists(file_path):
        print("이메일이 이미 인증되었습니다. 프로그램을 계속 실행합니다.")
        root.destroy()  # 루트 창 닫기
        return
    
    # 이메일 입력 받기
    user_email = simpledialog.askstring("이메일 입력", "이메일 주소를 입력하세요:", parent=root)

    # 구글 시트에서 이메일 리스트 가져오기
    emails = get_emails_from_sheet()

    # 이메일 확인
    if user_email in emails:
        print("이메일이 확인되었습니다. 프로그램을 계속 실행합니다.")
        # 이메일 인증 성공을 기록하는 파일 생성
        with open(file_path, "w") as f:
            f.write("verified")
        root.destroy()  # 루트 창 닫기
    else:
        print("제함됨!")
        # 사용자에게 알림창 표시
        messagebox.showinfo(
            "이메일 인증 실패", 
            "관리자(eusy1327@istec.co.kr)에게 본인의 Gmail 계정과 함께 사용 신청하세요!",
            parent=root
        )
        root.destroy()  # 루트 창 닫기
        sys.exit()
        exit_program()



# 연결 시작 함수
def start_connection():
    global ser, send_thread, receive_thread, stop_signal, send_count, receive_count
    port = port_var.get()
    baudrate = baudrate_var.get()
    
    try:        
        ser = serial.Serial(port, baudrate)  # 시리얼 포트 열기
        stop_signal.clear()

        # 현재 선택된 모드 가져오기
        handle_mode_selection = current_mode.get()

        match handle_mode_selection:
            case "계량기조회":
                # 계량기 조회 모드일 때 동작
                send_text.insert(tk.END, f"[계량기 조회모드]\n")        
            case "계량기응답":
                # 계량기 응답 모드일 때 동작
                send_text.insert(tk.END, f"[검침기 검침요청 패킷]\n")        
            case _:
                # 기본 동작 (예외 처리 또는 로그 출력)
                send_text.insert(tk.END, "[알 수 없는 모드입니다]")

        send_thread = threading.Thread(target=send)
        receive_thread = threading.Thread(target=receive)
        send_thread.start()
        receive_thread.start()
        start_button.config(state=tk.DISABLED)
        stop_button.config(state=tk.NORMAL)
        exit_button.config(state=tk.NORMAL)
        status_label.config(text="연결 중")        
        update_title()
        
    except Exception as e:
        status_label.config(text=f"연결 오류: {str(e)}")




# 데이터 전송 함수
def send():
    global send_count
    first_send = True  # 처음 전송 여부를 나타내는 변수

    # 현재 선택된 모드 가져오기
    handle_mode_selection = current_mode.get()

    while not stop_signal.is_set():
        try:
            # [스레드간 전역변수 사용법]
            global rx_status
            with rx_status_lock:  # Lock을 사용해 전역 변수 접근 보호
                rx_status = "READY"
            update_status_message()

            # select_mode에 따라 전송할 데이터 설정
            dataToSend = None
            match handle_mode_selection:
                case "계량기조회":  # (기본값)
                    dataToSend = REQ_METER_PACKET.start_bytes + bytes([0x5B, 0x01, 0x5C]) + REQ_METER_PACKET.end_bytes + bytes([0x0d, 0x0a])  # 5 bytes 검침 요청 데이터

                case "계량기응답":
                        # 나중에 구현할 내용
                        pass

                case "TCP":  # TCP mode
                    dataToSend = bytes([0x11, 0x22, 0x33, 0x44, 0x55])  # 예시 데이터

                case _:  # 기본값을 처리                   
                    handle_mode_selection("계량기조회")
                    send_text.insert(tk.END, f"[알수없는 모드!]\n")

            if not first_send:
                send_text.insert(tk.END, f"검침요청> {send_count}  ")
                send_count += 1  # 첫 번째 전송이 아닌 경우에만 send_count 증가
                # latin-1로 데이터 디코딩
                send_text.insert(tk.END, f"{dataToSend.decode('latin-1')}")
                send_text.see(tk.END)  # 자동 스크롤
                send_label.config(text=f"TX({send_count})")

            # 데이터 전송
            if handle_mode_selection == "계량기조회":
                ser.write(dataToSend)
                first_send = False  # 첫 번째 전송 후에는 False로 변경
                time.sleep(SEND_INTERVAL)
            else:
                status_label.config(text="검침기의 조회명령 대기중...")
                break

        except Exception as e:
            status_label.config(text=f"데이터 전송 오류: {str(e)}")
            break




def receive():
    global receive_count, rx_status
    packet = b''  # 패킷 초기화
    VIF = None  # VIF 초기화
    meter_value_index = 0  # 검침값 인덱스 초기화
    
    # 현재 선택된 모드 가져오기
    selected_mode = current_mode.get()

    while not stop_signal.is_set():
        try:
            data = ser.read(1)  # 1바이트씩 읽음
            if data:
                match selected_mode:
                    case "계량기조회":  # Serial mode (기본값)
                        if data == ACK_METER_PACKET.end_bytes and len(packet) >= ACK_METER_PACKET.length:
                            current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                            # [스레드간 전역변수 사용법]
                            with rx_status_lock:  # Lock을 사용해 전역 변수 접근 보호
                                rx_status = "OK"

                            # 비프음 추가 (200ms 동안)
                            winsound.Beep(1000, 200)  # Frequency: 1000Hz, Duration: 200ms
                            
                            update_status_message()
                            print(f"{current_time} : {len(packet)} {packet.hex()}")

                            if len(packet) <= 21:  # 패킷의 길이가 21바이트인 경우 정상
                                receive_count += 1

                                # Status 변수에 packet[13] 값 저장-시작에 0x00 이 붙어있음 주의!
                                Status = 0b00000000  # packet[13]

                                # Batt 변수에 Status의 하위 5bits 저장
                                Batt = (LITUM37 - (Status & 0b00011111)) * 0.1

                                # 검침값 ,기물번호, 체크섬, VIF 추출 및 소수점 위치 결정
                                if len(packet) == 21:  # 신동아 15mm 가 아닌 모든 수도미터인 경우
                                    meter_value_hex = f"{packet[19]:02x}{packet[18]:02x}{packet[17]:02x}{packet[16]:02x}"
                                    VIF = packet[15] & 0x0F
                                    meter_num_hex = f"{packet[12]:02x}{packet[11]:02x}{packet[10]:02x}{packet[9]:02x}"
                                    check_sum = sum(packet[18:4:-1])

                                if len(packet) == 20:  # 신동아 15mm 인 경우 1byte 차이가 있음!
                                    meter_value_hex = f"{packet[18]:02x}{packet[17]:02x}{packet[16]:02x}{packet[15]:02x}"
                                    VIF = packet[14] & 0x0F
                                    meter_num_hex = f"{packet[11]:02x}{packet[10]:02x}{packet[9]:02x}{packet[8]:02x}"
                                    check_sum = sum(packet[17:3:-1])

                                check_sum &= 0xFF  # 하위 1바이트만 사용
                                check_sum = check_sum.to_bytes(1, 'big').hex()
                                meter_value_int = int(meter_value_hex, 10)

                                if check_sum == f"{packet[-1]:02x}":  # 계산된 체크섬 Vs 배열의 체크섬 비교
                                    pass
                                else:
                                    print("CheckSum ERROR !!!!!!!!!!!!!!")

                                # 실수형 검침값 계산
                                if VIF == 3:
                                    meter_value = meter_value_int / 1000  # 소수점 위치가 3일 때
                                elif VIF == 2:
                                    meter_value = meter_value_int / 100  # 소수점 위치가 2일 때
                                elif VIF == 1:
                                    meter_value = meter_value_int / 10  # 소수점 위치가 1일 때
                                else:
                                    meter_value = meter_value_int

                                # 현재 시간 표시
                                current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                # 소수점 위치 계산
                                decimal_places = 8 - len(str(int(meter_value)))
                                if decimal_places < 0:
                                    decimal_places = 0

                                # 검침값 출력
                                receive_text.insert(tk.END, f"\r{current_time} : {meter_value_index % 1000:03d} : {meter_value} : {meter_num_hex}\n")

                                meter_value_index = (meter_value_index + 1) % 1000  # 다음 검침값 인덱스 계산
                                receive_text.see(tk.END)  # 자동 스크롤
                                receive_label.config(text=f"RX({receive_count})")
                                packet = b''  # 패킷 초기화
                            else:
                                packet = b''  # 패킷 초기화
                        else:
                            packet += data

                            
                    case "계량기응답":
                        if data == REQ_METER_PACKET.end_bytes and len(packet) >= REQ_METER_PACKET.length:
                            current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                            # [스레드간 전역변수 사용법]
                            with rx_status_lock:  # Lock을 사용해 전역 변수 접근 보호
                                rx_status = "OK"
                            # 비프음 추가 (1000ms 동안)
                            winsound.Beep(1000, 1000)  # Frequency: 1000Hz, Duration: 1000ms
                            update_status_message()
                            print(f"{current_time} : {len(packet)} {packet.hex()}")
                            send_text.insert(tk.END, f"{packet.hex()}\n")  # Output the hex value of the packet
                            receive_count += 1                               
                            receive_text.see(tk.END)  # 자동 스크롤
                            receive_label.config(text=f"RX({receive_count})")
                            # 패킷 초기화
                            packet = b''
                            # 패킷 전송 (winsound.Beep 후 전송)
                            dataToSend = (
                                ACK_METER_PACKET.start_bytes + bytes([
                                    0x0F, 0x0F, 0x68, 0x08, 0x01, 0x78, 0x0F, 0x97, 0x10, 0x11, 0x23, 0x00,
                                    0x1C, 0x13, 0x17, 0x02, 0x00, 0x00, 0xB3]) + ACK_METER_PACKET.end_bytes
                            )  # 21 bytes 검침 응답 데이터: 0.217 톤
                            ser.write(dataToSend)  # 패킷 전송
                            send_text.insert(tk.END, f"응답 패킷 전송: {dataToSend.hex()}\n")
                            send_text.see(tk.END)  # 자동 스크롤
                        else:
                            packet += data
                        # 3초 대기
                        time.sleep(3)  # Delay for 3 seconds
                        with rx_status_lock:  # Lock을 사용해 전역 변수 접근 보호
                            rx_status = "READY"
                        update_status_message()


                        
                    case "TCP":  # TCP mode (미구현)
                        # 나중에 구현할 내용
                        pass

                    case _:  # 기본값을 처리
                        # 기본값 처리
                        pass

        except Exception as e:
            status_label.config(text=f"데이터 수신 오류: {str(e)}")




# 프로그램 멈춤 함수
def stop_connection():
    global ser, send_thread, receive_thread
    try:
        stop_signal.set()  # 데이터 전송 및 수신 스레드 종료 플래그 설정
        if send_thread is not None:
            send_thread.join()  # 데이터 전송 스레드 종료 대기
##        if receive_thread is not None:
##            receive_thread.join()  # 데이터 수신 스레드 를 주석처리해야 먹통방지된다
        if ser is not None and ser.is_open:
            ser.close()  # 시리얼 포트 닫기
        # [스레드간 전역변수 사용법]
        global rx_status
        with rx_status_lock:  # Lock을 사용해 전역 변수 접근 보호
             rx_status = "READY"
        update_status_message()
        start_button.config(state=tk.NORMAL)
        stop_button.config(state=tk.DISABLED)
        exit_button.config(state=tk.NORMAL)
        status_label.config(text="연결 중지됨")
    except Exception as e:
        status_label.config(text=f"연결 종료 오류: {str(e)}")

# 화면 지우기 함수
def clear_screen():
    global send_count, receive_count  # 전역 변수임을 명시

    send_text.delete(1.0, tk.END)
    receive_text.delete(1.0, tk.END)
    send_count = 0  # global 키워드 추가
    receive_count = 0  # global 키워드 추가
    send_label.config(text=f"TX({send_count})")
    receive_label.config(text=f"RX({receive_count})")

        
# 프로그램 종료 함수
def exit_program():
    if send_thread is not None:
        stop_connection()
    root.quit()


# Tkinter 창 생성
root = tk.Tk()
port_var = StringVar()
PROGRAM_VERSION = "2024-10-11"
root.title(f"유무선시험기 {PROGRAM_VERSION} ")

# 현재 모드 상태를 표시할 레이블
current_mode = tk.StringVar(value="계량기조회")
mode_label = tk.Label(root, textvariable=current_mode, font=("Arial", 24, "bold"))
mode_label.pack(pady=10)

# 메뉴 바 추가
menubar = tk.Menu(root)

# 모드 메뉴
file_menu = tk.Menu(menubar, tearoff=0)

# 모드 선택 핸들러
def handle_mode_selection(mode):  
    current_mode.set(mode)

# 메뉴에서 모드 선택 시 사용하는 코드도 수정
file_menu.add_radiobutton(label="계량기응답", variable=current_mode, command=lambda: handle_mode_selection("계량기응답"))
file_menu.add_radiobutton(label="계량기조회", variable=current_mode, command=lambda: handle_mode_selection("계량기조회"))
file_menu.add_radiobutton(label="TCP", variable=current_mode, command=lambda: handle_mode_selection("TCP"))
file_menu.add_separator()
file_menu.add_command(label="종료", command=exit_program)
menubar.add_cascade(label="모드", menu=file_menu)

# 편집 메뉴
edit_menu = tk.Menu(menubar, tearoff=0)
edit_menu.add_command(label="복사")
edit_menu.add_command(label="붙여넣기")
menubar.add_cascade(label="편집", menu=edit_menu)

# 도움말 메뉴
help_menu = tk.Menu(menubar, tearoff=0)
help_menu.add_command(label="사용법")
help_menu.add_command(label="정보")
menubar.add_cascade(label="도움말", menu=help_menu)

# 메뉴 바 설정
root.config(menu=menubar)




check_email()



# 송수신 카운트 업데이트 함수
def update_title():
    global send_count, receive_count

    # 폰트 생성 및 크기 설정
    custom_font = font.Font(size=20)

    # 송수신 카운트 업데이트 및 폰트 적용
    send_label.config(text=f"TX({send_count})", font=custom_font)
    receive_label.config(text=f"RX({receive_count})", font=custom_font)
    

    if receive_count - send_count == 1:
        send_count += 1

    success_rate = 0 if send_count == 0 else receive_count / send_count * 100
    root.title(f"유무선시험기 {PROGRAM_VERSION} (Success Rate: {success_rate:.2f}%)")


    root.after(1000, update_title)  # 1초마다 업데이트

# 창 크기 설정
window_width = 800
window_height = 800
root.geometry(f"{window_width}x{window_height}")


# status_message_label을 정의했다고 가정
status_message_label = Label(root, text="READY")
status_message_label.pack()



def update_status_message():
    global rx_status
    # 폰트 크기와 굵은체 설정 (기본 폰트 크기의 2배, 굵은체)
    font_size = ("Arial", 24, "bold")  # Arial 폰트, 크기 24, 굵은체

    # 글씨 색 설정
    if rx_status == "OK":
        text_color = "blue"  # OK일 때 파란색
    elif rx_status == "ERROR":
        text_color = "red"    # ERROR일 때 적색
    else:
        text_color = "black"  # READY 또는 그 외 상태일 때 흑색

    # 배경색 설정 (연한 회색)
    background_color = "lightgray"

    # 고정 너비를 설정하여 바탕색이 변하지 않도록 설정 (ERROR의 글자 수에 맞춤)
    fixed_width = 10  # 가장 긴 "ERROR"에 맞춘 너비

    # 상태 메시지 업데이트 (폰트 크기, 굵은체, 글씨 색, 배경색, 고정 너비 적용)
    status_message_label.config(text=f"{rx_status}", font=font_size, fg=text_color, bg=background_color, width=fixed_width)



update_status_message()

# 연결 상태 라벨
status_label = tk.Label(root, text="연결되지 않음")
status_label.pack(pady=1)

# 시리얼 포트 및 보레이트 설정
port_var = tk.StringVar(value=SERIAL_PORT)  # 포트 선택 기본값
baudrate_var = tk.StringVar(value=DEFAULT_BAUDRATE)  # 보레이트 선택 기본값

port_label = tk.Label(root, text="포트 선택:")
port_label.pack()

# 시스템에서 사용 가능한 포트를 검색
available_ports = [port.device for port in serial.tools.list_ports.comports()]

port_combo = ttk.Combobox(root, textvariable=port_var)
port_combo['values'] = available_ports  # 동적으로 사용 가능한 포트 설정
port_combo.pack()

baudrate_label = tk.Label(root, text="보레이트 선택:")
baudrate_label.pack()

baudrate_combo = ttk.Combobox(root, textvariable=baudrate_var)
baudrate_combo['values'] = [1200, 2400, 4800, 9600, 19200, 38400, 57600, 115200]
baudrate_combo.pack()

# 텍스트 박스 추가 (데이터 전송)
send_frame = tk.Frame(root)
send_frame.pack(side=tk.LEFT, padx=10)

send_label = tk.Label(send_frame, text="TX(0)")
send_label.pack()

send_text = scrolledtext.ScrolledText(send_frame, width=30, height=40)
send_text.pack()

# 텍스트 박스 추가 (데이터 수신)
receive_frame = tk.Frame(root)
receive_frame.pack(side=tk.RIGHT, padx=10)

receive_label = tk.Label(receive_frame, text="RX(0)")
receive_label.pack()

receive_text = scrolledtext.ScrolledText(receive_frame, width=50, height=40)
receive_text.pack()

# 버튼 추가
start_button = tk.Button(root, text="시작", command=start_connection)
start_button.pack()
start_button.config(state=tk.NORMAL)

stop_button = tk.Button(root, text="멈춤", command=stop_connection)
stop_button.pack()
stop_button.config(state=tk.DISABLED)

exit_button = tk.Button(root, text="종료", command=exit_program)
exit_button.pack()
exit_button.config(state=tk.NORMAL)

# 화면 지우기 버튼 추가
clear_button = tk.Button(root, text="화면 지우기", command=clear_screen)
clear_button.pack()



# 프로그램 실행
root.mainloop()
