import pywinauto
import os
from pywinauto.application import Application
from pywinauto.application import ProcessNotFoundError
from bs4 import BeautifulSoup
from mlstripper import strip_tags
from time import sleep
import logging
from datetime import datetime as dt
from datetime import timedelta as td

class NateonMemo:
    SENDER_DESC = '보낸사람    :'
    RECEIVER_DESC = '받는사람    :'
    CC_DESC = '참조          :'

    def __init__(self, file_name):
        self.sender = ''
        self.receiver = None
        self.cc = None
        self.memo_time = ''
        self.file_name = file_name        


    def parse_memo(self):
        with open(self.file_name, 'r', encoding='UTF-16') as f:            
            soup = BeautifulSoup(f, 'html.parser')        
            divs_tag = soup.find_all('div')
            
            sender_index = -1
            receiver_index = -1

            for i, d in enumerate(divs_tag):
                text = strip_tags(d.get_text())
                if NateonMemo.SENDER_DESC in text:  
                    # sender가 이미 있으면 skip
                    if len(self.sender) > 0 or len(self.memo_time) > 0:
                        continue
                    self.parse_sender_and_time_(text)        

                    sender_index = i

                elif NateonMemo.RECEIVER_DESC in text: 
                    receiver_index_start = text.find(':') + 1
                    self.receiver = [s.split('/')[0].strip() for s in text[receiver_index_start:].split(';') if len(s.strip()) > 0]
                    
                    receiver_index = i

                elif NateonMemo.CC_DESC in text:
                    cc_index_start = text.find(':') + 1
                    self.cc = [s.split('/')[0].strip() for s in text[cc_index_start:].split(';') if len(s.strip()) > 0]
                    break   

                else:
                    if receiver_index != -1 and receiver_index + 1 == i:
                        break
            return True


    def parse_sender_and_time_(self, line):
        time_index_start = line.find('(') + 1
        time_index_end = line.find(')')
        memo_time = line[time_index_start : time_index_end]
        self.memo_time = memo_time.replace(':','')     

        sender_index_start = line.find(':') + 1
        sender = line[sender_index_start:time_index_start-1] 
        self.sender = sender.split('/')[0].strip()


    def make_file_name(self):
        # [2018-11-30 100241]-[sender]-R[receivers]-C[ccs].html
        ret = f'[{self.memo_time}]-[{self.sender}]-R['
        
        for r in self.receiver:
            ret += r
            ret += ','
            
        ret = ret[:-1]
        ret += ']'
            
        if self.cc is not None:
            ret += '-C['
            for c in self.cc:
                ret += c
                ret += ','        
            ret = ret[:-1]
            ret += ']'
        
        return ret[:125] + '.html'    


    def memo_date(self):
        return self.memo_time[:10]
    
BIZMEMOBOX_PATH = os.path.expanduser('~\\AppData\\Local\\SK Communications\\NATEON BIZ\\Addin\\7716FE45-BADF-460a-8376-11E3691206EF\\BizMemoBox.exe')

class NateonMemoAutomation:
    log = logging.getLogger("auto")
    log.setLevel(logging.INFO)
    formatter = logging.Formatter('[%(asctime)s][%(levelname)s] %(message)s')
    stream_handler = logging.StreamHandler()
    stream_handler.setFormatter(formatter)
    log.addHandler(stream_handler)

    def __init__(self, recv_path, send_path):
        self.recv_path = recv_path
        self.send_path = send_path

        self.bizmemobox_app = None
        self.bizmemobox_wnd = None
        self.tree_send_box = None
        self.tree_recv_box = None
        self.ctrl_memo_list = None
        self.ctrl_save_btn = None
        self.ctrl_save_dlg_filename_edit = None
        self.ctrl_save_dlg_ok_btn = None
        self.ctrl_from_date = None
        self.ctrl_to_date = None
        self.ctrl_send_recv_person = None
        self.ctrl_search_word = None
        self.search_btn = None
        self.ctrl_period_check = None
    
    
    def initialize(self):
        try:
            self.bizmemobox_app = Application().connect(path=BIZMEMOBOX_PATH)
        except ProcessNotFoundError:
            self.bizmemobox_app = Application().start(BIZMEMOBOX_PATH)    

        sleep(5)
            
        # 쪽지 윈도우
        self.bizmemobox_wnd = self.bizmemobox_app.window(title='통합메시지함')     
        
        # 좌측 트리 메뉴 - 받은쪽지함, 보낸쪽지함
        tree = self.bizmemobox_wnd.child_window(title="Tree1", class_name="SysTreeView32")
        self.tree_send_box = tree.get_item(['쪽지함','보낸쪽지함'])
        self.tree_recv_box = tree.get_item(['쪽지함','받은쪽지함'])      
        
        # 메모 목록
        self.ctrl_memo_list = self.bizmemobox_wnd.child_window(class_name='MsgBoxCtrl', control_id=0x4EEB)
        # 저장 버튼
        self.ctrl_save_btn = self.bizmemobox_wnd.child_window(class_name='Button', control_id=0x7D5)

        # 저장 시 뜨는 다이얼로그
        save_dlg = self.bizmemobox_app.다른_이름으로_저장
        # 파일명 입력
        self.ctrl_save_dlg_filename_edit = save_dlg.child_window(class_name='Edit', control_id=0x3E9)
        # OK 버튼
        self.ctrl_save_dlg_ok_btn = save_dlg.child_window(class_name='Button', control_id=0x1)        
        # 조회기간-from
        self.ctrl_from_date = self.bizmemobox_wnd.child_window(class_name='SysDateTimePick32', control_id=0x47B)
        # 조회기간-시작
        self.ctrl_to_date = self.bizmemobox_wnd.child_window(class_name='SysDateTimePick32', control_id=0x47D)      
        # 조회기간-check
        self.ctrl_period_check = self.bizmemobox_wnd.child_window(control_id=0x47C)

        # 보낸/받은사람
        self.ctrl_send_recv_person = self.bizmemobox_wnd.child_window(control_id=0x406)

        # 검색어
        self.ctrl_search_word = self.bizmemobox_wnd.child_window(control_id=0x405)

        # 찾기
        self.search_btn = self.bizmemobox_wnd.child_window(control_id=0x1780)


    # 받은쪽지함 선택
    def select_recv_box_tree(self):
        sleep(1)
        self.tree_recv_box.select()


    # 보낸쪽지함 선택
    def select_send_box_tree(self):
        sleep(1)
        self.tree_send_box.select()


    # 조회버튼 클릭
    def click_search_btn(self):
        sleep(1)
        self.search_btn.click()


    # 조회조건 설정
    def set_search_conditions(self, from_date, to_date, check_period = True, send_recv_person = '', search_word = ''):        
        # 날짜 parsing
        from_y, from_m, from_d = int(from_date[:4]), int(from_date[4:6]), int(from_date[6:])
        to_y, to_m, to_d = int(to_date[:4]), int(to_date[4:6]), int(to_date[6:])
        
        # 기간 설정
        self.ctrl_from_date.set_time(year=from_y, month=from_m, day=from_d)
        self.ctrl_to_date.set_time(year=to_y, month=to_m, day=to_d)
        
        # 기간체크박스 설정
        if check_period:
            self.ctrl_period_check.check()    
        else:
            self.ctrl_period_check.uncheck()

        # 보낸/받은사람
        self.ctrl_send_recv_person.set_edit_text(send_recv_person)
        # 검색어
        self.ctrl_search_word.set_edit_text(search_word)

        self.date_list = self.enum_period(from_date, to_date)


    # 기간 내의 모든 날짜를 yyyymmdd 형태의 list로 반환 (폴더 생성을 위함)
    def enum_period(self, from_date, to_date):
        day = 0
        days = []
        while(True):
            cur_dt = dt.strptime(from_date, '%Y%m%d') + td(day)
            days.append(cur_dt.strftime('%Y-%m-%d'))   
            if cur_dt == dt.strptime(to_date, '%Y%m%d'):
                break
            day += 1         
        return days        


    # 날짜별 폴더 생성
    def make_date_folder(self, parent_folder):
        for d in self.date_list:
            date_folder = os.path.join(parent_folder, d)     
            if not os.path.exists(date_folder):
                os.mkdir(date_folder)        


    # 받은쪽지함 저장
    def save_recv_memo_box(self):
        sleep(3)
        self.make_date_folder(self.recv_path)
        self.save_(self.recv_path) 
    

    # 보낸쪽지함 저장
    def save_send_memo_box(self):
        sleep(3)
        self.make_date_folder(self.send_path)
        self.save_(self.send_path) 
 

    def move_to_first_memo_pos(self):
        # 메모 목록 최상단 위치 찾아서 클릭 이벤트 발생
        OFFSET_X, OFFSET_Y = 200, 180        
        rect = self.bizmemobox_wnd.rectangle()
        pywinauto.mouse.move(coords=(rect.left + OFFSET_X, rect.top + OFFSET_Y))
        pywinauto.mouse.click(coords=(rect.left + OFFSET_X, rect.top + OFFSET_Y))


    # 진짜 save~
    def save_(self, save_path):
        # 종료 조건 체크위함 (직전 저장 파일명과 동일하면 끝났다고 간주)
        last_saved_file_name = ''
        
        # 임시 파일
        tmp_file_name = os.path.join(save_path, 'tmp.html')        
        # 기존의 tmp.html 삭제      
        if os.path.exists(tmp_file_name):
            os.remove(tmp_file_name)   
        
        while True:    
            # 쪽지 저장 버튼 클릭
            self.ctrl_save_btn.click()
            # temp 파일명 입력
            self.ctrl_save_dlg_filename_edit.set_edit_text(tmp_file_name)
            # enter
            self.ctrl_save_dlg_ok_btn.type_keys('{ENTER}')

            # 1초간 텀, 파일 내부 parsing 시도 시 아직 파일 저장이 끝나지 않을 수도 있어서 루프 돌며 체크
            sleep(1)
            file_found = False
            
            while not file_found:
                try:
                    m = NateonMemo(tmp_file_name) 
                    m.parse_memo()

                    file_found = True

                    new_file_folder = os.path.join(save_path, m.memo_date())
                    new_file_name = os.path.join(new_file_folder, m.make_file_name())
                    # 직전 파일명과 현재 파일명 동일하면 tmp파일 삭제후 종료
                    if new_file_name == last_saved_file_name:
                        os.remove(tmp_file_name)             
                        return
                    else:
                        last_saved_file_name = new_file_name
                        # 예전에 저장한거 또 저장하면 FileExistsError exception 발생하므로 skip. tmp파일 삭제
                        try:
                            os.rename(tmp_file_name, new_file_name)
                            NateonMemoAutomation.log.info(f'{new_file_name} saved!')

                        except FileExistsError as e:
                            os.remove(tmp_file_name)
                            NateonMemoAutomation.log.info(f'{new_file_name} exists! skip!')

                        # key down event
                        sleep(1)
                        self.ctrl_memo_list.type_keys('{DOWN}')    

                except FileNotFoundError as e:
                    NateonMemoAutomation.log.info('wait for file save...')
                    sleep(1)
            