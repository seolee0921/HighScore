import sys
from PyQt5.QtWidgets import QApplication, QLabel,QWidget, QScrollArea, QFrame, QPushButton, QVBoxLayout, QHBoxLayout, QLineEdit, QTextEdit, QStackedWidget, QListWidget, QListWidgetItem, QButtonGroup, QSlider, QFileDialog
from PyQt5.QtGui import QIcon, QFont, QFontDatabase, QColor, QFontMetrics, QBrush, QPainter, QPen
from PyQt5.QtWidgets import QApplication, QWidget, QDesktopWidget, QGraphicsDropShadowEffect
from PyQt5.QtCore import Qt, QRect, QPropertyAnimation, pyqtProperty, pyqtSignal
import os
from dataclasses import dataclass

from pyhwpx import Hwp
import time
import clipboard
import spacy
import random
from langdetect import detect
import string
import os

## github seolee0921
 
@dataclass
class Hwp:
    path: str
    file_number: int
    pos_type: list[str]
    full_blank: bool
    percentage: int
    save_path: str

def mainProcess(location: str, pos_type: list[str], full_blank: bool, percentage: int, save_path: str):
    
    nlp = spacy.load("en_core_web_sm") # 영어 자연어 처리

    hwp = Hwp() # 한글 실행
    hwp.set_visible(False) # 실행 화면 가리기

    hwp.open(location) # 파일 열기


    # 표 내의 특정 셀 탐색
    def SetTableCellfind(addr):
        init_addr = hwp.KeyIndicator()[-1][1:].split(")")[0]  # 함수를 실행할 때의 주소를 기억
        if not hwp.CellShape:  # 표 안에 있을 때만 CellShape 오브젝트를 리턴함
            return
            # raise AttributeError("현재 캐럿이 표 안에 있지 않습니다.")
        if addr == hwp.KeyIndicator()[-1][1:].split(")")[0]:  # 시작하자 마자 해당 주소라면
            return True # 바로 종료
        hwp.Run("CloseEx")  # 그렇지 않다면 표 밖으로 나가서
        hwp.FindCtrl()  # 표를 선택한 후
        hwp.Run("ShapeObjTableSelCell")  # 표의 첫 번째 셀로 이동함(A1으로 이동하는 확실한 방법 & 셀선택모드)
        while True:
            current_addr = hwp.KeyIndicator()[-1][1:].split(")")[0]  # 현재 주소를 기억해둠
            hwp.Run("TableRightCell")  # 우측으로 한 칸 이동(우측끝일 때는 아래 행 첫 번째 열로)
            if current_addr == hwp.KeyIndicator()[-1][1:].split(")")[0]:  # 이동했는데 주소가 바뀌지 않으면?(표 끝 도착)
                # == 한 바퀴 돌았는데도 목표 셀주소가 안 나타났다면?(== addr이 표 범위를 벗어난 경우일 것)
                SetTableCellfind(init_addr)  # 최초에 저장해둔 init_addr로 돌려놓고
                hwp.Run("Cancel")  # 선택모드 해제
                return False
                # raise AttributeError("입력한 셀주소가 현재 표의 범위를 벗어납니다.")
            if addr == hwp.KeyIndicator()[-1][1:].split(")")[0]:  # 목표 셀주소에 도착했다면?
                # print(hwp.KeyIndicator()[-1][1:].split(")")[0])
                return True # 함수 종료

    # 고루고루 단어 선정하기기
    def select_percent(arr):
        if percentage == 100:
            selected_items = arr
            return selected_items

        # 배열을 5개 정도의 작은 묶음으로 나누기
        num_groups = 5
        group_size = len(arr) // num_groups
        
        # 배열을 고르게 분할
        groups = [arr[i*group_size: (i+1)*group_size] for i in range(num_groups)]
        
        # 퍼센트 샘플링을 위한 리스트
        selected_items = []
        tmp_items = []
        
        for group in groups:
            # 각 그룹에서 percentage 만큼 랜덤으로 선택
            num_to_select = int(len(group) * (percentage / 100))
            tmp_items = random.sample(group, num_to_select)
            selected_items += sorted(tmp_items, key=lambda x: group.index(x))
        
        # 선택된 항목들을 원본 배열에서의 순서대로 정렬
        # selected_items_sorted = sorted(selected_items, key=lambda x: arr.index(x))

        return selected_items

    def extract_single_word(text):
        # spaCy 텍스트 처리
        doc = nlp(text)
        
        # 영어 단어만 추출 (기호 제외)
        for token in doc:
            if token.is_alpha:  # 알파벳 문자로만 이루어진 단어인지 확인
                return token.text  # 첫 번째 알파벳으로 이루어진 단어 반환
        return ""  # 알파벳으로 이루어진 단어가 없으면 None 반환

    def blank(arr):
        while True:
            if not hwp.MoveNextWord():
                break
            fmove = 0
            if full_blank == 0: fmove = 1
            bmove = 0
            # 단어 가져오기 
            hwp.MoveSelWordEnd()
            hwp.Copy()
            copy_word = clipboard.paste()
            
            if copy_word != " " and copy_word != "\n":
                refine_word = extract_single_word(copy_word) # 단어 정제
                
                # 앞에 특수기호가 존재
                if copy_word[0] in string.punctuation + '"' + '“':
                    for text in copy_word:
                        if text in string.punctuation + '"' + '“':
                            fmove += 1    
                        else: break

                # 뒤에 특수기호가 존재
                if copy_word[-1] in string.punctuation + '"' + '”':
                    for text in reversed(copy_word):
                        if text in string.punctuation + '"' + '”':
                            bmove += 1
                        else: break

                # 정제된 단어가 뚫기 리스트에 들어가 있을때
                if refine_word in arr:
                    print(f"{copy_word} --> fmove: {fmove}, bmove: {bmove}")

                    # 커서를 단어 앞으로 이동
                    hwp.MovePrevWord()
                    hwp.MoveNextWord()
                    
                    # 전체 뚫기기
                    hwp.MoveSelWordEnd()
                    hwp.set_font(UnderlineColor="Black", UnderlineType=1, TextColor="White")
                    
                    if fmove != 0 or bmove != 0:
                        invisible_sign(fmove, bmove, copy_word)


            # gui 제작 후 예외처리 작성하기, 너무 많이 불러오면 traceback 발생
            for _ in range(3):
                try:
                    clipboard.copy(" ")
                    break
                except clipboard.PyperclipWindowsException:
                    print("Clipboard access error, retrying...")
                    time.sleep(2)

    def invisible_sign(f, b, word):
        if f:
            hwp.MovePrevWord()
            hwp.MoveNextWord()

            for i in range(f):
                hwp.MoveSelNextChar()
            
            hwp.insert_text(word[0:f])

            hwp.MoveWordBegin()
            
            for i in range(f):
                hwp.MoveSelNextChar()
            
            hwp.set_font(UnderlineColor="White", TextColor="Black")

        if b:
            hwp.MoveWordEnd()
            for i in range(b):
                hwp.MoveSelPrevChar()

            hwp.insert_text(word[-b:])

            hwp.MoveWordEnd()
            for i in range(b):
                hwp.MoveSelPrevChar()
            
            hwp.set_font(UnderlineColor="White", TextColor="Black")

    table_pieces = 1
    while hwp.get_into_nth_table(table_pieces):
        table_pieces += 1


    # 모든 표를 탐색하며 빈칸뚫기 시작

    i = 1
    while hwp.get_into_nth_table(i):

        print(f"{i}번째 표 정보")

        pos_cnt = 0 # 총 품사 개수 카운팅
        if SetTableCellfind('A5'): # 표에서 영어 본문 셀 위치
            
            hwp.SelectAll() # 해당 셀 전체 선택

            # 영어 본문 가져오기
            hwp.Copy() 
            sentence = clipboard.paste()
            if sentence:
                if sentence != "":
                    if detect(sentence) == 'en':
                        doc = nlp(sentence) # 품사분석

                        pos_dict = {} # 품사 별 단어 저장할 dictionary

                        # 각 품사에 해당하는 단어들 리스트에 추가
                        for token in doc:
                            if token.pos_ != 'SPACE':
                                if token.pos_ not in pos_dict:
                                    pos_dict[token.pos_] = []  # 품사별 리스트 초기화
                                pos_dict[token.pos_].append(token.text) # token.pos는 품사를 뜻하는 정수 반환, token.pos_는 문자열 반환

                        # 결과 출력
                        for type in pos_type:
                            pos_cnt += len(pos_dict[type])

                        print("\n=========품사=========")
                        print(f"{pos_type} 개수: ", pos_cnt)
                        for pos, words in pos_dict.items():
                            print(f"{pos}: {', '.join(words)}")

                        # 지정 품사 목록에서 랜덤으로 단어 추출
                        pos_words = []
                        for type in pos_type:
                            pos_words += pos_dict[type]
                        blank_list = select_percent(pos_words)
                        
                        print("\n=========품사 선별=========")
                        print(f"단어 리스트 : {pos_words}")
                        print(f"선별된 단어 : {blank_list}")
                        print("빈칸 뚫기 시작")
                        hwp.MoveLeft()
                        hwp.MoveRight()

                        #빈칸뚫기 시작작
                        blank(blank_list)

                        print("\n")

        i += 1

    try:   
        hwp.save_as(f"{save_path}")
        
        if os.path.exists(f"{save_path}"):
            print("파일이 성공적으로 저장되었습니다.")
        else:
            print("파일 저장에 실패했습니다.")

        os._exit(0)

    except:
        print("Save Error")
        hwp.set_visible(True)

setting_box_width = 0

class BlankTest(QWidget):
    window_size_x = 1280
    window_size_y = 720
    hwp = []
    file_cnt = 0
    def __init__(self):
        super().__init__()
        

        # 윈도우 타이틀 설정
        self.setWindowTitle('Blank Test')
        self.setWindowIcon(QIcon("./image/logo.png"))
        self.setGeometry(0, 0, BlankTest.window_size_x, BlankTest.window_size_y)
        self.setStyleSheet("background-color: rgb(243, 243, 243);")
        self.setAcceptDrops(True)  # 드래그 앤 드롭 활성화

        def center():
            qr = self.frameGeometry()
            cp = QDesktopWidget().availableGeometry().center()
            qr.moveCenter(cp)
            self.move(qr.topLeft())
        center()
        
        # 드래그 박스 생성
        self.SettingBox()
        self.FileBox()

        # 파일 드롭시 불투명도 적용될 위젯
        self.drop_background = QWidget(self)
        self.drop_background.setStyleSheet("background-color: rgba(16, 39, 60, 0); margin: 10px; border-radius: 18px")

        sd2 = QGraphicsDropShadowEffect()
        sd2.setOffset(5, 5)
        sd2.setBlurRadius(15)
        sd2.setColor(QColor(0, 0, 0, 80))
        self.file_box.setGraphicsEffect(sd2)

        sd1 = QGraphicsDropShadowEffect()
        sd1.setOffset(5, 5)
        sd1.setBlurRadius(25)
        sd1.setColor(QColor(0, 0, 0, 100))
        self.setting_box.setGraphicsEffect(sd1)

        self.layout = QHBoxLayout(self)

        # self.file_box.setMaximumWidth(1300)
        self.setting_box.setFixedWidth(400)
        self.layout.addWidget(self.setting_box, 2)
        self.layout.addWidget(self.file_box, 6)
        
        self.setLayout(self.layout)
        
    def resizeEvent(self, event):
        print(f"setting box: {self.setting_box.width()}" )
        print(f"file box: {self.file_box.width()}")
        global setting_box_width
        setting_box_width = self.setting_box.width() 

        self.drop_background.resize(self.width(), self.height())
        super().resizeEvent(event)

    def SettingBox(self):
        self.setting_box = QStackedWidget()
        # self.setting_box.setMaximumWidth(450)
        self.setting_box.setObjectName("setting_box")
        self.setting_box.setStyleSheet("""QStackedWidget{background-color: rgb(255, 255, 255); border-radius: 18px;}""")


    def FileBox(self):
        # 파일 경로 및 버튼을 표시할 스택 위젯젯
        self.file_box = QStackedWidget()
        self.file_box.setObjectName("file_box")
        self.file_box.setStyleSheet("""background-color: rgb(255, 255, 255);border-radius: 18px;  """)
        

        # 파일 경로들을 표시할 위젯
        self.file_path = QWidget()
        self.file_path.setObjectName("file_path_box")
        self.file_path.setStyleSheet("""background-color: rgb(255, 255, 255);border-radius: 18px;  """)
        self.file_path_layout = QVBoxLayout(self.file_path)
        self.file_path_layout.setAlignment(Qt.AlignCenter)
        self.file_path_layout.setSpacing(10)
        scroll_area = QScrollArea()
        scroll_area.setWidget(self.file_path)
        scroll_area.setStyleSheet("""
                                  QScrollArea {
                                  margin: 7px;
                                  }
                                QScrollBar:vertical {
                                    background: #F0F0F0;  /* 스크롤바 배경 */
                                    width: 10px;         /* 스크롤바 너비 */
                                    margin: 0px 0px 0px 0px;  /* 스크롤바 여백 */
                                    border: none;
                                }
                                QScrollBar::handle:vertical {
                                    background: #888; /* 스크롤바 색상 */
                                    min-height: 20px;  /* 핸들의 최소 크기 */
                                    border-radius: 4px;
                                }
                                QScrollBar::handle:vertical:hover {
                                    background: #555;  /* 마우스가 올려졌을 때 핸들 색상 */
                                }
                                QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                                    background: none;  /* 화살표 버튼 숨기기 */
                                    height: 0px;
                                }
                                QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
                                    background: rgb(252, 252, 252);  /* 빈 영역 색상 */
                                }
                                  QScrollBar:horizontal {
                                    background: #F0F0F0;  /* 스크롤바 배경색 */
                                    height: 10px;         /* 수평 스크롤바 높이 */
                                    margin: 0px 0px 0px 0px;  /* 스크롤바 여백 */
                                    border: none;
                                }
                                QScrollBar::handle:horizontal {
                                    background: #888;      /* 스크롤바 핸들 색상 */
                                    border-radius: 4px;    /* 둥근 모서리 */
                                    min-width: 10px;       /* 핸들의 최소 너비 */
                                }
                                QScrollBar::handle:horizontal:hover {
                                    background: #555;      /* 마우스를 올렸을 때 핸들 색상 */
                                }
                                QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
                                    background: none;  /* 화살표 버튼 숨기기 */
                                    width: 0px;
                                }
                                QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {
                                    background: rgb(252, 252, 252);  /* 빈 영역 색상 */
                                }
        """)
        # 스크롤이 자동으로 활성화되도록 설정
        scroll_area.setWidgetResizable(True)

        self.file_box.addWidget(scroll_area)

        #버튼들 넣을 버튼 그룹 생성성
        self.file_group = QButtonGroup()
        self.file_group.setExclusive(True)

        # 파일 여는 버튼
        # self.file_open_button = QWidget()
        # self.file_box.addWidget(self.file_open_button)

        self.file_box.setCurrentIndex(0)

    def dragEnterEvent(self, event):
        # 드래그된 파일이 수락 가능하면 이벤트를 허용
        if event.mimeData().hasUrls():
            event.accept()
            self.drop_background.raise_()
            self.drop_background.setStyleSheet("""
                                                background-color: rgba(47, 128, 237, 0.2);
                                                margin: 10px;
                                                border-radius: 18px;
                                                border: 1.2px inset rgba(68, 69, 89, 0.9);
                                               """)
        else:
            event.ignore()
    
    def dragLeaveEvent(self, event):
        """파일이 창을 벗어나면 원래 색으로 돌아가기"""
        self.drop_background.lower()
        self.drop_background.setStyleSheet("background-color: rgba(0, 0, 0, 0); margin: 10px; border-radius: 18px;")

    def dropEvent(self, event):
        # 드롭된 파일 경로를 텍스트로 표시
        if event.mimeData().hasUrls():
            file_urls = event.mimeData().urls()
            file_paths = [url.toLocalFile() for url in file_urls]
            for path in file_paths:
                if path[-3:] == "hwp":
                    self.file_cnt += 1
                    self.hwp.append(Hwp(path, self.file_cnt, [], False, 50, ""))
                    self.addListWidget(path)

            # for hwp in self.hwp:
            #     print(f"path: {hwp.path}")
            #     print(f"file number: {hwp.file_number}")
            #     print(f"type: {hwp.pos_type}")
            #     print(f"first letter: {hwp.full_blank}")
            #     print(f"percentage: {hwp.percentage}")
            #     print()
                
            self.drop_background.lower()
            self.drop_background.setStyleSheet("background-color: rgba(0, 0, 0, 0); margin: 10px; border-radius: 18px;")
            
    def addListWidget(self, path):
        
        # file_path_box에 표시될 widget
        button_name = f'{self.file_cnt}'
        parent_button = QPushButton(os.path.basename(path))
        parent_button.setObjectName(button_name)
        parent_button.setCheckable(True)
        self.file_group.addButton(parent_button)
        parent_button.setFont(QFont(QFontDatabase.applicationFontFamilies(Nanum_UltraLight_id)[0]))
        parent_button.setStyleSheet("""
                                QPushButton {
                                    background-color: rgb(248, 248, 254);
                                    text-align: left;
                                    font-size: 15px;
                                    padding-left: 10px;
                                    border-radius: 15px;
                                }
                                QPushButton:hover {
                                    background-color: rgb(248, 248, 250);
                                    color: rgb(47, 128, 237);
                                    padding-top: 2px;
                                 }
                                QPushButton:checked {
                                    background-color: rgb(234, 242, 253);
                                    color: rgb(47, 128, 237);
                                    padding-top: 4px;
                                }
                                
        """)
        parent_button.setMaximumHeight(50)
        parent_button.setMinimumHeight(50)

        parent_button.clicked.connect(self.change_setting_box)
        
        self.file_path_layout.addWidget(parent_button)

        self.button_layout = QHBoxLayout(parent_button)
        self.button_layout.setContentsMargins(5, 5, 5, 5)
        self.button_layout.addStretch()
        

        file_info = QLabel()
        file_info.setObjectName(button_name)
        file_info.setStyleSheet("margin-left: 5px; background-color: rgba(0,0,0,0); color: rgb(249, 47, 96);")
        file_info.setFont(QFont("나눔바른고딕", 10))

        removeButton = QPushButton('❌')
        removeButton.setObjectName(button_name) 
        removeButton.setStyleSheet("""
                                QPushButton {
                                   text-align: center;
                                   padding-left: 0px;
                                   font-size: 16px;
                                }
                                QPushButton:hover {
                                   background-color: rgb(251, 251, 251);
                                }
        """)

        removeButton.setMaximumSize(40, 40)
        removeButton.setMinimumSize(40,40)

        sd = QGraphicsDropShadowEffect()
        sd.setOffset(0, 0)
        sd.setBlurRadius(10)
        sd.setColor(QColor(0, 0, 0, 120))
        removeButton.setGraphicsEffect(sd)  

        def path_delete(item: QPushButton):  
            # print(self.sender().objectName())
            try:
                # print(item.objectName()) 
                button_to_remove = self.file_path.findChild(QPushButton, self.sender().objectName())       
                self.file_path_layout.removeWidget(button_to_remove)
                button_to_remove.deleteLater()
                print(f"Delete file{self.sender().objectName()}")

            except:
                print("None")

        removeButton.clicked.connect(lambda: path_delete(parent_button))
        
        def duplicate():
            self.file_cnt += 1
            self.hwp.append(Hwp(path, self.file_cnt, [], False, 50, ""))
            self.addListWidget(path)

        duplicate_button = QPushButton("복사")
        duplicate_button.setStyleSheet("""
                                QPushButton {
                                   text-align: center;
                                   padding-left: 0px;
                                   font-size: 16px;
                                    margin: 2px;
                                }
                                QPushButton:hover {
                                   background-color: rgb(251, 251, 251);
                                }
        """)
        duplicate_button.clicked.connect(lambda: duplicate())
        

        self.button_layout.addWidget(file_info)
        self.info(button_name)
        
        self.button_layout.addWidget(duplicate_button)
        self.button_layout.addWidget(removeButton)


        # setting_box에 표시될 widget
        setting = QWidget()
        setting.setStyleSheet("QWidget{background-color: rgba(0, 0, 0, 0); margin:5px; border-radius: 18px; padding-top: 10px; padding-bottom: 10px}")
        setting_layout = QVBoxLayout(setting)
        
        file_name = ResizableLabel(path)

        pos_button_style = """QPushButton {
                                  color: rgb(68, 69, 89);
                                  background-color: rgb(252, 252, 252);
                                  border-radius: 10px;
                                  }
                                  QPushButton:hover {
                                    background-color: rgb(234, 242, 253);
                                    color: rgb(68, 69, 89);
                                  }
                                  QPushButton:checked {
                                    background-color: rgb(234, 242, 253);
                                    color: rgb(74, 113, 240);
                                    }
                                  """

        # 그림자 효과 추가 (아래쪽에만 적용)
        shadow1 = QGraphicsDropShadowEffect()
        shadow1.setBlurRadius(0)  # 그림자 흐림 정도
        shadow1.setOffset(0, 1)  # 그림자 위치 (아래쪽으로 이동)
        shadow1.setColor(QColor(0, 0, 0, 90))  # 반투명 검은색 그림자
        shadow2 = QGraphicsDropShadowEffect()
        shadow2.setBlurRadius(0)  # 그림자 흐림 정도
        shadow2.setOffset(0, 1)  # 그림자 위치 (아래쪽으로 이동)
        shadow2.setColor(QColor(0, 0, 0, 90))  # 반투명 검은색 그림자
        shadow3 = QGraphicsDropShadowEffect()
        shadow3.setBlurRadius(0)  # 그림자 흐림 정도
        shadow3.setOffset(0, 1)  # 그림자 위치 (아래쪽으로 이동)
        shadow3.setColor(QColor(0, 0, 0, 90))  # 반투명 검은색 그림자
        shadow4 = QGraphicsDropShadowEffect()
        shadow4.setBlurRadius(0)  # 그림자 흐림 정도
        shadow4.setOffset(0, 1)  # 그림자 위치 (아래쪽으로 이동)
        shadow4.setColor(QColor(0, 0, 0, 90))  # 반투명 검은색 그림자

        noun_button = QPushButton("NOUN")
        noun_button.setObjectName(f"NOUN {button_name}")
        noun_button.setCheckable(True)
        noun_button.setMaximumWidth(100)
        noun_button.setFixedHeight(45)
        noun_button.setGraphicsEffect(shadow1)
        noun_button.setStyleSheet(pos_button_style)
        
        verb_button = QPushButton("VERB")
        verb_button.setObjectName(f"VERB {button_name}")
        verb_button.setCheckable(True)
        verb_button.setMaximumWidth(100)
        verb_button.setFixedHeight(45)
        verb_button.setGraphicsEffect(shadow2)
        verb_button.setStyleSheet(pos_button_style)

        adj_button = QPushButton("ADJ")
        adj_button.setObjectName(f"ADJ {button_name}")
        adj_button.setCheckable(True)
        adj_button.setMaximumWidth(100)
        adj_button.setFixedHeight(45)
        adj_button.setGraphicsEffect(shadow3)
        adj_button.setStyleSheet(pos_button_style)

        adv_button = QPushButton("ADV")
        adv_button.setObjectName(f"ADV {button_name}") 
        adv_button.setCheckable(True)
        adv_button.setMaximumWidth(100)
        adv_button.setFixedHeight(45)
        adv_button.setGraphicsEffect(shadow4)
        adv_button.setStyleSheet(pos_button_style)

        noun_button.clicked.connect(lambda: self.insert_type(noun_button)) 
        verb_button.clicked.connect(lambda: self.insert_type(verb_button)) 
        adj_button.clicked.connect(lambda: self.insert_type(adj_button)) 
        adv_button.clicked.connect(lambda: self.insert_type(adv_button)) 
        
        # 버튼 사이에 들어갈 수직선 생성 함수
        def create_divider():
            """ 둥근 모양의 칸막이(구분선) 생성 """
            divider = QFrame()
            divider.setFixedSize(2, 35)  # 가로 4px, 세로 30px 크기
            divider.setStyleSheet("""
                QFrame {
                    background-color: #888;
                    border-radius: 2px;  /* 둥근 모양 */
                }
            """)
            return divider
        
        button_layout = QHBoxLayout()
        button_layout.setSpacing(0)
        button_layout.addWidget(noun_button)
        button_layout.addWidget(create_divider())
        button_layout.addWidget(verb_button)
        button_layout.addWidget(create_divider())
        button_layout.addWidget(adj_button)
        button_layout.addWidget(create_divider())
        button_layout.addWidget(adv_button)

        blank_type_layout = QHBoxLayout()
        blank_type = ToggleSwitch()
        blank_type_text = QLabel("앞 글자 힌트")
        blank_type_text.setAlignment(Qt.AlignCenter)
        blank_type_text.setFont(QFont("나눔바른고딕", 12))
        blank_type_text.setFixedWidth(90)
        blank_type_layout.setAlignment(Qt.AlignCenter)
        blank_type_layout.addWidget(blank_type)
        blank_type_layout.addWidget(blank_type_text)
        blank_type.setStyleSheet("margin: -2px;")

        def label_update(state):
            if state:
                blank_type_text.setText("전체 빈칸")
            else:
                blank_type_text.setText("앞 글자 힌트")
            
            self.hwp[int(button_name) - 1].full_blank = state

            self.info(button_name)

        blank_type.toggled_signal.connect(label_update)

        percentage_slide = PercentageSlider()

        def percentage_update(value):
            self.hwp[int(button_name) - 1].percentage = value

            self.info(button_name)

        percentage_slide.slide_signal.connect(percentage_update)
        
        save_file = SaveFileDialog()

        def save_path_update(info):
            self.hwp[int(button_name) - 1].save_path = info

            self.info(button_name)

        save_file.save_signal.connect(save_path_update)

        # 레이아웃에 위젯 추가
        setting_layout.addStretch(1)
        setting_layout.addWidget(file_name)
        setting_layout.addStretch(7)
        setting_layout.addLayout(blank_type_layout)
        setting_layout.addLayout(button_layout)
        setting_layout.addWidget(percentage_slide)
        setting_layout.addWidget(save_file)
        # 스택위젯에 위젯 추가
        self.setting_box.addWidget(setting)

    def info(self, number):
        try:
            change_text = self.file_path.findChild(QPushButton, number).findChild(QLabel, number)
            if change_text:
                if not self.hwp[int(number) - 1].pos_type:
                    change_text.setStyleSheet("margin-left: 5px; background-color: rgba(0,0,0,0); color: rgb(249, 47, 96);")
                    if self.hwp[int(number) - 1].full_blank:
                        if not self.hwp[int(number) - 1].save_path:
                            change_text.setText(f"None | 전체 빈칸 | {self.hwp[int(number) - 1].percentage} | 저장 경로 없음")
                        else:
                            change_text.setText(f"None | 전체 빈칸 | {self.hwp[int(number) - 1].percentage} | {os.path.basename(self.hwp[int(number) - 1].save_path)}")
                    else:
                        if not self.hwp[int(number) - 1].save_path:
                            change_text.setText(f"None | 앞 글자 힌트 | {self.hwp[int(number) - 1].percentage} | 저장 경로 없음")
                        else:
                            change_text.setText(f"None | 앞 글자 힌트 | {self.hwp[int(number) - 1].percentage} | {os.path.basename(self.hwp[int(number) - 1].save_path)}")
                else:
                    change_text.setStyleSheet("margin-left: 5px; background-color: rgba(0,0,0,0); color: rgb(0, 0, 0);")
                    text = ""
                    for i in self.hwp[int(number) - 1].pos_type:
                        text += i
                        if i == self.hwp[int(number) - 1].pos_type[-1]:
                            text += " "
                        else:
                            text += ", "

                    if self.hwp[int(number) - 1].full_blank: # 전체 빈칸
                        if not self.hwp[int(number) - 1].save_path: # 저장 경로 없음
                            change_text.setStyleSheet("margin-left: 5px; background-color: rgba(0,0,0,0); color: rgb(249, 47, 96);")
                            change_text.setText(f"{text}| 전체 빈칸 | {self.hwp[int(number) - 1].percentage} | 저장 경로 없음")
                        else: # {os.path.basename(self.hwp[int(number) - 1].save_path)}
                            if self.hwp[int(number) - 1].percentage == 0: 
                                change_text.setStyleSheet("margin-left: 5px; background-color: rgba(0,0,0,0); color: rgb(249, 47, 96);")
                            change_text.setText(f"{text}| 전체 빈칸 | {self.hwp[int(number) - 1].percentage} | {os.path.basename(self.hwp[int(number) - 1].save_path)}")
                    else: # 앞 글자 힌트 
                        if not self.hwp[int(number) - 1].save_path: # 저장 경로 없음
                            change_text.setStyleSheet("margin-left: 5px; background-color: rgba(0,0,0,0); color: rgb(249, 47, 96);")
                            change_text.setText(f"{text}| 앞 글자 힌트 | {self.hwp[int(number) - 1].percentage} | 저장 경로 없음")
                        else: # {os.path.basename(self.hwp[int(number) - 1].save_path)}
                            if self.hwp[int(number) - 1].percentage == 0: 
                                change_text.setStyleSheet("margin-left: 5px; background-color: rgba(0,0,0,0); color: rgb(249, 47, 96);")
                            change_text.setText(f"{text}| 앞 글자 힌트 | {self.hwp[int(number) - 1].percentage} | {os.path.basename(self.hwp[int(number) - 1].save_path)}")
        except:
            print("QLabel is none")
     
    def insert_type(self, item: QPushButton):
        # 그림자 효과 추가 (아래쪽에만 적용)
        shadow = QGraphicsDropShadowEffect() 
        shadow.setBlurRadius(0)  # 그림자 흐림 정도
        shadow.setOffset(0, 1)  # 그림자 위치 (아래쪽으로 이동)
        shadow.setColor(QColor(0, 0, 0, 90))  # 반투명 검은색 그림자

        info = item.objectName().split(" ")
        if info[0] in self.hwp[int(info[1]) - 1].pos_type: # 이미 들어있다면
            self.hwp[int(info[1]) - 1].pos_type.remove(info[0])
            item.setGraphicsEffect(shadow)
            # item.move(item.x(), item.y() - 1)
        else:
            self.hwp[int(info[1]) - 1].pos_type.append(info[0]) # 없었다면
            item.setGraphicsEffect(None)
            # item.move(item.x(), item.y() + 1)
        
        self.info(info[1])

    def path_info(self, num):
        print("\n\n\n\n\n\n\n\n=================== file info ===================")
        print(f"path: {self.hwp[num].path}")
        print(f"file number: {self.hwp[num].file_number}")
        print(f"type: {self.hwp[num].pos_type}")
        print(f"full blank: {self.hwp[num].full_blank}")
        print(f"percentage: {self.hwp[num].percentage}")
        print(f"save path: {self.hwp[num].save_path}")
        print("=================================================")

    def change_setting_box(self):
        self.path_info(int(self.sender().objectName()) - 1)
        self.setting_box.setCurrentIndex(int(self.sender().objectName()) - 1)

class ToggleSwitch(QWidget):
    toggled_signal = pyqtSignal(bool)
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedSize(60, 30)  # 스위치 크기 설정
        self.setStyleSheet("background: transparent; border: none;")  # 배경 투명

        self._position = 3  # 버튼(원)의 초기 위치 (OFF 상태)
        self.type = False  # 기본 상태 (OFF)

        # 애니메이션 설정 (버튼 내부 원이 움직이는 효과)
        self.animation = QPropertyAnimation(self, b"position")
        self.animation.setDuration(180)  # 애니메이션 속도 (밀리초)

    def paintEvent(self, event):
        """스위치 디자인 커스텀"""
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        # 배경 색상 (OFF 상태)
        bg_color = QColor("#d6d6d6")
        border_color = QColor("#d6d6d6")  # 테두리 색상
        circle_color = QColor("#f7f7f7")  # 버튼 색상

        if self.type:
            bg_color = QColor("#4a71f0")  # 활성화 배경
            circle_color = QColor("#f7f7f7")  # 활성화된 버튼 색상


        pen = QPen(border_color, 1)
        painter.setPen(pen)

        # 배경 그리기 (둥근 모서리)
        painter.setBrush(QBrush(bg_color))
        painter.drawRoundedRect(0, 0, self.width(), self.height(), 17, 17)

        shadow_color = QColor(0, 0, 0, 100)  # 반투명 검은색 (투명도 100)
        painter.setBrush(QBrush(shadow_color))
        painter.setPen(Qt.NoPen)
        painter.drawEllipse(self._position + 4, 7, 22, 22)  # 그림자 위치 (살짝 아래)

        # 버튼(슬라이더) 위치 설정 (애니메이션 적용)
        painter.setBrush(QBrush(circle_color))
        painter.setPen(Qt.NoPen)
        painter.drawEllipse(self._position, 3, 24, 24)  # 버튼(원) 그리기

    def toggle_state(self):
        """스위치 상태 변경 (애니메이션 적용)"""
        self.type = not self.type
        start_pos = self._position
        end_pos = self.width() - 27 if self.type else 3  # ON: 오른쪽, OFF: 왼쪽

        # 애니메이션 설정
        self.animation.setStartValue(start_pos)
        self.animation.setEndValue(end_pos)
        self.animation.start()

        self.toggled_signal.emit(self.type)

    def mousePressEvent(self, event):
        """마우스로 클릭하면 상태 변경"""
        self.toggle_state()
        super().mousePressEvent(event)

    def get_position(self):
        """애니메이션이 적용될 속성 값 가져오기"""
        return self._position

    def set_position(self, value):
        """애니메이션이 적용될 속성 값 설정"""
        self._position = value
        self.update()

    position = pyqtProperty(int, get_position, set_position)

class SaveFileDialog(QWidget):
    save_signal = pyqtSignal(str)
    def __init__(self):
        super().__init__()

        # QLabel: 선택한 파일 경로 표시
        self.label = QLabel("파일 저장", self)
        self.label.setFixedHeight(45)
        self.label.setMidLineWidth(280)
        self.label.setAlignment(Qt.AlignCenter)
        self.label.setFont(QFont("나눔바른고딕", 10))
        self.label.setStyleSheet("border: 1px solid rgb(249, 47, 96); padding: 5px; color: rgba(0, 0, 0, 0.8); border-radius: 3px;")

        # QPushButton: 파일 저장 경로 선택 버튼
        self.button = QPushButton("···", self)
    
        shadow = QGraphicsDropShadowEffect()
        shadow.setOffset(0, 0)
        shadow.setBlurRadius(10)
        shadow.setColor(QColor(0, 0, 0, 90))
        self.button.setGraphicsEffect(shadow)

        self.button.setStyleSheet("border-radius: 3px;")
        self.button.clicked.connect(self.select_save_path)
        self.button.setFont(QFont("나눔바른고딕", 10))

        # 레이아웃 설정
        layout = QHBoxLayout()
        layout.addWidget(self.label, 9)
        layout.addWidget(self.button, 1)
        self.setLayout(layout)

    def select_save_path(self):
        """파일 저장 경로 선택 다이얼로그 (.hwp 형식 제한)"""
        file_path, _ = QFileDialog.getSaveFileName(self, "경로 탐색", "", "HWP Files (*.hwp)")
        
        if file_path:  # 사용자가 경로를 선택한 경우
            if not file_path.endswith(".hwp"):  # 확장자 자동 추가
                file_path += ".hwp"
            self.label.setText(f"  {file_path}")
            self.label.setStyleSheet("border: 1px solid rgb(74, 113, 240); padding: 5px; border-radius: 3px;")
            self.save_signal.emit(file_path)

class PercentageSlider(QWidget):
    slide_signal = pyqtSignal(int)
    def __init__(self):
        super().__init__()

        # Percentage 라벨
        self.label = QLabel("Percentage", self)

        # 슬라이더 (0~100 범위)
        self.slider = QSlider(Qt.Horizontal, self)
        self.slider.setMinimum(0)
        self.slider.setMaximum(100)
        self.slider.setValue(50)  # 기본값
        self.slider.setTickInterval(10)
        self.slider.setTickPosition(QSlider.TicksBelow)

        # 값 입력을 위한 QLineEdit
        self.value_input = QLineEdit(self)
        self.value_input.setFixedHeight(36)
        self.value_input.setFixedWidth(55)  # 입력 창 크기
        self.value_input.setAlignment(Qt.AlignCenter)  # 숫자 중앙 정렬
        self.value_input.setFont(QFont("나눔바른고딕", 10, QFont.Bold))
        self.value_input.setText(str(self.slider.value()))  # 초기값 설정
 
        # 스타일 변경: 핸들(원형) 크기 키우기 + 트랙 색상 조정
        self.setStyleSheet("""
            QLabel {
                font-size: 14px;
                color: #000000;
            }
            QSlider::groove:horizontal {
                height: 4px;
                background: rgb(215, 215, 216);
            }
            QSlider::handle:horizontal {
                background: #4A71F0;
                border: 2px solid #ffffff;
                width: 16px;  /* 핸들 크기 증가 */
                height: 18px;
                margin: -8px 0;  /* 핸들이 트랙 위에 더 크게 올라오도록 설정 */
                border-radius: 10px;  /* 원형 핸들 */
            }
            QLineEdit {
                border: 1px solid rgba(215, 215, 216, 0.89);
                padding: -9px;
                border-radius: 6px;
                color: #4A71F0;
            }
        """)

        # 레이아웃 구성
        layout = QVBoxLayout()
        hlayout = QHBoxLayout()
        hlayout.addWidget(self.label)
        hlayout.addWidget(self.slider)
        hlayout.addWidget(self.value_input)
        layout.addLayout(hlayout)
        self.setLayout(layout)

        # 연결: 슬라이더 값 변경 -> 입력창 업데이트
        self.slider.valueChanged.connect(self.update_value)
        self.value_input.textChanged.connect(self.update_slider)

    def update_value(self, value):
        """슬라이더 값이 변경되면 입력창 업데이트"""
        self.value_input.setText(str(value))            
        self.slide_signal.emit(value)

    def update_slider(self):
        """입력창 값이 변경되면 슬라이더 업데이트"""
        text = self.value_input.text()
        if text.isdigit():
            value = int(text)
            if value > 100:
                self.value_input.setText("100")
                self.slider.setValue(100)
            elif 0 <= value <= 100:  # 0~100 범위 확인
                self.slider.setValue(value)

class ResizableLabel(QLabel):
    def __init__(self, path, parent=None):
        super().__init__(parent) 
        self.path = path  # 파일 경로 저장
        print(self.path)
        self.setAlignment(Qt.AlignCenter)  # 텍스트 중앙 정렬
        self.setTextInteractionFlags(Qt.TextInteractionFlag.NoTextInteraction)  # 텍스트 상호작용 비활성화
        self.update_elided_text()  # 초기 텍스트 설정

    def resizeEvent(self, event):
        """위젯 크기 변경 시 텍스트 업데이트"""
        global setting_box_width
        # print(setting_box_width)
        self.update_elided_text()
        super().resizeEvent(event)

    def update_elided_text(self): 
        """현재 QLabel 크기에 맞춰 텍스트를 ... 처리"""
        try:
            font_metrics = QFontMetrics(self.font())
            elided_text = font_metrics.elidedText(os.path.basename(self.path), Qt.ElideRight, int(setting_box_width * 0.85))
            self.setText(elided_text)
        except:
            print("update elided text is error")

if __name__ == '__main__':
    # 애플리케이션 객체 생성
    app = QApplication(sys.argv)
    
    """https://noonnu.cc/font_page/1456"""

    fontDB = QFontDatabase()
    font_count = 3
    Nanum_Light_id = fontDB.addApplicationFont('./font/나눔바른고딕/NanumBarunGothicLight.ttf')
    Nanum_UltraLight_id = fontDB.addApplicationFont('./font/나눔바른고딕/NanumBarunGothicUltraLight.ttf')
    Nanum_Bold_id = fontDB.addApplicationFont('./font/나눔바른고딕/NanumBarunGothicBold.ttf')
    app.setFont(QFont("나눔바른고딕"))

    for i in range(font_count):
        print(fontDB.applicationFontFamilies(i)[0])

    # 윈도우 객체 생성
    window = BlankTest()

    # 윈도우 표시
    window.show()

    # 애플리케이션 실행
    sys.exit(app.exec())