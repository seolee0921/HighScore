from pyhwpx import Hwp
import time
import clipboard
import spacy
import random
from langdetect import detect
import string
import os

## github seolee0921

#################################################################

## 문서 경로 설정 ##
location = "D:\HighScore\hwp_resource\둔촌1_교과서_7과.hwp"

# 빈칸 뚫고 싶은 품사 지정
pos_type = ["NOUN", "VERB", "ADJ"]

# 전체 빈칸 여부 (True or False)
full_blank = True

# 빈칸 빈도
percentage = 60

# 저장할 파일 이름
test_name = "둔촌1_교과서_7과 - 빈칸테스트 2단계"

# 저장할 위치
save_path = "D:\HighScore\exam"

#################################################################

nlp = spacy.load("en_core_web_sm")

hwp = Hwp()
hwp.set_visible(False)

hwp.open(location)


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
    hwp.save_as(f"{save_path}\{test_name}.hwp")
    
    if os.path.exists(f"{save_path}\{test_name}.hwp"):
        print("파일이 성공적으로 저장되었습니다.")
    else:
        print("파일 저장에 실패했습니다.")

    os._exit(0)

except:
    print("Save Error")
    hwp.set_visible(True)