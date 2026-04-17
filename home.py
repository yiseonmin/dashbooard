# -*- coding: utf-8 -*-
# ── 서버 시작 ──────────────────────────────────────────────
# streamlit run home.py --server.address 0.0.0.0
# ── 접속 방법 ──────────────────────────────────────────────
# http://12.127.200.209:8501/

import os
import re
import pandas as pd
import streamlit as st
import xlwings as xw
from streamlit_option_menu import option_menu   # 메뉴 라이브러리
from streamlit_calendar import calendar

import datetime
import time
from time import strftime
from datetime import date, timedelta

import json
import openpyxl
from win32con import PRINTRATEUNIT_PPM

from excel_manager import *

# ── 페이지 설정 ──────────────────────────────────────────────
st.set_page_config(page_title="IT 업무 스케쥴", page_icon="📋", layout="wide")

EXCEL_FILE = "data/basic_data.xlsx"
SHEET_NAME1 = "schedule"
SHEET_NAME2 = "holiday"

TODAY = pd.Timestamp(date.today())  # 오늘 날짜

# ──────────────────────────────────────────
# MENU - HOME
# 엑셀 파일 가져와서 보여주기
# ──────────────────────────────────────────
def read_excel_with_xlwings_IT(filename, sheet_name1, sheet_name2):
    print("=" * 50)
    print("엑셀 파일 열기:", filename)
    print("=" * 50)

    # ── 1) Excel 앱을 통해 파일 열기 ──────────────────
    # visible=False : 엑셀 창을 화면에 띄우지 않음
    # visible=True : 엑셀 창이 실제로 열림 (디버깅할 때 유용)
    app = xw.App(visible=False)
    book = app.books.open(filename)

    try:
        # ── 2) 시트 선택(IT) ──────────────────────────────
        st.subheader("오늘 우리팀의 할일 🖥️ : ",divider=True)

        sheet = book.sheets[sheet_name1]  # 이름으로 선택
        # sheet = book.sheets[0] # 순서로 선택가능 (0 = 첫번째)
        print(f"- 시트 선택 완료: [{sheet.name}]")

        # 데이터 전체를 DataFrame으로 읽기 ───────
        # used_range : 데이터가 있는 영역을 자동으로 가져옴
        df = sheet.used_range.options(pd.DataFrame, index=False, header=True).value

        df["Date"] = pd.to_datetime(df["Date"])  # 날짜 열만
        today = pd.Timestamp(date.today())
        today_df = df[df["Date"] == today]
        # print(f" - 오늘은 {today}")
        # print(f" - 가져온 날짜\n {df['date']}")
        # print(f" - 가져온 오늘의 업무\n {today_df}")

        if today_df.empty:
            st.subheader("데이타 없음")
        else:
            for _, row in today_df.iterrows():
                print("오늘 우리팀의 할일  : ")
                print(f"   날짜:  {row['Date'], strftime('%Y-%m-%d')}")

        today_df["Date"] = today_df["Date"].dt.strftime('%Y-%m-%d')
        today_df =today_df.fillna(" ")  # None -> 공백으로 표시
        st.dataframe(today_df, width="stretch")

        # print(f"\n전체 데이터 shape: {df.shape} (행 수, 열 수)")
        # print(f" 컬럼 목록: {list(df.columns)}")
        # print()
        # # ── 한 줄씩 읽기 ───────────────────────────
        # print("-" * 50)
        # print("한 줄씩 읽기 시작")
        # print("-" * 50)
        #
        # for row_index, row in df.iterrows():
        #     # row_index : 0, 1, 2, 3 ... (pandas 기준 번호)
        #     # row : 각 행의 데이터 (Series 형태)
        #
        #     print(f"\n[{row_index + 1}번째 줄]")
        #
        #     # 열 이름과 값을 함께 출력
        #     for col_name, value in row.items():
        #         print(f"  {col_name} : {value}")
        #
        # print("\n" + "=" * 50)
        # print("읽기 완료!")
        # print("=" * 50)

        #return df # DataFrame을 반환 (다른 곳에서 활용 가능)

        # ── 3) 시트 선택 (휴가자 현황) ──────────────────────────────
        st.subheader("오늘 팀 현황 🕶️ : ", divider=True)
        sheet = book.sheets[sheet_name2]  # 이름으로 선택
        # print(f"- 시트 선택 완료: [{sheet_name2}]")
        df = sheet.used_range.options(pd.DataFrame, index=False, header=True).value
        df["Date"] = pd.to_datetime(df["Date"])  # 날짜 열만
        today = pd.Timestamp(date.today())
        today_df = df[df["Date"] == today]
        if today_df.empty:
            st.subheader("데이타 없음")
        # else:
            # for _, row in today_df.iterrows():


        today_df["Date"] = today_df["Date"].dt.strftime('%Y-%m-%d')
        today_df =today_df.fillna("")  # None -> 공백으로 표시
        st.dataframe(today_df, width="stretch")

    finally:
        # ── 반드시 정리! ───────────────────────────
        # 파일 닫기 + Excel 앱 종료 (안하면 Excel이 백그라운드에 남아있음)
        book.close()
        app.quit()
        print("엑셀 파일 닫기 완료")
        print("E N D ")

# ──────────────────────────────────────────
# MENU - fileupload
# ──────────────────────────────────────────
# 파일 업로드 메뉴 선택
def upload_file():
    # 파일 업로드
    attached_file = st.file_uploader('파일을 업로드 하세요.', type=['xlsx', 'py', 'sh', 'doc', 'ppt'])
    if attached_file is not None:  # 파일을 넣을 경우에만 실행 함
        # 실제 파일 저장
        save_upload_file('sharedFiles', attached_file)

    # 업로드전에 공유된 파일 목록 가져와서 보여주기
    folder_path = './sharedFiles'  # 파일이 저장된 폴더 경로 지정
    if os.path.exists(folder_path):
        files = os.listdir(folder_path)
        st.write(f"### {folder_path}  파일 리스트")
        # st.table(files)  # 또는 st.write(files)

        for file_name in files:
            file_path = os.path.join(folder_path, file_name)

            with open(file_path, "rb") as file:
                st.download_button(
                    label=f"📤 {file_name} 다운로드",
                    data=file,
                    file_name=file_name,
                    mime="application/octet-stream"
                )
    else:
        st.error("폴더를 찾을 수 없습니다.")

# 실제 업로드 파일 저장
def save_upload_file(directory, file):
    st.title("파일 업로드")
    if not os.path.exists(directory):  # 해당 이름의 폴더가 존재하는지 여부 확인
        os.makedirs(directory)  # 폴더가 없다면 폴더를 생성한다.
    file_path = os.path.join(directory, file.name)
    with open(file_path, 'wb') as f: #해당 경로의 폴더에서 파일의 이름으로 생성하겠다.
        f.write(file.getbuffer())  # 해당 내용은 Buffer로 작성하겠다.
        # 기본적으로 이미즈는 buffer로 저장되고 출력할때도 buffer로 출력한다.
    return st.success('파일 업로드 성공!')

# 엑셀 레포트 취합 화면
def export_excel_report():
    st.title("엑셀 레포트 생성 파이썬 수행")
    # py
    # report_creator.py - i
    # IR20260306 - b
    # TR260206_IDCEVO_SOP28_720.xlsx

    st.markdown("""
        <style>
        div.stButton > button:first-child {
            height: 2em;
            width: 15em;
            font-size: 20px;
            font-weight: bold;
        }
        </style>""", unsafe_allow_html=True)
    in_val1, in_val2 = st.columns(2)

    with in_val1:
        input_ir = st.text_input("입력 - i :IR20260306", placeholder="IR20260306 입력")

    with in_val2:
        input_xls = st.text_input("입력 - b :TR260206_IDCEVO_SOP28_720.xlsx", placeholder="TR260206_IDCEVO_SOP28_720.xlsx 입력")

    if st.button("명령 라인 추출", use_container_width=True):
        st.session_state.page = "excel_export"
        print("aAAddddddddddddddddddA")
        txt_ir = input_ir.replace(" ", "")
        if txt_ir == "":
            st.warning("IR을 입력하세요.")
            return
        txt_xls = input_xls.replace(" ", "")
        if txt_xls == "":
            st.warning("엘셀 레포트 파일명을 입력하세요.")
            return

        st.subheader("py report_creator.py - i "+ txt_ir + " - b "+ txt_xls)
        # st.code(f"py report_creator.py - i {txt_ir} - b {txt_xls}")

        if st.button("실행", type="primary"):
            with st.spinner("실행 중"):
                import subprocess

                result = subprocess.run(
                    ["python", "파이썬.py", "-i", txt_ir, "-b", txt_xls],
                    capture_output=True,
                    text=True,
                    encoding="utf-8"
                )

            st.markdown("결과")

            if result.returncode == 0:
                st.success("완료")
                if result.stdout:
                    st.code(result.stdout)

            else:
                st.error("오류")
                if result.stdout:
                    st.code(result.stderr)


# ──────────────────────────────────────────
# 실행
# ──────────────────────────────────────────
def main():
    try:
        # ------------------------------------------
        # 오른쪽 화면
        # ------------------------------------------
        # calendar() # 달력

        # 반환된 DataFrame 추가 활용 예시
        # print("\n 특정 열만 보기 예시:")
        # print(df["날짜"])  # 날짜 열만
        # print(df.iloc[0]) # 첫 번째 행만
        # print(df.iloc[2, 1]) # 3번째 행, 2번째 열 값만

        # ------------------------------------------
        # 왼쪽 메뉴
        # ------------------------------------------
        with st.sidebar:
            st.sidebar.title('MENU')
            # 1 메뉴형식 - 버튼식
            # if st.button("홈", use_container_width=True):
            #     st.session_state.page = "home"

            # 2. 메뉴형식 - 라이브러리 사용
            selected = option_menu(
                menu_title="PEG1 협업",
                options=["홈", "TC현황", "엑셀레포트 취합", "파일 업로드"],
                icons=["check2-square", "filetype-xls", "folder"],
                menu_icon="cast",
                default_index=0  # This sets 'Settings' as default
            )
            # # 3. 메뉴형식 - selectbox
            # menu = ['메뉴 선택', '파일 업로드']
            # selected = st.sidebar.selectbox('메뉴를 선택하세요.', menu)

            # 4. 달력표시
            selected_date = st.date_input("날짜 선택 : ")

        print("----------->", selected)
        # if st.session_state.page == "home":
        #     read_excel_with_xlwings_IT(EXCEL_FILE, SHEET_NAME1, SHEET_NAME2)

        if selected == "홈":
            # 1. 오늘 날짜 및 요일 가져오기
            now = datetime.datetime.now()
            date_str = now.strftime("%Y-%m-%d")  # 연-월-일 포맷
            weekday_str = now.strftime("%A")  # 요일 (영어)

            # 요일을 한글로 변경 (옵션)
            days_kr = ["월", "화", "수", "목", "금", "토", "일"]
            weekday_kr = days_kr[now.weekday()]

            st.set_page_config(page_title="PEG1팀 IT 일정", page_icon="", layout="wide")
            st.header(f"Hi! PEG1 ~  오늘은 **{date_str}** ({weekday_kr}요일) 입니다")
            st.subheader("✔️ 이번주 주간 미팅은 없습니다.")
            st.subheader("✔️ 다음 주 : 세미나, 1 on 1 계속..", divider="gray")
            
            read_excel_with_xlwings_IT(EXCEL_FILE, SHEET_NAME1, SHEET_NAME2)
        elif selected =="TC현황":
            upload_file()
        elif selected == "엑셀레포트 취합":
            st.subheader("엑셀파일 취합하기 : ", divider=True)
            export_excel_report()
        elif selected == "파일 업로드":
            st.subheader("파일 업로드하기 : ", divider=True)
            upload_file()
        else:
            print("end")

    except TypeError as e:
        print(f"오류 발생: {e}")
        print("E N D ")


if __name__ == "__main__":
    print("S T A R T" )
    print("1 *" * 20 )
    main()
