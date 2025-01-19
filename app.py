import streamlit as st
import pandas as pd
from io import BytesIO
from make import engine

# 이미지 표시
st.image("image.png", caption="와인 재고 관리 시스템", use_container_width=True)

# 엑셀 파일 업로드
st.write("OFF, ON, REMAIN 엑셀 파일을 업로드하세요.")
off_file = st.file_uploader("OFF 파일", type=["xlsx"], key="off")
on_file = st.file_uploader("ON 파일", type=["xlsx"], key="on")
remain_file = st.file_uploader("REMAIN 파일", type=["xlsx"], key="remain")

# 파일이 모두 업로드된 경우
if off_file and on_file and remain_file:
    # 엑셀 파일을 pandas DataFrame으로 읽기
    df_off = pd.read_excel(off_file)
    df_on = pd.read_excel(on_file)
    df_remain = pd.read_excel(remain_file)

    # 처리된 결과 엑셀 파일 생성
    result_file = engine(df_off, df_on, df_remain)

    # 결과 엑셀 파일 다운로드 버튼
    with open(result_file, "rb") as f:
        st.download_button(
            label="결과 엑셀 다운로드",
            data=f,
            file_name="result_with_formulas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
