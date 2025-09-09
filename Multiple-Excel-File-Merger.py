import streamlit as st
import pandas as pd
from io import BytesIO

# Streamlit 페이지 설정
st.set_page_config(
    page_title="다중 엑셀 파일 병합기",
    layout="wide",
)

# 제목과 헤더 설정
st.title('다중 엑셀 파일 병합기')
st.markdown('### 여러 개의 엑셀 파일을 하나의 데이터프레임으로 병합')
st.markdown(':red[관세청(unipass) 수출 이행 내역과 같이, 동일한 형태의 컬럼을 가진 여러 개의 파일을 하나로 병합할 때 유용합니다.]')

# ---
# 사이드바에 파일 업로드 섹션 생성
# ---

with st.sidebar:
    st.header('파일 업로드')
    # .xls와 .xlsx 확장자를 가진 여러 파일 업로드 허용
    uploaded_files = st.file_uploader(
        "엑셀 파일을 선택하세요",
        type=['xls', 'xlsx'],
        accept_multiple_files=True
    )

# ---
# 메인 화면에 내용 표시
# ---

# 업로드된 파일이 있는지 확인
if uploaded_files:
    # 데이터프레임을 저장할 빈 리스트 초기화
    all_dfs = []
    
    # 업로드된 각 파일 처리
    for file in uploaded_files:
        try:
            # 각 엑셀 파일을 읽어 데이터프레임으로 변환
            df = pd.read_excel(file)
            all_dfs.append(df)
        except Exception as e:
            st.error(f"파일을 읽는 중 오류가 발생했습니다: {file.name}: {e}")
    
    # 성공적으로 읽힌 데이터프레임이 하나라도 있을 경우
    if all_dfs:
        # 모든 데이터프레임을 하나의 데이터프레임으로 병합 (행 기준으로 추가)
        merged_df = pd.concat(all_dfs, ignore_index=True)
        
        # 중복되는 행 제거
        merged_df.drop_duplicates(inplace=True)
        
        st.success("파일이 성공적으로 병합되었으며 중복이 제거되었습니다!")
        st.markdown("### 병합된 데이터프레임 미리보기")
        
        # 병합된 데이터프레임 화면에 표시 (인덱스 숨김)
        st.dataframe(merged_df, hide_index=True)
        
        # 병합된 데이터프레임을 다운로드할 수 있는 버튼 생성
        # 데이터프레임을 메모리 내에서 엑셀 파일 형식으로 변환
        @st.cache_data
        def convert_df_to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Merged_Data')
            processed_data = output.getvalue()
            return processed_data
        
        excel_data = convert_df_to_excel(merged_df)
        
        st.download_button(
            label="병합된 엑셀 파일 다운로드",
            data=excel_data,
            file_name="merged_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("병합할 유효한 엑셀 파일을 찾을 수 없습니다.")

else:
    st.info("사이드바를 사용하여 하나 이상의 엑셀 파일을 업로드해주세요.")
