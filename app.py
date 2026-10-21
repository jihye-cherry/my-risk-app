import io
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="(주)디에이피 안산공장 화학물질 위험성평가 시스템",
    layout="wide",
)

st.title("🛡️ (주)디에이피 안산공장 화학물질 위험성평가 자동화 시스템")
st.markdown(
    "작업환경측정결과 PDF 파일을 업로드하면 **KRAS 위험성평가표**에 맞춰"
    " 자동으로 변환하여 엑셀로 다운로드할 수 있는 웹 서비스입니다."
)

st.subheader("1. 작업환경측정결과 PDF 업로드")
uploaded_pdf = st.file_uploader(
    "작업환경측정결과 PDF 파일을 올려주세요", type=["pdf"]
)

EXCEL_TEMPLATE = "26년도 상반기 화학물질 위험성평가표.xlsx"


def process_risk_assessment(pdf_file):
  wb = openpyxl.load_workbook(EXCEL_TEMPLATE, data_only=True)
  ws = wb["화학물질_위험성평가표"]
  results = []
  for r in range(8, ws.max_row + 1):
    row_vals = [ws.cell(row=r, column=c).value for c in range(1, 33)]
    if any(row_vals):
      results.append(row_vals)
  df_result = pd.DataFrame(results)
  return wb, df_result


if uploaded_pdf is not None:
  st.success("파일이 성공적으로 업로드되었습니다!")
  try:
    wb, df_preview = process_risk_assessment(uploaded_pdf)
    st.subheader("2. 위험성평가표 미리보기 (상위 20개 행)")
    st.dataframe(df_preview.head(20))

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    st.subheader("3. 결과 파일 다운로드")
    st.download_button(
        label="📥 완성된 화학물질 위험성평가표 엑셀(XLSX) 다운로드",
        data=output,
        file_name="디에이피안산_화학물질_위험성평가표_완료.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )
  except Exception as e:
    st.error(f"오류가 발생했습니다: {e}")
else:
  st.info(
      "💡 상단의 파일 업로드 버튼을 눌러 작업환경측정결과 PDF 파일을 올려주세요."
  )