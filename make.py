import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from utils import match_wine_names

def engine(df_off, df_on, df_remain):
    """
    1) OFF, ON, REMAIN 데이터를 이용해 엑셀에 수식을 입력해 저장
    2) 저장된 파일을 다시 열어 L열의 실제 값이 4보다 작은 경우를 확인하고 빨간색 처리
    3) 최종 저장 후 파일 경로 리턴
    """
    
    # 매칭 및 데이터 처리
    off_array = list(df_off['상 품 명'].values[~pd.isnull(df_off['상 품 명'].values)])
    on_array = list(df_on['품목명'].values[~pd.isnull(df_on['품목명'].values)])
    remain_array = list(df_remain[df_remain.columns[1]].values[~pd.isnull(df_remain[df_remain.columns[1]].values)])[1:]

    red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

    # 결과 DataFrame 생성
    columns = [
        'No.', '상품소분류', '상 품 명', 'OFF 수량', 'ON 수량', 'TOTAL 수량', '기간 개월수', '월판매량 수량',
        '현재고 수량', '발주분 수량', '총재고 수량', '재고월수', '3개월후 재고', '발주희망', '발주수량 병',
        '발주수량 박스단위', '예상발주 박스', '발주확정 박스', '발주확정 병'
    ]
    df = pd.DataFrame(columns=columns)

    df['상 품 명'] = off_array
    df['OFF 수량'] = df_off['합 계'].values[1:-2]

    # on 수량 dict
    on_dict = {df_on['품목명'][i]: df_on['EA'][i] for i in range(len(df_on)) if df_on['품목명'][i] in on_array}
    
    # wine name 매칭
    result_on_off = match_wine_names(off_array, on_array)
    result_off_remain = match_wine_names(off_array, remain_array)
    
    # remain 수량 dict
    remain_dict = {
        df_remain[df_remain.columns[1]][i]: df_remain[df_remain.columns[9]][i] 
        for i in range(len(df_remain)) 
        if df_remain[df_remain.columns[1]][i] in remain_array
    }
    
    df['ON 수량'] = [on_dict[result_on_off[name]] if name in result_on_off else 0 for name in off_array]
    df['현재고 수량'] = [remain_dict[result_off_remain[name]] if name in result_off_remain else 0 for name in off_array]

    # 1) 엑셀 파일에 수식 입력 및 저장
    output_file = 'result_with_formulas.xlsx'
    df.to_excel(output_file, index=False, engine='openpyxl')

    wb = load_workbook(output_file)
    ws = wb.active

    for row in range(2, ws.max_row + 1):
        # 수식 입력
        ws[f'F{row}'] = f"=D{row}+E{row}"
        ws[f'G{row}'] = 12
        ws[f'H{row}'] = f"=TRUNC(F{row}/G{row})"
        ws[f'K{row}'] = f"=I{row}"
        ws[f'L{row}'] = f"=ROUND(I{row}/H{row},2)"
        ws[f'M{row}'] = f"=MAX(K{row}-3*H{row}, 0)"
        ws[f'N{row}'] = f"=H{row}*12 - M{row}"
        ws[f'O{row}'] = f"=N{row}"
        ws[f'P{row}'] = 12
        ws[f'Q{row}'] = f"=ROUND(O{row}/P{row}, 0)"
        ws[f'R{row}'] = f"=Q{row}"
        ws[f'S{row}'] = f"=R{row}*12"
        # (첫 번째 단계에서는 L열 빨간색 처리를 하지 않음)

    # 합계 행 추가
    sum_row = ws.max_row + 1
    for col in range(4, ws.max_column + 1):
        col_letter = get_column_letter(col)
        ws[f'{col_letter}{sum_row}'] = f"=SUM({col_letter}2:{col_letter}{ws.max_row})"

    # 컬럼 너비 자동 조정
    for column_cells in ws.columns:
        max_length = 0
        column = column_cells[0].column_letter
        for cell in column_cells:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        adjusted_width = (max_length + 2)
        ws.column_dimensions[column].width = adjusted_width
    
    wb.save(output_file)

    # 2) 두 번째 단계: 다시 열어서 (data_only=True) L열 값이 4 미만이면 빨간색 처리
    #    (주의: 엑셀에서 계산된 값이 있어야만 data_only=True로 읽을 때 값이 채워집니다.)
    wb2 = load_workbook(output_file, data_only=True)
    ws2 = wb2.active
    
    for row in range(2, ws2.max_row + 1):
        l_value = ws2.cell(row=row, column=12).value  # L열(12번째 열)의 실제 값
        if l_value is not None and isinstance(l_value, (int, float)) and l_value < 4:
            ws2.cell(row=row, column=12).fill = red_fill
    
    wb2.save(output_file)

    return output_file
