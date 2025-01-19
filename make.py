import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from utils import match_wine_names

def engine(df_off, df_on, df_remain):
    # 매칭 및 데이터 처리
    off_array = list(df_off['상 품 명'].values[~pd.isnull(df_off['상 품 명'].values)])
    on_array = list(df_on['품목명'].values[~pd.isnull(df_on['품목명'].values)])
    remain_array = list(df_remain[df_remain.columns[1]].values[~pd.isnull(df_remain[df_remain.columns[1]].values)])[1:]

    # 결과 DataFrame 생성
    columns = [
        'No.', '상품소분류', '상 품 명', 'OFF 수량', 'ON 수량', 'TOTAL 수량', '기간 개월수', '월판매량 수량',
        '현재고 수량', '발주분 수량', '총재고 수량', '재고월수', '3개월후 재고', '발주희망', '발주수량 병',
        '발주수량 박스단위', '예상발주 박스', '발주확정 박스', '발주확정 병'
    ]
    df = pd.DataFrame(columns=columns)

    df['상 품 명'] = off_array
    df['OFF 수량'] = df_off['합 계'].values[1:-2]

    # on 수량 dictionalry로 만들기
    on_dict = {df_on['품목명'][i]: df_on['EA'][i] for i in range(len(df_on)) if df_on['품목명'][i] in on_array}
    
    result_on_off = match_wine_names(off_array, on_array)
    result_off_remain = match_wine_names(off_array, remain_array)
    # remain 수량 dictionary
    remain_dict = {df_remain[df_remain.columns[1]][i]: df_remain[df_remain.columns[9]][i] for i in range(len(df_remain)) if df_remain[df_remain.columns[1]][i] in remain_array}
    
    df['ON 수량'] = [on_dict[result_on_off[name]] if name in result_on_off else 0 for name in off_array]
    df['현재고 수량'] = [remain_dict[result_off_remain[name]] if name in result_off_remain else 0 for name in off_array]

    output_file = 'result_with_formulas.xlsx'
    df.to_excel(output_file, index=False, engine='openpyxl')

    # 엑셀 파일에 수식 및 서식 적용
    wb = load_workbook(output_file)
    ws = wb.active

    # 수식 및 서식 추가
    for row in range(2, ws.max_row + 1):
        ws[f'F{row}'] = f"=D{row}+E{row}"
        ws[f'G{row}'] = 12
        ws[f'H{row}'] = f"=TRUNC(F{row}/G{row})"
        ws[f'K{row}'] = f"=I{row}"
        ws[f'L{row}'] = f"=ROUND(I{row}/H{row},2)"
        ws[f'M{row}'] = f"=K{row}-3*H{row}"
        ws[f'N{row}'] = f"=H{row}*12 - M{row}"
        ws[f'O{row}'] = f"=N{row}"
        ws[f'P{row}'] = 12
        ws[f'Q{row}'] = f"=ROUND(O{row}/P{row}, 0)"
        ws[f'R{row}'] = f"=Q{row}"
        ws[f'S{row}'] = f"=R{row}*12"
    
    sum_row = ws.max_row + 1
    for col in range(4, ws.max_column + 1):
        col_letter = get_column_letter(col)
        ws[f'{col_letter}{sum_row}'] = f"=SUM({col_letter}2:{col_letter}{ws.max_row})"
    # 컬럼 너비 자동 조정
    for column_cells in ws.columns:
        max_length = 0
        column = column_cells[0].column_letter  # 해당 컬럼 이름(A, B, C...)
        for cell in column_cells:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        adjusted_width = (max_length + 2)
        ws.column_dimensions[column].width = adjusted_width
    
    wb.save(output_file)
    return output_file
