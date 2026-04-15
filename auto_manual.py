import xlwings as xw
import os

def process_excel_files(file1_path, file2_path):
    """
    엑셀 파일 2개를 비교하여 데이터를 복사하는 함수
    
    Args:
        file1_path (str): 참조용 엑셀 파일 경로
        file2_path (str): 수정할 엑셀 파일 경로
    """
    
    # 파일 존재 여부 확인
    if not os.path.exists(file1_path):
        print(f"엑셀1 파일을 찾을 수 없습니다: {file1_path}")
        return
    if not os.path.exists(file2_path):
        print(f"엑셀2 파일을 찾을 수 없습니다: {file2_path}")
        return
    
    try:
        # 엑셀 어플리케이션 실행 (보이지 않게 설정)
        app = xw.App(visible=False)
        
        # 엑셀 파일 열기
        wb1 = app.books.open(file1_path)
        wb2 = app.books.open(file2_path)
        
        # 시트 선택 (첫 번째 시트를 사용합니다. 필요에 따라 수정 가능)
        ws1 = wb1.sheets[0]
        ws2 = wb2.sheets[0]
        
        # 데이터 범위 확인
        # 엑셀2의 마지막 행 찾기
        last_row_ws2 = ws2.range('A' + str(ws2.cells.last_cell.row)).end('up').row
        
        print(f"엑셀2 처리할 총 행 수: {last_row_ws2 - 2}행 (3행부터 {last_row_ws2}행까지)")
        
        # 엑셀1의 데이터를 미리 읽어서 딕셔너리로 저장 (성능 최적화)
        # E~I 컬럼을 키로, AU 컬럼 값을 값으로 저장
        ref_data = {}
        
        # 엑셀1의 마지막 행 찾기
        last_row_ws1 = ws1.range('A' + str(ws1.cells.last_cell.row)).end('up').row
        
        print(f"엑셀1 참조 데이터 수: {last_row_ws1 - 1}행")
        
        # 엑셀1 데이터 읽기
        for row in range(2, last_row_ws1 + 1):  # 1행은 헤더라고 가정
            # E~I 컬럼 값 읽기
            e_col = ws1.range(f'E{row}').value
            f_col = ws1.range(f'F{row}').value
            g_col = ws1.range(f'G{row}').value
            h_col = ws1.range(f'H{row}').value
            i_col = ws1.range(f'I{row}').value
            
            # AU 컬럼 값 읽기
            au_col = ws1.range(f'AU{row}').value
            
            # 키 생성 (튜플로 E~I 컬럼 값을 조합)
            key = (e_col, f_col, g_col, h_col, i_col)
            
            # 딕셔너리에 저장
            if key not in ref_data:
                ref_data[key] = au_col
        
        print(f"참조 데이터 딕셔너리 크기: {len(ref_data)}")
        
        # 엑셀2 데이터 처리 (3행부터 시작)
        processed_count = 0
        copied_count = 0
        
        for row in range(3, last_row_ws2 + 1):
            # 엑셀2의 E~I 컬럼 값 읽기
            e2_col = ws2.range(f'E{row}').value
            f2_col = ws2.range(f'F{row}').value
            g2_col = ws2.range(f'G{row}').value
            h2_col = ws2.range(f'H{row}').value
            i2_col = ws2.range(f'I{row}').value
            
            # 키 생성
            key = (e2_col, f2_col, g2_col, h2_col, i2_col)
            
            # 참조 데이터에서 찾기
            if key in ref_data:
                # AU 컬럼 값 복사
                au_value = ref_data[key]
                ws2.range(f'AU{row}').value = au_value
                copied_count += 1
                
                if copied_count <= 5:  # 처음 5개만 예시 출력
                    print(f"행 {row}: 데이터 복사 완료 - {au_value}")
            
            processed_count += 1
            
            # 진행 상황 표시 (100행마다)
            if processed_count % 100 == 0:
                print(f"처리 중... {processed_count}행 완료")
        
        # 파일 저장
        wb2.save()
        print(f"\n처리 완료!")
        print(f"총 처리 행 수: {processed_count}")
        print(f"복사된 데이터 수: {copied_count}")
        
        # "자동화\n매뉴얼 구분" 컬럼 처리 확인 메시지
        print("\n'자동화\\n매뉴얼 구분' 컬럼(AU열)에 데이터가 복사되었습니다.")
        
    except Exception as e:
        print(f"오류 발생: {str(e)}")
        import traceback
        traceback.print_exc()
        
    finally:
        # 워크북 닫기
        try:
            wb1.close()
        except:
            pass
        try:
            wb2.close()
        except:
            pass
        
        # 엑셀 어플리케이션 종료
        try:
            app.quit()
        except:
            pass

def main():
    """메인 함수"""
    print("=== Excel 데이터 복사 프로그램 ===")
    print("설명:")
    print("1. 엑셀1과 엑셀2의 E~I 컬럼 데이터를 비교합니다.")
    print("2. 동일한 데이터가 있으면 엑셀1의 AU컬럼 값을 엑셀2의 AU컬럼으로 복사합니다.")
    print("3. 엑셀2의 3행부터 아래로 처리합니다.")
    print()
    
    # 파일 경로 설정 (실제 경로로 수정 필요)
    excel1_path = input("엑셀1 파일 경로를 입력하세요 (참조용): ").strip()
    excel2_path = input("엑셀2 파일 경로를 입력하세요 (수정할 파일): ").strip()
    
    # 경로에 따옴표가 있으면 제거
    excel1_path = excel1_path.strip('"\'')
    excel2_path = excel2_path.strip('"\'')
    
    print("\n처리를 시작합니다...")
    process_excel_files(excel1_path, excel2_path)
    
    input("\n완료되었습니다. 엔터 키를 누르면 종료합니다.")

if __name__ == "__main__":
    main()
