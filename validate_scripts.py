"""
복잡한 스크립트 패턴 매칭 검증 프로그램
- 한 엑셀 행에 여러 스크립트 가능
- Level3 생략 가능
- .py 앞에 숫자 있을 수 있음
"""

import os
import re
import pandas as pd
import xlwings as xw
from typing import List, Dict, Any

class ComplexScriptValidator:
    def __init__(self, excel_path: str, 
                 sheet_name: str, 
                 excel_range: str, 
                 scripts_folder: str):
        """
        복잡한 패턴 검증기 초기화
        """
        self.excel_path = excel_path
        self.sheet_name = sheet_name
        self.excel_range = excel_range
        self.scripts_folder = scripts_folder
        self.raw_data = None
        self.filtered_data = None
        self.filters = []
        self.expected_scripts = []  # 각 행별 예상 스크립트 리스트
        self.found_scripts = {}     # 실제 찾은 스크립트
        self.matched_results = []   # 매칭 결과
        
    def read_excel_data(self) -> pd.DataFrame:
        """
        엑셀 데이터 읽기 (2행이 헤더)
        """
        print("엑셀 파일 읽는 중...")
        
        try:
            app = xw.App(visible=False)
            wb = app.books.open(self.excel_path)
            sheet = wb.sheets[self.sheet_name]
            
            # 모든 데이터 읽기
            df = sheet.range(self.excel_range).options(pd.DataFrame, header=1, index=False).value
            # 동일한 명칭이 두개 있을 때 마지막 컬럼을 사용함
            df = df.loc[:, ~df.columns.duplicated(keep='last')]
            
            wb.close()
            app.quit()
            
        except Exception as e:
            print(f"xlwings 실패: {e}")
            try:
                df = pd.read_excel(self.excel_path, engine='openpyxl')
            except Exception as e2:
                print(f"openpyxl 실패: {e2}")
                raise
        
        self.raw_data = df
        print(f"읽은 데이터: {len(df)}행, 컬럼: {list(df.columns)[:10]}...")
        return df
    
    def add_equal_filter(self, column: str, value: Any, description: str = ""):
        """
        특정 값과 일치하는 조건 추가
        """
        self.add_filter(column, value, f"{description or f'{column} = {value}'}")
    
    def add_in_filter(self, column: str, values: list, description: str = ""):
        """
        특정 값들 중 하나인 조건 추가
        """
        self.add_filter(column, values, f"{description or f'{column} in {values}'}")
    
    def add_notnull_filter(self, column: str, description: str = ""):
        """
        값이 null이 아닌 조건 추가
        """
        self.add_filter(column, 'notnull', f"{description or f'{column} is not null'}")

    def add_not_equal_filter(self, column: str, value: Any, description: str = ""):
        """
        특정 값이 아닌 조건 추가
        """
        self.add_filter(column, lambda x: x!= value, 
                        f"{description or f'{column} != {value}'}")
        
    def add_not_in_filter(self, column: str, values: list, description: str = ""):
        """
        특정 값들에 포함되지 않는 조건 추가
        """
        self.add_filter(column, lambda x: x not in values,
                        f"{description or f'{column} not in {values}'}")
    
    def add_filter(self, column: str, condition, description: str = ""):
        """필터 추가"""
        self.filters.append({
            'column': column,
            'condition': condition,
            'description': description
        })
    
    def apply_filters(self) -> pd.DataFrame:
        """
        설정된 필터들을 적용하여 데이터 필터링
        """
        if self.raw_data is None:
            self.read_excel_data()
        
        df = self.raw_data.copy()
        
        print(f"\n{len(self.filters)}개의 필터 적용 중...")
        
        for i, filter_item in enumerate(self.filters, 1):
            column = filter_item['column']
            condition = filter_item['condition']
            description = filter_item['description']
                
            if condition == 'notnull':
                # null이 아닌 행
                before = len(df)
                df = df[df[column].notna()]
                after = len(df)
                print(f"{i}. {description}: {before} → {after} 행 (제외: {before - after})")
                
            elif isinstance(condition, list):
                # 리스트 값 중 하나인 경우
                before = len(df)
                df = df[df[column].isin(condition)]
                after = len(df)
                print(f"{i}. {description}: {before} → {after} 행 (제외: {before - after})")
                
            elif callable(condition):
                # 함수 조건
                before = len(df)
                df = df[df[column].apply(condition)]
                after = len(df)
                print(f"{i}. {description}: {before} → {after} 행 (제외: {before - after})")
                
            else:
                # 단일 값과 일치하는 경우
                before = len(df)
                df = df[df[column] == condition]
                after = len(df)
                print(f"{i}. {description}: {before} → {after} 행 (제외: {before - after})")
        
        self.filtered_data = df
        print(f"필터링 완료: {len(self.raw_data)} → {len(df)} 행")
        
        return df
    
    def generate_expected_script_patterns(self, row: pd.Series) -> str:
        """
        스크립트 패턴 생성
        """
        
        # 기본 값 가져오기
        category = str(row['category2']).strip() if pd.notna(row.get('category2')) else ""
        level1 = str(row['Level1']).strip() if pd.notna(row.get('Level1')) else ""
        level2 = str(row['Level2']).strip() if pd.notna(row.get('Level2')) else ""
        level3 = str(row['Level3']).strip() if pd.notna(row.get('Level3')) else ""
        
        if not category or not level1:
            return ""  # 최소 category2와 Level1은 필요
        
        # 모든 조합 생성
        parts = [category, level1]

        if level2:
            parts.append(level2)

        if level3:
            parts.append(level3)

        pattern = f"IBD_{'_'.join(parts)}"

        return pattern
    
    def find_all_scripts(self) -> Dict[str, Dict]:
        """
        모든 스크립트 찾기 (파일명 → 정보 매핑)
        """
        script_dict = {}
        
        for root, dirs, files in os.walk(self.scripts_folder):
            for file in files:
                if file.endswith('.py'):
                    file_path = os.path.join(root, file)
                    rel_path = os.path.relpath(file_path, self.scripts_folder)
                    
                    # 파일명에서 정보 추출
                    file_info = {'raw_name': file.replace('.py', '')}
                    
                    script_dict[file] = {
                        'file_name': file,
                        'file_path': file_path,
                        'relative_path': rel_path,
                        'folder': root,
                        'extracted_info': file_info
                    }
        
        self.found_scripts = script_dict
        print(f"찾은 스크립트: {len(script_dict)}개")
        return script_dict
    
    def validate_all_rows(self) -> List[Dict]:
        """
        모든 행 검증 실행
        """
        if self.filtered_data is None:
            self.apply_filters()
        
        if not self.found_scripts:
            self.find_all_scripts()
        
        print("\n행별 스크립트 검증 시작...")
        
        results = []
        for idx, row in self.filtered_data.iterrows():
            excel_row_num = idx + 3  # 엑셀 행 번호 (헤더 2행 + 1)
            
            # 이 행의 스크립트 패턴
            pattern = self.generate_expected_script_patterns(row)

            found_for_row = []
            for script_name in self.found_scripts.keys():
                if re.search(pattern, script_name):
                    found_for_row.append({
                        'script_name': script_name,
                        'path': self.found_scripts[script_name]['relative_path']
                    })
            
            # 결과 저장
            row_result = {
                '엑셀_행': excel_row_num,
                'category2': row.get('category2', ''),
                'Level1': row.get('Level1', ''),
                'Level2': row.get('Level2', ''),
                'Level3': row.get('Level3', ''),
                '예상_패턴_수': len(pattern),
                '발견_스크립트_수': len(found_for_row),
                '모두_있음': len(found_for_row) >= 1,  # 최소 하나라도 있으면 OK
                '예상_패턴들': pattern,
                '발견_스크립트들': found_for_row,
                '상태': '✓ 있음' if found_for_row else '✗ 없음'
            }
            
            # 추가 컬럼 정보
            for col in ['ccIC', region, '자동화 매뉴얼 구분']:
                if col in row:
                    row_result[col] = row[col]
            
            results.append(row_result)
        
        self.matched_results = results
        
        # 통계
        total_rows = len(results)
        rows_with_script = len([r for r in results if r['발견_스크립트_수'] > 0])
        total_found = sum(r['발견_스크립트_수'] for r in results)
        
        print(f"\n검증 완료:")
        print(f"  - 검증 행: {total_rows}")
        print(f"  - 스크립트 있는 행: {rows_with_script}/{total_rows} ({rows_with_script/total_rows*100:.1f}%)")
        print(f"  - 총 발견 스크립트: {total_found}")
        
        return results
    
    def generate_detailed_report(self, output_path: str = None) -> str:
        """
        상세 리포트 생성
        """
        if output_path is None:
            base_name = os.path.splitext(os.path.basename(self.excel_path))[0]
            timestamp = pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')
            output_path = f"상세_검증_결과_{base_name}_{timestamp}.xlsx"
        
        results = self.validate_all_rows() if not self.matched_results else self.matched_results
        
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # 1. 요약 시트
            self._create_summary_sheet(writer, results)
            
            # 2. 행별 상세 결과
            self._create_row_detail_sheet(writer, results)
            
            # 3. 누락된 항목
            self._create_missing_sheet(writer, results)
            
            # 4. 발견된 스크립트 전체 목록
            self._create_all_scripts_sheet(writer)
            
            # 5. 필터 조건
            self._create_filter_sheet(writer)
        
        print(f"\n리포트 생성 완료: {output_path}")
        return output_path
    
    def _create_summary_sheet(self, writer, results):
        """요약 시트 생성"""
        summary_data = []
        
        # 기본 통계
        total_rows = len(results)
        rows_with_script = len([r for r in results if r['발견_스크립트_수'] > 0])
        rows_without_script = total_rows - rows_with_script
        
        summary_data.append(['총 검증 행 수', total_rows])
        summary_data.append(['스크립트 있는 행', rows_with_script])
        summary_data.append(['스크립트 없는 행', rows_without_script])
        summary_data.append(['검증률', f'{rows_with_script/total_rows*100:.1f}%'])
        
        # 카테고리별 통계
        categories = {}
        for r in results:
            cat = r.get('category2', 'Unknown')
            if cat not in categories:
                categories[cat] = {'total': 0, 'with_script': 0}
            categories[cat]['total'] += 1
            if r['발견_스크립트_수'] > 0:
                categories[cat]['with_script'] += 1
        
        summary_data.append(['', ''])
        summary_data.append(['카테고리별 통계', ''])
        for cat, stats in sorted(categories.items()):
            if stats['total'] > 0:
                rate = stats['with_script']/stats['total']*100
                summary_data.append([f"  {cat}", f"{stats['with_script']}/{stats['total']} ({rate:.1f}%)"])
        
        df_summary = pd.DataFrame(summary_data, columns=['구분', '값'])
        df_summary.to_excel(writer, sheet_name='요약', index=False)
    
    def _create_row_detail_sheet(self, writer, results):
        """행별 상세 결과"""
        detail_data = []
        
        for r in results:            
            # 발견 스크립트들을 문자열로
            found_str = '\n'.join([f"{s['script_name']} ({s['path']})" 
                                  for s in r['발견_스크립트들']]) if r['발견_스크립트들'] else ''
            
            detail_data.append({
                '행번호': r['엑셀_행'],
                'category2': r['category2'],
                'Level1': r['Level1'],
                'Level2': r.get('Level2', ''),
                'Level3': r.get('Level3', ''),
                'ccIC': r.get('ccIC', ''),
                region: r.get(region, ''),
                '구분': r.get('자동화 매뉴얼 구분', ''),
                '발견스크립트수': r['발견_스크립트_수'],
                '상태': r['상태'],
                '발견_스크립트들': found_str
            })
        
        df_detail = pd.DataFrame(detail_data)
        df_detail.to_excel(writer, sheet_name='행별_상세', index=False)
    
    def _create_missing_sheet(self, writer, results):
        """누락된 항목"""
        missing_data = []
        
        for r in results:
            if r['발견_스크립트_수'] == 0:  # 스크립트 하나도 없으면
                missing_data.append({
                    '행번호': r['엑셀_행'],
                    'category2': r['category2'],
                    'Level1': r['Level1'],
                    'Level2': r.get('Level2', ''),
                    'Level3': r.get('Level3', ''),
                    'ccIC': r.get('ccIC', ''),
                    region: r.get(region, ''),
                    '구분': r.get('자동화 매뉴얼 구분', '')
                })
        
        if missing_data:
            df_missing = pd.DataFrame(missing_data)
            df_missing.to_excel(writer, sheet_name='누락_항목', index=False)
        else:
            pd.DataFrame(['모든 행에 스크립트가 있습니다.']).to_excel(
                writer, sheet_name='누락_항목', index=False, header=False
            )
    
    def _create_all_scripts_sheet(self, writer):
        """모든 스크립트 목록"""
        script_data = []
        
        for script_name, info in self.found_scripts.items():
            script_data.append({
                # '파일명': script_name,
                '경로': info['relative_path'],
                # '폴더': info['folder'],
            })
        
        if script_data:
            df_scripts = pd.DataFrame(script_data)
            df_scripts.to_excel(writer, sheet_name='전체_스크립트', index=False)
    
    def _create_filter_sheet(self, writer):
        """필터 조건"""
        filter_data = []
        
        for i, f in enumerate(self.filters, 1):
            filter_data.append({
                '번호': i,
                '컬럼': f['column'],
                '조건': f['description'],
            })
        
        if filter_data:
            df_filters = pd.DataFrame(filter_data)
            df_filters.to_excel(writer, sheet_name='적용_필터', index=False)


# 사용 예시
if __name__ == "__main__":
    # 검증기 생성
    """
    region: [국내, 유럽, 미국, 호주/뉴질랜드, 캐나다, 중국, 러시아, 브라질, 인도네시아, 일본, 인도, 중동] 중 하나
    excel_path: 성적서 경로
    sheet_name: 검증하려는 Sheet이름
    excel_range: 검증하려는 셀 범위
    scripts_folder: 스크립트 폴더
    """
    region = "유럽"
    excel_path=r"D:\IBD\files\업무보고\성적서\유럽 26년 1차\(유럽 ccIC) 인포.Big DATA_수집 항목_v1.25.2_'25년2차_r11_자동화매뉴얼구분.xlsx"
    sheet_name = "NAVI"
    excel_range = "C2:BS671"
    scripts_folder=r"D:\NOVA\com.auto.script\Test Script\IBD\EUR\ccIC"

    validator = ComplexScriptValidator(
        excel_path=excel_path,
        sheet_name=sheet_name,
        excel_range=excel_range,
        scripts_folder=scripts_folder,
    )

    """
    add_equal_filter: 특정 값과 일치하는 조건 추가
    add_not_equal_filter: 특정 값이 아닌 조건 추가
    add_in_filter: 특정 값들 중 하나인 조건 추가
    add_not_in_filter: 특정 값들에 포함되지 않는 조건 추가
    add_notnull_filter: 값이 null이 아닌 조건 추가
    """

    validator.add_not_equal_filter("ccIC", "X", "ccIC가 X가 아닌 것")
    validator.add_not_equal_filter(region, "X", f"{region} X가 아닌 것")
    validator.add_not_equal_filter("Test 결과", "N/A", "테스트 결과가 N/A 아닌 것")
    validator.add_in_filter("자동화\n매뉴얼 구분", ["자동화", "통합"], "자동화 또는 통합인 항목만")

    # 검증 실행 및 리포트 생성
    report_path = validator.generate_detailed_report()
    print(f"검증 완료! 리포트: {report_path}")

