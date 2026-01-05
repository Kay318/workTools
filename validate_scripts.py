"""
엑셀 필터링 조건에 따라 파이썬 스크립트 자동 검증 프로그램
"""

import os
import re
import pandas as pd
from pathlib import Path
import xlwings as xw
import warnings
from typing import Dict, List, Any, Optional, Callable
warnings.filterwarnings('ignore')

class AdvancedScriptValidator:
    def __init__(self, excel_path: str, scripts_folder: str):
        """
        고급 스크립트 검증기 초기화
        
        Args:
            excel_path (str): 엑셀 파일 경로
            scripts_folder (str): 스크립트 폴더 경로
        """
        self.excel_path = excel_path
        self.scripts_folder = scripts_folder
        self.raw_data = None
        self.filtered_data = None
        self.expected_scripts = []
        self.found_scripts = []
        self.missing_scripts = []
        self.filters = []  # 필터 조건 저장
        self.script_name_pattern = "IBD_{category2}_{Level1}_{Level2}_{Level3}__.py"
        
    def read_excel_data(self) -> pd.DataFrame:
        """
        엑셀에서 모든 데이터 읽어오기
        """
        print("엑셀 파일 읽는 중...")
        
        try:
            # xlwings로 엑셀 열기
            app = xw.App(visible=False)
            wb = app.books.open(self.excel_path)
            sheet = wb.sheets[0]
            
            # 모든 데이터 읽기
            data_range = sheet.range('A1').expand('down')
            df = sheet.range('A1').options(pd.DataFrame, header=1, index=False, expand='table').value
            
            wb.close()
            app.quit()
            
        except Exception as e:
            print(f"xlwings로 엑셀 읽기 실패: {e}")
            print("openpyxl로 시도 중...")
            try:
                df = pd.read_excel(self.excel_path, engine='openpyxl')
            except Exception as e2:
                print(f"openpyxl로도 실패: {e2}")
                raise
        
        self.raw_data = df
        print(f"엑셀에서 {len(df)}행 데이터 읽음")
        print(f"컬럼 목록: {list(df.columns)}")
        
        return df
    
    def add_filter(self, column: str, condition: Any, description: str = ""):
        """
        필터 조건 추가
        
        Args:
            column (str): 필터링할 컬럼명
            condition: 필터 조건 (문자열, 리스트, 함수 등)
            description (str): 필터 설명
        """
        self.filters.append({
            'column': column,
            'condition': condition,
            'description': description
        })
        print(f"필터 추가: {column} -> {condition} ({description})")
    
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
    
    def add_custom_filter(self, filter_func: Callable, description: str = ""):
        """
        사용자 정의 필터 함수 추가
        """
        self.filters.append({
            'column': '__custom__',
            'condition': filter_func,
            'description': description or '사용자 정의 필터'
        })
        print(f"사용자 정의 필터 추가: {description}")
    
    def apply_filters(self) -> pd.DataFrame:
        """
        설정된 필터들을 적용하여 데이터 필터링
        """
        if self.raw_data is None:
            self.read_excel_data()
        
        df = self.raw_data.copy()
        filtered_rows = []
        
        print(f"\n{len(self.filters)}개의 필터 적용 중...")
        
        for i, filter_item in enumerate(self.filters, 1):
            column = filter_item['column']
            condition = filter_item['condition']
            description = filter_item['description']
            
            if column == '__custom__':
                # 사용자 정의 함수 필터
                mask = df.apply(condition, axis=1)
                before = len(df)
                df = df[mask]
                after = len(df)
                print(f"{i}. {description}: {before} → {after} 행 (제외: {before - after})")
                
            elif condition == 'notnull':
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
    
    def generate_expected_scripts(self) -> List[Dict]:
        """
        필터링된 데이터로부터 예상 스크립트명 생성
        """
        if self.filtered_data is None:
            self.apply_filters()
        
        df = self.filtered_data
        expected_scripts = []
        
        # 필수 컬럼 확인
        required_columns = ['category2', 'Level1', 'Level2', 'Level3']
        
        for col in required_columns:
            if col not in df.columns:
                print(f"경고: 필수 컬럼 '{col}'가 없습니다.")
                # 대체 컬럼 찾기
                possible_cols = [c for c in df.columns if col.lower() in c.lower()]
                if possible_cols:
                    print(f"  가능한 대체 컬럼: {possible_cols}")
                return []
        
        print("\n예상 스크립트명 생성 중...")
        
        for idx, row in df.iterrows():
            # 각 레벨 값 가져오기
            category2 = str(row['category2']).strip() if pd.notna(row['category2']) else ""
            level1 = str(row['Level1']).strip() if pd.notna(row['Level1']) else ""
            level2 = str(row['Level2']).strip() if pd.notna(row['Level2']) else ""
            level3 = str(row['Level3']).strip() if pd.notna(row['Level3']) else ""
            
            # 스크립트명 생성
            script_name = self._generate_script_name(category2, level1, level2, level3)
            
            if script_name:  # 유효한 스크립트명만 추가
                script_info = {
                    '엑셀_행': idx + 2,  # 1-based + 헤더
                    'category2': category2,
                    'Level1': level1,
                    'Level2': level2,
                    'Level3': level3,
                    '예상_스크립트명': script_name,
                    'key_name': f"{category2}_{level1}_{level2}_{level3}".rstrip('_')
                }
                
                # 추가 컬럼 정보도 포함
                for col in df.columns:
                    if col not in required_columns and col not in script_info:
                        script_info[col] = row[col]
                
                expected_scripts.append(script_info)
        
        self.expected_scripts = expected_scripts
        print(f"생성된 예상 스크립트: {len(expected_scripts)}개")
        
        return expected_scripts
    
    def _generate_script_name(self, category2: str, level1: str, level2: str, level3: str) -> str:
        """
        keyname으로부터 예상 스크립트명 생성
        """
        parts = []
        
        # 빈 값이 아닌 부분만 추가
        for part in [category2, level1, level2, level3]:
            if part and str(part).strip() and str(part).strip().lower() != 'nan':
                parts.append(str(part).strip())
        
        if parts:
            # 파일명 생성 (IBD_접두사, __접미사)
            script_name = f"IBD_{'_'.join(parts)}__.py"
            return script_name
        
        return ""
    
    def find_python_scripts(self) -> List[Dict]:
        """
        특정 폴더에서 모든 파이썬 스크립트 찾기
        """
        print(f"\n폴더 {self.scripts_folder} 에서 파이썬 스크립트 검색 중...")
        
        found_scripts = []
        
        # 모든 하위 폴더 탐색
        for root, dirs, files in os.walk(self.scripts_folder):
            for file in files:
                if file.endswith('.py'):
                    file_path = os.path.join(root, file)
                    rel_path = os.path.relpath(file_path, self.scripts_folder)
                    
                    # 파일명에서 keyname 추출 시도
                    key_parts = self._extract_key_from_filename(file)
                    
                    found_scripts.append({
                        'file_name': file,
                        'file_path': file_path,
                        'relative_path': rel_path,
                        'folder': root,
                        'key_parts': key_parts
                    })
        
        self.found_scripts = found_scripts
        print(f"총 {len(found_scripts)}개의 파이썬 스크립트 찾음")
        
        return found_scripts
    
    def _extract_key_from_filename(self, filename: str) -> Dict:
        """
        파일명에서 keyname 부분 추출
        """
        # IBD_로 시작하고 __.py로 끝나는 패턴
        pattern = r'^IBD_(.+?)__\.py$'
        match = re.match(pattern, filename)
        
        if match:
            key_str = match.group(1)
            parts = key_str.split('_')
            
            # 최대 4개의 파트로 분할
            result = {
                'full_key': key_str,
                'category2': parts[0] if len(parts) > 0 else "",
                'Level1': parts[1] if len(parts) > 1 else "",
                'Level2': parts[2] if len(parts) > 2 else "",
                'Level3': parts[3] if len(parts) > 3 else "",
            }
            return result
        
        return {'full_key': filename.replace('.py', '')}
    
    def validate_scripts(self) -> tuple:
        """
        필터링된 keyname과 실제 스크립트 비교 검증
        """
        if not self.expected_scripts:
            self.generate_expected_scripts()
        
        if not self.found_scripts:
            self.find_python_scripts()
        
        # 검증 결과 저장
        validation_results = []
        missing_scripts = []
        
        print("\n스크립트 검증 중...")
        
        # 각 예상 스크립트에 대해 검증
        for expected in self.expected_scripts:
            expected_name = expected['예상_스크립트명']
            found = False
            found_info = None
            
            # 실제 스크립트에서 찾기
            for script in self.found_scripts:
                if script['file_name'] == expected_name:
                    found = True
                    found_info = script
                    break
            
            # 결과 저장
            result = expected.copy()
            result['상태'] = '✓ 있음' if found else '✗ 없음'
            result['실제_경로'] = found_info['relative_path'] if found else ''
            result['파일_존재'] = '예' if found else '아니오'
            
            validation_results.append(result)
            
            if not found:
                missing_scripts.append(result)
        
        self.missing_scripts = missing_scripts
        
        # 통계 계산
        total = len(validation_results)
        found_count = len([r for r in validation_results if r['파일_존재'] == '예'])
        missing_count = len(missing_scripts)
        
        print(f"검증 완료: {found_count}/{total} (누락: {missing_count})")
        
        return validation_results, missing_scripts
    
    def generate_report(self, output_path: str = None) -> str:
        """
        검증 결과 리포트 생성
        """
        if output_path is None:
            # 기본 리포트 파일명
            base_name = os.path.splitext(os.path.basename(self.excel_path))[0]
            timestamp = pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')
            output_path = f"스크립트_검증_결과_{base_name}_{timestamp}.xlsx"
        
        validation_results, missing_scripts = self.validate_scripts()
        
        # DataFrame으로 변환
        df_results = pd.DataFrame(validation_results)
        df_missing = pd.DataFrame(missing_scripts)
        
        # ExcelWriter로 여러 시트에 저장
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # 1. 전체 결과
            df_results.to_excel(writer, sheet_name='전체_검증결과', index=False)
            
            # 2. 누락된 스크립트
            if not df_missing.empty:
                df_missing.to_excel(writer, sheet_name='누락_스크립트', index=False)
            else:
                # 빈 시트 생성
                pd.DataFrame(['모든 스크립트가 존재합니다.']).to_excel(
                    writer, sheet_name='누락_스크립트', index=False, header=False
                )
            
            # 3. 요약 통계
            self._create_summary_sheet(writer, validation_results, missing_scripts)
            
            # 4. 필터 조건 시트
            self._create_filter_sheet(writer)
            
            # 5. 발견된 스크립트 목록
            self._create_found_scripts_sheet(writer)
        
        print(f"\n검증 결과 리포트 저장: {output_path}")
        
        # 콘솔에 요약 출력
        self._print_summary(validation_results, missing_scripts)
        
        return output_path
    
    def _create_summary_sheet(self, writer, validation_results, missing_scripts):
        """요약 통계 시트 생성"""
        total = len(validation_results)
        found = len([r for r in validation_results if r['파일_존재'] == '예'])
        missing = len(missing_scripts)
        
        summary_data = {
            '구분': ['총 검증 대상', '스크립트 있음', '스크립트 없음', '검증률'],
            '값': [total, found, missing, f"{found/total*100:.1f}%" if total > 0 else "0%"]
        }
        
        df_summary = pd.DataFrame(summary_data)
        df_summary.to_excel(writer, sheet_name='요약_통계', index=False)
    
    def _create_filter_sheet(self, writer):
        """적용된 필터 조건 시트 생성"""
        filter_data = []
        
        for i, f in enumerate(self.filters, 1):
            filter_data.append({
                '번호': i,
                '컬럼': f['column'],
                '조건': str(f['condition']),
                '설명': f['description']
            })
        
        if filter_data:
            df_filters = pd.DataFrame(filter_data)
            df_filters.to_excel(writer, sheet_name='적용_필터', index=False)
        else:
            pd.DataFrame(['적용된 필터가 없습니다.']).to_excel(
                writer, sheet_name='적용_필터', index=False, header=False
            )
    
    def _create_found_scripts_sheet(self, writer):
        """발견된 스크립트 목록 시트 생성"""
        script_data = []
        
        for script in self.found_scripts:
            script_data.append({
                '파일명': script['file_name'],
                '상대경로': script['relative_path'],
                '폴더': script['folder']
            })
        
        if script_data:
            df_scripts = pd.DataFrame(script_data)
            df_scripts.to_excel(writer, sheet_name='발견_스크립트', index=False)
    
    def _print_summary(self, validation_results, missing_scripts):
        """
        검증 결과 요약 출력
        """
        print("\n" + "="*70)
        print("스크립트 검증 결과 요약")
        print("="*70)
        
        total = len(validation_results)
        found = len([r for r in validation_results if r['파일_존재'] == '예'])
        missing = len(missing_scripts)
        
        print(f"총 검증 대상: {total}개")
        print(f"스크립트 있음: {found}개")
        print(f"스크립트 없음: {missing}개")
        print(f"검증률: {found/total*100:.1f}%" if total > 0 else "0%")
        
        if missing_scripts:
            print(f"\n[상위 10개 누락 스크립트]")
            for i, missing in enumerate(missing_scripts[:10], 1):
                print(f"{i:3d}. {missing['예상_스크립트명']}")
            
            if len(missing_scripts) > 10:
                print(f"... 외 {len(missing_scripts)-10}개 더")
        
        # 카테고리별 통계
        if validation_results:
            print("\n[카테고리별 통계]")
            categories = {}
            for result in validation_results:
                cat = result.get('category2', 'Unknown')
                if cat not in categories:
                    categories[cat] = {'total': 0, 'found': 0}
                categories[cat]['total'] += 1
                if result['파일_존재'] == '예':
                    categories[cat]['found'] += 1
            
            for cat, stats in sorted(categories.items()):
                if stats['total'] > 0:
                    found_rate = stats['found']/stats['total']*100
                    print(f"  {cat}: {stats['found']}/{stats['total']} ({found_rate:.1f}%)")
        
        print("="*70)
    
    def find_similar_scripts(self, threshold: float = 0.6):
        """
        유사한 파일명 찾기
        """
        if not self.missing_scripts:
            self.validate_scripts()
        
        suggestions = []
        
        for missing in self.missing_scripts:
            expected_name = missing['예상_스크립트명'].lower().replace('.py', '')
            similar_files = []
            
            for script in self.found_scripts:
                script_name = script['file_name'].lower().replace('.py', '')
                
                # 유사도 계산
                similarity = self._calculate_similarity(expected_name, script_name)
                
                if similarity >= threshold:
                    similar_files.append({
                        'file_name': script['file_name'],
                        'similarity': similarity,
                        'path': script['relative_path']
                    })
            
            if similar_files:
                # 유사도 순 정렬
                similar_files.sort(key=lambda x: x['similarity'], reverse=True)
                suggestions.append({
                    'expected': missing['예상_스크립트명'],
                    'key_name': missing.get('key_name', ''),
                    'similar_files': similar_files[:3]
                })
        
        return suggestions
    
    def _calculate_similarity(self, str1: str, str2: str) -> float:
        """
        두 문자열 간의 유사도 계산
        """
        # 간단한 Jaccard 유사도
        set1 = set(str1.split('_'))
        set2 = set(str2.split('_'))
        
        if not set1 or not set2:
            return 0.0
        
        intersection = len(set1.intersection(set2))
        union = len(set1.union(set2))
        
        return intersection / union if union > 0 else 0.0


# 사용 예시 함수들
def example_usage():
    """
    사용 예시
    """
    # 1. 검증기 생성
    validator = AdvancedScriptValidator(
        excel_path="path/to/your/excel.xlsx",
        scripts_folder="path/to/scripts/folder"
    )
    
    # 2. 필터 조건 추가 (예시)
    validator.add_equal_filter("ccIC", "O", "ccIC이 O인 항목")
    validator.add_equal_filter("미국", "O", "미국이 O인 항목")
    validator.add_in_filter("자동화 매뉴얼 구분", ["자동화", "통화"], "자동화 또는 통화")
    
    # 3. 사용자 정의 필터 예시
    def custom_filter(row):
        """복잡한 조건의 사용자 정의 필터"""
        # 예: 특정 값이 포함된 경우
        if pd.notna(row.get('비고')) and '중요' in str(row.get('비고')):
            return True
        return False
    
    validator.add_custom_filter(custom_filter, "비고에 '중요'가 포함된 항목")
    
    # 4. 검증 실행
    report_path = validator.generate_report()
    
    print(f"\n검증 완료! 리포트: {report_path}")
    
    # 5. 유사 파일 검색
    suggestions = validator.find_similar_scripts(threshold=0.5)
    if suggestions:
        print(f"\n유사한 파일 {len(suggestions)}개 발견")
        for sug in suggestions[:5]:
            print(f"\n예상: {sug['expected']}")
            for sim in sug['similar_files']:
                print(f"  유사: {sim['file_name']} ({sim['similarity']:.1%})")


def interactive_mode():
    """
    대화형 모드 실행
    """
    print("="*70)
    print("엑셀 필터링 기반 스크립트 검증 프로그램")
    print("="*70)
    
    # 입력 받기
    excel_path = input("엑셀 파일 경로: ").strip()
    scripts_folder = input("스크립트 폴더 경로: ").strip()
    
    if not os.path.exists(excel_path):
        print(f"에러: 파일 없음 - {excel_path}")
        return
    
    if not os.path.exists(scripts_folder):
        print(f"에러: 폴더 없음 - {scripts_folder}")
        return
    
    # 검증기 생성
    validator = AdvancedScriptValidator(excel_path, scripts_folder)
    
    # 엑셀 데이터 읽기
    df = validator.read_excel_data()
    
    # 컬럼 목록 표시
    print(f"\n[엑셀 컬럼 목록]")
    for i, col in enumerate(df.columns, 1):
        print(f"{i:2d}. {col}")
    
    # 필터 설정
    print("\n[필터 설정]")
    print("필터를 추가하시겠습니까? (y/n): ", end="")
    if input().lower() == 'y':
        while True:
            print("\n필터 유형 선택:")
            print("1. 특정 값 일치 (column = value)")
            print("2. 여러 값 중 하나 (column in [value1, value2, ...])")
            print("3. null이 아닌 값")
            print("4. 사용자 정의 함수")
            print("5. 필터 설정 완료")
            
            choice = input("선택 (1-5): ").strip()
            
            if choice == '5':
                break
            
            if choice == '1':
                col = input("컬럼명: ").strip()
                value = input("값: ").strip()
                desc = input("설명 (선택사항): ").strip()
                validator.add_equal_filter(col, value, desc or f"{col} = {value}")
                
            elif choice == '2':
                col = input("컬럼명: ").strip()
                values = input("값들 (쉼표로 구분): ").strip().split(',')
                values = [v.strip() for v in values]
                desc = input("설명 (선택사항): ").strip()
                validator.add_in_filter(col, values, desc or f"{col} in {values}")
                
            elif choice == '3':
                col = input("컬럼명: ").strip()
                desc = input("설명 (선택사항): ").strip()
                validator.add_notnull_filter(col, desc or f"{col} is not null")
                
            elif choice == '4':
                print("사용자 정의 함수는 코드 내에서만 추가 가능합니다.")
                print("프로그램 코드를 수정하시거나 다음 옵션을 사용해주세요.")
    
    # 검증 실행
    print("\n검증을 시작합니다...")
    report_path = validator.generate_report()
    
    print(f"\n✓ 검증 완료!")
    print(f"리포트 파일: {report_path}")
    
    # 유사 파일 검색
    print("\n유사 파일을 검색하시겠습니까? (y/n): ", end="")
    if input().lower() == 'y':
        suggestions = validator.find_similar_scripts(threshold=0.5)
        if suggestions:
            print(f"\n[유사 파일 {len(suggestions)}개 발견]")
            for i, sug in enumerate(suggestions[:10], 1):
                print(f"\n{i}. 예상: {sug['expected']}")
                for sim in sug['similar_files']:
                    print(f"   - {sim['file_name']} ({sim['similarity']:.1%})")


if __name__ == "__main__":
    # 대화형 모드 실행
    interactive_mode()
    
    # 또는 직접 설정
    # example_usage()
