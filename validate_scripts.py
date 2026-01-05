"""
엑셀의 keyname들과 파이썬 스크립트 파일 매칭 검증 프로그램
"""

import os
import re
import pandas as pd
from pathlib import Path
import xlwings as xw
import warnings
warnings.filterwarnings('ignore')

class ScriptValidator:
    def __init__(self, excel_path, scripts_folder):
        """
        스크립트 검증기 초기화
        
        Args:
            excel_path (str): 엑셀 파일 경로
            scripts_folder (str): 스크립트 폴더 경로
        """
        self.excel_path = excel_path
        self.scripts_folder = scripts_folder
        self.expected_scripts = []
        self.found_scripts = []
        self.missing_scripts = []
        
    def read_excel_keys(self):
        """
        엑셀에서 keyname들 읽어오기
        """
        print("엑셀 파일에서 keyname 읽는 중...")
        
        try:
            # xlwings로 엑셀 열기
            app = xw.App(visible=False)
            wb = app.books.open(self.excel_path)
            sheet = wb.sheets[0]
            
            # 데이터 읽기 (A:D 열 - category2, Level1, Level2, Level3)
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
            
        # 필요한 열 확인
        required_columns = ['category2', 'Level1', 'Level2', 'Level3']
        missing_cols = [col for col in required_columns if col not in df.columns]
        
        if missing_cols:
            print(f"경고: 다음 열이 없습니다: {missing_cols}")
            # 열 이름 확인
            print(f"엑셀에 있는 열: {list(df.columns)}")
            return []
        
        # keyname 생성
        expected_scripts = []
        
        for idx, row in df.iterrows():
            # 빈 값 건너뛰기
            if pd.isna(row['category2']) or pd.isna(row['Level1']):
                continue
                
            # 각 레벨 값 가져오기
            category2 = str(row['category2']).strip()
            level1 = str(row['Level1']).strip() if not pd.isna(row['Level1']) else ""
            level2 = str(row['Level2']).strip() if not pd.isna(row['Level2']) else ""
            level3 = str(row['Level3']).strip() if not pd.isna(row['Level3']) else ""
            
            # 스크립트명 생성
            script_name = self._generate_script_name(category2, level1, level2, level3)
            expected_scripts.append({
                'row': idx + 2,  # 엑셀 행 번호 (헤더 제외)
                'category2': category2,
                'Level1': level1,
                'Level2': level2,
                'Level3': level3,
                'expected_script': script_name,
                'key_name': f"{category2}_{level1}_{level2}_{level3}".rstrip('_')
            })
        
        self.expected_scripts = expected_scripts
        print(f"엑셀에서 {len(expected_scripts)}개의 keyname 읽음")
        return expected_scripts
    
    def _generate_script_name(self, category2, level1, level2, level3):
        """
        keyname으로부터 예상 스크립트명 생성
        """
        parts = [category2, level1, level2, level3]
        # 빈 부분 제거
        parts = [p for p in parts if p and str(p).strip() != '']
        
        # 파일명 생성
        if len(parts) >= 1:
            script_name = f"IBD_{'_'.join(parts)}__.py"
        else:
            script_name = ""
            
        return script_name
    
    def find_python_scripts(self):
        """
        특정 폴더에서 모든 파이썬 스크립트 찾기
        """
        print(f"폴더 {self.scripts_folder} 에서 파이썬 스크립트 검색 중...")
        
        found_scripts = []
        
        # 모든 하위 폴더 탐색
        for root, dirs, files in os.walk(self.scripts_folder):
            for file in files:
                if file.endswith('.py'):
                    file_path = os.path.join(root, file)
                    rel_path = os.path.relpath(file_path, self.scripts_folder)
                    
                    found_scripts.append({
                        'file_name': file,
                        'file_path': file_path,
                        'relative_path': rel_path,
                        'folder': root
                    })
        
        self.found_scripts = found_scripts
        print(f"총 {len(found_scripts)}개의 파이썬 스크립트 찾음")
        return found_scripts
    
    def validate_scripts(self):
        """
        엑셀의 keyname과 실제 스크립트 비교 검증
        """
        if not self.expected_scripts:
            self.read_excel_keys()
        
        if not self.found_scripts:
            self.find_python_scripts()
        
        # 검증 결과 저장
        validation_results = []
        missing_scripts = []
        
        # 각 예상 스크립트에 대해 검증
        for expected in self.expected_scripts:
            expected_name = expected['expected_script']
            found = False
            found_path = ""
            
            # 실제 스크립트에서 찾기
            for script in self.found_scripts:
                if script['file_name'] == expected_name:
                    found = True
                    found_path = script['relative_path']
                    break
            
            # 결과 저장
            result = {
                '엑셀_행': expected['row'],
                'category2': expected['category2'],
                'Level1': expected['Level1'],
                'Level2': expected['Level2'],
                'Level3': expected['Level3'],
                'KeyName': expected['key_name'],
                '예상_스크립트명': expected_name,
                '상태': '✓ 있음' if found else '✗ 없음',
                '실제_경로': found_path if found else ''
            }
            
            validation_results.append(result)
            
            if not found:
                missing_scripts.append(result)
        
        self.missing_scripts = missing_scripts
        return validation_results, missing_scripts
    
    def generate_report(self, validation_results, missing_scripts, output_path=None):
        """
        검증 결과 리포트 생성
        """
        if output_path is None:
            output_path = os.path.join(os.path.dirname(self.excel_path), 
                                      '스크립트_검증_결과.xlsx')
        
        # DataFrame으로 변환
        df_results = pd.DataFrame(validation_results)
        df_missing = pd.DataFrame(missing_scripts)
        
        # ExcelWriter로 두 개의 시트에 저장
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df_results.to_excel(writer, sheet_name='전체_검증결과', index=False)
            df_missing.to_excel(writer, sheet_name='없는_스크립트', index=False)
            
            # 시트별 컬럼 너비 조정
            workbook = writer.book
            worksheet_all = writer.sheets['전체_검증결과']
            worksheet_missing = writer.sheets['없는_스크립트']
            
            # 컬럼 너비 조정
            for column in worksheet_all.columns:
                max_length = 0
                column = [cell for cell in column]
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(cell.value)
                    except:
                        pass
                adjusted_width = (max_length + 2)
                worksheet_all.column_dimensions[column[0].column_letter].width = adjusted_width
            
            for column in worksheet_missing.columns:
                max_length = 0
                column = [cell for cell in column]
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(cell.value)
                    except:
                        pass
                adjusted_width = (max_length + 2)
                worksheet_missing.column_dimensions[column[0].column_letter].width = adjusted_width
        
        print(f"\n검증 결과 리포트 저장: {output_path}")
        
        # 콘솔에 요약 출력
        self._print_summary(validation_results, missing_scripts)
        
        return output_path
    
    def _print_summary(self, validation_results, missing_scripts):
        """
        검증 결과 요약 출력
        """
        print("\n" + "="*60)
        print("스크립트 검증 결과 요약")
        print("="*60)
        
        total_expected = len(validation_results)
        total_found = len([r for r in validation_results if r['상태'] == '✓ 있음'])
        total_missing = len(missing_scripts)
        
        print(f"총 예상 스크립트: {total_expected}개")
        print(f"발견된 스크립트: {total_found}개")
        print(f"누락된 스크립트: {total_missing}개")
        print(f"검증률: {total_found/total_expected*100:.1f}%")
        
        if missing_scripts:
            print("\n[누락된 스크립트 목록]")
            for i, missing in enumerate(missing_scripts[:20], 1):  # 상위 20개만 출력
                print(f"{i:3d}. {missing['예상_스크립트명']}")
            
            if len(missing_scripts) > 20:
                print(f"... 외 {len(missing_scripts)-20}개 더")
        
        # 카테고리별 통계
        if validation_results:
            print("\n[카테고리별 통계]")
            categories = {}
            for result in validation_results:
                cat = result['category2']
                status = result['상태']
                if cat not in categories:
                    categories[cat] = {'total': 0, 'found': 0}
                categories[cat]['total'] += 1
                if status == '✓ 있음':
                    categories[cat]['found'] += 1
            
            for cat, stats in categories.items():
                found_rate = stats['found']/stats['total']*100 if stats['total'] > 0 else 0
                print(f"  {cat}: {stats['found']}/{stats['total']} ({found_rate:.1f}%)")
        
        print("="*60)
    
    def validate_and_fix_naming(self):
        """
        파일명이 약간 다른 경우 유사도 검사 및 제안
        """
        print("\n파일명 유사도 검사 중...")
        
        suggestions = []
        
        for expected in self.expected_scripts:
            expected_name = expected['expected_script'].lower().replace('.py', '')
            found_exact = False
            
            # 정확히 일치하는지 먼저 확인
            for script in self.found_scripts:
                script_name_lower = script['file_name'].lower().replace('.py', '')
                if script_name_lower == expected_name:
                    found_exact = True
                    break
            
            if not found_exact:
                # 유사한 파일명 찾기
                similar_files = []
                for script in self.found_scripts:
                    script_name = script['file_name'].lower().replace('.py', '')
                    
                    # 간단한 유사도 검사
                    if (expected_name in script_name or 
                        script_name in expected_name or
                        self._calculate_similarity(expected_name, script_name) > 0.7):
                        
                        similarity = self._calculate_similarity(expected_name, script_name)
                        similar_files.append({
                            'file_name': script['file_name'],
                            'similarity': similarity,
                            'path': script['relative_path']
                        })
                
                if similar_files:
                    # 유사도 순 정렬
                    similar_files.sort(key=lambda x: x['similarity'], reverse=True)
                    suggestions.append({
                        'expected': expected['expected_script'],
                        'key_name': expected['key_name'],
                        'similar_files': similar_files[:3]  # 상위 3개
                    })
        
        if suggestions:
            print(f"\n[{len(suggestions)}개의 유사한 파일 발견]")
            for i, suggestion in enumerate(suggestions[:10], 1):  # 상위 10개만 출력
                print(f"\n{i}. 예상: {suggestion['expected']}")
                print(f"   Key: {suggestion['key_name']}")
                print("   유사 파일:")
                for sim_file in suggestion['similar_files']:
                    print(f"     - {sim_file['file_name']} ({sim_file['similarity']:.1%})")
            
            if len(suggestions) > 10:
                print(f"\n... 외 {len(suggestions)-10}개 더")
        
        return suggestions
    
    def _calculate_similarity(self, str1, str2):
        """
        두 문자열 간의 간단한 유사도 계산
        """
        # Jaccard 유사도
        set1 = set(str1.split('_'))
        set2 = set(str2.split('_'))
        
        if not set1 or not set2:
            return 0.0
            
        intersection = len(set1.intersection(set2))
        union = len(set1.union(set2))
        
        return intersection / union if union > 0 else 0.0


def main():
    """
    메인 실행 함수
    """
    print("="*60)
    print("엑셀 Keyname - 파이썬 스크립트 자동 검증 프로그램")
    print("="*60)
    
    # 사용자 입력
    excel_path = input("엑셀 파일 경로를 입력하세요: ").strip()
    scripts_folder = input("스크립트 폴더 경로를 입력하세요: ").strip()
    
    # 경로 확인
    if not os.path.exists(excel_path):
        print(f"에러: 엑셀 파일이 없습니다 - {excel_path}")
        return
    
    if not os.path.exists(scripts_folder):
        print(f"에러: 스크립트 폴더가 없습니다 - {scripts_folder}")
        return
    
    # 검증 실행
    validator = ScriptValidator(excel_path, scripts_folder)
    
    try:
        # 1. 엑셀에서 keyname 읽기
        expected_scripts = validator.read_excel_keys()
        if not expected_scripts:
            print("엑셀에서 keyname을 읽을 수 없습니다.")
            return
        
        # 2. 파이썬 스크립트 찾기
        found_scripts = validator.find_python_scripts()
        
        # 3. 검증 실행
        print("\n스크립트 검증 중...")
        validation_results, missing_scripts = validator.validate_scripts()
        
        # 4. 결과 리포트 생성
        report_path = validator.generate_report(validation_results, missing_scripts)
        
        # 5. 유사 파일명 검사 (선택사항)
        run_similarity = input("\n유사 파일명 검사를 실행하시겠습니까? (y/n): ").lower()
        if run_similarity == 'y':
            validator.validate_and_fix_naming()
        
        print(f"\n✓ 검증 완료! 리포트 파일: {report_path}")
        
    except Exception as e:
        print(f"에러 발생: {e}")
        import traceback
        traceback.print_exc()
    
    print("\n프로그램 종료")


if __name__ == "__main__":
    main()