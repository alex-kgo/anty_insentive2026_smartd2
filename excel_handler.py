import win32com.client
import pythoncom
import time
import os

class ExcelHandler:
    def __init__(self):
        self.workbook = None
        self.excel_app = None

    def connect_to_active_excel(self, retries=5):
        """
        Connect to the active Excel instance using ROT (Running Object Table).
        Specific for Excel 2016 and scenarios where direct binding fails.
        """
        print("엑셀 연결 시도 (ROT 방식)...")
        
        for i in range(retries):
            try:
                # ROT(Running Object Table)를 순회하며 엑셀 찾기
                context = pythoncom.CreateBindCtx(0)
                running_objects = pythoncom.GetRunningObjectTable()
                monikers = running_objects.EnumRunning()
                
                excel_found = False
                
                for moniker in monikers:
                    name = moniker.GetDisplayName(context, None)
                    # 엑셀 워크북인지 확인
                    # 한국어 엑셀: "통합 문서"
                    # 영어/일반: "Excel", "Book", ".xlsx", ".xls"
                    if "Excel" in name or "Book" in name or "xlsx" in name or "xls" in name or "통합 문서" in name:
                        try:
                            obj = running_objects.GetObject(moniker)
                            # IDispatch 인터페이스 확인
                            workbook = win32com.client.Dispatch(obj.QueryInterface(pythoncom.IID_IDispatch))
                            
                            # 엑셀 워크북인지 확실히 하기 위해 Application 속성 접근 확인
                            app = workbook.Application
                            if app.Name == "Microsoft Excel":
                                self.workbook = workbook
                                self.excel_app = app
                                print(f"엑셀 연결 성공: {name}")
                                excel_found = True
                                break
                        except Exception:
                            continue
                
                if excel_found:
                    return True
                
                print(f"엑셀을 찾는 중... ({i+1}/{retries})")
                time.sleep(2)
                
            except Exception as e:
                print(f"ROT 조회 중 오류: {e}")
                time.sleep(2)
                
        print("엑셀 연결 실패 (ROT).")
        return False

    def extract_data(self):
        """
        Extract data from the active sheet using pywin32.
        Returns a list of dictionaries.
        """
        if not self.workbook:
            raise Exception("엑셀 워크북이 연결되지 않았습니다.")

        try:
            # 활성 시트 가져오기
            sheet = self.workbook.ActiveSheet
            
            # 사용된 범위 가져오기
            used_range = sheet.UsedRange
            # 데이터 값을 2차원 튜플로 가져옴
            raw_data = used_range.Value
            
            # 데이터가 없는 경우
            if not raw_data: # None checking
                 print("엑셀에 데이터가 없습니다.")
                 return []
            
            # 행 개수 확인 (튜플의 튜플 형태)
            # raw_data가 튜플(행들)이고, 각 행은 값들의 튜플
            # 예: ((val1, val2), (val3, val4))
            
            # win32com은 1-based index 범위를 0-based tuple로 리턴함.
            # 하지만 range.Value는 바로 튜플을 줌.
            
            rows = list(raw_data)
            row_count = len(rows)
            
            if row_count < 2:
                print("데이터 행이 부족합니다.")
                return []
            
            # 헤더 제외하고 데이터 파싱 (2번째 줄부터)
            parsed_data = []
            
            # rows[0] is header, start from rows[1]
            for row_idx in range(1, row_count):
                row = rows[row_idx]
                
                # row is a tuple. 
                # Check column length
                if len(row) < 7:
                    continue
                
                # row content: code, name, cust_in, driver_in, cust_out, driver_out, sum
                # Handle None
                code = str(row[0]).strip() if row[0] is not None else ""
                name = str(row[1]).strip() if row[1] is not None else ""
                
                # 성명 필드에 '합계'가 포함된 행 제외
                if "합계" in name:
                    continue

                def to_int(val):
                    try:
                        return int(float(val)) if val is not None else 0
                    except:
                        return 0

                cust_in = to_int(row[2])
                driver_in = to_int(row[3])
                cust_out = to_int(row[4])
                driver_out = to_int(row[5])
                
                sum_in = cust_in + driver_in
                sum_out = cust_out + driver_out
                total_sum = sum_in + sum_out
                
                parsed_data.append({
                    "code": code,
                    "name": name,
                    "in_sum": sum_in,
                    "out_sum": sum_out,
                    "total_sum": total_sum
                })
                
            return parsed_data
            
        except Exception as e:
            print(f"데이터 추출 중 오류: {e}")
            return []

    def close_workbook(self):
        """Close the workbook without saving."""
        if self.workbook:
            try:
                # SaveChanges=False
                self.workbook.Close(False)
            except Exception as e:
                print(f"워크북 닫기 실패: {e}")
            finally:
                self.workbook = None
                self.excel_app = None
