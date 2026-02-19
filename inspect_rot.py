import pythoncom
import win32com.client

def check_rot():
    print("=== Running Object Table (ROT) 검사 시작 ===")
    
    try:
        context = pythoncom.CreateBindCtx(0)
        running_objects = pythoncom.GetRunningObjectTable()
        monikers = running_objects.EnumRunning()
        
        found_count = 0
        for moniker in monikers:
            try:
                name = moniker.GetDisplayName(context, None)
                print(f"발견된 객체: {name}")
                found_count += 1
                
                if "Excel" in name or "Book" in name or "xlsx" in name:
                    print(f"  -> 엑셀 관련 객체로 추정됨!")
                    try:
                        obj = running_objects.GetObject(moniker)
                        disp = win32com.client.Dispatch(obj.QueryInterface(pythoncom.IID_IDispatch))
                        print(f"  -> 연결 성공! Application Name: {disp.Application.Name}")
                    except Exception as e:
                        print(f"  -> 연결 시도 중 에러: {e}")
                        
            except Exception as e:
                print(f"객체 이름 확인 중 에러: {e}")
                
        print(f"=== 검사 종료 (총 {found_count}개 발견) ===")
        
    except Exception as e:
        print(f"ROT 접근 중 치명적 에러: {e}")

if __name__ == "__main__":
    print("이 프로그램은 현재 실행 중인 프로그램들(엑셀 포함)이 윈도우에 어떻게 등록되어 있는지 확인합니다.")
    print("반드시 **로지 프로그램을 통해 엑셀이 열려 있는 상태**에서 실행해주세요.")
    input("준비되었으면 엔터를 누르세요...")
    check_rot()
