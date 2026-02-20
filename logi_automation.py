import time
import datetime
from pywinauto import Application, Desktop
from pywinauto.findwindows import ElementNotFoundError
import config

class LogiApp:
    def __init__(self):
        self.app = None
        self.dlg = None

    def connect(self):
        """Connect to the running Logi application using PID."""
        print("Attempting to locate Logi application window...")
        try:
            import win32gui
            import win32process
            
            # Find the Logi window
            hwnd = win32gui.FindWindow("SmartD2-", None)
            if not hwnd:
                print("Window class 'SmartD2-' not found. Scanning all windows...")
                # Optional: Add title-based search here if class search becomes unreliable
            
            if hwnd:
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                print(f"Target window found! HWND: {hwnd}, PID: {pid}")
                self.app = Application(backend="win32").connect(process=pid)
                try:
                    self.dlg = self.app.window(handle=hwnd)
                    self.dlg.set_focus()
                    print("Connection established and Logi window focused.")
                except Exception as focus_err:
                     print(f"Connected but could not set focus: {focus_err}")
            else:
                 print("Error: Could not find Logi window. Please make sure the program is open.")
                 raise Exception("Logi application window not found.")
            
        except Exception as e:
            print(f"Connection error details: {e}")
            raise Exception("Logi application connection failed.")

    def login(self, user_id, user_pw):
        """Handle login if the login window is present."""
        try:
            # Check for Login Button (auto_id="2074")
            # Using wait for exists to ensure window is ready
            login_btn = self.dlg.child_window(auto_id="2074", control_type="Button")
            
            if login_btn.exists(timeout=5):
                print("Login screen detected.")
                
                # ID Input (auto_id="1649")
                self.dlg.child_window(auto_id="1649", control_type="Edit").set_text(user_id)
                
                # PW Input (auto_id="2246")
                self.dlg.child_window(auto_id="2246", control_type="Edit").set_text(user_pw)
                
                # Click Login
                login_btn.click()
                print("Clicked Login. Waiting for Main Window...")
                
                # Wait for Main Window (Logic to be added after inspecting Main Window)
                time.sleep(10) 
            else:
                print("Login button not found. Assuming already logged in.")
                
        except Exception as e:
            print(f"Login skip or error: {e}")

    def set_search_period(self, start_dt, end_dt):
        """
        Set the search date range.
        start_dt, end_dt: datetime objects
        Format: YYYY-MM-DD HH:MM
        """
        start_str = start_dt.strftime("%Y-%m-%d %H:%M")
        end_str = end_dt.strftime("%Y-%m-%d %H:%M")
        
        print(f"Setting range: {start_str} ~ {end_str}")
        
        try:
            if config.DATE_SETTING_METHOD == 'coordinates':
                self._set_date_by_coordinates(start_str, end_str)
            else:
                self._set_date_by_controls(start_str, end_str)
            
        except Exception as e:
            print(f"Error setting date: {e}")
            raise

    def _set_date_by_coordinates(self, start_str, end_str):
        """Set dates using relative coordinates from config."""
        coords = config.LOGI_COORDINATES
        
        # Parse date and time separately
        try:
            start_dt = datetime.datetime.strptime(start_str, "%Y-%m-%d %H:%M")
            end_dt = datetime.datetime.strptime(end_str, "%Y-%m-%d %H:%M")
            
            s_year = start_dt.strftime("%Y")
            s_month = start_dt.strftime("%m")
            s_day = start_dt.strftime("%d")
            s_time = start_dt.strftime("%H")
            
            e_year = end_dt.strftime("%Y")
            e_month = end_dt.strftime("%m")
            e_day = end_dt.strftime("%d")
            e_time = end_dt.strftime("%H")
            
        except ValueError:
            print(f"Error parsing date string: {start_str}, {end_str}")
            return

        # Helper to click and type
        def click_and_type(coord_key, value):
            pos = coords.get(coord_key)
            if pos and pos != (0, 0):
                print(f"Setting {coord_key} to {value} at {pos}")
                self.dlg.click_input(coords=pos)
                time.sleep(0.3)
                self.dlg.type_keys("^a{BACKSPACE}", with_spaces=True)
                self.dlg.type_keys(value, with_spaces=True)
                self.dlg.type_keys("{ENTER}")
                time.sleep(0.3)
            else:
                print(f"Warning: Coordinate for {coord_key} not set.")

        # Set Start Date Parts
        click_and_type("start_year", s_year)
        click_and_type("start_month", s_month)
        click_and_type("start_day", s_day)
        click_and_type("start_time", s_time)

        # Set End Date Parts
        click_and_type("end_year", e_year)
        click_and_type("end_month", e_month)
        click_and_type("end_day", e_day)
        click_and_type("end_time", e_time)


    def _set_date_by_controls(self, start_str, end_str):
        """Set dates using control identification (legacy)."""
        # 1. Focus Start Date
        edit_start = self.dlg.child_window(class_name="Edit", found_index=0)
        edit_start.click_input()
        
        # 2. Type Start Date
        time.sleep(0.5)
        edit_start.type_keys("^a{BACKSPACE}", with_spaces=True)
        edit_start.type_keys(start_str, with_spaces=True)
        edit_start.type_keys("{ENTER}")
        
        # 3. Move to End Date
        edit_end = self.dlg.child_window(class_name="Edit", found_index=1)
        edit_end.click_input()
        time.sleep(0.5)
        edit_end.type_keys("^a{BACKSPACE}", with_spaces=True)
        edit_end.type_keys(end_str, with_spaces=True)
        edit_end.type_keys("{ENTER}")

    def click_search_button(self):
        """Click the Search (갱신) button."""
        try:
            # Check for coordinate override
            coords = config.LOGI_COORDINATES.get("search_button")
            if coords and coords != (0, 0):
                print(f"Clicking Search button at coordinates {coords}")
                self.dlg.click_input(coords=coords)
            else:
                # Button text is "갱신" (Update/Refresh) based on inspection
                self.dlg.child_window(title="갱신", class_name="Button", found_index=0).click()
            
            print("Clicked Search. Waiting 5s...")
            time.sleep(5)
            
        except Exception as e:
            print(f"Error clicking search: {e}")
            raise

    def open_excel_view(self):
        """Trigger 'Excel View' (엑셀로 보기)."""
        try:
            # Check for coordinate override for grid click
            coords = config.LOGI_COORDINATES.get("grid_click")
            
            if coords and coords != (0, 0):
                print(f"Right-clicking grid at coordinates {coords}")
                self.dlg.click_input(button='right', coords=coords)
            else:
                 # 1. Find the Main Grid (XTPReport)
                reports = self.dlg.descendants(class_name="XTPReport")
                
                if not reports:
                    print("No XTPReport grid found.")
                    return

                largest_grid = None
                max_area = 0
                
                for grid in reports:
                    rect = grid.rectangle()
                    area = (rect.right - rect.left) * (rect.bottom - rect.top)
                    if area > max_area:
                        max_area = area
                        largest_grid = grid
                
                print(f"Targeting largest grid: {largest_grid} (Area: {max_area})")
                
                # 2. Right Click on the grid
                # Click slightly inside the top-left (e.g., 50, 50) to hit content
                largest_grid.click_input(button='right', coords=(50, 50))

            time.sleep(1)
            
            # 3. Select '엑셀로 보기' from Context Menu
            try:
                # Try keyboard shortcut 'e' (common for Excel export in Korean menus often)
                # Or arrow down + enter?
                # Best: Access menu item
                self.app.PopupMenu.menu_item("엑셀로 보기").click_input()
                print("Clicked 'Excel View'. Waiting for Excel...")
                time.sleep(5) 
            except Exception as menu_error:
                print(f"Menu item error: {menu_error}")
                # Fallback: Try sending 'x' key or arrow keys if needed
                pass

        except Exception as e:
            print(f"Error opening Excel view: {e}")
            raise

    def close_auth_popup(self):
        """
        Close the authentication wizard popup if coordinates are provided.
        """
        coord = config.LOGI_COORDINATES.get("auth_popup_close")
        if coord and coord != (0, 0):
            print(f"Closing auth popup at {coord}...")
            try:
                # 팝업이 떴을 때 로지 창 기준이 아니라, 팝업이 활성화되었을 수 있음.
                # 하지만 사용자가 '로지 프로그램 상대좌표'로 달라고 했으므로 
                # 로지 창 내의 위치로 가정하고 클릭하거나, 
                # 혹은 그냥 화면 절대 좌표로 클릭해야 할 수도 있음.
                # 일단 기존 로지 창 기준 click_input 사용.
                self.dlg.click_input(coords=coord)
                time.sleep(1) # 닫히는 시간 대기
            except Exception as e:
                print(f"Warning: Failed to click auth popup close: {e}")
        else:
            pass # 설정 안되어 있으면 패스
