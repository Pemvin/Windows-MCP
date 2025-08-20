from uiautomation import Control, GetRootControl, ControlType, GetFocusedControl, SetWindowTopmost, IsTopLevelWindow, IsZoomed, IsIconic, IsWindowVisible, ControlFromHandle
from src.desktop.config import EXCLUDED_CLASSNAMES,BROWSER_NAMES, AVOIDED_APPS
from src.desktop.views import DesktopState,App,Size
from src.desktop.translations import APP_TRANSLATIONS
from fuzzywuzzy import process
from psutil import Process
from src.tree import Tree
from time import sleep
from io import BytesIO
from PIL import Image
import subprocess
import pyautogui
import logging
import csv
import io

class Desktop:
    def __init__(self):
        self.desktop_state=None
        self.logger = logging.getLogger('mcp_logger')
        
    def _get_translated_app_name(self, name: str, target_lang: str) -> str:
        """
        Translates the application name to the target language.
        将应用程序名称翻译成目标语言。

        Args:
            name (str): The name of the application to translate.
            target_lang (str): The target language (e.g., 'en', 'zh').

        Returns:
            str: The translated application name.
        """
        name_lower = name.lower()
        for canonical_name, translations in APP_TRANSLATIONS.items():
            if name_lower in translations.values():
                return translations.get(target_lang, canonical_name)
        return name

    def get_state(self,use_vision:bool=False)->DesktopState:
        tree=Tree(self)
        tree_state=tree.get_state()
        if use_vision:
            nodes=tree_state.interactive_nodes
            annotated_screenshot=tree.annotated_screenshot(nodes=nodes,scale=0.5)
            screenshot=self.screenshot_in_bytes(screenshot=annotated_screenshot)
        else:
            screenshot=None
        apps=self.get_apps()
        active_app,apps=(apps[0],apps[1:]) if len(apps)>0 else (None,[])
        self.desktop_state=DesktopState(apps=apps,active_app=active_app,screenshot=screenshot,tree_state=tree_state)
        return self.desktop_state
    
    
    def get_app_status(self,control:Control)->str:
        if IsIconic(control.NativeWindowHandle):
            return 'Minimized'
        elif IsZoomed(control.NativeWindowHandle):
            return 'Maximized'
        elif IsWindowVisible(control.NativeWindowHandle):
            return 'Normal'
        else:
            return 'Hidden'
    
    def get_window_element_from_element(self,element:Control)->Control|None:
        while element is not None:
            if IsTopLevelWindow(element.NativeWindowHandle):
                return element
            element = element.GetParentControl()
        return None

    def get_element_under_cursor(self)->Control:
        return GetFocusedControl()
    
    def get_default_browser(self):
        mapping = {
            "ChromeHTML": "Google Chrome",
            "FirefoxURL": "Mozilla Firefox",
            "MSEdgeHTM": "Microsoft Edge",
            "IE.HTTP": "Internet Explorer",
            "OperaStable": "Opera",
            "BraveHTML": "Brave",
            "SafariHTML": "Safari"
        }
        command= "(Get-ItemProperty HKCU:\\Software\\Microsoft\\Windows\\Shell\\Associations\\UrlAssociations\\http\\UserChoice).ProgId"
        browser,_=self.execute_command(command)
        return mapping.get(browser.strip())
        
    def get_default_language(self)->str:
        command="Get-Culture | Select-Object -ExpandProperty Name"
        response,_=self.execute_command(command)
        return response.strip()
    
    def get_apps_from_start_menu(self)->dict[str,str]:
        """
        Retrieves a dictionary of installed applications from the Start Menu.
        It specifies UTF-8 encoding for the CSV conversion to handle non-ASCII characters.
        从“开始”菜单检索已安装应用程序的字典。
        它为CSV转换指定了UTF-8编码，以处理非ASCII字符。

        Returns:
            dict[str,str]: A dictionary mapping lowercased application names to their AppIDs.
        """
        command='Get-StartApps | ConvertTo-Csv -NoTypeInformation'
        apps_info,_=self.execute_command(command)
        reader=csv.DictReader(io.StringIO(apps_info))
        return {row.get('Name').lower():row.get('AppID') for row in reader}
    
    def execute_command(self,command:str)->tuple[str,int]:
        """
        Executes a PowerShell command and returns the output and status code.
        It ensures that the PowerShell output is UTF-8 encoded.
        执行一个PowerShell命令并返回输出和状态码。
        该方法确保PowerShell的输出是UTF-8编码的。

        Args:
            command (str): The PowerShell command to execute.

        Returns:
            tuple[str,int]: A tuple containing the command's stdout and return code.
        """
        try:
            # Prepend command to set output encoding to UTF-8
            # 在命令前添加设置输出编码为UTF-8的指令
            full_command = f"$OutputEncoding = [Console]::InputEncoding = [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding; {command}"
            
            # Execute the command. Note: we pass `full_command` as a single argument, avoiding `split()`.
            # 执行命令。注意：我们将`full_command`作为一个单独的参数传递，避免使用`split()`。
            result = subprocess.run(['powershell', '-Command', full_command], 
                                    capture_output=True, 
                                    check=True)
            return (result.stdout.decode('utf-8', errors='ignore'), result.returncode)
        except subprocess.CalledProcessError as e:
            # Decode stdout from the exception object and log the error
            # 从异常对象中解码stdout并记录错误
            stdout = e.stdout.decode('utf-8', errors='ignore') if e.stdout else ""
            stderr = e.stderr.decode('utf-8', errors='ignore') if e.stderr else ""
            self.logger.error(f"Command '{command}' failed with exit code {e.returncode}. Stderr: {stderr}")
            return (stdout, e.returncode)
        
    def is_app_browser(self,node:Control):
        process=Process(node.ProcessId)
        return process.name() in BROWSER_NAMES
    
    def resize_app(self,name:str,size:tuple[int,int]=None,loc:tuple[int,int]=None)->tuple[str,int]:
        apps=self.get_apps()
        self.logger.debug(f"Available running apps: {[app.name for app in apps]}")
        
        system_lang = self.get_default_language().split('-')[0]
        translated_name = self._get_translated_app_name(name, system_lang)
        self.logger.debug(f"Translated app name: '{translated_name}' for system language '{system_lang}'")

        matched_app:tuple[App,int]|None=process.extractOne(translated_name,apps,score_cutoff=60)
        if matched_app is None:
            self.logger.warning(f"Application '{name}' not found in running applications with translated name '{translated_name}'")
            return (f'Application {name.title()} not open.',1)
        
        app,_=matched_app
        self.logger.info(f"Found running app: '{app.name}' with handle {app.handle}")
        
        app_control=ControlFromHandle(app.handle)
        if loc is None:
            x=app_control.BoundingRectangle.left
            y=app_control.BoundingRectangle.top
            loc=(x,y)
        if size is None:
            width=app_control.BoundingRectangle.width()
            height=app_control.BoundingRectangle.height()
            size=(width,height)
        
        x,y=loc
        width,height=size
        self.logger.debug(f"Resizing app '{app.name}' to {width}x{height} at position ({x},{y})")
        
        app_control.MoveWindow(x,y,width,height)
        self.logger.info(f"Successfully resized app '{app.name}'")
        return (f'Application {name.title()} resized to {width}x{height} at {x},{y}.',0)
        
    def launch_app(self,name:str)->tuple[str,int]:
        apps_map=self.get_apps_from_start_menu()
        self.logger.debug(f"Available apps in start menu: {apps_map}")
        
        system_lang = self.get_default_language().split('-')[0]
        translated_name = self._get_translated_app_name(name, system_lang)
        self.logger.debug(f"Translated app name: '{translated_name}' for system language '{system_lang}'")

        matched_app=process.extractOne(translated_name,apps_map.keys(),score_cutoff=80)

        if matched_app is None:
            self.logger.warning(f"Application '{name}' not found in start menu with translated name '{translated_name}'.")
            return (f'Application {name.title()} not found in start menu.',1)
        
        app_name_in_map, _ = matched_app
        self.logger.debug(f"Fuzzy matched app name: '{app_name_in_map}'")

        app_id = apps_map[app_name_in_map]
        self.logger.info(f"Found AppID: '{app_id}' for application '{app_name_in_map}'")

        if '!' in app_id or '{' in app_id:
            _,status=self.execute_command(f'Start-Process "shell:AppsFolder\\{app_id}"')
        else:
            _,status=self.execute_command(f'Start-Process "{app_id}"')
        response=f'Launched {name.title()}. Wait for the app to launch...'
        return response,status
    
    def switch_app(self,name:str)->tuple[str,int]:
        apps={app.name:app for app in self.desktop_state.apps}
        matched_app:tuple[str,float]=process.extractOne(name,apps,score_cutoff=70)
        if matched_app is None:
            return (f'Application {name.title()} not found.',1)
        app_name,_=matched_app
        app=apps.get(app_name)
        if SetWindowTopmost(app.handle,isTopmost=True):
            return (f'{app_name.title()} switched to foreground.',0)
        else:
            return (f'Failed to switch to {app_name.title()}.',1)
    
    def get_app_size(self,control:Control):
        window=control.BoundingRectangle
        if window.isempty():
            return Size(width=0,height=0)
        return Size(width=window.width(),height=window.height())
    
    def is_app_visible(self,app)->bool:
        is_minimized=self.get_app_status(app)!='Minimized'
        size=self.get_app_size(app)
        area=size.width*size.height
        is_overlay=self.is_overlay_app(app)
        return not is_overlay and is_minimized and area>10
    
    def is_overlay_app(self,element:Control) -> bool:
        no_children = len(element.GetChildren()) == 0
        is_name = "Overlay" in element.Name.strip()
        return no_children or is_name
        
    def get_apps(self) -> list[App]:
        try:
            sleep(0.75)
            desktop = GetRootControl()  # Get the desktop control
            elements = desktop.GetChildren()
            apps = []
            for depth, element in enumerate(elements):
                if element.ClassName in EXCLUDED_CLASSNAMES or element.Name in AVOIDED_APPS or self.is_overlay_app(element):
                    continue
                if element.ControlType in [ControlType.WindowControl, ControlType.PaneControl]:
                    status = self.get_app_status(element)
                    size=self.get_app_size(element)
                    apps.append(App(name=element.Name, depth=depth, status=status, size=size, process_id=element.ProcessId, handle=element.NativeWindowHandle))
        except Exception as ex:
            print(f"Error: {ex}")
            apps = []
        return apps
    
    def screenshot_in_bytes(self,screenshot:Image.Image)->bytes:
        io=BytesIO()
        screenshot.save(io,format='PNG')
        bytes=io.getvalue()
        return bytes

    def get_screenshot(self,scale:float=0.7)->Image.Image:
        screenshot=pyautogui.screenshot()
        size=(screenshot.width*scale, screenshot.height*scale)
        screenshot.thumbnail(size=size, resample=Image.Resampling.LANCZOS)
        return screenshot