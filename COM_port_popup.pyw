
# This app should be added to your startup folder in windows:
# to find this folder, hit the windows key and type "run"  then in the box
# type "shell:startup".  Drop this file there and restart the computer 
# When you plug in a COM port, it should automatically pop up with the name.




# Reference for polling USB ports for changes: https://stackoverflow.com/a/62614107

from infi.devicemanager import DeviceManager
from plyer import notification

import win32api, win32con, win32gui
import win32com.client 
from ctypes import *

# #refernce for no command window: https://stackoverflow.com/questions/764631/how-to-hide-console-window-in-python
the_program_to_hide = win32gui.GetForegroundWindow()
win32gui.ShowWindow(the_program_to_hide , win32con.SW_HIDE)


#
# Device change events (WM_DEVICECHANGE wParam)
#
DBT_DEVICEARRIVAL = 0x8000
DBT_DEVICEQUERYREMOVE = 0x8001
DBT_DEVICEQUERYREMOVEFAILED = 0x8002
DBT_DEVICEMOVEPENDING = 0x8003
DBT_DEVICEREMOVECOMPLETE = 0x8004
DBT_DEVICETYPESSPECIFIC = 0x8005
DBT_CONFIGCHANGED = 0x0018

#
# type of device in DEV_BROADCAST_HDR
#
DBT_DEVTYPE_PORT = 0x00000003



DWORD = c_ulong


class DEV_BROADCAST_HDR(Structure):
    _fields_ = [
        ("dbch_size", DWORD),
        ("dbch_devicetype", DWORD),
        ("dbch_reserved", DWORD)
    ]



class Notification:
    def __init__(self):
        message_map = {
            win32con.WM_DEVICECHANGE: self.onDeviceChange
        }

        wc = win32gui.WNDCLASS()
        hinst = wc.hInstance = win32api.GetModuleHandle(None)
        wc.lpszClassName = "DeviceChangeDemo"
        wc.style = win32con.CS_VREDRAW | win32con.CS_HREDRAW
        wc.hCursor = win32gui.LoadCursor(0, win32con.IDC_ARROW)
        wc.hbrBackground = win32con.COLOR_WINDOW
        wc.lpfnWndProc = message_map
        classAtom = win32gui.RegisterClass(wc)
        style = win32con.WS_OVERLAPPED | win32con.WS_SYSMENU
        self.hwnd = win32gui.CreateWindow(
            classAtom,
            "Device Change Demo",
            style,
            0, 0,
            win32con.CW_USEDEFAULT, win32con.CW_USEDEFAULT,
            0, 0,
            hinst, None
        )

    def onDeviceChange(self, hwnd, msg, wparam, lparam):
        #
        # WM_DEVICECHANGE:
        #  wParam - type of change: arrival, removal etc.
        #  lParam - what's changed?
        #    if it's a volume then...
        #  lParam - what's changed more exactly
        #
        dev_broadcast_hdr = DEV_BROADCAST_HDR.from_address(lparam)

        if wparam == DBT_DEVICEARRIVAL:
            # print("Something's arrived")
            
            if dev_broadcast_hdr.dbch_devicetype == DBT_DEVTYPE_PORT:
                # print("It's a port")
                dm = DeviceManager()
                dm.root.rescan()
                devices = dm.all_devices
                for device in devices:
                    if 'COM' in str(device.has_property):

                        deviceString = str(device.has_property)
                        startPos = deviceString.find('COM')
                        com_string = deviceString[startPos:]
                        endPos = com_string.find(')')
                        clean_com_name = com_string[:endPos]
                        # print(clean_com_name)

# # reference for popup notifiction example: https://stackoverflow.com/questions/15921203/how-to-create-a-system-tray-popup-message-with-python-windowsfrom plyer.utils import platform

                        notification.notify(
                            title=clean_com_name,
                            message='Com Port Detected',
                            # app_name='Here is the application name',
                            # app_icon='path/to/the/icon.{}'.format(
                            #     # On Windows, app_icon has to be a path to a file in .ICO format.
                            #     'ico' if platform == 'win' else 'png'
                            # )
                        )
                
        return 1


if __name__ == '__main__':
    w = Notification()
    win32gui.PumpMessages()