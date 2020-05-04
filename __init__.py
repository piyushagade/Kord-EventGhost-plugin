import eg
import wx
from threading import Event, Thread
from time import sleep
from eg.WinApi import (
    SendMessageTimeout,
    GetWindowText,
    FindWindow,
    WM_COMMAND,
    WM_USER
)
from eg.WinApi.Utils import BringHwndToFront
from win32gui import FindWindow, MessageBox
from win32api import ShellExecute

global sendWAActive
sendWAActive=False

ICON_B64 = "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAMAAAAoLQ9TAAAABGdBTUEAALGPC/xhBQAAAAFzUkdCAK7OHOkAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAAbxQTFRFAAAAJCQkLy8vOTk5KCgoWlpaXFxcMDAwNjY2MTExMjIyXV1dGhoaLS0tIiIiLS0tLi4uLy8vLy8vLy8vLS0tLi4uMDAwNDQ0OTk5PDw8PDw8Ojo6NTU1Ojo6YWFhLy8vMjIyOzs7UlJSX19fMTExLi4uMjIyPj4+NTU1MDAwOjo6Pj4+MTExKCgoMzMzNTU1LS0tLi4uODg4Ojo6Ly8vLi4uOjo6PT09MDAwLi4uOjo6PT09MDAwLS0tNzc3Ojo6Ly8vJiYmMzMzNTU1LS0tMDAwOjo6Pj4+MDAwLS0tMjIyPT09bm5uNDQ0LS0tLy8vMjIyOjo6UVFRXl5eMTExMDAwMzMzODg4Ozs7Ozs7OTk5Ozs7YmJiKysrLi4uLy8vLy8vLi4uLCwsRUVFSUlJSkpKTk5OVlZWUFBQaGhob29vS0tLXFxcdHR0fHx8fX19d3d3YGBga2trREREW1tbgICAf39/goKCfn5+R0dHSEhITU1NcnJygYGBpKSki4uLUlJSenp6paWlra2tsLCwmpqaoqKiqqqqcXFxo6OjiYmJQ0NDWlpae3t7dXV1XV1dampqVFRUZ2dn////XZP5+AAAAGN0Uk5TAAAAAAAAAAAAAAAAAAAACSlFSC8NBkyu4vLz5rpeDA2I8/mjGQSE+6ND7/pjA5+/EBzV6jcx6PdTMOf2URrT6TQCmrsOPuv4XAN7+f6aCgp97vaXFEKh2e3u31MJBSA5OyUItopo4QAAAAFiS0dEk+ED37YAAAAJcEhZcwAACxMAAAsTAQCanBgAAAEBSURBVBjTY2BgYGBk4hcQFBIWYWZhAAMmUTFxCUkpaRlZOVY2IJ9dXkExOQUIkpWUVTg4GThV1dRTUtPS0zNSMrM0RIACmlop2Tm5efkFhUX52jpcDLp6xakleaVlZaX55RWV+gYMhkZV1TWlZaW1ZRV19anGJgymZikNjWVlTc3l+S2tqeYWDJZWYIG29lqQgLUNg61dVXUHUEtpWUVnV6q9A4OjU3dqTy/I0LzSikpnFwZOV7eU1J6a3ry+/gn57h5cDJyeXt4pqdUTJ6alTPLx9eMEOt0/ILC4Cuj04qDgEKDTgZ5TDQ0Lj4iMEo+O4YZ4l4c3Ni4+ITGJnQ/IAQCyokrouXaVHgAAACV0RVh0ZGF0ZTpjcmVhdGUAMjAyMC0wNS0wM1QyMToyNTozMSswMDowMLGbPEAAAAAldEVYdGRhdGU6bW9kaWZ5ADIwMjAtMDUtMDNUMjE6MjU6MzErMDA6MDDAxoT8AAAARnRFWHRzb2Z0d2FyZQBJbWFnZU1hZ2ljayA2LjcuOC05IDIwMTQtMDUtMTIgUTE2IGh0dHA6Ly93d3cuaW1hZ2VtYWdpY2sub3Jn3IbtAAAAABh0RVh0VGh1bWI6OkRvY3VtZW50OjpQYWdlcwAxp/+7LwAAABh0RVh0VGh1bWI6OkltYWdlOjpoZWlnaHQAMTkyDwByhQAAABd0RVh0VGh1bWI6OkltYWdlOjpXaWR0aAAxOTLTrCEIAAAAGXRFWHRUaHVtYjo6TWltZXR5cGUAaW1hZ2UvcG5nP7JWTgAAABd0RVh0VGh1bWI6Ok1UaW1lADE1ODg1NDExMzGbkU8AAAAAD3RFWHRUaHVtYjo6U2l6ZQAwQkKUoj7sAAAAVnRFWHRUaHVtYjo6VVJJAGZpbGU6Ly8vbW50bG9nL2Zhdmljb25zLzIwMjAtMDUtMDMvNmYyMDBiZDYyMjZiNjA3YWM3MmNhZjYwOGVlMzQ0ZDcuaWNvLnBuZwm9NeMAAAAASUVORK5CYII=";
KORD_APP_UID = None
KORD_INSTALL_PATH = None

eg.RegisterPlugin(
    name = "Kord",
    guid = "{A26297A6-4123-4262-92DD-DDE818F7A155}",
    author = "Piyush Agade",
    version = "1.0.1",
    kind = "other",
    icon = ICON_B64,
    createMacrosOnAdd = True,
    description = "This is an extension for Kord keyboard shortcuts application.\n\n<a href=\"https://piyushagade.xyz/kord\">Click here</a> to visit website."
)

class Kord(eg.PluginBase):

    def Configure(self, myString=""):
        panel = eg.ConfigPanel()

        staticText1 = panel.StaticText("This is the configuration window of Kord plugin. \n\nIf the plugin doesn't work, chances are the installation path in not found or incorrectly filled in below.\nIf the plugin still doesn't work, email the author at piyushagade@gmail.com")
        staticText2 = panel.StaticText("Input the location of the Kord installation folder. The plugin needs the location to integrate with Kord.")
        tcpBox1 = panel.BoxedGroup(
            "Kord",
            ("", staticText1),
            ("", staticText2),
        )
        eg.EqualizeWidths(tcpBox1.GetColumnItems(0))
        panel.sizer.Add(tcpBox1, 0, wx.EXPAND)
        
        textControl = panel.TextCtrl("C:/Program Files (x86)/Kord")
        tcpBox2 = panel.BoxedGroup(
            "Installation path",
            ("", textControl),
        )
        eg.EqualizeWidths(tcpBox2.GetColumnItems(0))
        panel.sizer.Add(tcpBox2, 0, wx.EXPAND)

        while panel.Affirmed():
            panel.SetResult(textControl.GetValue())
            
    def __init__(self):
                
        # self.AddAction(Test)
        # group_trigger = self.AddGroup(
        #     "Events",
        #     "Pressing a chord triggers a specified event"
        # )
        # group_trigger.AddAction(Trigger)

        pass
        

    def __start__(self, params):
        global KORD_INSTALL_PATH
        KORD_INSTALL_PATH = self.GetInstallPath()[0]
        if(KORD_INSTALL_PATH == None):
            KORD_INSTALL_PATH = params
            print "Kord installation not found"
        else:
            print "Kord plugin successfully initialized"
            
        self.stopThreadEvent = Event()
        thread = Thread(
            target=ThreadLoop,
            args=(self, self.stopThreadEvent)
        )
        thread.start()
    
    def __stop__(self):
        self.stopThreadEvent.set()

    def GetInstallPath(self):
        import _winreg
        try:
            fb = _winreg.OpenKey(
                _winreg.HKEY_LOCAL_MACHINE,
                "SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{7909EE4C-D8F5-43B6-9D31-7189250DA5BA}}_is1"
            )
            install_location, dummy =_winreg.QueryValueEx(fb, "InstallLocation")
            version, dummy =_winreg.QueryValueEx(fb, "DisplayVersion")
            _winreg.CloseKey(fb)
        except WindowsError:
            install_location= None
            version = None
        return [install_location, version]


def ThreadLoop(self, stopThreadEvent):
    while not stopThreadEvent.isSet():
        import time
        import codecs
        # KORD_INSTALL_PATH = "C:\\Users\\Piyush\\OneDrive - University of Florida\\Laboratory\\2. Electron\Kord\\"
        input = codecs.open(KORD_INSTALL_PATH + "eg", 'r', eg.systemEncoding, 'ignore')
        data = input.readlines()
        input.close()
        if len(data) > 0:
            for line in data:
                if len(line.strip()) > 0:
                    trigger_id = line.strip().split(":-:")[0]
                    chord_id = line.strip().split(":-:")[1]
                    payload = line.strip().split(":-:")[2]
                    
                    if chord_id == "Notification":
                        message = payload.replace("#>#", " | ")
                        eg.plugins.EventGhost.ShowOSD(message, "0;-24;0;0;0;700;0;0;0;1;0;0;2;26;Arial", (255, 255, 255), (0, 0, 0), 4, (0, 0), 0, 3.0, True)
                    else:
                        eg.TriggerEvent(chord_id, prefix = "Kord", payload = [payload, trigger_id])
        stopThreadEvent.wait(1.0)

# class Trigger(eg.ActionBase):
#     name = "Trigger Event"
#     description = "Triggers an Event Ghost event when use plays a chord"

#     def Configure(self, myString=""):
#         panel = eg.ConfigPanel()
#         panel.AddLabel("Chord ID:")
#         textControl = wx.TextCtrl(panel, -1, "100")
#         panel.AddCtrl(textControl)
#         while panel.Affirmed():
#             panel.SetResult(textControl.GetValue())

#     def __call__(self, payload):
#         self.payload = payload[0]
#         self.id = payload[1]
#         print self.id
#         # self.TriggerEvent("MyTimerEvent")
