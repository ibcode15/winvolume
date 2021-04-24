from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume,ISimpleAudioVolume
winvolume = cast(AudioUtilities.GetSpeakers().Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None), POINTER(IAudioEndpointVolume))


def NormalToScalar(value:int):
    value_str = str(value)
    if len(value_str) == 3:
        return 1.00
    elif len(value_str) == 2:
        return float("0." +value_str)
    else:
        return float("0.0" +value_str)
def ScalarToNormal(value:int):
    try:
        if str(value)[4] == "9":
            value += 0.01
    except:
        pass
    value = str(value)[0:4]
    if value[0] == "1":
        return 100
    value = value.split(".")[1]
    if value[0] == "0":
        return int(value[1])
    return int(value)
    
def Mute():
    winvolume.SetMute(1, None)
def Unmute():
   winvolume.SetMute(0, None)
def SetVolume(vol:int):
    winvolume.SetMasterVolumeLevelScalar(NormalToScalar(vol), None)
def GetCurrentVolumeLevel():
    current = winvolume.GetMasterVolumeLevelScalar()
    return ScalarToNormal(current)
def VolumeUp(vol=2):
    SetVolume(int(GetCurrentVolumeLevel() + int(vol)))

def VolumeDown(vol=2):
    SetVolume(int(GetCurrentVolumeLevel() - int(vol)))
    
def ListSessions():
    Sessions = AudioUtilities.GetAllSessions()
    SessionList = []
    for session in Sessions:
        if session == None:
            continue
        else:
            SessionList.append(session.Process)
    
    SessionList.remove(None)
    return SessionList
            
            
def GetSessionByName(ProcessName):
    Sessions = AudioUtilities.GetAllSessions()
    for session in Sessions:
        volume = session._ctl.QueryInterface(ISimpleAudioVolume)
        try:
            if session.Process.name() == ProcessName:
                
                return volume
        except:
            pass
    
def GetSessionByPID(PID):
    Sessions = AudioUtilities.GetAllSessions()
    for session in Sessions:
        volume = session._ctl.QueryInterface(ISimpleAudioVolume)
        try:
            if session.Process.pid == PID:
                return volume
        except:
            pass
    
    
class VolumeSession:
    def __init__(self, Session):
        self.Session = Session
        if str(type(self.Session)) == "<class 'str'>":
            self.Session = GetSessionByName(self.Session)
        elif str(type(self.Session)) == "<class 'int'>":
            self.Session = GetSessionByPID(int(self.Session))
        else:
            pass
            
    def SetVolume(self,vol):
        self.Session.SetMasterVolume(NormalToScalar(vol), None)
    def Mute(self):
        self.Session.SetMute(1, None)
    def Unmute(self):
        self.Session.SetMute(0, None)
    def GetCurrentVolumeLevel(self):
        return ScalarToNormal(self.Session.GetMasterVolume())
    def GetCurrentActualVolumeLevel(self):
        GetVolume1 = ScalarToNormal(self.Session.GetMasterVolume())
        GetVolume2 = round(winvolume.GetMasterVolumeLevelScalar(), 2)
        return int(GetVolume1*GetVolume2)
