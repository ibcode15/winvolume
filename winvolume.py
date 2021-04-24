from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume,ISimpleAudioVolume
import fire
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

def ListSessionsCLI():
    Sessions = AudioUtilities.GetAllSessions()
    SessionList = []
    for session in Sessions:
        if session == None:
            continue
        else:
            SessionList.append(session.Process)
    
    SessionList.remove(None)
    for i in SessionList:
        print(str(i))
                      
            
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
    
def TypeSwitch(input_):
    if str(type(input_)) == "<class 'str'>":
        input_ = GetSessionByName(input_)
    elif str(type(self.Session)) == "<class 'int'>":
        input_ = GetSessionByPID(int(input_))
    else:
        pass
    return input_
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
        return str(ScalarToNormal(self.Session.GetMasterVolume())) + "%"
    def GetCurrentActualVolumeLevel(self):
        return int(ScalarToNormal(self.Session.GetMasterVolume())*round(winvolume.GetMasterVolumeLevelScalar(), 2))

def SetSessionVolume(session, vol:int):
    session = TypeSwitch(session)
    session.SetMasterVolume(NormalToScalar(vol), None)
def MuteSession(session):
    session = TypeSwitch(session)
    session.SetMute(1, None)
def UnMuteSession(session):
    session = TypeSwitch(session)
    session.SetMute(0, None)
def GetSessionCurrentVolumeLevel(session):
    session = TypeSwitch(session)
    return str(ScalarToNormal(session.GetMasterVolume())) + "%" 
def GetRealSessionCurrentVolumeLevel(session):
    session = TypeSwitch(session)
    return int(ScalarToNormal(session.GetMasterVolume()) * round(winvolume.GetMasterVolumeLevelScalar(), 2))
if __name__ == '__main__':
  fire.Fire({
      'M': Mute,
      'UM': Unmute,
      'SV': SetVolume,
      'CV': GetCurrentVolumeLevel,
      'up': VolumeUp,
      'down': VolumeDown,
      'LS': ListSessionsCLI,
      'SSV': SetSessionVolume,
      'MS': MuteSession,
      'UMS': UnMuteSession,
      'CSV': GetSessionCurrentVolumeLevel,
      'RCSV': GetRealSessionCurrentVolumeLevel
  })
