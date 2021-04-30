from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
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
