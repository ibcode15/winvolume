
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume,ISimpleAudioVolume
from threading import Thread
import fire
from time import sleep
import sys
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
def Muted():
    return winvolume.GetMute() == 1
def Mute():
    return winvolume.SetMute(1, None)
def Unmute():
    return winvolume.SetMute(0, None)
def SetVolume(vol:int):
    winvolume.SetMasterVolumeLevelScalar(NormalToScalar(vol), None)
def GetVolumeLevel():
    current = winvolume.GetMasterVolumeLevelScalar()
    return ScalarToNormal(current)
def VolumeUp(vol=2):
    SetVolume(int(GetVolumeLevel() + int(vol)))

def VolumeDown(vol=2):
    SetVolume(int(GetVolumeLevel() - int(vol)))


def ListProcesss():
    Processs = AudioUtilities.GetAllSessions()
    ProcessList = []
    for Process in Processs:
        if Process == None:
            continue
        else:
            ProcessList.append(Process.Process)
    
    ProcessList.remove(None)
    return ProcessList

def ListProcesssCLI():
    Processs = AudioUtilities.GetAllSessions()
    ProcessList = []
    for Process in Processs:
        if Process == None:
            continue
        else:
            ProcessList.append(Process.Process)
    
    ProcessList.remove(None)
    for i in ProcessList:
        print(str(i))
                      
            
def GetProcessByName(ProcessName):
    Processs = AudioUtilities.GetAllSessions()
    for Process in Processs:
        volume = Process._ctl.QueryInterface(ISimpleAudioVolume)
        try:
            if Process.Process.name() == ProcessName:
                
                return volume
        except:
            pass
    
def GetProcessByPID(PID):
    Processs = AudioUtilities.GetAllSessions()
    for Process in Processs:
        volume = Process._ctl.QueryInterface(ISimpleAudioVolume)
        try:
            if Process.Process.pid == PID:
                return volume
        except:
            pass
    
def TypeSwitch(input_):
    if str(type(input_)) == "<class 'str'>":
        input_ = GetProcessByName(input_)
    elif str(type(input_)) == "<class 'int'>":
        input_ = GetProcessByPID(int(input_))
    else:
        return None 
    return input_

class VolumeProcess:
    def __init__(self, Process):
        self.Process = TypeSwitch(Process)
    def SetVolume(self,vol):
        self.Process.SetMasterVolume(NormalToScalar(vol), None)
    def Mute(self):
        self.Process.SetMute(1, None)
    def Unmute(self):
        self.Process.SetMute(0, None)
    def Muted(self):
        return self.Process.GetMute() == 1
    def GetVolumeLevel(self):
        return str(ScalarToNormal(self.Process.GetMasterVolume())) + "%"
    def GetActualVolumeLevel(self):
        return int(ScalarToNormal(self.Process.GetMasterVolume())*round(winvolume.GetMasterVolumeLevelScalar(), 2))

def SetProcessVolume(Process, vol:int):
    Process = TypeSwitch(Process)
    Process.SetMasterVolume(NormalToScalar(vol), None)
def MuteProcess(Process):
    Process = TypeSwitch(Process)
    Process.SetMute(1, None)
def UnMuteProcess(Process):
    Process = TypeSwitch(Process)
    Process.SetMute(0, None)
def GetProcessVolumeLevel(Process):
    Process = TypeSwitch(Process)
    return str(ScalarToNormal(Process.GetMasterVolume())) + "%" 
def GetRealProcessVolumeLevel(Process):
    Process = TypeSwitch(Process)
    return int(ScalarToNormal(Process.GetMasterVolume()) * round(winvolume.GetMasterVolumeLevelScalar(), 2))

#                 Listener
# modes
#
# V - if the volume level has any change, run script
#
# CV (certain volume - not current volume) - if volume is not a certain volume, run script.
# These certain value can be added with the following line, CValue=[values]
# Here is an example, Listener("CV", script="print('hello')", CValue=[100, 50, 0])
# if the volume is not 100,50 or 0 then it will print hello
#
# UM - if unmuted, run script.
#
# M - if muted, run script.


class Listener:
    def __init__(self, mode,script=None,CValue=[]):
        self.mode = mode
        self.script = script
        self.CValue = CValue
        self.StopThread = False
    def func(self):
        
        if self.mode == "V":
            
            self.VolumeLevel = GetVolumeLevel()
            while True:
                if self.StopThread == True:
                    break
                if self.VolumeLevel == GetVolumeLevel():
                    continue
                self.VolumeLevel = GetVolumeLevel()
                exec(self.script)
        elif self.mode == "CV":
            
            while True:
                if self.StopThread == True:
                    break
                self.VolumeLevel = GetVolumeLevel()
                if self.VolumeLevel in self.CValue:
                    continue
                exec(self.script)
        elif self.mode == "UM":
            while True:
                if self.StopThread == True:
                    break
                if Muted():
                    continue
                exec(self.script)
        elif self.mode == "M":
            
            while True:
                if self.StopThread == True:
                    break
                if Muted():
                    exec(self.script)
        else:
            print("help me")
                
    def start(self):
        if "idlelib" in sys.modules == True:
            print()
        self.MyThread = Thread(target = self.func)
        self.MyThread.start()
    def stop(self):
        self.StopThread = True

class ProcessListener:
    def __init__(self,Process,mode,script=None,CValue=[]):
        self.Process = TypeSwitch(Process)
        self.mode = mode
        self.script = script
        self.CValue = CValue
        self.StopThread = False
    def func(self):
        
        if self.mode == "V":
            
            self.VolumeLevel = ScalarToNormal(self.Process.GetMasterVolume())
            while True:
                if self.StopThread == True:
                    break
                if self.VolumeLevel == ScalarToNormal(self.Process.GetMasterVolume()):
                    continue
                self.VolumeLevel = ScalarToNormal(self.Process.GetMasterVolume())
                exec(self.script)
        elif self.mode == "CV":
            
            while True:
                if self.StopThread == True:
                    break
                self.VolumeLevel = ScalarToNormal(self.Process.GetMasterVolume())
                if self.VolumeLevel in self.CValue:
                    continue
                exec(self.script)
        elif self.mode == "UM":
            while True:
                if self.StopThread == True:
                    break
                if self.Process.GetMute() == 1:
                    continue
                exec(self.script)
        elif self.mode == "M":
            
            while True:
                if self.StopThread == True:
                    break
                if self.Process.GetMute() == 1:
                    exec(self.script)
        else:
            print("help me")
                
    def start(self):
        if "idlelib" in sys.modules == True:
            print()
        self.MyThread = Thread(target = self.func)
        self.MyThread.start()
    def stop(self):
        self.StopThread = True    

if __name__ == '__main__':
  fire.Fire({
      'M': Mute,
      'UM': Unmute,
      'SV': SetVolume,
      'CV': GetVolumeLevel,
      'up': VolumeUp,
      'down': VolumeDown,
      'LS': ListProcesssCLI,
      'SSV': SetProcessVolume,
      'MS': MuteProcess,
      'UMS': UnMuteProcess,
      'CSV': GetProcessVolumeLevel,
      'RCSV': GetRealProcessVolumeLevel
  })
