import sys
import tkinter
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from os import path
winvolume = cast(AudioUtilities.GetSpeakers().Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None), POINTER(IAudioEndpointVolume))



def split_every_n(list_, n):
    return [list_[i:i+n] for i in range(0, len(list_), n)]


debug_on = True


def volume_driver(input_):
    try:
        input_ = int(input_)
    except:
        return "{2} Error input_ value is not able to be converted to a int as it is not a number."
    return{
        0 : -60,
        1 : -56,
        2 : -51,
        3 : -46,
        4 : -44,
        5 : -41,
        6 : -39,
        7 : -37,
        8 : -36,
        9 : -34,
        10 : -33,
        11 : -31.5,
        12 : -30.5,
        13 : -29.5,
        14 : -28.5,
        15 : -27.7,
        16 : -26.5,
        17 : -25.5,
        18 : -25,
        19 : -24,
        20 : -23.5,
        21 : -22.7,
        22 : -22,
        23 : -21.5,
        24 : -21,
        25 : -20.3,
        26 : -19.7,
        27 : -19.3,
        28 : -18.7,
        29 : -18.3,
        30 : -17.7,
        31 : -17.3,
        32 : -16.7,
        33 : -16.3,
        34 : -16,
        35 : -15.5,
        36 : -15,
        37 : -14.7,
        38 : -14.3,
        39 : -14,
        40 : -13.6,
        41 : -13.1,
        42 : -12.9,
        43 : -12.5,
        44 : -12.2,
        45 : -11.75,
        46 : -11.5,
        47 : -11.1,
        48 : -10.8,
        49 : -10.5,
        50 : -10.3,
        51 : -10,
        52 : -9.7,
        53 : -9.45,
        54 : -9.1,
        55 : -8.8,
        56 : -8.6,
        57 : -8.3,
        58 : -8,
        59 : -7.8,
        60 : -7.5,
        61 : -7.3,
        62 : -7.1,
        63 : -6.8,
        64 : -6.6,
        65 : -6.4,
        66 : -6.1,
        67 : -5.9,
        68 : -5.7,
        69 : -5.5,
        70 : -5.3,
        71 : -5.1,
        72 : -4.9,
        73 : -4.7,
        74 : -4.5,
        75 : -4.3,
        76 : -4.1,
        77 : -3.9,
        78 : -3.7,
        79 : -3.5,
        80 : -3.3,
        81 : -3.1,
        82 : -2.9,
        83 : -2.7,
        84 : -2.6,
        85 : -2.4,
        86 : -2.2,
        87 : -2.1,
        88 : -1.9,
        89 : -1.7,
        90 : -1.5,
        91 : -1.4,
        92 : -1.2,
        93 : -1.05,
        94 : -0.9,
        95 : -0.7,
        96 : -0.6,
        97 : -0.4,
        98 : -0.3,
        99 : -0.1,
        100: 0
        }.get(input_,"{1} Error input_ value is too high or low ")
def mute():
    winvolume.SetMute(1, None)
def unmute():
    winvolume.SetMute(0, None)

def set_volume(vol):
    winvolume.SetMasterVolumeLevel(volume_driver(vol), None)
def current_volume():
    current = winvolume.GetMasterVolumeLevelScalar()
    if current == 1.0:
        return 100
    else:
        current = str(current)
        item = split_every_n(current, 5)[0]
        while True:
            try:
                if int(list(item)[4]) == 9:
                    return (int("".join(list(item)[2:4]))) + 1;break
                else:
                    return int("".join(list(item)[2:4]));break
            except:
                item = item + str(0)
                
def volume_up(vol=None):
    if vol == None:
        vol = 1
    set_volume(int(current_volume() + int(vol)))

def volume_down(vol=None):
    if vol == None:
        vol = 1
    set_volume(int(current_volume() - int(vol)))

def set_debug(input_):
    f = open("debug.txt","w").write(str(input_).capitalize())

def debug():
    current_vol = current_volume()
    test1 = None
    for i in range(1,101):
        set_volume(i)
        if current_volume() == i:
            pass
        else:
            test1 = False
            set_volume(current_vol)
            print("[winvolume] test on current_volume and set_volume: failed")
            break
    
    if test1 == None:
        set_volume(current_vol)
        print("[winvolume] test on current_volume and set_volume: Passed")
        
def checks():
    if path.exists("debug.txt") == True:
        pass
    else:
        open("debug.txt","w").write("True")
        
    if sys.platform == "win32":
        pass
    else:
        print("This module can only run on win32 computers.")
        quit()
        
    if (open("debug.txt", "r").read()).strip() == "False":
        pass
    else:
        if debug_on == False:
            pass
        else:
            debug()

def set_value(vol):
    if int(vol) == 0:
        mute()
    else:
        set_volume(vol)
        unmute()

def volume_slider(length_=750):
    root = tkinter.Tk()
    slider = tkinter.Scale(root,orient='vertical', from_=0, to=100, command=set_value,length=length_)
    slider.set(current_volume())
    slider.pack()
    root.mainloop()

def main():
    checks()
if __name__ == "__main__":
    main()

