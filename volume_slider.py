import tkinter
from winvolume import set_volume,current_volume,unmute,mute
def set_value(vol):
    if int(vol) == 0:
        mute()
    else:
        set_volume(vol)
        unmute()

def volume_slider():
    root = tkinter.Tk()
    slider = tkinter.Scale(root,orient='vertical', from_=0, to=100, command=set_value,length=750)
    slider.set(current_volume())
    slider.pack()
    root.mainloop()


