from tkinter.filedialog import askopenfilename
from tkinter import *
import cli
import gettext

window = Tk()
window.title("ofx_to_xlsx")

def close_window():
    window.destroy()


def callback():
    ofx = askopenfilename()
    cli.run(ofx)


gettext.install('ofx_to_xlsx')
t = gettext.translation('gui_i18n', 'locale', fallback=True)
_ = t.gettext

frame = Frame(window)
frame.pack()

w1 = Label (frame,text = _("Select a OFX file to convert it to Excel"))
w1.pack()
arq = Button (frame, text = _("Select File"), command = callback)
arq.pack()
sair = Button (frame, text = _("Quit"), command = close_window)
sair.pack()

window.mainloop()
exit()
