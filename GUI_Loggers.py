# noinspection PyUnresolvedReferences
from tkinter import *

# noinspection PyUnresolvedReferences
import loggins

# noinspection PyUnresolvedReferences
import arrow

# noinspection PyUnresolvedReferences
from PIL import ImageTk, Image

#RAIZ
raiz=Tk()
raiz.title("SMG GROUP // " + arrow.now().format('DD-MM-YYYY'))
raiz.geometry("650x350")

#FONDO
img = ImageTk.PhotoImage(Image.open("SwissMedical.png"))
panel = Label(raiz, image = img)
panel.place(x=0, y=0, relwidth=1, relheight=1)

#LABELS
text2= Label(panel, text= "Estado Bajada", font=("Arial", 7))
text2.grid(row=0, column=3)



#BOTONES
botonarchivotransfer=Button(panel, text="Log to VPN", height=1, width= 24,command=lambda:loggins.get_to_vpn())
botonarchivotransfer.grid(row=1,column=0, padx=5, pady=9)

botonarchivotransfer=Button(panel, text="Focus", height=1, width= 24,command=lambda:loggins.zcl04())
botonarchivotransfer.grid(row=2,column=0, padx=5, pady=9)

raiz.mainloop()