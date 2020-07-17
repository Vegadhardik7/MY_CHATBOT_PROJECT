from chatterbot import ChatBot
from chatterbot.trainers import ListTrainer
from chatterbot.logic import logic_adapter
from tkinter import *
from tkinter import font
from tkinter import messagebox
import speech_recognition as s
import threading
import pyaudio
import pyttsx3 as pp
import win32com.client
import os

main=Tk()
frame = Frame(main)
ment=StringVar()

main.geometry("500x700")
main.title("CS Department")
img=PhotoImage(file='unnamed.png')
photo=Label(main,image=img)
photo.place(x=1,y=440)


eng=pp.init()

voices = eng.getProperty('voices')
print(voices)

i=IntVar()

def speak(word):
    if i.get()==1:
        eng.setProperty('voice', voices[0].id)
        eng.say(word)
        eng.runAndWait()
    elif i.get()==2:
        eng.setProperty('voice', voices[1].id)
        eng.say(word)
        eng.runAndWait()
    else:
        eng.runAndWait()

lbl=Label(main,text="Select Gender of voice :",font=40).place(x=100,y=50)

r1=Radiobutton(main,text="Male Voice",value=1,variable=i,command=speak,font=35).place(x=100,y=100)

r2=Radiobutton(main,text="Female Voice",value=2,variable=i,command=speak,font=35).place(x=100,y=150)

r3=Radiobutton(main,text="Sound ON/OF",value=3,variable=i,command=speak,font=35).place(x=100,y=200)

#import conversation
chatbot=ChatBot('Bot')

for _file in os.listdir('files'):
    chats=open('files/' + _file,'r').readlines()
    trainer = ListTrainer(chatbot)
    trainer.train(chats)


sy=Scrollbar(main)

fnt = font.Font(size=10)
lbl2=Label(main,text="List of Questions you may ask :",font=35).place(x=1150,y=15)
txt=Listbox(main,width=55,height=44,font=fnt,yscrollcommand=sy.set, bd=5)

txt.insert(1,"  ADMISSION INFORMATION:  "+"\n")
txt.insert(2,"\n")
txt.insert(3," what is the registration process"+"\n")
txt.insert(4," what document will be required"+"\n")
txt.insert(5," what are the timing for admission"+"\n")
txt.insert(6," Form where do we have to collect forms"+"\n")
txt.insert(7," How to fill form"+"\n")
txt.insert(8," can we cancel admission"+"\n")
txt.insert(9," Refund"+"\n")
txt.insert(10,"\n")
txt.insert(11,"\n")
txt.insert(12,"  FEES INFORMATION:  "+"\n")
txt.insert(13,"\n")
txt.insert(14," What is the Fees for computer science"+"\n")
txt.insert(15," what document will be required"+"\n")
txt.insert(16," Fees for open categories"+"\n")
txt.insert(17," What is Fees for NT cast"+"\n")
txt.insert(18," Whats is Fees for Other Backward Class"+"\n")
txt.insert(19," What is the Fees for extracurricular activities"+"\n")
txt.insert(20,"\n")
txt.insert(21,"\n")
txt.insert(22,"  EXTRACURRICULAR INFORMATION:  "+"\n")
txt.insert(23,"\n")
txt.insert(24," extracurricular activities are there in you college"+"\n")
txt.insert(25," what document will be required"+"\n")
txt.insert(26," what link of extracurricular activities will be provided"+"\n")
txt.insert(27," how will extracurricular activities will help to my child"+"\n")
txt.insert(28,"\n")
txt.insert(29,"\n")
txt.insert(30,"  CONCESSION INFORMATION:  "+"\n")
txt.insert(31,"\n")
txt.insert(32," what kind of concession will be provided"+"\n")
txt.insert(32," Any specific concession for girls"+"\n")
txt.insert(33,"\n")
txt.insert(34,"\n")
txt.insert(35,"  PAYMENT INFORMATION:  "+"\n")
txt.insert(36,"\n")
txt.insert(37," What are the Payment procedure"+"\n")
txt.insert(38," Can payment be done in installment"+"\n")
txt.insert(39," Do you support UPI"+"\n")
txt.insert(40," Any cash back is available for specific bank"+"\n")
txt.insert(41,"\n")
txt.insert(42,"\n")
txt.insert(43,"  STAFF INFORMATION:  "+"\n")
txt.insert(44,"\n")
txt.insert(45," how is your teaching staff"+"\n")
txt.insert(46," how teachers will inform us about our childs progress"+"\n")
txt.insert(47," how technically advance are you laboratory"+"\n")
txt.insert(48," what placement opportunity will be provided to over child"+"\n")
txt.insert(49,"\n")
txt.insert(50,"\n")
txt.insert(51,"  CONTACT INFORMATION:  "+"\n")
txt.insert(52,"\n")
txt.insert(53," I need a help"+"\n")
txt.insert(54," what is your contact number"+"\n")
txt.insert(55," Help"+"\n")
txt.insert(56," How can I contact you"+"\n")
txt.insert(57," how can I get further information"+"\n")
txt.insert(58," I do not get you"+"\n")
txt.insert(59," calling which number might help me"+"\n")

txt.config()
sy.pack(side=RIGHT,fill=Y)
txt.place(x=1125,y=45)

#I do not get you
lbl3=Label(main,text="Help us to Advance, Enter your unanswered Query :",font=35).place(x=525,y=585)
# txt1=Entry(main,bd=5,width=20)
# txt1.place(x=515,y=580)

txt1 = Entry(main,font=10, bd=5)
txt1.place(width=500,height=50,x=515,y=650)



def Unsolved():
    q = txt1.get()

    if len(q) < 5 :
        messagebox.showinfo("Error","Please Enter Your Unsolved Query")
    else:
        f=open('Unsolved_Query.txt','a')
        f.write(q+"\n")
        f.close()
        messagebox.showinfo("Congratulations", "Your Query is Accepted Successfully")
        txt1.delete(0, END)



btn1=Button(main, text="Enter Query", font=("Times", 16),command=Unsolved).place(x=710,y=710)

sc=Scrollbar(frame)
scx=Scrollbar(frame,orient=HORIZONTAL)

msg=Listbox(frame,width=100,height=25,yscrollcommand=sc.set,xscrollcommand=scx.set, bd=5)
scx.pack(side=BOTTOM,fill=X)

scx.config(command=msg.xview)
msg.config()

msg.pack(side=LEFT,fill=BOTH,pady=10)
sc.pack(side=RIGHT,fill=Y)
frame.pack()

def take_query():
    sr=s.Recognizer()
    sr.pause_threshold=1
    print("Your Bot is listening try to speak")
    with s.Microphone() as m:
        try:
            audio = sr.listen(m)
            question = sr.recognize_google(audio, language='eng-in')
            print(question)
            text.delete(0, END)
            text.insert(0, question)
            Ask_from_Bot()
        except Exception as e:
            print(e)
            print("Not recognized")

def Ask_from_Bot():
    question = text.get()

    if question == 'Bye' or question == 'bye':
        print("ChatBot:Bye")
        main.destroy()
    else:
        ans = chatbot.get_response(question)
        msg.insert(END, "You : " + question + '\n')
        print(type(ans))
        msg.insert(END, "Bot : " + str(ans))
        speak(ans)
        text.delete(0, END)
        msg.xview()
        msg.yview(END)


# Creating text field
text = Entry(main, bd=5,font=("Times", 15))
text.place(x=650,y=450,width=220,height=35)

btn = Button(main, text="Ask From Bot", font=("Times", 16), command=Ask_from_Bot)
btn.place(x=695,y=510)

# Press Enter & get Output
def Enter_fun(event):
    btn.invoke()


# going to bind main window with enter key

main.bind('<Return>', Enter_fun)
def repeatL():
    while True:
        take_query()
t=threading.Thread(target=repeatL)
t.start()
main.mainloop()
