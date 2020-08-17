from datetime import datetime

from openpyxl import Workbook










def checkPW():
    
    workbook = Workbook()
    sheet = workbook.active

    now = datetime.now()
    date = now.today()
    current_time = now.strftime("%m/%d -- %H:%M:%S")
    #print("Current Time =", current_time)
    #https://stackoverflow.com/questions/50491839/python-openpyxl-find-strings-in-column-and-return-row-number
    if(pw.get()=="01"):
        msg.set("張三歡迎登入,您簽到的時間是:"+current_time)
        sheet["A1"] = "張三01"
        sheet["B1"] = now
    elif(pw.get()=="02"):
        msg.set("李四歡迎登入,您簽到的時間是:"+current_time)
    elif(pw.get()=="03"):
        msg.set("王五歡迎登入,您簽到的時間是:"+current_time)
    elif(pw.get()=="04"):
        msg.set("陳六歡迎登入,您簽到的時間是:"+current_time)
    else:
        msg.set("您輸入錯誤工號,請重新輸入!")
        
        
    workbook.save(filename="hello_world.xlsx")
		
import tkinter as tk


worktime=tk.Tk()
pw=tk.StringVar()
msg=tk.StringVar()
worktime.geometry("720x300")
worktime.title("上班打卡")
label=tk.Label(worktime,text="請輸入員工工號:")
label.pack()
entry=tk.Entry(worktime,textvariable=pw)
entry.pack()
button=tk.Button(worktime,text="登入",command=checkPW)
button.pack()
lblmsg=tk.Label(worktime,fg="red",textvariable=msg)
lblmsg.pack()
worktime.mainloop()


