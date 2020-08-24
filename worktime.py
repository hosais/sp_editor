from datetime import datetime

from openpyxl import Workbook











    
    
		
import tkinter as tk


    




class WorkTime:
    def __init__(self):
        self.worktime_ui = tk.Tk()
        self.employee_num=tk.StringVar()
        self.msg=tk.StringVar()
        self.worktime_ui.geometry("720x300")
        self.worktime_ui.title("上班打卡")
        self.employee_num_lb=tk.Label(self.worktime_ui,text="請輸入員工工號:")
        self.employee_num_lb.pack()
        self.employee_entry = tk.Entry(self.worktime_ui,textvariable=self.employee_num)
        self.employee_entry.pack()
        self.login_btn = tk.Button(self.worktime_ui,text="登入",command=self.check_pw)
        self.login_btn.pack()
        self.result_msg=tk.Label(self.worktime_ui,fg="red",textvariable=self.msg)
        self.result_msg.pack()
        self.worktime_ui.mainloop()
    def get_user_name(employee_num): 
        '''
            Open spreadsheet workers (ID, name) and search the matched user name
            return user name
            
        '''
    def is_login_history_normal(self):
        '''
            exist 2 or times of 2 login records 
        '''
    def build_log_record(self):
        '''
            is_login_history?
            Correct -> create and save log record
            Wrong -> create and save log record 
                     Show error message that need to contact with manager.
        '''
    def save_log_record2sp(self):
        '''
        
        '''
    def check_pw(self):

        workbook = Workbook()
        sheet = workbook.active

        now = datetime.now()
        date = now.today()
        current_time = now.strftime("%m/%d -- %H:%M:%S")
        #print("Current Time =", current_time)
        #https://stackoverflow.com/questions/50491839/python-openpyxl-find-strings-in-column-and-return-row-number
        #
       
        if(self.employee_num.get()=="01"):
            self.msg.set("張三歡迎登入,您簽到的時間是:"+current_time)
            sheet["A1"] = "張三01"
            sheet["B1"] = now
        elif(self.employee_num.get()=="02"):
            self.msg.set("李四歡迎登入,您簽到的時間是:"+current_time)
        elif(self.employee_num.get()=="03"):
            self.msg.set("王五歡迎登入,您簽到的時間是:"+current_time)
        elif(self.employee_num.get()=="04"):
            self.msg.set("陳六歡迎登入,您簽到的時間是:"+current_time)
        else:
            self.msg.set("您輸入錯誤工號,請重新輸入!")


        workbook.save(filename="hello_world.xlsx")
        
    


def main():
    
    wt = WorkTime()    

    
    

if __name__ == "__main__":
    main()
    
    
    
    


