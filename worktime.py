from datetime import datetime

from openpyxl import Workbook
from openpyxl import load_workbook


    
		
import tkinter as tk

#################################################################3

# Openpyxl searching data functions
# reference: 
# get worksheet by name: https://stackoverflow.com/questions/36814050/openpyxl-get-sheet-by-name

# code source: https://stackoverflow.com/questions/50491839/python-openpyxl-find-strings-in-column-and-return-row-number
# wb = load_workbook(filename = 'empty_book.xlsx')
# ws = wb.active
# col_idx, row = search_value_in_col_index(ws, "_my_search_string_")

def search_value_in_column(ws, search_string, column="A"):    
    for row in range(1, ws.max_row + 1):
        coordinate = "{}{}".format(column, row)        
        if ws[coordinate].value == search_string:             
            return column, row   
    return column, None


def search_value_in_col_idx(ws, search_string, col_idx=1):
    for row in range(1, ws.max_row + 1):
        if ws[row][col_idx].value == search_string:
            return col_idx, row
    return col_idx, None


def search_value_in_row_index(ws, search_string, row=1):
    for cell in ws[row]:
        if cell.value == search_string:
            return cell.column, row
    return None, row 

def get_last_new_row(ws):
    '''
    column = "A"
    for row in range(1, ws.max_row + 1):
        coordinate = "{}{}".format(column, row)        
        if ws[coordinate].value == None: 
            print("value = ")
            print(ws[coordinate].value)
            break
    return row
    WE ASSUME: There is no empty row in the database
    '''
    column,row =  search_value_in_column(ws, None, "A")
    if row==None:
        return ws.max_row+1
    return row

#######################################3333
# WorkTime

class WorkTime:
    def __init__(self, filename="hello_world.xlsx"):
        # Openpyxl
        self.data_filename = filename
        self.worker_table_name="workers"
        self.data_wb = load_workbook(filename = self.data_filename)
        
        
        #UI
        self.worktime_ui = tk.Tk()
        self.employee_num=tk.StringVar()
        self.msg=tk.StringVar()
        self.worktime_ui.geometry("720x300")
        self.worktime_ui.title("上班打卡")
        self.employee_num_lb=tk.Label(self.worktime_ui,text="請輸入員工工號:")
        self.employee_num_lb.pack()
        self.employee_entry = tk.Entry(self.worktime_ui,textvariable=self.employee_num)
        self.employee_entry.pack()
        self.login_btn = tk.Button(self.worktime_ui,text="登入",command=self.register)
        self.login_btn.pack()
        self.result_msg=tk.Label(self.worktime_ui,fg="red",textvariable=self.msg)
        self.result_msg.pack()
        self.worktime_ui.mainloop()
    def get_user_name(self, employee_num): 
        '''
            Open spreadsheet workers (ID, name) and search the matched user name
            return user name
            
        '''                
        workers_ws = self.data_wb[self.worker_table_name]        
        column, row = search_value_in_column(workers_ws, employee_num, column="A")
              
        if row == None: 
            # not found
            return None       
        coordinate = "{}{}".format("B", row)        
        return workers_ws[coordinate].value
        
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
    def save_sp(self):
        self.data_wb.save(self.data_filename)
        
    def is_no_open_log(self):
        return True
        
    def save_log_record2sp(self,user_name,current_time):
        '''
        
        '''
        employee_ws = self.data_wb[user_name] 
        row = get_last_new_row(employee_ws)
        
        # get today :
        now = datetime.now()
        
        ###### search today's record. 
        #  if no today record 
        #      save ABC 
        #  else 
        #      D is empty  Save D##################################
        
        ################# Save ABC #############333333
        # name
        coordinate = "{}{}".format("A", row) 
        employee_ws[coordinate] = user_name
        # checking there is no WRONG checkout before
        # check in time
        coordinate = "{}{}".format("B", row) 
        employee_ws[coordinate] = current_time.date()
        
        coordinate = "{}{}".format("C", row) 
        employee_ws[coordinate] = current_time.time()
        ############# save D #####################3
        coordinate = "{}{}".format("D", row) 
        employee_ws[coordinate] = current_time.time()
        
        self.save_sp()
    def check_pw(self):


        now = datetime.now()
        date = now.today()
        current_time = now.strftime("%m/%d -- %H:%M:%S")
        #print("Current Time =", current_time)
        #https://stackoverflow.com/questions/50491839/python-openpyxl-find-strings-in-column-and-return-row-number
        user_id = self.employee_num.get()
        user_name = self.get_user_name(user_id)
        if user_name == None: 
            user_name = "user not Found"        
        msg_str = "歡迎 "+user_name+" 登入,您簽到的時間是:"+current_time
        self.msg.set(msg_str)
        return user_name, now
 

        #workbook.save(self.data_filename)
    def register(self):
        user_name, now = self.check_pw()
         ############################333 
         
        self.save_log_record2sp(user_name,now)
        
         
############################################33
######## End of work time
##########################################
def main():
    
    wt = WorkTime(filename="work_log.xlsx")    

    
    

if __name__ == "__main__":
    main()
    
    
    
    


