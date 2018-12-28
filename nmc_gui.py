from tkinter import *
from datetime import datetime, timedelta

class NateonMemoAutomationWindow:
    def __init__(self, window):
        self.window = window
        self.window.title("Welcome to LikeGeeks app")
        self.window.geometry('400x200')
        
        self.send_rcv_person_lbl = Label(self.window, text="보낸/받은사람")
        self.send_rcv_person_lbl.grid(column=0, row=0)

        self.send_rcv_person_edit = Entry(self.window, width=20)
        self.send_rcv_person_edit.grid(column=1, row=0)

        self.search_word_lbl = Label(self.window, text="검색어")
        self.search_word_lbl.grid(column=0, row=1)
        
        self.search_word_edit = Entry(self.window, width=30)
        self.search_word_edit.grid(column=1, row=1)

        self.period_lbl = Label(self.window, text="기간")
        self.period_lbl.grid(column=0, row=2)

        yesterday = StringVar()
        one_week_ago = StringVar()    

        yesterday.set((datetime.today() - timedelta(1)).strftime('%Y%m%d'))
        one_week_ago.set((datetime.today() - timedelta(8)).strftime('%Y%m%d'))

        self.period_from_edit = Entry(self.window, width=10, textvariable=one_week_ago)
        self.period_from_edit.grid(column=1, row=2)

        self.period_to_edit = Entry(self.window, width=10, textvariable=yesterday)
        self.period_to_edit.grid(column=2, row=2)

        self.selected_save_opt = IntVar()
        self.selected_save_opt.set(1)
        self.save_opt_radio1 = Radiobutton(self.window, text='전체 저장', value=1, variable=self.selected_save_opt)
        self.save_opt_radio2 = Radiobutton(self.window, text='현재 위치부터 저장', value=2, variable=self.selected_save_opt)

        self.save_opt_radio1.grid(column=0, row=3) 
        self.save_opt_radio2.grid(column=1, row=3)

        self.start_btn = Button(self.window, text='시작', command=self.start_btn_clicked)
        self.start_btn.grid(column=3, row=5)

        self.stop_btn = Button(self.window, text='중지', command=self.stop_btn_clicked)
        self.stop_btn.grid(column=4, row=5)

    def start_btn_clicked(self):
        print(self.selected_save_opt.get())

    def stop_btn_clicked(self):
        pass    


if __name__ == '__main__':
    window = Tk()
    gui = NateonMemoAutomationWindow(window)
    window.mainloop()