import requests
import xlwt
from bs4 import BeautifulSoup
import tkinter as tk
from tkinter import messagebox as tm, Menu as Mn
import socket
import tkinter.ttk as ttk

# Create a new excel file
book = xlwt.Workbook()
sheet1 = book.add_sheet('sheet1', cell_overwrite_ok=True)
url = "http://oecteaching.ujs.edu.cn/default.aspx"
url2 = ""
s = requests.Session()
headers1 = []
headers2 = []
new_list1 = []
new_list2 = []
head = []
new_list = []
ID = ""


def is_connected():
    try:
        # connect to the host -- tells us if the host is actually
        # reachable
        socket.create_connection((url.split('/')[2], 80))
        return True
    except OSError:
        pass
    return False


class DisplayInfo(tk.Toplevel):
    def __init__(self, original, s):
        """Constructor"""
        self.original_frame = original
        tk.Toplevel.__init__(self)
        global url, url2
        self.title("Omega's Result Checker")
        self.wm_iconbitmap('JU_logo.ico')
        width = self.winfo_screenwidth()
        height = 600
        self.geometry("%dx%d+%d+%d" % (width, height, 0, 0))
        menu = tk.Frame(self)
        menu.grid(row=0, columnspan=2, sticky=(tk.N, tk.S, tk.W, tk.E))
        self.init_menu(menu)

        self.frame = tk.Frame(self)
        self.frame.grid(row=1, column=1, rowspan=4, columnspan=6, sticky=(tk.N, tk.S, tk.W, tk.E))
        bop = tk.Frame(self)
        bop.grid(row=1, column=0, sticky=(tk.N, tk.W))

        tt = tk.Button(bop, text="Display Timetable", font=('times', 14), relief='raised', highlightthickness=2,
                       bd=2, command=self.chk_tt)
        sv_tt = tk.Button(bop, text="Save Timetable", font=('times', 14), relief='raised', highlightthickness=2,
                          bd=2, command=self.save_tt)
        rs = tk.Button(bop, text="Display All Courses", font=('times', 14), relief='raised', highlightthickness=2,
                       bd=2, command=self.chk_rs)
        sv_rs = tk.Button(bop, text="Save All Results", font=('times', 14), relief='raised', highlightthickness=2,
                          bd=2, command=self.save_rs)
        fld = tk.Button(bop, text="Display Failed Courses", font=('times', 14), relief='raised', highlightthickness=2,
                        command=self.chk_fld)
        rs.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        fld.grid(row=1, column=0, padx=5, pady=5, sticky="ew")
        sv_rs.grid(row=2, column=0, padx=5, pady=5, sticky="ew")
        tt.grid(row=3, column=0, padx=5, pady=5, sticky="ew")
        sv_tt.grid(row=4, column=0, padx=5, pady=5, sticky="ew")
        close = tk.Button(bop, text="Exit", relief='raised', font=('times', 14), highlightthickness=2, bd=2, bg="red",
                          command=self.on_close)
        close.grid(row=5, column=0, padx=5, pady=5, sticky="ew")

    # ----------------------------------------------------------------------
    def on_close(self):
        self.destroy()
        self.original_frame.show()

    def chk_fld(self):
        global headers2, new_list2
        if headers2:
            self.create_ui(headers2)
            self.load_table(new_list2)
        else:
            tm.showerror('Error !', 'Please check all results before checking failed!')

    def save_rs(self):
        global headers1, new_list1, headers2, new_list2, ID
        if headers1:
            # writing all the courses
            sheet1.write(0, 5, "List Of All The Courses")
            ii = 0  # initial column number for each header of the second table

            for hd in headers1:
                sheet1.write(2, ii, hd)
                ii += 1
            i = 3  # initial row number for all the courses

            for lst in new_list1:
                j = 0  # column number for each data
                for data in lst:
                    sheet1.write(i, j, data)
                    j += 1
                i += 1

            # writing the failed courses
            i += 1  # row number of the Title of the second table
            sheet1.write(i, 5, "List Of Failed Courses")
            i += 2  # row number of the Headers of the second table
            ii = 0  # initial column number for each header of the second table

            for hd in headers2:
                sheet1.write(i, ii, hd)
                ii += 1
            i += 1  # initial row number for the failed courses

            for lst in new_list2:
                j = 0  # column number for each data
                for data in lst:
                    sheet1.write(i, j, data)
                    j += 1
                i += 1
            book.save('{} results.xls'.format(ID))
            tm.showinfo("Saving info", "Results saved to {}.xls!".format(ID))
        else:
            tm.showerror('Error Saving !', 'Please check all results before saving!')

    def chk_tt(self):
        global url, url2, s, head, new_list
        timetable_page_url = "http://oecteaching.ujs.edu.cn/{}{}".format(url2.split('/')[3],
                                                                         "/publicen/kebiaoall.aspx")
        timetable_page = s.get(timetable_page_url)
        soup = BeautifulSoup(timetable_page.content, 'lxml')
        tables = soup.find_all('table')  # all the tables present on the page

        time_table = []  # the list of all the data present in the first table
        # 2 ~ 3
        for data in tables[2].find_all('td'):
            time_table.append(data.text.strip())
        head = []
        for i in range(8):
            head.append(time_table[0])
            del time_table[0]

        new_list = [time_table[i:i + 8] for i in range(0, len(time_table), 8)]  # the list of all the courses
        for i in range(len(new_list)):
            for ii in range(len(new_list[i])):
                new_list[i][ii] = new_list[i][ii].replace(' ', '')
                new_list[i][ii] = new_list[i][ii].replace('\r\n', '')
                new_list[i][ii] = new_list[i][ii].replace('◇', '\r\n')
        self.create_ui(head, 120, 120)
        self.load_table(new_list)
        tm.showinfo('Success !', 'Successfully loaded timetable!')

    def save_tt(self):
        global head, new_list, ID
        if head:
            # writing all the courses
            sheet1.write(0, 5, "Timetable")
            ii = 0  # initial column number for each header of the second table

            for hd in head:
                sheet1.write(2, ii, hd)
                ii += 1
            i = 3  # initial row number for all the courses

            for lst in new_list:
                j = 0  # column number for each data
                for data in lst:
                    sheet1.write(i, j, data)
                    j += 1
                i += 1

            book.save('{} timetable.xls'.format(ID))
            tm.showinfo("Saving info", "Timetable saved to {}.xls!".format(ID))
        else:
            tm.showerror('Error Saving !', 'Please check the timetable before saving it!')

    def chk_rs(self):
        global url, url2, s, new_list1, new_list2, headers1, headers2
        results_page_url = "http://oecteaching.ujs.edu.cn/{}{}".format(url2.split('/')[3],
                                                                       "/studenten/chengji.aspx")
        results_page = s.get(results_page_url)
        soup = BeautifulSoup(results_page.content, 'lxml')
        tables = soup.find_all('table')  # all the tables present on the page

        data_list1 = []  # the list of all the data present in the first table
        for data in tables[1].find_all('td'):
            data_list1.append(data.text.strip())

        data_list2 = []  # the list of all the data present in the second table
        for data in tables[2].find_all('td'):
            data_list2.append(data.text.strip())

        headers_list = []  # the list of all the headers on the page
        for header in soup.find_all('th'):
            headers_list.append(header.text.strip())

        headers1 = [headers_list[i] for i in range(0, 10)]  # headers for the table of all the courses

        headers2 = [headers_list[i] for i in range(10, len(headers_list))]  # headers for the list of failed courses

        new_list1 = [data_list1[i:i + 10] for i in range(0, len(data_list1), 10)]  # the list of all the courses

        new_list2 = [data_list2[i:i + 11] for i in range(0, len(data_list2), 11)]  # the list of failed courses

        self.create_ui(headers1)
        self.load_table(new_list1)
        tm.showinfo('Success !', 'Successfully loaded results!')

    def create_ui(self, heads, height=50, width=50):
        style = ttk.Style(self)
        style.configure('Treeview', rowheight=height, rowwidth=width)  # SOLUTION
        tv = ttk.Treeview(self)
        tv['columns'] = heads
        tv.column("#0", anchor="w", width=0)
        for hd in heads:
            tv.heading(hd, text=hd, anchor = 'center')
            tv.column(hd, anchor='center', width=width)

        tv.grid(row=0, column=1, rowspan=6, columnspan=6,
                sticky=(tk.N, tk.S, tk.W, tk.E))

        # vertical scrollbar
        vsb = ttk.Scrollbar(self, orient="vertical", command=tv.yview)
        vsb.grid(row=1, column=2, sticky=(tk.N, tk.S, tk.W, tk.E))
        tv.configure(yscrollcommand=vsb.set)
        # horizontal scrollbar
        hsb = ttk.Scrollbar(self, orient="horizontal", command=tv.xview)
        hsb.grid(row=2, column=1, sticky=(tk.N, tk.S, tk.W, tk.E))
        tv.configure(xscrollcommand=vsb.set)
        self.treeview = tv
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(1, weight=1)

    def load_table(self, courses):
        for lst in courses:
            self.treeview.insert('', 'end', text="", values=lst)

    def init_menu(self, menu):
        menu_bar = Mn(menu)
        self.config(menu=menu_bar)

        display_menu = Mn(menu_bar)
        display_menu.add_command(label="Timetable", command=self.chk_tt)
        display_menu.add_command(label="All Scores", command=self.chk_rs)
        display_menu.add_command(label="Failed Courses", command=self.chk_fld)
        display_menu.add_command(label="Exit", command=self.on_close)
        menu_bar.add_cascade(label="Display", menu=display_menu)

        save_menu = Mn(menu_bar)
        save_menu.add_command(label="Save Timetable", command=self.save_tt)
        save_menu.add_command(label="Save All Scores", command=self.chk_rs)
        save_menu.add_command(label="Exit", command=self.on_close)
        menu_bar.add_cascade(label="Save", menu=save_menu)

        exit_menu = Mn(menu_bar)
        exit_menu.add_command(label="Exit", command=self.on_close)
        menu_bar.add_cascade(label="Exit", menu=exit_menu)


class Login(object):
    # ----------------------------------------------------------------------
    def __init__(self, parent):
        """Constructor"""
        self.root = parent
        self.root.title("Omega's Result Checker")
        self.root.wm_iconbitmap('JU_logo.ico')
        width = 400
        height = 280
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width / 2) - (width / 2)
        y = (screen_height / 2) - (height / 2)
        self.root.geometry("%dx%d+%d+%d" % (width, height, x, y))
        self.root.resizable(0, 0)
        self.frame = tk.Frame(parent)
        self.frame.pack()
        lbl_title = tk.Label(self.frame, text="Login", font=('times', 15, "bold"))
        lbl_title.pack()
        self.root.bind('<Return>', self.check_login)
        tk.Label(self.frame, text="Student ID: ", font=('times', 14)).pack(padx=30, pady=5)
        self.user = tk.Entry(self.frame, width=30)
        self.user.pack()
        tk.Label(self.frame, text="Password: ", font=('times', 14)).pack(padx=30, pady=5)
        self.pwd = tk.Entry(self.frame, width=30, show="*")
        self.pwd.pack()
        self.login = tk.Button(self.frame, text="Login", relief='raised', highlightthickness=3, bd=5,
                          bg="gray", fg="blue", state=tk.DISABLED)
        self.login.bind('<Button-1>', self.check_login)
        self.login.pack(padx=30, pady=12)
        close = tk.Button(self.frame, text="Close", relief='raised', highlightthickness=1, bd=5,
                          bg="red", fg="blue", command=self.on_close)
        close.pack(padx=30, pady=12)
        if is_connected():
            tk.Label(self.frame, text="[+] Connected to Internet!", font=('times', 14, "italic")).pack(padx=0, pady=1)
            self.login.config(state="normal")
        else:
            tk.Label(self.frame, text="[!!] Not connected, please check your connection.",
                     font=('times', 14, "italic")).pack(padx=0, pady=1)
            tm.showerror("Network error", "You are not connected to internet!")

    # ----------------------------------------------------------------------
    def hide(self):
        self.root.withdraw()

    # ----------------------------------------------------------------------
    def check_login(self, event):
        global ID
        ID = self.user.get()
        pwd = self.pwd.get()
        global url, url2, s
        if ID == "" or pwd == "":
            tm.showerror("Empty fields", "Please complete the required fields!")
            return
        with s:
            get1 = s.get(url)
            url2 = get1.url
            soup0 = BeautifulSoup(get1.content, 'xml')
            token = str(soup0.find('input')['value'])  # the security cookie of the website
            payload = {'TextBox1': ID, 'TextBox2': pwd, '__VIEWSTATE': token, 'Button1': '登陆', 'js': 'RadioButton3'}
            r = s.post(url2, data=payload, headers={"Referer": "http://oecteaching.ujs.edu.cn/main/main.aspx"})
            r = r.text.split(' ')[0][1]
        if r == "h":
            tm.showinfo("Login info", "Successfully Logged In!")
            self.user.delete(0, tk.END)
            self.pwd.delete(0, tk.END)
            self.hide()
            subFrame = DisplayInfo(self, s)
        else:
            tm.showerror("Login error", "Incorrect ID/Password")
            return

    # ----------------------------------------------------------------------
    def show(self):
        self.root.update()
        self.root.deiconify()
        return

    # ----------------------------------------------------------------------
    def on_close(self):
        root.destroy()
        return


if __name__ == "__main__":
    root = tk.Tk()
    app = Login(root)
    root.mainloop()
