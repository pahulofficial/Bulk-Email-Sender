from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
from PIL import ImageTk
import os
import pandas as panda
import email_func
import time
import threading


class bulk_mail_sender:
    def __init__(self, root):
        self.root = root
        self.root.title("Bulk Email Sender | Coded By Pahul Preet Singh")
        self.root.geometry("1000x550+200+50")
        self.root.resizable(False, False)
        self.root.config(bg="white")
        self.root.iconbitmap("icons/main_icon.ico")
        # Icons
        self.setting_icon = ImageTk.PhotoImage(file="icons/settings.png")
        self.mail_icon = ImageTk.PhotoImage(file="icons/mail.png")
        # self.main_ico = ImageTk.PhotoImage(file="icons/main_icon.png")

        # Header_Title() (mere leye top bar ka kee window header hotee hai mtb top mai titile not in window of tkinter)

        title = Label(self.root, text="Bulk Email Sender", image=self.mail_icon, compound=LEFT, padx=40, bg="#222A35",
                      fg="white", font=("Fixedsys", 58, "bold"))
        title.pack(fill=X, )

        desc = Label(self.root,
                     text="Only Use Excel File to Send the Bulk (multiple) Emails at once, with just one click. (Ensure the Email Column Name must be  email)",
                     bg="orange", fg="black", font=("Calibri (body)", 13))
        desc.pack(fill=X)
        #############Buttons##################
        setting_btn = Button(self.root, image=self.setting_icon, cursor="hand2", bg="#222A35",
                             activebackground="#222A35", command=self.setting_screen).place(x=900, y=5)
        self.var_choice = StringVar()
        self.single_btn = Radiobutton(self.root, text="Single", command=self.single_or_bulk, value="Single",
                                      variable=self.var_choice, font=("times new roman", 30, "bold"), bg="white",
                                      activebackground="white", fg="#262626")
        self.single_btn.place(x=50, y=150)

        self.bulk_btn = Radiobutton(self.root, text="Bulk", value="Bulk", command=self.single_or_bulk,
                                    variable=self.var_choice, font=("times new roman", 30, "bold"), bg="white",
                                    activebackground="white", fg="#262626")
        self.bulk_btn.place(x=250, y=150)

        self.var_choice.set("Single")

        # .......................
        self.to_send_mail = Label(self.root, text="To Email (Email Address)", font=("times new roman", 18),
                                  bg="white").place(x=50, y=250)
        self.subject_label = Label(self.root, text="SUBJECT", font=("times new roman", 18), bg="white").place(x=50,
                                                                                                              y=300)
        self.message = Label(self.root, text="MESSAGE", font=("times new roman", 18), bg="white").place(x=50, y=350)

        # ............................
        self.to_mail_entry = Entry(self.root, font=("times new roman", 14),
                                   bg="lightyellow")
        self.to_mail_entry.place(x=350, y=250, width=450, height=30)

        global to_subject
        self.to_subject = Entry(self.root, font=("times new roman", 14), bg="lightyellow")
        self.to_subject.place(x=350, y=300, width=600, height=30)

        self.btn_browse = Button(self.root, text="BROWSE", command=self.browse_func,
                                 font=("times new roman", 15, "bold"), bg="#262626", activebackground="#262626",
                                 activeforegroun="white", fg="white", cursor="hand2", state=DISABLED)
        self.btn_browse.place(x=810, y=250, width=120, height=35)

        global to_message
        self.to_message = Text(self.root, font=("times new roman", 12), bg="lightyellow")
        self.to_message.place(x=350, y=350, width=600, height=140)

        self.btn_send = Button(self.root, text="SEND", command=self.stuck_solution,font=("times new roman", 18, "bold"), bg="#00B0F0", activebackground="#00B0F0",
                               activeforegroun="white", fg="white", cursor="hand2").place(x=700, y=510, width=120,
                                                                                          height=30)
        self.file_exists_or_not()
        #################Status bar ###################################
        self.total_mail = Label(self.root, font=("times new roman", 28), bg="white", fg="blue")
        self.total_mail.place(x=20, y=500)

        self.sent_mail = Label(self.root, font=("times new roman", 28), bg="white", fg="green")
        self.sent_mail.place(x=230, y=500)

        self.left_mail = Label(self.root, font=("times new roman", 28), bg="white", fg="orange")
        self.left_mail.place(x=350, y=500)

        self.fail_mail = Label(self.root, font=("times new roman", 28), bg="white", fg="red")
        self.fail_mail.place(x=470, y=500)

        global btn_clear
        self.btn_clear = Button(self.root, text="CLEAR", command=self.clear1, font=("times new roman", 18, "bold"),
                                bg="#262626", activebackground="#262626", activeforegroun="white", fg="white",
                                cursor="hand2")
        self.btn_clear.place(x=830, y=510, width=120, height=30)

    def stuck_solution(self):
        self.stuck_thread = threading.Thread(target=self.send_mail)
        self.stuck_thread.daemon = True
        self.stuck_thread.start()

    def send_mail(self):
        a = len(self.to_message.get("0.1", END))
        if self.to_mail_entry.get() == "" or self.to_subject.get() == "" or a == 1:
            messagebox.showerror("Error", "All Fields are Required")
        else:
            if self.var_choice.get() == "Single":
                status = email_func.email_send_func(self.to_mail_entry.get(), self.to_subject.get(),
                                                    self.to_message.get('1.0', END), self.from_, self.pass_)
                if status == "Sent":
                    messagebox.showinfo("Sent", "Your Email has been sent")
                if status == "Failed":
                    messagebox.showerror("Failed", "Email has been not Sent, Please try again")
            if self.var_choice.get() == "Bulk":
                self.failed = []
                self.sent_count = 0
                self.failed_count = 0
                for x in self.emails:
                    status = email_func.email_send_func(x, self.to_subject.get(), self.to_message.get('1.0', END),
                                                        self.from_, self.pass_)
                    if status == "Sent":
                        self.sent_count += 1
                    if status == "Failed":
                        self.failed_count += 1
                    self.statu_bar()
                    time.sleep(1)

                messagebox.showinfo("Sent", "Your Email has been sent")

    def statu_bar(self):
        self.total_mail.config(text="Status Bar: " + str(len(self.emails)))
        self.sent_mail.config(text="Sent: " + str(self.sent_count))
        self.left_mail.config(text="Left: " + str(len(self.emails) - (self.sent_count + self.failed_count)))
        self.fail_mail.config(text="Fail: " + str(self.failed_count))
        self.total_mail.update()
        self.sent_mail.update()
        self.left_mail.update()
        self.fail_mail.update()

    def single_or_bulk(self):
        if self.var_choice.get() == "Single":
            self.btn_browse.config(state=DISABLED)
            self.to_mail_entry.delete(0, END)
            self.to_mail_entry.config(state=NORMAL)
            self.clear1()
        if self.var_choice.get() == "Bulk":
            self.btn_browse.config(state=NORMAL)
            self.to_mail_entry.delete(0, END)
            self.to_mail_entry.config(state=DISABLED)

    def clear1(self):
        self.to_mail_entry.config(state=NORMAL)
        self.to_mail_entry.delete(0, END)
        self.to_subject.delete(0, END)
        self.to_message.delete("0.1", END)
        self.var_choice.set("Single")
        self.btn_browse.config(state=DISABLED)
        self.total_mail.config(text="")
        self.sent_mail.config(text=" ")
        self.left_mail.config(text=" ")
        self.fail_mail.config(text=" ")

    def browse_func(self):
        self.open_file = filedialog.askopenfile(initialdir='/', title="Select Excel Files",
                                                filetypes=(("All Files", "*.*"), ("Excel Files", ".xlsx")))
        if self.open_file != None:
            self.data = panda.read_excel(self.open_file.name)
            # print(data['Email'])
            if 'email' in self.data.columns:
                self.emails = list(self.data['email'])
                c = []
                for i in self.emails:
                    if panda.isnull(i) == False:
                        c.append(i)
                self.emails = c
                if len(self.emails) > 0:
                    self.to_mail_entry.config(state=NORMAL)
                    self.to_mail_entry.delete(0, END)
                    self.to_mail_entry.insert(0, str(self.open_file.name.split("/")[-1]))
                    self.to_mail_entry.config(state="readonly")
                    self.total_mail.config(text="Total : " + str(len(self.emails)))
                    self.sent_mail.config(text="Sent : ")
                    self.left_mail.config(text="Left : ")
                    self.fail_mail.config(text="Fail : ")

                else:
                    messagebox.showerror("Error", "The files is not having any email", parent=self.root)
            else:
                messagebox.showerror("Error", "There is no coloumn name with email", parent=self.root)

    # setting wale button ka code
    def setting_screen(self):
        self.file_exists_or_not()
        self.root2 = Toplevel()
        self.root2.iconbitmap("icons/main_icon.ico")
        self.root2.title("Settings")
        self.root2.geometry(("700x350+350+90"))
        self.root2.resizable(False, False)
        self.root2.focus_force()
        self.root2.grab_set()
        self.root2.config(bg="white")
        title_setting = Label(self.root2, text="Email Settings", image=self.setting_icon, compound=LEFT, padx=10,
                              bg="#222A35", fg="white", font=("Fixedsys", 58, "bold"))
        title_setting.pack(fill=X, )
        desc_2 = Label(self.root2, text="Enter the Email address and password from which to send all  emails.",
                       bg="orange", fg="black", font=("Calibri (body)", 13))
        desc_2.pack(fill=X)

        self.user_email = Label(self.root2, text="Email Address", font=("times new roman", 18), bg="white").place(x=50,
                                                                                                                  y=150)
        self.user_password = Label(self.root2, text="PASSWORD", font=("times new roman", 18), bg="white").place(x=50,
                                                                                                                y=200)

        self.user_email_text = Entry(self.root2, font=("times new roman", 14), bg="lightyellow")
        self.user_email_text.place(x=200, y=150, width=400, height=30)

        self.user_password_text = Entry(self.root2, font=("times new roman", 14), bg="lightyellow", show="*")
        self.user_password_text.place(x=200, y=200, width=400, height=30)

        self.btn_clear2 = Button(self.root2, text="CLEAR", command=self.clear2_for_second_window,
                                 font=("times new roman,", 18, "bold"), bg="#262626", activebackground="#262626",
                                 activeforegroun="white", fg="white", cursor="hand2")
        self.btn_clear2.place(x=480, y=270, width=120, height=30)

        self.btn_save = Button(self.root2, text="SAVE", command=self.save_settings_root2,
                               font=("times new roman,", 18, "bold"), bg="#00B0F0", activebackground="#00B0F0",
                               activeforegroun="white", fg="white", cursor="hand2")
        self.btn_save.place(x=350, y=270, width=120, height=30)

        self.user_email_text.insert(0, self.from_)
        self.user_password_text.insert(0, self.pass_)

    def clear2_for_second_window(self):
        self.user_email_text.delete(0, END)
        self.user_password_text.delete(0, END)

    # this function file_exists_or_not make a file if files is not present or get the email and pass from the user

    def file_exists_or_not(self):
        if os.path.exists('credentials.txt') == False:
            save_data_file = open('credentials.txt', 'w')
            save_data_file.write(",")
            save_data_file.close()
        d = open('credentials.txt', 'r')
        self.email_and_pass = []
        for i in d:
            self.email_and_pass.append([i.split(",")[0], i.split(",")[1]])

        self.from_ = self.email_and_pass[0][0]
        self.pass_ = self.email_and_pass[0][1]

    def save_settings_root2(self):
        if self.user_email_text.get() == "" or self.user_password_text.get() == "":
            messagebox.showerror("Error", "All field required", parent=self.root2)

        else:
            save_data_file = open('credentials.txt', 'w')
            save_data_file.write(self.user_email_text.get() + ", " + self.user_password_text.get())
            save_data_file.close()
            messagebox.showinfo("Saved", "Your Data is Saved", parent=self.root2)
            self.file_exists_or_not()


root = Tk()

mail_sender = bulk_mail_sender(root)
root.mainloop()
