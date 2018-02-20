import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
from tkinter import ttk
from SendMail import SendMail
from ReadExcel import ReadExcel


class Application(tk.Tk):

    # Initialize the window
    def __init__(self):
        super().__init__()
        self.createWidgets()

    # Create the widgets
    def createWidgets(self):
        self.title('EmailSender')
        # the size of the window
        width, height = 800, 600
        # To make the window be centered
        self.geometry('%dx%d+%d+%d' % (width, height, (self.winfo_screenwidth() -
                                                       width) / 2, (self.winfo_screenheight() - height) / 2))
        # Max size of the window
        self.maxsize(800, 600)
        # Min size of the window
        self.minsize(800, 600)

        # the email part
        lable = tk.Label(self, text="SMTP Host：")
        lable.grid(row=0, column=0)
        self.hostbox = tk.Entry(self)
        self.hostbox.grid(row=0, column=1, pady=10)

        lable = tk.Label(self, text="Your address username：")
        lable.grid(row=0, column=2)
        self.userbox = tk.Entry(self)
        self.userbox.grid(row=0, column=3, pady=10)

        lable = tk.Label(
            self, text="Your address password\n(sometimes\n authorization code):")
        lable.grid(row=1, column=0, rowspan=2)
        self.pwdbox = tk.Entry(self,)
        self.pwdbox['show'] = '●'
        self.pwdbox.grid(row=1, column=1, pady=12)

        sh = ttk.Separator(self, orient=tk.HORIZONTAL)
        sh.grid(row=3, column=0, columnspan=5, sticky="we")

        lable = tk.Label(self, text="The sender's address:")
        lable.grid(row=4, column=0, pady=12)
        self.senderbox = tk.Entry(self)
        self.senderbox.grid(row=4, column=1, pady=12,
                            columnspan=3, sticky="we")

        lable = tk.Label(self, text="The receiver's address:")
        lable.grid(row=5, column=0, pady=4)
        self.receiverbox = tk.Entry(self)
        self.receiverbox.grid(row=5, column=1, pady=4,
                              columnspan=3, sticky="we")

        lable = tk.Label(self, text="The subject of your mail:")
        lable.grid(row=6, column=0, pady=10)
        self.subjectbox = tk.Text(self, height=5)
        self.subjectbox.grid(row=6, column=1, pady=10,
                             columnspan=3, sticky="we")

        lable = tk.Label(self, text="The text of your mail:")
        lable.grid(row=7, column=0, pady=10)
        self.subjectbox = tk.Text(self, height=12)
        self.subjectbox.grid(row=7, column=1, pady=10,
                             columnspan=3, sticky="we")

        sh = ttk.Separator(self, orient=tk.HORIZONTAL)
        sh.grid(row=8, column=0, columnspan=5, sticky="we")

        # the file part
        lable = tk.Label(self, text="The File's type:")
        lable.grid(row=9, column=0, pady=10)

        frame_choosetype = tk.Frame(self, width=140)
        frame_choosetype.grid(row=9, column=1, columnspan=4, sticky="we")

        v = tk.IntVar()
        v.set(1)

        type1 = tk.Radiobutton(
            frame_choosetype, variable=v, value=0, text=".txt")
        type1.grid(row=0, column=0, sticky="we")
        type2 = tk.Radiobutton(
            frame_choosetype, variable=v, value=1, text=".md")
        type2.grid(row=0, column=1, sticky="we")
        type3 = tk.Radiobutton(
            frame_choosetype, variable=v, value=2, text=".pdf")
        type3.grid(row=0, column=2, sticky="we")
        type4 = tk.Radiobutton(frame_choosetype, variable=v, value=3,
                               text=".png/.jpg/.jpeg/.gif")
        type4.grid(row=0, column=3, sticky="we")
        type5 = tk.Radiobutton(
            frame_choosetype, variable=v, value=4, text=".doc")
        type5.grid(row=0, column=4, sticky="we")
        type6 = tk.Radiobutton(
            frame_choosetype, variable=v, value=5, text=".docx")
        type6.grid(row=0, column=5, sticky="we")
        type7 = tk.Radiobutton(
            frame_choosetype, variable=v, value=6, text=".xls/.xlsx")
        type7.grid(row=0, column=6, sticky="we")

        lable = tk.Label(self, text="File's name：")
        lable.grid(row=11, column=0, pady=4)

        self.filebox = tk.Entry(self, width=40)
        self.filebox.grid(row=11, column=1, pady=4,
                          columnspan=3, sticky="we")

        button = tk.Button(self, text="select file",
                           command=self._openfile, height=1)
        button.grid(row=11, column=4, pady=4)

        button = tk.Button(self, text="send mail",
                           command=self._sendmail, width=30)
        button.grid(row=12, column=2, columnspan=2, pady=4, sticky="w")

    def _openfile(self):
        self.filebox.delete(0, tk.END)
        filename = filedialog.askopenfilename()
        self.filebox.insert(0, filename)
        self.filebox.update()

    def _sendmail(self):
        a = messagebox.askquestion(
            'Warning', 'Are you sure to send the mail?')
        if a == 'yes':
            readxlsx = ReadExcel()
            print("测试："+self.filebox.get()+"：测试")
            htmltext = readxlsx.readexcel(self.filebox.get())

            mail_host = self.hostbox.get()
            mail_user = self.userbox.get()
            mail_pass = self.pwdbox.get()
            sender = self.senderbox.get()
            receivers = self.receiverbox.get()
            subject = self.subjectbox.get("0.0", "end")
            sendemail = SendMail()
            msg = sendemail.sendemail(sender, receivers,
                                      mail_host, mail_user, mail_pass,
                                      subject, htmltext)
            messagebox.showinfo(msg)
        else:
            return
