"""
Dwyer Engineering LIMS is a gui that was designed with the intention of being used as a Laboratory Information
Management System (LIMS) for Dwyer Instruments, Inc. located in Michigan City, IN. It exists in the same directory
where all the pertinent information is stored for ease of access to all necessary information. As time goes on, there
are additional components that will be added to the program such that it meets the requirements of ISO 17025 for an
acceptable LIMS.

Copyright(c) 2018, Robert Adam Maldonado.
"""

import LIMSVarConfig
import PIL.Image
from PIL import ImageTk
from tkinter import *
from tkinter import messagebox as tm
from tkinter import ttk


class LoginWindow(Frame):

    # ================================METHODS========================================= #

    def __init__(self, title, parent):
        global username, password
        Frame.__init__(self, parent)
        self.frame = Frame(parent)
        self.parent = parent
        self.frame.pack()

        parent.width = 525
        parent.height = 180
        screen_width = parent.winfo_screenwidth()
        screen_height = parent.winfo_screenheight()
        x = (screen_width / 2) - (parent.width / 2)
        y = (screen_height / 2) - (parent.height / 2)
        parent.geometry("%dx%d+%d+%d" % (parent.width, parent.height, x, y))
        parent.title(title)
        parent.focus_force()
        parent.iconbitmap("Required Images\\DwyerLogo.ico")

        # ============================VARIABLES=================================== #

        username_array = StringVar()
        password_array = StringVar()

        # ============================SUB FRAMES================================== #

        top = Frame(parent)
        top.pack(side=LEFT, padx=10)
        form = Frame(parent)
        form.pack(side=TOP)

        # ===========================MAIN LABELS================================== #

        image = PIL.Image.open("Required Images\\DwyerLogo_Calibration.jpg")
        photo = ImageTk.PhotoImage(image)

        lbl_logo = Label(top, image=photo)
        lbl_logo.image = photo
        lbl_logo.pack(fill=BOTH, pady=5)

        # Software version - Major.Minor.Maintenance.Build
        lbl_software_version = ttk.Label(form, text="Version 2.0", font=('serif', '10'))
        lbl_software_version.grid(row=0, columnspan=2, pady=8)

        lbl_username = ttk.Label(form, text="Username:", font=('serif', '10'))
        lbl_username.grid(row=1, pady=2)

        lbl_password = ttk.Label(form, text="Password:", font=('serif', '10'))
        lbl_password.grid(row=2, pady=2)

        lbl_text = Label(form)
        lbl_text.grid(row=3, columnspan=2)

        # =======================MAIN ENTRY WIDGETS================================#

        username = ttk.Entry(form, textvariable=username_array)
        username.config(width=20)
        username.focus()
        username.grid(row=1, column=1, sticky="w")

        password = ttk.Entry(form, textvariable=password_array, show="*")
        password.config(width=20)
        password.grid(row=2, column=1, sticky="w")

        # ======================MAIN BUTTON WIDGETS===============================#

        btn_login = ttk.Button(form, text="Login", width=30, command=lambda: self.lims_login())
        btn_login.bind("<Return>", lambda event: self.lims_login())
        btn_login.grid(pady=0, row=4, columnspan=2)

        # btn_register = ttk.Button(form, text="Register New User", width=30, command=lambda: self.lims_register())
        # btn_register.bind("<Return>", lambda event: self.lims_register())
        # btn_register.grid(pady=0, row=5, columnspan=2)

        btn_exit = ttk.Button(form, text="Exit", width=30, command=lambda: self.lims_exit())
        btn_exit.bind("<Return>", lambda event: self.lims_exit())
        btn_exit.grid(pady=0, row=6, columnspan=2)

    # --------------------------------------------------------------------------------#

    # # Allows New User to Create Login Credentials (Writes to .txt File)
    # def lims_register(self):
    #     self.do_nothing()
    #
    #     new_username_array = []
    #     new_password_array = []
    #
    #     for line in open("\\\\BDC5\\Dwyer Engineering LIMS\\Required Files\\Usernames and Passwords.txt",
    #                      "r").readlines():
    #         register_info = line.split()
    #         new_username_array.append(register_info[0])
    #         new_password_array.append(register_info[1])
    #
    #     if username.get() == "" or password.get() == "":
    #         tm.showerror("Registration Error", "Please complete the required fields with a username and password of "
    #                                            "your choice (fields cannot be empty.)")
    #     elif ((username.get() != "") and (username.get() not in new_username_array) and
    #           (password.get() != "" or password.get() != " ") and
    #           password.get() not in new_password_array):
    #         user_file = open("\\\\BDC5\\Dwyer Engineering LIMS\\Required Files\\Usernames and Passwords.txt", "a")
    #         user_file.write("\n")
    #         user_file.write(username.get())
    #         user_file.write(" ")
    #         user_file.write(password.get())
    #         tm.showinfo("Success!", "Account created!")
    #         user_file.close()
    #     else:
    #         for n in range(0, len(new_username_array)):
    #             if (username.get() != "") and username.get() == new_username_array[n]:
    #                 tm.showerror("Bad Entry", "Those credentials are already taken! Please use different "
    #                                           "account credentials.")
    #                 break

    # --------------------------------------------------------------------------------#

    # Opens Login File to Verify User Login Credentials
    def lims_login(self):

        account_number_array = []
        username_array = []
        password_array = []
        rights_array = []
        name_array = []
        title_array = []

        i = 1
        with open(r"Required Files\\Accounts.txt", "r") as f:
            for line in f:
                account_information = line.split(';')
                account_number_array.append(i)
                username_array.append(account_information[0])
                password_array.append(account_information[1])
                rights_array.append(account_information[2])
                name_array.append(account_information[3])
                title_array.append(account_information[4])
                i += 1

        if username.get() == "" or password.get() == "":
            tm.showerror("Login Error",
                         "Please complete the required fields with a registered username and password")
        elif username.get() not in username_array and password.get() in password_array:
            tm.showerror("Login Error", "Invalid Username! Please try logging in again.")
        elif username.get() in username_array and password.get() not in password_array:
            tm.showerror("Login Error", "Invalid Password! Please try logging in again.")

        for i in range(0, len(account_number_array)):
            if username.get() != username_array[i] and password.get() != password_array[i]:
                i += 1
            elif username.get() == username_array[i] and password.get() == password_array[i]:
                LIMSVarConfig.certificate_technician_name = name_array[i]
                LIMSVarConfig.certificate_technician_job_title = title_array[i]
                LIMSVarConfig.user_access = rights_array[i]
                self.close_root_open_home()

    # --------------------------------------------------------------------------------#

    # Closes Calibration Program
    def lims_exit(self):
        self.parent.destroy()

    # --------------------------------------------------------------------------------#

    # Close Login Screen and Open Main LIMS Window
    def close_root_open_home(self):
        self.parent.withdraw()
        from LIMSHomeWindow import AppHomeWindow
        new_window = AppHomeWindow()
        new_window.home_window()

    # --------------------------------------------------------------------------------#

    # Function that does nothing
    def do_nothing(self):
        pass

# ========================INITIALIZATION===================================#


if __name__ == '__main__':
    root = Tk()
    app = LoginWindow(title='Dwyer Instruments, Inc. - Laboratory Information Management System', parent=root)
    root.mainloop()


# =====================METHOD TO REOPEN LOGIN=============================#

class LoginWindowRestoration:

    # ================================METHODS========================================= #

    def __init__(self):
        pass

    # --------------------------------------------------------------------------------#

    # Function to restore the Login Window. Designed for use in other modules
    def root_window_restore(self):
        LoginWindow(title='Dwyer Instruments, Inc. - Laboratory Information Management System', parent=Toplevel())
