# Import customtkinter module
import time
import tkinter.messagebox
import customtkinter as ctk
import threading
from tkinter import filedialog
import pywhatkit
import pandas as pd
import phonenumbers

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("dark-blue")


# Create App class
class App(ctk.CTk):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.title("WhatsApp Automator V1.0.0")
        self.geometry("750x500")
        self.resizable(False, False)
        # ==== GUI Start ====
        custom_font = ("Roboto", 15)
        self.contact_fl = None
        self.msg_fl = None
        self.choice=0
        self.cancel=False

        self.contact = ctk.CTkButton(self, text="Choose Contacts", font=custom_font, command=self.choose_contacts)
        self.contact.grid(row=0, column=0, padx=15, pady=15)
        self.chosen_contact_file = ctk.CTkEntry(self, width=250, state="disabled")
        self.chosen_contact_file.grid(row=0, column=1)
        self.columnconfigure(1, weight=1)

        self.img = ctk.CTkButton(self, text="Choose Image", font=custom_font, command=self.choose_img)
        self.img.grid(row=1, column=0)

        self.chosen_img_file = ctk.CTkEntry(self, width=250, state="disabled")
        self.chosen_img_file.grid(row=1, column=1)

        self.msg = ctk.CTkButton(self, text="Choose Message", font=custom_font, command=self.choose_msg)
        self.msg.grid(row=1, column=2, padx=15)

        self.prg_bar = ctk.CTkProgressBar(self, width=600, height=20, fg_color="white", progress_color="forestgreen")
        self.prg_bar.grid(row=2, column=0, columnspan=3, pady=15)
        self.prg_bar.set(0)

        self.text_area = ctk.CTkTextbox(self, width=600, height=300, state="disabled", font=custom_font)
        self.text_area.grid(row=3, column=0, columnspan=3)

        self.Frame = ctk.CTkFrame(self, fg_color="transparent")
        self.Frame.grid(row=4, column=0, columnspan=3, pady=15)

        self.start_btn = ctk.CTkButton(self.Frame, text="Start", command=self.start)
        self.start_btn.grid(row=0, column=0, padx=(0, 50))

        self.stop_btn = ctk.CTkButton(self.Frame, text="Stop", fg_color="orangered", hover_color="red",
                                      command=self.stop)
        self.stop_btn.grid(row=0, column=1, padx=(50, 0))

        self.display_instructions()

    def display_instructions(self):
        self.write_text("""*** Instructions ***\n
* Choose your Contacts. Make sure they are in an Excel file and the first cell has \n written "Phone Number" in it. \n
* If you want to send a Text message, click the "Message Button". Choose a TXT file\n  which contains the message in it. Make sure the entire message is in one line!\n	
* If you want to send an Image, click the "Image Button". PNG and JPG files are\n chosen by default.\n
*** Credits: Author of this Software is """, "black")
        self.write_text("Edita Rexhepi",'red')
        self.write_text(". Contact Info: ", 'black')
        self.write_text("edita.rexhepi@ogr.sakarya.edu.tr\n \n", 'red')

    def write_text(self, text, color):
        self.text_area.tag_config('blue', foreground="blue")
        self.text_area.tag_config('black', foreground="black")
        self.text_area.tag_config('red', foreground="red")
        self.text_area.tag_config('green', foreground="green")
        self.text_area.tag_config('white', foreground="white")
        self.text_area.configure(state="normal")
        self.text_area.insert("insert", text, color)
        self.text_area.configure(state="disabled")
        self.text_area.see("end")

    def choose_contacts(self):
        filetypes = (('Excel Files', '*.xlsx'), ('All files', '*.*'))
        file_path = filedialog.askopenfilename(filetypes=filetypes)
        if len(file_path) == 0:
            tkinter.messagebox.showwarning(title="Warning", message="No file has been chosen!")
            self.chosen_contact_file.configure(state="normal")
            self.chosen_contact_file.delete(0, "end")
            self.chosen_contact_file.configure(state="disabled")
            self.contact_fl = None
        else:
            file_name = file_path.split("/")[-1]
            self.chosen_contact_file.configure(state="normal")
            self.chosen_contact_file.delete(0, "end")
            self.chosen_contact_file.insert(0, file_name)
            self.chosen_contact_file.configure(state="disabled")
            self.contact_fl = file_path

    def choose_img(self):
        filetypes = (('Image Files', '*.jpg;*.png'), ('All files', '*.*'))
        file_path = filedialog.askopenfilename(filetypes=filetypes)
        if len(file_path) == 0:
            tkinter.messagebox.showwarning(title="Warning", message="No file has been chosen!")
            self.chosen_img_file.configure(state="normal")
            self.chosen_img_file.delete(0, "end")
            self.chosen_img_file.configure(state="disabled")
            self.msg_fl = None
        else:
            file_name = file_path.split("/")[-1]
            self.chosen_img_file.configure(state="normal")
            self.chosen_img_file.delete(0, "end")
            self.chosen_img_file.insert(0, file_name)
            self.chosen_img_file.configure(state="disabled")
            self.msg_fl = file_path
            self.choice=1

    def choose_msg(self):
        filetypes = (('Text Files', '*.txt'), ('All files', '*.*'))
        file_path = filedialog.askopenfilename(filetypes=filetypes)
        if len(file_path) == 0:
            tkinter.messagebox.showwarning(title="Warning", message="No file has been chosen!")
            self.chosen_img_file.configure(state="normal")
            self.chosen_img_file.delete(0, "end")
            self.chosen_img_file.configure(state="disabled")
            self.msg_fl = None
        else:
            file_name = file_path.split("/")[-1]
            self.chosen_img_file.configure(state="normal")
            self.chosen_img_file.delete(0, "end")
            self.chosen_img_file.insert(0, file_name)
            self.chosen_img_file.configure(state="disabled")
            with open(file_path,'r',encoding='utf-8') as file:
                self.msg_fl = file.read()
            self.choice = 2

    def start(self):
        self.cancel = False
        t1 = threading.Thread(target=self.start_thread)
        t1.start()

    def start_thread(self):
        if self.contact_fl == None or self.msg_fl == None:
            tkinter.messagebox.showwarning(title="Warning",
                                           message="You did not choose contacts and/or mesage to be sent!")
        else:
            df = pd.read_excel(self.contact_fl)
            div = len(df.index)
            self.start_val=0
            self.total_contacts = len(df.index)
            for index, row in df.iterrows():
                self.total_contacts-=1
                if self.cancel == True:
                    exit()
                    
                phone_number = row['Phone Number']
                # Format the Phone Number properly
                my_number = phonenumbers.parse('+' + str(phone_number))
                formated = str(my_number).split()
                selected_elements = ['+', formated[2], formated[-1]]
                phone_number_formated = (' '.join(selected_elements))
                time.sleep(5)
                try:
                    if self.choice == 1:
                        pywhatkit.sendwhats_image(str(phone_number_formated), self.msg_fl, tab_close=True, wait_time=7)
                    elif self.choice == 2:
                        pywhatkit.sendwhatmsg_instantly(str(phone_number_formated), self.msg_fl, tab_close=True,
                                                        wait_time=7)
                    self.write_text(text=f"Message sent to: {phone_number_formated}\n", color="green")
                    self.write_text(text=f"{self.total_contacts} Contacts Left\n \n", color="green")
                except :
                    self.write_text(text=f"Failed to send message to {phone_number_formated}\n \n", color="red")
                    self.write_text(text=f"{self.total_contacts} Contacts Left\n \n", color="green")
                time.sleep(1)
                self.start_val += 1/div
                self.prg_bar.set(self.start_val)
            self.write_text(text="Finished\n", color="blue")

    def stop(self):
        t2=threading.Thread(target=self.cancel_prc)
        t2.start()

    def cancel_prc(self):
        self.cancel=True
        self.write_text(text="*** Process Stopped! ***\n", color="red")
        self.write_text(text="Finishing Current Contact\n", color="red")


if __name__ == "__main__":
    app = App()
    app.mainloop()
