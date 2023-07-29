import tkinter
from tkinter import ttk
import pandas as pd
from tkinter import messagebox
window =tkinter.Tk()
window.title('timetable data entry')
frame=tkinter.Frame(window)
frame.pack()
list_hall_no=[1, 8, 10, 101, 102, 103, 104, 105, 106, 107, 108,
              109, 110, 111, 201, 202, 203, 206, 207, 208, 209,
              210, 211, 212, 213, 301, 302, 304, 306, 309]


list_div=['A1', 'A2', 'A3', 'B1,B2', 'D', 'E', 'F', 'G', 'H', 'I',
          'J(AIIMS)', 'K(AIIMS)', 'M(AIIMS)', 'A1-IIT', 'A2-IIT',
          'A3', 'B', 'C-GEN', 'D-PCB', 'E-PCB', 'F-PCB', 'G-PCB',
          'H-PCB', 'I-AIIMS', 'J-AIIMS', 'I-1', 'I-2', 'I-3', 'I-4', 'I-5']


listx=['AKG', 'AVM', 'ABK', 'BII', 'BIII', 'B', 'CS', 'CMS', 'CIII',
       'CII', 'C', 'DK', 'DPS', 'DK', 'ELE', 'GSG', 'G', 'HR', 'KP', 'KRS', 'MHB',
       'MII', 'M', 'NPS', 'NMP', 'OM', 'PK', 'PBD', 'PRS', 'RAVI', 'RK', 'RD', 'RPB',
       'SSK', 'SPT', 'SAHU', 'SSS', 'SBG', 'SRM', 'SSP', 'SPK', 'SBG', 'SKS', 'SRB', 'SSP',
       'SRS', 'STU', 'SD', 'UBK', 'UBR', 'VP', 'V', 'VD', 'VNY', 'VAD', 'VSB', 'YP', 'YKJ', 'ZAK']

timetable_data = []

def data_extraction():
    div = division_name_combobox.get()
    hall_no = hall_no_combobox.get()

    lectures = [(f"Lecture {i+1}", lecture_combobox.get()) for i, lecture_combobox in enumerate([
        first_lec_combobox, second_lec_combobox, third_lec_combobox,
        fourth_lec_combobox, fifth_lec_combobox, sixth_lec_combobox,
        seventh_lec_combobox, eight_lec_combobox, nine_lec_combobox,
        ten_lec_combobox, elevn_lec_combobox, twlv_lec_combobox
    ])]

    print('Div:', div, 'Hall No:', hall_no)
    for lecture, teacher in lectures:
        print(lecture + ":", teacher)

    try:
        existing_df = pd.read_excel('timetable_output.xlsx')
    except FileNotFoundError:
        existing_df = pd.DataFrame()

    # Add new data to the existing DataFrame
    data = {'Division': [div], 'Hall No.': [hall_no]}
    data.update({lecture: [teacher] for lecture, teacher in lectures})
    new_data_df = pd.DataFrame(data)
    updated_df = pd.concat([existing_df, new_data_df], ignore_index=True)

 
    updated_df.to_excel('timetable_output.xlsx', index=False)

    print('∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞∞')



def save_to_excel():
    df = pd.DataFrame(timetable_data)
    df.to_excel("timetable_data.xlsx", index=False)


style=ttk.Style(window)
window.tk.call("source","forest-light.tcl")
window.tk.call("source","forest-dark.tcl")
style.theme_use('forest-dark')
##############pt1
user_timetable_entry_frame1=tkinter.LabelFrame(frame,text='Timetable entry UI')
user_timetable_entry_frame1.grid(row=0,column=0,padx=20,pady=10)

division_name_label=tkinter.Label(user_timetable_entry_frame1,text='Division')
division_name_combobox=ttk.Combobox(user_timetable_entry_frame1,values=list_div)
division_name_label.grid(row=0,column=0)
division_name_combobox.grid(row=1,column=0)

hall_no_label=tkinter.Label(user_timetable_entry_frame1,text='Hall.No')
hall_no_combobox=ttk.Combobox(user_timetable_entry_frame1,values=list_hall_no)
hall_no_label.grid(row=0,column=1)
hall_no_combobox.grid(row=1,column=1)


lectures_frame=tkinter.LabelFrame(frame)
lectures_frame.grid(row=1,column=0,sticky='news',padx=20,pady=10)


first_lec_label=tkinter.Label(lectures_frame,text='Lecture(7:00-8:00)')
first_lec_label.grid(row=0,column=0)
first_lec_combobox=ttk.Combobox(lectures_frame,values=listx)
first_lec_combobox.grid(row=0,column=1)

second_lec_label=tkinter.Label(lectures_frame,text='Lecture(8:00-9:00)')
second_lec_label.grid(row=0,column=2)
second_lec_combobox=ttk.Combobox(lectures_frame,values=listx)
second_lec_combobox.grid(row=0,column=3)

third_lec_label=tkinter.Label(lectures_frame,text='Lecture(9:00-10:00)')
third_lec_label.grid(row=0,column=4)
third_lec_combobox=ttk.Combobox(lectures_frame,values=listx)
third_lec_combobox.grid(row=0,column=5)

fourth_lec_label=tkinter.Label(lectures_frame,text='Lecture(10:00-11:00)')
fourth_lec_label.grid(row=0,column=6)
fourth_lec_combobox=ttk.Combobox(lectures_frame,values=listx)
fourth_lec_combobox.grid(row=0,column=7)

fifth_lec_label=tkinter.Label(lectures_frame,text='Lecture(11:00-12:00)')
fifth_lec_label.grid(row=1,column=0)
fifth_lec_combobox=ttk.Combobox(lectures_frame,values=listx)
fifth_lec_combobox.grid(row=1,column=1)

sixth_lec_label=tkinter.Label(lectures_frame,text='Lecture(12:00-13:00)')
sixth_lec_label.grid(row=1,column=2)
sixth_lec_combobox=ttk.Combobox(lectures_frame,values=listx)
sixth_lec_combobox.grid(row=1,column=3)

seventh_lec_label=tkinter.Label(lectures_frame,text='Lecture(14:00-14:45)')
seventh_lec_label.grid(row=1,column=4)
seventh_lec_combobox=ttk.Combobox(lectures_frame,values=listx)
seventh_lec_combobox.grid(row=1,column=5)

eight_lec_label=tkinter.Label(lectures_frame,text='Lecture(14:00-15:00)')
eight_lec_label.grid(row=1,column=6)
eight_lec_combobox=ttk.Combobox(lectures_frame,values=listx)
eight_lec_combobox.grid(row=1,column=7)

nine_lec_label=tkinter.Label(lectures_frame,text='Lecture(14:45-15:45)')
nine_lec_label.grid(row=2,column=0)
nine_lec_combobox=ttk.Combobox(lectures_frame,values=listx)
nine_lec_combobox.grid(row=2,column=1)

ten_lec_label=tkinter.Label(lectures_frame,text='Lecture(15:00-15:45)')
ten_lec_label.grid(row=2,column=2)
ten_lec_combobox=ttk.Combobox(lectures_frame,values=listx)
ten_lec_combobox.grid(row=2,column=3)

elevn_lec_label=tkinter.Label(lectures_frame,text='Lecture(15:45-16:45)')
elevn_lec_label.grid(row=2,column=4)
elevn_lec_combobox=ttk.Combobox(lectures_frame,values=listx)
elevn_lec_combobox.grid(row=2,column=5)

twlv_lec_label=tkinter.Label(lectures_frame,text='Lecture(16:45-17:45)')
twlv_lec_label.grid(row=2,column=6)
twlv_lec_combobox=ttk.Combobox(lectures_frame,values=listx)
twlv_lec_combobox.grid(row=2,column=7)


button_final_label = tkinter.Button(frame, text='Continue', command=lambda: [data_extraction(), save_to_excel()])
button_final_label.grid(row=4,column=0,padx=5,pady=7)



for widget in lectures_frame.winfo_children():
    widget.grid_configure(padx=7,pady=16,)

window.mainloop()


#fu
