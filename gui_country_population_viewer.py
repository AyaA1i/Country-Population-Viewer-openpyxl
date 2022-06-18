# This is the graphical interface for the excel app that shows the population of states
# Authors: Youssef Morad, Aya Ali, Mariam Ayman
# Date: 08/03/2022

from tkinter import *
from console_country_population_viewer import *
from time import sleep


def destroy_last_ouput():
    try:
        output1.destroy()
    except:
        pass
    try:
        output2.destroy()
    except:
        pass


def destroy_load_message():
    try:
        not_found_label.destroy()
    except:
        pass
    try:
        found_label.destroy()
    except:
        pass


def destroy_choices():
    try:
        option1_label.destroy()
        option2_label.destroy()
        option1_button.destroy()
        option2_button.destroy()
    except:
        pass


def load_click():
    global country, sheet, found_label, not_found_label
    country = country_entry.get().lower().strip()
    destroy_last_ouput()
    destroy_load_message()
    if country not in ls:
        not_found_label = Label(window, text="Country is not available..", bg='black', fg='red', font='none 24 bold', width=27, anchor='w')
        not_found_label.place(x=1035, y=130, anchor='w')
        destroy_choices()
    else:
        found_label = Label(window, text="Country is loaded successfully..", bg='black', fg='green', font='none 24 bold')
        found_label.place(x=1035, y=130, anchor='w')
        sheet = wb[country.title()]
        label_of_choices()


def label_of_choices():
    global option1_label, option2_label, option1_button, option2_button
    option1_label = Label(window, text="Option 1: To display the population of each province and the total population of the country: ", bg='black', fg='white', font='none 20 bold')
    option1_label.place(x=5, y=230, anchor='w')

    option2_label = Label(window, text="Option 2: To display the population of each province and the total population of the country: ", bg='black', fg='white', font='none 20 bold')
    option2_label.place(x=5, y=350, anchor='w')

    option1_button = Button(window, text="Option 1", width=15, height=2, command=option1_click)
    option1_button.place(x=20, y=280, anchor='w')

    option2_button = Button(window, text="Option 2", width=15, height=2, command=option2_click)
    option2_button.place(x=20, y=400, anchor='w')


def option1_click():
    global output1
    destroy_load_message()
    destroy_last_ouput()
    output1 = Text(window, background="white", width=60, height=35, wrap=WORD, state='normal')
    output1.place(x=150, y=720, anchor="w")
    output1.insert(END, option_1(country, sheet))
    output1.config(state='disabled')


def option2_click():
    global output2
    destroy_load_message()
    destroy_last_ouput()
    output2 = Text(window, background="white", width=80, height=5, wrap=WORD, state='normal', font='none 15 bold')
    output2.place(x=50, y=530, anchor="w")
    output2.insert(END, option_2(country, sheet))
    output2.config(state='disabled')


def exit_():
    window.destroy()


def excel_program_gui():
    global window, country_entry
    # main
    window = Tk()
    window.title("Country Excel Viewer")
    window.configure(bg="black")

    # Main label
    head = Label(window, text="Countries' Population Program", bg='black', fg='orange', font='none 34 bold')
    head.place(x=960, y=50, anchor="center")

    # Country name label
    country_label = Label(window, text="Please choose a country to load its file:", bg='black', fg='white', font='none 25 bold')
    country_label.place(x=5, y=130, anchor="w")

    # create a text entry box and button to take input
    country_entry = Entry(window, width=30, bg="white", font='none 13 bold')
    country_entry.place(x=650, y=130, anchor='w')
    country_button = Button(window, text="Enter", width=10, command=load_click)
    country_button.place(x=950, y=117)

    # button exit
    exit_button = Button(window, text="Exit", command=exit_, font='none 25 bold')
    exit_button.place(x=1920, y=990, anchor='e')

    window.mainloop()


excel_program_gui()

