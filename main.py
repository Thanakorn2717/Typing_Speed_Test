from tkinter import *
from tkinter import messagebox
import time
import math
import pandas
import openpyxl
import random
from openpyxl.reader.excel import load_workbook

RED = "#e7305b"
GREEN = "#9bdeac"
VANILLA = "#fddfaf"
FONT_NAME = "Courier"
IS_ENTER = True
sentence = ""
last_sentence = ""
time_used = ""
is_retry = False
wrong_sentence = ""
global my_timer

sentence_df = pandas.read_excel("sentence.xlsx", sheet_name="sentence")
all_sentences = sentence_df["sentence"].tolist()


def start_stop():
    global IS_ENTER
    global sentence
    global last_sentence
    global time_used
    global wrong_sentence
    global is_retry

    if IS_ENTER:
        start_btn.config(text="Stop!", font=(FONT_NAME, 20, "bold"), bg=RED)
        entry_word.delete("1.0", END)

        if is_retry is False:
            sentence = random.choice(all_sentences)
            while sentence == last_sentence:  # make sure not to show duplicate sentence
                sentence = random.choice(all_sentences)
        else:
            sentence = wrong_sentence
            is_retry = False

        canvas.itemconfig(random_word, text=sentence)
        count_down(1)

        IS_ENTER = False

    else:
        start_btn.config(text="Start!", font=(FONT_NAME, 20, "bold"), bg=GREEN)
        answer = entry_word.get("1.0", END)[:-1]  # remove a newline character

        if answer != sentence:
            # only the right answer will be recorded.
            messagebox.showerror(title="Incorrect", message="Your answer is incorrect. Please try again.")
            canvas.itemconfig(timer, text="00:00")
            entry_word.delete("1.0", END)
            window.after_cancel(my_timer)
            wrong_sentence = sentence
            is_retry = True

        else:
            messagebox.showinfo(title="Correct!", message="Correct +1 point!")
            sentence_id = sentence_df.loc[sentence_df["sentence"] == answer, "sentence_id"].to_string(index=False)

            record = {"sentence": [answer],
                      "sentence_id": [sentence_id],
                      "typing_time": [time_used]}
            record_df = pandas.DataFrame(record)

            with pandas.ExcelWriter("sentence.xlsx", mode='a', if_sheet_exists='overlay', engine='openpyxl') as writer:
                record_df.to_excel(writer, sheet_name='summary', startrow=writer.sheets['summary'].max_row, index=False,
                                   header=False)

            last_sentence = sentence
            count_down(0)

        IS_ENTER = True


def on_entry_click(event):
    if entry_word.get("1.0", END).strip() == "Type here...":
        entry_word.delete("1.0", END)


def adjust_wrap_length(event):
    canvas_width = canvas.winfo_width()
    canvas.itemconfig(random_word, width=canvas_width - 20)  # Subtracting a little to account for padding
    adjust_canvas_height()


def adjust_canvas_height():
    # Calculate the bounding box of the text item
    bbox = canvas.bbox(random_word)  # bbox represent to the total height of Canvas after the random word appears
    # print(bbox)
    if bbox:
        # Adjust the canvas height based on the text bounding box
        canvas_height = bbox[3]  # bbox[3] means The y-coordinate of the bottom-right corner of the bounding box.
        canvas.config(height=canvas_height)
        window.grid_rowconfigure(1, minsize=canvas_height)  # row 1 is for Canvas


def count_down(count):
    global my_timer
    global time_used
    count_min = math.floor(count / 60)
    count_sec = count % 60

    if count_sec < 10:
        count_sec = f"0{count_sec}"

    canvas.itemconfig(timer, text=f"{count_min}:{count_sec}")
    time_used = canvas.itemcget(timer, "text")
    if count != 0:
        my_timer = window.after(1000, count_down, count + 1)
    else:
        window.after_cancel(my_timer)


window = Tk()
window.title("Pomodoro")
window.config(padx=50, pady=50, bg=VANILLA)

window.after(1000, )

canvas = Canvas(width=500, height=300, bg=VANILLA, highlightthickness=0)
title = canvas.create_text(250, 30, text="Typing Speed Test", fill="black", font=(FONT_NAME, 30, "bold"))
timer = canvas.create_text(250, 130, text="00:00", fill="blue", font=(FONT_NAME, 30, "bold"))
sentence_title = canvas.create_text(10, 200, anchor="nw", text="Sentence:", fill="black", font=(FONT_NAME, 20, "bold"))
random_word = canvas.create_text(10, 240, anchor="nw", text=sentence, fill="blue", font=(FONT_NAME, 15, "bold"))
canvas.grid(column=1, row=1)

entry_word = Text(height=4, width=40, font=(FONT_NAME, 15, "bold"))
# Add some text to begin with
entry_word.insert(END, "Type here...")
entry_word.grid(column=1, row=2, pady=10)

start_btn = Button(text="Start!", font=(FONT_NAME, 20, "bold"), bg=GREEN, command=start_stop)
start_btn.grid(column=1, row=3, pady=20)

window.bind('<Configure>', adjust_wrap_length)
# window.bind('<Return>', start_stop)
entry_word.bind("<FocusIn>", on_entry_click)

window.mainloop()
