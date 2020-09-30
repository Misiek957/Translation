from tkinter import Button, Tk, filedialog, Radiobutton, StringVar, Label, Entry, END, LabelFrame
import tkinter.font as font
import pandas as pd
import docx2txt

# Initiate the GUI.
root = Tk()
root.title("Translation Analysis")
root.geometry('780x400')

# Define initialization variables.
select = "Sociology"
position = []
position_index = -1
change = 0
score = 0
table_remaining = pd.read_excel("working_wordlist_" + select + ".xlsx", sheet_name=select)
original_table = pd.read_excel("working_wordlist_" + select + ".xlsx", sheet_name=select)
previous_phrase = 'first'
ttFile, phrase, final, pos, neg = None, None, None, None, None


def open_file():
    # Function to import the selected word document into the script as ttFile.
    global ttFile
    global position_index
    root.filename = filedialog.askopenfilename(initialdir="", title="Select A fle", filetypes=(("Text documents", "*.docx"), ("all files", "*.")))
    entry_file.delete(0, END)
    entry_file.insert(0, root.filename)
    ttFile = root.filename
    position_index = -1


def radio_change(value):
    # Global change of variable to the change of selected subject radio button.
    global select
    select = value


def score_change(points):
    # Global change of variable to the change of selected score radio button.
    global score
    score = points


def next_phrase(remaining, score_add, new_phrase):
    # Function for after the button 'Next' is pressed. The function prepares to display the next phrase out of the list
    # of the phrases not previously found by the t_score() function. Simultaneously the previous entry within the entry
    # box is stored to its corresponding previous phrase in relation to the applied score, and saved in the appropriate
    # cell, updating the excel file
    global phrase
    global position
    global position_index
    global match
    global final
    global pos
    global neg
    global original_table
    global previous_phrase
    print('POSITION:')
    print(position)
    print('position_index = %d' % position_index)
    finish_label = Label(root, text="                 ")  # Blank finish label for when the work is not yet finished.
    finish_label['font'] = font.Font(size=14)
    finish_label.grid(row=12, column=0, columnspan=2)
    entry.delete(0, END)
    if previous_phrase == new_phrase:
        # A sanity check checking for a double click of the 'Next' button, in which case it would repeat itself.
        print('double input')
        return

    # Save the entered phrase present in the previous state's entry box.
    if int(score_add) > 0:
        # Check if the score entered is either positive.
        pos += 1    # Add score.
        for i in range(2, 50):
            # Scan across the positive columns.
            if position_index == -1:
                # position_index is -1 during the first scan, therefore the entry is invalid and the loop is broken.
                break
            elif original_table['1' + 3 * '0' + str(i)][position[position_index]] == ' ':
                # Check if the cell in the positive '1000x' region is unoccupied.
                original_table['1' + 3 * '0' + str(i)][position[position_index]] = new_phrase   # Add the phrase from entry box.
                break
        else:
            print('more cells required')  # Error message for when the amount of cells filled is exceeded.
    elif int(score_add) < 0:
        # Check if the score entered is negative.
        neg -= 1
        for i in range(2, 50):
            if original_table['-1' + 3 * '0' + str(i)][position[position_index]] == ' ':
                # Check if the cell in the positive '-1000x' region is unoccupied.
                original_table['-1' + 3 * '0' + str(i)][position[position_index]] = new_phrase  # Add the phrase from entry box.
                break
        else:
            print('more cells required')  # Error message for when the amount of cells filled is exceeded.
    elif int(score_add) == 0:
        # Check if the score entered is 0.
        for i in range(2, 50):
            if original_table['1' + 3 * '0' + str(i)][position[position_index]] == ' ':
                # Check if the cell in the positive '0000x' region is unoccupied.
                original_table['0' + 3 * '0' + str(i)][position[position_index]] = new_phrase   # Add the phrase from entry box.
                break
        else:
            print('more cells required')  # Error message for when the amount of cells filled is exceeded.
        pass
    original_table.to_excel("working_wordlist_"+select+".xlsx", sheet_name=select)  # Save the changes into the corresponding excel file.
    final += int(score_add)  # Sum the added value to the final score.
    # Setup the next phrase to be displayed on the GUI.
    previous_phrase = new_phrase
    position_index += 1
    phrase = remaining[0][position[position_index]]
    match = Label(root, text=phrase, padx=80)
    match.grid(row=7, column=0)
    result_label = Label(frame, text="    The result = " + str(final) + "   ")
    detail_label = Label(frame, text="Positive hits = " + str(pos) + "  Negative hits = " + str(neg))
    result_label['font'] = font.Font(size=16)
    detail_label['font'] = font.Font(size=10)
    result_label.grid(row=5, column=0, columnspan=2)
    detail_label.grid(row=6, column=0)
    if position_index == (len(position)-1):
        # Check for final condition.
        finish_label = Label(root, text="Finished")
        finish_label['font'] = font.Font(size=14)
        finish_label.grid(row=12, column=0, columnspan=2)
        match = Label(root, text="                ", padx=80)
        match.grid(row=7, column=0)
        position_index = -1


def process_tt(file_location):
    # Check for a presence of the text within the script, then proceed into score analysis.
    if ttFile is None:
        ''' insert warning '''
        print("no file selected")
    else:
        t_text = docx2txt.process(file_location)  # stores the selected TT document file location
        t_score(t_text)


def t_score(text):
    # Using the string from the word document, using a string-matching algorithm check for the amount of phrases present
    # in the wordlist that match with the ones present in the text. For when the phrases match, the corresponding score
    # is recorded and tallied up to return the total positive and negative hits and a list of remaining equivalent
    # phrases that have not been found.
    global table_remaining
    global original_table
    global position
    global position_index
    # Import the corresponding wordlist for the analysis.
    table = pd.read_excel("working_wordlist_" + select + ".xlsx", sheet_name=select)
    table_remaining = table
    original_table = table
    pos_score = 0
    neg_score = 0
    # Check negative occurrences.
    for idx, val in enumerate(table[-1]):
        # Scan the items in the -1 column corresponding to the score -1.
        if val == ' ':
            continue
        elif str(val).lower() in str(text).lower():
            # String match comparison of the phrase (val) within text, in lower case to minimalise error due to grammar.
            print("%s is in the text = -1" % val)
            table_remaining = table_remaining.drop(idx)
            neg_score -= 1
    minus_list = []
    for i in range(2, 50):
        minus_list.append('-1' + 3 * '0' + str(i))
    for j in minus_list:
        # Scan the additional appended -1 columns, '-1000x'.
        for idx, val in enumerate(table[j]):
            if val == ' ':
                continue
            if str(val).lower() in str(text).lower():
                print("%s is in the text = -1" % val)
                table_remaining = table_remaining.drop(idx)
                neg_score -= 1
    # Check positive occurrences.
    for idx, val in enumerate(table[1]):
        # Scan the items in the 1 column corresponding to the score +1.
        if val == ' ':
            continue
        elif str(val).lower() in str(text).lower():
            # String match comparison of the phrase (val) within text, in lower case to minimalise error due to grammar.
            print("%s is in the text = +1" % val)
            table_remaining = table_remaining.drop(idx)
            pos_score += 1
    plus_list = []
    for i in range(2, 50):
        plus_list.append('1' + 3 * '0' + str(i))
    for j in plus_list:
        for idx, val in enumerate(table[j]):
            # Scan the additional appended -1 columns, '-1000x'.
            if val == ' ':
                continue
            elif str(val).lower() in str(text).lower():
                print("%s is in the text = +1" % val)
                table_remaining = table_remaining.drop(idx)
                pos_score += 1
    # Check for zero score occurrences.
    zero_list = []
    for i in range(2, 50):
        zero_list.append('0' + 3 * '0' + str(i))
    for j in zero_list:
        for idx, val in enumerate(table[j]):
            if val == ' ':
                continue
            elif str(val).lower() in str(text).lower():
                print("%s is in the text = +0" % val)
                table_remaining = table_remaining.drop(idx)
    disp_score(pos_score, neg_score)
    position = table_remaining.index
    if table_remaining.index is not None:
        # Final condition negative check, if it is not met, a label prompting to continue searching is placed.
        detail_label_2 = Label(frame, text="No full match, press next to continue matching")
        detail_label_2['font'] = font.Font(size=12)
        detail_label_2.grid(row=6, column=1, columnspan=2)


def disp_score(positive, negative):
    # Function to display the updated score on the GUI.
    global final
    global pos
    global neg
    pos, neg = positive, negative
    print("pos_score = %d" % pos)
    print("neg_score = %d" % neg)
    final = pos + neg
    print("final score = %d" % final)
    result_label = Label(frame, text="    The result =" + str(final) + "   ")
    detail_label = Label(frame, text="Positive hits = " + str(pos) + "  Negative hits = " + str(neg))
    result_label['font'] = font.Font(size=16)
    detail_label['font'] = font.Font(size=10)
    result_label.grid(row=5, column=0, columnspan=2)
    detail_label.grid(row=6, column=0)


# Define initial button variables.
radio = StringVar()
radio.set("Sociology")
match = Label(root, text=phrase, padx=200).grid(row=7, column=0)

# Define the GUI button, with their names and related functions on which they act upon.
option_one = Radiobutton(root, text="Sociology", variable=radio, value="Sociology", command=lambda: radio_change(radio.get()))
option_two = Radiobutton(root, text="Psychology", variable=radio, value="Psychology", command=lambda: radio_change(radio.get()))
option_three = Radiobutton(root, text="History", variable=radio, value="Historia", command=lambda: radio_change(radio.get()))
entry_file = Entry(root, width=60, borderwidth=5)
open_btn = Button(root, text="Open a file", width=20, command=open_file)
process_btn = Button(root, text="Process", width=20, padx=50, pady=10, command=lambda: process_tt(ttFile))
entry = Entry(root, width=60, borderwidth=5)
score_plus = Radiobutton(root, text="+1", variable=radio, value=+1, command=lambda: score_change(radio.get())).grid(row=8, column=0, columnspan=2)
score_zero = Radiobutton(root, text="0", variable=radio, value=0, command=lambda: score_change(radio.get())).grid(row=9, column=0, columnspan=2)
score_minus = Radiobutton(root, text="-1", variable=radio, value=-1, command=lambda: score_change(radio.get())).grid(row=10, column=0, columnspan=2)
next_button = Button(root, text="Next", command=lambda: next_phrase(table_remaining, score, entry.get()))
frame = LabelFrame(root, padx=5, pady=5).grid(row=5, column=0, padx=10, pady=10)

# Button and entry field placement within the GUI.
entry_file.grid(row=0, column=0)
open_btn.grid(row=0, column=1)
option_one.grid(row=1, column=0, columnspan=2)
option_two.grid(row=2, column=0, columnspan=2)
option_three.grid(row=3, column=0, columnspan=2)
process_btn.grid(row=4, column=0, columnspan=2)
entry.grid(row=7, column=1)
option_one.grid(row=1, column=0, columnspan=2)
option_two.grid(row=2, column=0, columnspan=2)
option_three.grid(row=3, column=0, columnspan=2)
next_button.grid(row=11, column=0, columnspan=2)

root.mainloop()
