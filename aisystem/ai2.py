from docx import Document
import pandas as pd
import ollama
import os
import tkinter as tk
from tkinter import filedialog

selected_file_path = 'vvvvvv'
ms = ''

def browse_file(root):
    global selected_file_path
    # Open a file dialog and allow the user to select a file
    file_path = filedialog.askopenfilename(
        title="Select a file",
        filetypes=[("Text Files", "*.txt"), ("Word Documents", "*.docx"), ("Excel Files", "*.xlsx"),
                   ("All Files", "*.*")]
    )

    if file_path:
        # Store the selected file path in the global variable
        selected_file_path = file_path
        print(f"Selected file: {file_path}")
    else:
        print("No file selected.")

    root.destroy()

def mainTk():
    # Create the main window
    root = tk.Tk()
    root.title("File Browser")

    # Create a button to trigger the file dialog
    browse_button = tk.Button(root, text="Browse", command=lambda: browse_file(root))
    browse_button.pack(pady=20)

    # Start the Tkinter event loop
    root.mainloop()

# normal question
def ask_question(md):
    desirable_model = md

    question_to_ask = input("Question: ")
    response = ollama.chat(model=desirable_model, messages=[
        {
            'role': 'user',
            'content': question_to_ask,
        },
    ], stream=True)

    print("Model's thinking process:")
    for chunk in response:
        print(chunk['message']['content'], end='', flush=True)
    print("\n")

# text files
def file_question(md):
    desirable_model = md

    mainTk()
    print(selected_file_path)
    file_path = str(selected_file_path)
    print(file_path, type(file_path))
    if not os.path.isabs(file_path):
        print("Please provide an absolute path.")
        return

    print("File selected:", file_path)
    try:
        with open(file_path, "r") as f:
            file_content = f.read()

        question = input("Question on the file: ")
        query = f"From {file_content}, {question}"

        response = ollama.chat(model=desirable_model, messages=[
            {
                'role': 'user',
                'content': query,
            },
        ], stream=True)

        print("Model's thinking process:")
        for chunk in response:
            print(chunk['message']['content'], end='', flush=True)
        print("\n")
    except FileNotFoundError:
        print("File not found. Please check the path and try again.")

# list or arrays
def list_question(md):
    desirable_model = md
    lt = []
    print("Add list values (enter '0' to finish):")
    while True:
        m = input()
        if m != '0':
            lt.append(m)
        else:
            break

    question = input("Question on the list: ")
    query = f"From {lt}, {question}"

    response = ollama.chat(model=desirable_model, messages=[
        {
            'role': 'user',
            'content': query,
        },
    ], stream=True)

    print("Model's thinking process:")
    for chunk in response:
        print(chunk['message']['content'], end='', flush=True)
    print("\n")

# word files
def doc_question(md):
    desirable_model = md
    mainTk()
    print(selected_file_path)
    file_path = str(selected_file_path)
    print(file_path, type(file_path))
    if not os.path.isabs(file_path):
        print("Please provide an absolute path.")
        return

    try:
        doc = Document(file_path)
        full_text = [para.text for para in doc.paragraphs]
        doc_content = "\n".join(full_text)

        question = input("Question on the document: ")
        query = f"From the following document content: {doc_content}, {question}"

        response = ollama.chat(model=desirable_model, messages=[
            {
                'role': 'user',
                'content': query,
            },
        ], stream=True)

        print("Model's thinking process:")
        for chunk in response:
            print(chunk['message']['content'], end='', flush=True)
        print("\n")
    except Exception as e:
        print(f"Error reading the document: {e}")

# excel files
def excel_question(md):
    desirable_model = md
    mainTk()
    print(selected_file_path)
    file_path = str(selected_file_path)
    print(file_path, type(file_path))
    if not os.path.isabs(file_path):
        print("Please provide an absolute path.")
        return

    try:
        df = pd.read_excel(file_path)
        excel_content = df.to_string()

        question = input("Question on the Excel file: ")
        query = f"From the following Excel content: {excel_content}, {question}"

        response = ollama.chat(model=desirable_model, messages=[
            {
                'role': 'user',
                'content': query,
            },
        ], stream=True)

        print("Model's thinking process:")
        for chunk in response:
            print(chunk['message']['content'], end='', flush=True)
        print("\n")
    except Exception as e:
        print(f"Error reading the Excel file: {e}")

# choosing option
def main():
    while True:
        global ms

        print("Choose model")
        print("1--'deepseek-r1:1.5b'--Less efficient but fast")
        print("2--'deepseek-r1:8b'--Slow but efficient")
        mdl = input()
        if mdl == '1':
            ms = 'deepseek-r1:1.5b'
        if mdl == '2':
            ms = 'deepseek-r1:8b'
        print("Choose an option:")
        choice = input("F - File, Q - Question, L - List, D - Doc, E - Excel: ").strip().lower()

        if choice == 'f':
            file_question(ms)
        elif choice == 'q':
            ask_question(ms)
        elif choice == 'l':
            list_question(ms)
        elif choice == 'd':
            doc_question(ms)
        elif choice == 'e':
            excel_question(ms)
        else:
            print("Invalid option. Please try again.")
            continue

        continue_choice = input("Continue? [1 - Yes, 2 - No]: ").strip()
        if continue_choice == '2':
            break

main()