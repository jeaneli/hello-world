# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.


import tkinter as tk
from tkinter import messagebox, simpledialog
import random
import pandas as pd

def add_word_to_excel(word, meaning, file_name='words_meanings.xlsx'):
    try:
        # Read the existing Excel file
        df = pd.read_excel(file_name)
    except FileNotFoundError:
        # If the file does not exist, create a new DataFrame
        df = pd.DataFrame(columns=['Word', 'Meaning'])

    # Append the new word and meaning as a new row
    new_row = {'Word': word, 'Meaning': meaning}
    df = df.append(new_row, ignore_index=True)

    # Save the updated DataFrame back to the Excel file
    df.to_excel(file_name, index=False)
    print(f"Added word: {word} with meaning: {meaning} to {file_name}")


# Function to read Excel data as a dictionary
def read_excel_as_dict(file_name='words_meanings.xlsx'):
    # Read the Excel file into a DataFrame
    df = pd.read_excel(file_name)

    # Convert the DataFrame to a dictionary
    words_dict = df.set_index('Word')['Meaning'].to_dict()

    return words_dict

class WordApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Word Memory App")
        self.words = read_excel_as_dict()

        self.word_label = tk.Label(root, text="Word:")
        self.word_label.pack()
        self.word_entry = tk.Entry(root)
        self.word_entry.pack()

        self.meaning_label = tk.Label(root, text="Meaning:")
        self.meaning_label.pack()
        self.meaning_entry = tk.Entry(root)
        self.meaning_entry.pack()

        self.add_button = tk.Button(root, text="Add Word", command=self.add_word)
        self.add_button.pack()

        self.test_button = tk.Button(root, text="Test Me", command=self.test_me)
        self.test_button.pack()

    def add_word(self):
        word = self.word_entry.get()
        meaning = self.meaning_entry.get()
        if word and meaning:
            self.words[word] = meaning
            messagebox.showinfo("Success", f"Added word: {word}")
        else:
            messagebox.showwarning("Input Error", "Please enter both word and meaning")
        add_word_to_excel(word, meaning)

    def test_me(self):
        if self.words:
            word = random.choice(list(self.words.keys()))
            meaning = self.words[word]
            user_meaning = simpledialog.askstring("Test", f"What is the meaning of {word}?")
            if user_meaning == meaning:
                messagebox.showinfo("Correct", "You got it right!")
            else:
                messagebox.showinfo("Incorrect", f"The correct meaning is {meaning}")
        else:
            messagebox.showwarning("No Words", "No words to test")



# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    root = tk.Tk()
    app = WordApp(root)
    root.mainloop()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/


