# Activate a virtual environment for program to work:
# cd SpellGame
# conda activate venv
# python Spell.py

import tkinter as tk
from tkinter import font
import win32com.client as wincl
import random
from Fuzzy_Logic_Class import FuzzyLogic
from fuzzywuzzy import process


class SpellSmart(FuzzyLogic):

    def __init__(self, master):
        """ This makes the GUI front end and graphics """
        
        # Window
        self.master = master
        self.height = 600
        self.width = 900
        self.title = self.master.title("Spell Smart")
        self.canvas = tk.Canvas(height=self.height, width=self.width)
        self.canvas.pack()

        # Frame
        self.master_frame = tk.Frame(self.master, bg="#9CE0FF")
        self.master_frame.place(anchor="center", relx=0.5, rely=0.5, relheight=1, relwidth=1)

        # Fake entry widget, so nothing happens when it is used
        self.phony = tk.Entry(self.master_frame, bd=3, relief="groove", font=("news gothic", 60))
        self.phony.place(anchor="sw", relx=0.175, rely=0.975, relwidth=0.8, relheight=1/3)

        # text_rounds Box - Rounds
        self.text_rounds = tk.Entry(self.master_frame, bd=3, relief="groove", font=("news gothic", 60))
        self.text_rounds.bind("<Return>", self.validation) # The enter key or <return> makes an event occur when pressed. this is what bind does
        self.text_rounds.focus_set()

        # Option_misspelt box - misspelt words
        self.misspelt_words = tk.Entry(self.master_frame, bd=3, relief="groove", font=("news gothic", 60))
        self.misspelt_words.bind("<Return>", self.misspelt)
        self.misspelt_words.focus_set()

        # Option correct words box - to correct misspelt words
        self.correct_words = tk.Entry(self.master_frame, bd=3, relief="groove", font=("news gothic", 60))
        self.correct_words.bind("<Return>", self.correct)
        self.correct_words.focus_set()

        # Spelling words entry text box widget
        self.spelling_words = tk.Entry(self.master_frame, bd=3, relief="groove", font=("news gothic", 60))
        self.spelling_words.bind("<Return>", self.spelling_enter)
        self.spelling_words.focus_set()

        # Super words entry box
        self.super_question = tk.Entry(self.master_frame, bd=3, relief="groove", font=("news gothic", 60))
        self.super_question.bind("<Return>", self.spelling_super)
        self.super_question.focus_set()


        # Label
        self.instructions = tk.Message(self.master_frame, fg="white", bg="#9CE0FF", text="Instructions: \nUse the text box below"
        " to answer the spoken questions and spell the words. Hit the 'start' button to begin. Hit the"
        " speaker button if you want a word repeated.", font=("news gothic", 20), justify="left", aspect=300, padx=0, pady=0)
        self.instructions.place(anchor="nw", relx=0.025, rely=0.02)

        # Sound Button
        self.sound_image = tk.PhotoImage(file="SoundButtonFinal2.png")
        self.speaker = tk.Button(self.master_frame, relief="flat", bg="#9CE0FF", image=self.sound_image, command=self.sound)
        self.speaker.place(anchor="n", relheight=1/2, relwidth=1/3, relx=0.808333, rely=0.025)

        # Start Button
        self.start_button = tk.Button(self.master_frame, fg="white", bg="#00a8e9", relief="groove", text="Start", font=("news gothic", 30), command=self.start)
        self.start_button.place(anchor="sw", rely=0.975, relx=0.025, relheight=1/3, relwidth=0.15)
        
        # Voice
        self.speak = wincl.Dispatch("SAPI.SpVoice")

        # Open file and put all words into a list
        self.open_file = open("SpellingWord.txt", "r")
        self.list_words = self.open_file.read().split()
        self.open_file.close()

        self.repeat = 0
        self.input_rounds = lambda : self.text_rounds.get() # The lambda makes a function, so when self.input is called it goes to lambda function. put () around self.input to retrieve remember
        self.input_misspelt = lambda : self.misspelt_words.get()
        self.input_correct = lambda : self.correct_words.get()
        self.super_input = lambda : self.super_question.get()
        self.user_spelt = ""
        self.input_word = [] 
        self.check_misspelt = False
        self.check_correct = False
        self.full_rotation = 0
        self.random_rotation = 0
        self.random_words = 0

        self.new_list = [] # List of words user is to spell
        self.new_dict = {}
        self.first = True
        self.input_spell = ""
        self.iterate = 0
        self.correct_spelt = 0
        self.total_spelt = 0
        self.first_dict = True
        self.FuzzyLogic = FuzzyLogic    
        self.speaker["state"] = "disable" 
        self.super_word = ""

        tk.mainloop()  # Must be at end

    def sound(self):
        """ This will repeat a specific word """
        self.speak.Speak(self.new_dict[self.iterate])
 
    def start(self):
        """ This will initiate the game when start button is pressed """

        self.correct = 0
        self.total = 0
        self.super_correct = 0

        self.text_rounds.place(anchor="sw", relx=0.175, rely=0.975, relwidth=0.8, relheight=1/3)
        self.speak.Speak("How many rounds do you want to play?")

    def validation(self, *event):
        """ Validates input so it's an integer equal to or less than 20 """

        try:
            self.input_round = int(self.input_rounds()) # self.input_round CANNOT be called self.input_rounds otherwise it will not work because of the same name
            if self.input_round > 20 or self.input_round < 1: 
                raise Exception
            elif self.input_round <= 20:
                pass
            self.rounds()
            
        except Exception:
            self.speak.Speak("Please enter an integer less than or equal to 20.")

    def rounds(self, *event): # event is the "<Return>" event pressed or the enter button
        """ This calculates how many rounds the user wants to play and leads onto the next question """

        self.input_round = int(self.input_rounds())
        self.text_rounds.destroy()
        self.misspelt_words.place(anchor="sw", relx=0.175, rely=0.975, relwidth=0.8, relheight=1/3)
        if self.input_round < 5:
            self.random_words = self.input_round
        if self.input_round >= 5:
            self.full_rotation = self.input_round // 5
            self.random_rotation = self.input_round % 5
        self.speak.Speak("Do you want me to spell out your attempts?")

    def misspelt(self, *event):
        """ This checks if the user wants the computer to read out their misspelt words """

        self.input_word.append(self.input_misspelt())
        yes = process.extract("yes", self.input_word)
        no = process.extract("no", self.input_word)

        if yes[0][1] > no[0][1]: # If response is most similar to yes
            self.check_misspelt = True

        elif no[0][1] > yes[0][1]: # If response is most similar to no
            self.check_misspelt = False

        self.input_word.remove(self.input_misspelt())
        self.misspelt_words.destroy()
        self.correct_words.place(anchor="sw", relx=0.175, rely=0.975, relwidth=0.8, relheight=1/3)
        self.speak.Speak("If you spell a word wrong, do you want me to dictate the correct spelling?")

    def correct(self, *event):
        """ This checks if the user wants to correct their misspelt words """
    
        self.input_word.append(self.input_correct())

        yes = process.extract("yes", self.input_word)
        no = process.extract("no", self.input_word)

        if yes[0][1] > no[0][1]: # If response is most similar to yes
            self.check_correct = True

        elif no[0][1] > yes[0][1]: # If response is most similar to no
            self.check_correct = False
        
        self.input_word.remove(self.input_correct())
        self.correct_words.destroy()
        self.spelling_words.place(anchor="sw", relx=0.175, rely=0.975, relwidth=0.8, relheight=1/3)
        
        if self.input_round >= 5:
            self.spelling_progressive(self.full_rotation, self.random_rotation)

        elif self.input_round < 5:
            self.spelling_random(self.random_words)

    def spelling_progressive(self, full_rotation=None, random_rotation=None):
        """ Makes a list of words in increasing difficulty based on 
        number of words the user wants """
        
        for x in range(full_rotation):
            self.new_list.append(str(self.list_words[random.randint(1, 60) - 1])) # Random Really easy word
        for y in range(full_rotation):
            self.new_list.append(str(self.list_words[random.randint(61, 150) - 1])) # Random easy word
        for z in range(full_rotation):
            self.new_list.append(str(self.list_words[random.randint(151, 290) - 1])) # Random normal word
        for w in range(full_rotation):
            self.new_list.append(str(self.list_words[random.randint(291, 390) - 1])) # Random hard word
        for r in range(full_rotation):
            self.new_list.append(str(self.list_words[random.randint(391, 450) - 1])) # Random really hard word. Impossible is 451 - 459

        for x in range(self.random_rotation):
            self.new_list.append(str(self.list_words[random.randint(391, 459) - 1]))
        
        self.spelling()
        
    def spelling_random(self, random_words):

        for x in range(self.random_words):
            self.new_list.append(self.list_words[random.randint(61, 391) - 1])
        
        self.spelling()

    def spelling(self, *event):
        """ This is the actual spelling game """
        self.speaker["state"] = "normal"

        if self.first_dict is True:
            for x in range(len(self.new_list)): # Key for dictionary starts at 0
                self.new_dict[x] = self.new_list[x]
            print(self.new_dict)
            self.first_dict = False

        if self.first_dict is False:
            pass
        
        if self.new_dict[self.iterate] in self.list_words[451 : ]: # If its an impossible word 
            self.super_question.place(anchor="sw", relx=0.175, rely=0.975, relwidth=0.8, relheight=1/3) # Just place this over the spelling_words
            # entry widget. There is no need to delete the entry words widget. Also, it didn't work because it was placed in the wrong spot.
            self.speak.Speak("You have encountered a super word. Would you like to spell it?")

        elif self.new_dict[self.iterate] in self.list_words[ : 451]:
            if self.first is True:
                self.speak.Speak("Your first word is ")
                self.first = False
            elif self.first is False:
                self.speak.Speak("Your next word is ")
            self.speak.Speak(self.new_dict[self.iterate])
        
        
    def spelling_super(self, *event):
        """ This is what happens when the user gets a 'super' word """
        
        self.input_word.append(self.super_input())
        yes = process.extract("yes", self.input_word)
        no = process.extract("no", self.input_word)

        if yes[0][1] > no[0][1]: # If response is most similar to yes
            self.input_word.remove(self.super_input())
            self.super_question.destroy()
            self.speak.Speak("Your super word is ")
            self.input_word.clear()
            self.speak.Speak(self.new_dict[self.iterate])

        elif no[0][1] > yes[0][1]: # If response is most similar to no
            self.iterate += 1
            self.super_question.delete("0", "end")
            self.super_question.destroy()
            self.input_word.clear() # Clears entire list
            self.spelling()
        
        
        

    def spelling_enter(self, *event):
        """ Allows user to enter words """
        self.user_spelt = str(self.spelling_words.get()).lower()
        self.spelling_words.delete("0", "end") # do 0 instead of "1" otherwise the first character will not delete. use this for word spell

        if self.user_spelt == self.new_dict[self.iterate]:
            self.speak.Speak("Correct!")
            self.correct_spelt += 1
        
        if self.user_spelt != self.new_dict[self.iterate]:
            self.speak.Speak("Incorrect!")

            if self.check_misspelt is True:
                self.speak.Speak("You spelt ")
                for x in self.user_spelt:
                    self.speak.Speak(x)
            
            if self.check_correct is True:
                self.speak.Speak("The correct spelling is ")
                for i in self.new_dict[self.iterate]:
                    self.speak.Speak(i)
            
            self.speak.Speak("This is " + str(self.FuzzyLogic(self.user_spelt, self.new_dict[self.iterate]).percentage()) + " percent correct")
        
        self.iterate += 1 # This placement is crucial
        self.total_spelt += 1

        if self.iterate == self.input_round:
            self.finish()
        elif self.iterate != self.input_round:
            self.spelling()

    def finish(self):
        self.speak.Speak("Well done! you got " + str(self.correct_spelt) + " out of " + str(self.total_spelt) + " correct.")
        self.speaker["state"] = "disable"  


if __name__ == "__main__":
    root = tk.Tk()    
    SpellSmart(root)

