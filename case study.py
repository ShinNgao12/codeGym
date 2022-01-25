# IMPORT MODULE HERE
import pyttsx3
import datetime
import time
import xlsxwriter
from docx import Document
import speech_recognition as sr
import webbrowser as wb
import os

class virtual_assistant:
    def __init__(self):
        # tạo trợ lý
        self.assistant = pyttsx3.init()
        # giúp nhận diện trên 1 dòng
        self.voice = self.assistant.getProperty('voices')
         # chọn giọng nam/ nữ (0,1)
        self.assistant.setProperty('voice',self.voice[0].id)
        
    def speak(self,audio):
        self.assistant.say(audio)
        self.assistant.runAndWait()
        
    def welcome(self):
        hour = datetime.datetime.now().hour
        if 4 < hour < 10:
            self.speak('Good moring, Boss')
        elif hour < 18:
            self.speak('Good afternoon, Boss')
        else:
            self.speak('Good evening, Boss') 
        
    
    def showtime(self):
        time = datetime.datetime.now().strftime("%H:%M:%p")
        print("S.I.R.I: It ís", time)
        self.speak("It is" + time)
        
    def showdate(self):
        date = datetime.datetime.now().strftime('%A %d %B %Y')
        print("S.I.R.I: It ís", date)
        self.speak("It is" + date)
        
    def wait(self):
            cmd = sr.Recognizer()
            # xác nhận nguồn nhận --> microphone
            with sr.Microphone() as source:
                # setup thời gian đợi 
                cmd.pause_threshold = 2
                audio = cmd.listen(source)
                try:
                    query = cmd.recognize_google(audio,language = 'en').lower()
                    if 'hello' in query or 'wake' in query:
                        return False
                    return True
                except:
                    return True
                
    def sleep(self):
        check = self.wait()
        while check:
            time.sleep(1)
            check = self.wait()
    
    def command(self):
        # tạo đối tượng giúp nhận diện giọng nói
        cmd = sr.Recognizer()
        
        # xác nhận nguồn nhận --> microphone
        with sr.Microphone() as source:
            # setup thời gian đợi 
            cmd.pause_threshold = 2
            audio = cmd.listen(source)
            
            try:
                query = cmd.recognize_google(audio,language = 'en')
                print('-- YOU: ' + query)
            
            except sr.UnknownValueError:
                print('Please repeat or typing your command: ')
                query = input('Your command is: ')
        return query
    
    def create_Excel(self):
        self.speak('Please input name of excel file')
        name_file = input('Name of excel file: ')
        workbook = xlsxwriter.Workbook(f'{name_file}.xlsx')
        worksheet = workbook.add_worksheet()
        workbook.close()
        
        os.startfile(f'{name_file}.xlsx')
        
    def create_Word(self):
        self.speak('Please input name of word file')
        name_file = input('Name of word file: ')
        document = Document()
        document.save(f'{name_file}.docx')
        
        os.startfile(f'{name_file}.docx')
        
def main():
    siri = virtual_assistant()
    siri.welcome()
    while True:
        print("S.I.R.I: How can I help you ?")
        siri.speak('How can I help you ?')
        
        query = siri.command().lower()
        
        if 'google' in query:
            print('S.I.R.I: What do you want to search ?')
            siri.speak('What do you want to search ?')
            search = siri.command().lower()
            url = f'https://www.google.com.vn/search?q={search}'
            wb.get().open(url)
            siri.speak(f'Here you are')
            
        elif 'youtube' in query:
            print('S.I.R.I: What do you want to watch ?')
            siri.speak('What do you want to watch ?')
            search = siri.command().lower()
            url = f'https://www.youtube.com/results?search_query={search}'
            wb.get().open(url)
            siri.speak(f'Here you are')
        
        elif 'time' in query:
            siri.showtime()
            
        elif 'date' in query or 'day' in query:
            siri.showdate()
            
        elif 'notepad' in query:
            siri.speak(f'Openning Notepad')
            os.system('Notepad')
            
        elif 'excel' in query:
            siri.create_Excel()
            siri.speak(f'Openning Microsolf Excel')
            
        elif 'word' in query:
            siri.create_Word()
            siri.speak(f'Openning Microsolf Word')
        
        elif 'sleep' in query or 'pause' in query:
            print('S.I.R.I: Call me when you need')
            siri.speak('Call me when you need')
            siri.sleep()
            
        elif 'exit' in query or 'quit' in query:
            print('Goodbye, boss. Later')
            siri.speak('Goodbye, Boss.')
            siri.speak('Later')
            return
        
        else:
            print("S.I.R.I: Sorry I don't understand ?")
            siri.speak("Sorry I don't understand ?")
        
main()