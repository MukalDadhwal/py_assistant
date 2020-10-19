from tkinter import *
from tkinter import ttk
from ttkthemes import ThemedTk
from tkinter import messagebox
from PIL import ImageTk, Image
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import requests, json, webbrowser
import win32com.client, datetime
import smtplib, email
from math import *

window  = ThemedTk(theme='breeze')

window.geometry('600x550+300+70')

window.title('Py Assistant')

window.resizable(0,0)

speaker = win32com.client.Dispatch('SAPI.SpVoice')

speaker.Speak("Hey my name is py tell me what can I do for you")

# Defining all the functions

def sendMail(senderEmail, recEmail, password, subject, msg):

    try:
        message = MIMEMultipart()
        message['From'] = senderEmail
        message['To'] = recEmail
        message['Subject'] = subject

        message.attach(MIMEText(msg, 'plain'))
        text = message.as_string()

        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.ehlo()
        server.starttls()
        server.ehlo()
        server.login(senderEmail, password)
        server.sendmail(senderEmail, recEmail, text)

    finally:
        server.close()

def kelvin_to_celsius(x):

    return round(x - 273, 2)

def showWeather(city):

    api_key = "59830b3bf04fb5a6d65e035f3888ab20"

    base_url = "http://api.openweathermap.org/data/2.5/weather?"

    complete_url = base_url + "appid=" + api_key + "&q=" + city

    try:
        response = requests.get(complete_url)
    except:
        return 'No internet connection'

    x = response.json()

    # code 200 -> site can be accessed
    # code 404 -> site is not available

    try:
        if x["cod"] != "400":

            if x["cod"] != "404":

                y = x["main"]

                current_temperature = y["temp"]
                current_pressure = y["pressure"]
                current_humidiy = y["humidity"]
                current_temp = kelvin_to_celsius(current_temperature)

                z = x["weather"]

                weather_description = z[0]["description"]

                return (f" THE WEATHER IN {city.upper()} IS \n"+

                        "\n Temperature (in celsius unit) => " + str(current_temp).upper()+

                        "\n Atmospheric pressure (in hPa unit) => " + str(current_pressure).upper()+

                        "\n Humidity (in percentage) => " + str(current_humidiy).upper()+

                        "\n Description => " + str(weather_description))

            else:
                return " City Not Found"
        else:
            return " Please Enter a City name"
    except:
        return "No internet connection"


def search(query):

    webbrowser.open_new_tab(f'https://google.com/search?q={query}')

# Main function -> responsible for showing all results
def showResult(*args):

    query = entryField.get().strip()

    emptyLable.config(text=" ")

    if 'showweatherof' in query.lower().replace(" ",""):

        city = query.lower().replace(" ", "")[13:]
        result = showWeather(city)

        emptyLable.config(text=result)
        speaker.Speak(result)

    elif  query.lower().replace(" ","") == 'calculator' or query.lower().replace(" ","") == 'opencalculator':

        speaker.Speak('Opening the calculator')

        calculatorApp = Toplevel(window)

        calculatorApp.title('CALCULATOR')

        calculatorApp.geometry('480x420')

        calculatorApp.resizable(0,0)

        def press(Bt_num):

            exp = Ent.get()

            if exp == 'Error':
                equation.set(str(Bt_num))
            else:
                equation.set(str(exp)+str(Bt_num))

        def clear():
            var = ''
            equation.set(var)

        def sine(angle: int) -> str:
            radian = radians(angle)
            return str(round(sin(radian), 5))

        def cosine(angle: int) -> str:
            r = radians(angle)
            return str(round(cos(r), 5))

        def tangent(angle: int) -> str:
            r = radians(angle)
            return str(round(tan(r), 5))

        def evalute():
            expression = Ent.get()

            try:
                exp = expression.replace(" ", "")

                if 'sin' in expression:
                    value = sine(int(exp[3:]))
                    equation.set(value)
                elif 'cos' in expression:
                    value = cosine(int(exp[3:]))
                    equation.set(value)
                elif 'tan' in expression:
                    value = tangent(int(exp[3:]))
                    equation.set(value)
                elif 'sqrt' in expression:
                    value = exp[4:]
                    equation.set(sqrt(int(value)))
                else:
                    val = round(eval(str(expression)), 5)
                    equation.set(val)

            except:      # Checking for 0 division or some other error
                equation.set('Error')
        def delete():

            exp2 = Ent.get()
            afterDel = exp2[:-1]

            equation.set(afterDel)

        TopFrame = Frame(calculatorApp)

        equation = StringVar()

        Ent = ttk.Entry(TopFrame, textvariable=equation, width=55)

        Ent.grid(row=0, column=0, pady=10, padx=45)

        TopFrame.grid()

        DownFrame = Frame(calculatorApp)

        ## All the buttons

        Bt_1 = Button(DownFrame, text=1, width=12, height=3, font='Courier 10 bold', command=lambda: press(1))

        Bt_1.grid(row=1, column=0)

        Bt_2 = Button(DownFrame, text=2, width=12, height=3, font='Courier 10 bold', command=lambda: press(2))

        Bt_2.grid(row=1, column=1)

        Bt_3 = Button(DownFrame, text=3, width=12, height=3, font='Courier 10 bold', command=lambda: press(3))

        Bt_3.grid(row=1, column=2)

        Bt_add = Button(DownFrame, text='+', width=12, height=3, font='Courier 10 bold', command=lambda: press('+'))

        Bt_add.grid(row=1, column=3)

        Bt_4 = Button(DownFrame, text=4, width=12, height=3, font='Courier 10 bold', command=lambda: press(4))

        Bt_4.grid(row=2, column=0)

        Bt_5 = Button(DownFrame, text=5, width=12, height=3, font='Courier 10 bold', command=lambda: press(5))

        Bt_5.grid(row=2, column=1)

        Bt_6 = Button(DownFrame, text=6, width=12, height=3, font='Courier 10 bold', command=lambda: press(6))

        Bt_6.grid(row=2, column=2)

        Bt_subtract = Button(DownFrame, text='-', width=12, height=3, font='Courier 10 bold', command=lambda: press('-'))

        Bt_subtract.grid(row=2, column=3)

        Bt_7 = Button(DownFrame, text=7, width=12, height=3, font='Courier 10 bold', command=lambda: press(7))

        Bt_7.grid(row=3, column=0)

        Bt_8 = Button(DownFrame, text=8, width=12, height=3, font='Courier 10 bold', command=lambda: press(8))

        Bt_8.grid(row=3, column=1)

        Bt_9 = Button(DownFrame, text=9, width=12, height=3, font='Courier 10 bold', command=lambda: press(9))

        Bt_9.grid(row=3, column=2)

        Bt_multiply = Button(DownFrame, text='X', width= 12, height=3, font='Courier 10 bold', command=lambda: press('*'))

        Bt_multiply.grid(row=3, column=3)

        Bt_del = Button(DownFrame, text='DELETE', width=12, height=3, font='Courier 10 bold', command=delete)

        Bt_del.grid(row=4, column=0)

        Bt_0 = Button(DownFrame, text=0, width=12, height=3, font='Courier 10 bold', command=lambda: press(0))

        Bt_0.grid(row=4, column=1)

        Bt_equal = Button(DownFrame, text='=', width=12, height=3, font='Courier 10 bold', command=evalute)

        Bt_equal.grid(row=4, column=2)

        Bt_divide = Button(DownFrame, text='/', width=12, height=3, font='Courier 10 bold', command=lambda: press('/'))

        Bt_divide.grid(row=4, column=3)

        Bt_sin = Button(DownFrame, text='sin', width=12, height=3, font='Courier 10 bold', command=lambda: press('sin'))

        Bt_sin.grid(row=5, column=0)

        Bt_cos = Button(DownFrame, text='cos', width=12, height=3, font='Courier 10 bold', command=lambda: press('cos'))

        Bt_cos.grid(row=5, column=1)

        Bt_tan = Button(DownFrame, text='tan', width=12, height=3, font='Courier 10 bold', command=lambda: press('tan'))

        Bt_tan.grid(row=5, column=2)

        Bt_sqrt = Button(DownFrame, text='sqrt', width=12, height=3, font='Courier 10 bold', command=lambda: press('sqrt'))

        Bt_sqrt.grid(row=5, column=3)

        Bt_clear = Button(DownFrame, text='CLEAR', width=52, height=3, font='Courier 10 bold', command=clear)

        Bt_clear.grid(row=6, columnspan=4)

        DownFrame.grid()

    elif query.lower().replace(" ", "") == 'sendmail' or query.lower().replace(" ", "") == 'mail' or query.lower().replace(" ", "") == 'mailsend':

        speaker.Speak('Opening the Mail Sender app')

        mailSenderApp = Toplevel(window)

        mailSenderApp.resizable(0,0)

        mailSenderApp.geometry('600x670')

        mailSenderApp.title('Mail Sender App')

        def sm():

            try:
                sendMail(str(Ent2.get()), str(Ent4.get()), str(Ent3.get()), str(Ent5.get()), str(textBox.get('1.0', END)))

                blankText.configure(text='Mail sended')
                speaker.Speak('Mail sended')

            except:

                blankText.configure(text='Something went wrong !')
                speaker.Speak('Something went wrong!')


        loginText = Label(mailSenderApp, text='LOGIN DETAILS', bg='Red', fg='Yellow')

        loginText.grid(row=0, column=0)

        senderId = Label(mailSenderApp, text='Enter your email id here', height=5)

        senderId.grid(row=1, column=0)

        Ent2 = Entry(mailSenderApp, width=40)

        Ent2.grid(row=1, column=1)

        Ent2.focus()

        Password = Label(mailSenderApp, text='Password', pady=20)

        Password.grid(row=2, column=0)

        Ent3 = Entry(mailSenderApp, width=40, show='*')

        Ent3.grid(row=2, column=1)

        composeText = Label(mailSenderApp, text='COMPOSE EMAIL', fg='Yellow', bg='Red')

        composeText.grid(row=3, columnspan=1)

        receiverId = Label(mailSenderApp, text='Type the receiver\'s email id', height=5)

        receiverId.grid(row=4, column=0)

        Ent4 = Entry(mailSenderApp, width=40)

        Ent4.grid(row=4, column=1)

        Subject = Label(mailSenderApp, text='Type the subject of your email id here', height=2)

        Subject.grid(row=5, column=0)

        Ent5 = Entry(mailSenderApp, width=40)

        Ent5.grid(row=5, column=1)

        contentText = Label(mailSenderApp, text='Write the content here :', height=20)

        contentText.grid(row=6, column=0)

        textBox = Text(mailSenderApp, width=40, height=10, wrap=WORD, pady=20)

        textBox.grid(row=6, column=1)

        Bt_send = Button(mailSenderApp, text='Send Mail',command=sm, bg='Yellow', fg='Blue')

        Bt_send.grid(row=7, columnspan=2)

        blankText = Label(mailSenderApp, text='')

        blankText.grid(row=7, column=1)

        mailSenderApp.mainloop()

    elif not entryField.get():

        emptyLable.config(text='PLEASE ENTER SOMETHING IN THE ENTRY FIELD')

        speaker.Speak("PLEASE ENTER SOMETHING IN THE ENTRY FIELD")

    else:

        toSearch = entryField.get()

        speaker.Speak(f"Showing web results for your query {toSearch}")

        search(toSearch)


def showGuide():

    searchWindow = Toplevel(window)

    searchWindow.geometry('300x300+920+70')

    searchWindow.title('Help')

    searchWindow.resizable(0,0)

    guideText = Label(searchWindow, text='GUIDE FOR APP', font='Times 20')

    guideText.grid(row=0, columnspan=2)

    list = Listbox(searchWindow, width=42, height=15)

    list.insert(1, 'To find the weather of a place')
    list.insert(2, 'Write: show weather of <city name>')

    list.insert(3, '')

    list.insert(4, 'To open calculator')
    list.insert(5, 'Write: calculator or open calculator')

    list.insert(6, '')

    list.insert(7, 'To open Mail Sender App')
    list.insert(8, 'Write: send mail or mail')

    list.insert(9, '')

    list.insert(10, 'To search Web')
    list.insert(11, 'Write: <Query>')

    list.insert(12, '')

    list.grid(row=1)

    searchWindow.mainloop()

def about():

    aboutWindow = Toplevel(window)

    aboutWindow.geometry('340x200+920+70')

    aboutWindow.title('About')

    aboutWindow.resizable(0,0)

    text = """Hey, I am a 16 year old app developer.
    Reach out to me at:
    github id => github.com/MukalDadhwal
    facebook id => blabla2facebook.com
    gmail id => test@gmail.com"""

    someText= Label(aboutWindow, text=text, font='10')

    someText.grid(row=0)

def exit():

    window.destroy()

menuBar = Menu(window)

help = Menu(menuBar, tearoff=0)

help.add_command(label='How To Search', command=showGuide)

menuBar.add_cascade(label='Help', menu=help)

menuBar.add_cascade(label='About', command=about)

menuBar.add_cascade(label='Exit', command=exit)

window.config(menu=menuBar)

icon_path = "drivename:\\xxxxxx\\xxxxxxx\\app_icon.ico"

window.iconbitmap(icon_path)

topFrame = ttk.Frame(window)

image_path = "drivename:\\xxxxxxx\\xxxxxx\\assistant_logo.png"

img = ImageTk.PhotoImage(Image.open(image_path))

image = ttk.Label(topFrame, image=img)

image.grid(row=0,column=0,pady=10)

topFrame.grid()

downFrame = ttk.Frame(window)

text = ttk.Label(downFrame, text='HEY MY NAME IS PY TELL ME WHAT CAN I DO FOR YOU', font='bold 9' )

text.grid(row=0, column=0, padx=150)

entryField = ttk.Entry(downFrame, width=70)

entryField.bind("<Return>", showResult)

entryField.grid(row=2, columnspan=2, padx=20, pady=15)

entryField.focus()

button = ttk.Button(downFrame, text='Go', width=10,command=showResult)

button.grid(row=3, columnspan=2, pady=10)

emptyLable = ttk.Label(downFrame, text=' ', font=("Times", 10, "bold"), justify=LEFT)

emptyLable.grid(row=4, columnspan=2, pady=50)

downFrame.grid()

window.mainloop()