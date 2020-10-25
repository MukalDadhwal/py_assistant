from tkinter import *
from tkinter import ttk
from ttkthemes import ThemedTk
from tkinter import messagebox
from PIL import ImageTk, Image
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import requests, json, webbrowser
import win32com.client, datetime
import smtplib, email, pygame, random
from newsapi import NewsApiClient
from math import *

window  = ThemedTk(theme='breeze')

window.geometry('600x550+300+70')

window.title('Py Assistant')

window.resizable(0,0)

speaker = win32com.client.Dispatch('SAPI.SpVoice')

speaker.Speak("Hey my name is py tell me what can I do for you")

speak = IntVar()
speak.set(True)

# Defining all the functions

def speakResult():
    if speak.get() == 1:
        return True
    return False


def game():

    pygame.init()

    width = 600
    height = 600

    screen = pygame.display.set_mode((height,width))

    pygame.display.set_caption('Snake')

    red = (255,0,0)
    white = (255,255,255)
    green = (0,155,0)

    clock = pygame.time.Clock()

    font = pygame.font.SysFont(None , 25 )

    def snake(block_size, snakeList):
        for XnY in snakeList:
            pygame.draw.rect(screen, green, (XnY[0], XnY[1], block_size, block_size))


    def message(msg,color):
        screen_text = font.render(msg, True, color)
        screen.blit(screen_text, (200,200))

    def gameloop():
        x_change = width/2
        y_change = height/2

        block_size = 10
        AppleSize = 10

        lead_x_change = 0
        lead_y_change = 0

        snakeList = []
        snakeLength = 1

        GameExit = False
        GameOver = False

        randAppleX = random.randrange(0, width - 30, 10)
        randAppleY = random.randrange(0, height - 30, 10)

        while not GameExit:

            while GameOver == True:
                screen.fill(white)
                message('GAME OVER , Press P to play again or Q to quit',red)
                pygame.display.update()

                for event in pygame.event.get():
                    if event.type == pygame.KEYDOWN:
                        if event.key == pygame.K_q:
                            GameExit = True
                            GameOver = False
                        if event.key == pygame.K_p:
                            gameloop()

            for event in pygame.event.get():
                if event.type == pygame.QUIT:
                    GameExit = True

                if event.type == pygame.KEYDOWN:
                    if event.key == pygame.K_LEFT:
                        lead_x_change = -10
                        lead_y_change = 0
                    elif event.key == pygame.K_RIGHT:
                        lead_x_change = 10
                        lead_y_change = 0
                    elif event.key == pygame.K_UP:
                        lead_y_change = -10
                        lead_x_change = 0
                    elif event.key == pygame.K_DOWN:
                        lead_y_change = 10
                        lead_x_change = 0


            if x_change >= 600 or x_change < 0 or y_change >= 600 or y_change < 0:
                GameOver = True

            x_change += lead_x_change
            y_change += lead_y_change

            screen.fill(white)


            snakehead = []
            snakehead.append(x_change)
            snakehead.append(y_change)
            snakeList.append(snakehead)

            if len(snakeList) > snakeLength:
                del snakeList[0]

            for eachElement in snakeList[:-1]:
                if eachElement == snakehead:
                    GameOver = True

            snake(block_size, snakeList)


            pygame.draw.rect(screen, red, [randAppleX, randAppleY, AppleSize, AppleSize])
            clock.tick(30)
            pygame.display.update()

            if x_change == randAppleX and y_change == randAppleY:
                randAppleX = random.randrange(0, width - 30, 10)
                randAppleY = random.randrange(0, height - 30, 10)
                snakeLength += 1


        pygame.display.update()
        pygame.quit()
        # quit()
    try:
        gameloop()
    except:
        print('')

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

        s = False
        city = query.lower().replace(" ", "")[13:]
        result = showWeather(city)

        emptyLable.config(text=result)

        if speakResult():
            speaker.Speak(result)

    elif  query.lower().replace(" ", "") == 'game' or query.lower().replace(" ", "") == 'snakegame':
        if speakResult():
            speaker.Speak('Opening the snake game')
        game()

    elif  'latestnewson' in query.lower().replace(" ", ""):
        topic = query.lower().replace(" ", "")[12:]
        title = query.lower()[12:]

        if speakResult():
            speaker.Speak(f'Showing the latest news on {topic}')

        NewsWindow = Toplevel(window)

        NewsWindow.title('News App')

        NewsWindow.geometry('600x650')

        NewsWindow.resizable(0,100)

        TopFrame = Frame(NewsWindow)

        def callback(url):
            webbrowser.open_new_tab(url)

        def showNews():

            try:
                api = NewsApiClient(api_key="26c55621d6204b9dbaaebf85bd238831")

                data = api.get_everything(q=topic, page_size=5)

                articles = data["articles"]
                articlesList = list(articles)

                lb1 = Label(TopFrame, text=f"\n1.{articlesList[0]['title']}", justify=LEFT, wraplength=595, font='bold')
                lb1.grid(row=1, column=0)

                lb2 = Label(TopFrame, text=f"Description => {articlesList[0]['description']}", justify=LEFT, wraplength=595)

                lb2.grid(row=2, columnspan=2)

                link1 = Label(TopFrame, text=f"Click Url => {articlesList[0]['url']}", justify=LEFT, wraplength=595)

                link1.grid(row=3, columnspan=2)
                link1.bind('<Button-1>', lambda e: callback(articlesList[0]['url']))

                lb3 = Label(TopFrame, text=f"\n2.{articlesList[1]['title']}", justify=LEFT, wraplength=595, font='bold')
                lb3.grid(row=4, column=0)

                lb4 = Label(TopFrame, text=f"Description => {articlesList[1]['description']}", justify=LEFT, wraplength=595)

                lb4.grid(row=5, columnspan=2)

                link2 = Label(TopFrame, text=f"Click Url => {articlesList[1]['url']}", justify=LEFT, wraplength=595)

                link2.grid(row=6, columnspan=2)
                link2.bind('<Button-1>', lambda e: callback(articlesList[1]['url']))

                lb5 = Label(TopFrame, text=f"\n3.{articlesList[2]['title']}", justify=LEFT, wraplength=595, font='bold')
                lb5.grid(row=7, column=0)

                lb6 = Label(TopFrame, text=f"Description => {articlesList[2]['description']}", justify=LEFT, wraplength=595)

                lb6.grid(row=8, columnspan=2)

                link3 = Label(TopFrame, text=f"Click Url => {articlesList[2]['url']}", justify=LEFT, wraplength=595)

                link3.grid(row=9, columnspan=2)
                link3.bind('<Button-1>', lambda e: callback(articlesList[2]['url']))

                lb7 = Label(TopFrame, text=f"\n4.{articlesList[3]['title']}", justify=LEFT, wraplength=595, font='bold')
                lb7.grid(row=10, column=0)

                lb8 = Label(TopFrame, text=f"Description => {articlesList[3]['description']}", justify=LEFT, wraplength=595)

                lb8.grid(row=11, columnspan=2)

                link4 = Label(TopFrame, text=f"Click Url => {articlesList[3]['url']}", justify=LEFT, wraplength=595)

                link4.grid(row=12, columnspan=2)
                link4.bind('<Button-1>', lambda e: callback(articlesList[3]['url']))

                lb9 = Label(TopFrame, text=f"\n5.{articlesList[4]['title']}", justify=LEFT, wraplength=595, font='bold')
                lb9.grid(row=13, column=0)

                lb10 = Label(TopFrame, text=f"Description => {articlesList[4]['description']}", justify=LEFT, wraplength=595)

                lb10.grid(row=14, columnspan=2)

                link5 = Label(TopFrame, text=f"Click Url => {articlesList[4]['url']}", justify=LEFT, wraplength=595)

                link5.grid(row=15, columnspan=2)
                link5.bind('<Button-1>', lambda e: callback(articlesList[4]['url']))

            except IndexError:
                speaker.Speak('Not enough information to show!')

            except:
                speaker.Speak('No internet connection!')


        lb = Label(TopFrame, text = f'Latest News on {title}', font='Courier 12 bold', fg='blue', bg='yellow')

        lb.grid(row=0, columnspan=2)

        TopFrame.grid()

        showNews()

        NewsWindow.mainloop()


    elif  query.lower().replace(" ","") == 'calculator' or query.lower().replace(" ","") == 'opencalculator':

        if speakResult():
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
        if speakResult():
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
        if speakResult():
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
    list.insert(11, 'Write: <your query>')

    list.insert(12, '')

    list.insert(13, 'To Play games')
    list.insert(14, 'Write: game or snake game')

    list.insert(15, '')

    list.insert(16, 'To see the latest news')
    list.insert(17, 'Write: latest news on <your topic>')

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

help.add_checkbutton(label = 'Don\'t speak result', variable = speak, onvalue = 0, offvalue = 1)

menuBar.add_cascade(label='Help', menu=help)

menuBar.add_cascade(label='About', command=about)

menuBar.add_cascade(label='Exit', command=exit)

window.config(menu=menuBar)

currFileLocation = __file__

icon_path = currFileLocation.replace("app_v1.0.py", "app_icon.ico")

window.iconbitmap(icon_path)

topFrame = ttk.Frame(window)

image_path = currFileLocation.replace("app_v1.0.py", "assistant_logo.png")

img = ImageTk.PhotoImage(Image.open(image_path))

image = ttk.Label(topFrame, image = img)

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