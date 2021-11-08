from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC, wait
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException, NoSuchElementException
import time
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
import tkinter as tk
import tkinter.font as tkFont
from tkinter import *
from tkinter import messagebox, LabelFrame, Frame


ExcelFileName = 'AutomationTest.xlsx'
wb = load_workbook(ExcelFileName)
sheetindex = 0
sheets = wb.sheetnames
print("SheetNames: " + str(sheets))
ws = wb.active
i = 4
questionType = ws['C'+str(i)]
QuestionName = ws['B'+str(i)]
InstructionsToBeEntered = ws['D'+str(i)]
HelpTextToBeEntered = ws['H'+str(i)]
DragOptionToBeEntered = ws['E'+str(i)]
DropTargetToBeEntered = 'Drop Target Entered'
CorrectAnswerCell = ws['F'+str(i)]
CorrectAnswerCellList = ws['F'+str(i)]
ObjectivesCell = ws['J'+str(i)]
ObjectivesCell = str(ObjectivesCell.value).split('\n')
Objective = str(ObjectivesCell[0])
SubObjective = str(ObjectivesCell[1])
TotalSheets = len(sheets)
correctanswerAmount = len(str(CorrectAnswerCellList.value).split('\n'))

amountofAnswers = len(str(DragOptionToBeEntered.value).split('\n'))
answerIndex = 0
allAnswersClicked = 0
correctAnswerIndex = 0
#print("Amount of Total Answers"+str(amountofAnswers))


def SheetChecker():
    global i, questionType, QuestionName, InstructionsToBeEntered, HelpTextToBeEntered, DropTargetToBeEntered, CorrectAnswerCell, Objective, SubObjective, correctanswerAmount, amountofAnswers
    print(str(i)+" before")
    i += 1
    if ws.max_row != (i-1):
        print(str(ws.max_row) + ' Max Row Number')
        global CorrectAnswerCellList, ObjectivesCell
        questionType = ws['C'+str(i)]
        QuestionName = ws['B'+str(i)]
        InstructionsToBeEntered = ws['D'+str(i)]
        HelpTextToBeEntered = ws['H'+str(i)]
        DragOptionToBeEntered = ws['E'+str(i)]
        DropTargetToBeEntered = 'Drop Target Entered'
        CorrectAnswerCell = ws['F'+str(i)]
        CorrectAnswerCellList = ws['F'+str(i)]
        ObjectivesCell = ws['J'+str(i)]
        ObjectivesCell = str(ObjectivesCell.value).split('\n')
        Objective = str(ObjectivesCell[0])
        SubObjective = str(ObjectivesCell[1])
        correctanswerAmount = len(str(CorrectAnswerCellList.value).split('\n'))
        amountofAnswers = len(str(DragOptionToBeEntered.value).split('\n'))
        print(str(i)+" after")
        QuestionTypeClicker()
        QuestionCreation()
    else:
        print("Finished cycling!!")
    # print(str(DragOptionToBeEntered.value).split('\n'))
    #print(str(i) + " :This is the i value")


def TimeoutErrorMessage():
    global QIDCell
    QIDCell.fill = PatternFill(fgColor='34B1EB', fill_type='solid')
    QIDCell.value = "Question Failed To Create"
    wb.save(ExcelFileName)


def ClickMoreAnswer():
    global driver
    if amountofAnswers > answerIndex:
        # Create Another Answer
        CreateAnotherAnswer = driver.find_element_by_id(
            "AddAnswer0"+str(answerIndex))
        CreateAnotherAnswer.click()


def FinallyCreateQuestionButton():
    global driver, questionType
    if questionType.value == 'IFrame':
        FinalCreate = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable(
                (By.XPATH, '//*[@id="createQuestion"]/div[3]/div/input'))
        ).click()
    else:
        FinalCreate = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.ID, 'OnClickCreateQuestion'))
        ).click()
    time.sleep(2)


def RetreiveQID():
    global driver, i, QIDCell, QIDName
    QIDName = driver.find_element_by_class_name("modal-body").text
    QIDCell = ws['N'+str(i)]
    if questionType.value == "Multiple Choice":
        QIDCell.fill = PatternFill(fgColor='29FF49', fill_type='solid')
    # elif ws['C'+str(i)].value.lower() == "mcwi":
     #   QIDCell.fill = PatternFill(fgColor='34B1EB', fill_type='solid')
    else:
        QIDCell.fill = PatternFill(fgColor='FF4229', fill_type='solid')

    QIDCell.value = str(QIDName).strip(
        "Question successfully created. Question ID: ")
    wb.save(ExcelFileName)
    print(QIDName)


def QuestionTypeClicker():
    global driver, questionType
    openQuestionDD = WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.XPATH, ('//*[@id="questionTypeContainer"]/div[2]/span[1]/span/span[1]'))))
    openQuestionDD.click()
    if questionType.value.lower() == 'multiple choice' or questionType.value.lower() == 'choose' or questionType.value.lower() == 'multiple' or questionType.value.lower() == 'mcwi':
        questionType.value = 'Multiple Choice'
    elif questionType.value.lower() == "dnd" or questionType.value.lower() == "drag and match":
        questionType.value = 'Drag and Match'
    elif questionType.value.lower() == 'applications':
        questionType.value = 'Application'
    elif questionType.value.lower() == 'demo' or questionType.value.lower() == 'iframe' or questionType.value == 'Demo':
        questionType.value = 'IFrame'
    time.sleep(1)
    questionTypeSelected = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="QuestionType-list"]//li[text() = "'+questionType.value+'"]'
                                    ))
    )
    print("Question Type: " + questionType.value)
    questionTypeSelected.click()


def QuestionCreation():
    global questionType, QuestionName, answerIndex, amountofAnswers, DragOptionToBeEntered, DropTargetToBeEntered, InstructionsToBeEntered, HelpTextToBeEntered, i, driver, allAnswersClicked, Objective, SubObjective, correctanswerAmount
    global correctAnswerIndex, CorrectAnswerCell, QIDCell, ErrorMessage
    # Inputting Question Name
    time.sleep(1)
    questionNameEntry = WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.ID, 'QuestionName')))
    #questionNameEntry = driver.find_element_by_id("QuestionName")
    questionNameEntry.send_keys(QuestionName.value)
    # Opens Objective Menu
    try:
        objectiveDD = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="ObjectiveId"]'
                                            ))
        )
        objectiveDD.click()
    except:
        print("Object failed to Match any listed under this product")
        ErrorMessage = "Object failed to Match any listed under this product"
        TimeoutErrorMessage()
        ErrorWindowDefault()
        SheetChecker()
        ResetWindow()
        LoginAndOpenQuestionInput()

    time.sleep(1)
    selectionOBJ = Select(driver.find_element_by_xpath(
        "//select[@name='ObjectiveId']"))
    selectionOBJ.select_by_visible_text(Objective.strip(" "))
    # Opens subObjective Drop Down
    subObjectiveDD = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located(
            (By.XPATH, '//*[@id="ddlSubObjective"]'))
    )
    subObjectiveDD.click()
    time.sleep(1)
    selectionSubOBJ = Select(driver.find_element_by_xpath(
        "//select[@name='ddlSubObjective']"))
    try:
        selectionSubOBJ.select_by_visible_text(SubObjective.strip(" "))
    except TimeoutException or ElementClickInterceptedException or NoSuchElementException:
        ErrorMessage = "SubObjective does not match any listed under this product, Click OK to cycle to next Question"
        ErrorWindowDefault()
        TimeoutErrorMessage()
        SheetChecker()
        print("SubObjective does not match any listed under this product")

    time.sleep(1)
    # Instructions
    InstructionFrame = WebDriverWait(driver, 20).until(
        EC.frame_to_be_available_and_switch_to_it((By.XPATH, '//*[@id="cke_1_contents"]/iframe')))
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable(
        (By.XPATH, "//body[@class='cke_editable cke_editable_themed cke_contents_ltr cke_show_borders']/p"))).send_keys(InstructionsToBeEntered.value)
    driver.switch_to.default_content()
    # HelpText
    HelpFrame = WebDriverWait(driver, 20).until(
        EC.frame_to_be_available_and_switch_to_it((By.XPATH, '//*[@id="cke_2_contents"]/iframe')))
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable(
        (By.XPATH, "//body[@class='cke_editable cke_editable_themed cke_contents_ltr cke_show_borders']/p"))).send_keys(HelpTextToBeEntered.value)
    driver.switch_to.default_content()

    if questionType.value == 'Multiple Choice':
        DragOptionToBeEntered = str(DragOptionToBeEntered.value).split('\n')
        CorrectAnswerCell = str(CorrectAnswerCell.value).split('\n')
        while amountofAnswers > answerIndex:
            MultipleChoice1 = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((
                    By.XPATH, "//input[@name='InstructionComponent[0].AnswersBlock.Answers["+str(answerIndex)+"].AnswerText']"))).send_keys(DragOptionToBeEntered[answerIndex])
            if correctanswerAmount != allAnswersClicked:
                if str(CorrectAnswerCell[correctAnswerIndex]) == str(DragOptionToBeEntered[answerIndex]):
                    CorrectAnswerClicked = WebDriverWait(driver, 20).until(
                        EC.element_to_be_clickable(
                            (By.XPATH, '//*[@id="InstructionComponent_0__AnswersBlock_Answers_'+str(answerIndex)+'__IsCorrect"]'))
                    ).click()
                    correctAnswerIndex += 1
                    allAnswersClicked += 1
                    time.sleep(1)
            if (int(amountofAnswers)-1) != answerIndex:
                ClickMoreAnswer()
                time.sleep(1)
                #print(str(answerIndex) + "current answer index")
            answerIndex += 1

    FinallyCreateQuestionButton()
    RetreiveQID()
    CloseQIDMenu = WebDriverWait(driver, 15).until(
        EC.presence_of_element_located(
            (By.XPATH, '//*[@id="successfulUpdateModal"]/div/div/div[1]/button/span'))
    )
    CloseQIDMenu.click()
    SheetChecker()


def LoginAndOpenQuestionInput():
    global Username, Password, driver, categoryName, productName, ErrorMessage, ErrorWindow
    if not driver:

        Username = Username_var.get()
        Password = Password_var.get()
        categoryName = " " + Category_var.get()
        productName = Product_var.get()
        PATH = "C:\Program Files (x86)\chromedriver.exe"
        driver = webdriver.Chrome(PATH)
        driver.get("https://author.gmetrix.net")
        # This is opening the webpage^^^
        driver.maximize_window()
        search = driver.find_element_by_id("Email")
        # Your Credentials here
        search.send_keys(Username)
        search = driver.find_element_by_id("Password")
        search.send_keys(Password)
        time.sleep(4)

        signIn = driver.find_element_by_xpath('//*[@id="buttonSignIn"]')
        signIn.click()
        # This is Signing the user in ^^

        questionTab = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.LINK_TEXT, "Questions"))
        )
        questionTab.click()
        createButton = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.LINK_TEXT, "Create"))
        )
        createButton.click()
        # Opening Category Menu
        openCategoryDD = driver.find_element_by_css_selector(
            'span[class="k-input"]')
        openCategoryDD.click()

        # Clicking Category Selected
        categorySelected = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="Category-list"]//li[text() = "'+categoryName+'"]'
                                        ))
        )
        categorySelected.click()
        # Opening Product Drop Down List
        openProductDD = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.XPATH, ('/html/body/div[2]/form/div[1]/div/div/div[2]/div[2]/span[1]/span/span[1]'))))
        openProductDD.click()
        time.sleep(2)
        # Clicking Product Name
        productSelected = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="Product-list"]//li[text() = "'+productName+'"]'
                                        ))
        )
        productSelected.click()

        time.sleep(2)
        # Opening Question Drop Down List
        # Inputing the Question Type
        QuestionTypeClicker()
        firstCreateButton = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable(
                (By.XPATH, '/html/body/div[2]/form/div[2]/input'))
        )
        firstCreateButton.click()
        QuestionCreation()


def DisplayStartWindow():
    global driver, check, PasswordInput, Password_var, UsernameInput, Username_var, Category_var, Product_var, ExcelName_var, root
    root = Tk()
    ErrorMessage = ""
    Username_var = StringVar()
    Password_var = StringVar()
    Category_var = StringVar()
    Product_var = StringVar()
    ExcelName_var = StringVar()
    root.title("Question Creation")
    driver = None
    root.geometry('1000x800')
    root.configure(background="#1e2124")
    LabelFontStyle = tkFont.Font(family="Lucida Grande", size=12)

    #Color Library
    mainColor = '#3e4e69'
    secondColor = '#435e8a'
    textColor ="#FFF"

    # Creating menuFrame
    menuFrame = LabelFrame(root, bg=mainColor)
    menuFrame.pack(pady=50)

    entryFrame = Frame(menuFrame, bg=secondColor)
    entryFrame.pack(padx=50, pady=50)

    productFrame= Frame(menuFrame, bg=secondColor)
    productFrame.pack(padx=10, pady=10)

    ProjectLabel = Label(menuFrame, text='Question Automation Tool', font=("Verdana bold", 20),
                         bg=mainColor, fg=textColor)
    ProjectLabel.place(x=80, y=0)

    UsernameLabel = Label(entryFrame, text='Username :',
                          bg=secondColor, fg=textColor, font=LabelFontStyle)
    UsernameLabel.grid(row=1, column=1, padx=0)

    UsernameInput = Entry(entryFrame, textvariable=Username_var,
                          bg='#cfcfcf', width=25).grid(row=1, column=2, padx=0)

    PasswordLabel = Label(entryFrame, text='Password :', font=LabelFontStyle,
                          bg=secondColor, fg=textColor).grid(row=2, column=1, padx=20)

    PasswordInput = Entry(entryFrame, textvariable=Password_var,
                          bg='#cfcfcf', width=25, show='*')
    PasswordInput.grid(row=2, column=2, padx=20)

    check = Checkbutton(entryFrame, text='Show Password',
                        command=show, bg=secondColor,fg=textColor,font=LabelFontStyle)
    check.grid(row=2, column=3, padx=10)

    CategoryLabel = Label(productFrame, text='Enter Category:',
                           bg=secondColor, fg=textColor, font=LabelFontStyle).grid(row=1, column=1, padx=0)
    CategoryInput = Entry(productFrame, textvariable=Category_var,
                          bg='#cfcfcf', width=30).grid(row=1, column=2, padx=0)

    ProductLabel = Label(productFrame, text='Enter Product:',
                          bg=secondColor, fg=textColor,font=LabelFontStyle).grid(row=2, column=1, padx=0)
    ProductInput = Entry(productFrame, textvariable=Product_var,
                          bg='#cfcfcf', width=30).grid(row=2, column=2, padx=0)

    Start = Button(root, text="Start", command=(
        LoginAndOpenQuestionInput), fg='#FFF', bg='#63cbff', width=25).place(x=400, y=400)

    # MessageDisplayed = Label(root, text="Status: " + ErrorMessage,
    #                        bg='#1e2124',  fg='#FFF').place(x=10, y=250)
    root.mainloop()


def show():
    PasswordInput.configure(show='')
    check.configure(command=hide, text='Hide Password')


def hide():
    PasswordInput.configure(show='*')
    check.configure(command=show, text='Show Password')


def ErrorWindowDefault():
    messagebox.showerror(title="Error", message=ErrorMessage)


def ResetWindow():
    root.destroy()
    DisplayStartWindow()


DisplayStartWindow()

# .split('\n')
"""        categorySelected = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="Category-list"]//li[text() = "'+categoryName+'"]'
                                        ))"""


# ~~~~~~~~~~~Made By Diana Cervantes ~~~~~~~~~~~~~ Project: Question Creation Automation


# Just in Case Graveyard Code

# elif questionType.value == 'Drag and Match':

#     while answerIndex < amountofAnswers:
#         if answerIndex == 0:
#             DragOption = WebDriverWait(driver, 15).until(
#                 EC.frame_to_be_available_and_switch_to_it(
#                     (By.XPATH, '//*[@id="cke_3_contents"]/iframe'))
#             )
#             WebDriverWait(driver, 15).until(EC.element_to_be_clickable((
#                 By.XPATH, '/html/body/p'))).send_keys(DragOptionToBeEntered)
#             driver.switch_to.default_content()

#             DropTarget = WebDriverWait(driver, 15).until(
#                 EC.frame_to_be_available_and_switch_to_it(
#                     (By.XPATH, '//*[@id="cke_4_contents"]/iframe'))
#             )
#             WebDriverWait(driver, 15).until(EC.element_to_be_clickable((
#                 By.XPATH, '/html/body/p'))).send_keys(DropTargetToBeEntered)
#             driver.switch_to.default_content()
#         elif answerIndex == 2:
#             DragOption = WebDriverWait(driver, 15).until(
#                 EC.frame_to_be_available_and_switch_to_it(
#                     (By.XPATH, '//*[@id="cke_298_contents"]/iframe'))
#             )
#             WebDriverWait(driver, 15).until(EC.element_to_be_clickable((
#                 By.XPATH, '/html/body/p'))).send_keys(DragOptionToBeEntered)
#             driver.switch_to.default_content()

#             DropTarget = WebDriverWait(driver, 15).until(
#                 EC.frame_to_be_available_and_switch_to_it(
#                     (By.XPATH, '//*[@id="cke_299_contents"]/iframe'))
#             )
#             WebDriverWait(driver, 15).until(EC.element_to_be_clickable((
#                 By.XPATH, '/html/body/p'))).send_keys(DropTargetToBeEntered)
#             driver.switch_to.default_content()
#         elif answerIndex == 1:
#             DragOption = WebDriverWait(driver, 15).until(
#                 EC.frame_to_be_available_and_switch_to_it(
#                     (By.XPATH, '//*[@id="cke_214_contents"]/iframe'))
#             )
#             WebDriverWait(driver, 15).until(EC.element_to_be_clickable((
#                 By.XPATH, '/html/body/p'))).send_keys(DragOptionToBeEntered)
#             driver.switch_to.default_content()

#             DropTarget = WebDriverWait(driver, 15).until(
#                 EC.frame_to_be_available_and_switch_to_it(
#                     (By.XPATH, '//*[@id="cke_215_contents"]/iframe'))
#             )
#             WebDriverWait(driver, 15).until(EC.element_to_be_clickable((
#                 By.XPATH, '/html/body/p'))).send_keys(DropTargetToBeEntered)
#             driver.switch_to.default_content()
##answerIndex += 1
# ClickMoreAnswer()
