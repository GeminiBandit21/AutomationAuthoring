from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC, wait
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException
import time
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
import tkinter as tk
from tkinter import *


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
    global driver, i
    QIDName = driver.find_element_by_class_name("modal-body").text
    QIDCell = ws['N'+str(i)]
    if questionType.value == "Multiple Choice":
        QIDCell.fill = PatternFill(fgColor='29FF49', fill_type='solid')
    #elif ws['C'+str(i)].value.lower() == "mcwi":
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
    global questionType, QuestionName, answerIndex, amountofAnswers, DragOptionToBeEntered, DropTargetToBeEntered, InstructionsToBeEntered, HelpTextToBeEntered, i, driver, allAnswersClicked, Objective, SubObjective, correctanswerAmount, correctAnswerIndex, CorrectAnswerCell
    # Inputting Question Name
    time.sleep(1)
    questionNameEntry = WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.ID, 'QuestionName')))
    #questionNameEntry = driver.find_element_by_id("QuestionName")
    questionNameEntry.send_keys(QuestionName.value)
    # Opens Objective Menu
    objectiveDD = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="ObjectiveId"]'
                                        ))
    )
    objectiveDD.click()
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
    selectionSubOBJ.select_by_visible_text(SubObjective.strip(" "))
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


def LoginAndOpenQuestionInput():
    global Username, Password, driver, categoryName, productName

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
        categorySelected = WebDriverWait(driver, 20).until(
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


root = Tk()

Username_var = StringVar()
Password_var = StringVar()
Category_var = StringVar()
Product_var = StringVar()
ExcelName_var = StringVar()
root.title("Question Creation")
driver = None
root.geometry('600x300')
root.configure(background="#1e2124")


def show():
    PasswordInput.configure(show='')
    check.configure(command=hide, text='hide password')


def hide():
    PasswordInput.configure(show='*')
    check.configure(command=show, text='show password')


ProjectLabel = Label(root, text='Question Automation Tool',
                     background="#1e2124", foreground='#FFF').place(x=5, y=0)
UsernameLabel = Label(root, text='Username:',
                      background="#1e2124", foreground='#FFF').place(x=55, y=30)
UsernameInput = Entry(root, textvariable=Username_var,
                      background='#cfcfcf', width=25).place(x=125, y=30)

PasswordLabel = Label(root, text='Password:',
                      background='#1e2124', foreground='#FFF').place(x=56, y=50)

PasswordInput = Entry(root, textvariable=Password_var,
                      background='#cfcfcf', width=25, show='*')
PasswordInput.place(x=125, y=50)

check = Checkbutton(root, text='show password',
                    command=show)
check.place(x=290, y=50)

CategoryLabel = Label(root, text='Enter Category:',
                      background='#1e2124', foreground='#FFF').place(x=20, y=85)
CategoryInput = Entry(root, textvariable=Category_var,
                      background='#cfcfcf', width=30).place(x=110, y=85)

ProductLabel = Label(root, text='Enter Product:',
                     background='#1e2124', foreground='#FFF').place(x=310, y=85)
ProductInput = Entry(root, textvariable=Product_var,
                     background='#cfcfcf', width=30).place(x=400, y=85)

Start = Button(root, text="Start", command=(
    LoginAndOpenQuestionInput), fg='#FFF', bg='#63cbff', width=25).place(x=200, y=200)

root.mainloop()


# .split('\n')
"""        categorySelected = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="Category-list"]//li[text() = "'+categoryName+'"]'
                                        ))"""

# ~~~~~~~~~~~Made By Diana Cervantes ~~~~~~~~~~~~~ Project: Question Creation Automation


# Notes
# Working Question Types:
# Multiple Choice Works without issue- Marks Green
# IFrame Works missing template input - Marks Red will mark Yellow since Template issue
# Application- Creates QID does not add Sample Doc or Templates!!
# Drag and Match or DND: Creates Question does not add answers!! Marks Red
# Categorize-Marks Red missing categories
# Drop Down List-Code Red Will need to input the drop downs
# Short Answer - Marks red

# Fails
# Hotspot- Fills Name, Obj, SubObj,HelpText, Instructions. doesn't input template or hotspot locations will fail!
# Drag to Paragraph-Does not Working
# Fill in the Blank-Fails could work would need to add the fill in the blank info
# Paragraph Dropdown - Needs the Dropdowns added
# Simulation - Working just needs tweaking on the final create Question button would mark Red
# Multiple Yes/No- fails could work with tweaking requires to answers
# Single True False  - could work with tweaking requires answers to work


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
