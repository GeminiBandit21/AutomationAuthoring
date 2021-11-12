# Notes

To create new styles type: python -m ttkcreator

Working Question Types:
Multiple Choice Works without issue- Marks Green
IFrame Works missing template input - Marks Red will mark Yellow since Template issue
Application- Creates QID does not add Sample Doc or Templates!!
Drag and Match or DND: Creates Question does not add answers!! Marks Red
Categorize-Marks Red missing categories
Drop Down List-Code Red Will need to input the drop downs
Short Answer - Marks red

Fails:
Hotspot- Fills Name, Obj, SubObj,HelpText, Instructions. doesn't input template or hotspot locations will fail!
Drag to Paragraph-Does not Working
Fill in the Blank-Fails could work would need to add the fill in the blank info
Paragraph Dropdown - Needs the Dropdowns added
Simulation - Working just needs tweaking on the final create Question button would mark Red
Multiple Yes/No- fails could work with tweaking requires to answers
Single True False - could work with tweaking requires answers to work

--To-Do--

-Added File Insert for Excel Sheets 11/10/21
-Added Theme selection 11/10/21
-Change Category and Product Entry boxes to comboboxes 11/11/21
Save users Perfered Theme
Need to add Sheet iterator or Toggle Function
-Error Handling 11/12/21
Update question types for more finished question types from beginning to end.


-- Just in Case Graveyard Code

elif questionType.value == 'Drag and Match':

     while answerIndex < amountofAnswers:
         if answerIndex == 0:
            DragOption = WebDriverWait(driver, 15).until(
                 EC.frame_to_be_available_and_switch_to_it(
                     (By.XPATH, '//*[@id="cke_3_contents"]/iframe'))
             )
             WebDriverWait(driver, 15).until(EC.element_to_be_clickable((
                 By.XPATH, '/html/body/p'))).send_keys(DragOptionToBeEntered)
             driver.switch_to.default_content()

             DropTarget = WebDriverWait(driver, 15).until(
                 EC.frame_to_be_available_and_switch_to_it(
                     (By.XPATH, '//*[@id="cke_4_contents"]/iframe'))
             )
             WebDriverWait(driver, 15).until(EC.element_to_be_clickable((
                 By.XPATH, '/html/body/p'))).send_keys(DropTargetToBeEntered)
             driver.switch_to.default_content()
         elif answerIndex == 2:
             DragOption = WebDriverWait(driver, 15).until(
                 EC.frame_to_be_available_and_switch_to_it(
                     (By.XPATH, '//*[@id="cke_298_contents"]/iframe'))
             )
             WebDriverWait(driver, 15).until(EC.element_to_be_clickable((
                 By.XPATH, '/html/body/p'))).send_keys(DragOptionToBeEntered)
             driver.switch_to.default_content()
             DropTarget = WebDriverWait(driver, 15).until(
                 EC.frame_to_be_available_and_switch_to_it(
                     (By.XPATH, '//*[@id="cke_299_contents"]/iframe'))
             )
             WebDriverWait(driver, 15).until(EC.element_to_be_clickable((
                 By.XPATH, '/html/body/p'))).send_keys(DropTargetToBeEntered)
             driver.switch_to.default_content()
         elif answerIndex == 1:
             DragOption = WebDriverWait(driver, 15).until(
                 EC.frame_to_be_available_and_switch_to_it(
                     (By.XPATH, '//*[@id="cke_214_contents"]/iframe'))
             )
             WebDriverWait(driver, 15).until(EC.element_to_be_clickable((
                 By.XPATH, '/html/body/p'))).send_keys(DragOptionToBeEntered)
             driver.switch_to.default_content()

             DropTarget = WebDriverWait(driver, 15).until(
                 EC.frame_to_be_available_and_switch_to_it(
                     (By.XPATH, '//*[@id="cke_215_contents"]/iframe'))
             )
             WebDriverWait(driver, 15).until(EC.element_to_be_clickable((
                 By.XPATH, '/html/body/p'))).send_keys(DropTargetToBeEntered)
             driver.switch_to.default_content()

answerIndex += 1
ClickMoreAnswer()

--Theme Colors-

Samhain:
 -#EB6121
 -#C8522B
 -#A64435
 -#83353E
 -#612748
 -#3E1852

SamhainVariant:
 -#020203
 -#5A2969
 -#BD624F
 -#FFA95E
 -#FFE59E

CodeRed:
 -#3F1515
 -#9E2830
 -#CE313D
 -#FD3A4A
 -#6F1E22
 -#100C08

EarthlyTones:
 -#9D9379
 -#888574
 -#3B3739
 -#312129
 -#5C4332
 -#705C44