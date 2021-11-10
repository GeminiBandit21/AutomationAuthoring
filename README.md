
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

# To-Do
#22023d
#460180
#a8cbed
Added File Insert for Excel Sheets 11/10/21
Need to add Sheet iterator
Need to finish Theme selection
Update question types

# Just in Case Graveyard Code

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

