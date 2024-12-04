# ------------- Selenium Python Test Frameowrk AOL(Phase-1)------------#


""" Automating the test cases for AOL Site"""
import re
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from excel_data import Data
from excel_data import XML
from selenium.webdriver.support.select import Select

browser = Service("C:\\msedgedriver.exe")

# Initialising driver to access web elements of AOL
driver = webdriver.Edge(service=browser)
driver.get(XML.url)
driver.maximize_window()
if "aol" in driver.current_url:
    driver.find_element(By.CSS_SELECTOR, "input[name=name]").send_keys(XML.user_name)
    driver.find_element(By.CSS_SELECTOR, "input[name=pass]").send_keys(XML.password)
    driver.find_element(By.CSS_SELECTOR, "input[value=Login]").click()
    driver.maximize_window()


elif "ldr" in driver.current_url:
    driver.find_element(By.CSS_SELECTOR, "input[name=username]").send_keys(XML.user_name)
    driver.find_element(By.CSS_SELECTOR, "input[name=password]").send_keys(XML.password)
    driver.find_element(By.CSS_SELECTOR, "input[value=Login]").click()
    driver.maximize_window()

#global hi_elm, low_elm, amd_chk

class TestFramework:
    """ Main class function declaration where looping each instrument through instrument index"""

    def test_aol(self):
        """" Declaring Framework test function"""

        for instrument_index, _ in enumerate(Data.instru_list):
            global run
            driver.refresh()
            driver.find_element(By.LINK_TEXT, "Browse").click()  # Navigating to Browse Section
            sel = Select(driver.find_element(By.XPATH, "//form[2]//select[@name='browsetype']"))
            sel.select_by_visible_text(Data.sheet.cell(row=instrument_index + 2, column=1).value)
            close_tab = []
            # ******************************  Search Section ************************#
            # No of characters present in 'Browse Versioned Legislation'
            # Defining search elements
            if "aol" in driver.current_url:
                aol_search = driver.find_elements(By.XPATH, "//body/div[2]/form[2]"
                                                            "/div/table/tbody/tr[2]/td[2]/input")
                search_char = len(aol_search)

            else:
                ldr_search = driver.find_elements(By.XPATH,
                                                  "//body[1]/div[2]/form[2]/div[1]/table[1]"
                                                  "/tbody[1]/tr[2]/td[2]/input")
                search_char = len(ldr_search)
            print("No of Alphabets present in 'Browse by Legislation section'=" + str(search_char))
            if "aol" in driver.current_url:
                alphabets_list = driver.find_element(By.XPATH, "//body/div[2]/form[2]/div/table/"
                                                               "tbody/tr[2]/td[2]/input")
            elif "ldr" in driver.current_url:
                alphabets_list = driver.find_element(By.XPATH,
                                                     "//body[1]/div[2]/form[2]/div[1]/"
                                                     "table[1]/tbody[1]/tr[2]/td[2]/input")

            my_char = ord('A')  # convert char to ascii
            if Data.instru_list[instrument_index] is not None:
                # Iterating from 'A' till the first character of instrument
                while my_char <= ord(Data.char_list[instrument_index]):
                    if my_char == ord(Data.char_list[instrument_index]):
                        print(Data.char_list[instrument_index])
                        # Opening the link based on first character
                        alphabets_list.find_element(By.XPATH,
                                                    "//input[@value='" + Data.char_list[
                                                        instrument_index] + "']").click()
                        def dupinstru():
                            if len(duplicate_instru) > 1:
                                print("Found more than one result for"
                                  " the instrument-" + Data.instru_list[
                                      instrument_index]
                                  + ",Total Results = " + str(len(duplicate_instru)))
                                Data.sheet.cell(row=instrument_index + 2,
                                            column=3).value = "Yes"
                                Data.sheet.cell(row=instrument_index + 2,
                                            column=4).value = len(duplicate_instru)
                            else:
                                Data.sheet.cell(row=instrument_index + 2,
                                            column=3).value = "No"
                                Data.sheet.cell(row=instrument_index + 2,
                                            column=4).value = int(0)
                        try:
                            print(Data.instru[instrument_index])
                            element = driver.find_element(By.LINK_TEXT, "show all")
                            element.is_displayed()
                            element.click()
                            if "aol" in driver.current_url:
                                hi_elm = driver.find_elements(By.XPATH,
                                                              "(//tr[@class='highlite'])/td[1]/a")
                                low_elm = driver.find_elements(By.XPATH,
                                                               "(//tr[@class='lowlite'])/td[1]/a")
                            elif "ldr" in driver.current_url:
                                hi_elm = driver.find_elements(By.XPATH,
                                                            "(//tr[@class='odd'])/td[1]/ul/li/a")
                                low_elm = driver.find_elements(By.XPATH,
                                                            "(//tr[@class='even'])/td[1]/ul/li/a")
                            run = 0
                            for h_elm in hi_elm:
                                conv_web = re.sub("[^A-Za-z0-9]", "", h_elm.text).lower()
                                if Data.instru_list[instrument_index] == conv_web:
                                    duplicate_instru = driver.find_elements(By.LINK_TEXT, h_elm.text)
                                    dupinstru()
                                    driver.find_element(By.LINK_TEXT, h_elm.text).click()
                                    run = 1
                                    break
                            if run == 0:
                                for l_elm in low_elm:
                                    conv_web = re.sub("[^A-Za-z0-9]", "", l_elm.text).lower()
                                    if Data.instru_list[instrument_index] == conv_web:
                                        duplicate_instru = driver.find_elements(By.LINK_TEXT, l_elm.text)
                                        dupinstru()
                                        driver.find_element(By.LINK_TEXT, l_elm.text).click()
                                        run = 1
                                        break
                        except NoSuchElementException:
                            if "aol" in driver.current_url:
                                hi_elm = driver.find_elements(By.XPATH,
                                                              "(//tr[@class='highlite'])/td[1]/a")
                                low_elm = driver.find_elements(By.XPATH,
                                                               "(//tr[@class='lowlite'])/td[1]/a")
                            elif "ldr" in driver.current_url:
                                hi_elm = driver.find_elements(By.XPATH,
                                                            "(//tr[@class='odd'])/td[1]/ul/li/a")
                                low_elm = driver.find_elements(By.XPATH,
                                                            "(//tr[@class='even'])/td[1]/ul/li/a")
                            run = 0
                            for h_elm in hi_elm:
                                conv_web = re.sub("[^A-Za-z0-9]", "", h_elm.text).lower()
                                if Data.instru_list[instrument_index] == conv_web:
                                    duplicate_instru = driver.find_elements(By.LINK_TEXT, h_elm.text)
                                    dupinstru()
                                    driver.find_element(By.LINK_TEXT, h_elm.text).click()
                                    run = 1
                                    break
                            if run == 0:
                                for l_elm in low_elm:
                                    conv_web = re.sub("[^A-Za-z0-9]", "", l_elm.text).lower()
                                    if Data.instru_list[instrument_index] == conv_web:
                                        duplicate_instru = driver.find_elements(By.LINK_TEXT, l_elm.text)
                                        dupinstru()
                                        driver.find_element(By.LINK_TEXT, l_elm.text).click()
                                        run = 1
                                        break
                            try:
                                assert run == 0
                                print("No instrument were found for the title - " +
                                      Data.instru[instrument_index])


                            except AssertionError:
                                pass


                    my_char += 1  # iterate over alphabets_list
            else:
                continue

            parent = driver.window_handles[0]
            # __________Checking whether instrument is available with nodes/not______

            try:
                driver.find_element(By.ID, "expand_button")
            except NoSuchElementException:
                if run != 0:
                    print("No nodes available:" + Data.instru_list[instrument_index])
                    Data.sheet.cell(row=instrument_index + 2, column=5).value = int(0)
                    Data.sheet.cell(row=instrument_index + 2, column=6).value = "No nodes are present"
                    Data.Workbook.save("Files\\UserFile.xlsx")
                continue
            # ********** Nodes Loading Section *********#
            if "aol" in driver.current_url:
                aol_instru = "//body[1]/div[2]/div[1]/div[3]/div[4]" \
                             "/div[1]/div[2]/div[1]/table/tbody/tr/td"
                inst_node = driver.find_elements(By.XPATH, aol_instru)
            elif "ldr" in driver.current_url:
                aol_instru = "//body[1]/div[1]/div[1]/div[3]/div[3]/" \
                             "div[1]/div[1]/div[2]/div[1]/table[1]/" \
                             "tbody[1]/tr[1]/td "
                inst_node = driver.find_elements(By.XPATH, aol_instru)

            node_count = len(inst_node)
            Data.sheet.cell(row=instrument_index + 2,
                            column=5).value = node_count  # Writing the no of nodes to workbook
            Data.Workbook.save("Files\\UserFile.xlsx")
            print("No of Nodes available=" + str(node_count))
            node = 0
            explicit_wait = WebDriverWait(driver, 5)
            amend_list = []

            while node <= node_count:  # Initialising the node and iterating till the nth node
                driver.find_element(By.ID, "expand_button").click()
                node = node + 1
                amend_link = driver.find_elements(By.XPATH, aol_instru + "[" +
                                                  str(node) + "]" + "/a[last()]")
                for links in amend_link:  # To get the instrument document link
                    amend_url = links.text
                    amend_list.append(amend_url)
                if node <= node_count:  # Loading the nodes one by one till the nth node
                    date_element = driver.find_element(By.XPATH, aol_instru +
                                                       "[" + str(node) + "]")
                    get_date = date_element.get_attribute("innerText")
                    node_date = get_date.split()[0]
                    explicit_wait.until(EC.element_to_be_clickable((By.XPATH,
                                                                    aol_instru + "[" + str(
                                                                        node) + "]"))).click()
                    node_errlist = ["The Versioned Legislation Database System is"
                                    " presently unable to locate the "
                                    "legislation contemplated for the hyperlink"
                                    " that you have selected.  You may wish "
                                    "to locate the legislation concerned by using"
                                    " the Search or Browse functions of "
                                    "the Versioned Legislation Database System.Thank you."]
                    if "aol" in driver.current_url:
                        node_text = driver.execute_script('return document.getElementsByTagName'
                                                          '("span")[3].textContent')
                    elif "ldr" in driver.current_url:
                        node_text = driver.execute_script('return document.getElementsByTagName'
                                                          '("span")[1].textContent')
                    for node_check, _ in enumerate(node_errlist):
                        try:
                            assert node_errlist[node_check] not in node_text
                        except AssertionError:
                            print("The link is broken for the inst_node:" + str(node))
                            node_url = driver.execute_script('return document.URL')
                            print(node_url)
                            ###### Drafting Output for Nodes Section ###########
                            # ***********  Node Number output section  ************

                            if Data.sheet.cell(row=instrument_index + 2, column=6).value is None:
                                Data.sheet.cell(row=instrument_index + 2, column=6).value = str(
                                    node) + "-" + node_date + ","
                            elif not Data.sheet.cell(row=instrument_index + 2,
                                                     column=6).value is None:
                                prev_node = Data.sheet.cell(row=instrument_index + 2,
                                                            column=6).value
                                Data.sheet.cell(row=instrument_index + 2,
                                                column=6).value = "".join(
                                    [prev_node, "\n", str(node), "-", node_date]) + ","

                            Data.Workbook.save("Files\\UserFile.xlsx")
                            # *************   Broken Node URL ***************

                            if Data.sheet.cell(row=instrument_index + 2, column=7).value is None:
                                Data.sheet.cell(row=instrument_index + 2,
                                                column=7).value = node_url + ","
                            elif not Data.sheet.cell(row=instrument_index + 2,
                                                     column=7).value is None:
                                prev_amend = Data.sheet.cell(row=instrument_index + 2,
                                                             column=7).value
                                Data.sheet.cell(row=instrument_index + 2, column=7).value = "".join(
                                    [prev_amend, "\n", node_url]) + ","

                            Data.Workbook.save("Files\\UserFile.xlsx")
                    driver.back()

            # ***********   Document Loading Section ******#
            print("The list of Amending documents present are:")
            for amend in range(0, node_count):  # To iterate over the amending links
                if amend < node_count:
                    if amend_list[amend] != '':
                        driver.find_element(By.LINK_TEXT, amend_list[amend]).click()
                        docerr_list = ["The Versioned Legislation Database System is "
                                       "presently unable to locate the "
                                       "legislation contemplated for the hyperlink that you "
                                       "have selected.  You may "
                                       "wish to locate the legislation concerned by using "
                                       "the Search or Browse "
                                       "functions of the Versioned Legislation Database "
                                       "System.Thank you.",
                                       "No records found."]
                        if "aol" in driver.current_url:
                            doc_text = driver.execute_script(
                                'return document.getElementsByTagName("span")[3].textContent')
                        elif "ldr" in driver.current_url:
                            doc_text = driver.execute_script(
                                'return document.getElementsByTagName("span")[1].textContent')
                        for amend_check, _ in enumerate(docerr_list):
                            try:  # Type-1 Error Check
                                assert docerr_list[amend_check] not in doc_text
                            except AssertionError:
                                ###########   Drafting Output for Amend Documents Section #######
                                # *******  AmendNote  *****
                                print(amend_list[amend] + "-" + docerr_list[amend_check])
                                doc_url = driver.execute_script('return document.URL')

                                if Data.sheet.cell(row=instrument_index + 2,
                                                   column=8).value is None:
                                    Data.sheet.cell(row=instrument_index + 2,
                                                    column=8).value = amend_list[amend] + ","
                                elif not Data.sheet.cell(row=instrument_index + 2,
                                                         column=8).value is None:
                                    prev_node = Data.sheet.cell(row=instrument_index + 2,
                                                                column=8).value
                                    Data.sheet.cell(row=instrument_index + 2,
                                                    column=8).value = "".join(
                                        [prev_node, "\n", amend_list[amend]]) + ","

                                Data.Workbook.save("Files\\UserFile.xlsx")
                                # *********   AmendNote url  ***************

                                if Data.sheet.cell(row=instrument_index + 2,
                                                   column=9).value is None:
                                    Data.sheet.cell(row=instrument_index + 2,
                                                    column=9).value = doc_url + ","
                                elif not Data.sheet.cell(row=instrument_index + 2,
                                                         column=9).value is None:
                                    prev_amend = Data.sheet.cell(row=instrument_index + 2,
                                                                 column=9).value
                                    Data.sheet.cell(row=instrument_index + 2,
                                                    column=9).value = "".join([prev_amend, "\n",
                                                                               doc_url]) + ","

                                Data.Workbook.save("Files\\UserFile.xlsx")

                        driver.back()
                        explicit_wait.until(EC.element_to_be_clickable((By.ID,
                                                                        "expand_button"))).click()
                        print(amend_list[amend])

            # ********* PDF Loading Section   *****#
            no_pdf = driver.find_elements(By.XPATH, "//td/img[@alt='Blank']")
            print("No of Blanks(without PDF links)=" + str(len(no_pdf)))
            pdf_doc = driver.find_elements(By.XPATH, "//td/a/img[@alt='View PDF']")

            print("No of PDF links=" + str(len(pdf_doc)))
            pdf_len = len(pdf_doc)
            pdf_err = ["Page not found", "Could not find pdf file with url:",
                       "No records found.", "The Versioned Legislation Database "
                                            "System is presently unable to locate the legislation "
                                            "contemplated for the hyperlink "
                                            "that you have selected. "
                                            "You may wish to locate the legislation "
                                            "concerned by using the Search or Browse "
                                            "functions of the "
                                            "Versioned Legislation Database System."]

            if "aol" in driver.current_url:
                for link in pdf_doc:
                    link.click()
                    driver.switch_to.window(parent)
                amend = 1

                while amend <= pdf_len:
                    child = driver.window_handles[amend]
                    driver.switch_to.window(child)
                    for pdf_chk, _ in enumerate(pdf_err):
                        try:
                            pdf_text = driver.execute_script('return document.body.innerText')
                            pdf_url = driver.execute_script('return document.URL')
                            assert pdf_err[pdf_chk] not in pdf_text
                        except AssertionError:
                            ######    Drafting Output for PDF Section ##########
                            # ***********  PDF URL ***************

                            if Data.sheet.cell(row=instrument_index + 2, column=10).value is None:
                                Data.sheet.cell(row=instrument_index + 2,
                                                column=10).value = pdf_url + ","
                            elif not Data.sheet.cell(row=instrument_index + 2,
                                                     column=10).value is None:
                                prev_node = Data.sheet.cell(row=instrument_index + 2,
                                                            column=10).value
                                Data.sheet.cell(row=instrument_index + 2,
                                                column=10).value = "".join(
                                    [str(prev_node), "\n", pdf_url]) + ","

                            Data.Workbook.save("Files\\UserFile.xlsx")
                            # *****  PDF Number *******

                            if Data.sheet.cell(row=instrument_index + 2, column=11).value is None:
                                Data.sheet.cell(row=instrument_index + 2,
                                                column=11).value = str(amend) + ","
                            elif not Data.sheet.cell(row=instrument_index + 2,
                                                     column=11).value is None:
                                prev_amend = Data.sheet.cell(row=instrument_index + 2,
                                                             column=11).value
                                Data.sheet.cell(row=instrument_index + 2,
                                                column=11).value = "".join(
                                    [str(prev_amend), "\n", str(amend)]) + ","

                            Data.Workbook.save("Files\\UserFile.xlsx")
                    close_tab.append(driver.window_handles[amend])
                    amend += 1

            if "ldr" in driver.current_url:
                for idx, _ in enumerate(pdf_doc, 1):
                    driver.find_element(By.XPATH, "(//img[@alt='View PDF'])["
                                        + str(idx) + "]").click()
                    for pdf_chk, _ in enumerate(pdf_err):
                        try:
                            pdf_text = driver.execute_script('return document.body.innerText')
                            pdf_url = driver.execute_script('return document.URL')
                            assert pdf_err[pdf_chk] not in pdf_text
                        except AssertionError:
                            ######    Drafting Output for PDF Section ##########
                            # ***********  PDF URL  ***************

                            if Data.sheet.cell(row=instrument_index + 2, column=10).value is None:
                                Data.sheet.cell(row=instrument_index + 2,
                                                column=10).value = pdf_url + ","
                            elif not Data.sheet.cell(row=instrument_index + 2,
                                                     column=10).value is None:
                                prev_node = Data.sheet.cell(row=instrument_index + 2,
                                                            column=10).value
                                Data.sheet.cell(row=instrument_index + 2,
                                                column=10).value = "".join(
                                    [str(prev_node), "\n", pdf_url]) + ","

                            Data.Workbook.save("Files\\UserFile.xlsx")
                            # *****  PDF Number  *******

                            if Data.sheet.cell(row=instrument_index + 2, column=11).value is None:
                                Data.sheet.cell(row=instrument_index + 2,
                                                column=11).value = str(idx) + ","
                            elif not Data.sheet.cell(row=instrument_index + 2,
                                                     column=11).value is None:
                                prev_amend = Data.sheet.cell(row=instrument_index + 2,
                                                             column=11).value
                                Data.sheet.cell(row=instrument_index + 2,
                                                column=11).value = "".join(
                                    [str(prev_amend), "\n", str(idx)]) + ","

                            Data.Workbook.save("Files\\UserFile.xlsx")
                    driver.back()
                    explicit_wait.until(EC.element_to_be_clickable((By.ID,
                                                                    "expand_button"))).click()

            for tab, _ in enumerate(close_tab):
                driver.switch_to.window(close_tab[tab])
                driver.close()
                driver.switch_to.window(parent)

            # Node Icon and mouse_hover Text Comparison
            #### Image link ###
            images = driver.find_elements(By.XPATH, "//a//span//img")
            image_link = []
            for element in images:
                image_link.append(element.get_attribute("src"))

            ## Icon text ####
            for icon, _ in enumerate(image_link):
                if "aol" in driver.current_url:
                    total_data = driver.find_element(By.XPATH,
                                                     "//body[1]/div[2]/div[1]/div[3]/div[4]"
                                                     "/div[1]/div[2]/div[1]/table/tbody/tr/td["
                                                     + str(icon + 1) + "]")
                if "ldr" in driver.current_url:
                    total_data = driver.find_element(By.XPATH,
                                                     "//body[1]/div[1]/div[1]/div[3]/div[3]"
                                                     "/div[1]/div[1]/div[2]/div[1]/table[1]"
                                                     "/tbody[1]/tr[1]/td["
                                                     + str(icon + 1) + "]")

                node_data = total_data.get_attribute("innerText")
                data1 = " ".join(node_data.split())
                print(data1)
                mouse_hover = data1.split()[1].lower()
                if mouse_hover in image_link[icon]:
                    print("The Mouse hover text matched with node display icon")
                elif (mouse_hover == "spent") and (
                        ("formal-cons-cur" in image_link[icon]) or
                        ("formal-cons" in image_link[icon])):
                    print("Its a Spent Node - The Mouse hover text matched with node display icon")
                elif (mouse_hover == "informal") and ("retro" in image_link[icon]):
                    print("Its a Retro Node - The Mouse hover text matched with node display icon")
                else:
                    print("Not Matched" + "No," + "Node:" + str(icon + 1))
                    if Data.sheet.cell(row=instrument_index + 2,
                                       column=12).value is None:
                        Data.sheet.cell(row=instrument_index + 2,
                                        column=12).value = "No," + "Node:" + str(icon + 1)
                    elif not Data.sheet.cell(row=instrument_index + 2,
                                             column=12).value is None:
                        prev_ico = Data.sheet.cell(row=instrument_index + 2,
                                                   column=12).value
                        Data.sheet.cell(row=instrument_index + 2,
                                        column=12).value = prev_ico + "\n" \
                                                           + "No," + "Node:" + str(
                            icon + 1)
                    Data.Workbook.save("Files\\UserFile.xlsx")
                ####Check for Amended by string on nodes ###
                if icon + 1 == 1:
                    amd_chk = " ".join(data1.split(' ')[6:])
                if (icon + 1 == 1) and (("Act" in data1) or ("RevEd" in data1) or ("SL" in data1)):
                    print("Its a OE node")
                elif (icon + 1 == 1) and ("Amended by" in data1):
                    print("The First node doesn't seems to be a OE node:" + str(icon + 1))
                    Data.sheet.cell(row=instrument_index + 2,
                                    column=13).value = "The First node doesn't seems to " \
                                                       "be a OE node:" + str(icon + 1)
                elif ("Spent" in data1) or ("RevEd" in data1) or \
                        ("Reprint" in data1) or (amd_chk in data1):
                    print("As it is a Spent/Revised/Reprint node, it not starts with Amended by")
                elif ("Amended by" in data1) or ("Repealed by" in data1):
                    print("The node starts with Amended by")
                else:
                    print("The node doesn't starts with Amended by:" + str(icon + 1))
                    if Data.sheet.cell(row=instrument_index + 2,
                                       column=13).value is None:
                        Data.sheet.cell(row=instrument_index + 2,
                                        column=13).value = "Not begins with Amended by," \
                                                           "for the node:" + str(icon + 1)
                    elif not Data.sheet.cell(row=instrument_index + 2,
                                             column=13).value is None:
                        prev_string = Data.sheet.cell(row=instrument_index + 2,
                                                      column=13).value
                        Data.sheet.cell(row=instrument_index + 2,
                                        column=13).value = prev_string + "\n" \
                                                           + "Not begins with Amended by,for the " \
                                                             "node:" + str(icon + 1)

                    Data.Workbook.save("Files\\UserFile.xlsx")

            close_tab = None
        driver.quit()
