# ------------- Selenium Python Test Framework AOL(Phase-1)------------#


""" Automating the test cases for AOL Site for Table Of Contents"""
import re
from selenium.common.exceptions import NoSuchElementException, JavascriptException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from excel_data import Data
from excel_data import XML
from selenium.webdriver.support.select import Select


browser = Service("C:\\msedgedriver.exe")


def remove_string(onlytext):
    """Function for filtering only text in section headings"""
    string_only = onlytext.split(' ')
    if string_only.index(string_only[-1]) == 1:
        output = string_only[-1]
    else:
        output = ' '.join(string_only[1:])
    return output


# Initialising driver to access web elements of AOL
driver = webdriver.Edge(service=browser)
driver.get(XML.URL)
driver.maximize_window()
if "aol" in driver.current_url:
    driver.find_element(By.CSS_SELECTOR, "input[name=name]").send_keys(XML.UserName)
    driver.find_element(By.CSS_SELECTOR, "input[name=pass]").send_keys(XML.Password)
    driver.find_element(By.CSS_SELECTOR, "input[value=Login]").click()
    driver.maximize_window()


elif "ldr" in driver.current_url:
    driver.find_element(By.CSS_SELECTOR, "input[name=username]").send_keys(XML.UserName)
    driver.find_element(By.CSS_SELECTOR, "input[name=password]").send_keys(XML.Password)
    driver.find_element(By.CSS_SELECTOR, "input[value=Login]").click()
    driver.maximize_window()



class TestFramework:
    """ Main class function declaration where looping each instrument through instrument index"""

    def test_aol(self):

        """" Declaring Framework test function"""

        for instrument_index, _ in enumerate(Data.instru_list):
            global run
            driver.refresh()
            # Navigating to Browse Section
            driver.find_element(By.LINK_TEXT, "Browse").click()
            sel = Select(driver.find_element(By.XPATH, "//form[2]//select[@name='browsetype']"))
            sel.select_by_visible_text(Data.sheet.cell(row=instrument_index + 2, column=1).value)
            # ******************************  Search Section ************************#
            # No of characters present in 'Browse Versioned Legislation'
            # Defining search elements
            if driver.current_url == "http://103.107.198.34:8077/aol/home.w3p":
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
                                      " the instrument-" + Data.instru[
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
                                    instrument_title = h_elm.text
                                    duplicate_instru = driver.find_elements(By.LINK_TEXT,
                                                                            h_elm.text)
                                    dupinstru()
                                    driver.find_element(By.LINK_TEXT, h_elm.text).click()
                                    run = 1
                                    break
                            if run == 0:
                                for l_elm in low_elm:
                                    conv_web = re.sub("[^A-Za-z0-9]", "", l_elm.text).lower()
                                    if Data.instru_list[instrument_index] == conv_web:
                                        instrument_title = l_elm.text
                                        duplicate_instru = driver.find_elements(By.LINK_TEXT,
                                                                                l_elm.text)
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
                                    instrument_title = h_elm.text
                                    duplicate_instru = driver.find_elements(By.LINK_TEXT,
                                                                            h_elm.text)
                                    dupinstru()
                                    driver.find_element(By.LINK_TEXT, h_elm.text).click()
                                    run = 1
                                    break
                            if run == 0:
                                for l_elm in low_elm:
                                    conv_web = re.sub("[^A-Za-z0-9]", "", l_elm.text).lower()
                                    if Data.instru_list[instrument_index] == conv_web:
                                        instrument_title = l_elm.text
                                        duplicate_instru = driver.find_elements(By.LINK_TEXT,
                                                                                l_elm.text)
                                        dupinstru()
                                        driver.find_element(By.LINK_TEXT, l_elm.text).click()
                                        run = 1
                                        break
                    my_char += 1  # iterate over alphabets_list
            else:
                continue
            explicit_wait = WebDriverWait(driver, 20)
            parent = driver.window_handles[0]

            # Whole Document Section(Comparing by either last section heading or last section
            # with page tail content or the whole page content
            if "aol" in driver.current_url:
                toc_parts = "//body[1]/div[2]/div[1]/div[1]" \
                            "/div[2]/form[1]/div[1]/div[2]/p"
                toc_sections = "//body[1]/div[2]/div[1]/div[1]" \
                               "/div[2]/form[1]/div[1]/div[2]/blockquote"
            elif "ldr" in driver.current_url:
                toc_parts = "//html[1]/body[1]/div[1]/div[1]/div[1]" \
                            "/div[2]/form[1]/div[1]/div[2]/p"
                toc_sections = "//html[1]/body[1]/div[1]/div[1]/div[1]" \
                               "/div[2]/form[1]/div[1]/div[2]/blockquote"

            def whole_doc():
                """ Function for whole document page loading for each node"""
                if "aol" in driver.current_url:
                    last_toc_heading = driver.find_element(By.XPATH, "//body[1]/div[2]/div[1]/"
                                                                     "div[1]/div[2]/form["
                                                        "1]/div[1]/div[2]/p[last()]/a").text
                    last_toc_section = driver.find_element(By.XPATH, "//body[1]/div[2]/div[1]"
                                                                     "/div[1]/div[2]/form["
                                                        "1]/div[1]/div[2]/blockquote[last( "
                                                                     ")]/div/a").text

                elif "ldr" in driver.current_url:
                    last_toc_heading = driver.find_element(By.XPATH, "//html[1]/body[1]/div[1]"
                                                                     "/div[1]/div[1]/div["
                                                                     "2]/form[1]/div[1]/div[2]"
                                                                     "/p[last()]").text
                    last_toc_section = driver.find_element(By.XPATH, "//html[1]/body[1]/div[1]"
                                                                     "/div[1]/div[1]/div["
                                                                     "2]/form[1]/div[1]/div[2]"
                                                                     "/blockquote[last("
                                                                     ")]/div/a").text
                # ------ Whole Document Section Check ------ #
                driver.find_element(By.CSS_SELECTOR, "input[value='Whole Document']").click()
                toc_check = remove_string(last_toc_section)
                page_tail = driver.find_element(By.XPATH, "//div[@class='tail']").text
                page_content = driver.execute_script('return document.querySelector'
                                                     '("#mainText").textContent')
                try:
                    assert (last_toc_heading in page_tail) or (last_toc_section in page_tail) or \
                           (toc_check in page_content) or (page_content in last_toc_heading)
                    print("The 'Whole Document' loaded successfully.")
                except AssertionError:
                    print("The 'Whole Document' didn't load successfully.")
                    if Data.sheet.cell(row=instrument_index + 2,
                                       column=5).value is None:
                        Data.sheet.cell(row=instrument_index + 2,
                                        column=5).value = "No-" + node_point
                    elif not Data.sheet.cell(row=instrument_index + 2,
                                             column=5).value is None:
                        prev_print = Data.sheet.cell(row=instrument_index + 2,
                                                     column=5).value
                        Data.sheet.cell(row=instrument_index + 2,
                                        column=5).value = "".join([str(prev_print),
                                                                   "\n", "No-", node_point])
                Data.Workbook.save("Files\\UserFile.xlsx")
                driver.back()
                # ------ Instrument Title/ TOC Heading Check ------ #
                toc_heading = driver.find_element(By.XPATH, "//span[@class='currentTocItem']")
                explicit_wait.until(EC.element_to_be_clickable((By.XPATH,
                                                                "//span[@class="
                                                                "'currentTocItem']"))).click()
                if toc_heading.text.split('\n')[0] in page_content:
                    print(toc_heading.text)
                    print("The Page heading is present in loaded page")
                elif driver.find_element(By.XPATH, "//div[@class='SLTitle']").text.\
                    split('\n')[0].lower() in toc_heading.text.split('\n')[0].lower():
                    print(toc_heading.text)
                    print("The Page heading is present in loaded page")
                elif driver.find_element(By.XPATH, "//div[@id='aT-.']").text.\
                    split('\n')[0].lower() in toc_heading.text.split('\n')[0].lower():
                    print(toc_heading.text)
                    print("The Page heading is present in loaded page")
                else:
                    if Data.sheet.cell(row=instrument_index + 2,
                                       column=8).value is None:
                        Data.sheet.cell(row=instrument_index + 2,
                                        column=8).value = "Not for " + node_point + \
                                                          "TOC Title-" + toc_heading.text
                    elif not Data.sheet.cell(row=instrument_index + 2,
                                             column=8).value is None:
                        prev_print = Data.sheet.cell(row=instrument_index + 2,
                                                     column=8).value
                        Data.sheet.cell(row=instrument_index + 2,
                                        column=8).value = "".join([str(prev_print),
                                                                   "\n", "Not for ", node_point,
                                                                   "TOC Title-", toc_heading.text])
                    Data.Workbook.save("Files\\UserFile.xlsx")

            # ----- Section, Parts, and Division Loading section ----- #
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
                            column=6).value = node_count
            Data.Workbook.save("Files\\UserFile.xlsx")

            def parts_check(part_input):
                """ Function for checking parts for each node """
                parttext = driver.find_elements(By.XPATH, part_input)
                for part_chk, _ in enumerate(parttext, 1):
                    if part_chk in (1, 2, 3):
                        continue
                    explicit_wait.until(EC.element_to_be_clickable((By.XPATH,
                                                                    part_input + "[" + str(part_chk)
                                                                    + "]" + "/a"))).click()
                    part_txt = driver.find_element(By.XPATH,
                                                   part_input + "[" + str(part_chk) + "]").text
                    if ("Long Title" in part_txt) or ("Enacting Formula" in part_txt):
                        try:
                            page_content = driver.execute_script('return document.querySelector'
                                                                 '("#mainText").textContent')
                            assert ("An Act" in page_content) or \
                                   ("Be it enacted" in page_content)
                        except AssertionError:
                            print("The Part heading is not present in the loaded page")
                            if Data.sheet.cell(row=instrument_index + 2,
                                               column=8).value is None:
                                Data.sheet.cell(row=instrument_index + 2,
                                                column=8).value = "Not for Part-" + \
                                                                  node_point + "-" + part_txt
                            elif not Data.sheet.cell(row=instrument_index + 2,
                                                     column=8).value is None:
                                prev_print = Data.sheet.cell(row=instrument_index + 2,
                                                             column=8).value
                                Data.sheet.cell(row=instrument_index + 2,
                                                column=8).value = "".join([str(prev_print),
                                                                           "\n", "Not for Part-",
                                                                    node_point, "-", part_txt])
                            Data.Workbook.save("Files\\UserFile.xlsx")
                    else:
                        try:
                            page_content = driver.execute_script('return document.querySelector'
                                                                 '("#mainText").textContent')
                            part_text = part_txt.split(' ')
                            assert (part_text[-1] in page_content) or \
                                   (part_text[-1].upper() in page_content)
                        except AssertionError:
                            print("The Part heading is not present in the loaded page")
                            if Data.sheet.cell(row=instrument_index + 2,
                                               column=8).value is None:
                                Data.sheet.cell(row=instrument_index + 2,
                                                column=8).value = "Not for Part-" + \
                                                                  node_point + "-" + part_txt
                            elif not Data.sheet.cell(row=instrument_index + 2,
                                                     column=8).value is None:
                                prev_print = Data.sheet.cell(row=instrument_index + 2,
                                                             column=8).value
                                Data.sheet.cell(row=instrument_index + 2,
                                                column=8).value = "".join([str(prev_print),
                                                                    "\n", "Not for Part-",
                                                                    node_point, "-", part_txt])
                        Data.Workbook.save("Files\\UserFile.xlsx")

            def section_check(sect_input):
                """ Function for checking section heading for each node """
                secttext = driver.find_elements(By.XPATH, sect_input)
                for sect_check, _ in enumerate(secttext, 1):
                    explicit_wait.until(EC.visibility_of_element_located((By.XPATH,
                        sect_input + "[" + str(sect_check) + "]" + "/div/a"))).click()
                    sect_text = driver.find_element(By.XPATH,
                                                    sect_input + "[" + str(sect_check) + "]").text
                    try:
                        loaded_page_text = driver.execute_script('return document.querySelector'
                                                                 '("#mainText").textContent')
                        if not remove_string(sect_text) in loaded_page_text:
                            assert remove_string(sect_text).split(" ")[0] in loaded_page_text and\
                                   remove_string(sect_text).split(" ")[-1] in loaded_page_text
                    except AssertionError:
                        print("The Section heading is not present in the loaded page")
                        if Data.sheet.cell(row=instrument_index + 2,
                                           column=8).value is None:
                            Data.sheet.cell(row=instrument_index + 2,
                                            column=8).value = "Not for Section-" + \
                                                              node_point + "-" + sect_text
                        elif not Data.sheet.cell(row=instrument_index + 2,
                                                 column=8).value is None:
                            prev_print = Data.sheet.cell(row=instrument_index + 2,
                                                         column=8).value
                            Data.sheet.cell(row=instrument_index + 2,
                                            column=8).value = "".join([str(prev_print),
                                                                       "\n", "Not for Section-",
                                                                       node_point, "-", sect_text])
                    Data.Workbook.save("Files\\UserFile.xlsx")

            def print_check(print_input):
                """ Function for checking Print Functionality for each node"""
                global check1, aol_instru, node_text
                condition_chk = driver.find_element(By.XPATH, toc_sections + "[last()]/div/a").text
                driver.find_element(By.CSS_SELECTOR, print_input).click()
                child = driver.window_handles[1]
                driver.switch_to.window(child)
                each_sections = driver.find_elements(By.XPATH, "(//blockquote)")
                sect_list = []
                for num in each_sections:
                    sect_list.append(remove_string(num.text))

                each_parts = driver.find_elements(By.XPATH, "//body[1]/div[1]/div[1]/div[1]"
                                                            "/div[1]/form[1]/div[3]/p")
                toc = "//body[1]/div[1]/div[1]/div[1]/div[1]/form[1]/div[3]/div[1]/p"
                driver.find_element(By.CSS_SELECTOR, "input[value = ' Select All']").click()
                driver.find_element(By.XPATH, "(//input[@name='next'])[2]").click()
                child1 = driver.window_handles[2]
                driver.switch_to.window(child1)
                print_content = driver.execute_script \
                    ('return document.querySelector(".body").textContent')
                # -------- Select All option check ---------- #

                try:
                    assert remove_string(condition_chk) in print_content
                except AssertionError:
                    print("The Print Functionality didn't load the document successfully")
                    if Data.sheet.cell(row=instrument_index + 2,
                                       column=9).value is None:
                        Data.sheet.cell(row=instrument_index + 2,
                                        column=9).value = "Not for node-" \
                                                          + node_point + "-" + "Select All"
                    elif not Data.sheet.cell(row=instrument_index + 2,
                                             column=9).value is None:
                        prev_print = Data.sheet.cell(row=instrument_index + 2,
                                                     column=9).value
                        Data.sheet.cell(row=instrument_index + 2,
                                        column=9).value = "".join([str(prev_print),
                                                                   "\n", "Not for node-",
                                                                   node_point, "-", "Select All"])
                    Data.Workbook.save("Files\\UserFile.xlsx")
                driver.close()
                driver.switch_to.window(child)
                driver.find_element(By.CSS_SELECTOR, "input[value=' Clear All']").click()

                # ---- Expanding of the Parts and each Parts check if exists ---- #
                title = driver.find_element(By.XPATH, "//html[1]/body[1]/div[1]/div[1]/div[1]"
                                                      "/div[1]/form[1]/div[3]/div[1]/input"
                                                      "[@type='checkbox']")
                title.click()
                driver.find_element(By.XPATH, "(//input[@name='next'])[2]").click()
                child1 = driver.window_handles[2]
                driver.switch_to.window(child1)
                toc_title = driver.find_element(By.XPATH, "//div[@class='front']").text

                driver.close()
                driver.switch_to.window(child)
                driver.find_element(By.CSS_SELECTOR, "input[value=' Clear All']").click()
                driver.find_element(By.XPATH, toc + "/input[@type='checkbox']").click()
                driver.find_element(By.XPATH, "(//input[@name='next'])[2]").click()
                child1 = driver.window_handles[2]
                driver.switch_to.window(child1)
                toc_title1 = driver.find_element(By.CSS_SELECTOR, "#tocView").text
                driver.close()
                driver.switch_to.window(child)
                driver.find_element(By.CSS_SELECTOR, "input[value=' Clear All']").click()
                if not (title.text in toc_title) and (condition_chk in toc_title1):
                    if Data.sheet.cell(row=instrument_index + 2,
                                       column=10).value is None:
                        Data.sheet.cell(row=instrument_index + 2,
                                        column=10).value = "Not for node-" + node_point \
                                                          + "-" + "TOC/TOC Title"
                    elif not Data.sheet.cell(row=instrument_index + 2,
                                             column=10).value is None:
                        prev_print = Data.sheet.cell(row=instrument_index + 2,
                                                     column=10).value
                        Data.sheet.cell(row=instrument_index + 2,
                                        column=10).value = "".join([str(prev_print),
                                                                   "\n", "Not for node-"
                                                        , node_point, "-", "TOC/TOC Title"])
                    Data.Workbook.save("Files\\UserFile.xlsx")
                expander = driver.find_elements(By.XPATH, "//html[1]/body[1]/div[1]"
                                                          "/div[1]/div[1]/div[1]/"
                                                          "form[1]/div[3]/p/a")
                for click, _ in enumerate(expander, 6):
                    try:
                        driver.find_element(By.XPATH, "//html[1]/body[1]/div[1]"
                                                  "/div[1]/div[1]/div[1]/form[1]/div[3]/p["
                                            + str(click) + "]/a").click()
                    except NoSuchElementException:
                        pass
                try:
                    driver.find_element(By.XPATH, "//html[1]/body[1]/div[1]"
                            "/div[1]/div[1]/div[1]/form[1]/div[3]/p[" + str(5) + "]/a").click()
                except NoSuchElementException:
                    pass
                for click, _ in enumerate(each_parts, 1):
                    if click in (1, 2, 3):
                        continue
                    driver.find_element(By.XPATH, "//body[1]/div[1]/div[1]/"
                                                  "div[1]/div[1]/form[1]/div[3]/p["
                                        + str(click) + "]/input[@type='checkbox']").click()
                    driver.find_element(By.XPATH, "(//input[@name='next'])[2]").click()
                    child1 = driver.window_handles[2]
                    driver.switch_to.window(child1)
                    parts_text = driver.find_element(By.XPATH, "//html[1]/body[1]"
                                                               "/div[1]/div[1]/div[3]").text
                    driver.close()
                    driver.switch_to.window(child)
                    driver.find_element(By.CSS_SELECTOR, "input[value=' Clear All']").click()
                    try:
                        assert (any(parts_text in i for i in sect_list[0:-1])) \
                               or ("An Act", "Be it enacted" in parts_text)
                    except AssertionError:
                        print("The Parts didn't load print page")
                        if Data.sheet.cell(row=instrument_index + 2,
                                           column=11).value is None:
                            Data.sheet.cell(row=instrument_index + 2,
                                            column=11).value = "Not for node-" + \
                                                               node_point + "-" + parts_text
                        elif not Data.sheet.cell(row=instrument_index + 2,
                                                 column=11).value is None:
                            prev_print = Data.sheet.cell(row=instrument_index + 2,
                                                         column=11).value
                            Data.sheet.cell(row=instrument_index + 2,
                                            column=11).value = "".join([str(prev_print),
                                                                "\n", "Not for node-",
                                                                node_point, "-", parts_text])
                        Data.Workbook.save("Files\\UserFile.xlsx")
                        driver.close()
                        driver.switch_to.window(child)
                # ------ Each Section content check by selecting checkbox ----- #
                for sect_idx, _ in enumerate(each_sections, 1):
                    explicit_wait.until(EC.element_to_be_clickable((By.XPATH, "(//blockquote)["
                        + str(sect_idx) + "]/div[1]/input[@type='checkbox']"))).click()
                    driver.find_element(By.XPATH, "(//input[@name='next'])[2]").click()
                    try:
                        child1 = driver.window_handles[2]
                        driver.switch_to.window(child1)
                        prov_text = driver.find_element(By.XPATH, "//html[1]/body[1]/div[1]"
                                                                  "/div[1]/div[3]/div[1]"
                                                                  "/div[1]").text
                        # Slicing each sections list to filter all the sections except current index
                        check1 = sect_list[sect_idx - 1]
                        copy_list = sect_list[:sect_idx - 1] + sect_list[sect_idx:]
                        assert (check1 in prov_text) and \
                               (prov_text not in copy_list[0:-1])
                    except AssertionError:
                        print("The Print functionality for Section didn't loaded successfully for "
                              + sect_list[sect_idx])
                        if Data.sheet.cell(row=instrument_index + 2,
                                           column=12).value is None:
                            Data.sheet.cell(row=instrument_index + 2,
                                            column=12).value = "Not for node-" + \
                                                               node_point + "-" + check1
                        elif not Data.sheet.cell(row=instrument_index + 2,
                                                 column=12).value is None:
                            prev_print = Data.sheet.cell(row=instrument_index + 2,
                                                         column=12).value
                            Data.sheet.cell(row=instrument_index + 2,
                                            column=12).value = "".join([str(prev_print),
                                                                        "\n", "Not for node-",
                                                                        node_point, "-", check1])
                        Data.Workbook.save("Files\\UserFile.xlsx")
                    driver.close()
                    driver.switch_to.window(child)
                    driver.find_element(By.CSS_SELECTOR, "input[value=' Clear All']").click()
                driver.close()
                driver.switch_to.window(parent)
                driver.find_element(By.CSS_SELECTOR, "input[value='Whole Document']").click()

            def vieworinforce_link(view_inforce, name):
                """ Function for checking View & Inforce link versions for each node"""
                driver.find_element(By.XPATH, view_inforce) \
                    .click()
                child = driver.window_handles[1]
                driver.switch_to.window(child)
                try:
                    view_page = driver.execute_script('return document.querySelector'
                                                      '(".front").innerText')
                    page_title = instrument_title.upper()
                    if (page_title not in view_page) and \
                            (page_title.split(' ')[0] not in view_page):
                        print("The Link to " + name + "didn't load successfully")
                        if Data.sheet.cell(row=instrument_index + 2,
                                           column=13).value is None:
                            Data.sheet.cell(row=instrument_index + 2,
                                            column=13).value = "Not for node-" + \
                                                               node_point + "-" + name
                        elif not Data.sheet.cell(row=instrument_index + 2,
                                                 column=13).value is None:
                            prev_print = Data.sheet.cell(row=instrument_index + 2,
                                                         column=13).value
                            Data.sheet.cell(row=instrument_index + 2,
                                            column=13).value = "".join([str(prev_print),
                                     "\n", "Not for node-", node_point, "-", name])
                        Data.Workbook.save("Files\\UserFile.xlsx")
                    driver.close()
                    driver.switch_to.window(parent)

                except (NoSuchElementException, JavascriptException):
                    view_page = driver.find_element(By.XPATH, "//p/span").text
                    if err_list[0] or err_list[1] or err_list[2] in view_page:
                        print("The link to " + name + "is broken")
                        if Data.sheet.cell(row=instrument_index + 2,
                                           column=13).value is None:
                            Data.sheet.cell(row=instrument_index + 2,
                                            column=13).value = "Not for node-" + \
                                                               node_point + "-" + name
                        elif not Data.sheet.cell(row=instrument_index + 2,
                                                 column=13).value is None:
                            prev_print = Data.sheet.cell(row=instrument_index + 2,
                                                         column=13).value
                            Data.sheet.cell(row=instrument_index + 2,
                                            column=13).value = "".join([str(prev_print),
                                       "\n", "Not for node-", node_point, "-", name])
                        Data.Workbook.save("Files\\UserFile.xlsx")
                    driver.close()
                    driver.switch_to.window(parent)

            def get_prov():
                """ Function for checking GetProvision timeline nodes for each nodes"""
                global amd_chk1

                def excel(val):
                    if Data.sheet.cell(row=instrument_index + 2,
                                       column=val).value is None:
                        Data.sheet.cell(row=instrument_index + 2,
                                        column=val).value = node_point + "-Prov node-" + \
                                                            prov_date + section_text
                    elif not Data.sheet.cell(row=instrument_index + 2,
                                             column=val).value is None:
                        prev_print = Data.sheet.cell(row=instrument_index + 2,
                                                     column=val).value
                        Data.sheet.cell(row=instrument_index + 2,
                                        column=val).value = "".join([str(prev_print),
                                                "\n", node_point, "-Prov node-", prov_date,
                                                        section_text])
                    Data.Workbook.save("Files\\UserFile.xlsx")
                def excel1(val):
                    if Data.sheet.cell(row=instrument_index + 2,
                                           column=val).value is None:
                        Data.sheet.cell(row=instrument_index + 2, column=val).value\
                             = node_point + "-" + section_text + "-" + amd_url[-1]

                    elif not Data.sheet.cell(row=instrument_index + 2,
                                                 column=val).value is None:
                        prev_print = Data.sheet.cell(row=instrument_index + 2,
                                                         column=val).value
                        Data.sheet.cell(row=instrument_index + 2,
                                            column=val).value = "".join([str(prev_print),
                                                    "\n", node_point, "-", section_text,
                                                        "-", amd_url[-1]])
                    Data.Workbook.save("Files\\UserFile.xlsx")
                if "aol" in driver.current_url:
                    check_input = "//html[1]/body[1]/div[2]/div[1]/div[1]/div[2]" \
                                  "/form[1]/div[1]/div[2]/blockquote"
                    checkbox = driver.find_elements(By.XPATH, check_input + "/div/input")
                else:
                    check_input = "//html[1]/body[1]/div[1]/div[1]/div[1]/div[2]/" \
                                  "form[1]/div[1]/div[2]/blockquote"
                    checkbox = driver.find_elements(By.XPATH, check_input + "/div/input")
                for check, _ in enumerate(checkbox,1):
                    section_text = driver.find_element(By.XPATH, check_input +
                                            "[" + str(check) + "]/div/a").text

                    elem = driver.find_element(By.XPATH, check_input + "["
                                        + str(check) + "]/div/input")
                    driver.execute_script("arguments[0].scrollIntoView(true);", elem)
                    driver.execute_script("arguments[0].click()", elem)
                    driver.find_element(By.XPATH, "//input[@name='provisions']").click()
                    try: #To check if Get Provision timeline is present
                        prov_nodes = driver.find_elements(By.XPATH,
                                        "//div[@id='partTimeline1']/table/tbody/tr/td")
                    except NoSuchElementException:
                        print("No Provisional timeline was found")
                        if Data.sheet.cell(row=instrument_index + 2,
                                           column=14).value is None:
                            Data.sheet.cell(row=instrument_index + 2,
                                            column=14).value = "No Provisional timeline-" \
                                                               + node_point + section_text
                        elif not Data.sheet.cell(row=instrument_index + 2,
                                                 column=14).value is None:
                            prev_print = Data.sheet.cell(row=instrument_index + 2,
                                                         column=14).value
                            Data.sheet.cell(row=instrument_index + 2,
                                            column=14).value = "".join([str(prev_print),
                                            "\n", "No Provisional timeline-", node_point,
                                                                        section_text])
                        Data.Workbook.save("Files\\UserFile.xlsx")
                    # Amend link and Diff link check

                    amd_links = driver.find_elements(By.XPATH,
                                    "//div[@id='partTimeline1']/table/tbody/tr/td/a[2]")
                    amd_url = []
                    for links in amd_links:
                        amd_url.append(links.text)
                        href = links.get_attribute('href')
                        # open in new tab
                        driver.execute_script("window.open('%s', '_blank')" % href)
                        # Switch to new tab
                        driver.switch_to.window(driver.window_handles[-1])
                        # Diff Link Check
                        try:
                            diff_text = driver.execute_script('return document.querySelector'
                                                          '(".fragview").textContent')
                            if (diff_text == "") and ("Show Difference" not in diff_text):
                                excel1(18)
                            else:
                                act_no1 = driver.execute_script('return document.querySelector'
                                                    '(".actNo").textContent').split(' ')[1:4]
                                result = " ".join(act_no1).replace(")", "")
                                if ("Act" + " " + str(result)) not in amd_url[-1]:
                                    excel1(18)
                        except (NoSuchElementException, JavascriptException):
                            try:
                                err_page = driver.find_element(By.XPATH,
                                                            "//body/div/div/div").text
                                if err_list[0] and err_list[1] and err_list[2] in err_page:
                                    excel1(18)
                            except NoSuchElementException:
                                excel1(18)
                        driver.close()
                        driver.switch_to.window(parent)
                    # Provisional node check
                    for node, _ in enumerate(prov_nodes, 1):
                        date_element = driver.find_element(By.XPATH,
                                                           "//div[@id='partTimeline1']/table"
                                                           "/tbody/tr/td[" + str(node) + "]")
                        prov_text = date_element.get_attribute("innerText").split('\n')
                        amd_chk = " ".join(prov_text[4:])
                        prov_date = prov_text[1]
                        icon_txt = prov_text[2].split(' ')
                        icon = driver.find_element(By.XPATH, "//div[@id='partTimeline1']/table"
                                "/tbody/tr/td[" + str(node) + "]/a/span/img").get_attribute("src")
                        driver.find_element(By.XPATH,
                                            "//div[@id='partTimeline1']/table/tbody/tr/td["
                                            + str(node) + "]/a").click()
                        try:
                            pg_text = driver.execute_script('return document.querySelector'
                                                            '("#mainText").textContent')
                            #  Provision node loaded page content check
                            if not remove_string(section_text) in pg_text:
                                if not remove_string(section_text)[0] in pg_text:
                                    excel(14)
                        except (NoSuchElementException, JavascriptException):
                            if "aol" in driver.current_url:
                                err = driver.find_element(By.XPATH, "//tbody//tr[2]//td"
                                                                "//div//p/span").text
                            else:
                                err = driver.find_element(By.XPATH, "//div[@id='page-wrapper']"
                                                                    "/div[2]/p/span").text
                            if err_list[0] or err_list[1] or err_list[2] in err:
                                excel(14)
                        driver.back()

                        # Match check for mousehover and icon
                        if not icon_txt[0].lower() in icon:
                            print("The mousehover text didn't get matched with icon")
                            excel(15)
                        # 'Amended by' string check and Provision number check
                        if node == 1: # Provision number check
                            amd_chk1 = " ".join(prov_text[4:])
                            if not prov_text[0] == "pr" + section_text.split(' ')[0] + "-.":
                                excel(17)
                            if ("Act" not in amd_chk) and ("RevEd" not in amd_chk):
                                excel(16)
                        elif ("Spent" not in amd_chk) and ("Repealed by" not in amd_chk) \
                            and ("RevEd" not in amd_chk) and \
                                ("Amended by" not in amd_chk) and (amd_chk1 not in amd_chk):
                            print("Not starts with Amended by")
                            excel(16)

            # ------      MAIN TEST Initiation --------
            # __________Checking whether instrument is available with nodes/not______
            if run == 0:
                print("No instrument were found for the title - "+Data.instru[instrument_index])
                continue
            try:
                driver.find_element(By.ID, "expand_button")
            except NoSuchElementException:
                node_point = ""
                print("No nodes available:" + Data.instru_list[instrument_index])
                Data.sheet.cell(row=instrument_index + 2, column=6).value = int(0)
                Data.sheet.cell(row=instrument_index + 2,
                                column=7).value = "No nodes are present"
                Data.Workbook.save("Files\\UserFile.xlsx")
                # As the instrument don't have nodes, checking other functionalities like
                # Section headings,Print and Inforce Versions
                #  --> Whole document and title page
                whole_doc()
                # --> TOC Parts check
                parts_check(toc_parts)
                # --> TOC Sections check
                section_check(toc_sections)
                if "aol" in driver.current_url:
                    # --> Print Check
                    print_check("img[title='Print']")
                    # ---> Viewed or In-Force Version link check
                    vieworinforce_link("//img[@title='Link to Viewed Version']", "Viewed Version")
                    driver.switch_to.window(parent)
                    vieworinforce_link("//img[@title='Link to In-Force Version']",
                                       "In-Force Version")
                    driver.switch_to.window(parent)
                # --> Get Provisions check
                #get_prov()

            # Checking functionalities for Instrument with nodes available
            if node_count != 0:
                for node, _ in enumerate(inst_node, 1):
                    driver.find_element(By.ID, "expand_button").click()
                    explicit_wait.until(EC.element_to_be_clickable((By.XPATH,
                                                                    aol_instru + "[" + str(
                                                                        node) + "]"))).click()
                    err_list = ["The Versioned Legislation Database System is "
                                "presently unable to locate the "
                                "legislation contemplated for the hyperlink that you "
                                "have selected.  You may "
                                "wish to locate the legislation concerned by using "
                                "the Search or Browse "
                                "functions of the Versioned Legislation Database "
                                "System.Thank you.",
                                "No records found.", "Page not found"]
                    if "aol" in driver.current_url:
                        node_text = driver.execute_script('return document.getElementsByTagName'
                                                          '("span")[3].textContent')
                    elif "ldr" in driver.current_url:
                        node_text = driver.execute_script('return document.getElementsByTagName'
                                                          '("span")[1].textContent')
                    to_skip = None
                    for node_check, _ in enumerate(err_list):
                        if err_list[node_check] in node_text:
                            to_skip = " Exiting the current Iteration as the node is broken"
                            print(to_skip)
                            driver.back()
                    if to_skip is not None:
                        continue
                    date_element = driver.find_element(By.XPATH, aol_instru +
                                                       "[" + str(node) + "]")
                    get_date = date_element.get_attribute("innerText")
                    node_date = get_date.split()[0]
                    node_point = str(node) + "-" + node_date
                    #  --> Whole document and title page
                    whole_doc()
                    # --> TOC Parts check
                    parts_check(toc_parts)
                    # --> Section heading check
                    section_check(toc_sections)
                    if "aol" in driver.current_url:
                        # --> Print Check
                        print_check("img[title='Print']")
                        # ---> Viewed or In-Force Version link check
                        vieworinforce_link("//img[@title='Link to Viewed Version']",
                                           "Viewed Version")
                        driver.switch_to.window(parent)
                        vieworinforce_link("//img[@title='Link to In-Force Version']",
                                           "In-Force Version")
                        driver.switch_to.window(parent)
                    # --> Get Provisions check
                    get_prov()
