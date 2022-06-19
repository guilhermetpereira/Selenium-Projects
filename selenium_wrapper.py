from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from time import sleep
import datetime as dt
import json

DC_CELL_XPATH = "//table[@class='Calendar']/tbody/tr/td/a"
PRJ_REPORT_TABLE_ID = "ctl00_ContentPlaceHolder1_ctl00_GridViewUserMissions"
PRJ_REPORT_ROWS_XPATH = "//table[@id='{}']/tbody/tr[not(@class)]".format(PRJ_REPORT_TABLE_ID)
PRJ_REPORT_CELLS_XPATH = "{}/td[not(@align)]".format(PRJ_REPORT_ROWS_XPATH)
DAY_MESSAGE_DATE_ID = "ctl00_ContentPlaceHolder1_ctl00_lblErrorDate"
DAY_MESSAGE_MSG_ID = "ctl00_ContentPlaceHolder1_ctl00_lblAlert"
REPORT_PROJ_BTN_ID = "ctl00_ContentPlaceHolder1_ctl00_btnReportMission"

class SeleniumWrapper:
    def __init__(self, url="http://subsbrdcp/dc/"):
        # The options is just to shut up some warnings
        options = webdriver.ChromeOptions()
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        self.driver = webdriver.Chrome(options=options)
        self.set_url(url)
        self.wait = WebDriverWait(self.driver, 20)
        
    def click_by_xpath(self, path):
        try:
            self.wait.until(EC.element_to_be_clickable((By.XPATH, path)))
        except:
            print("Couldn't find xpath {}".format(path))

    def click_by_id(self, id):
        try:
            element = self.wait.until(EC.element_to_be_clickable((By.ID, id)))
            element.click()
        except:
            print("Couldn't find element {}".format(id))

    def set_text_by_id(self, id, text,enter=False):
        try:
            element = self.wait.until(EC.element_to_be_clickable((By.ID, id)))
            self.driver.execute_script(f"arguments[0].value='{text}'", element)
            if enter:
                element.send_keys(Keys.ENTER)
        except Exception as e:
            print("Couldn't find element {}".format(id))
            print("Error {}".format(e))

    def set_url(self, url):
        try:
            self.driver.get(url)
            sleep(5)
        except:
            raise AssertionError("Could not reach page {}".format(url))

    def get_elements_by_xpath(self, xpath):
        try:
            elements = self.driver.find_elements(By.XPATH, xpath)
            return elements
        except Exception as e:
            print("Failed to get elements with error: {}".format(e))

    def set_select_by_id(self, id, text):
        element_select = self.get_elements_by_xpath(f'//*[@id="{id}"]')
        select = Select(element_select[0])
        select.select_by_visible_text(text)


class ParserFromTdToDays():
    def __init__(self,webDriver):
        assert webDriver
        self.webDriver = webDriver
        self.tds = self.days  = []
        self.curr_td = None
        self.update_days()
        self.update_curr_td()

    def get_day_element(self, day):
        self.update_days()
        for idx, val in enumerate(self.tds):
            itr_day = self.title_to_date_obj( val.get_attribute("title") ) 
            if itr_day == day:
                return val

        return False

    def title_to_date_obj(self, title):
        try:
            return dt.datetime.strptime(title + " {}".format(dt.date.today().year), "%B %d %Y")
        except:
            return None

    def update_days(self):
        self.tds = self.webDriver.get_elements_by_xpath(DC_CELL_XPATH)
        for element in self.tds:
            title = element.get_attribute('title')
            date_obj = self.title_to_date_obj(title)
            if date_obj is not None:
                self.days.append( date_obj )

    def update_curr_td(self):
        updated = False
        self.update_days()
        for td in self.tds:
            parent = td.find_element(by=By.XPATH, value="..")
            if "CalendarSelectedCellImgae" in parent.get_attribute("class"):
                if self.curr_td != td:
                    updated = True
                self.curr_td = td
                break
        return updated

    def wait_for_update(self, day):
        timeout = 5
        while timeout > 0:
            try:
                curr_date_element = self.webDriver.get_elements_by_xpath(f"//*[@id='{DAY_MESSAGE_DATE_ID}']")
                curr_date = dt.datetime.strptime(curr_date_element[0].text, "%d/%m/%Y")
                self.update_curr_td()
                if curr_date == day:
                    break
            except:
                pass
            timeout -= 1
            sleep(2)

class ProjectReporter():
    def __init__(self, webDriver):
        assert webDriver
        self.webDriver = webDriver
        self.day = self.start = self.type = self.proj = self.operation = self.description = self.total = None
        self.update()

    def set_proj_and_pca(self, proj, pca):
        ENTRY_TYPE_SELECTOR_ID = "ctl00_ContentPlaceHolder1_ctl00_ddlHRTP"
        PROJ_INPUT_ID = "ctl00_ContentPlaceHolder1_ctl00_txtPrj"
        PCA_INPUT_ID = "ctl00_ContentPlaceHolder1_ctl00_txtMission"
        self.webDriver.set_select_by_id( id=ENTRY_TYPE_SELECTOR_ID, text="Projeto" )
        self.webDriver.set_text_by_id( id=PROJ_INPUT_ID, text=proj )
        self.webDriver.set_text_by_id( id=PCA_INPUT_ID, text=pca )

    def update(self):
        self.day, self.start, self.type, self.proj, self.operation, self.description, self.total = self.webDriver.get_elements_by_xpath(xpath=PRJ_REPORT_CELLS_XPATH)
        return self.webDriver.get_elements_by_xpath(xpath=PRJ_REPORT_CELLS_XPATH)

    def wait_for_update(self):
        timeout = 15
        while timeout > 0:
            try:
                self.update()
                if self.description.text != "":
                    break
            except:
                pass
            timeout -= 1
            sleep(3)


def go_to_date(parserFromTdToDays, day):
    parserFromTdToDays.update_days()
    dayElement = parserFromTdToDays.get_day_element(day)
    dayElement.click()
    parserFromTdToDays.wait_for_update(day)
    if day.day >= dt.datetime.today().day and day.month >= dt.datetime.today().month:
        return False
    return True

if __name__ == "__main__":
    file = open("config.json")
    config = json.load(file)
    seleniumWrapper = SeleniumWrapper()
    seleniumWrapper.set_text_by_id(id="txtWorkerId", text=config["username"])
    seleniumWrapper.set_text_by_id(id="TxtPassword", text=config["password"], enter=True)
    projectReporter = ProjectReporter(seleniumWrapper)
    parserFromTdToDays = ParserFromTdToDays(seleniumWrapper)

    workDays = [ day for day in parserFromTdToDays.days if day.weekday() < 5 ]
    for day in workDays:
        if not go_to_date(parserFromTdToDays, day):
            break
        projectReporter.update()
        message = seleniumWrapper.get_elements_by_xpath(f"//*[@id='{DAY_MESSAGE_MSG_ID}']")[0]
        if message.text != "OK" or projectReporter.description.text == "":
            projectReporter.set_proj_and_pca(config["project"], config["pca"])
            seleniumWrapper.click_by_id(REPORT_PROJ_BTN_ID)
            projectReporter.wait_for_update()   

    seleniumWrapper.driver.close()