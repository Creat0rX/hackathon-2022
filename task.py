
### Settings ###
"""Get coronavirus case statistics from the world and USA states, store them in a database, and create a powerpoint with findings"""
import os
from datetime import datetime
from selenium import webdriver
from time import sleep
from pptx import Presentation
from pptx.util import Pt
import sqlite3 as sqlite

### Variables ###
XPATH_MAIN = "//table[@id='main_table_countries_today']/tbody/tr[@role='row'][@class='odd' or @class='even']"
XPATH_USA = "//table[@id='usa_table_countries_today']/tbody/tr[not(@class='total_row_usa odd' or @class='total_row')]"
XPATH_GRAPH = "//div[@class='col-md-12']/div"
COUNTRY_LIST = ["Country link", "Rank", "Country", "Total cases", "New cases", "Total deaths", "New deaths",
                "Total recovered", "New recovered", "Active cases", "Serious cases", "Cases per mil",
                "Deaths per mil", "Total tests", "Tests per mil", "Population"]
STATE_LIST = ["State link", "Rank", "State", "Total cases", "New cases", "Total deaths", "New deaths",
                "Total recovered", "Active cases", "Cases per mil",
                "Deaths per mil", "Total tests", "Tests per mil", "Population"]
SLASH = os.sep
sql_bool = False
driver = None
conn = None
main_table_countries = []
main_table_states = []

### Keywords ###
def open_browser():
    global driver
    driver = webdriver.Chrome()
    driver.get("https://www.worldometers.info/coronavirus/")

def scrape_table_from_website():
    table = driver.find_elements_by_xpath(XPATH_MAIN)
    for element in table:
        main_table_countries.append(insert_values_safely(element, COUNTRY_LIST))

def scrape_us_table_from_website():
    driver.get("https://www.worldometers.info/coronavirus/country/us/")
    table = driver.find_elements_by_xpath(XPATH_USA)
    for element in table:
        main_table_states.append(insert_values_safely(element, STATE_LIST))
    print("done")

def screenshot_us_graphs_from_website():
    for state in main_table_states[:10]:
        driver.get(state["State link"])
        total_cases = driver.find_element_by_xpath(f"{XPATH_GRAPH}[@class='tabbable-panel-cases']")
        daily_new_cases = driver.find_element_by_xpath(f"{XPATH_GRAPH}[@id='graph-cases-daily']")
        active_cases = driver.find_element_by_xpath(f"{XPATH_GRAPH}[@id='graph-active-cases-total']")
        total_deaths = driver.find_element_by_xpath(f"{XPATH_GRAPH}[@class='tabbable-panel-deaths']")
        total_cases.screenshot(f"output{SLASH}total_cases.png")
        daily_new_cases.screenshot(f"output{SLASH}daily_new_cases.png")
        active_cases.screenshot(f"output{SLASH}active_cases.png")
        total_deaths.screenshot(f"output{SLASH}total_deaths.png")
        add_to_powerpoint(state["State"])
        print("done with 1")

def insert_values_safely(element, list_choice):
    table_values = {}
    for index, header in enumerate(list_choice):
        try:
            if index == 0:
                table_values[header] = element.find_element_by_xpath(f"./td[2]/a").get_attribute("href")
                continue
            else:
                table_values[header] = element.find_element_by_xpath(f"./td[{index}]").text
        except:
            table_values[header] = ""
    return table_values

def connect_to_sql_database(conn):
    conn = sqlite.connect(f"output{SLASH}corona.db")
    return conn

def add_sql_tables(conn, sql_bool):
    if conn is not None:
        cursor = conn.cursor()
        tables_exist = len(list(cursor.execute("SELECT name FROM sqlite_master WHERE type='table';").fetchall()))
        if not tables_exist:
            sql_bool = True

        sql_countries = f"CREATE TABLE IF NOT EXISTS countries(Rank integer PRIMARY KEY, {' text,'.join(COUNTRY_LIST[2:])} text);"
        sql_states = f"CREATE TABLE IF NOT EXISTS american_states(Rank integer PRIMARY KEY, {' text,'.join(STATE_LIST[2:])} text);"

        cursor.execute(sql_countries)
        cursor.execute(sql_states)
        conn.commit()
        add_values_to_sql_tables(conn, sql_bool)
    else:
        print("Failed to establish a connection with the database.")

    return sql_bool

def add_values_to_sql_tables(conn, sql_bool):
    if sql_bool:
        pass
    else:
        pass
        
        # for country in COUNTRY_LIST:
        #     pass

        # for state in STATE_LIST:
        #     pass


def disconnect_from_sql_database(conn):
    if conn is not None:
        conn.close()

def add_to_powerpoint(title):
    try:
        ppt = Presentation(f"output{SLASH}presentation-{str(datetime.now().date())}.pptx")
    except:
        ppt = Presentation()
    title_slide = ppt.slide_layouts[5]
    slide = ppt.slides.add_slide(title_slide)
    slide.shapes.title.text = f"USA - {title}"
    slide.shapes.add_picture(f"output{SLASH}total_cases.png", left=Pt(50), top=Pt(100), height=Pt(200), width=Pt(300))
    slide.shapes.add_picture(f"output{SLASH}active_cases.png", left=Pt(50), top=Pt(310), height=Pt(200), width=Pt(300))
    slide.shapes.add_picture(f"output{SLASH}daily_new_cases.png", left=Pt(400), top=Pt(100), height=Pt(200), width=Pt(300))
    slide.shapes.add_picture(f"output{SLASH}total_deaths.png", left=Pt(400), top=Pt(310), height=Pt(200), width=Pt(300))
    ppt.save(f"output{SLASH}presentation-{str(datetime.now().date())}.pptx")

def close_browser():
    if driver is not None:
        driver.close()

def main():
    try:
        open_browser()
        scrape_table_from_website(driver)
    except Exception as e:
        print(f"An error occurred: {str(e)}")
    finally:
        close_browser()

### Tasks ###
if __name__ == "__main__":
    print(__doc__)
    open_browser()
    scrape_us_table_from_website()
    screenshot_us_graphs_from_website()
    close_browser()
