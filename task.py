
### Settings ###
"""Get coronavirus case statistics from the world and USA states, store them in a database, and create a powerpoint with findings"""
import os
from datetime import datetime
from selenium import webdriver
from pptx import Presentation
from pptx.util import Pt
import sqlite3 as sqlite

### Variables ###
XPATH_MAIN = "//table[@id='main_table_countries_today']/tbody/tr[@role='row'][@class='odd' or @class='even']"
XPATH_USA = "//table[@id='usa_table_countries_today']/tbody/tr[not(@class='total_row_usa odd' or @class='total_row')]"
XPATH_GRAPH = "//div[@class='col-md-12']/div"
COUNTRY_LIST = ["Country_link", "Rank", "Country", "Total_cases", "New_cases", "Total_deaths", "New_deaths",
                "Total_recovered", "New_recovered", "Active_cases", "Serious_cases", "Cases_per_mil",
                "Deaths_per_mil", "Total_tests", "Tests_per_mil", "Population"]
STATE_LIST = ["State_link", "Rank", "State", "Total_cases", "New_cases", "Total_deaths", "New_deaths",
                "Total_recovered", "Active_cases", "Cases_per_mil", "Deaths_per_mil", "Total_tests",
                "Tests_per_mil", "Population"]
SLASH = os.sep
driver = None
conn = None
main_table_countries = []
main_table_states = []

### Keywords ###
def delete_powerpoint_today_if_exists():
    print("delete_powerpoint_today_if_exists")
    if os.path.exists(f"output{SLASH}presentation-{str(datetime.now().date())}.pptx"):
        os.remove(f"output{SLASH}presentation-{str(datetime.now().date())}.pptx")

def open_browser():
    print("open_browser")
    global driver
    driver = webdriver.Chrome()
    driver.implicitly_wait(7)
    driver.get("https://www.worldometers.info/coronavirus/")

def scrape_table_from_website():
    print("scrape_table_from_website")
    table = driver.find_elements_by_xpath(XPATH_MAIN)
    for element in table:
        main_table_countries.append(insert_values_safely(element, COUNTRY_LIST))

def scrape_us_table_from_website():
    print("scrape_us_table_from_website")
    driver.get("https://www.worldometers.info/coronavirus/country/us/")
    table = driver.find_elements_by_xpath(XPATH_USA)
    for element in table:
        main_table_states.append(insert_values_safely(element, STATE_LIST))

def screenshot_us_graphs_from_website():
    print("screenshot_us_graphs_from_website")
    for state in main_table_states[:10]:
        driver.get(state["State_link"])
        total_cases = driver.find_element_by_xpath(f"{XPATH_GRAPH}[@class='tabbable-panel-cases']")
        daily_new_cases = driver.find_element_by_xpath(f"{XPATH_GRAPH}[@id='graph-cases-daily']")
        active_cases = driver.find_element_by_xpath(f"{XPATH_GRAPH}[@id='graph-active-cases-total']")
        total_deaths = driver.find_element_by_xpath(f"{XPATH_GRAPH}[@class='tabbable-panel-deaths']")
        total_cases.screenshot(f"output{SLASH}total_cases.png")
        daily_new_cases.screenshot(f"output{SLASH}daily_new_cases.png")
        active_cases.screenshot(f"output{SLASH}active_cases.png")
        total_deaths.screenshot(f"output{SLASH}total_deaths.png")
        add_to_powerpoint(state["State"])

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

def connect_to_sql_database():
    print("connect_to_sql_database")
    conn = sqlite.connect(f"output{SLASH}corona.db")
    return conn

def add_sql_tables(conn):
    print("add_sql_tables")
    if conn is not None:
        cursor = conn.cursor()

        tables_exist = len(list(cursor.execute("SELECT name FROM sqlite_master WHERE type='table';").fetchall()))
        if not tables_exist:
            sql_bool = False
        else:
            sql_bool = True

        sql_countries = f"""CREATE TABLE IF NOT EXISTS countries(
                        Rank integer PRIMARY KEY, {' text,'.join(COUNTRY_LIST[2:])} text);"""
        sql_states = f"""CREATE TABLE IF NOT EXISTS american_states(
                        Rank integer PRIMARY KEY, {' text,'.join(STATE_LIST[2:])} text);"""

        cursor.execute(sql_countries)
        cursor.execute(sql_states)
        conn.commit()
        add_values_to_sql_tables(conn, sql_bool)
    else:
        print("Failed to establish a connection with the database.")

def add_values_to_sql_tables(conn, sql_bool):
    print("add_values_to_sql_tables")
    cursor = conn.cursor()
    if sql_bool:
        # UPDATE SQL IF TABLES EXISTED
        query1 = f"""UPDATE countries SET {' = ? , '.join(COUNTRY_LIST[2:])} = ? WHERE Rank = ?;"""
        query2 = f"""UPDATE american_states SET {' = ? , '.join(STATE_LIST[2:])} = ? WHERE Rank = ?;"""
        for country in main_table_countries:
            country_details = (country["Country"], country["Total_cases"], country["New_cases"], 
                country["Total_deaths"], country["New deaths"], country["Total_recovered"], 
                country["New_recovered"], country["Active_cases"], country["Serious_cases"], 
                country["Cases_per_mil"], country["Deaths_per_mil"], country["Total_tests"], 
                country["Tests_per_mil"], country["Population"], int(country["Rank"]))
            cursor.execute(query1, country_details)
            conn.commit()
        for state in STATE_LIST:
            state_details = (state["State"], state["Total_cases"], state["New_cases"], state["Total_deaths"],
                state["New_deaths"], state["Total_recovered"], state["Active_cases"], state["Cases_per_mil"],
                state["Deaths_per_mil"], state["Total_tests"], state["Tests_per_mil"], state["Population"],
                int(state["Rank"]))
            cursor.execute(query2, state_details)
            conn.commit()
    else:
        # INSERT SQL IF TABLES NEVER EXISTED
        query1 = f"""INSERT INTO countries ({', '.join(COUNTRY_LIST[1:])})
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?);"""
        query2 = f"""INSERT INTO american_states ({', '.join(STATE_LIST[1:])})
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?);"""
        for country in main_table_countries:
            country_details = (int(country["Rank"]), country["Country"], country["Total_cases"], 
                country["New_cases"], country["Total_deaths"], country["New_deaths"],
                country["Total_recovered"], country["New_recovered"], country["Active_cases"],
                country["Serious_cases"], country["Cases_per_mil"], country["Deaths_per_mil"],
                country["Total_tests"], country["Tests_per_mil"], country["Population"])
            cursor.execute(query1, country_details)
            conn.commit()
        for state in STATE_LIST:
            state_details = (int(state["Rank"]), state["State"], state["Total_cases"], state["New_cases"],
                state["Total_deaths"], state["New_deaths"], state["Total_recovered"],
                state["Active_cases"], state["Cases_per_mil"], state["Deaths_per_mil"],
                state["Total_tests"], state["Tests_per_mil"], state["Population"])
            cursor.execute(query2, state_details)
            conn.commit()


def disconnect_from_sql_database():
    print("disconnect_from_sql_database")
    if conn is not None:
        conn.close()

def add_to_powerpoint(title):
    print("add_to_powerpoint")
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
    print("close_browser")
    if driver is not None:
        driver.close()

def main():
    try:
        open_browser()
        scrape_table_from_website()
        scrape_us_table_from_website()
        # conn = connect_to_sql_database()
        # add_sql_tables(conn)
        delete_powerpoint_today_if_exists()
        screenshot_us_graphs_from_website()
    except Exception as e:
        print(f"An error occurred: {str(e)}")
    finally:
        disconnect_from_sql_database()
        close_browser()

### Tasks ###
if __name__ == "__main__":
    print(__doc__)
    main()
