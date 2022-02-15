from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
import time
import pandas as pd

## data processing
def data_processing(data):
    # delete first none row
    data = data.drop(labels=[0], axis=0)

    # choose needed columns
    needed_columns = ['Last Name', 'First Name', 'Middle Name', 'Birth date']
    data = data.loc[data['Country'] == "Russia"]
    new_df = data[needed_columns]

    return new_df

# Get driver and connecting to WebPage
def get_driver():
    # webpage
    url = "https://diploma.olimpiada.ru/full-diplomas"

    # Chrome driver
    driver = webdriver.Chrome(
        executable_path="C:\\Users\\asus\\PycharmProjects\\MyPetProject\\chromedriver\\chromedriver."
                        "exe")

    # connecting to the webpage
    driver.get(url=url)
    return driver

# Process of birth date to right format
def get_bd(birth_date):
    return birth_date[-4::1] + birth_date[2:6] + birth_date[:2]

#inserting parameters in form of page
def insert_data(last_name,first_name,middle_name,bd):

    ln = driver.find_element_by_id("field_last-name")
    ln.send_keys(last_name)
    # time.sleep(1)

    fn = driver.find_element_by_id("field_first-name")
    fn.send_keys(first_name)
    # time.sleep(1)

    mn = driver.find_element_by_id("field_middle-name")
    mn.send_keys(middle_name)
    # time.sleep(1)

    birth_date = driver.find_element_by_id("field_birth-date")
    birth_date.send_keys(bd)

#clicking on button
def click():
    driver.find_element_by_class_name("enter").click()
    time.sleep(1)
    driver.find_element_by_class_name("button").click()

# Check diplomas of olympiads
def check_diplomas(xpath):
    try:
        driver.find_element_by_class_name(xpath)
    except NoSuchElementException:
        return False
    return True

# Get a list of information about diplomas
def get_info():
    return driver.find_elements_by_tag_name('tbody')

# Get a list of diplomas
def process_of_data(lst):
    x, names = [],[]
    for i in range(len(lst)-1):
        x = lst[i].text.split(sep="\n")
        name_diploma = x[0].split(sep=',')[0]
        class_diploma = x[-1].split()[-1]
        if int(class_diploma) > 9:
            names.append(name_diploma)
    return names

# 1 - read data
data = pd.read_excel("куча данных ШК.xlsx",sheet_name="R10")

# 2 - data processing
new_data = data_processing(data)

# 3 - preparing work with data
count_of_rows = new_data.shape[0]
driver = get_driver()
list_of_fullname,list_of_all_diplomas = [], []

# 4 - work with data
for i in range(count_of_rows):
    # fill parameters
    last_name = new_data.iloc[i]['Last Name']
    first_name = new_data.iloc[i]['First Name']
    middle_name = new_data.iloc[i]['Middle Name']
    birth_date = get_bd(str(new_data.iloc[i]['Birth date']))

    # inserting data on the webpage
    insert_data(last_name,first_name,middle_name,birth_date)

    # open result of inserting
    click()
    # time.sleep(2)

    # checking result of inserting
    result = check_diplomas("beauty_table")
    if result == True:
        list_diplomas = []
        # Get a list of diplomas
        info = get_info()
        full_name = last_name + " " + first_name + " " + middle_name
        # Make a records for dictionary
        list_diplomas = process_of_data(info)
        if len(list_diplomas) > 0:
           list_of_all_diplomas += list_diplomas
           list_of_fullname += [full_name] * len(list_diplomas)

    # refresh page
    print(i,"Прогресс", (i+1)/count_of_rows * 100, "%")
    driver.refresh()

# Close webpage
driver.close()

# 5 - write data

# create Dictionary
resultDict = {'Full name':list_of_fullname,'Diplomas':list_of_all_diplomas}
# Write to DataFrame
df = pd.DataFrame(resultDict)
# Write in excel file
writer = pd.ExcelWriter('output.xlsx')
df.to_excel(writer)
writer.save()
