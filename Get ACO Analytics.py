from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from prettytable import PrettyTable
from prettytable import MSWORD_FRIENDLY
import smtplib 
import datetime as DT
from email.message import EmailMessage

""" Scrape_analytics_data() logs into the analytics portal of ACO to quickly gather data to find if there are
any obvious problem extensions
Pre: 
url - The url for the specific portion of the analytics portal 
name - log-in email
pw - log-in password """
def scrape_analytics_data(url, name, pw):
    data = [1]
    try:
        Option = Options()
        Option.add_argument('--enable-javascript')
        driver = webdriver.Chrome(options=Option)

        #Go to ACO analytics
        driver.get(url)

        #WebDriverWait.until() is a dynamic timer that forces the code to wait until the relevant data piece is shown
        #Here, we search for the 'Authorize' button
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '/html/body/div/div/div/div/div[1]/div[2]/button')))
        #driver.find_element() searches the HTML for a specific data piece to manipulate
        #Click the 'Authorize' button
        driver.find_element(By.XPATH, '/html/body/div/div/div/div/div[1]/div[2]/button').click()
        
        #Finding and inputting the supplied username from the function
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.NAME, 'credential')))
        log = driver.find_element(By.NAME, 'credential')
        log.send_keys(name)
        
        #Finding and clicking the 'Next' button that will reveal the password box
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div[5]/div[2]/div[2]/div/form/div/div/div/div[2]/div[2]/button')))
        driver.find_element(By.XPATH, '/html/body/div[2]/div[5]/div[2]/div[2]/div/form/div/div/div/div[2]/div[2]/button').click()
        
        #Finding and inputting the supplied password from the function
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.NAME, "Password")))
        passWORD = driver.find_element(By.NAME, "Password")
        passWORD.send_keys(pw)
        
        #Finding and clicking the 'Sign In' button to finish the sign in process
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div[5]/div[2]/div[2]/div/form/div/div[1]/div/div/div[3]/div/div/button[2]')))
        driver.find_element(By.XPATH, '/html/body/div[2]/div[5]/div[2]/div[2]/div/form/div/div[1]/div/div/div[3]/div/div/button[2]').click()

        #Finding and gathering the 'Active Users' data
        ##The 'Active Users' data is the top 10 extensions with the most calls
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.CLASS_NAME, "row--NDt7h")))
        data = driver.find_elements(By.CLASS_NAME, "row--NDt7h")
        
        #Data set containing the found data
        data_list = []

        #Adding data to a blank list for easy manipulation 
        for p in range(len(data)):
            if p == 0:
                continue
            data_list.append(data[p].text)
            data_list[p-1] = data_list[p-1].replace('\n', ' ')

        #Data set containing the separate instances of each row of data
        data_list_smol = []
        for p in range(len(data)):
            if p == 0:
                continue
            data_list_smol.append(data[p].text)
            data_list_smol[p-1] = data_list_smol[p-1].split('\n')

        #Output data is a neat table containing important information  
        Tabular_format = PrettyTable(['Name', 'Extension', 'Calls', 'Call Percent', 'Minutes', 'Minutes Percent'])
        for p in range(len(data_list)):
            Tabular_format.add_row([data_list_smol[p][0], data_list_smol[p][1], data_list_smol[p][2], data_list_smol[p][3], data_list_smol[p][4], data_list_smol[p][5]])
        print(Tabular_format)

        #Converting the pretty table so the email functions are able to understand it
        data_str = Tabular_format.get_string()
        
        #Write the pretty table into an output file that will be sent along with the email for easy reading
        output = open("C:\\Users\\ytc3512\\OneDrive - HMESRETL\\Documents\\Python Analytics\\output.txt", "w")
        output.write("The following are the top 10 most used extensions over the last 7 days:\n")
        output.write(data_str)
        output.close()

        #Sends the email containing all information found
        mail()

        return data
    except Exception as e:
        print(f"Error: {e}")
        return ""  

#Function to mail out the results
def mail():
    #Setting up the email server that will send
    smtp = smtplib.SMTP()
    smtp.connect('mta.yrcw.com')
    from_addr = "{fromEmail}"
    to_addr = "{toEmail}"
   
    #Preparing the date for the subject line of the email
    day = DT.date.today()
    day = day - DT.timedelta(days=1)
    one_week = day - DT.timedelta(days=6)
    day_text = day.strftime("%m/%d/%y")
    week_text = one_week.strftime("%m/%d/%y")
    
    #Preparing the email that will be sent
    string = EmailMessage()
    string['Subject'] = f"ACO Analytics {week_text}-{day_text}"
    with open("C:\\Users\\ytc3512\\OneDrive - HMESRETL\\Documents\\Python Analytics\\output.txt", "rb") as f:
        string.add_attachment(
            f.read(),
            filename="output.txt",
            maintype="application",
            subtype="txt"
        )

    #Sends the prepared email to the recipient(s)
    smtp.sendmail(from_addr, to_addr, string.as_string())
    smtp.quit()

if __name__ == "__main__":
    scrape_analytics_data("https://analytics.cloudoffice.avaya.com/adoption-and-usage/phone", "{loginEmail}", "{Password}")
