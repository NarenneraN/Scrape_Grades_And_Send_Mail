


from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.wait import WebDriverWait
import pywhatkit
import pandas as pd
from tabulate import tabulate
website = 'https://aumscb.amrita.edu/cas/login?service=https%3A%2F%2Faumscb.amrita.edu%2Faums%2FJsp%2FCore_Common%2Findex.jsp%3Ftask%3Doff'
path = 'E:\MY_WORKS\chromedriver'

service = Service(executable_path=path)
driver = webdriver.Chrome(service=service)
driver.get(website)


# 1. Filling the username and password

# Username Xpath - //input[@id="username"]
# Password Xpath - //input[@id="password"]
# Submit Button - //input[@class="btn-submit"]
driver.implicitly_wait(5)
username_box = driver.find_element(by="xpath",value='//input[@id="username"]')
password_box = driver.find_element(by="xpath",value='//input[@id="password"]')
submit_btn = driver.find_element(by="xpath",value='//input[@class="btn-submit"]')
username_box.send_keys('CB.EN.U4CSE20616')
password_box.send_keys('droideronline')
submit_btn.click()
# WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.ID,"maincontentframe")))
# WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.ID,"Iframe1")))
# WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.ID,"sakaiframeId")))

# 2.
driver.find_element(by="xpath",value='//div[@class="sidebar-toggler"]').click()
# //ul[@id="navbar-menu"]/li[5]/a
driver.find_element(by="xpath",value='//ul[@id="navbar-menu"]/li[5]/a').click()
# //li[@data-url="../StudentGrade/StudentPerformanceWithSurvey.jsp?action=UMS-EVAL_STUDPERFORMSURVEY_INIT_SCREEN&isMenu=true"]/a
driver.find_element(by="xpath",value='//li[@data-url="../StudentGrade/StudentPerformanceWithSurvey.jsp?action=UMS-EVAL_STUDPERFORMSURVEY_INIT_SCREEN&isMenu=true"]/a').click()
WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.ID,"maincontentframe")))
driver.find_element(by="xpath",value='//select[@name="htmlPageTopContainer_selectStep"]/option[text()="3"]').click()
# //div[@id="Grades"]/table/tbody/tr
table_len = driver.find_elements(by="xpath",value='//div[@id="Grades"]/table/tbody/tr')
print(len(table_len))
if(len(table_len)>2):
    msg='You have recieved your 4th semester grades ! Check out'
    rows = driver.find_elements(by="xpath",value="//div[@id='Grades']/table/tbody/tr")
    # cols = driver.find_elements(by="xpath",value="//div[@id='Grades']/table/tbody/tr[1]/td")
    data=[]
    for i in rows:
        sub_data=[]
        cols=i.find_elements(by="xpath",value="./td")
        for j in cols:
            # s = '{0: <50}'.format(j.text)
            sub_data.append(j.text)
            # print(s,end=' | ')
        # print(' ')
        data.append(sub_data)
    header=data.pop(0)
    print(header)
    print(data)
    from datetime import datetime
    now = datetime.now()
    m_d_y = now.strftime("%m%d%y")
    dataFrame = pd.DataFrame(data = data, columns = header)
    print(dataFrame)
    name='Result'+str(m_d_y)+'.csv'
    dataFrame.to_csv(name)
    ans=tabulate(data, headers=header)
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.base import MIMEBase
    from email import encoders

    fromaddr = "narendhiran.pugazhendhi@gmail.com"
    toaddr = "rajmiracle2023@gmail.com"

    # instance of MIMEMultipart
    msg = MIMEMultipart()

    # storing the senders email address
    msg['From'] = fromaddr

    # storing the receivers email address
    msg['To'] = toaddr

    # storing the subject
    msg['Subject'] = "Your grades"

    # string to store the body of the mail
    body = "Results of the semester in csv format"

    # attach the body with the msg instance
    msg.attach(MIMEText(body, 'plain'))

    # open the file to be sent
    filename = name
    attachment = open(name, "rb")

    # instance of MIMEBase and named as p
    p = MIMEBase('application', 'octet-stream')

    # To change the payload into encoded form
    p.set_payload((attachment).read())

    # encode into base64
    encoders.encode_base64(p)

    p.add_header('Content-Disposition', "attachment; filename= %s" % filename)

    # attach the instance 'p' to instance 'msg'
    msg.attach(p)

    # creates SMTP session
    s = smtplib.SMTP('smtp.gmail.com', 587)

    # start TLS for security
    s.starttls()

    # Authentication
    s.login(fromaddr, "rfdjiokatgjhomtn")

    # Converts the Multipart msg into a string
    text = msg.as_string()

    # sending the mail
    s.sendmail(fromaddr, toaddr, text)

    # terminating the session
    s.quit()
else:
    msg='4th semester results --- Not yet published'
# driver.quit()
