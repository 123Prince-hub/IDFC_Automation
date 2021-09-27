from selenium import webdriver
import xlwings as xw
from time import sleep
import os, easyocr

ws = xw.Book(r'card_detail.xlsx').sheets("data")
rows = ws.range("A2").expand().options(numbers=int).value
driver = webdriver.Chrome(executable_path="C:\Program Files (x86)\chromedriver.exe")
driver.maximize_window()
driver.get('https://reporting.idfcfirstbank.com/QuickPay/QPInfo_Customer.aspx')
sleep(10)
 
num = 2
for row in rows:

    exp_date = row[3]
    mm = exp_date[0:2]
    yy = exp_date[5:]
    expiry = exp_date.replace("/", "-")
    expiry = str(expiry) 

    col = ws.range("H"+str(num)).value
    
    if (col != "Success"):
        try:
            driver.get('https://reporting.idfcfirstbank.com/QuickPay/QpinfoNew.aspx')

            account = driver.find_element_by_xpath("//a[contains(text(), 'Account Number')]//following::input").send_keys(row[0])

            image_link = driver.find_element_by_xpath('//*[@id="imgcapcha"]').get_attribute('src')
            driver.execute_script("window.open('');")
            sleep(1)

            driver.switch_to.window(driver.window_handles[1])
            driver.get(image_link)
            driver.find_element_by_xpath('/html/body/img').screenshot('img.png')

            reader = easyocr.Reader(['en'], gpu=False) 
            txt = reader.readtext('img.png',detail=0,paragraph=True)
            a = " "
            a = a.join(txt)
            driver.close()
            sleep(1)
            driver.switch_to.window(driver.window_handles[0])
            captchaInput = driver.find_element_by_id('txtCaptcha').send_keys(a)
            button = driver.find_element_by_xpath('//input[@type="submit"]').click()

            try:
                sleep(1)
                error = driver.find_element_by_xpath('//span[contains(text(), "Invalid Captcha...Please try again.")] | //span[contains(text(), "Please Enter Captcha")] ')
                sleep(1)
                check=error.text
                i = 0
                while ( (check) ):
                    driver.find_element_by_id("txtCaptcha").clear()
                    image_link = driver.find_element_by_xpath('//*[@id="imgcapcha"]').get_attribute('src')
                    driver.execute_script("window.open('');")
                    sleep(1)

                    driver.switch_to.window(driver.window_handles[1])
                    driver.get(image_link)
                    driver.find_element_by_xpath('/html/body/img').screenshot('img.png')

                    reader = easyocr.Reader(['en'], gpu=False) 
                    txt = reader.readtext('img.png',detail=0,paragraph=True)
                    a = " "
                    a = a.join(txt)
                    driver.close()
                    sleep(1)
                    driver.switch_to.window(driver.window_handles[0])
                    driver.execute_script("document.querySelector('#lblmessage').replaceWith('')")
                    captchaInput = driver.find_element_by_id('txtCaptcha').send_keys(a)
                    button = driver.find_element_by_xpath('//input[@type="submit"]').click()
                    sleep(1)
                    i += 1  
                
            except:
                os.remove('img.png')

                sleep(2)
                Pay_Now = driver.find_element_by_xpath('//input[@type="submit"]').click()

                sleep(2)
                amount = driver.find_element_by_xpath('//p[contains(text(), "Enter the amount you want to pay")]//following::input').send_keys(row[1])
                Payment_option = driver.find_element_by_xpath('/html/body/form/div[3]/div[2]/div/div/div[3]/div[5]/div[2]/div/label/span/img').click()
                sleep(1)
                # button = driver.find_element_by_xpath('//*[contains(text(),"Make Payment")]').click() 
                button = driver.find_element_by_xpath('//*[@id="btnProceed"]').click() 

                sleep(2)
                checkbox = driver.find_element_by_xpath('//input[@type="checkbox"]').click()
                sleep(2)
                button2 = driver.find_element_by_xpath('//*[contains(text(),"Confirm Payment")]').click() 

                sleep(1)
                # payment_method = driver.find_element_by_xpath('//*[@id="app"]/main/div[3]/div[4]/section[1]/section/div[1]/div/label/input').click() 
                sleep(1)
                payment_method = driver.find_element_by_xpath('//span[contains(text(), "Debit Card")]').click() 
                sleep(1)
                card_number = driver.find_element_by_xpath('//span[contains(text(), "Card Number")]//following::input').send_keys(row[2]) 
                sleep(1)
                exp_mm = driver.find_element_by_xpath('//span[contains(text(), "Card Expiry Date")]//following::input').send_keys(mm) 
                sleep(1)
                exp_yy = driver.find_element_by_xpath('//span[contains(text(), "Card Expiry Date")]//following::input[2]').send_keys(yy) 
                sleep(1)
                cvv = driver.find_element_by_xpath('//span[contains(text(), "CVV")]//following::input').send_keys(row[4]) 
                sleep(1)
                button = driver.find_element_by_xpath('//span[contains(text(), "PAY")]').click()

                sleep(15)
                authentication = driver.find_element_by_xpath('//*[contains(text(),"ATM PIN")]').click()
                # authentication = driver.find_element_by_xpath('//*[@id="tab-B-label"]/span').click()
                sleep(2)
                
                exp_date2 = driver.execute_script("document.getElementById('expDate').value= '"+expiry+"'") 
                Pin_Number = driver.find_element_by_xpath('//b[contains(text(), "Pin Number")]//following::input').send_keys(row[5]) 
                button = driver.find_element_by_xpath('//*[@id="submitButtonIdForPin"]').click()

                sleep(30)
                Transaction_Reference_No = driver.find_element_by_xpath('//div[contains(text(), "Transaction Reference No")]//following::span').text
                Transaction_status = driver.find_element_by_xpath('//div[contains(text(), "Transaction Status")]//following::span').text
                if "Successful" in Transaction_status:
                    ws.range("G"+str(num)).value = Transaction_Reference_No
                    ws.range("H"+str(num)).value = "Success"
                else:
                    ws.range("G"+str(num)).value = Transaction_Reference_No
                    ws.range("H"+str(num)).value = "Pending"

        except:
            ws.range("G"+str(num)).value = "NA"
            ws.range("H"+str(num)).value = "Server Error"

    num += 1 


driver.close()