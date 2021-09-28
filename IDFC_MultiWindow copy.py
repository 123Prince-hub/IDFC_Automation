from multiprocessing.pool import ThreadPool as Pool
from selenium import webdriver
import xlwings as xw
import time

pool_size = 2

ws = xw.Book(r'card_detail.xlsx').sheets("data")
rows = ws.range("A2").expand().options(numbers=int).value

num = 2

def automation(row, num):   
    if (col != "Success"):
        try:
            driver = webdriver.Chrome(executable_path="C:\Program Files (x86)\chromedriver.exe")
            driver.maximize_window()
            url = driver.get('https://reporting.idfcfirstbank.com/QuickPay/QPInfo_Customer.aspx')
            account = driver.find_element_by_xpath("//a[contains(text(), 'Account Number')]//following::input").send_keys("49121607")
            tim = time.sleep(20)

            if num == 3:
                del url, account, tim

            driver.get('https://reporting.idfcfirstbank.com/QuickPay/QpLoanDetails.aspx')
            Pay_Now = driver.find_element_by_xpath('//input[@type="submit"]').click()

            time.sleep(1)
            amount = driver.find_element_by_xpath('//p[contains(text(), "Enter the amount you want to pay")]//following::input').send_keys(row[1])
            Payment_option = driver.find_element_by_xpath('/html/body/form/div[3]/div[2]/div/div/div[3]/div[5]/div[2]/div/label/span/img').click()
            button = driver.find_element_by_xpath('//button[contains(text(),"Make Payment")]').click() 

            time.sleep(1)
            checkbox = driver.find_element_by_xpath('//input[@type="checkbox"]').click()
            button2 = driver.find_element_by_xpath('//*[contains(text(),"Confirm Payment")]').click() 

            time.sleep(1)
            payment_method = driver.find_element_by_xpath('//span[contains(text(), "Debit Card")]').click() 
            card_number = driver.find_element_by_xpath('//span[contains(text(), "Card Number")]//following::input').send_keys(row[2]) 
            exp_mm = driver.find_element_by_xpath('//span[contains(text(), "Card Expiry Date")]//following::input').send_keys(mm) 
            exp_yy = driver.find_element_by_xpath('//span[contains(text(), "Card Expiry Date")]//following::input[2]').send_keys(yy) 
            cvv = driver.find_element_by_xpath('//span[contains(text(), "CVV")]//following::input').send_keys(row[4]) 
            button = driver.find_element_by_xpath('//span[contains(text(), "PAY")]').click()

            time.sleep(15)
            authentication = driver.find_element_by_xpath('//*[contains(text(),"ATM PIN")]').click()
            exp_date2 = driver.execute_script("document.getElementById('expDate').value= '"+expiry+"'") 
            Pin_Number = driver.find_element_by_xpath('//b[contains(text(), "Pin Number")]//following::input').send_keys(row[5]) 
            button = driver.find_element_by_xpath('//*[@id="submitButtonIdForPin"]').click()

            time.sleep(30)
            Transaction_Reference_No = driver.find_element_by_xpath('//div[contains(text(), "Transaction Reference No")]//following::span').text
            Transaction_status = driver.find_element_by_xpath('//div[contains(text(), "Transaction Status")]//following::span').text
            if "Successful" in Transaction_status:
                ws.range("G"+str(num)).value = "Transaction_Reference_No"
                ws.range("H"+str(num)).value = "Success"
            else:
                ws.range("G"+str(num)).value = "Transaction_Reference_No"
                ws.range("H"+str(num)).value = "Pending"

        except:
            ws.range("G"+str(num)).value = "NA"
            ws.range("H"+str(num)).value = "Server Error"


pool = Pool(pool_size)

for row in rows:
    exp_date = row[3]
    mm = exp_date[0:2]
    yy = exp_date[5:]
    expiry = exp_date.replace("/", "-")
    expiry = str(expiry) 
    col = ws.range("H"+str(num)).value

    pool.apply_async(automation, (row, num))
    num += 1

pool.close()
pool.join()
