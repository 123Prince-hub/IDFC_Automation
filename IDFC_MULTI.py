import os
import threading
import xlwings as xw
from time import sleep
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager

def task1():
    ws = xw.Book(r'card_detail.xlsx').sheets("data1")
    rows = ws.range("A2").expand().options(numbers=int).value

    account = rows[0][0]
    driver = webdriver.Chrome(ChromeDriverManager().install())
    driver.implicitly_wait(60)
    driver.maximize_window()
    driver.get('https://reporting.idfcfirstbank.com/QuickPay/QPInfo_Customer.aspx')
    account = driver.find_element_by_xpath("//a[contains(text(), 'Account Number')]//following::input").send_keys(account)
    sleep(60)
    
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
                driver.get('https://reporting.idfcfirstbank.com/QuickPay/QpLoanDetails.aspx')
                
                Pay_Now = driver.find_element_by_xpath('//input[@type="submit"]').click()

                amount = driver.find_element_by_xpath('//p[contains(text(), "Enter the amount you want to pay")]//following::input').send_keys(row[1])
                Payment_option = driver.find_element_by_xpath('/html/body/form/div[3]/div[2]/div/div/div[3]/div[5]/div[2]/div/label/span/img').click()
                button = driver.find_element_by_xpath('//button[contains(text(),"Make Payment")]').click() 
                checkbox = driver.find_element_by_xpath('//input[@type="checkbox"]').click()
                button2 = driver.find_element_by_xpath('//*[contains(text(),"Confirm Payment")]').click() 

                payment_method = driver.find_element_by_xpath('//span[contains(text(), "Debit Card")]').click() 
                card_number = driver.find_element_by_xpath('//span[contains(text(), "Card Number")]//following::input').send_keys(row[2]) 
                exp_mm = driver.find_element_by_xpath('//span[contains(text(), "Card Expiry Date")]//following::input').send_keys(mm) 
                exp_yy = driver.find_element_by_xpath('//span[contains(text(), "Card Expiry Date")]//following::input[2]').send_
                keys(yy) 
                cvv = driver.find_element_by_xpath('//span[contains(text(), "CVV")]//following::input').send_keys(row[4]) 
                button = driver.find_element_by_xpath('//span[contains(text(), "PAY")]').click()

                authentication = driver.find_element_by_xpath('//*[contains(text(),"ATM PIN")]').click()            
                exp_date2 = driver.execute_script("document.getElementById('expDate').value= '"+expiry+"'") 
                Pin_Number = driver.find_element_by_xpath('//b[contains(text(), "Pin Number")]//following::input').send_keys(row[5]) 
                button = driver.find_element_by_xpath('//*[@id="submitButtonIdForPin"]').click()

                # sleep(40)
                Transaction_Reference_No = driver.find_element_by_xpath('//div[contains(text(), "Transaction Reference No")]//following::span').text
                Transaction_status = driver.find_element_by_xpath('//div[contains(text(), "Transaction Status")]//following::span').text
                if "Successful" in Transaction_status:
                    ws.range("G"+str(num)).value = Transaction_Reference_No
                    ws.range("H"+str(num)).value = "Success"
                else:
                    ws.range("G"+str(num)).value = "Transaction_No Not Available"
                    ws.range("H"+str(num)).value = "Pending"

            except:
                ws.range("G"+str(num)).value = "NA"
                ws.range("H"+str(num)).value = "Server Error"

        num += 1 
    driver.close()










def task2():
    ws = xw.Book(r'card_detail.xlsx').sheets("data2")
    rows = ws.range("A2").expand().options(numbers=int).value

    account = rows[0][0]
    driver = webdriver.Chrome(ChromeDriverManager().install())
    driver.implicitly_wait(60)
    driver.maximize_window()
    driver.get('https://reporting.idfcfirstbank.com/QuickPay/QPInfo_Customer.aspx')
    account = driver.find_element_by_xpath("//a[contains(text(), 'Account Number')]//following::input").send_keys(account)
    sleep(60)
    
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
                driver.get('https://reporting.idfcfirstbank.com/QuickPay/QpLoanDetails.aspx')
                
                Pay_Now = driver.find_element_by_xpath('//input[@type="submit"]').click()

                amount = driver.find_element_by_xpath('//p[contains(text(), "Enter the amount you want to pay")]//following::input').send_keys(row[1])
                Payment_option = driver.find_element_by_xpath('/html/body/form/div[3]/div[2]/div/div/div[3]/div[5]/div[2]/div/label/span/img').click()
                button = driver.find_element_by_xpath('//button[contains(text(),"Make Payment")]').click() 
                checkbox = driver.find_element_by_xpath('//input[@type="checkbox"]').click()
                button2 = driver.find_element_by_xpath('//*[contains(text(),"Confirm Payment")]').click() 

                payment_method = driver.find_element_by_xpath('//span[contains(text(), "Debit Card")]').click() 
                card_number = driver.find_element_by_xpath('//span[contains(text(), "Card Number")]//following::input').send_keys(row[2]) 
                exp_mm = driver.find_element_by_xpath('//span[contains(text(), "Card Expiry Date")]//following::input').send_keys(mm) 
                exp_yy = driver.find_element_by_xpath('//span[contains(text(), "Card Expiry Date")]//following::input[2]').send_
                keys(yy) 
                cvv = driver.find_element_by_xpath('//span[contains(text(), "CVV")]//following::input').send_keys(row[4]) 
                button = driver.find_element_by_xpath('//span[contains(text(), "PAY")]').click()

                authentication = driver.find_element_by_xpath('//*[contains(text(),"ATM PIN")]').click()            
                exp_date2 = driver.execute_script("document.getElementById('expDate').value= '"+expiry+"'") 
                Pin_Number = driver.find_element_by_xpath('//b[contains(text(), "Pin Number")]//following::input').send_keys(row[5]) 
                button = driver.find_element_by_xpath('//*[@id="submitButtonIdForPin"]').click()

                # sleep(40)
                Transaction_Reference_No = driver.find_element_by_xpath('//div[contains(text(), "Transaction Reference No")]//following::span').text
                Transaction_status = driver.find_element_by_xpath('//div[contains(text(), "Transaction Status")]//following::span').text
                if "Successful" in Transaction_status:
                    ws.range("G"+str(num)).value = Transaction_Reference_No
                    ws.range("H"+str(num)).value = "Success"
                else:
                    ws.range("G"+str(num)).value = "Transaction_No Not Available"
                    ws.range("H"+str(num)).value = "Pending"

            except:
                ws.range("G"+str(num)).value = "NA"
                ws.range("H"+str(num)).value = "Server Error"

        num += 1 
    driver.close()








def task3():
    ws = xw.Book(r'card_detail.xlsx').sheets("data3")
    rows = ws.range("A2").expand().options(numbers=int).value

    account = rows[0][0]
    driver = webdriver.Chrome(ChromeDriverManager().install())
    driver.implicitly_wait(60)
    driver.maximize_window()
    driver.get('https://reporting.idfcfirstbank.com/QuickPay/QPInfo_Customer.aspx')
    account = driver.find_element_by_xpath("//a[contains(text(), 'Account Number')]//following::input").send_keys(account)
    sleep(60)
    
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
                driver.get('https://reporting.idfcfirstbank.com/QuickPay/QpLoanDetails.aspx')
                
                Pay_Now = driver.find_element_by_xpath('//input[@type="submit"]').click()

                amount = driver.find_element_by_xpath('//p[contains(text(), "Enter the amount you want to pay")]//following::input').send_keys(row[1])
                Payment_option = driver.find_element_by_xpath('/html/body/form/div[3]/div[2]/div/div/div[3]/div[5]/div[2]/div/label/span/img').click()
                button = driver.find_element_by_xpath('//button[contains(text(),"Make Payment")]').click() 
                checkbox = driver.find_element_by_xpath('//input[@type="checkbox"]').click()
                button2 = driver.find_element_by_xpath('//*[contains(text(),"Confirm Payment")]').click() 

                payment_method = driver.find_element_by_xpath('//span[contains(text(), "Debit Card")]').click() 
                card_number = driver.find_element_by_xpath('//span[contains(text(), "Card Number")]//following::input').send_keys(row[2]) 
                exp_mm = driver.find_element_by_xpath('//span[contains(text(), "Card Expiry Date")]//following::input').send_keys(mm) 
                exp_yy = driver.find_element_by_xpath('//span[contains(text(), "Card Expiry Date")]//following::input[2]').send_
                keys(yy) 
                cvv = driver.find_element_by_xpath('//span[contains(text(), "CVV")]//following::input').send_keys(row[4]) 
                button = driver.find_element_by_xpath('//span[contains(text(), "PAY")]').click()

                authentication = driver.find_element_by_xpath('//*[contains(text(),"ATM PIN")]').click()            
                exp_date2 = driver.execute_script("document.getElementById('expDate').value= '"+expiry+"'") 
                Pin_Number = driver.find_element_by_xpath('//b[contains(text(), "Pin Number")]//following::input').send_keys(row[5]) 
                button = driver.find_element_by_xpath('//*[@id="submitButtonIdForPin"]').click()

                # sleep(40)
                Transaction_Reference_No = driver.find_element_by_xpath('//div[contains(text(), "Transaction Reference No")]//following::span').text
                Transaction_status = driver.find_element_by_xpath('//div[contains(text(), "Transaction Status")]//following::span').text
                if "Successful" in Transaction_status:
                    ws.range("G"+str(num)).value = Transaction_Reference_No
                    ws.range("H"+str(num)).value = "Success"
                else:
                    ws.range("G"+str(num)).value = "Transaction_No Not Available"
                    ws.range("H"+str(num)).value = "Pending"

            except:
                ws.range("G"+str(num)).value = "NA"
                ws.range("H"+str(num)).value = "Server Error"

        num += 1 
    driver.close()











def task4():
    ws = xw.Book(r'card_detail.xlsx').sheets("data4")
    rows = ws.range("A2").expand().options(numbers=int).value

    account = rows[0][0]
    driver = webdriver.Chrome(ChromeDriverManager().install())
    driver.implicitly_wait(60)
    driver.maximize_window()
    driver.get('https://reporting.idfcfirstbank.com/QuickPay/QPInfo_Customer.aspx')
    account = driver.find_element_by_xpath("//a[contains(text(), 'Account Number')]//following::input").send_keys(account)
    sleep(60)
    
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
                driver.get('https://reporting.idfcfirstbank.com/QuickPay/QpLoanDetails.aspx')
                
                Pay_Now = driver.find_element_by_xpath('//input[@type="submit"]').click()

                amount = driver.find_element_by_xpath('//p[contains(text(), "Enter the amount you want to pay")]//following::input').send_keys(row[1])
                Payment_option = driver.find_element_by_xpath('/html/body/form/div[3]/div[2]/div/div/div[3]/div[5]/div[2]/div/label/span/img').click()
                button = driver.find_element_by_xpath('//button[contains(text(),"Make Payment")]').click() 
                checkbox = driver.find_element_by_xpath('//input[@type="checkbox"]').click()
                button2 = driver.find_element_by_xpath('//*[contains(text(),"Confirm Payment")]').click() 

                payment_method = driver.find_element_by_xpath('//span[contains(text(), "Debit Card")]').click() 
                card_number = driver.find_element_by_xpath('//span[contains(text(), "Card Number")]//following::input').send_keys(row[2]) 
                exp_mm = driver.find_element_by_xpath('//span[contains(text(), "Card Expiry Date")]//following::input').send_keys(mm) 
                exp_yy = driver.find_element_by_xpath('//span[contains(text(), "Card Expiry Date")]//following::input[2]').send_
                keys(yy) 
                cvv = driver.find_element_by_xpath('//span[contains(text(), "CVV")]//following::input').send_keys(row[4]) 
                button = driver.find_element_by_xpath('//span[contains(text(), "PAY")]').click()

                authentication = driver.find_element_by_xpath('//*[contains(text(),"ATM PIN")]').click()            
                exp_date2 = driver.execute_script("document.getElementById('expDate').value= '"+expiry+"'") 
                Pin_Number = driver.find_element_by_xpath('//b[contains(text(), "Pin Number")]//following::input').send_keys(row[5]) 
                button = driver.find_element_by_xpath('//*[@id="submitButtonIdForPin"]').click()

                # sleep(40)
                Transaction_Reference_No = driver.find_element_by_xpath('//div[contains(text(), "Transaction Reference No")]//following::span').text
                Transaction_status = driver.find_element_by_xpath('//div[contains(text(), "Transaction Status")]//following::span').text
                if "Successful" in Transaction_status:
                    ws.range("G"+str(num)).value = Transaction_Reference_No
                    ws.range("H"+str(num)).value = "Success"
                else:
                    ws.range("G"+str(num)).value = "Transaction_No Not Available"
                    ws.range("H"+str(num)).value = "Pending"

            except:
                ws.range("G"+str(num)).value = "NA"
                ws.range("H"+str(num)).value = "Server Error"

        num += 1 
    driver.close()










def task5():
    ws = xw.Book(r'card_detail.xlsx').sheets("data5")
    rows = ws.range("A2").expand().options(numbers=int).value

    account = rows[0][0]
    driver = webdriver.Chrome(ChromeDriverManager().install())
    driver.implicitly_wait(60)
    driver.maximize_window()
    driver.get('https://reporting.idfcfirstbank.com/QuickPay/QPInfo_Customer.aspx')
    account = driver.find_element_by_xpath("//a[contains(text(), 'Account Number')]//following::input").send_keys(account)
    sleep(60)
    
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
                driver.get('https://reporting.idfcfirstbank.com/QuickPay/QpLoanDetails.aspx')
                
                Pay_Now = driver.find_element_by_xpath('//input[@type="submit"]').click()

                amount = driver.find_element_by_xpath('//p[contains(text(), "Enter the amount you want to pay")]//following::input').send_keys(row[1])
                Payment_option = driver.find_element_by_xpath('/html/body/form/div[3]/div[2]/div/div/div[3]/div[5]/div[2]/div/label/span/img').click()
                button = driver.find_element_by_xpath('//button[contains(text(),"Make Payment")]').click() 
                checkbox = driver.find_element_by_xpath('//input[@type="checkbox"]').click()
                button2 = driver.find_element_by_xpath('//*[contains(text(),"Confirm Payment")]').click() 

                payment_method = driver.find_element_by_xpath('//span[contains(text(), "Debit Card")]').click() 
                card_number = driver.find_element_by_xpath('//span[contains(text(), "Card Number")]//following::input').send_keys(row[2]) 
                exp_mm = driver.find_element_by_xpath('//span[contains(text(), "Card Expiry Date")]//following::input').send_keys(mm) 
                exp_yy = driver.find_element_by_xpath('//span[contains(text(), "Card Expiry Date")]//following::input[2]').send_
                keys(yy) 
                cvv = driver.find_element_by_xpath('//span[contains(text(), "CVV")]//following::input').send_keys(row[4]) 
                button = driver.find_element_by_xpath('//span[contains(text(), "PAY")]').click()

                authentication = driver.find_element_by_xpath('//*[contains(text(),"ATM PIN")]').click()            
                exp_date2 = driver.execute_script("document.getElementById('expDate').value= '"+expiry+"'") 
                Pin_Number = driver.find_element_by_xpath('//b[contains(text(), "Pin Number")]//following::input').send_keys(row[5]) 
                button = driver.find_element_by_xpath('//*[@id="submitButtonIdForPin"]').click()

                # sleep(40)
                Transaction_Reference_No = driver.find_element_by_xpath('//div[contains(text(), "Transaction Reference No")]//following::span').text
                Transaction_status = driver.find_element_by_xpath('//div[contains(text(), "Transaction Status")]//following::span').text
                if "Successful" in Transaction_status:
                    ws.range("G"+str(num)).value = Transaction_Reference_No
                    ws.range("H"+str(num)).value = "Success"
                else:
                    ws.range("G"+str(num)).value = "Transaction_No Not Available"
                    ws.range("H"+str(num)).value = "Pending"

            except:
                ws.range("G"+str(num)).value = "NA"
                ws.range("H"+str(num)).value = "Server Error"

        num += 1 
    driver.close()








if __name__ == "__main__":
	# creating threads
	t1 = threading.Thread(target=task1, name='t1')
	t2 = threading.Thread(target=task2, name='t2')
	t3 = threading.Thread(target=task3, name='t3')
	t4 = threading.Thread(target=task4, name='t4')
	t5 = threading.Thread(target=task5, name='t5')

	# starting threads
	t1.start()
	t2.start()
	t3.start()
	t4.start()
	t5.start()

	# wait until all threads finish
	t1.join()
	t2.join()
	t3.join()
	t4.join()
	t5.join()
