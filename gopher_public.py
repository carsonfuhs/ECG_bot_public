from datetime import datetime;

import os;
import requests;
import schedule;
import smtplib;
import time;
import win32com;
import win32com.client;

from selenium import webdriver;

from selenium.webdriver.common.keys import Keys;
from selenium.webdriver.common.by import By;
from selenium.webdriver.support import expected_conditions as EC;
from selenium.webdriver.support.ui import WebDriverWait;
from selenium.webdriver.chrome.options import Options;

from selenium.common.exceptions import NoSuchElementException;
from selenium.common.exceptions import TimeoutException;
from selenium.common.exceptions import ElementClickInterceptedException;
from selenium.common.exceptions import WebDriverException;





'''

Note: This is a public version of the code. I've taken a lot of things out,
such as credentials, URLs, filepaths, and anything else that might not be a good
idea to display publicly. This wouldn't work if I ran the code on my computer,
but I do have an non-public copy I use.

'''





'''
            Gopher: an email bot that fetches data

The purpose of Gopher is to get test data for business users who can see the data
in the reporting tool, but can’t access Sterling Integrator to pull the data
themselves.

Gopher will scan my emails for a filename, and if it finds it, it'll go pull
the file from PepsiCo's ecommerce system (IBM Sterling Integrator) and email it
back to the initial requestor.

'''





class gopher_email:

    # outbound email messages for receipts and failures
    successful_subject_receipt  =   '''Successful Receipt via Gopher'''
    successful_body_receipt     =   '''Gopher has received your data request and is currently fetching.'''
    successful_data_request     =   '''Requested Test Data from Gopher'''
    successful_data_body        =   '''Your requested data is attached.'''
    failure_subject             =   '''Failure fetching via Gopher''' 
    failure_body                =   '''Gopher has failed fetching one of your files.\n\nPlease contact ECG Support for more information.'''

    def __init__(self):

        # naming outlook API. Mostly standard notation as per win32com documentation.
        self.outlook                    = win32com.client.Dispatch("Outlook.Application")   
        self.gopher_inbox               = self.outlook.GetNamespace("MAPI").GetDefaultFolder(6).Folders[1]   
        self.gopher_outbox              = self.outlook.GetNamespace("MAPI").GetDefaultFolder(6).Folders[2]
        self.gopher_failbox             = self.outlook.GetNamespace("MAPI").GetDefaultFolder(6).Folders[2].Folders[0]
        self.gopher_email_count         = self.gopher_inbox.Items.Count
        self.most_recent_email_index    = self.gopher_inbox.Items.Count - 1 # highest number = most recent. Subtracting 1 makes it a 0 index.

    def get_most_recent_email_details(self):

        self.email_body                             = self.gopher_inbox.Items[self.most_recent_email_index].Body
        self.sender_email_address_outlook_object    = self.gopher_inbox.Items[self.most_recent_email_index].SenderEmailAddress
        
    def send_outbound_email(self, subject, body, paths = None):

        mail            = self.outlook.CreateItem(0)
        mail.To         = self.sender_email_address_outlook_object
        mail.Subject    = subject
        mail.Body       = body

        # this adds attachments if I need them. paths parameter should always be a list, even if there's only 1 path.
        if paths is not None:          
            for y in paths:
                mail.Attachments.Add(y)
                
        mail.Send()

    def move_email_from_inbox_to_different_folder(self, inbox, target_folder):

        # marks as "read" if it wasn't already
        if self.gopher_inbox.Items[self.most_recent_email_index].UnRead == True:
            self.gopher_inbox.Items[self.most_recent_email_index].UnRead = False
            
        self.gopher_inbox.Items[self.most_recent_email_index].Move(target_folder)

    def check_if_sender_is_an_authorized_sender(self):

        # list of authorized users for Gopher (changed for public display)
        list_of_authorized_users   =    [   '''firstname_lastname1@example.com''',
                                            '''firstname_lastname2@example.com''',
                                            '''firstname_lastname3@example.com'''] 

        # this gets the actual email address as a string (ex: John.Doe@Pepsico.com). Other one in earlier method is an outlook object.
        self.sender_email_address = self.gopher_inbox.Items[self.most_recent_email_index].Sender.GetExchangeUser().PrimarySmtpAddress

        # using list comprehension to lower all list entries for proper comparison
        list_of_authorized_users = [x.lower() for x in list_of_authorized_users]
        if ( self.sender_email_address.lower() ) in list_of_authorized_users:
            print('Successful Authentication!')
            return True
        
        else:
            return False

        


    
def get_filenames_from_body_of_email(email_body):

    raw_filenames = []
    offset = 0 
    sub = '/specific/filename/format/for/pepsico'   # the data that users are requesting has a very specific filename format.
    i = email_body.find(sub, offset)                # I've changed it to a generic version for public display.

    if i == -1:
        print('Filename not found')
        raise SystemExit(1)

    # finding the filename(s) within the body
    while i >= 0:
        raw_filenames.append(email_body[i:i+58]) # each filename is 58 characters long
        i = email_body.find(sub, i + 1)

    return raw_filenames




    
class file_validation:

    def __init__(self):

        self.validation_failure = False

    # notice how it accepts a singular filename and not the list of raw_filenames!
    # I'll loop the raw_filenames list in the main method before this function call.
    def split_strings(self, filename):

        # splitting strings to validate
        try:
            self.receive_date_1  = int(filename[24:32])
            self.receive_date_2  = int(filename[33:41])
            self.receive_time    = int(filename[42:47])
            
        except ValueError:
            self.validation_failure = True

    # also only accepts a singular filename
    def validate_file(self, filename):

        # validation sequence - checks to see if characters are in the right place.
        if len(filename) != 58:
            self.validation_failure = True
                                            
        elif filename.find('/', 31) != 32:
            self.validation_failure = True
            
        elif filename.find('_', 14) != 15:
            self.validation_failure = True
            
        elif filename.find('_', 40) != 41:
            self.validation_failure = True
            
        elif filename.find('_', 47) != 48:
            self.validation_failure = True
            
        elif self.receive_date_1 != self.receive_date_2:
            self.validation_failure = True
            
        elif self.receive_date_1 < 20000000:    # these validate if the dates are possible. This simply checks
            self.validation_failure = True      # if it's over the year 2000.
            
        elif self.receive_time > 240000:    # validation based on a 24 hour clock
            self.validation_failure = True
            
        elif self.receive_time < 000000:
            self.validation_failure = True
            
        else:
            print('Successful Validation!')



            

def get_user_password():

    user_password = input('Password: ')

    return user_password




        
# "Reading" the xpath is impossible, so I'll try to put descriptions by
# each xpath value. This will be helpful if it ever breaks by IBM changing the
# HTML code of Sterling Integrator. Hopefully I'll be able to change it in the
# future, as absolute xpath is very prone to breakage.





class sterling_integrator_navigation():

    def __init__(self):

        # starting chromedriver and going to SI. Webdriver has a 20 second wait time or it'll throw a timeout exception.    
        options = Options()
        options.headless = True
        self.driver = webdriver.Chrome(executable_path=r'C:\chromedriverFilepath', options=options)
        self.wait = WebDriverWait(self.driver, 20)

    def get_to_SI_home_page(self):

        self.driver.get('Pepsico IBM Sterling Integrator WebApp URL') # actual URL changed for public display

    def login_to_SI(self, driver, wait, user_gpid, user_password):

        # gets to the splash page and logs in
        wait.until(EC.element_to_be_clickable((By.NAME, 'submit')))
        driver.find_element_by_name('username').send_keys(user_gpid)
        driver.find_element_by_name('password').send_keys(user_password)
        driver.find_element_by_name('submit').click()

    def navigate_to_GET_ARCHIVE(self, driver, wait):

        # waits until the "Administration Menu" is loaded, then clicks through to
        # get to the "Manager" option, always waiting for the elements.
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, 'IframeConsolePanel')))
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, 'menu')))
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="mainframe"]/table/tbody/tr[2]/td/ul/li[1]/a')))#business process menu button
        driver.find_element_by_xpath('//*[@id="mainframe"]/table/tbody/tr[2]/td/ul/li[1]/a').click()   
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="mainframe"]/table/tbody/tr[2]/td/ul/li[1]/ul/li[1]/a')))#manager menu button
        driver.find_element_by_xpath('//*[@id="mainframe"]/table/tbody/tr[2]/td/ul/li[1]/ul/li[1]/a').click()

        # searches for GET_ARCHIVE by jumping through frames and to the search
        driver.switch_to.default_content()
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, 'IframeConsolePanel')))
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, 'basefrm')))
        wait.until(EC.element_to_be_clickable((By.ID,'autoCompleteBPNameSource')))
        driver.find_element_by_id('autoCompleteBPNameSource').send_keys('PepsiCo Archive File Process Name') # changed for public display
        wait.until(EC.visibility_of_element_located((By.NAME, 'Go2')))
        wait.until(EC.element_to_be_clickable((By.NAME,'Go2')))
        driver.find_element_by_name('Go2').click() 

        # switching back to main frame so I can switch to the other frames
        driver.switch_to.default_content()
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID,'IframeConsolePanel')))
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, 'basefrm')))
        wait.until(EC.element_to_be_clickable((By.NAME,'Exe1422')))# execution manager button to the left
        driver.find_element_by_name('Exe1422').click()

    def search_for_file_and_pull_data(self, driver, wait, raw_filename):

        # Identifying main window first. Necessary for repetitive function.
        sterling_integrator = driver.window_handles[0]
        driver.switch_to.window(sterling_integrator)
        driver.switch_to.default_content()
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID,'IframeConsolePanel')))
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, 'basefrm')))
        wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/table/tbody/tr[2]/td[2]/table[2]/tbody/tr[9]/td[8]/a/img')))#green execute octogon under ECG_GET_ARCHIVE_FILE
        driver.find_element_by_xpath('/html/body/table/tbody/tr[1]/td/table/tbody/tr[2]/td[2]/table[2]/tbody/tr[9]/td[8]/a/img').click()

        # SI will open a pop up after searching.
        ECG_GET_ARCHIVE_FILE_window = driver.window_handles[1]
        driver.switch_to.window(ECG_GET_ARCHIVE_FILE_window)

        # finding the search bar in the new window
        wait.until(EC.element_to_be_clickable((By.NAME, 'unixfile')))
        driver.find_element_by_name('unixfile').send_keys(raw_filename) 
        wait.until(EC.element_to_be_clickable((By.ID,'oGoAnchor')))
        driver.find_element_by_id('oGoAnchor').click()

        # waiting for the completed text to show up, then clicking on the document link
        wait.until(EC.text_to_be_present_in_element((By.XPATH,'/html/body/table[1]/tbody/tr/td/table/tbody/tr[7]/td[2]'), 'Completed')) #"Completed, Status: Success" label
        wait.until(EC.element_to_be_clickable((By.XPATH,'/html/body/table[1]/tbody/tr/td/table/tbody/tr[8]/td[2]/table/tbody/tr[6]/td[35]/a/img')))#"info" button under document
        driver.find_element_by_xpath('/html/body/table[1]/tbody/tr/td/table/tbody/tr[8]/td[2]/table/tbody/tr[6]/td[35]/a/img').click()

        # switching to another window, scraping the text
        target_file_window = driver.window_handles[2]
        driver.switch_to.window(target_file_window)
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME,'DocFrame')))

        # getting the actual data 
        wait.until(EC.visibility_of_element_located((By.XPATH,'/html/body/table/tbody/tr/td/table/tbody/tr[8]/td[2]/table/tbody/tr/td[2]/pre')))
        self.data = driver.find_element_by_xpath('/html/body/table/tbody/tr/td/table/tbody/tr[8]/td[2]/table/tbody/tr/td[2]/pre').text #actual document text





class data_txt_file():

    def __init__(self):

        self.path_list = []

    def write_data_to_txt_file(self, data, bp_number):

        # writing the data to a text file. For however many files there are, this changes the filepath
        # based on a unique identifier of the filename called a bp number.
        path = 'C:\ECG Projects\Gopher files\\' + str(bp_number) + '.txt'
        file_object = open(path, 'w')
        file_object.write(data)
        file_object.close()

        # adds filepath to list for later usage.
        self.path_list.append(path)





def check_if_webdriver_is_still_running(driver):

    try:
        driver.title
        return True

    except WebDriverException:
        return False





def main():

    # any hard coded variables I might need go here. If I were to develop out gopher more, I would
    # probably be pulling credentials from a secure file. I might revisit in the future.
    user_gpid       = ''

    # initializing class. This only gets gopher's emails when initialized.
    ge = gopher_email()

    while ge.gopher_email_count != 0:

        if ge.most_recent_email_index > -1: 
            ge.get_most_recent_email_details() 

            # checks if the sender is authorized via hard coded list
            if ge.check_if_sender_is_an_authorized_sender() == False:
                break

            # returns a list (raw_filenames) from the body of the email request
            raw_filenames = get_filenames_from_body_of_email(ge.email_body)

            fv = file_validation()

            # looping through a list to validate
            for x in raw_filenames:

                # if validation fails, it sends a failure email and breaks from the for loop
                fv.split_strings(x)
                if fv.validation_failure == True:
                    ge.send_outbound_email(ge.failure_subject, ge.failure_body)
                    break

                fv.validate_file(x)
                if fv.validation_failure == True:
                    ge.send_outbound_email(ge.failure_subject, ge.failure_body)
                    break

            if fv.validation_failure == False:      # if there are no validation errors

                # sends a successful receipt
                ge.send_outbound_email(ge.successful_subject_receipt, ge.successful_body_receipt)

                user_password = get_user_password()

                try:    # sterling integrator webapp navigation. Check classes and methods for more details.
                    si_nav = sterling_integrator_navigation()
                    
                    try:
                        si_nav.get_to_SI_home_page()

                        try:
                            si_nav.login_to_SI(si_nav.driver, si_nav.wait, user_gpid, user_password)

                            try:
                                si_nav.navigate_to_GET_ARCHIVE(si_nav.driver, si_nav.wait)

                                try:
                                    file = data_txt_file()
                                    
                                    # looping for as many files are in the list
                                    for y in raw_filenames:

                                        # pulls data, writes to a .txt file
                                        si_nav.search_for_file_and_pull_data(si_nav.driver, si_nav.wait, y)
                                        file.write_data_to_txt_file(si_nav.data, str(y[49:58]) )# this pulls the bp number from the filename 

                                    # successful ending sequence
                                    ge.send_outbound_email(ge.successful_data_request, ge.successful_data_body, file.path_list)
                                    ge.move_email_from_inbox_to_different_folder(ge.gopher_inbox, ge.gopher_outbox)

                                    # reinitializing the email class to check/update Gopher's emails
                                    ge = gopher_email()

                                    if check_if_webdriver_is_still_running(si_nav.driver) == True:
                                        if ge.gopher_email_count == 0:
                                            si_nav.driver.quit()

                                # these exception handlers get common errors, send an email to the requestor saying it 
                                # failed, moves the failed request from inbox to a failure box, and quits the driver.
                                except (NoSuchElementException, ElementClickInterceptedException, TimeoutException, UnicodeEncodeError): # UnicodeEncodeError handles garbage characters
                                    ge.send_outbound_email(ge.failure_subject, ge.failure_body) 
                                    ge.move_email_from_inbox_to_different_folder(ge.gopher_inbox, ge.gopher_failbox)
                                    si_nav.driver.quit()

                            except (NoSuchElementException, ElementClickInterceptedException, TimeoutException):
                                ge.send_outbound_email(ge.failure_subject, ge.failure_body)
                                ge.move_email_from_inbox_to_different_folder(ge.gopher_inbox, ge.gopher_failbox)
                                si_nav.driver.quit()

                        except (NoSuchElementException, ElementClickInterceptedException, TimeoutException):
                            ge.send_outbound_email(ge.failure_subject, ge.failure_body)
                            ge.move_email_from_inbox_to_different_folder(ge.gopher_inbox, ge.gopher_failbox)
                            si_nav.driver.quit()
                        
                    except (NoSuchElementException, ElementClickInterceptedException, TimeoutException):
                        ge.send_outbound_email(ge.failure_subject, ge.failure_body)
                        ge.move_email_from_inbox_to_different_folder(ge.gopher_inbox, ge.gopher_failbox)
                        si_nav.driver.quit()

                except (NoSuchElementException, ElementClickInterceptedException, TimeoutException):
                    ge.send_outbound_email(ge.failure_subject, ge.failure_body)
                    ge.move_email_from_inbox_to_different_folder(ge.gopher_inbox, ge.gopher_failbox)
                    si_nav.driver.quit()





if __name__ == "__main__": 

    main()

# this will run every one minute afterwards
schedule.every(1).minute.do(main)

# this keeps the script running forever at 1 minute intervals
while True:
    schedule.run_pending()
    time.sleep(1)


    
