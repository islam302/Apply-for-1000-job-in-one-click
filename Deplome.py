# -*- coding: utf-8 -*-
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from openpyxl import Workbook, load_workbook
from django.http import Http404, HttpResponse
from ChromeDriver import WebDriver
from selenium import webdriver
from bs4 import BeautifulSoup
from random import randint
import pandas as pd
import requests
import random
import json
import time
import sys
import os



class LMSBot:

    def __init__(self):
        self.url = "https://bblms.kfu.edu.sa/"
        self.driver = None

    def start_driver(self):
        self.driver = WebDriver.start_driver(self)
        return self.driver

    def login(self, username, password):
        try:
            self.driver.get(self.url)

            username_field = self.driver.find_element(By.XPATH, '//*[@id="user_id"]')
            username_field.send_keys(username)

            password_field = self.driver.find_element(By.XPATH, '//*[@id="password"]')
            password_field.send_keys(password)

            login_button = self.driver.find_element(By.XPATH, "//button[contains(text(), 'تسجيل الدخول')]")
            login_button.click()

            accept1 = self.driver.find_element(By.XPATH, '//*[@id="btnYes"]')
            accept1.click()
            time.sleep(3)
        except :
            pass

    def get_subjects(self) -> list:
        div = WebDriverWait(self.driver, 10).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, '.portletList-img.courseListing.coursefakeclass')))
        links = div.find_elements(By.TAG_NAME, "a")
        subjects = []
        for link in links:
            subject_name = link.text
            subject_url = link.get_attribute("href")
            subjects.append({"name": subject_name, "url": subject_url})
        return subjects

    def classes_links(self) -> list:
        time.sleep(2)
        link_text = 'المحاضرات المسجلة والمحتوى الرقمي'
        link_element = self.driver.find_element(By.LINK_TEXT, link_text)
        link_element.click()

        time.sleep(3)
        ul = self.driver.find_element(By.CLASS_NAME, 'contentList')
        li = ul.find_elements(By.TAG_NAME, 'li')
        links = []
        del li[0:3]

        for element in li:
            try:
                div = element.find_element(By.CSS_SELECTOR, '.item.clearfix')
                link = div.find_element(By.TAG_NAME, 'h3').find_element(By.TAG_NAME, 'a').get_attribute('href')
                links.append(link)
            except Exception as e:
                pass
        return links

    def get_units_links(self, link):
        self.driver.get(link)
        ul = self.driver.find_element(By.XPATH, '//*[@id="content_listContainer"]')
        li = ul.find_elements(By.TAG_NAME, 'li')
        li = li[1:-1]
        links = []
        for element in li:
            try:
                div = element.find_element(By.CSS_SELECTOR, '.item.clearfix')
                label = div.find_element(By.TAG_NAME, 'h3').find_element(By.TAG_NAME, 'a').text
                if 'التقييم الذاتي' not in label:
                    l = div.find_element(By.TAG_NAME, 'h3').find_element(By.TAG_NAME, 'a').get_attribute('href')
                    links.append({"link": l, "label": label})
            except Exception as e:
                pass
        return links

    def watch_video(self, link):
        try:
            try:
                self.check_secure_connection_error()
            except:
                pass
            self.driver.get(link)
            video_link = self.driver.find_element(By.ID, 'content_listContainer').find_elements(By.TAG_NAME, 'li')[3] \
                .find_element(By.TAG_NAME, 'a').get_attribute('href')
            self.driver.get(video_link)
            WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located((By.TAG_NAME, 'video')))
            play_button = WebDriverWait(self.driver, 5).until(
                EC.element_to_be_clickable((By.TAG_NAME, 'video')))

            if play_button:
                play_button.click()
                time.sleep(random.randint(6, 7))
                return True
            else:
                return False
        except:
            return False

    def listen_audio(self, link):
        try:
            try:
                self.check_secure_connection_error()
            except:
                pass
            self.driver.get(link)
            third_item_link = self.driver.find_element(By.ID, 'content_listContainer') \
                .find_elements(By.TAG_NAME, 'li')[2] \
                .find_element(By.TAG_NAME, 'a').get_attribute('href')
            self.driver.get(third_item_link)
            play_button = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.TAG_NAME, "video"))
            )
            if play_button:
                time.sleep(random.randint(6, 7))
                return True
        except Exception as e:
            return False

    def logout(self):
        self.driver.get('https://bblms.kfu.edu.sa/')
        self.driver.find_element(By.ID, "topframe.logout.label").click()

    def login_report(self, username):
        try:
            file_exists = os.path.isfile("Login_Report.csv")
            if file_exists:
                df = pd.read_csv("Login_Report.csv", encoding="utf-8-sig")

                if username in df['اسم المستخدم'].values:
                    return
                else:
                    new_row = pd.DataFrame({"اسم المستخدم": [username], "بيانات تسجيل الدخول خاطئة": [""]})
                    df = pd.concat([pd.DataFrame(new_row), df], ignore_index=True, sort=False)
            else:
                df = pd.DataFrame(columns=["اسم المستخدم", "بيانات تسجيل الدخول خاطئة"])
                new_row = pd.DataFrame({"اسم المستخدم": [username], "بيانات تسجيل الدخول خاطئة": [""]})
                df = pd.concat([pd.DataFrame(new_row), df], ignore_index=True, sort=False)

            df.to_csv("Login_Report.csv", index=False, encoding="utf-8-sig")
        except Exception as e:
            pass

    def report(self, username, inter):
        try:
            file_path = "REPORT.xlsx"
            if os.path.isfile(file_path):
                wb = load_workbook(file_path)
                ws = wb.active
            else:
                wb = Workbook()
                ws = wb.active
                ws.append(["اسم المستخدم", "الحالة"])

            username_exists = False
            for row in ws.iter_rows(values_only=True):
                if row[0] == username:
                    # Remove the prefix "APP-3270-103-CRN65994-Term144420-M: " from the status
                    inter = [s.replace("APP-3270-103-CRN65994-Term144420-M: ", "") for s in inter]
                    row = [username] + inter
                    username_exists = True
                    break

            if not username_exists:
                if inter:
                    # Remove the prefix "APP-3270-103-CRN65994-Term144420-M: " from the status
                    inter = [s.replace("APP-3270-103-CRN65994-Term144420-M: ", "") for s in inter]
                    row = [username] + inter
                else:
                    row = [username, "هذا الحساب جيد"]
                ws.append(row)
            wb.save(file_path)

        except Exception as e:
            return False

    def check_secure_connection_error(self):
        error_message_element = WebDriverWait(self.driver, 5).until(
            EC.presence_of_element_located(
                (By.XPATH, "//h1[contains(text(),'This site can't provide a secure connection')]"))
        )
        if error_message_element:
            logging.error('There is an error here. Please try again')
            time.sleep(3)
            self.driver.quit()
            sys.exit(1)

    def check_server_error(self):
        error_message_element = self.driver.find_element(By.XPATH,"//h2[contains(text(),'404 - File or directory not found.')]")
        if error_message_element:
            return True
        else:
            return False

    @staticmethod
    def data():
        try:
            data = pd.read_excel('users1.xlsx')
            collage_numbers = data['الرقم الجامعي'].tolist()
            paswordat = data['الرقم السري'].tolist()
            usernames = collage_numbers
            passwords = paswordat
            login_data = zip(usernames, passwords)
            return dict(login_data)
        except:
            return False

    def main(self, number=0):
        users = self.data()
        if not users:
            print("Check your users.xlsx file")
            return False
        self.start_driver()

        for username, password in users.items():
            list_error = []
            self.login(username, password)
            try:
                if self.driver.current_url == "https://bblms.kfu.edu.sa/webapps/login/":
                    print(f"Credentials for {username} are incorrect. Skipping...")
                    self.login_report(username)
                    time.sleep(3)
                    continue
            except TimeoutException:
                pass
            try:
                accept_button = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="agree_button"]')))
                accept_button.click()
            except TimeoutException:
                pass

            subjects = self.get_subjects()

            for subject in subjects:
                subject_name = subject["name"]
                subject_href = subject["url"]
                self.driver.get(subject_href)

                for link in self.classes_links()[-number:]:
                    lecture_links = self.get_units_links(link)
                    self.driver.refresh()

                    for inner_lecture in lecture_links:
                        audio_available = self.listen_audio(inner_lecture['link'])
                        if not audio_available:
                            video_available = self.watch_video(inner_lecture['link'])

                            if video_available:
                                list_error.append(
                                    f"{subject_name} >> {inner_lecture['label']} لا يعمل (الملف الصوتي) ولكن يعمل الفيديو")
                            else:
                                error = self.check_server_error()
                                if error:
                                    list_error.append(
                                        f"{subject_name} >> {inner_lecture['label']} لا يعمل (الملف الصوتي والفيديو)")
                                    continue
                                else:
                                    print(f'There is a Problem in this user : {username}')
                                    list_error.append(
                                        f"{subject_name} >> {inner_lecture['label']} لا يعمل (الملف الصوتي والفيديو)")
                                    self.report(username, list_error)
                                    self.driver.quit()
                                    time.sleep(5)
                                    exit()

            self.logout()
            self.report(username, list_error)
            print(f"User {username}  done...")
            time.sleep(5)

        print('All users Done, Have a nice day')
        time.sleep(2)
        self.driver.quit()


if __name__ == '__main__':
    bot = LMSBot()
    use = input("1- Start all Lectures ?\n2- Specific Classes ?\n")
    if use == "1":
        try:
            bot.main()
        except:
            print("There is an Error please try again")
    elif use == "2":
        x = int(input('Enter The Number Of Classes : \n'))
        try:
            bot.main(x)
        except:
            print("There is an Error please try again")
    else:
        os.system('exit')

