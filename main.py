from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from ChromeDriver import WebDriver
from selenium import webdriver
from random import randint
import pandas as pd
import requests
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
        except:
            pass

    def get_subjects(self):
        ul = self.driver.find_element(By.CSS_SELECTOR, '.portletList-img.courseListing.coursefakeclass')
        li = ul.find_elements(By.TAG_NAME, 'li')
        links = {}
        for element in li:
            links[element.find_element(By.TAG_NAME, 'a').text] = element.find_element(By.TAG_NAME,
                                                                                        'a').get_attribute('href')
        return links

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
                l = div.find_element(By.TAG_NAME, 'h3').find_element(By.TAG_NAME, 'a').get_attribute('href')
                s = div.find_element(By.TAG_NAME, 'h3').find_element(By.TAG_NAME, 'a').find_element(By.TAG_NAME,
                                                                                                     'span').text
                links.append({"link": l, "label": s})
            except Exception as e:
                pass
        return links

    def watch_video(self, link):
        try:
            self.driver.get(link)
            video_link = self.driver.find_element(By.ID, 'content_listContainer').find_elements(By.TAG_NAME, 'li')[3] \
                .find_element(By.TAG_NAME, 'a').get_attribute('href')

            self.driver.get(video_link)

            WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located((By.TAG_NAME, 'video')))

            play_button = WebDriverWait(self.driver, 5).until(
                EC.element_to_be_clickable((By.TAG_NAME, 'video')))

            if play_button:
                play_button.click()
                time.sleep(7)
                return True
            else:
                return False
        except:
            return False

    def listen_audio(self, link):
        try:
            self.driver.get(link)

            self.driver.get(
                self.driver.find_element(By.ID, 'content_listContainer').find_elements(By.TAG_NAME, 'li')[2]
                    .find_element(By.TAG_NAME, 'a').get_attribute('href'))

            play = WebDriverWait(self.driver, 10).until(
                        EC.presence_of_element_located((By.TAG_NAME, "video"))
                    )
            if play:
                time.sleep(7)
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
            file_path = "REPORT.csv"
            file_exists = os.path.isfile(file_path)

            if file_exists:
                df = pd.read_csv(file_path, encoding="utf-8-sig")
            else:
                df = pd.DataFrame(columns=["اسم المستخدم", "الحالة"])

            if inter:
                if username in df['اسم المستخدم'].values:
                    df = df[df['اسم المستخدم'] != username]

                error_dict = {"اسم المستخدم": username}
                for n, error_text in enumerate(inter, start=1):
                    text_parts = error_text.split(":")
                    key = f"{n}"
                    text = ':'.join(text_parts[1:]).strip()
                    error_dict[key] = text
                df = pd.concat([df, pd.DataFrame(error_dict, index=[0])], ignore_index=True)

            else:
                text = "هذا الحساب جيد".split(":")[0].strip()
                if username in df['اسم المستخدم'].values:
                    df = df[df['اسم المستخدم'] != username]

                df = pd.concat([df, pd.DataFrame({"اسم المستخدم": [username], "الحالة": [text]})],
                               ignore_index=True)
            df.to_csv(file_path, mode='w', index=False, encoding="utf-8-sig")

        except Exception as e:
            pass

    def main(self, number=0):
        users = self.data()
        if not users:
            print("check your users.xlsx file ")
            return False
        self.start_driver()

        for username, password in users.items():
            list_error = []
            self.login(username, password)

            try:
                if self.driver.current_url == "https://bblms.kfu.edu.sa/webapps/login/":
                    print(f"Credentials for {username} are incorrect. Skipping...")
                    self.login_report(username)
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

            for label_S, subject in subjects.items():
                self.driver.get(subject)

                for link in self.classes_links()[-number:]:
                    lecture_links = self.get_units_links(link)
                    self.driver.refresh()

                    try:
                        for inner_lecture in lecture_links:
                            audio_available = self.listen_audio(inner_lecture['link'])

                            if not audio_available:
                                video_available = self.watch_video(inner_lecture['link'])

                                if video_available:
                                    list_error.append(
                                        f"{label_S} >> {inner_lecture['label']} لا يعمل (الملف الصوتي) ولكن يعمل الفيديو")
                                else:
                                    list_error.append(
                                        f"{label_S} >> {inner_lecture['label']} لا يعمل (الملف الصوتي والفيديو)")

                    except Exception as outer_error:
                        list_error.append(
                            f"{label_S} >> {inner_lecture['label']} لا يعمل (الملف الصوتي والفيديو)")

            self.logout()
            self.report(username, list_error)
            print(f"User {username} is done...")

    @staticmethod
    def data():
        try:
            data = pd.read_excel('users.xlsx')
            collage_numbers = data['الرقم الجامعي'].tolist()
            paswordat = data['الرقم السري'].tolist()
            usernames = collage_numbers
            passwords = paswordat
            login_data = zip(usernames, passwords)
            return dict(login_data)
        except:
            return False


if __name__ == '__main__':
    bot = LMSBot()
    use = input("1- Start all Lectures ?\n2- Specific Classes ?\n ")
    if use == "1":
        try:
            bot.main()
        except Exception as ex:
            print('There is something Wrong Please try again')
    elif use == "2":
        try:
            time.sleep(1)
            x = int(input('Enter The Number Of Classes : \n'))
            bot.main(x)
        except:
            print('There is something Wrong Please try again')
    else:
        os.system('exit')