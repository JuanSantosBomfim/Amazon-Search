import sys, os, csv, time, xlsxwriter
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

class AmazonProductSearch(object):

    def __init__(self, search_list):
        self.url = "https://www.amazon.com/"
        self.url_first_page_searched = ""
        self.search_list = search_list
        self.products = []

        geckodriverLocation = os.path.dirname(os.path.abspath(sys.argv[0])) + '/External/geckodriver-v0.27.0-win64/geckodriver.exe'
        self.driver = webdriver.Firefox(executable_path=geckodriverLocation)

        self.html = self.driver.page_source
        self.soup = BeautifulSoup(self.html, 'html.parser')
        self.html = self.soup.prettify('utf-8')
        self.wait_time = 10

    def check_all_products(self):
        index_number = 0

        for product in self.search_list:
            self.search_product(self.url, product)
            self.url_first_page_searched = self.driver.current_url

            while True:
                try:
                    product_index = self.driver.find_element_by_xpath('//div[@data-index="' + str(index_number)+'"]')
                    asin = product_index.get_attribute("data-asin")
                    print("VERIFY PRODUCT / index_number = " + str(index_number) + " / asin = " + str(asin) + " / product_index = " + str(product_index))
                except NoSuchElementException:
                    print("FISRT PAGE ALL PRODUCTS CHECKED")
                    break

                if asin != "":
                    url = "https://www.amazon.com/dp/" + asin
                    self.driver.get(url)

                    try:
                        WebDriverWait(self.driver, self.wait_time).until(EC.presence_of_all_elements_located)
                    except TimeoutException:
                        print("timeoutException")

                    price = self.get_product_price()
                    name = self.get_product_name()

                    product = {}
                    product['Name'] = name
                    product['Price'] = price
                    self.products.append(product)

                    print("ACCESS PRODUCT PAGE / Name = " + name + " / Price = " + price)
                    time.sleep(2)

                    self.driver.get(self.url_first_page_searched)

                index_number = index_number + 1

    def search_product(self, url, product):
        self.driver.get(url)

        try:
            WebDriverWait(self.driver, self.wait_time).until(EC.presence_of_all_elements_located)
        except TimeoutException:
            print("timeoutException")
        search_input = self.driver.find_element_by_id("twotabsearchtextbox")
        search_input.send_keys(product)

        search_button = self.driver.find_element_by_xpath('//*[@id="nav-search"]/form/div[2]/div/input')
        search_button.click()

    def get_product_name(self):
        product_name = "not available"
        try:
            product_name = self.driver.find_element_by_id("productTitle").text
        except:
            pass

        return product_name

    def get_product_price(self):
        price = "not available"
        try:
            price = self.driver.find_element_by_id("priceblock_ourprice").text
        except:
            pass

        try:
            price = self.driver.find_element_by_id("priceblock_dealprice").text
        except:
            pass

        return price

    def create_spreadsheets_xlsx(self):
        workbook = xlsxwriter.Workbook('Amazon Product Spreadsheets.xlsx')
        worksheet = workbook.add_worksheet("Product Data")
        row = 0
        col = 0
        for product_data in self.products:
            worksheet.write(row, col, product_data['Name'])
            worksheet.write(row, col + 1, product_data['Price'])
            row += 1
        workbook.close()
        print("CREATE spreadsheets with names and prices")

search_list = ["IPHONE"]

amazon = AmazonProductSearch(search_list)
amazon.check_all_products()
amazon.create_spreadsheets_xlsx()

# Objective:
# Open Amazon Site, Search By IPHONE,
# Get Name and Price of ALL Products in the First Page,
# Create a Spreadsheets with the Data
