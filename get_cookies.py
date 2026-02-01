import pickle
from selenium import webdriver

browser = webdriver.Firefox()


def get_cookies():
    #browser.get("https://login.aliexpress.com/buyer.htm?return=https%3A%2F%2Fwww.aliexpress.com%2F&random=CEA73DF4D81D4775227F78080B9B6126")
    browser.get("https://aliexpress.ru/account?spm=a2g2w.home.profil.1.5a505586y4WLcS")	
    print('input your username and password in Firefox and hit Submit')
    input('Hit Enter here if you have summited the form: <Enter>')
    cookies = browser.get_cookies()
    pickle.dump(cookies, open("cookies.pickle", "wb"))
    browser.close()


def set_cookies():
    browser.get("https://aliexpress.ru")
    cookies = pickle.load(open("cookies.pickle", "rb"))
    for cookie in cookies:
        browser.add_cookie(cookie)
    browser.get("https://aliexpress.ru/one-price")


if __name__ == '__main__':
    get_cookies()
    #set_cookies()
