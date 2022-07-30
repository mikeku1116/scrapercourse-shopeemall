from webbrowser import Chrome
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
import time
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import re


driver = webdriver.Chrome(ChromeDriverManager().install())

result = []
for page in range(1, 3):

    driver.get(
        f'https://shopee.tw/mall/%E5%B1%85%E5%AE%B6%E7%94%9F%E6%B4%BB-cat.11040925/popular?pageNumber={page}')

    ActionChains(driver).move_by_offset(100, 100).click().perform()

    locator = (By.CSS_SELECTOR,
               "div[class='col-xs-2 recommend-products-by-view__item-card-wrapper']")

    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located(locator),
        "找不到指定的元素"
    )

    cards = driver.find_elements(
        By.CSS_SELECTOR, "div[class='col-xs-2 recommend-products-by-view__item-card-wrapper']")

    items = []
    for card in cards:
        # ActionChains(driver).move_to_element(card).perform()

        title = card.find_element(
            By.CSS_SELECTOR, "div[class='ie3A+n bM+7UW Cve6sh']").text
        price = card.find_element(
            By.CSS_SELECTOR, "div[class='vioxXd rVLWG6']").text
        link = card.find_element(By.TAG_NAME, "a").get_attribute('href')
        items.append((title, price, link))

    # print(items)

    for item in items:
        driver.get(item[2])

        for i in range(5):
            driver.execute_script(
                "window.scrollTo(0, document.body.scrollHeight)")
            time.sleep(3)

        comments = driver.find_elements(By.CSS_SELECTOR, "div[class='Em3Qhp']")
        for comment in comments:
            result.append((item[0], item[1], comment.text))
        break

    # print(f"第{page}頁")
    # print(result)

data = pd.DataFrame(result, columns=['title', 'price', 'comment'])

# 單一字元資料清理
data['title'] = data['title'].str.replace('【', '')

# 多個字元資料清理
signs = ['】', '【']
data['title'] = data['title'].replace(dict.fromkeys(signs, ''), regex=True)

# emoji表情符號清理
emoji_pattern = re.compile(
    "(["
    "\U0001F1E0-\U0001F1FF"  # flags (iOS)
    "\U0001F300-\U0001F5FF"  # symbols & pictographs
    "\U0001F600-\U0001F64F"  # emoticons
    "\U0001F680-\U0001F6FF"  # transport & map symbols
    "\U0001F700-\U0001F77F"  # alchemical symbols
    "\U0001F780-\U0001F7FF"  # Geometric Shapes Extended
    "\U0001F800-\U0001F8FF"  # Supplemental Arrows-C
    "\U0001F900-\U0001F9FF"  # Supplemental Symbols and Pictographs
    "\U0001FA00-\U0001FA6F"  # Chess Symbols
    "\U0001FA70-\U0001FAFF"  # Symbols and Pictographs Extended-A
    "\U00002702-\U000027B0"  # Dingbats
    "])"
)

data['comment'] = data['comment'].str.replace(emoji_pattern, '')

# print(data)

data.to_excel('shopeemall.xlsx', sheet_name='shopeemall', index=False)
