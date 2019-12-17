from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup
import re
import pandas as pd
from time import localtime, strftime, sleep
options = webdriver.ChromeOptions()
options.add_argument('headless')



def num_puri(num): return int(re.findall('\d+', num)[0])


def page_cal(num): return (num // 10) + 1


driver = webdriver.PhantomJS('../script/phantomjs/bin/phantomjs.exe')
driver.get('https://new.abb.com/search/results#query=af80')
sleep(1)

try:
    max_page = page_cal(
        num_puri(driver.find_element_by_css_selector(".OneABBSearchAside > span:nth-child(1)").text)
    )
    if not max_page:
        sleep(3)
        max_page = page_cal(
            num_puri(driver.find_element_by_css_selector(".OneABBSearchAside > span:nth-child(1)").text)
        )
except:
    try:
        sleep(3)
        max_page = page_cal(
            num_puri(driver.find_element_by_css_selector(".OneABBSearchAside > span:nth-child(1)").text)
        )
        if not max_page:
            sleep(3)
            max_page = page_cal(
                num_puri(driver.find_element_by_css_selector(".OneABBSearchAside > span:nth-child(1)").text)
            )
    except:
        print("check the internet.")
        exit(1)

print(f'All {max_page}Pages.')
now_page = 1

result_df = pd.DataFrame(columns=['width', 'depth', 'height', 'weight'])

while now_page <= max_page:
    search_df = pd.DataFrame(columns=['width', 'depth', 'height', 'weight'])
    names = []
    widths = []
    depths = []
    heights = []
    weights = []

    print(f'@@@ now on page {str(now_page)} / {str(max_page)}')
    list_html = driver.page_source
    list_soup = BeautifulSoup(list_html, 'html.parser')
    item_list = list_soup.select("li.OneABBSearchList-item > a:nth-child(4)")

    for item in item_list:
        item = re.sub('<[^>]*>', '', item.text)
        driver.execute_script(f"window.open('{item}')")
        driver.switch_to.window(driver.window_handles[1])
        sleep(5)

        try:
            item_html = driver.page_source
            item_soup = BeautifulSoup(item_html, 'html.parser')
            name = re.sub('<[^>]*>', '', str(item_soup.select(".display-name-repeat"))[1:-1])
            if not name:
                sleep(3)
                name = re.sub('<[^>]*>', '', str(item_soup.select(".display-name-repeat"))[1:-1])

        except AttributeError:
            try:
                sleep(5)
                item_html = driver.page_source
                item_soup = BeautifulSoup(item_html, 'html.parser')
                name = re.sub('<[^>]*>', '', str(item_soup.select(".display-name-repeat"))[1:-1])
                if not name:
                    sleep(3)
                    name = re.sub('<[^>]*>', '', str(item_soup.select(".display-name-repeat"))[1:-1])
            except AttributeError:
                print("인터넷 상태를 확인해주세요.")
                exit(1)

        names.append(name)
        widths.append(re.sub('<[^>]*>', '', str(item_soup.select(
            "div.attribute-group:nth-child(7) > ul:nth-child(2) > li:nth-child(1) > dl:nth-child(1) > dd:nth-child(2)"
        ))[1:-1]))
        depths.append(re.sub('<[^>]*>', '', str(item_soup.select(
            "div.attribute-group:nth-child(7) > ul:nth-child(2) > li:nth-child(2) > dl:nth-child(1) > dd:nth-child(2)"
        ))[1:-1]))
        heights.append(re.sub('<[^>]*>', '', str(item_soup.select(
            "div.attribute-group:nth-child(7) > ul:nth-child(2) > li:nth-child(3) > dl:nth-child(1) > dd:nth-child(2)"
        ))[1:-1]))
        weights.append(re.sub('<[^>]*>', '', str(item_soup.select(
            "div.attribute-group:nth-child(7) > ul:nth-child(2) > li:nth-child(4) > dl:nth-child(1) > dd:nth-child(2)"
        ))[1:-1]))
        driver.close()
        driver.switch_to.window(driver.window_handles[0])
        search_df = pd.DataFrame(
            {'name': names, 'width': widths, 'depth': depths, 'height': heights, 'weight': weights},
            columns=['width', 'depth', 'height', 'weight'], index=names)

    result_df = pd.concat([result_df, search_df])
    result_df
    print(search_df)

    if now_page != max_page:
        now_page += 1
        driver.find_element_by_css_selector(".OneABBSearchPagination-item-next").send_keys(Keys.CONTROL + "\n")

print(result_df)
writer = pd.ExcelWriter(f'../DB/AF80_검색결과_{strftime("%y_%m_%d_%H-%M", localtime())}.xlsx', engine='xlsxwriter')
result_df.to_excel(writer, sheet_name='Sheet1')
writer.save()
