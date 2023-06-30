import pandas as pd
import time
import json
import requests
import os
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

#.envの読み込み
load_dotenv('.env')
api_key = os.environ.get("CUSTOM_SEARCH_API_KEY")
cx = os.environ.get("CX")



# エクセルファイルの読み込み
df = pd.read_excel("/Users/bocmitsuhashi/Desktop/fixed_application/nutirtion_calc_on_excel/calcsheet.xlsx",engine='openpyxl')  # change to your excel file path

def get_nutritional_info(food_items, selected_elements, api_key, cx):
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    webdriver_service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=webdriver_service, options=chrome_options)

    nutritional_info = {}

    for food_name, food_weight in food_items.items():
        try:
            # Google Custom Search APIを使って、食品DBのURLを取得する
            search_query = "日本食品標準成分表 " + " "+ food_name
            api_key = api_key
            cx = cx
            url = f"https://www.googleapis.com/customsearch/v1?key={api_key}&cx={cx}&q={search_query}&num=10&siteSearch=fooddb.mext.go.jp"
            response = requests.get(url)
            response_json = json.loads(response.text)

            first_result_url = None
            for item in response_json.get('items', []):
                if 'https://fooddb.mext.go.jp' in item['link']:
                    first_result_url = item['link']
                    break

            if not first_result_url:
                print(f"Could not find suitable search result for {food_name}")
                continue

            driver.get(first_result_url)
            time.sleep(1)  

            food_DB_name = driver.find_element(By.XPATH,'//*[@id="head"]/div[2]').text

            input_box = driver.find_element(By.XPATH,'//*[@id="main"]/form/ul/li[2]/input[1]')
            input_box.clear()
            input_box.send_keys(str(food_weight))

            calc_button = driver.find_element(By.CLASS_NAME,'designed_button')
            calc_button.click()

            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, selected_elements[1]['xpath'])))
            result_elements = {}
            for num, element in selected_elements.items():
                try:
                    result_elements[element['name']] = driver.find_element(By.XPATH, element['xpath']).text
                except:
                    result_elements[element['name']] = "N/A" 

            nutritional_info[food_name] = {
                "DB_name": food_DB_name,
                "weight": food_weight,
                "url": first_result_url,
                "nutritional_values": result_elements
            }
        except Exception as e:
            print(f"An error occurred with {food_name}: {str(e)}")
            continue

    driver.quit()

    return nutritional_info

# 食品DBからエクセルに転記
def update_dataframe(df,nutrition_data, nutrition_columns):
    for index, row in df.iterrows():
        food_name = row['食品名']  
        if food_name in nutrition_data:
            for nutrient in nutrition_columns:
                if nutrient['name'] in df.columns:
                    df.loc[index, nutrient['name']] = nutrition_data[food_name]["nutritional_values"].get(nutrient['name'], "N/A")
            
            
            if 'DB_name' in df.columns:
                df.loc[index, 'DB_name'] = nutrition_data[food_name].get('DB_name', "N/A")
            else:
                df['DB_name'] = pd.Series(dtype=str)  
                df.loc[index, 'DB_name'] = nutrition_data[food_name].get('DB_name', "N/A")

           
            if 'url' in df.columns:
                df.loc[index, 'url'] = nutrition_data[food_name].get('url', "N/A")
            else:
                df['url'] = pd.Series(dtype=str)  
                df.loc[index, 'url'] = nutrition_data[food_name].get('url', "N/A")


    df = df[[col for col in df.columns if col not in ['DB_name', 'url']] + ['DB_name', 'url']]
    return df

food_items = df.set_index('食品名')['使用量'].to_dict()  
selected_elements = {
    1: {'name': '廃棄率', 'xpath': '//*[@id="nut"]/tbody/tr[1]/td[2]'},
    2: {'name': 'エネルギー(kcal)', 'xpath': '//*[@id="nut"]/tbody/tr[2]/td[2]'},
    3: {'name': 'エネルギー(kJ)', 'xpath': '//*[@id="nut"]/tbody/tr[3]/td[1]'},
    4: {'name': '水分', 'xpath': '//*[@id="nut"]/tbody/tr[4]/td[2]'},
    5: {'name': 'アミノ酸組成によるたんぱく質', 'xpath': '//*[@id="nut"]/tbody/tr[6]/td[2]'},
    6: {'name': 'たんぱく質', 'xpath': '//*[@id="nut"]/tbody/tr[7]/td[2]'},
    7: {'name': '脂肪酸のトリアシルグリセロール当量', 'xpath': '//*[@id="nut"]/tbody/tr[9]/td[2]'},
    8: {'name': 'コレステロール', 'xpath': '//*[@id="nut"]/tbody/tr[10]/td[2]'},
    9: {'name': '脂質', 'xpath': '//*[@id="nut"]/tbody/tr[11]/td[2]'},
    10: {'name': '利用可能炭水化物（単糖当量）', 'xpath': '//*[@id="nut"]/tbody/tr[12]/td[5]'},
    11: {'name': '利用可能炭水化物（質量計）', 'xpath': '//*[@id="nut"]/tbody/tr[13]/td[3]'},
    12: {'name': '差引き法による利用可能炭水化物', 'xpath': '//*[@id="nut"]/tbody/tr[14]/td[2]'},
    13: {'name': '食物繊維総量', 'xpath': '//*[@id="nut"]/tbody/tr[15]/td[2]'},
    14: {'name': '糖アルコール', 'xpath': '//*[@id="nut"]/tbody/tr[16]/td[2]'},
    15: {'name': '炭水化物', 'xpath': '//*[@id="nut"]/tbody/tr[17]/td[2]'},
    16: {'name': '有機酸', 'xpath': '//*[@id="nut"]/tbody/tr[18]/td[2]'},
    17: {'name': '灰分', 'xpath': '//*[@id="nut"]/tbody/tr[19]/td[2]'},
    18: {'name': 'ナトリウム', 'xpath': '//*[@id="nut"]/tbody/tr[20]/td[3]'},
    19: {'name': 'カリウム', 'xpath': '//*[@id="nut"]/tbody/tr[21]/td[2]'},
    20: {'name': 'カルシウム', 'xpath': '//*[@id="nut"]/tbody/tr[22]/td[2]'},
    21: {'name': 'マグネシウム', 'xpath': '//*[@id="nut"]/tbody/tr[23]/td[2]'},
    22: {'name': 'リン', 'xpath': '//*[@id="nut"]/tbody/tr[24]/td[2]'},
    23: {'name': '鉄', 'xpath': '//*[@id="nut"]/tbody/tr[25]/td[2]'},
    24: {'name': '亜鉛', 'xpath': '//*[@id="nut"]/tbody/tr[26]/td[2]'},
    25: {'name': '銅', 'xpath': '//*[@id="nut"]/tbody/tr[27]/td[2]'},
    26: {'name': 'マンガン', 'xpath': '//*[@id="nut"]/tbody/tr[28]/td[2]'},
    27: {'name': 'ヨウ素', 'xpath': '//*[@id="nut"]/tbody/tr[29]/td[2]'},
    28: {'name': 'セレン', 'xpath': '//*[@id="nut"]/tbody/tr[30]/td[2]'},
    29: {'name': 'クロム', 'xpath': '//*[@id="nut"]/tbody/tr[31]/td[2]'},
    30: {'name': 'モリブデン', 'xpath': '//*[@id="nut"]/tbody/tr[32]/td[2]'},
    31: {'name': 'レチノール', 'xpath': '//*[@id="nut"]/tbody/tr[33]/td[3]'},
    32: {'name': 'α−カロテン', 'xpath': '//*[@id="nut"]/tbody/tr[34]/td[2]'},
    33: {'name': 'β−カロテン', 'xpath': '//*[@id="nut"]/tbody/tr[35]/td[2]'},
    34: {'name': 'β−クリプトキサンチン', 'xpath': '//*[@id="nut"]/tbody/tr[36]/td[2]'},
    35: {'name': 'β−カロテン当量', 'xpath': '//*[@id="nut"]/tbody/tr[37]/td[2]'},
    36: {'name': 'レチノール活性当量', 'xpath': '//*[@id="nut"]/tbody/tr[38]/td[2]'},
    37: {'name': 'ビタミンD', 'xpath': '//*[@id="nut"]/tbody/tr[39]/td[2]'},
    38: {'name': 'α−トコフェロール', 'xpath': '//*[@id="nut"]/tbody/tr[40]/td[2]'},
    39: {'name': 'β−トコフェロール', 'xpath': '//*[@id="nut"]/tbody/tr[41]/td[2]'},
    40: {'name': 'γ−トコフェロール', 'xpath': '//*[@id="nut"]/tbody/tr[42]/td[2]'},
    41: {'name': 'δ−トコフェロール', 'xpath': '//*[@id="nut"]/tbody/tr[43]/td[2]'},
    42: {'name': 'ビタミンK', 'xpath': '//*[@id="nut"]/tbody/tr[44]/td[2]'},
    43: {'name': 'ビタミンB1', 'xpath': '//*[@id="nut"]/tbody/tr[45]/td[2]'},
    44: {'name': 'ビタミンB2', 'xpath': '//*[@id="nut"]/tbody/tr[46]/td[2]'},
    45: {'name': 'ナイアシン', 'xpath': '//*[@id="nut"]/tbody/tr[47]/td[2]'},
    46: {'name': 'ナイアシン当量', 'xpath': '//*[@id="nut"]/tbody/tr[48]/td[2]'},
    47: {'name': 'ビタミンB6', 'xpath': '//*[@id="nut"]/tbody/tr[49]/td[2]'},
    48: {'name': 'ビタミンB12', 'xpath': '//*[@id="nut"]/tbody/tr[50]/td[2]'},
    49: {'name': '葉酸', 'xpath': '//*[@id="nut"]/tbody/tr[51]/td[2]'},
    50: {'name': 'パントテン酸', 'xpath': '//*[@id="nut"]/tbody/tr[52]/td[2]'},
    51: {'name': 'ビオチン', 'xpath': '//*[@id="nut"]/tbody/tr[53]/td[2]'},
    52: {'name': 'ビタミンC', 'xpath': '//*[@id="nut"]/tbody/tr[54]/td[2]'},
    53: {'name': 'アルコール', 'xpath': '//*[@id="nut"]/tbody/tr[55]/td[2]'},
    54: {'name': '食塩相当量', 'xpath': '//*[@id="nut"]/tbody/tr[56]/td[2]'}
}

nutrition_data = get_nutritional_info(food_items, selected_elements)
update_dataframe(df,nutrition_data, selected_elements.values())

# 保存
df.to_excel("/Users/bocmitsuhashi/Desktop/fixed_application/nutirtion_calc_on_excel/calcsheet.xlsx", index=False) 
