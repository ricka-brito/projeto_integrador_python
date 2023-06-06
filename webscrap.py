import os
import time
import pandas as pd
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from xlsxwriter import *

class scrap:

    def __init__(self):
        self.servico = Service(ChromeDriverManager().install())
        self.options = webdriver.ChromeOptions()
        self.options.add_argument("--headless=new")

    def pesquisa(self):
        self.navegador = webdriver.Chrome(service=self.servico, options=self.options)
        time.sleep(5)
        self.navegador.get("https://projetosemds.com.br/hbrito/produtos.html")
        infos = {"nome": [], "preco": [], "img": []}
        dir = './imgs'
        os.mkdir(dir)
        for i in range(1, 100):
            try:
                nome = self.navegador.find_element("xpath", f"/html/body/div[2]/div[2]/div[{i}]/h1").text
                preco = self.navegador.find_element("xpath", f"/html/body/div[2]/div[2]/div[{i}]/div[2]/p/span").text
                img = self.navegador.find_element('xpath', f'/html/body/div[2]/div[2]/div[{i}]/div[1]/img').get_attribute('src')
                self.navegador.get(img)
                self.navegador.save_screenshot(f"./imgs/{nome}.png")
                self.navegador.get("https://projetosemds.com.br/hbrito/produtos.html")
                infos["nome"].append(nome)
                infos["preco"].append(float(preco.replace("R$", "").replace(",", ".").strip()))
                infos["img"].append(f"./imgs/{nome}.png")
                time.sleep(0.2)
            except:
                df = pd.DataFrame(infos)
                writer = pd.ExcelWriter('output.xlsx', engine='xlsxwriter')
                df.to_excel(writer, index=False)

                workbook = writer.book
                ws = writer.sheets['Sheet1']

                format1 = workbook.add_format({'num_format': 'R$ #,##0.00'})

                (max_row, max_col) = df.shape

                column_settings = [{"header": column} for column in df.columns]

                ws.add_table(0, 0, max_row, max_col - 1, {"columns": column_settings})

                for column in df:
                    column_width = max(df[column].astype(str).map(len).max(), len(column))
                    col_idx = df.columns.get_loc(column)
                    writer.sheets['Sheet1'].set_column(col_idx, col_idx, column_width)

                ws.set_column(1, 1, None, format1)
                writer.close()
                break


scrap().pesquisa()