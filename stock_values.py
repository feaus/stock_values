import os, openpyxl, requests, bs4
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC

class Shares:
    def __init__(self):
        pass
    
    def get_shares(self, sheet, shares):
        """
        Toma todas las acciones que están en el excel.
        """
        for k in range(1, 1001):
            name = sheet.cell(row=3, column=k).value
            if name in ["Dólar", "Dólar "]:
                break
            if name != None:
                USAmarket = sheet.cell(row=4, column=k+1).value.upper()
                if name == "BAC":
                    shares.append("BCBA-" + name[0:2] + "." + name[2:3])
                else:
                    shares.append("BCBA-" + name)
                shares.append(USAmarket + "-" + name)
                
        print("Estas son las acciones que se van a buscar:")
        print(" - ".join(shares))
                
    def new_share(self, sheet, shares):
        """
        Pregunta si querés agregar alguna otra acción, en qué bolsa cotiza, y
        la agrega a la lista.
        """
        print("Querés agregar alguna otra acción?")
        answer = input()
        while answer.lower() in ["sí", "si", "s", "yes", "y"]:
            print("Introducir el nombre de la acción.")
            new_share_name = input()
            print("Introducir la bolsa en la que cotiza en USA (NASDAQ o NYSE)")
            new_share_market = input()
            new_share1 = "BCBA-" + new_share_name.upper()
            new_share2 = new_share_market.upper() + "-" + new_share_name.upper()
            shares.append(new_share1)
            shares.append(new_share2)
            fa.new_action_excel(new_share_name, sheet, new_share_market)
            print("Querés agregar alguna otra acción?")
            answer = input()
            if answer.lower() not in ["sí", "si", "s", "yes", "y"]:
                break
            
    def new_action_excel(self, new_share, sheet, new_share_market):
        """
        Agrega las columnas para la acción antes del dólar.
        """
        for c in range(1, 1001):
            if sheet.cell(row=4, column=c).value == "Blue":
                break
            if c == 100:
                raise Exception("No hay celda donde insertar la acción.")
        sheet.insert_cols(c)
        sheet.cell(row=4, column=c).value = "T C implicito"
        sheet.insert_cols(c)
        sheet.cell(row=4, column=c).value = new_share_market.upper()
        sheet.insert_cols(c)
        sheet.cell(row=4, column=c).value = "Cot CEDEAR"
        sheet.merge_cells(start_row=3, start_column=c, end_row=3, end_column=(c+2))
        sheet.cell(row=3, column=c).value = new_share.upper()
    
    def get_price(self, shares, sheet):
        """
        Obtiene todos los valores de las acciones y los guarda en un Excel.
        """
        i = 2
        global r
        for k in shares:
            browser.get("https://es.tradingview.com/symbols/%s" % (k))
            WebDriverWait(browser, 20).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "div.js-symbol-last > span:nth-child(1)")))
            tradingValue = browser.find_element_by_css_selector(".js-last-price-block-value-row > div:nth-child(1) > span:nth-child(1)")
            shareValue = str(tradingValue.text)
            tradingValue2 = browser.find_element_by_css_selector("div.js-symbol-last > span:nth-child(1)")
            shareValue2 = str(tradingValue2.text)   
            
            if shareValue != "":
                sheet.cell(row=r, column=i).value = float(shareValue)
                print(k + ": " + shareValue)  
            else:
                sheet.cell(row=r, column=i).value = float(shareValue2)
                print(k + ": " + shareValue2)
                
            if (shareValue == "" and shareValue2 == "") or (float(shareValue) == 0 and float(shareValue2) == 0):
                browser.quit()
                raise Exception("Hubo un error en la recolección de datos. Intentar nuevamente.")
                
            i += 1
            if i % 3 == 1:
                i += 1
        browser.quit()
        print("\nNavegador cerrado.")
    

if __name__ == "__main__":
    try:
        fa = Shares()
        
        #os.chdir("")
        directory = os.getcwd()
        
        print("Abriendo Excel...")
        workbook = openpyxl.load_workbook("Cotizaciones Cedear.xlsx")
        sheet = workbook["Hoja1"]
        
        r = 1
        while datetime.today().strftime('%Y-%m-%d 00:00:00') != str(sheet.cell(row = r, column = 1).value):
            r += 1
            if r == 1000:
                raise Exception("No se pudo encontrar la fecha de hoy en el Excel.")
                break
        print("Excel abierto.")
        
        shares = []
        
        fa.get_shares(sheet, shares)
        fa.new_share(sheet, shares)
        
        print("Abriendo internet...")
        options = Options()
        options.headless = True
        browser = webdriver.Firefox(options=options)
        #browser = webdriver.Firefox() PARA DEBUG
        print("Navegador abierto.")
        
        print("Obteniendo datos...\n")
        fa.get_price(shares, sheet)
        
        workbook.save("Cotizaciones Cedear1.xlsx")
        workbook.close()
        print("Todo listo!")
        
        print('Archivo guardado como "Cotizaciones Cedear1.xlsx" en ' + directory + ".")
    
    except:
        print("Hubo un error. Intentar nuevamente más tarde.")
        browser.quit()
        workbook.close()
