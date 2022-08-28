from PIL import Image
import requests
from io import BytesIO
import xlsxwriter as xwrite
import json
import os 
import time



row = 1
col = 0
with open("skulist.json") as list: #### READS SKU LIST
    skusdata = json.load(list)
    skus = skusdata['skus']


workbook = xwrite.Workbook('Sku_List.xlsx') ##### CREATES EXCEL FILE
worksheet = workbook.add_worksheet()
for sku in skus: #### GOES THROUGH SKU LIST


    imageURL = (f"https://images.footlocker.com/is/image/EBFL2/{sku}?wid=60&hei=60&fmt=png-alpha")

    r = requests.get(imageURL)
    img = Image.open(BytesIO(r.content))
    imgSave = img.save(f"SKU: {sku}.png") #### PULLS IMAGE FROM URL

    worksheet.insert_image(row - 1, col + 2,  f"SKU: {sku}.png") 
    worksheet.write(row, col, sku)
    print(f'{sku} : Done', end= '\n')
    row += 2 # GOES DOWN TWO ROWS


workbook.close() #### WRITES EXCEL FILE


