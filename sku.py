from PIL import Image
import requests
from io import BytesIO
import xlsxwriter as xwrite
import json
import os 
import time


newSku_input = input("Add a New Sku? [y / n]: ")

if newSku_input == "n":
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
# remove = input("Remove Images? [y / n]: ")
# if remove == 'y':
#     time.sleep(5)
#     print(f"Sleeping for {5} Seconds...")
#     for sku in skus:
#         os.remove(f"SKU: {sku}.png")

if newSku_input == "y":
    with open("skulist.json") as list:
        skusdata = json.load(list)
        skus = skusdata['skus']
        #print(skus)

    workbook = xwrite.Workbook('test_excel.xlsx')
    worksheet = workbook.add_worksheet()
    new_sku = input("Enter Sku: ")
    imageURL = (f"https://images.footlocker.com/is/image/EBFL2/{new_sku}?wid=60&hei=60&fmt=png-alpha")

    r = requests.get(imageURL)
    img = Image.open(BytesIO(r.content))
    imgSave = img.save(f"test{new_sku}.png")

    row = len(skus)*2
    col = len(skus)*2
    worksheet.insert_image(row + 3, col + 4,  f"test{new_sku}.png")
    worksheet.write(row + 1, col + 1, new_sku)
    print(f"G{row}", f'test{new_sku}.png')
    print(row)

