# from openpyxl import Workbook, load_workbook
#
#
# result_article = Workbook()
# article_sheet = result_article.active
# articles=[]
# r_a = 2
# id=1
#
# article_sheet.cell(row=1,column=1).value = 'Article'
# article_sheet.cell(row=1,column=2).value = 'Prix Unitaire'
#
# book = load_workbook('C:/Users/ATL/Desktop/excel projet/all_factures.xlsx')
# sheet = book.active
#
# for row in range(2,sheet.max_row+3):
#     if str(sheet.cell(row=row , column=8).value).strip() not in articles:
#         articles.append(str(sheet.cell(row=row , column=8).value))
#         article_sheet.cell(row=r_a, column=1).value =id
#         article_sheet.cell(row=r_a ,column=2).value=str(sheet.cell(row=row , column=8).value)
#         article_sheet.cell(row=r_a, column=3).value = sheet.cell(row=row, column=10).value
#         r_a+=1
#         id+=1
#
#
# result_article.save('C:/Users/ATL/Desktop/excel projet/Articles.xlsx')
strrr = '  dfghjkl '
print(strrr.strip()+" dklqciqsqsmlckqslk")