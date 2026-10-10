import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

url = "https://top.baidu.com/board?tab=realtime"
headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3"}

r = requests.get(url, headers=headers,timeout=10)
r.encoding = 'utf-8'
soup = BeautifulSoup(r.text, 'html.parser')
items = soup.find_all("div",class_="category-wrap_iQLoo")

wb = Workbook()
ws = wb.active
ws.append(["排名","标题","热度"])

print(f"抓到{len(items)}条热搜")
for i,item in enumerate(items[:10],1):
    title = item.find("div",class_="c-single-text-ellipsis").text.strip()
    hot = item.find("div",class_="hot-index_1Bl1a").text.strip()
    ws.append([i, title, int(hot)])

header_font = Font(bold=True, color="FFFFFF")
header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
for cell in ws[1]:
    cell.font = header_font
    cell.fill = header_fill

ws.column_dimensions['A'].width = 8
ws.column_dimensions['B'].width = 50
ws.column_dimensions['C'].width = 15

wb.save("百度热搜.xlsx")
print(f"完成！共抓取{len(items)}条热搜，已保存到百度热搜.xlsx")
input("按任意键退出")