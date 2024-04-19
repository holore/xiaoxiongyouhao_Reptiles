import requests
from bs4 import BeautifulSoup
import xlwt

# 构造请求
headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko)'
                         ' Chrome/100.0.4896.127 Safari/537.36'}
url = 'https://www.xiaoxiongyouhao.com/chxi_report_list.php'

# 获得网页
rsp = requests.get(url, headers)
# 解析网页
soup = BeautifulSoup(rsp.text, 'lxml')

# 创建一个新的Excel文件和工作表用于储存 车系和车辆型号链接
workbook = xlwt.Workbook()
sheet = workbook.add_sheet('小熊油耗信息')
# 定义行号
row = 0


# 获取车系和车辆型号链接
data = soup.select('body > div.wrap_outer.wrap_contact > div.container.wrap_inner_body >.row')
# 遍历数据并将其写入Excel文件
for item in data:
    name = item.text  # 提取所有文字
    links = item.find_all('a')  # 提取所有a标签

    name_parts = [part.strip() for part in name.split('\n') if part.strip()]  # 空格去掉
    sheet.write(row, 0, name_parts[0])  # 将品牌名写入第一行

    for link in links:  # 循环提取车型链接
        href = link['href']
        sheet.write(row + 1 + links.index(link), 1, href)  # 将链接写入第二列
        sheet.write(row + 1 + links.index(link), 0, "https://www.xiaoxiongyouhao.com/" + name_parts[1 + links.index(link)])
    row = row + len(links) + 1
    print(row)

# 保存Excel文件
workbook.save('车系信息.xls')

