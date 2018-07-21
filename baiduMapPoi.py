from urllib.parse import urlencode
import requests
from requests.exceptions import RequestException
import json
import xlwt
import xlrd
from xlutils.copy import copy

search_Poi = '学校'
search_city_Num = 340
search_city_Name ='深圳'
Save_file_name = search_city_Name + search_Poi + ".xls"
'''
百度地图中下一页nn改变10
北京市|131
上海市|289
广州市|257
深圳市|340
成都市|75
天津市|332
南京市|315
杭州市|179
武汉市|218
重庆市|132
'''
def get_page_index(nn,search_name,city_num):
    data = {
        'newmap':'1',
        'reqflag':'pcmap',
        'biz': '1',
        'from':'webmap',
        'da_par':'direct',
        'pcevaname':'pc4.1',
        'qt':'s',
        'da_src':'searchBox.button',
        'wd': search_name,
        'c': city_num,
        'src':'0',
        'wd2':'',
        'sug':'0',
        'l':'12',
        'biz_forward':'{"scaler": 1, "styles": "pl"}',
        'sug_forward':'',
        'tn':'B_NORMAL_MAP',
        'nn':nn,
        'ie':'utf - 8'
    }
    url = "http://map.baidu.com/?" + urlencode(data)
    response = requests.get(url)
    try:
        if response.status_code == 200:
            return response.text
        return None
    except RequestException:
        print("初始请求页出错")
        return None
def parse_html_detail(html):
    data = json.loads(html)
    name, addr, diPointY, diPointX, phone, shop_hours, image = [], [], [], [], [], [], []
    if 'content' not in data.keys():
        print("爬取结束")
        exit()
    if data and 'content' in data.keys():
        for content in data.get('content'):
            name_ = content.get('name')
            addr_ = content.get('addr')
            diPointX_ = content.get('diPointX')
            diPointY_ = content.get('diPointY')
            ext = content.get('ext')
            try:
                detail_info = ext.get('detail_info')
            except:
                print("此Poi没有detail_info记录")
            phone_ = detail_info.get('phone')
            shop_hours_ = detail_info.get('shop_hours')
            image_ = detail_info.get('image')
            name.append(name_)
            addr.append(addr_)
            diPointX.append(diPointX_)
            diPointY.append(diPointY_)
            phone.append(phone_)
            shop_hours.append(shop_hours_)
            image.append(image_)
    return name, addr, diPointX, diPointY, phone, shop_hours, image
def write_to_excel(name, addr, diPointX, diPointY, phone, shop_hours, image,nn):
    workbook_previous = xlrd.open_workbook(Save_file_name, formatting_info=True)
    workbook_new = copy(workbook_previous)
    data_sheet = workbook_new.get_sheet(0)
    for i in range(len(name)):
        data_sheet.write(i + nn, 0, name[i])
        data_sheet.write(i + nn, 1, addr[i])
        data_sheet.write(i + nn, 2, diPointX[i])
        data_sheet.write(i + nn, 3, diPointY[i])
        data_sheet.write(i + nn, 4, phone[i])
        data_sheet.write(i + nn, 5, shop_hours[i])
        data_sheet.write(i + nn, 6, image[i])
        # 保存文件
    workbook_new.save(Save_file_name)
def main(nn,search_name,city_num):
    html = get_page_index(nn,search_name, city_num)
    name, addr, diPointX, diPointY, phone, shop_hours, image = parse_html_detail(html)
    write_to_excel(name,addr,diPointX,diPointY,phone,shop_hours,image,nn)

if __name__ == '__main__':
    i = 0
    workbook = xlwt.Workbook(encoding='utf-8')
    data_sheet = workbook.add_sheet('sheet')
    workbook.save(Save_file_name)
    while(True):
        i = i*10
        main(i, search_Poi, search_city_Num)
        print("已下载%d条数据" % i)
        i = int(i / 10 + 1)
