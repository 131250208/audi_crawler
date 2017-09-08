# -*- coding:utf-8 -*-
import xlrd
import requests
from bs4 import BeautifulSoup
import re
import xlsxwriter

def get_agencies():
    wkbook = xlrd.open_workbook(u"./经销商子站问题整理.xlsx")
    booksheet = wkbook.sheet_by_name(u"奥迪域名-按时间（完整）")

    row_index = 3
    agencies_list = []

    while True:
        # 读取经销商名字
        agency_name = booksheet.cell(row_index, 2).value

        # 读取经销商子站网址 第 row_index + 1 行 的 第 4 列
        url_original = booksheet.cell(row_index, 3).value
        # 并将空格、回车、换行、制表符等字符去掉
        illegal_label = re.compile("[ \n\t\r]")# 非法字符，空格等
        agency_url = illegal_label.sub("", url_original)
        # 为了排除url前后还有其他非法字符，用正则匹配出真正的url部分。如：承德庞大奥兴奥迪的url开头有不明空格（不是普通的" "）
        url_pattern = re.compile("[a-zA-Z0-9_.~!*'();:@&=+\-$,/?#\[\]]*")
        search_ob = re.search(url_pattern, agency_url)
        agency_url = search_ob.group()

        agency = {"name": agency_name, "url": agency_url}
        agencies_list.append(agency)

        if row_index == 452: # 如果读到末尾行,跳出循环
            break

        row_index += 1 # 行索引递增

    return agencies_list

def get_topicid(url):
    responce = requests.get(url)

    bs = BeautifulSoup(responce.text, "lxml")
    div_top =bs.find("div", class_ = "main").find("ul", class_ = "act-ul").find("div", class_ = "top")# 找到 class = top div

    if div_top == None: return None # 如果找不到 top，返回 None

    a = div_top.find("h2", class_ = "act-h2").find("a") # 找到 a 标签
    onclick =  a["onclick"] # e.g. showNews(5890)

    match_obj = re.match("showNews\((.*?)\)", onclick)

    return match_obj.group(1) # 返回topicid

def get_activity(url, act_type, topicid):
    if topicid == None: return None

    showTopicNews_url = "http://" + url + "/showTopicNews.html?topicid=" + topicid + "&modelid=100&modeltwoid="
    if act_type == "service":
        showTopicNews_url += "101"
    else:
        showTopicNews_url += "102"

    responce = requests.get(showTopicNews_url)

    bs = BeautifulSoup(responce.text, "lxml")
    date = bs.find("div", class_ = "m-tt").find("p").get_text().split(":")[1]
    title = bs.find("div", class_ = "m-text").find("h4").get_text()

    activity = {"title": title, "date": date}
    return activity

def check(agencies_list):
    for index, agency in enumerate(agencies_list):
        try:
            url = agency["url"]
            name = agency["name"]
            print u"正在写入 ",name,u" 的信息..."

            sheet.write(index + 1, 0, name) # 写入经销商名称

            service_activities_url = "http://" + url + "/subfrontfwhd.html?modelid=100&modeltwoid=101" # 服务活动页面
            market_activities_url = "http://" + url + "/subfrontfwhd.html?modelid=100&modeltwoid=102" # 市场活动页面

            service_activity_topicid = get_topicid(service_activities_url) # 获取top服务活动的topicid
            market_activity_topicid = get_topicid(market_activities_url) # 获取top市场活动的topicid

            service_activity =  get_activity(url, "service", service_activity_topicid)# 获取活动的标题与日期，封装在dict中
            market_activity =  get_activity(url, "market", market_activity_topicid)

            if service_activity == None:
                sheet.write(index + 1, 1, u"未添加")
                sheet.write(index + 1, 2, u"未添加")
            else:
                sheet.write(index + 1, 1, service_activity["title"])
                sheet.write(index + 1, 2, service_activity["date"])

            if market_activity == None:
                sheet.write(index + 1, 3, u"未添加")
                sheet.write(index + 1, 4, u"未添加")
            else:
                sheet.write(index + 1, 3, market_activity["title"])
                sheet.write(index + 1, 4, market_activity["date"])

            print u"写入 ", name, u" 的信息完毕!"
        except requests.exceptions.ConnectionError:
            print agency["url"] + u"可能无法访问，请手动尝试检查！"
            sheet.write(index + 1, 1, u"url链接无法访问")
            sheet.write(index + 1, 2, u"url链接无法访问")
            sheet.write(index + 1, 3, u"url链接无法访问")
            sheet.write(index + 1, 4, u"url链接无法访问")
        except Exception,e:
            print agency["url"] + u"发生其他未知错误"
            sheet.write(index + 1, 1, u"其他未知错误")
            sheet.write(index + 1, 2, u"其他未知错误")
            sheet.write(index + 1, 3, u"其他未知错误")
            sheet.write(index + 1, 4, u"其他未知错误")
            print e


 # 先打开一个新的excel
print u"打开一个excel"
wb = xlsxwriter.Workbook(u"./经销商任务完成情况.xlsx")
sheet = wb.add_worksheet('sheet1')
# 写入标题
print u"写入标题中..."
sheet.write(0, 0, u'经销商')
sheet.write(0, 1, u'服务活动标题')
sheet.write(0, 2, u'服务活动发布日期')
sheet.write(0, 3, u'市场活动标题')
sheet.write(0, 4, u'市场活动发布日期')
print u"写入标题完毕!"

check(get_agencies())
wb.close()
print u'写入完毕，excel文件已生成！'