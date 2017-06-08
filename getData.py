#coding=utf-8
import requests
import csv
import random
import time
import socket
# import http.client
# import urllib.request
from bs4 import BeautifulSoup
import httplib, urllib
import urllib2
import cookielib
import sys
import re
import xlrd
import xlwt
import os
import requests
import cookielib
from selenium import webdriver
sum=1
browser = webdriver.Firefox(executable_path="geckodriver.exe")
workbook = xlwt.Workbook()
def firstget(url):
    global browser
    out1={ u'domain':u'www.factual.com',u'expiry':2147385600, u'httpOnly':True ,u'secure': False ,u'name':u'_www_session',
           u'value':u'ekRNZHU2YkxUK3JiNTlJcEhWWGs5czBHY2ZiZHlITnUwOU1yWmp2dXJiVllVRXJLcjBUQjdycERGYzM3SzRnek5lY01pbkI5eWNjUnFQQy9FbUFkUDIxcG9qWFBnd3lldktoaDVGUFgwYk9ocy93NDRvVi91S0VsTS91aDRYZ1dVbEw4Um9VMjNHVlBNUjcxS1pmcVR2R1c1djFiWEJtdStzcnUydDVmdnVFNjhneTBlUkRxUjhzcGZvMjdLNDBtWjhqeWJUa3FqVGYwZXc2eEZGTUIrMHRzdjNnVHhWdVlMQmZZaDdmRTR5WFlNR3VVcFUxTlI5VXpUUFg3dkt6MnVoMHVhWTd4Y3hvWStiMVJDQ2p2NWtBVFBwc1lTd1FuYXlwTC9mM09FSFZPSGZYSGxCZjY5eGtEbmxyY1N6c01FMU5yWXd5RUR4QllxOVBFcWUzMlZkOTQ0ekVFbysvd2JnVEp3OFFzNUNVPS0tODl6N0hkYXpMS3FtbUw1VzJsditvUT09--bdc3c540306fc56cc4a7fdd92ce8541170797a9d' , u'path':u'/'}
    browser.get(url)
    ck=browser.get_cookies()
    print ck
    browser.delete_all_cookies()
    browser.add_cookie(out1)
    time.sleep(3)
    ck=browser.get_cookies()
    print ck
    browser.refresh()
    ck=browser.get_cookies()
    print ck
    time.sleep(3)
    #print browser.page_source.encode('gbk', 'ignore')  # 这个函数获取页面的html
    browser.get_screenshot_as_file(""+str(sum)+".jpg")  # 获取页面截图
    print "Success To Create the screenshot & gather html"
    return browser.page_source.encode('gbk', 'ignore')

def get_content(url):
    print url
    browser.get(url)
    time.sleep(3)
    browser.get_screenshot_as_file(str(sum)+".jpg")  # 获取页面截图
    print "get success"
    return browser.page_source.encode('gbk', 'ignore')


def get_data(html_text,sheet):
    global sum
    bs = BeautifulSoup(html_text, "html.parser")  # 创建BeautifulSoup对象
    body=bs.body
    # a = body.find_all('div',{'class':'container-fluid'}) # 获取body部分
    # b = a[0].contents[1]
    # c = b.find('div',{'class':'col-sm-8'})
    # grid = c.contents[1]
    # gridcontext = grid.find('div',{'id':'data-grid-container'})
    # e = gridcontext[0].contents[0]
    # f = e.contents[3]
    #datacanvas = f.find_all('div',{'class':'grid-canvas'})
    datacanvas=body.find('div',{'class':'grid-canvas'})
    for data in datacanvas.children:    #data=每个酒店
        sheet.write(sum, 0, data.contents[0].contents[0].string)    #写第一个属性link
        print data.contents[1].string
        for i in range(1,len(data.contents)):    #simple单个酒店下各属性
            sheet.write(sum,i,data.contents[i-1].string)
        sum=sum+1

def set_style(name, height, bold=False):
    style = xlwt.XFStyle()  # 初始化样式

    font = xlwt.Font()  # 为样式创建字体
    font.name = name  # 'Times New Roman'
    font.bold = bold
    font.color_index = 4
    font.height = height

    # borders= xlwt.Borders()
    # borders.left= 6
    # borders.right= 6
    # borders.top= 6
    # borders.bottom= 6
    style.font = font
    # style.borders = borders
    return style

def initxls():
    global workbook
    systemdate=time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time()))
    systime=str(systemdate)
    systime=systime.replace(":","-")
    sheet = workbook.add_sheet('Data sheet in')
    row0 = [u'Link & ID',u'Name', u'Address', u'Locality', u'Region', u'Post', u'Latitude', u'Longitude',u'Air Con',u'Complimentary Breakfast', u'Low Price',u'High Price', u'Deposit',u'Room Count',u'Stars',u'Pets',u'Non Smoking Rooms',u'Smoking?',u'Internet',u'Pool',u'Fitness',u'Check in and out',u'Business center',u'Express Check in',u'Express Check out',u'Laundry',u'Cable TV',u'Room Service',u'Accessibility',u'Spa Service',u'Cribs',u'Restaurant',u'Concierge',u'Bar',u'24hr Desk',u'Meeting Rooms',u'Banquet Facilities',u'Event Catering',u'Complimentary Newspapers',u'Review Count',u'Rating',u'Type',u'Roll Out Beds',u'Neighborhood',u'Category Labels',u'Category IDs',u'Country',u'Parking',u'Address Extended']
    # 生成第一行
    for i in range(0, len(row0)):
        sheet.write(0, i, row0[i], set_style('Times New Roman', 220, True))
    return sheet

#xls文件会创建在该程序的文件夹下
#创建同时每一页面生成一张jpg截图
#运行后弹出firefox浏览器，请不要操作浏览器
if __name__ == '__main__':  #主函数，只需要更改这个函数的变量
    url ='https://www.factual.com/data/t/hotels-us#filters={"$and":[{"stars":{"$gt":"4"}}]}'    #将地址替换成你要检索的条件的地址（filters后面是过滤条件，更改条件后会有变化）
    sheet = initxls()
    fp=firstget(url)
    get_data(fp,sheet)
    pagelink=20
    max=1130/20+1   #总数据数除上单页显示数=总页数
    print "ready to Iteration data"
    for i in range(1,max+1):    #range为扫描所有页面，如果仅有3页有效数据，把max+1改为4
        print str(pagelink)
        newurl=url+"&offset="+str(pagelink*i)
        html = get_content(newurl)
        get_data(html,sheet)
        print "next"+str(pagelink)
    workbook.save('HotalData.xls')  #保存的xls名字