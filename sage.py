#!/usr/bin/python
# coding=utf-8
import requests
from bs4 import BeautifulSoup
import re
import xlwt
# http://www.promoplace.com/ws/ws.dll/QPic?SN=50307&P=154968900&PX=400&ReqFrameSize=1&I=1
URL_ARR = ["p1.html","p2.html","p3.html","p4.html","p5.html","p6.html","p7.html","p8.html","p9.html","p10.html","p11.html","p12.html","p13.html"]
BASE_URL = '/Users/apple/Desktop/'
LOCAL_PATH = r'/Users/apple/Desktop'
BASE_IMG_URL = "http://www.promoplace.com"
links = []
pic_urls = []
pic_names = []
def get_localhtml() :
    print "step 1...."
    for url in URL_ARR:
        url = BASE_URL + url
        print url
        htmlfile = open (url,'r')
        htmlpage = htmlfile.read()
        soup = BeautifulSoup(htmlpage, "html.parser")
        get_hrefs(soup)
    print "+++++++++++++++++++++++",len(links),links
    get_iframe_link(links)

def get_hrefs(soup):
    print "step 2...."
    result = soup.find_all("a",href=re.compile("http://www.checkmyproducts.com/p/"))
    for a in result:
        links.append(a.get('href'))

def get_iframe_link(links):
    print "step 3...."
    realLinks = []
    count = 0
    for url in links:
        count = count +1
        print count,"*****************step 3*****************",url
        html = requests.get(url)
        soup = BeautifulSoup(html.content, "html.parser")
        result = soup.find_all('iframe')
        for iframe in result:
            realLinks.append(iframe.get('src'))

    get_detail_inIframe(realLinks)

def get_detail_inIframe(links):
    print "step 4...."
    dataArr =[]
    for url in links:
        detailsArr = get_single_detail(url)
        dataArr.append(detailsArr)
        print "*****************step 4*****************️",len(dataArr)
    write_excel_file(dataArr)

def get_single_detail(url):
    picname = "NO image"
    html = requests.get(url)
    soup = BeautifulSoup(html.content, "html.parser")

    detailsArr = []

    pr_name = soup.find_all("h1")[0].string
    detailsArr.append(pr_name)

    pr_item = soup.find_all("p", class_="item-numb")

    for item in pr_item:
        detailsArr.append(item.text)

    pr_desc = soup.find_all("p", class_="item-desc")[0].string
    detailsArr.append(pr_desc)

    setup = soup.find_all("small")
    index = 0
    for small in setup:
        index += 1
        if index !=1:
            detailsArr.append(small.text)

    body = soup.find_all("div", class_="panel-body")
    count = 0
    longp = ""

    for panel in body:
        for p in panel.find_all("p"):
            str = p.text
            if count < 3 :
                detailsArr.append(str)
                count = count + 1

            else:
                longp = longp + p.text+" "
        detailsArr.append(longp)

    imgs = soup.find_all("img", src=re.compile("/ws/ws.dll/"))
    # 拼接路径
    if len(imgs)!=0:
        img = imgs[0]
        src = img.get("src")
        imgCount = len(imgs) / 2
        pretext = ["1","2","3","4","5"];
        for i in range(1, imgCount + 1):
            picname = pr_item[0].text
            picname = picname.replace(' ', '')
            imgRoute = BASE_IMG_URL + src[0:36]
            imgRoute = imgRoute + "&PX=400&ReqFrameSize=1&I=" + pretext[i-1]
            picname = picname+"_"+pretext[i]
            pic_urls.append(imgRoute)
            pic_names.append(picname)
    detailsArr.append(picname)
    # print "---------------",picname

    table = soup.find_all("table", class_="table rwd-table")
    for tab in table:
        tds = tab.find_all("td")
        for td in tds:
            detailsArr.append(td.get("data-th"))
            detailsArr.append(td.string)

    return detailsArr


def write_excel_file(datas) :
    print "step 5...."
    book = xlwt.Workbook(encoding='utf-8',style_compression=0)
    sheet = book.add_sheet('SAGE',cell_overwrite_ok=True)
    sheet.write(0, 0, "Name")
    sheet.write(0, 1, "Item #")
    sheet.write(0, 2, "SAGE #")
    sheet.write(0, 3, "Description")
    sheet.write(0, 4, "EEE")
    sheet.write(0, 5, "Setup")
    sheet.write(0, 6, "Colors")
    sheet.write(0, 7, "Themes")
    sheet.write(0, 8, "Imprint Information")
    sheet.write(0, 9, "Delivery Information")
    sheet.write(0, 10, "Picture Name")
    sheet.write(0, 11, "Quantity")
    sheet.write(0, 12, "Price")


    for j in range(0,len(datas)+1):
        print j,"❤️❤️❤️❤️❤️❤️❤️❤️❤️❤❤️❤️❤️❤️❤️❤️❤️❤️❤️❤️❤️❤️step 5❤️❤️❤️❤️❤️❤️❤️❤️❤️❤️❤️❤️❤️❤️❤️❤️❤️❤️",len(datas)
        data = datas[j-1]
        for i in range(0,len(data)):
            if j==0 and i>10 :
                sheet.write(0,i,"prices && quantity")

            if j !=0:
                sheet.write(j,i,data[i])

    book.save(LOCAL_PATH+'/sage.xls')

    download_imgs()

    print "Success!!"

def download_imgs():
    print "step 6...."
    for i in range(0,len(pic_urls)):
        each = pic_urls[i]
        name = pic_names[i]
        print "正在下载第", i, "张图片", "图片地址：",each
        try:
            pic = requests.get(each,timeout =15)
        except requests.exceptions.ConnectionError:
            print '【错误】当前图片无法下载'
            continue
        savepath = LOCAL_PATH + '/'+'sageimages/'+name+'.jpg'
        fp = open(savepath,'wb')
        fp.write(pic.content)


if __name__ == "__main__":

    get_localhtml()
    #
    # arr = get_single_detail("http://www.promoplace.com/ws/ws.dll/PrDtl?UID=9002&SPC=ayrna-kwslo")
    # print arr