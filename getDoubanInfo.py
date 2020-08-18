# -*- coding: utf-8 -*-
# @Time : 2020/8/13 23:02
# @Author : C4C
# @File : getDoubanInfo.py
# @Remark : 获取豆瓣电影信息
import random
import re
import time
import xlrd
import xlwt
import requests as rq
from bs4 import BeautifulSoup as bs

douban_apikey_list = [
    "02646d3fb69a52ff072d47bf23cef8fd",
    "0b2bdeda43b5688921839c8ecb20399b",
    "0dad551ec0f84ed02907ff5c42e8ec70",
    "0df993c66c0c636e29ecbb5344252a4a"
]
video_list =[]
time_list = []
info_list = []

def get_db_apikey() -> str:
    return random.choice(douban_apikey_list)

def get_videoname(file):
    book_read = xlrd.open_workbook(file)#'D:/python/pytest/1.xlsx'
    table = book_read.sheet_by_index(0)
    video_list = table.col_values(1)
    del video_list[0]
    time_list = table.col_values(5)
    del time_list[0]

    return video_list,time_list


def get_douban_link(video_list,time_list,header):
    if(len(video_list) == len(time_list)):
        for i in range(len(video_list)):
            videoname = video_list[i]
            videotime = str(time_list[i])[:-2]
            search_url = "https://www.douban.com/search?cat=1002&q=" + videoname
            print(search_url)
            print('进度  '+ str(i+1) + '/' + str(len(video_list)))

            response = rq.get(url=search_url, headers=header, timeout=30)
            time.sleep(3)
            # print(respones.status_code)
            if (response.status_code == 200):
                response.encoding = 'utf-8'
                page_html = response.text
                soup = bs(page_html, 'lxml')  # 返回网页
                # link
                title_class = soup.select_one('.title')
                link = title_class.select('a')[0].get('href')
                # 获取跳转连接
                response1 = rq.get(url=link, headers=header, timeout=30)
                time.sleep(2)
                link_subject = response1.url[:-1];
                # info
                info_class = soup.select_one('.rating-info')
                info = info_class.select('span')[3].text
                # 判断连接是否正确
                searchObj1 = re.search(videotime, info, re.I)
                if searchObj1:
                    try:
                        get_info(link_subject,header)
                    except:
                        print('获取信息时发生错误，请人工检查！')
                else:
                    data = {
                        '海报': 'Null',
                        '年代': 'Null',
                        '国家': 'Null',
                        '类别': 'Null',
                        '语言': 'Null',
                        '上映日期': 'Null',
                        '豆瓣评分': 'Null',
                        '豆瓣链接': 'Null',
                        'IMDB链接': 'Null',
                        '集数': 'Null',
                        '片长': 'Null',
                        '导演': 'Null',
                        '编剧': 'Null',
                        '主演': 'Null',
                        '标签': 'Null',
                        '简介': 'Null',
                    }
                    info_list.append(data)
                    print('年份未匹配，请人工确认！')

    else:
        print('名称与年份数量不匹配，请检查原表！')


def get_info(link_subject,header):

    url = link_subject
    print(url)
    reponse_detail = rq.get(url=url,headers=header,timeout=30)
    time.sleep(3)
    #print(reponse_detail.status_code)
    if(reponse_detail.status_code == 200):
        reponse_detail.encoding = 'utf-8'
        detail_page_html = reponse_detail.text
        detail_soup = bs(detail_page_html,'lxml')

        region_anchor = detail_soup.find("span", class_="pl", text=re.compile("制片国家/地区"))
        language_anchor = detail_soup.find("span", class_="pl", text=re.compile("语言"))
        episodes_anchor = detail_soup.find("span", class_="pl", text=re.compile("集数"))
        imdb_link_anchor = detail_soup.find("a", text=re.compile("tt\d+"))
        year_anchor = detail_soup.find("span", class_="year")

        year = detail_soup.find("span", class_="year").text[1:-1] if year_anchor else ""  # 年代
        region = fetch(region_anchor).split(" / ") if region_anchor else []  # 产地
        str_region = tostr(region)
        genre = list(map(lambda l: l.text.strip(), detail_soup.find_all("span", property="v:genre")))  # 类别
        str_genre = tostr2(genre)
        language = fetch(language_anchor).split(" / ") if language_anchor else []  # 语言
        str_language = tostr2(language)
        playdate = sorted(map(lambda l: l.text.strip(),  # 上映日期
                                      detail_soup.find_all("span", property="v:initialReleaseDate")))
        str_playdate  = tostr(playdate)
        imdb_link = imdb_link_anchor.attrs["href"] if imdb_link_anchor else ""  # IMDb链接
        str_imdb_link = tostr(imdb_link)
        imdb_id = imdb_link_anchor.text if imdb_link_anchor else ""  # IMDb号
        str_imdb_id = tostr(imdb_id)
        episodes = fetch(episodes_anchor) if episodes_anchor else ""  # 集数
        str_episodes = tostr(episodes)
#       获取片长
        duration_anchor = detail_soup.find("span", class_="pl", text=re.compile("单集片长"))
        runtime_anchor = detail_soup.find("span", property="v:runtime")

        duration = ""  # 片长
        if duration_anchor:
            duration = fetch(duration_anchor)
        elif runtime_anchor:
            duration = runtime_anchor.text.strip()

        douban_api_json = douban_api(url,header)
        #print(douban_api_json)
        # 豆瓣评分，简介，海报，导演，编剧，演员，标签
        douban_average_rating = douban_api_json["rating"]["average"] or 0
        douban_votes = douban_api_json["rating"]["numRaters"] or 0
        douban_rating = "{}/10 from {} users".format(douban_average_rating, douban_votes)#评分
        introduction = re.sub("^None$", "暂无相关剧情介绍", douban_api_json["summary"])#简介
        poster= poster = re.sub("s(_ratio_poster|pic)", r"l\1", douban_api_json["image"])#海报

        director = douban_api_json["attrs"]["director"] if "director" in douban_api_json["attrs"] else []#导演
        str_director = tostr2(director)
        writer = douban_api_json["attrs"]["writer"] if "writer" in douban_api_json["attrs"] else []#编剧
        str_writer = tostr2(writer)
        cast = douban_api_json["attrs"]["cast"] if "cast" in douban_api_json["attrs"] else ""#主演
        str_cast = tostr2(cast)
        tags = list(map(lambda member: member["name"], douban_api_json["tags"]))#标签
        str_tags = tostr2(tags)

        data = {
            '海报':poster,
            '年代':year,
            '国家':str_region,
            '类别':str_genre,
            '语言':str_language,
            '上映日期':str_playdate,
            '豆瓣评分':douban_rating,
            '豆瓣链接':url,
            'IMDB链接':str_imdb_link,
            '集数':str_episodes,
            '片长':duration,
            '导演':str_director,
            '编剧':str_writer,
            '主演':str_cast,
            '标签':str_tags,
            '简介':introduction,
        }
        info_list.append(data)
        #print(data)
    else:
        print('网络错误！')

def fetch(node):
    return node.next_element.next_element.strip()



def tostr(list):
    str = ''
    for i in list:
        str += i
    return str

def tostr2(list):
    str = ''
    for i in list:
        str += i+'/'
    return str

def douban_api(url,header):
    url1 = 'https://api.douban.com/v2/movie/'+url.rsplit('/',1)[1]+'?apikey='+random.choice(douban_apikey_list)
    print(url1)
    reponse_douban_api = rq.get(url=url1, headers=header, timeout=30)
    time.sleep(3)
    #print(reponse_douban_api.status_code)
    if (reponse_douban_api.status_code == 200):
        reponse_douban_api.encoding = 'utf-8'
        page_html = reponse_douban_api.json()
    return page_html

def write_excel():
    work_book = xlwt.Workbook(encoding='utf-8',style_compression=0)
    sheet = work_book.add_sheet('sheet1',cell_overwrite_ok=True)
    thead = ['county','time','mainstar','img','douban_link','imdb_link','genre','douban_rating','info']
    sheet.write(0, 1, thead[0])
    sheet.write(0, 2, thead[1])
    sheet.write(0, 3, thead[2])
    sheet.write(0, 8, thead[3])
    sheet.write(0, 9, thead[4])
    sheet.write(0, 10, thead[5])
    sheet.write(0, 11, thead[6])
    sheet.write(0, 12, thead[7])
    sheet.write(0, 13, thead[8])
    for i in range(len(info_list)):
        poster = info_list[i]['海报']
        year = info_list[i]['年代']
        region = info_list[i]['国家']
        genre = info_list[i]['类别']
        language = info_list[i]['语言']
        playdate = info_list[i]['上映日期']
        douban_rating = info_list[i]['豆瓣评分']
        url = info_list[i]['豆瓣链接']
        imdb_link = info_list[i]['IMDB链接']
        episodes = info_list[i]['集数']
        duration = info_list[i]['片长']
        director = info_list[i]['导演']
        writer = info_list[i]['编剧']
        cast  = info_list[i]['主演']
        tags = info_list[i]['标签']
        introduction = info_list[i]['简介']
    #组合信息
        packing_info = '◎年　　代　' + year + '<br>'\
        '◎产　　地　' + region + '<br>'\
        '◎类　　别　' + genre + '<br>'\
        '◎语　　言　' + language + '<br>'\
        '◎上映日期　' + playdate + '<br>'\
        '◎IMDb链接 ' + imdb_link + '<br>'\
        '◎豆瓣评分　' + douban_rating + '<br>'\
        '◎豆瓣链接　' + url + '<br>'\
        '◎集　　数　' + episodes + '<br>'\
        '◎片　　长　' + duration + '<br>'\
        '◎导　　演　' + director + '<br>'\
        '◎编　　剧　' + writer + '<br>'\
        '◎主　　演　' + cast + '<br>'\
        '◎标　　签　' + tags + '<br>'\
        '◎简　　介　' + introduction + '<br>'

        '''['county','time','mainstar','img','douban_link','imdb_link','genre','douban_rating','info']'''
        sheet.write(i + 1, 1, region)
        sheet.write(i + 1, 2, playdate)
        sheet.write(i + 1, 3, cast)
        sheet.write(i + 1, 8, poster)
        sheet.write(i + 1, 9, url)
        sheet.write(i + 1, 10, imdb_link)
        sheet.write(i + 1, 11, genre)
        sheet.write(i + 1, 12, douban_rating)
        sheet.write(i + 1, 13, packing_info)
        #print(packing_info)
    work_book.save('./info.xls')


if __name__ == '__main__':
    flag = 'n'
    # 定义随机header
    header1 = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/84.0.4147.125 Safari/537.36'
    }
    header2 = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:79.0) Gecko/20100101 Firefox/79.0'
    }
    header3 = {
        'User-Agent': 'Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; WOW64; Trident/5.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; Media Center PC 6.0; .NET4.0C; .NET4.0E; QQBrowser/7.0.3698.400)'
    }
    header4 = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 5.1) AppleWebKit/535.11 (KHTML, like Gecko) Chrome/17.0.963.84 Safari/535.11 SE 2.X MetaSr 1.0'
    }
    header5 = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/38.0.2125.122 UBrowser/4.0.3214.0 Safari/537.36'
    }
    headList = [header1, header2, header3, header4, header5]
    headerindex = random.randrange(0, len(headList))
    header = headList[headerindex]

    print('豆瓣影视作品信息获取    Create By C4C  ')
    flag = input(' 是否使用默认源数据文件(./1.xlsx)？\n y确定,n自定义路径,其他则提前结束程序  ')
    if(flag == 'y'):
        file = './1.xlsx'
    elif(flag == 'n'):
        file = input('请输入源数据Excel路径: ')
    else:
        exit(-1)

    video_list = get_videoname(file)[0]
    time_list = get_videoname(file)[1]
    get_douban_link(video_list,time_list,header)
    try:
        write_excel()
        print('文件写入完成')
    except:
        print('文件写入错误，请检查输出文件是否被占用！')
    input('按任意键结束程序！')