# -*- coding: utf-8 -*-
# @Time : 2020/8/26 22:59
# @Author : C4C
# @File : input.py
# @Remark : input douban link
import random
import requests as rq
import xlwt

info_list = []

def get_info(videoname,link_subject,header):

    url = 'https://ptgen.rhilip.workers.dev/?url='+link_subject
    print(url)
    reponse_detail = rq.get(url=url,headers=header,timeout=30)
    #print(reponse_detail.status_code)
    if(reponse_detail.status_code == 200):
        json = reponse_detail.json()

        name = videoname
        try:
            translated = json['aka']
            str_translated = tostr2(translated)
        except:
            str_translated = videoname
        poster = json['poster']
        year = json['year']
        region = json['region']
        str_region = tostr2(region)
        genre = json['genre']
        str_genre = tostr2(genre)
        language = json['language']
        str_language = tostr2(language)
        playdate = json['playdate']
        str_playdate =tostr2(playdate)
        douban_rating = json['douban_rating']
        douban_link = link_subject
        try:
            imdb_link = json['imdb_link']
        except:
            imdb_link = 'Null'
        episodes = json['episodes']
        duration = json['duration']
        director = json['director']
        str_director = ''
        for i in range(len(director)):
            if(i + 1 == len(director)):
                str_director += director[i]['name']
            else:
                str_director += director[i]['name'] + '/'
        writer = json['writer']
        str_writer = ''
        for i in range(len(writer)):
            if (i + 1 == len(writer)):
                str_writer+= writer[i]['name']
            else:
                str_writer += writer[i]['name'] + '/'
        cast = json['cast']
        str_cast = ''
        for i in range(len(cast)):
            if (i + 1 == len(cast)):
                str_cast += cast[i]['name']
            else:
                str_cast += cast[i]['name'] + '/'
        tags = json['tags']
        str_tags = tostr2(tags)
        introduction = json['introduction'].replace('\n','<br>')

        data = {
            '片名':name,
            '译名':str_translated,
            '海报':poster,
            '年代':year,
            '国家':str_region,
            '类别':str_genre,
            '语言':str_language,
            '上映日期':str_playdate,
            '豆瓣评分':douban_rating,
            '豆瓣链接':douban_link,
            'IMDB链接':imdb_link,
            '集数':episodes,
            '片长':duration,
            '导演':str_director,
            '编剧':str_writer,
            '主演':str_cast,
            '标签':str_tags,
            '简介':introduction,
        }
        info_list.append(data)
        print('信息获取成功！')
        #print(data)
    else:
        error(videoname)
        print('网络错误！(API)'+str(reponse_detail.status_code))

def tostr1(list):
    str = ''
    for i in list:
        str += i
    return str

def tostr2(list):
    str = ''
    for i in range(len(list)):
        if(i+1 == len(list)):
            str += list[i]
        else:
            str += list[i]+'/'
    return str

def error(videoname):
    data = {
        '片名': videoname,
        '译名': 'Null',
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

def write_excel():
    work_book = xlwt.Workbook(encoding='utf-8',style_compression=0)
    sheet = work_book.add_sheet('sheet1',cell_overwrite_ok=True)
    thead = ['name','county','time','mainstar','img','douban_link','imdb_link','genre','douban_rating','info']
    sheet.write(0, 0, thead[0])
    sheet.write(0, 1, thead[1])
    sheet.write(0, 2, thead[2])
    sheet.write(0, 3, thead[3])
    sheet.write(0, 8, thead[4])
    sheet.write(0, 9, thead[5])
    sheet.write(0, 10, thead[6])
    sheet.write(0, 11, thead[7])
    sheet.write(0, 12, thead[8])
    sheet.write(0, 13, thead[9])
    for i in range(len(info_list)):
        name = info_list[i]['片名']
        translated = info_list[i]['译名']
        poster = info_list[i]['海报']
        year = info_list[i]['年代']
        region = info_list[i]['国家']
        genre = info_list[i]['类别']
        language = info_list[i]['语言']
        playdate = info_list[i]['上映日期']
        douban_rating = info_list[i]['豆瓣评分']
        douban_link = info_list[i]['豆瓣链接']
        imdb_link = info_list[i]['IMDB链接']
        episodes = info_list[i]['集数']
        duration = info_list[i]['片长']
        director = info_list[i]['导演']
        writer = info_list[i]['编剧']
        cast  = info_list[i]['主演']
        tags = info_list[i]['标签']
        introduction = info_list[i]['简介']
    #组合信息
        packing_info = '◎译　　名　' + str(translated) + '<br>' \
        '◎年　　代　' + str(year) + '<br>'\
        '◎产　　地　' + str(region) + '<br>'\
        '◎类　　别　' + str(genre) + '<br>'\
        '◎语　　言　' + str(language) + '<br>'\
        '◎上映日期　' + str(playdate) + '<br>'\
        '◎IMDb链接 ' + str(imdb_link) + '<br>'\
        '◎豆瓣评分　' + str(douban_rating) + '<br>'\
        '◎豆瓣链接　' + str(douban_link) + '<br>'\
        '◎集　　数　' + str(episodes) + '<br>'\
        '◎片　　长　' + str(duration) + '<br>'\
        '◎导　　演　' + str(director) + '<br>'\
        '◎编　　剧　' + str(writer) + '<br>'\
        '◎主　　演　' + str(cast) + '<br>'\
        '◎标　　签　' + str(tags) + '<br>'\
        '◎简　　介　' + str(introduction)

        '''['county','time','mainstar','img','douban_link','imdb_link','genre','douban_rating','info']'''
        sheet.write(i + 1, 0, name)
        sheet.write(i + 1, 1, region)
        sheet.write(i + 1, 2, playdate)
        sheet.write(i + 1, 3, cast)
        sheet.write(i + 1, 7, poster)
        sheet.write(i + 1, 8, douban_link)
        sheet.write(i + 1, 9, imdb_link)
        sheet.write(i + 1, 10, genre)
        sheet.write(i + 1, 11, douban_rating)
        sheet.write(i + 1, 12, packing_info)
        #print(packing_info)
    work_book.save('./info.xls')


if __name__ == '__main__':
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
    try:
        while(1):
            videoname = input('请输入影片名称：')
            link = input('请输入豆瓣影片连接：')
            if (link == 'e' or videoname == 'e'):
                try:
                    write_excel()
                    print('文件写入完成')
                except:
                    print('文件写入错误，请检查输出文件是否被占用！')
                exit(0)
            elif(link == 'null'):
                error(videoname)
            else:
                get_info(videoname,link,header)
    except:
        print('致命错误，程序提前终止！')
        try:
            write_excel()
            print('文件写入完成')
        except:
            print('文件写入错误，请检查输出文件是否被占用！')
        exit(3)

