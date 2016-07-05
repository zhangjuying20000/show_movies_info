import re
import zlib
import queue
import urllib
import pymysql
import datetime
import xlsxwriter
import http.client
import urllib.request
from bs4 import BeautifulSoup

class Crawler():

    def __init__(self):
        self.conn = pymysql.connect(user='root', passwd='root', host='localhost', port=3306, db='movies', charset='UTF8')
        self.cursor = self.conn.cursor()
        self.carShouYeUrlSet = []

    def __del__(self):
        self.cursor.close()
        self.conn.close()

    #获取url 对应 HTML 源码
    def get_html(self, url):
        request = urllib.request.Request(url)

        try: #代理：110.81.238.173:8088
            proxy_support = urllib.request.ProxyHandler({'http': '110.81.238.173:8088'})
            opener = urllib.request.build_opener(proxy_support)
            urllib.request.install_opener(opener)

            page = urllib.request.urlopen(request)

            if page.headers.get('Content-Encoding') == 'gzip':
                return zlib.decompress(page.read(), 16+zlib.MAX_WBITS).decode('gbk', 'ignore')
            else:
                return page.read().decode(page.headers.get_content_charset(), 'ignore')

        except urllib.request.HTTPError as e:
            print('HTTPERROR: ', str(e))
            return urllib.request.HTTPError
        except http.client.HTTPException as e:
            print('http.client.HTTPException: ', str(e))
            return http.client.HTTPException

    #从电影详情页获取电影类型
    def get_movie_infos(self, url):

        html_doc = BeautifulSoup(self.get_html(url), 'html.parser')
        div_info = html_doc.select('#info')[0].getText().strip().split('\n')
        res_info = {}
        for div in div_info:
            div = div.split(': ')
            if div[0] == '导演':
                res_info.update({'导演' : ','.join(div[1:])})
            elif div[0] == '制片国家/地区':
                res_info.update({'制片国家/地区' : ','.join(div[1:])})
            elif div[0] == '语言':
                res_info.update({'语言' : ','.join(div[1:])})
            elif div[0] == '类型':
                res_info.update({'类型' : ','.join(div[1:])})
        return res_info

    #从豆瓣上获取我的观影信息
    def get_douban_infos(self, username):

        url_que = queue.Queue()
        url_que.put('https://movie.douban.com/people/' + username + '/collect?sort=time&amp;start=0&amp;filter=all&amp;mode=list&amp;tags_sort=count')
        while not url_que.empty():
            url = url_que.get()
            #爬取第一页
            html_doc = BeautifulSoup(self.get_html(url), 'html.parser')

            #获取下一页链接
            next_url = html_doc.select('#content > div.grid-16-8.clearfix > div.article > div.paginator > span.next > a')
            if len(next_url) > 0:
                next_url = next_url[0]['href']
                url_que.put(next_url)

            html_doc = BeautifulSoup(self.get_html(url), 'html.parser')
            li_lists = html_doc.select('#content > div')[1].select('div.article > ul > li')

            for li in li_lists:

                a_label = li.select('div.item-show > div > a')[0]

                movie_name = a_label.getText().strip()
                movie_name = movie_name.split('/')[0].strip()

                #去重
                query = 'SELECT COUNT(*) FROM douban_movie WHERE movie_name = (%s)'
                self.cursor.execute(query, (movie_name,))
                self.conn.commit()
                res = self.cursor.fetchall()
                if not res[0][0] == 0:
                    continue

                movie_detail_url = a_label['href']
                print(movie_detail_url)
                try:
                    movie_info = self.get_movie_infos(movie_detail_url)
                except TypeError as e:
                    print(str(e))
                    print(movie_name)
                    continue

                rate = li.select('div.item-show > div')[1].select('span')
                if len(rate) == 0:
                    rate = ''
                else:
                    rate = str(rate[0]['class'])
                    regx = re.compile('(\d)')
                    rate = regx.findall(rate)[0]

                watch_time = li.select('div.item-show > div')[1].getText().strip()

                comment = li.select('div.comment')
                #如果有评论
                if len(comment) > 0:
                    comment = comment[0].getText().strip()
                    if not comment.find('(1 有用)') == -1:
                        comment = comment.split(' ')[0].strip()
                else:
                    comment = ''

                #过滤emoij。。。
                regx = re.compile("[dc00-\udfff|\，]+")
                comment = ''.join(regx.findall(comment))
                print(movie_info)

                movie_types = movie_info['类型']
                movie_language = movie_info['语言']
                movie_country = movie_info['制片国家/地区']
                movie_director = movie_info['导演']

                print(comment)
                query = 'INSERT INTO douban_movie(movie_name, rate, watch_time, comment, movie_types, movie_language, movie_country, movie_director) VALUES (%s, %s, %s, %s, %s, %s, %s, %s)'
                self.cursor.execute(query, (movie_name, rate, watch_time, comment, movie_types, movie_language, movie_country, movie_director))
                self.conn.commit()
        return

    #将电影类型分类展示在柱状图中并统计数量
    def show_types_in_column(self, workbook):

        worksheet = workbook.add_worksheet()

        chart = workbook.add_chart({'type' : 'column'})

        query = 'SELECT movie_name, movie_types FROM douban_movie'
        self.cursor.execute(query)
        self.conn.commit()
        movies_lists = self.cursor.fetchall()

        movie_types_dict = {}

        for movie in movies_lists:
            movie_name, movie_types = movie
            types = movie_types.split('/')
            for type in types:
                type = type.strip()
                if type in movie_types_dict:
                    movie_types_dict[type] += 1
                else:
                    movie_types_dict.update({type : 1})

        # Add table data
        table_data = [['类型', '片数'],]
        for key, val in movie_types_dict.items():
            table_data.append([key, val])

        type = '类型'
        num = '数目'
        worksheet.write_column('A1', type)
        worksheet.write_column('B1', num)
        worksheet.write_column('A2', movie_types_dict.keys())
        worksheet.write_column('B2', movie_types_dict.values())

        chart.add_series({'name' : '影片数目',
                          'categories' : ['Sheet1', 1, 0, 1+len(movie_types_dict.keys()), 0],
                          'values' : ['Sheet1', 1, 1, 1+len(movie_types_dict.values()), 1],
                          'color' : 'red',
                          'data_labels' : {'value' : True}})

        #设置X轴属性
        chart.set_x_axis({
            'name' : '影片类型',
            'name_font' : {'size' : 14, 'bold' : False},
            'num_font' : {'size' : 12},
            'line' : {'none' : True},
            'major_gridlines' : {
                'visible' : True,
                'line' : {'width' : 1.5, 'dash_type' : 'dash'}
            },
            'text_axis' : True
        })

        #设置长宽
        chart.set_size({
            'x_scale' : 2.5,
            'y_scale' : 2
        })

        #设置标题
        chart.set_title({
            'name' : '15/07/01 - 16/06/30 观影类型统计'
        })

        #设置属性块
        chart.set_legend({
            'position' : 'left'
        })

        #图下方显示表格
        chart.set_table({
            'show_keys' : True,
            'font' : {'size' : 12}
        })

        worksheet.insert_chart('C1', chart)

    #将电影语言分类展示在饼状图中
    def show_languages_in_pie(self, workbook):

        worksheet = workbook.add_worksheet()

        chart = workbook.add_chart({'type' : 'pie'})

        query = 'SELECT movie_name, movie_language FROM douban_movie'
        self.cursor.execute(query)
        self.conn.commit()
        movies_lists = self.cursor.fetchall()

        movie_languages_dict = dict({})

        for movie in movies_lists:
            movie_name, movie_languages = movie
            languages = movie_languages.split('/')

            for language in languages:
                language = language.strip()
                if language in movie_languages_dict:
                    movie_languages_dict[language] += 1
                else:
                    movie_languages_dict.update({language : 1})

        total_sum = 0
        for key, val in movie_languages_dict.items():
            total_sum += val

        percent_lists = []
        # Add table data
        table_data = [['语言', '数目', '百分比'],]
        for key, val in movie_languages_dict.items():
            movie_languages_dict[key] = val
            tmp = format(val/total_sum, '.2%')
            percent_lists.append(tmp)
            table_data.append([key, val, tmp])

        language = '语言'
        num = '数目'
        percent = '百分比'

        worksheet.write_column('A1', language)
        worksheet.write_column('B1', num)
        worksheet.write_column('C1', percent)

        worksheet.write_column('A2', movie_languages_dict.keys())
        worksheet.write_column('B2', movie_languages_dict.values())
        worksheet.write_column('C2', percent_lists)

        chart.add_series({'name' : '影片语言',
                          'categories' : ['Sheet2', 1, 0, 1+len(movie_languages_dict.keys()), 0],
                          'values' : ['Sheet2', 1, 1, 1+len(movie_languages_dict.values()), 1],
                          'color' : 'red',
                          'data_labels' : {'percentage' : True}})

        #设置长宽
        chart.set_size({
            'x_scale' : 1.5,
            'y_scale' : 2
        })

        #设置标题
        chart.set_title({
            'name' : '15/07/01 - 16/06/30 观影语言统计'
        })

        #设置属性块
        chart.set_legend({
            'position' : 'bottom'
        })

        worksheet.insert_chart('D1', chart)

    #统计每个月观看电影数量并绘制趋势图
    def show_monthly_watched_movies_num(self, workbook):
        worksheet = workbook.add_worksheet()

        chart = workbook.add_chart({'type' : 'line'})

        query = 'SELECT movie_name, watch_time FROM douban_movie'
        self.cursor.execute(query)
        self.conn.commit()
        movies_lists = self.cursor.fetchall()

        movie_monthly_dict = dict({})

        for movie in movies_lists:
            movie_name, watch_time = movie
            watch_time = datetime.datetime.strptime(str(watch_time), "%Y-%m-%d %H:%M:%S")
            key = str(watch_time.year) + '-' + str(watch_time.month)
            if key in movie_monthly_dict:
                movie_monthly_dict[key] += 1
            else:
                movie_monthly_dict.update({key : 1})

        movie_monthly_list = sorted(movie_monthly_dict.items(), key=lambda d: d[0])

        time_list = []
        num_list = []

        # Add table data
        table_data = [['时间', '数目'],]
        for key, val in movie_monthly_list:
            table_data.append([key, val])
            time_list.append(key)
            num_list.append(val)

        watch_time = '时间'
        num = '数目'

        worksheet.write_column('A1', watch_time)
        worksheet.write_column('B1', num)

        worksheet.write_column('A2', time_list)
        worksheet.write_column('B2', num_list)

        chart.add_series({'name' : '观影数目',
                          'categories' : ['Sheet3', 1, 0, 1+len(movie_monthly_dict.keys()), 0],
                          'values' : ['Sheet3', 1, 1, 1+len(movie_monthly_dict.values()), 1],
                          'color' : 'red',
                          'marker' : {'type' : 'diamond'},
                          'smooth' : 'True',
                          'data_labels' : {'value' : True}})

        #设置长宽
        chart.set_size({
            'x_scale' : 1.5,
            'y_scale' : 2
        })

        #设置标题
        chart.set_title({
            'name' : '15/07/01 - 16/06/30 观影时间统计'
        })

        #设置属性块
        chart.set_legend({
            'position' : 'bottom'
        })

        #图下方显示表格
        chart.set_table({
            'show_keys' : True,
            'font' : {'size' : 12}
        })

        worksheet.insert_chart('C1', chart)

    #数据可视化展示
    def show_datas(self):

        workbook = xlsxwriter.Workbook('G:\\chart.xlsx')

        #统计分类信息
        self.show_types_in_column(workbook)

        #统计语言信息
        self.show_languages_in_pie(workbook)

        #统计月观影片数
        self.show_monthly_watched_movies_num(workbook)

        workbook.close()

    def execute_vba(self):
        import pychart_into_img
        pychart_into_img.main()

def main():
    crawler = Crawler()

    crawler.get_douban_infos('yinwoods')

    crawler.show_datas()

    crawler.execute_vba()

if __name__ == "__main__":
    main()

