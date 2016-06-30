import re
import zlib
import urllib
import xlsxwriter
import http.client
import urllib.request
import mysql.connector
from bs4 import BeautifulSoup

class Crawler():

    def __init__(self):
        self.conn = mysql.connector.connect(user='root', password='root', host='localhost', database='movies')
        self.cursor = self.conn.cursor(buffered=True)
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

    #获取电影名并写入数据库movie中
    def get_movies(self, url_lists):

        for url in url_lists:
            html_doc = BeautifulSoup(self.get_html(url), 'html.parser')
            div_lists = html_doc.select('#content > div > div > a > div')
            for div in div_lists:

                movie_name = div.getText()

                #查重
                query = 'SELECT COUNT(*) FROM movie WHERE movie_name = %s'
                self.cursor.execute(query, (movie_name,))
                self.conn.commit()
                res = self.cursor.fetchall()
                if not res[0][0] == 0:
                    continue

                print(movie_name)

                #插入
                query = 'INSERT INTO movie(movie_name) VALUES (%s)'
                self.cursor.execute(query, (movie_name,))
                self.conn.commit()

    #从数据库中读取电影名并通过豆瓣查询
    def search_movies(self, username='yinwoods'):
        query = 'SELECT movie_name FROM movie'
        self.cursor.execute(query)
        self.conn.commit()
        movie_lists = []
        res = self.cursor.fetchall()

        #self.get_douban_infos(username)

        movie_dicts = {
            '玩具总动员1' : '玩具总动员',
            '蝙蝠侠之黑暗骑士崛起' : '蝙蝠侠：黑暗骑士崛起',
            '我和厄尔及将死的女孩' : '我和厄尔以及将死的女孩',
            '寄生兽 完结篇' : '寄生兽：完结篇',
            '蜡笔小新 搬家物语' : '蜡笔小新：我的搬家物语 仙人掌大袭击',
            '歌曲改变人生' : '再次出发之纽约遇见你',
            '复仇者联盟2' : '复仇者联盟2：奥创纪元',
            '蝙蝠侠·黑暗骑士' : '蝙蝠侠：黑暗骑士',
            'Inside Out' : '头脑特工队',
            'Detachment' : '超脱',
            '心之谷' : '侧耳倾听',
            'Minions' : '小黄人大眼萌',
            '谍中谍5' : '碟中谍5：神秘国度',
            '钢琴师' : '钢琴家',
            '灵异第六感' : '第六感',
            '我是谁' : '我是谁：没有绝对安全的系统',
            '七号房的礼物' : '7号房的礼物',
            '小森林冬春篇' : '小森林 冬春篇',
            '本杰明巴顿奇事' : '本杰明·巴顿奇事',
            '进击的巨人' : '进击的巨人真人版：前篇',
            '王牌特工' : '王牌特工：特工学院',
            '海贼王之3D2Y' : '海贼王15周年纪念特别篇——幻之篇章「3D2Y 跨越艾斯之死！与路飞...',
            '小森林之夏秋篇' : '小森林 夏秋篇',
            '模仿游戏{祖师爷}' : '模仿游戏',
            '疯狂的麦克斯' : '疯狂的麦克斯4：狂暴之路',
            '重返二十岁' : '重返20岁',
            '哆啦a梦之伴我同行' : '哆啦A梦：伴我同行',
            '阴儿房' : '潜伏'
        }

        for movie in res:
            movie_name = movie[0]
            new_movie_name = movie_name

            if movie_name in movie_dicts:
                new_movie_name = movie_dicts[movie_name]

            query = 'SELECT rate, watch_time, comment, movie_types, movie_language, movie_country, movie_director FROM douban_movie WHERE movie_name = (%s)'
            self.cursor.execute(query, (new_movie_name,))
            self.conn.commit()
            res = self.cursor.fetchall()
            try:
                [rate, watch_time, comment, movie_types, movie_language, movie_country, movie_director] = res[0]
            except IndexError as e:
                print(movie_name)
                continue

            watch_time = str(watch_time).split(' ')[0]
            query = 'UPDATE movie SET rate = %s, watch_time = %s, comment = %s, movie_types = %s, movie_language = %s, movie_country = %s, movie_director = %s WHERE movie_name = %s'
            self.cursor.execute(query, (rate, watch_time, comment, movie_types, movie_language, movie_country, movie_director, movie_name))
            self.conn.commit()
        return

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

        #爬取第一页
        url_lists = []
        url_lists.append('https://movie.douban.com/people/' + username + '/collect?sort=time&amp;start=0&amp;filter=all&amp;mode=list&amp;tags_sort=count')

        html_doc = BeautifulSoup(self.get_html(url_lists[0]), 'html.parser')
        #获取所有页链接
        pages = html_doc.select('#content > div')[1].select('div.article > div > a')
        for url in pages:
            url_lists.append(url['href'])

        for url in url_lists:
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

                print(movie_info)

                movie_types = movie_info['类型']
                movie_language = movie_info['语言']
                movie_country = movie_info['制片国家/地区']
                movie_director = movie_info['导演']

                query = 'INSERT INTO douban_movie(movie_name, rate, watch_time, comment, movie_types, movie_language, movie_country, movie_director) VALUES (%s, %s, %s, %s, %s, %s, %s, %s)'
                self.cursor.execute(query, (movie_name, rate, watch_time, comment, movie_types, movie_language, movie_country, movie_director))
                self.conn.commit()
        return

    #数据可视化展示
    def show_datas(self):

        query = 'SELECT movie_name, movie_types FROM movie'
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

        workbook = xlsxwriter.Workbook('chart.xlsx')
        worksheet = workbook.add_worksheet()

        chart = workbook.add_chart({'type' : 'column'})

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
            'x_scale' : 3,
            'y_scale' : 2
        })

        #设置标题
        chart.set_title({
            'name' : '15/07/01 - 16/06/30 观影统计'
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
        workbook.close()

def main():
    crawler = Crawler()
    url_lists = []
    url_lists.append('http://blog.yinwoods.com/book-movie-list/movielist-2015.html')
    url_lists.append('http://blog.yinwoods.com/book-movie-list/movielist.html')

    #crawler.get_movies(url_lists)

    #crawler.search_movies('yinwoods')

    crawler.show_datas()

if __name__ == "__main__":
    main()

