import requests
import pandas as pd
from bs4 import BeautifulSoup
from time import sleep


def request_douban(url):
    """发送请求获取豆瓣页面内容"""
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3'
    }
    try:
        response = requests.get(url, headers=headers, timeout=10)
        return response.text if response.status_code == 200 else None
    except requests.RequestException as e:
        print(f"请求失败: {e}")
        return None


def parse_movie_data(soup):
    """解析电影数据"""
    movies = []
    movie_list = soup.find(class_='grid_view')

    if not movie_list:
        return movies

    for item in movie_list.find_all('li'):
        movie = {
            '排名': item.find(class_='pic').get_text().strip() if item.find(class_='pic') else '无',
            '名称': item.find(class_='title').get_text().strip(),
            '评分': item.find(class_='rating_num').get_text().strip(),
            '短评': item.find(class_='quote').get_text().strip() if item.find(class_='quote') else '无'
        }
        movies.append(movie)
    return movies


def save_to_excel(data, filename='豆瓣电影Top250.xlsx'):
    """保存数据到Excel文件"""
    try:
        df = pd.DataFrame(data)
        # 调整列顺序
        columns = ['排名', '名称', '评分', '短评']
        df = df[columns]

        # 设置Excel写入格式
        writer = pd.ExcelWriter(filename, engine='openpyxl')
        df.to_excel(writer, index=False, sheet_name='电影数据')

        # 自动调整列宽
        worksheet = writer.sheets['电影数据']
        for column in worksheet.columns:
            max_length = max(len(str(cell.value)) for cell in column)
            worksheet.column_dimensions[column[0].column_letter].width = max_length + 2

        writer.close()
        print(f"数据已成功保存到 {filename}")
    except Exception as e:
        print(f"保存文件时出错: {e}")


def main(page, all_movies):
    """主爬虫函数"""
    url = f'https://movie.douban.com/top250?start={page * 25}&filter='
    print(f"正在爬取第 {page + 1} 页数据...")

    html = request_douban(url)
    if not html:
        return

    soup = BeautifulSoup(html, 'lxml')
    movies = parse_movie_data(soup)
    all_movies.extend(movies)


if __name__ == '__main__':
    all_movies = []

    # 爬取前10页数据（共250条）
    for page in range(10):
        main(page, all_movies)
        sleep(1)  # 增加请求间隔

    # 保存数据到Excel
    if all_movies:
        save_to_excel(all_movies)
        print(f"共爬取到 {len(all_movies)} 条数据")
    else:
        print("未获取到有效数据")