import requests
from bs4 import BeautifulSoup
import xlwt
from xlwt import Workbook

URL = "https://www.imdb.com/chart/toptv/?ref_=nv_tvv_250"
page = requests.get(URL)
soup = BeautifulSoup(page.content, 'html.parser')
results = soup.find('table', class_="chart full-width")
# print(results.prettify())
shows = results.find_all('tr')

count_2019 = 0
p_rank = ""
p_rating = ""
wb = Workbook()
style = xlwt.easyxf('font: bold 1')
sheet1 = wb.add_sheet('Sheet 1')
sheet1.write(0, 0, 'NO.', style)
sheet1.write(0, 1, 'RANK', style)
sheet1.write(0, 2, 'TITLE', style)
sheet1.write(0, 3, 'YEAR', style)
sheet1.write(0, 4, 'PROTAGONISTS', style)
sheet1.write(0, 5, 'IMDB RATING', style)
i = 1
j = 1
for show in shows:
    j = 1
    # sheet1.write(i, 0, str(i))
    title = show.find('td', class_='titleColumn')
    rating = show.find('td', class_='ratingColumn imdbRating')

    if None in (title, rating):
        continue

    temp_list = title.text.split("\n")
    link = title.find('a')['title']
    if "2019" in title.text:
        count_2019 += 1
    if "pride and prejudice" in title.text.lower():
        p_rating = rating.text
        p_rank = temp_list[1]

    print("Rank: " + temp_list[1].strip())
    print("Title: " + temp_list[2].strip() + " " + temp_list[3].strip()[1:5])
    sheet1.write(i, 1, temp_list[1].strip())
    sheet1.write(i, 2, temp_list[2].strip())
    sheet1.write(i, 3, temp_list[3].strip()[1:5])
    print(f"{link}")
    sheet1.write(i, 4, link)
    print("IMDB Rating: " + rating.text.strip())
    sheet1.write(i, 5, rating.text.strip())
    print()
    sheet1.write(i, 0, str(i))
    i += 1

print(f"Number of top-rated shows from 2019: {count_2019}")
print(f"IMDB Rating for Pride and Prejudice: {p_rating.strip()}")
print(f"IMDB Rank for Pride and Prejudice: {p_rank.strip()}")
wb.save('imdb-web-scrape.xls')
