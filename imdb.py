import requests , openpyxl
from bs4 import BeautifulSoup

excel=openpyxl.Workbook()
print(excel.sheetnames)
sheet = excel.active
sheet.title = 'Top rated movies'
sheet.append(['Movvie Rank', 'Movie Name', 'Year of relase', 'IMDB rating'])



try:
    source = requests.get('https://www.imdb.com/chart/top/')
    source.raise_for_status()

    soup=BeautifulSoup(source.text,'html.parser')
    movies = soup.find('tbody', class_="lister-list").find_all('tr')

    for movie in movies:
        name=movie.find('td', class_="titleColumn").a.text
        #print(name)
        ranking=movie.find('td', class_="titleColumn").get_text(strip=True).split('.')[0]
        year=movie.find('td', class_="titleColumn").span.text.strip('()')
        rating=movie.find('td', class_="ratingColumn imdbRating").strong.text
        print (ranking, name, year, rating)
        sheet.append([ranking, name, year, rating])

except Exception as e:
    print(e)

excel.save('IMDB Movies.xlsx')
