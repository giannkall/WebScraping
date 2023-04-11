from bs4 import BeautifulSoup
import requests, openpyxl

excel = openpyxl.Workbook()

sheet = excel.active
sheet.title = 'Top Rated TV Shows'


sheet.append(['Movie Rank', 'Movie Name', 'Year of Realease', 'IMBD Rating'])

try:
    source = requests.get('https://www.imdb.com/chart/toptv?pf_rd_m=A2FGELUUNOQJNL&pf_rd_p=470df400-70d9-4f35-bb05-8646a1195842&pf_rd_r=GVAAC2H41AJ2P11ZVY26&pf_rd_s=right-4&pf_rd_t=15506&pf_rd_i=top&ref_=chttp_ql_6')
    source.raise_for_status() #if there is a problem with the web site it catches the error
    
    soup=BeautifulSoup(source.text, 'html.parser')
    
    tvshows = soup.find('tbody', class_="lister-list").find_all('tr')
    
    for tvshow in tvshows:
        
        name = tvshow.find('td', class_="titleColumn").a.text

        rank= tvshow.find('td', class_="titleColumn").get_text(strip=True).split('.')[0]

        year = tvshow.find('td', class_="titleColumn").span.text.strip('()')

        rating = tvshow.find('td', class_="ratingColumn imdbRating").strong.text

        print(rank, name, year, rating)
        sheet.append([rank, name, year, rating])

except Exception as e:
    print(e)

excel.save('IMBD TV Shows Ratings.xlsx')  