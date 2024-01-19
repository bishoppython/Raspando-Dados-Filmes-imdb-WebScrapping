# importando modulos
from bs4 import BeautifulSoup
import requests, openpyxl


# inserindo sites/targets
url = ['https://www.imdb.com/chart/toptv/', 'https://www.imdb.com/chart/top/']

# Habilitando o módulo Excel
excel = openpyxl.Workbook()
print(excel.sheetnames)
sheet = excel.active
sheet.title = "Top 250 Filmes"
print(excel.sheetnames)
sheet.append(['Ranking', 'Título', 'Ano de Lançamento', 'Nota IMDB' ])

# try/except
try:
    # Passando o url e requisitando
    origem = requests.get(url[1])
    origem.raise_for_status() #validando o status do target

    soup = BeautifulSoup(origem.text, 'html.parser') #convertendo o retorno em text

    #print(soup)

    filmes = soup.find('tbody', class_="lister-list").find_all('tr')

    for filme in filmes:

        titulo = filme.find('td', class_="titleColumn").a.text
        rank = filme.find('td', class_="titleColumn").get_text(strip=True).split('.')[0]
        ano = filme.find('td', class_="titleColumn").span.text.strip('()')
        nota = filme.find('td', class_="ratingColumn imdbRating").strong.text

        sheet.append([rank, titulo, ano, nota])
except Exception as e:
    print(e)

# Salvando em excel
excel.save('IMDB Filmes Memoráveis.xlsx')