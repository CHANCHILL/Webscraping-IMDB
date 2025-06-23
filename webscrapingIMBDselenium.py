from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
import time
from openpyxl import Workbook

#create a new workbook
excel = Workbook()

#access the active sheet
sheet = excel.active

#renaming the active sheet and appending required column headers
sheet.title = 'IMDB Top 250 Movies'
sheet.append(['rank', 'movie_name', 'releaseyear', 'rating'])

#setting up selenium in headless mode
options = Options()
options.headless = True 
driver = webdriver.Chrome(options=options)

#loading the IMDB page with Javascript
driver.get("https://www.imdb.com/chart/top/")
time.sleep(3)  # Wait for JS to load all movies

#using BeautifulSoup to parse the page
soup = BeautifulSoup(driver.page_source, 'html.parser')

#extracting movies data into object movies
movies = soup.find('ul', class_="ipc-metadata-list ipc-metadata-list--dividers-between sc-e22973a9-0 khSCXM compact-list-view ipc-metadata-list--base").find_all('div', class_="sc-995e3276-0 eXDZXb")

#checking the count of movies extracted
print(len(movies)) 

#looping through movies and loading the data to excel
for movie in movies:
    rank = movie.find('h3', class_="ipc-title__text ipc-title__text--reduced").text.split('.')[0]
    name = movie.find('h3', class_="ipc-title__text ipc-title__text--reduced").text.split('.')[1].strip()
    release_year = movie.find('span', class_="sc-dc48a950-8 gikOtO cli-title-metadata-item").text
    rating = movie.find('span', class_="ipc-rating-star--rating").text
   
   #printing the list of the movies during iteration
    print(rank, name, release_year, rating) 

  #loading the movie data during iteration into excel
    sheet.append([rank, name, release_year, rating])
    
#save the excel file 
excel.save('IMDB Top Rated Movies.xlsx')

#closing the browser running in background
driver.quit()
