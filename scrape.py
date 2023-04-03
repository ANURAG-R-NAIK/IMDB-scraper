#### TOP 250 MOVIES PAGE OF IMDB

from bs4 import BeautifulSoup
import requests, openpyxl

excel = openpyxl.Workbook()
print(excel.sheetnames)
sheet = excel.active
sheet.title = 'Top Rated Movies'  #title of the excel sheet
print(excel.sheetnames)
sheet.append(["Move Rank", "Movie Name", "Year of Release", "IMDB Rating"]) #title of each column


try:
    source = requests.get("https://www.imdb.com/chart/top/") #REQUESTS FOR THE DATA FROM THE WEBSITE

    source.raise_for_status() # CHECKS FOR IF THE URL IS CORRECT OR NOT SO,TRY AND XCEPTION ARE USED
    
    
    soup = BeautifulSoup(source.text, "html.parser")
    
    movies = soup.find('tbody', class_ ="lister-list").find_all('tr')#THIS FINDS A THE tr TAGS
    #EACH OF THE TBODY CONTAINS THE INFO ABOUT EACH OF THE MOVIES
    
    for movie in movies:
        
        name = movie.find('td', class_="titleColumn").a.text # EXTRACTS THE NAME OF A SINGLE MOVIE NAME
        
        rank = movie.find('td', class_="titleColumn").get_text(strip = True).split('.')[0]  #this accesses all the
                    #data related to 1 then we split it with the dot where the dot comes,  which returns a list
                    #then we take the first element with[0] which gives us the rank
                    
        year = movie.find('td', class_="titleColumn").span.text.strip('()') #strip removes the brackets from the year
        
        rating = movie.find('td', class_ = 'ratingColumn imdbRating').strong.text 
        
        
        print(rank, name, year, rating)
        sheet.append([rank, name, year, rating])
    
except Exception as e:
    print(e)

excel.save('IMDB Movies Ratings.xlsx')