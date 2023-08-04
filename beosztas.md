import cv2
import numpy as np
import openpyxl

# load image
img = cv2.imread("C:/Python/123.jpg")
#color = img[670, 1080]

size = img.shape  # (695, 1095, 3)
gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
gray_array = np.array(gray)

lower_gray = 0  #basicly Black 
upper_gray = 210  #almost white

# size[0] // 2 - a kÖzepe a képnek

# finde the edge of the squer from the mide of the picture (1080)
for i in range(gray.shape[1] - 1, gray.shape[0] // 2, -1):
    pixel_col = gray[:, i]
    gray_mask = cv2.inRange(pixel_col, lower_gray, upper_gray)
    if np.any(gray_mask):
        edge_number = i
        break

# find right cornel coordinates (ex: 670)
for i in range(gray.shape[0] // 2, size[0] , 1):
    pixel_col= gray[i, :]
    gray_mask = cv2.inRange(pixel_col, lower_gray, upper_gray)
    if not np.any(gray_mask):
        edge_cornel_cornel = i - 1
        break
    
#findig left  cornel coordinates (ex: 149)
for i in range(edge_number, 1 , -1):
    pixel = np.array(gray[edge_cornel_cornel, i])
    # print(f"Testing pixel ({edge_cornel_cornel}, {i})") testing 
    gray_mask = cv2.inRange(pixel , lower_gray, upper_gray)
    if not np.any(gray_mask):
        edge_cornel_left = i + 2
        break
     
 #findig upper left  cornel coordinates (ex: )    
for i in range(edge_cornel_cornel, 1 , -1):
    pixel = np.array(gray[i, edge_cornel_left]) 
    #print(f"Testing pixel ({i}, {edge_cornel_left})") #testing 
    gray_mask = cv2.inRange(pixel , lower_gray, upper_gray)
    if not np.any(gray_mask):
        edge_cornel_upper_left = i +2
        break    
     
     
     
#print(f"Edge_number: {edge_number}, Edge_cornel_ cornel: {edge_cornel_cornel}, Edge_Cornel_LEFT: {edge_cornel_left}, Edge_cornel_upper_left: {edge_cornel_upper_left}") 
#print(f"Right Cornel: ({edge_cornel_cornel}, {edge_number})     Left Cornel: ({edge_cornel_left},{edge_cornel_upper_left} )") 


#upper_left = (149, 49)
#bottom_right = (1080, 670)
#img_corp= img[49:670, 149:1080]
#img_crop = img[upper_left[1]:bottom_right[1], upper_left[0]:bottom_right[0]]

img= img[edge_cornel_upper_left:edge_cornel_cornel, edge_cornel_left:edge_number]

# save cropped image
cv2.imwrite("C:/Python/123c.jpg", img)
#img = cv2.imread("C:/Users/User/OneDrive/Dokumentumok/Python/123c.jpg")
##print(img.shape)

# Making a matrix and filling it up whit data based on tyhe picture
grid = (img.shape[0]//30, img.shape[1]//30)
matrix = np.zeros(grid, dtype=int)

for m in range(15, img.shape[0]+1, 30):
    for s in range(15, img.shape[1]+1, 30):
        pixel = img[m, s]
        #print(f"magaság: {m}, Szélesség: {s}, Pixel szín: {pixel_color}")
        if (pixel[1] >= 50 and pixel[0]< 180 and pixel[2]< 180): #IF it's GREEN
            matrix[m//30,s//30] = 1
        if (pixel[1] < 30 and pixel[0] < 30  and pixel[2] < 30 ): #IF it's BLACK
            matrix[m//30,s//30] = 2



#start the MAGIC EXEL TÁBLA SZERKERSZTÉSE
## ws fogjuk menten; A mátrixba érték 0- 0=van 1= ügyeletes 2= Szabadságon;  
## az alapj'n h a mátrix X kordinátájában hanyadik (0-19ig) jön ki a név; Y-ba meg a nap. 

wb = openpyxl.Workbook()
ws = wb.active
ex_row = 2
ex_col = 2
names = []   ## names az egy lista a Names.txt-bol

#open the file whit names
with open("C:/Python/Names.txt", encoding='UTF-8') as file:
    for line in file:
        names.append(str(line.strip()))

# Populate the top row with headers
ws.cell(row=1, column=1).value = ' '
ws.cell(row=1, column=2).value = 'Ügyeletes vezető'
ws.cell(row=1, column=3).value = 'Szülőszobás'
ws.cell(row=1, column=5).value = 'Szabadság'

# Populate the first column with dates
for i in range(2, 33):
    ws.cell(row=i, column=1, value=i-1)

for col in range(0,matrix.shape[1],1):
    for row in range(0,matrix.shape[0],1):
        if(matrix[row,col]==1):
            #print(f"Puting info in cell: {ex_col},{ex_row} - {names[row]} matrix steps row: {row} column: {col}")
            ws.cell(ex_row, ex_col, value= names[row]) 
            ex_col += 1   
            if (ex_col> 3):
                ex_col=2
                ex_row += 1
                col += 1
                break
 

ex_row = 1
                
for col in range(0,matrix.shape[1],1):
    ex_col = 5
    ex_row +=1
    for row in range(0,matrix.shape[0],1):
        if(matrix[row,col]==2 and row != 8 and row != 18): #IF it's BLACK és nem Stefi vagy Németh
            if(col != 0):
                if(matrix[row,col-1]!=1): #Post ügyeletesek kivételek
                #print(f"Puting info in cell: {ex_col},{row+1} - {names[row]} matrix steps row: {row} column: {col}")
                 ws.cell(ex_row, ex_col, value= names[row][:3]) 
                 ex_col += 1   
            else:
                ws.cell(ex_row, ex_col, value= names[row][:3]) 
                ex_col += 1
                      
wb.save("C:/Python/test.xlsx")
wb.close()

############################################################################################################################################
############################################################################################################################################
############################################################################################################################################


## Nagy beosztás   ??? A műtétes embereket MAJD 4-es számot kapnak 


## 0 FZIS BERAKON AZ ÜGYELETESEK AZ ÜGYELETES SORBA
wb = openpyxl.Workbook()
ws = wb.active

# 0.1 Az ügyeletesek beosztása
ex_row = 2
ex_col = 2
for i in range(2, 33):
    ws.cell(row=i, column=1, value=i-1)

for col in range(0,matrix.shape[1],1):  ## lévén hogy a másod ügyletes a 13. oszlopba van és a másod ügyeletes a 12. sorba meg kellet forditani
    for row in range(0,matrix.shape[0],1):
        if(matrix[row,col]==1):
            if (ex_col ==2) :
                ex_col = 13 
                ##print(f"Puting info in cell: {ex_col},{ex_row} - {names[row]} matrix steps row: {row} column: {col}") ##  TESZ
                ws.cell(ex_row , ex_col , value= names[row]) 
                ex_col -= 1   
            else:
                ##print(f"Puting info in cell: {ex_col},{ex_row} - {names[row]} matrix steps row: {row} column: {col}") ##  TESZT
                ws.cell(ex_row , ex_col , value= names[row]) 
                ex_col=2
                ex_row += 1
                col += 1
                break
            
 
#0.2   A távollevők beosztása

ex_row = 1
                
for col in range(0,matrix.shape[1],1):
    ex_col = 15 ## me a 15 sorba a post-ügyeletes megy
    ex_row +=1
    for row in range(0,matrix.shape[0],1):
        if(matrix[row,col]==2 and row != 8 and row != 18): #IF it's BLACK és NEM Stefi vagy NEM Németh
            if col != 0:
                if(matrix[row,col-1]!=1): #Post ügyeletesek kivételek
                 #print(f"Puting info in cell: {ex_col},{row+1} - {names[row]} matrix steps row: {row} column: {col}")
                 ws.cell(ex_row, ex_col, value= names[row][:3]) 
                 ex_col += 1   
            else:
                ws.cell(ex_row, ex_col, value= names[row][:3]) 
                ex_col += 1
                
#0.3  A post ügyeletesek 13 oszlopban 
fele = 1  ## nezi hogy megvan már a fele a postugyeletnek
ex_col = 14
ex_row = 2
postugy = ""
for i in range(2, 33):
    ws.cell(row=i, column=1, value=i-1)

for col in range(0,matrix.shape[1],1):
    for row in range(0,matrix.shape[0],1):
        if(matrix[row,col]==1):  
          #print(f"Puting info in cell: {ex_col},{ex_row} - {names[row]} matrix steps row: {row} column: {col}")  ##csökevény a Fázis 0,1-ből
         if (fele < 2):  
          postugy += names[row][:3] + ", "
          fele += 1 
          #print(f"Half info in cell: {ex_col},{ex_row} - {postugy}")     
         else:  # ha mar megvan egy ugyeletes 
          ex_row += 1
          fele = 1  ## visszálitunk uresre
          postugy += names[row][:3]
          ws.cell(ex_row , ex_col , value= postugy)   # beirjuk a stringet
          #print(f"Puting info in cell: {ex_col},{ex_row} - {postugy}")  
          col += 1 # ugrunk egyet a mátrix sorában
          postugy = ""
                
          break

## 0 Fázis  Csináljuk meg az alap táblázatott ügyelet + távolévők (+ szinezés és bordering)
import calendar
import holidays
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment
from openpyxl.styles import Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

ws.merge_cells(start_row=1, start_column=14, end_row=1, end_column=23)
merged_cell = ws.cell(row=1, column=14)
merged_cell.value = "Szabadság / Távol"

# Ask for the year and month
year = int(input("Enter the year: "))
month = int(input("Enter the month (1-12): "))

# Set the column names
column_names = [
    str(year), str(calendar.month_name[month])[:3], "I. Akut.", "II. Amb", "Referens","Terhes", "III Onkó Meno. GyeR.", "UH","Terhesg.", "Kisműtő", "Marcali", "I Ügyeletes", "II Ügyeletes", "post.ügy",
]
ws.cell(1 , 24 , value= "V") ## a verziót is beirjuk 202


# Set the column widths
## év/1,2,3..   Hónap/nap   "I. Akut.", "II. Amb", "Terhes", "III Onkó Meno. GyeR.", "UH",  "Terhesg.", "Kisműtő", "Marcali", "I Ügyeletes", "II Ügyeletes", "post.ügy"
column_widths = [5, 4, 10, 10, 10, 10, 18, 10, 10, 10, 10, 12, 12, 10, 20, 10]

for col in range(15, 24 + 1):
    ws.column_dimensions[get_column_letter(col)].width = 4


# Set the initial row and column indices
row_index = 1
column_index = 1

# Set the fill color for the desired columns
fill_color = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")

# Write the column names and set column widths
for name, width in zip(column_names, column_widths):
    cell = ws.cell(row=row_index, column=column_index, value=name)
    cell.alignment = Alignment(horizontal="center")
    ws.column_dimensions[get_column_letter(column_index)].width = width
    column_index += 1


# Get the number of days in the month
num_days = calendar.monthrange(year, month)[1]
maxdays = num_days

# Set the dates and day names in Hungarian
day_names_hungarian = ["H", "K", "Sz", "Cs", "P", "Sz", "V"]
for day in range(1, num_days + 1):
    # Get the day name in Hungarian
    day_name_hungarian = day_names_hungarian[calendar.weekday(year, month, day)]
    # Write the date and day name in the first two columns
    ws.cell(row=row_index + day , column=1, value=day)
    ws.cell(row=row_index + day , column=2, value=day_name_hungarian)


# Set the fill color for the desired columns
green_color = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
green_D_color = PatternFill(start_color="008400", end_color="008400", fill_type="solid")
pink_color = PatternFill(start_color="FF91AF", end_color="FF91AF", fill_type="solid")
cream_color = PatternFill(start_color="FFC300", end_color="FFC300", fill_type="solid")


# Color the cells in the 11th and 12th columns
for day in range(1, num_days + 1):
    day_of_week = calendar.weekday(year, month, day)
    if day_of_week >= 5:  # Saturday (5) or Sunday (6)
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row=day + 1, column=col)
            cell.fill = green_color

#seeing the holidays
holidays_hungary = holidays.Hungary(years=year)
for date, name in holidays_hungary.items():
 if date.month == month:
  print(f"{date}: {name}")
  for col in range(1, ws.max_column + 1):
   cell = ws.cell(row=date.day + 1, column=col)
   cell.fill = pink_color


# Color columns 13 and 14 till row 31
for row in range(2, 33):
    for col in range(12, 14):
        cell = ws.cell(row=row, column=col)
        cell.fill = green_color

for day in range(1, num_days + 1):
    day_of_week = calendar.weekday(year, month, day)
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=day + 1, column=col)
        if day_of_week >= 5:  # Weekend or columns 13 or 14
            if col == 12 or col == 13:
             cell.fill = green_D_color
             

       
border_style = Side(border_style="thin")

# Apply borders to the range of cells
for row in range(1, maxdays+2):
    for col in range(1, 25):
        cell = ws.cell(row=row, column=col)
        cell.border = Border(top=border_style, right=border_style, bottom=border_style, left=border_style)


## 1 Fázis Csináljuk meg a sablont (+ szinezés)

def TEST (name, day, hely):
    if(matrix[name, day] == 2 ):
        cell = ws.cell(day+2, hely)
        cell.fill = cream_color
        print(f"Nincs     : {day}.-án {names[name]} a {hely} ambulancián")      
    ws.cell(day+2, hely, value= names[name])
        
def Sablon(day, hely, name):
    hiba = 0
    cell = ws.cell(day, hely)
    day -= 2
    
    if (hely == 8) :                        ## Ha ultrahang
        if(matrix[name, day] == 2 ):
            cell.fill = cream_color
            print(f"Nincs     : {day}.-án {names[name]}")
            if (matrix [7, day] < 1):
                name = 7
            elif (matrix [2, day] != 2):
                name = 2
            elif (matrix [3, day] != 2):
                name = 3
            elif (matrix [1, day] != 2):
                name = 1
            elif (matrix [0, day] != 2):
                name = 0
            else: 
                print(f"NINCS ULTRAHANOGS !!!! {day}-án ")
                hiba = 1
            print(f"helyetesit: {names[name]}")
        if hiba == 0:                    
         ws.cell(day+2, hely, value= names[name])
    
    elif (hely == 6):    ## Terhes
        TEST(name, day, hely)   
    elif (hely == 7):   ## CSNT
        TEST(name, day, hely)  
    elif (hely == 9):   ## Gondozo 
        TEST(name, day, hely)       
    elif (hely == 2):   ## Akut  
        TEST(name, day, hely) 
################################################################################

    


""" # Mátrix kiirása's
filename = "C:\Python\matrix.txt"   
def print_matrix(matrix, filename):
    with open(filename, "w") as file:
        for row in matrix:
            file.write(" ".join(map(str, row)) + "\n")
         
print_matrix(matrix, filename)
""" #
## 1.1 UH 
Akut = 2 ## akut
Terhes = 6 ## Terhes ambulancia
CSNT = 7 ## III Amb. onko gyermek 
UH = 8 ## 7. oszlopban van az UH 
Gondoz = 9 ## Terhes Gondozó
day = 2

##for day in range(1, num_days + 1):
while (day <= num_days+1): 
    if(day <= num_days):
        day_of_week = calendar.weekday(year, month, day)
    else:
        if(day_of_week < 6):
            day_of_week  +=1
        else:
            day_of_week = 0
    if day_of_week == 1: ## Hétfő
        
        Sablon(day, UH, 3,)   ## Horvath 
        Sablon(day, CSNT, 2,)
        Sablon(day, Gondoz, 19,)
        
        
    if day_of_week == 2: ## Kedd
        Sablon(day, UH, 2,)
        Sablon(day, CSNT, 18,)
        Sablon(day, Terhes, 19,)
        Sablon(day, Gondoz, 20,)
        
    if day_of_week== 3: ## Szerda
        Sablon(day, UH, 1,)
        Sablon(day, CSNT, 8,)
        Sablon(day, Terhes, 17,)
        Sablon(day, Gondoz, 20,)
        
    if day_of_week== 4: ## Csütörtök
        ws.cell(day, UH , value= "Bencze") 
        Sablon(day, Gondoz, 17,)
        
    if day_of_week == 5: ## Péntek
        Sablon(day, UH, 0,)
        Sablon(day, Gondoz, 17,)
        
    #print(f"Day {day} day_of_week ,{day_of_week}, num_days{num_days}")   
    day += 1  
        
    
    
    
    

## 2 Fázis  Elenőrizzük a sablont (vannak-e az emberek akik be vannak irva a sablonba)
## 3 Fázis  A sorok popupációja (berakni az embereket az ambulanciákra)

wb.save("C:/Python/NagyBeosztás.xlsx")            
##print(matrix) #seeng the matrix 
##column_sums = np.sum(matrix, axis=0)  
##print(column_sums)      # testing the matrix 
