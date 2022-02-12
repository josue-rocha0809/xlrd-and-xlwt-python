import xlrd
import xlwt
bookWork = xlrd.open_workbook('C:/Cursos/bigData/Kardexjosue.xls')
newBook= xlwt.Workbook() #creates newBook
newSheet= newBook.add_sheet("page1")  #new sheet in the excel file

sh = bookWork.sheet_by_index(0)
rows = sh.nrows
val = sh.row(0)
cont=0
mayor=0
menor=0
lim=0

newSheet.write(0,0,"calificaciones")  # creates calificaciones column
newSheet.write(0,1,"entrenamiento")   # creates calificaciones column

for i in range(rows):
    lim=i
    fil=sh.row(i)
    cont+=(fil[4].value)  # counter for quali summation
    newSheet.write(i+1, 0, fil[4].value)   # added in the column of grades
    if fil[4].value >=80  :
        mayor+=(fil[4].value)          # value for grades greater than 80
    else:
        menor += (fil[4].value)         # value for grades under  80
    if fil[4].value<=70:
        newSheet.write(i+1, 1, 'Si')    # added in the newBook ´Si´ if the value is less than or equal to 70
    else :
        newSheet.write(i + 1, 1, 'No')  # added in the newBook ´no´ if the value is greater than or equal to 80
    if lim == rows - 1:
        print("promedio final = ", cont / rows)
        print("promedio 80s   = ", mayor / rows)
        print("promedio 70s   = ", menor / rows)
    newBook.save('C:/Cursos/bigData/newExcel/calificaciones.xls')

