import sqlite3
import datetime
#import os
import xlrd
import numpy as np
import matplotlib.pyplot as plt
import csv
import xlsxwriter

file_location="D:\Cloud Computing\Final_Project\bank_info.xlsx"
wbook= xlrd.open_workbook(file_location)

conn = sqlite3.connect('bank_info.db')
c = conn.cursor()

# Create table ID

Sheetid=wbook.sheet_by_index(0)
bank_id = []

c.execute("DROP TABLE IF EXISTS id")
c.execute('''CREATE TABLE id
               (User Name, 
		ID,
		CardNumber)''')

for rowx in xrange(Sheetid.nrows):
     bank_id.append(tuple(Sheetid.cell(rowx, colx).value 
                               for colx in xrange(Sheetid.ncols)))
c.executemany("INSERT INTO id VALUES (?,?,?)", bank_id)
conn.commit()

# Create table preferred_loc
Sheetpl=wbook.sheet_by_index(1)
bank_pl = []

c.execute("DROP TABLE IF EXISTS prefloc")
c.execute('''CREATE TABLE prefloc
               (ID,
		Preferred Location)''')

for rowx in xrange(Sheetpl.nrows):
     bank_pl.append(tuple(Sheetpl.cell(rowx, colx).value 
                               for colx in xrange(Sheetpl.ncols)))
c.executemany("INSERT INTO prefloc VALUES (?,?)", bank_pl)
conn.commit()

# Create table account

Sheetaccount=wbook.sheet_by_index(2)
bank_account = []

c.execute("DROP TABLE IF EXISTS account")
c.execute('''CREATE TABLE account
               (ID,
		Account)''')

for rowx in xrange(Sheetaccount.nrows):
     bank_account.append(tuple(Sheetaccount.cell(rowx, colx).value 
                               for colx in xrange(Sheetaccount.ncols)))
c.executemany("INSERT INTO account VALUES (?,?)", bank_account)
conn.commit()


# Table of Amount
SheetAmount=wbook.sheet_by_index(3)
bank_Amount = []

c.execute("DROP TABLE IF EXISTS Amount")
c.execute('''CREATE TABLE Amount
               (ID,
		amount,
		Account Number)''')

for rowx in xrange(Sheetaccount.nrows):
     bank_Amount.append(tuple(SheetAmount.cell(rowx, colx).value 
                               for colx in xrange(SheetAmount.ncols)))
c.executemany("INSERT INTO Amount VALUES (?,?,?)", bank_Amount)
conn.commit()


# Table of location

Sheetloc=wbook.sheet_by_index(4)
bank_loc = []

c.execute("DROP TABLE IF EXISTS Location")
c.execute('''CREATE TABLE Location
               (ID,
		location,
		amount,Date)''')

for rowx in xrange(Sheetloc.nrows):
     bank_loc.append(tuple(Sheetloc.cell(rowx, colx).value 
                               for colx in xrange(Sheetloc.ncols)))
c.executemany("INSERT INTO Location VALUES (?,?,?,?)", bank_loc)
conn.commit()


# Table of Frequency

SheetFreq=wbook.sheet_by_index(5)
bank_Freq = []

c.execute("DROP TABLE IF EXISTS Frequency")
c.execute('''CREATE TABLE Frequency
               (ID,
		Date,
		amount)''')

for rowx in xrange(SheetFreq.nrows):
     bank_Freq.append(tuple(SheetFreq.cell(rowx, colx).value 
                               for colx in xrange(SheetFreq.ncols)))
c.executemany("INSERT INTO Frequency VALUES (?,?,?)", bank_Freq)
conn.commit()


#############################################
# extracting all values for each id for non-preferred locations per user
c.execute('SELECT Location,amount,Date FROM Location WHERE  ID==1 AND Location!=6845 AND Location!=1032 AND Location!=1975 AND Location!=6045 AND Location!=2950')
id_1_loc = c.fetchall()

c.execute('SELECT Location, amount,Date FROM Location WHERE  ID==2 AND Location!=35345 AND Location!=59273 AND Location!=81813 AND Location!=75546 AND Location!=34076')
id_2_loc = c.fetchall()

c.execute('SELECT Location, amount,Date FROM Location WHERE  ID==3 AND Location!=39372 AND Location!=39372 AND Location!=18658 AND Location!=46643 AND Location!=75561')
id_3_loc = c.fetchall()

c.execute('SELECT Location, amount,Date FROM Location WHERE  ID==4 AND Location!=10111 AND Location!=19907 AND Location!=46082 AND Location!=59521 AND Location!=59524')
id_4_loc = c.fetchall()

# extracting all values for each id for amounts less than 10$ and their respective dates
c.execute('SELECT amount,Date FROM Frequency WHERE  ID==1 AND amount<=10')
id_1_amt = c.fetchall()

c.execute('SELECT amount,Date FROM Frequency WHERE  ID==2 AND amount<=10')
id_2_amt = c.fetchall()

c.execute('SELECT amount,Date FROM Frequency WHERE  ID==3 AND amount<=10')
id_3_amt = c.fetchall()

c.execute('SELECT amount,Date FROM Frequency WHERE  ID==4 AND amount<=10')
id_4_amt = c.fetchall()

# extracting all values for each id for consecutive dates
c.execute('SELECT amount,Date FROM Frequency WHERE  ID==1')
id_1 = c.fetchall()

c.execute('SELECT amount,Date FROM Frequency WHERE  ID==2')
id_2 = c.fetchall()

c.execute('SELECT amount,Date FROM Frequency WHERE  ID==3')
id_3 = c.fetchall()

c.execute('SELECT amount,Date FROM Frequency WHERE  ID==4')
id_4 = c.fetchall()
##############
#create dictionary for Date and amount(1 date, two related amount)
def dict_amt_date(id_0):
    l = len(id_0)
    id_0_fr={}
    for row in range (0,l-1):
            old_date = id_0[row][1]
            new_date = id_0[row+1][1]

            if (old_date == new_date):
                v1=id_0[row][0]
                v2 = id_0[row+1][0]
                k = old_date
                id_0_fr[k] = v1,v2
            
    return (id_0_fr)
    
id_1_fr = dict_amt_date(id_1)    
id_2_fr =dict_amt_date(id_2)
id_3_fr =dict_amt_date(id_3)
id_4_fr =dict_amt_date(id_4)    
#########################################3

#create dictionary for Date and amount<$10(1 date, 1 amount)
def dict_lessamt_date(id_0):
    l = len(id_0)
    id_0Dateamt={}
    for row in id_0:
               v1=row[0]      #amount
               k = row[1]  # date
               id_0Dateamt[k] = v1
            
    return (id_0Dateamt)

id_1_lessamt = dict_lessamt_date(id_1_amt)    
id_2_lessamt =dict_lessamt_date(id_2_amt)
id_3_lessamt =dict_lessamt_date(id_3_amt)
id_4_lessamt =dict_lessamt_date(id_4_amt)

#################################################
#create dictionary for non-preferred locations
def dict_nonPre_loc(id_0):
    l = len(id_0)
    id_nonPreLoc={}
    for row in id_0:
               v1=row[0]  # location
               k = row[1]     # amount
               id_nonPreLoc[v1] = k
               
    return (id_nonPreLoc)
id_1_nonPreLoc = dict_nonPre_loc(id_1_loc)    
id_2_nonPreLoc =dict_nonPre_loc(id_2_loc)
id_3_nonPreLoc =dict_nonPre_loc(id_3_loc)
id_4_nonPreLoc =dict_nonPre_loc(id_4_loc)

User1={}
User1={'LessAmount':[id_1_lessamt], 'FreqDate':[id_1_fr], 'NonPrefLocation':[id_1_nonPreLoc] }

User2={}
User2={'LessAmount':[id_2_lessamt], 'FreqDate':[id_2_fr], 'NonPrefLocation':[id_2_nonPreLoc] }

User3={}
User3={'LessAmount':[id_3_lessamt], 'FreqDate':[id_3_fr], 'NonPrefLocation':[id_3_nonPreLoc] }

User4={}
User4={'LessAmount':[id_4_lessamt], 'FreqDate':[id_4_fr], 'NonPrefLocation':[id_4_nonPreLoc] }
###################################################################################################


def createExeclFile(name, id_0_loc,id_0, id_0_amt):
     
     workbook = xlsxwriter.Workbook(name)

     worksheet = workbook.add_worksheet('NonPrefLocation')
     worksheet1 = workbook.add_worksheet('FreqDate')        
     worksheet2 = workbook.add_worksheet('LessAmount')
     row = 0
     col = 0
     for a,b,c in (id_0_loc):
         worksheet.write(row, col,     a)
         worksheet.write(row, col + 1, b)
         worksheet.write(row, col + 2, c)
         row += 1
     row = 0
     col = 0
     for a,b in (id_0):
         worksheet1.write(row, col,     a)
         worksheet1.write(row, col + 1, b)
         row += 1
     row = 0
     col = 0
     for a,b in (id_0_amt):
         worksheet2.write(row, col,     a)
         worksheet2.write(row, col + 1, b)
         row += 1

     workbook.close()


createExeclFile('USER1.xlsx', id_1_loc,id_1, id_1_amt)
createExeclFile('USER2.xlsx', id_2_loc,id_2, id_2_amt)
createExeclFile('USER3.xlsx', id_3_loc,id_3, id_3_amt)
createExeclFile('USER4.xlsx', id_4_loc,id_4, id_4_amt)
