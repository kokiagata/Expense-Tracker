import pandas as pd
from pathlib import Path
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from format import *
from utils import *

## Name of worksheet
#############################################
wsName = ${nameOfWorkSheet}
#############################################
## Connecting to Google Drive Spread Sheet
gc = gspread.service_account()
sh = gc.open(${nameOfGoogleSheet})

## Selecting the worksheet. If not exists, create anew

worksheet_list = sh.worksheets()   
titles = []
for sheet in worksheet_list:
    titles.append(sheet.title)

if wsName not in titles:
    print(f'{wsName} does not exist. Creating a new worksheet')
    sh.add_worksheet(title=wsName, rows=50, cols=10)
    
else:
    print(f'{wsName} already exists. Proceeding to update')

thisWorksheet = sh.worksheet(wsName)

## Creating a dataframe from downloaded CSV
####################################################
electricityAndWater = 99.97
utilities = 22 + 7.47 + electricityAndWater
dfCap1 = pd.read_csv('path/to/csv')
dfCap2 = pd.read_csv('path/to/csv/2')
dfAmex = pd.read_csv('path/to/csv/3')
dfTarget = pd.read_csv('path/to/csv/4')
dfCiti = pd.read_csv('path/to/csv/5')
####################################################
combinedCapOne = pd.concat([dfCap1, dfCap2], ignore_index=True, sort=False)
CapitalOne = combinedCapOne[['Posted Date', 'Description','Category', 'Debit', 'Credit']]
dfAmex = dfAmex[['Date','Description', 'Amount']]
AmericanExpress = dfAmex.rename(columns={'Date':'Posted Date','Amount':'Debit'})
dfTarget = dfTarget[['Category','Merchant Name','Amount']]
Target = dfTarget.rename(columns={'Merchant Name': 'Description', 'Amount':'Debit'})
Target = Target.loc[Target['Description'].str.contains("AUTO PAYMENT") == False]
Target['Category'] = 'Merchandise'
Target['Debit'] = Target['Debit'].str.replace('$','')
Target['Debit'] = pd.to_numeric(Target['Debit'])
Citi = dfCiti[['Description', 'Debit', 'Credit']]
Citi['Category'] = 'Grocery'
Citi = Citi.loc[Citi['Description'].str.contains("PAYMENT") == False]
CitiNonGroc = Citi.loc[Citi['Description'].str.contains("SAMS|ALDI|PRICE CHOPPER|HY-VEE|SPROUTS|GLOBAL PRODUCE MARKET|WAL-MART") == False]
Citi.loc[CitiNonGroc.index, 'Category'] = 'Dining'
print(Citi)

CapAmex = pd.concat([CapitalOne, AmericanExpress], ignore_index=True, sort=False)
TargetCiti = pd.concat([Target, Citi], ignore_index=True, sort=False)
print(TargetCiti)
allCols = pd.concat([CapAmex, TargetCiti], ignore_index=True, sort=False)

allCols = allCols.loc[allCols['Description'].str.contains("AUTOPAY PAYMENT|SPOTIFY USA|SUBMETA|CAPITAL ONE AUTOPAY PYMT|COVE SMART|HOMESERVE USA|WILDFIRE BRAZILIAN|APPLE.COM|HLU") == False].reset_index()
allCols[['Category']] = allCols[['Category']].fillna('Merchandise')
totalCredits = allCols['Credit'].sum() * -1
allCols = allCols[allCols['Debit'].notnull()].reset_index()

## List of descriptions that get transformed into "Merchandise" somehow during importing as CSV from capital one
groceries = ['ALDI', 'SEA TO TABLE', 'HY-VEE', 'ORIENTAL SUPERMARKET', 'SAMSCLUB', 'GLOBAL PRODUCE MARKET', '888 INTERNATIONAL', 'PRICE CHOPPER', 'WAL-MART']

## Changing those transactions' Category to "Grocery"
grocery = allCols.loc[allCols['Description'].str.contains("ALDI|SEA TO TABLE|HY-VEE|ORIENTAL SUPERMARKET|SAMSCLUB|GLOBAL PRODUCE MARKET|888 INTERNATIONAL|PRICE CHOPPER|WAL-MART|SAMS CLUB", case=False)]
grocery = grocery.query('Category != "Gas/Automotive"')
grocery['Category'] = 'Grocery'
allCols.loc[grocery.index, 'Category'] = 'Grocery'


## Narrowing down columns
totalByCat = allCols[['Category', 'Debit']]
refined = allCols[['Posted Date', 'Category', 'Debit']]

## Adding a new row with "Total"
add_new_rows(totalByCat, 'Credits', totalCredits)
sum = totalByCat['Debit'].sum()
add_new_rows(totalByCat, 'Total', sum)

## Adding up by categories
totes = totalByCat.groupby(['Category']).sum()


## Adding up by categories AND dates
totesByDate = refined.groupby(['Category','Posted Date']).sum()

## Finding the highest expense 
maxSpend = allCols['Debit'].max()
maxDay = allCols[allCols['Debit'] == maxSpend]


## Creating a Dataframe with fixed expenses
fixedCategory = ['Income','Mortgage','WiFi','Utility','Travel Fund','School Fund','Savings','Investment','Donation','Phone bills','Car Insurance','Motorcycle Ins.','Property Tax (car)','Spotify','Apple&Hulu','Cove','BJJ']
fixedExpense = [${totalIncome}, ${mortgage}, 65.00, utilities, 200.00, 0.00, 3050.00, 1080.00, 50.00, 34.00, 117.00, 30.00, 21.45, 9.99, 10.98, 17.99, 99.00]
fixed = {'FixedCategory': fixedCategory, 'FixedExpense': fixedExpense}

dfFixed = pd.DataFrame(fixed)

## Get the total fixed expense

dfFixedTotal = dfFixed.loc[dfFixed['FixedCategory'].str.contains("Income|Travel Fund|Savings|Investment") == False]

## Get the total savings

dfSavings = dfFixed.loc[dfFixed['FixedCategory'].str.contains("Income|Travel Fund|Savings|Investment")]
savingsSum = dfSavings['FixedExpense'].sum()

## Get the Income AFTER savings

fixedCats = []
fixedExps = []
for cat in dfFixed['FixedCategory']:
    fixedCats.append(cat)
for exp in dfFixed['FixedExpense']:
    fixedExps.append(exp)

fixedObj = dict(zip(fixedCats, fixedExps))

dispIncome = round(fixedObj['Income'] - fixedObj['Travel Fund'] - fixedObj['Savings'] - fixedObj['Investment'], 1)

## Add the income after savings to dfFixed

add_new_rows(dfFixed, 'Disposable Income', dispIncome)

## Add the total fixed expense to the dfFixed
fixedSum = dfFixedTotal['FixedExpense'].sum()
add_new_rows(dfFixed, 'Total Fixed Expense', fixedSum)

## Adding the total variable expense & net expense to dfFixed
netExpense = sum + fixedSum
add_new_rows(dfFixed, 'Total Variable Expense', sum)
add_new_rows(dfFixed, 'Net Expense', netExpense)


## Get the difference bw disposable income and net expense, then update cells

remains = (dispIncome - (sum + fixedSum))
add_new_rows(dfFixed, 'Remaining', remains)
print(dfFixed)
fixed_cat_cell_list = thisWorksheet.range('A2:A23')
fixed_exp_cell_list = thisWorksheet.range('B2:B23')
for cell, val in enumerate(dfFixed.FixedCategory):
    fixed_cat_cell_list[cell].value = val
thisWorksheet.update_cells(fixed_cat_cell_list)
for cell, val in enumerate(dfFixed.FixedExpense):
    fixed_exp_cell_list[cell].value = val
thisWorksheet.update_cells(fixed_exp_cell_list)
thisWorksheet.update('A1', dfFixed.columns.values[0])
thisWorksheet.update('B1', dfFixed.columns.values[1])



####### Updating worksheet ######
## Setting specific cells to update


## Sorting so "Total" is at the bottom if necessary, then uploading to Google Drive

removed = totes.index.tolist()
removed.remove('Total')
removed.append('Total')
sorted = totes.reindex(removed).reset_index()
sorted = sorted.rename(columns={'Debit':'Expenses'})

## Form the cell id for Total row
lastCellCat = 'C' + str(len(sorted.index) + 1)
lastCellExp = 'D' + str(len(sorted.index) + 1)

## Form the cell id for one before Total row
oneLessLast = 'C' + str(len(sorted.index))

## Get the range of cells for updating
cell_list = thisWorksheet.range(f'C2:{lastCellCat}')
cell_list2 = thisWorksheet.range(f'D2:{lastCellExp}')

## Update the cells with columns
for cell, val in enumerate(sorted.Category):
    cell_list[cell].value = val

thisWorksheet.update('C1', sorted.columns.values[0])
thisWorksheet.update('D1', sorted.columns.values[1])

## Update the cells with categories and expenses
for cell, val in enumerate(sorted.Category):
    cell_list[cell].value = val

thisWorksheet.update_cells(cell_list)

for cell, val in enumerate(sorted.Expenses):
    cell_list2[cell].value = val

thisWorksheet.update_cells(cell_list2)


## Get the expenses without travel expenses, then update
catDict = []
expDict = []

for cat in sorted['Category']:
    catDict.append(cat)

for exp in sorted['Expenses']:
    expDict.append(exp)

variableDict = dict(zip(catDict, expDict))

if 'Other Travel' not in catDict:
    withoutTravel = (dispIncome - (sum + fixedSum))
else:
    withoutTravel = (dispIncome - (sum + fixedSum)) + variableDict['Other Travel']
thisWorksheet.update('C23', 'Without Travel Expense')
thisWorksheet.update('D23', withoutTravel)

### Formatting cells ###

format_cells(thisWorksheet, oneLessLast, remains, withoutTravel, lastCellCat, lastCellExp)

print("Order sorted. Success")
print(sorted)
