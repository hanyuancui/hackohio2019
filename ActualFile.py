import tkinter as tk
import statsmodels.api as sm
import xlrd
from pandas import DataFrame
from sklearn import linear_model

# Import the spreadsheet of the unemployment rate
locUR = "UR.xlsx"

# To open Workbook
wbUR = xlrd.open_workbook(locUR)

# Get Excel sheet NO.0
sheetUR = wbUR.sheet_by_index(0)

# Store unemployment rate into an array, from Jan 2009 to Dec 2019
arrUnEmRate = []
for i in range(12, 23):
    for j in range(1, 13):
        arrUnEmRate.append(sheetUR.cell_value(i, j))
arrUnEmRate.pop()
arrUnEmRate.pop()

# Store years into an array, 12 duplicates per year, from 2009 to 2019
arrYear = []
for i in range(2009, 2020):
    for j in range(0, 12):
        arrYear.append(i)
arrYear.pop()
arrYear.pop()

# Store months into an array, from Jan to Dec, total 10 years
arrMonth = []
for i in range(0, 11):
    for j in range(1, 13):
        arrMonth.append(j)
arrMonth.pop()
arrMonth.pop()

# Store Interest Rate into an array, from Jan 2009 to Dec 2019
# Import the spreadsheet of the interest rate
locIR = "IR.xlsx"

# To open Workbook
wbIR = xlrd.open_workbook(locIR)

# Get Excel sheet NO.0
sheetIR = wbIR.sheet_by_index(0)

# Store unemployment rate into an array, from Jan 2009 to Dec 2019
arrIntRate = []

for i in range(0, 132):
    arrIntRate.append(sheetIR.cell_value(i, 1))
arrIntRate.pop()
arrIntRate.pop()

# Store Stock Index Price into an array, from Jan 2009 to Dec 2019
arrStockIndex = []
for i in range(1001, 1133):
    arrStockIndex.append(i)
arrStockIndex.pop()
arrStockIndex.pop()

# Allocate arrays into a matrix
Stock_Market = {
    'Year': arrYear,
    'Month': arrMonth,
    'Interest_Rate': arrIntRate,
    'Unemployment_Rate': arrUnEmRate,
    'Stock_Index_Price': arrStockIndex
    }

df = DataFrame(Stock_Market, columns=['Year', 'Month', 'Interest_Rate', 'Unemployment_Rate', 'Stock_Index_Price'])

X = df[['Interest_Rate',
        'Unemployment_Rate']]
# here we have 2 input variables for multiple regression. If you just want to use one variable for simple linear
# regression, then use X = df['Interest_Rate'] for example.Alternatively, you may add additional variables
# within the brackets
Y = df['Stock_Index_Price']  # output variable (what we are trying to predict)

# with sklearn
regr = linear_model.LinearRegression()
regr.fit(X, Y)

print('Intercept: \n', regr.intercept_)
print('Coefficients: \n', regr.coef_)

# with stats models
X = sm.add_constant(X)  # adding a constant

model = sm.OLS(Y, X).fit()
predictions = model.predict(X)

# tkinter GUI
root = tk.Tk()

canvas1 = tk.Canvas(root, width=1200, height=450)
canvas1.pack()

# with sklearn
Intercept_result = ('Intercept: ', regr.intercept_)
label_Intercept = tk.Label(root, text=Intercept_result, justify='center')
canvas1.create_window(260, 220, window=label_Intercept)

# with sklearn
Coefficients_result = ('Coefficients: ', regr.coef_)
label_Coefficients = tk.Label(root, text=Coefficients_result, justify='center')
canvas1.create_window(260, 240, window=label_Coefficients)

# with stats models
print_model = model.summary()
label_model = tk.Label(root, text=print_model, justify='center', relief='solid', bg='LightSkyBlue1')
canvas1.create_window(800, 220, window=label_model)

# New_Interest_Rate label and input box
label1 = tk.Label(root, text='Type Interest Rate: ')
canvas1.create_window(100, 100, window=label1)

entry1 = tk.Entry(root)  # create 1st entry box
canvas1.create_window(270, 100, window=entry1)

# New_Unemployment_Rate label and input box
label2 = tk.Label(root, text=' Type Unemployment Rate: ')
canvas1.create_window(120, 120, window=label2)

entry2 = tk.Entry(root)  # create 2nd entry box
canvas1.create_window(270, 120, window=entry2)


def values():
    global new_interest_rate  # our 1st input variable
    new_Interest_Rate = float(entry1.get())

    global new_unemployment_rate  # our 2nd input variable
    new_Unemployment_Rate = float(entry2.get())

    prediction_result = ('Predicted Stock Index Price: ', regr.predict([[new_Interest_Rate, new_Unemployment_Rate]]))
    label_prediction = tk.Label(root, text=prediction_result, bg='orange')
    canvas1.create_window(260, 280, window=label_prediction)


button1 = tk.Button(root, text='Predict Stock Index Price', command=values,
                    bg='orange')  # button to call the 'values' command above
canvas1.create_window(270, 150, window=button1)

root.mainloop()
