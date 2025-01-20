# -*- coding: utf-8 -*-
"""
Created on Wed Apr  3 23:24:23 2024

@author: Mukul Kapri
"""

# -*- coding: utf-8 -*-
"""
Created on Wed Apr  3 19:06:58 2024

@author: Mukul Kapri
"""

##Part 1 : Calculating finabcial metrices#############################################
#####################################################################################
#####################################################################################

#Net profit Value>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

# Net profit value TCS ....................................................  
import pandas as pd

def net_profit_margin_Tcs(net_profit, total_revenue):
    # Convert net_profit and total_revenue to numeric types (float)
    net_profit = float(net_profit)
    total_revenue = float(total_revenue)
   
    # Calculate the net profit margin
    net_profit_margin = (net_profit / total_revenue) * 100
    return net_profit_margin

# Read data from Excel file
file_path = r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Tata\Tata_Inc.xlsx'  
df = pd.read_excel(file_path, sheet_name='Income statement')

# Iterate over 4 years
for year in range(2, 6):
    # Extract net profit and total revenue as strings for the current year
    net_profit_str = df.iloc[31, year]  
    total_revenue_str = df.iloc[0, year]  

    # Calculation of the net profit margin for the current year
    margin = net_profit_margin_Tcs(net_profit_str, total_revenue_str)

    print(f"The net profit margin of TCS for the year {2025 - year} is: {margin:.2f}%")

    
# Net profit value IBM ......................................................   
import pandas as pd

def net_profit_margin_ibm(net_profit, total_revenue):
    # Convert net_profit and total_revenue to numeric types (float)
    net_profit = float(net_profit)
    total_revenue = float(total_revenue)
   
    # Calculate the net profit margin
    net_profit_margin = (net_profit / total_revenue) * 100
    return net_profit_margin

# Read data from Excel file
file_path = r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Ibm\ibm_inc.xlsx'  
df = pd.read_excel(file_path, sheet_name='Income statement')

# Iterate over 4 years
for year in range(2, 6):
    # Extract net profit and total revenue as strings for the current year
    net_profit_str = df.iloc[32, year]  
    total_revenue_str = df.iloc[0, year]  

    # Calculation of the net profit margin for the current year
    margin = net_profit_margin_ibm(net_profit_str, total_revenue_str)

    print(f"The net profit margin of IBM for the year {2025 - year} is: {margin:.2f}%")    
    
# Net profit value Capgemini ......................................................   
import pandas as pd

def net_profit_margin_cap(net_profit, total_revenue):
    # Convert net_profit and total_revenue to numeric types (float)
    net_profit = float(net_profit)
    total_revenue = float(total_revenue)
   
    # Calculate the net profit margin
    net_profit_margin = (net_profit / total_revenue) * 100
    return net_profit_margin

# Read data from Excel file
file_path = r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Capgemini\Cap_Inc.xlsx'  
df = pd.read_excel(file_path, sheet_name='Income statement')

# Iterate over 4 years
for year in range(2, 6):
    # Extract net profit and total revenue as strings for the current year
    net_profit_str = df.iloc[32, year]  
    total_revenue_str = df.iloc[0, year]  

    # Calculate the net profit margin for the current year
    margin = net_profit_margin_cap(net_profit_str, total_revenue_str)

    print(f"The net profit margin of Capgemini for the year {2025 - year} is: {margin:.2f}%")    

# Gross profit Value>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

# Gross profit value TCS ....................................................  
import pandas as pd

def gross_profit_margin_Tcs(gross_profit, total_revenue):
    # Convert gross_profit and total_revenue to numeric types (float)
    gross_profit = float(gross_profit)
    total_revenue = float(total_revenue)
   
    # Calculate the gross profit margin
    gross_profit_margin = (gross_profit / total_revenue) * 100
    return gross_profit_margin

# Read data from Excel file
file_path = r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Tata\Tata_Inc.xlsx'  
df = pd.read_excel(file_path, sheet_name='Income statement')

# Iterate over 4 years
for year in range(2, 6):
    # Extract net profit and total revenue as strings for the current year
    gross_profit_str = df.iloc[2, year]  
    total_revenue_str = df.iloc[0, year]  

    # Calculation of the gross profit margin for the current year
    margin = gross_profit_margin_Tcs(gross_profit_str, total_revenue_str)

    print(f"The gross profit margin of TCS for the year {2025 - year} is: {margin:.2f}%")
    
# Gross profit value IBM ......................................................   
import pandas as pd

def gross_profit_margin_ibm(gross_profit, total_revenue):
    # Convert GROSS_profit and total_revenue to numeric types (float)
    gross_profit = float(gross_profit)
    total_revenue = float(total_revenue)
   
    # Calculate the gross profit margin
    gross_profit_margin = (gross_profit / total_revenue) * 100
    return gross_profit_margin

# Read data from Excel file
file_path = r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Ibm\ibm_inc.xlsx'  
df = pd.read_excel(file_path, sheet_name='Income statement')

# Iterate over 4 years
for year in range(2, 6):
    # Extract net profit and total revenue as strings for the current year
    gross_profit_str = df.iloc[2, year]  
    total_revenue_str = df.iloc[0, year]  

    # Calculation of the gross profit margin for the current year
    margin = gross_profit_margin_ibm(gross_profit_str, total_revenue_str)

    print(f"The gross profit margin of {2025 - year} is:{margin:.2f}%")   
    
# Gross profit value Capgemini ......................................................   
import pandas as pd

def gross_profit_margin_cap(gross_profit, total_revenue):
    # Convert GROSS_profit and total_revenue to numeric types (float)
    gross_profit = float(gross_profit)
    total_revenue = float(total_revenue)
   
    # Calculate the gross profit margin
    gross_profit_margin = (gross_profit / total_revenue) * 100
    return gross_profit_margin

# Read data from Excel file
file_path = r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Capgemini\Cap_Inc.xlsx'  
df = pd.read_excel(file_path, sheet_name='Income statement')

# Iterate over 4 years
for year in range(2, 6):
    # Extract net profit and total revenue as strings for the current year
    gross_profit_str = df.iloc[2, year]  
    total_revenue_str = df.iloc[0, year]  

    # Calculate the gross profit margin for the current year
    margin = gross_profit_margin_cap(gross_profit_str, total_revenue_str)

    print(f"The gross profit margin of Capgemini for {2025 - year} is: {margin:.2f}%")  


#ROA>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
# ROA value TATA ....................................................  
import pandas as pd

def roa_Tcs(net_income, total_assets):
    # Convert net_income and total_assets to numeric types (float)
    net_income = float(net_income)
    total_assets = float(total_assets)
    
    # Calculate the ROA
    roa= (net_income / total_assets) 
    return roa

# Read data from Excel file
file_path = r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Tata\Tata_Inc.xlsx'  
df = pd.read_excel(file_path, sheet_name='Income statement')

# Iterate over 4 years
for year in range(2, 6):
    # Extract net income and total assets as strings for the current year
    net_income_str = df.iloc[34, year]  
    total_assets_str = df.iloc[44, year]  

    # Calculating the ROA for the current year
    margin = roa_Tcs(net_income_str, total_assets_str)

    print(f"The ROA( Return on assets) of TCS for {2025 - year} is: {margin:.2f}%")
    
# ROA value IBM ......................................................   
import pandas as pd

def roa_ibm(net_income, total_assets):
    # Convert net_income and total_assets to numeric types (float)
    net_income = float(net_income)
    total_assets = float(total_assets)
   
    # Calculate the ROA 
    roa= (net_income / total_assets) 
    return roa

# Read data from Excel file
file_path = r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Ibm\ibm_inc.xlsx'  
df = pd.read_excel(file_path, sheet_name='Income statement')

# Iterate over 4 years
for year in range(2, 6):
    # Extract net income and total assets as strings for the current year
    net_income_str = df.iloc[35, year]  
    total_assets_str = df.iloc[48, year]  

    # Calculating the net profit margin for the current year
    margin = roa_ibm(net_income_str, total_assets_str)

    print(f"The ROA( Return on assets) of IBM for {2025 - year} is: {margin:.2f}%")
    
# ROA Capgemini ......................................................   
import pandas as pd

def roa_cap(net_income, total_assets):
    # Convert net_income and total_assets to numeric types (float)
    net_income = float(net_income)
    total_assets = float(total_assets)
  
    # Calculate the ROA 
    roa= (net_income / total_assets) 
    return roa

# Read data from Excel file
file_path = r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Capgemini\Cap_Inc.xlsx'  
df = pd.read_excel(file_path, sheet_name='Income statement')

# Iterate over 4 years
for year in range(2, 6):
    # Extract net income and total assets as strings for the current year
    net_income_str = df.iloc[35, year]  
    total_assets_str = df.iloc[49, year]  

    # Calculate the ROA for the current year
    margin = roa_cap(net_income_str, total_assets_str)

    print(f"The ROA( Return on assets) of Capgemini for {2025 - year} is: {margin:.2f}%")

#ROE>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

# ROE value TCS ....................................................  
import pandas as pd

def roe_Tcs(net_income, shareholder_equity):
    # Convert net_income and shareholder_equity to numeric types (float)
    net_income = float(net_income)
    shareholder_equity = float(shareholder_equity)
    
    # Calculate the net profit margin
    roe= (net_income / shareholder_equity) 
    return roe

# Read data from Excel file
file_path = r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Tata\Tata_Inc.xlsx'  
df = pd.read_excel(file_path, sheet_name='Income statement')

# Iterate over 4 years
for year in range(2, 6):
    # Extract net income and shareholder equity as strings for the current year
    net_income_str = df.iloc[34, year]  
    shareholder_equity_str = df.iloc[59, year]  

    # Calculate the ROE for the current year
    margin = roe_Tcs(net_income_str, shareholder_equity_str)

    print(f"The ROE( Return on equity) of TCS for{2025 - year} is: {margin:.2f}%")
    
# ROE value Ibm ....................................................  
import pandas as pd

def roe_ibm(net_income, shareholder_equity):
    # Convert net_income and shareholder_equity to numeric types (float)
    net_income = float(net_income)
    shareholder_equity = float(shareholder_equity)
    
    # Calculate the net profit margin
    roe= (net_income / shareholder_equity) 
    return roe

# Read data from Excel file
file_path = r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Ibm\ibm_inc.xlsx'  
df = pd.read_excel(file_path, sheet_name='Income statement')

# Iterate over 4 years
for year in range(2, 6):
    # Extract net income and shareholder equity as strings for the current year
    net_income_str = df.iloc[35, year]  
    shareholder_equity_str = df.iloc[62, year]  

    # Calculate the ROE for the current year
    margin = roe_ibm(net_income_str, shareholder_equity_str)

    print(f"The ROE( Return on equity) of IBM for {2025 - year} is: {margin:.2f}%")
    
# ROE value Capgemini ....................................................  
import pandas as pd

def roe_cap(net_income, shareholder_equity):
    # Convert net_income and shareholder_equity to numeric types (float)
    net_income = float(net_income)
    shareholder_equity = float(shareholder_equity)
    
    # Calculate the ROE
    roe= (net_income / shareholder_equity) 
    return roe

# Read data from Excel file
file_path = r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Capgemini\Cap_Inc.xlsx'  
df = pd.read_excel(file_path, sheet_name='Income statement')

# Iterate over 4 years
for year in range(2, 6):
    # Extract net income and shareholder equity as strings for the current year
    net_income_str = df.iloc[35, year]  
    shareholder_equity_str = df.iloc[65, year]  

    # Calculate the ROE for the current year
    margin = roe_cap(net_income_str, shareholder_equity_str)

    print(f"The ROE( Return on equity) of Capgemini for {2025 - year} is: {margin:.2f}%")

# 5 Current ratios>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

# Current ratio TCS....................................................  
import pandas as pd

def current_ratio_Tcs(Current_assets, Current_liabilities):
    # Convert current assets and current liabilities to numeric types (float)
    Current_assets = float(Current_assets)
    Current_liabilities = float(Current_liabilities)
    
    # Calculate the Current ratio
    curratio= (Current_assets / Current_liabilities) 
    return curratio

# Read data from Excel file
file_path = r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Tata\Tata_Inc.xlsx'  
df = pd.read_excel(file_path, sheet_name='Income statement')

# Iterate over 4 years
for year in range(2, 6):
    # Extract current assets and current liabilities as strings for the current year
    Current_assets_str = df.iloc[64, year]  
    Current_liabilities_str = df.iloc[67, year]  

    # Calculate the Current ratio for the current year
    margin = current_ratio_Tcs(Current_assets_str, Current_liabilities_str)

    print(f"The Current ratio of TCS for {2025 - year} is: {margin:.2f}%")

# Current ratio ibm....................................................  
import pandas as pd

def current_ratio_ibm(Current_assets, Current_liabilities):
    # Convert current assets and current liabilities to numeric types (float)
    Current_assets = float(Current_assets)
    Current_liabilities = float(Current_liabilities)
    
    # Calculate the Current ratio
    curratio= (Current_assets / Current_liabilities) 
    return curratio

# Read data from Excel file
file_path = r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Ibm\ibm_inc.xlsx'  
df = pd.read_excel(file_path, sheet_name='Income statement')

# Iterate over 4 years
for year in range(2, 6):
    # Extract current assets and current liabilities as strings for the current year
    Current_assets_str = df.iloc[69, year]  
    Current_liabilities_str = df.iloc[72, year]  

    # Calculate the Current ratio for the current year
    margin = current_ratio_ibm(Current_assets_str, Current_liabilities_str)

    print(f"The Current ratio of IBM for {2025 - year} is: {margin:.2f}%")
    
# Current ratio Capgemini se....................................................  
import pandas as pd

def current_ratio_cap(Current_assets, Current_liabilities):
    # Convert current assets and current liabilities to numeric types (float)
    Current_assets = float(Current_assets)
    Current_liabilities = float(Current_liabilities)
    
    # Calculate the current ratio
    curratio= (Current_assets / Current_liabilities) 
    return curratio

# Read data from Excel file
file_path = r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Capgemini\Cap_Inc.xlsx'  
df = pd.read_excel(file_path, sheet_name='Income statement')

# Iterate over 4 years
for year in range(2, 6):
    # Extract current assets and current liabilities as strings for the current year
    Current_assets_str = df.iloc[73, year]  
    Current_liabilities_str = df.iloc[76, year]  

    # Calculate the current ratio for the current year
    margin = current_ratio_cap(Current_assets_str, Current_liabilities_str)

    print(f"The Current ratio of Capgemini for {2025 - year} is: {margin:.2f}%")  

#Quick Ratio>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

# Quick ratio tcs....................................................  
import pandas as pd

def quick_ratio_Tcs(Current_assets, Current_liabilities, inventory):
    # Convert current assets, inventory and current liabilities to numeric types (float)
    Current_assets = float(Current_assets)
    Current_liabilities = float(Current_liabilities)
    inventory= float(inventory_str)
    
    # Calculate the quick ratio
    quick_ratio = (Current_assets - inventory) / Current_liabilities
    return quick_ratio

# Read data from Excel file
file_path = r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Tata\Tata_Inc.xlsx'  
df = pd.read_excel(file_path, sheet_name='Income statement')

# Iterate over 4 years
for year in range(2, 6):
    # Extract current assets, current liabilities and inventory as strings for the current year
    Current_assets_str = df.iloc[64, year]  
    Current_liabilities_str = df.iloc[67, year] 
    inventory_str= df.iloc [74, year] 

    # Calculate the Quick ratio for the current year
    margin = quick_ratio_Tcs(Current_assets_str, Current_liabilities_str, inventory_str)

    print(f"The Quick ratio of TCS for {2025 - year} is: {margin:.2f}%")  
    
# Quick ratio ibm....................................................  
import pandas as pd

def quick_ratio_ibm(Current_assets, Current_liabilities, inventory):
    # Convert current assets, inventory and current liabilities to numeric types (float)
    Current_assets = float(Current_assets)
    Current_liabilities = float(Current_liabilities)
    inventory= float(inventory_str)
    
    # Calculate the quick ratio
    quick_ratio = (Current_assets - inventory) / Current_liabilities
    return quick_ratio

# Read data from Excel file
file_path = r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Ibm\ibm_inc.xlsx'  
df = pd.read_excel(file_path, sheet_name='Income statement')

# Iterate over 4 years
for year in range(2, 6):
    # Extract current assets, current liabilities and inventory as strings for the current year
    Current_assets_str = df.iloc[69, year]  
    Current_liabilities_str = df.iloc[72, year] 
    inventory_str= df.iloc [78, year] 

    # Calculate the quick ratio for the current year
    margin = quick_ratio_ibm(Current_assets_str, Current_liabilities_str, inventory_str)

    print(f"The Quick ratio of {2025 - year} is: {margin:.2f}%")  
    

# Quick ratio capgemini....................................................  
import pandas as pd

def quick_ratio_cap(Current_assets, Current_liabilities, inventory):
    # Convert current assets, inventory and current liabilities to numeric types (float)
    Current_assets = float(Current_assets)
    Current_liabilities = float(Current_liabilities)
    inventory= float(inventory_str)
    
    # Calculate the Quick ratio
    quick_ratio = (Current_assets - inventory) / Current_liabilities
    return quick_ratio

# Read data from Excel file
file_path = r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Capgemini\Cap_Inc.xlsx'  
df = pd.read_excel(file_path, sheet_name='Income statement')

# Iterate over 4 years
for year in range(2, 6):
    # Extract current assets, current liabilities and inventory as strings for the current year
    Current_assets_str = df.iloc[73, year]  
    Current_liabilities_str = df.iloc[76, year] 
    inventory_str= df.iloc [83, year] 

    # Calculate the quick ratio for the current year
    margin = quick_ratio_cap(Current_assets_str, Current_liabilities_str, inventory_str)

    print(f"The Quick ratio of Capgemini for {2025 - year} is: {margin:.2f}%")  

# 7 debt to equity ratio>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

# debt to equity ratio tata....................................................  
import pandas as pd

def debt_to_equity_ratio_tcs(total_debt, shareholder_equity):
    # Convert total debt and shareholder equity to numeric types (float)
    total_debt = float(total_debt)
    shareholder_equity = float(shareholder_equity)
    
    
    # Calculate the debt to equity ratio
    debt_to_equity_ratio = total_debt / shareholder_equity
    return debt_to_equity_ratio

# Read data from Excel file
file_path = r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Tata\Tata_Inc.xlsx'  
df = pd.read_excel(file_path, sheet_name='Income statement')

# Iterate over 4 years
for year in range(2, 6):
    # Extract total debt and shareholder equity as strings for the current year
    total_debt_str = df.iloc[54, year]  
    shareholder_equity_str = df.iloc[59, year] 
    

    # Calculate the debt to equity ratio for the current year
    margin = debt_to_equity_ratio_tcs(total_debt_str, shareholder_equity_str)

    print(f"The Current ratio of TCS for {2025 - year} is: {margin:.2f}%")  
    
# debt to equity ratio IBM....................................................  
import pandas as pd

def debt_to_equity_ratio_ibm(total_debt, shareholder_equity):
    # Convert total debt and shareholder equity to numeric types (float)
    total_debt = float(total_debt)
    shareholder_equity = float(shareholder_equity)
    
    
    # Calculate the debt to equity ratio
    debt_to_equity_ratio = total_debt / shareholder_equity
    return debt_to_equity_ratio

# Read data from Excel file
file_path = r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Ibm\ibm_inc.xlsx'  
df = pd.read_excel(file_path, sheet_name='Income statement')

# Iterate over 4 years
for year in range(2, 6):
    # Extract total debt and shareholder's equity as strings for the current year
    total_debt_str = df.iloc[57, year]  
    shareholder_equity_str = df.iloc[62, year] 
    

    # Calculate the debt to equityratio for the current year
    margin = debt_to_equity_ratio_ibm(total_debt_str, shareholder_equity_str)

    print(f"The Current ratio of IBM for {2025 - year} is: {margin:.2f}%")  

# debt to equity ratio capgemini....................................................  
import pandas as pd

def debt_to_equity_ratio_cap(total_debt, shareholder_equity):
    # Convert total debt and shareholder equity to numeric types (float)
    total_debt = float(total_debt)
    shareholder_equity = float(shareholder_equity)
    
    
    # Calculate the debt to equity ratio
    debt_to_equity_ratio = total_debt / shareholder_equity
    return debt_to_equity_ratio

# Read data from Excel file
file_path = r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Capgemini\Cap_Inc.xlsx'  
df = pd.read_excel(file_path, sheet_name='Income statement')

# Iterate over 4 years
for year in range(2, 6):
    # Extract total debt and shareholder equity as strings for the current year
    total_debt_str = df.iloc[59, year]  
    shareholder_equity_str = df.iloc[65, year] 
    

    # Calculate the debt to equity ratio for the current year
    margin = debt_to_equity_ratio_cap(total_debt_str, shareholder_equity_str)

    print(f"The Debt to Equity ratio of Capgemini for {2025 - year} is: {margin:.2f}%")  

#Debt ratio>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

#Debt ratio of TCS.........................................................

import pandas as pd

def debt_ratio_tcs(total_debt, total_assets):
    # Convert total debt and total assets to numeric types (float)
    total_debt = float(total_debt)
    total_assets = float(total_assets)
    
    # Calculate the debt ratio
    debt_ratio= (total_debt / total_assets) 
    return debt_ratio

# Read data from Excel file
file_path = r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Tata\Tata_Inc.xlsx'  
df = pd.read_excel(file_path, sheet_name='Income statement')

# Iterate over 4 years
for year in range(2, 6):
    # Extract total debt and total assets as strings for the current year
    total_debt_str = df.iloc[54, year]  
    total_assets_str = df.iloc[44, year]  

    # Calculate the debt ratio for the current year
    margin = debt_ratio_tcs(total_debt_str, total_assets_str)

    print(f"The debt ratio of TCS for {2025 - year} is: {margin:.2f}%")  
    
#Debt ratio of ibm.........................................................

import pandas as pd

def debt_ratio_ibm(total_debt, total_assets):
    # Convert total debt and total assets to numeric types (float)
    total_debt = float(total_debt)
    total_assets = float(total_assets)
    
    # Calculate the debt ratio
    debt_ratio= (total_debt / total_assets) 
    return debt_ratio

# Read data from Excel file
file_path = r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Ibm\ibm_inc.xlsx'  
df = pd.read_excel(file_path, sheet_name='Income statement')

# Iterate over 4 years
for year in range(2, 6):
    # Extract total debt and total assets as strings for the current year
    total_debt_str = df.iloc[57, year]  
    total_assets_str = df.iloc[48, year]  

    # Calculate the debt ratio for the current year
    margin = debt_ratio_ibm(total_debt_str, total_assets_str)

    print(f"The debt ratio of IBM for {2025 - year} is: {margin:.2f}%")  
    
#Debt ratio of capgemini.........................................................

import pandas as pd

def debt_ratio_cap(total_debt, total_assets):
    # Convert total debt and total assets to numeric types (float)
    total_debt = float(total_debt)
    total_assets = float(total_assets)
    
    # Calculate debt ratio
    debt_ratio= (total_debt / total_assets) 
    return debt_ratio

# Read data from Excel file
file_path = r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Capgemini\Cap_Inc.xlsx'  
df = pd.read_excel(file_path, sheet_name='Income statement')

# Iterate over 4 years
for year in range(2, 6):
    # Extract total debt and total assets as strings for the current year
    total_debt_str = df.iloc[59, year]  
    total_assets_str = df.iloc[49, year]  

    # Calculate the debt ratio for the current year
    margin = debt_ratio_cap(total_debt_str, total_assets_str)

    print(f"The debt ratio of Capgemini for {2025 - year} is: {margin:.2f}%")  

#Growth Rate>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
# Growth ratio TCS>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

import pandas as pd
def revenue_growth_ratio_tcs(Current_revenue, previous_revenue):
    # Convert current year revenue and previous year revenue to numeric types (float)
    Current_revenue = float(Current_revenue)
    previous_revenue = float(previous_revenue)
    
    # Calculate the growth rate
    growth_rate = ((Current_revenue - previous_revenue) / previous_revenue)*100
    return growth_rate

# Read data from Excel file
file_path = r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Tata\Tata_Inc.xlsx'  
df = pd.read_excel(file_path, sheet_name='Income statement')

# Iterate over 4 years
for year in range(2, 5):
    # Extract current year revenue and previous year revenue as strings for the current year
    Current_revenue_str = df.iloc[0, year]  
    previous_revenue_str = df.iloc[0, year + 1]  

# Calculate the growth rate for the current year
    growth_rate = revenue_growth_ratio_tcs(Current_revenue_str, previous_revenue_str)
  
 # Check if growth rate is negative 
    if growth_rate < 0:
        growth_rate_str = f"{growth_rate:.2f}% (Decrease)"
    else:
        growth_rate_str = f"{growth_rate:.2f}% (Increase)"

    print(f"The growth rate of TCS for year {2025 - year} is: {growth_rate_str}")

# Growth ratio ibm....................................................  
import pandas as pd

def growth_ratio_ibm(Current_revenue, previous_revenue):
    # Convert current year revenue and previous year revenue to numeric types (float)
    Current_revenue = float(Current_revenue)
    previous_revenue = float(previous_revenue)
    
    # Calculate the growth rate
    growth_rate = ((Current_revenue - previous_revenue) / previous_revenue) *100 
    return growth_rate

# Read data from Excel file
file_path = r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Ibm\ibm_inc.xlsx'  
df = pd.read_excel(file_path, sheet_name='Income statement')

# Iterate over 4 years
for year in range(2, 5):
    # Extract current year revenue and previous year revenue as strings for the current year
    Current_revenue_str = df.iloc[0, year]  
    previous_revenue_str = df.iloc[0, year + 1]  
   
# Calculate the growth rate for the current year
    growth_rate = growth_ratio_ibm(Current_revenue_str, previous_revenue_str)
  
 # Check if growth rate is negative 
    if growth_rate < 0:
        growth_rate_str = f"{growth_rate:.2f}% (Decrease)"
    else:
        growth_rate_str = f"{growth_rate:.2f}% (Increase)"

    print(f"The growth rate of IBM for year {2025 - year} is: {growth_rate_str}")
    
# Growth ratio capgemini....................................................  
import pandas as pd

def growth_ratio_cap(current_revenue, previous_revenue):
    # Convert current year revenue and previous year revenue to numeric types (float)
    current_revenue = float(current_revenue)
    previous_revenue = float(previous_revenue)
    
    # Calculate the growth rate
    growth_rate = ((current_revenue - previous_revenue) / previous_revenue) *100
    return growth_rate

# Read data from Excel file
file_path = r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Capgemini\Cap_Inc.xlsx'  
df = pd.read_excel(file_path, sheet_name='Income statement')

# Iterate over 3 years (range(2, 5) will iterate over 2, 3, 4)
for year in range(2, 5):
    # Extract current year revenue and previous year revenue as strings for the current year
    current_revenue_str = df.iloc[0, year]  
    previous_revenue_str = df.iloc[0, year + 1]  

    # Calculate the growth rate for the current year
    growth_rate = growth_ratio_cap(current_revenue_str, previous_revenue_str)

    # Check if growth rate is negative
    if growth_rate < 0:
        growth_rate_str = f"{growth_rate:.2f}% (Decrease)"
    else:
        growth_rate_str = f"{growth_rate:.2f}% (Increase)"

    print(f"The growth rate of Capgemnini for year {2025 - year} is: {growth_rate_str}")

# P/E ratio>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

#P/E ratio- TCS........................................................................................
import pandas as pd

def pe_ratio_Tcs(mps,eps):
    # Convert market price per share and earnings per share to numeric types (float)
    mps = float(mps)
    eps = float(eps)
    
    # Calculate the P/E ratio
    pe_ratio= (mps / eps) 
    return pe_ratio

# Read data from Excel file
file_path = r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Tata\Tata_Inc.xlsx'  
df = pd.read_excel(file_path, sheet_name='Income statement')

# Iterate over 4 years
for year in range(2, 6):
    # Extract market price per share and earnings per share as strings for the current year
    mps_str = df.iloc[82, year]  
    eps_str = df.iloc[83, year]  

    # Calculate the p/e ratio for the current year
    margin = pe_ratio_Tcs(mps_str,eps_str)

    print(f"The Price to earnings ratio of TCS for {2025 - year} is: {margin:.2f}%")

# P/E ratio- ibm.............................................................................. 
import pandas as pd

def pe_ratio_ibm(mps,eps):
    # Convert market price per share and earnings per share to numeric types (float)
    mps = float(mps)
    eps = float(eps)
    
    # Calculate the p/e ratio
    pe_ratio= (mps / eps) 
    return pe_ratio

# Read data from Excel file
file_path = r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Ibm\ibm_inc.xlsx'  
df = pd.read_excel(file_path, sheet_name='Income statement')

# Iterate over 4 years
for year in range(2, 6):
    # Extract market price per share and earnings per share as strings for the current year
    mps_str = df.iloc[87, year]  
    eps_str = df.iloc[88, year]  

    # Calculate the p/e ratio for the current year
    margin = pe_ratio_ibm(mps_str,eps_str)

    print(f"The Price to earnings ratio of IBM for {2025 - year} is: {margin:.2f}%")
    
# P/E ratio- capgemini.......................................................................  
import pandas as pd

def pe_ratio_cap(mps,eps):
    # Convert market price per share and earnings per share to numeric types (float)
    mps = float(mps)
    eps = float(eps)
    
    # Calculate the p/e ratio
    pe_ratio= (mps / eps) 
    return pe_ratio

# Read data from Excel file
file_path = r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Capgemini\Cap_Inc.xlsx'  
df = pd.read_excel(file_path, sheet_name='Income statement')

# Iterate over 4 years
for year in range(2, 6):
    # Extract market price per share and earnings per share as strings for the current year
    mps_str = df.iloc[92, year]  
    eps_str = df.iloc[93, year]  

    # Calculate the market price per share and earnings per share for the current year
    margin = pe_ratio_cap(mps_str,eps_str)

    print(f"The Price to earnings ratio of Capgemini for {2025 - year} is: {margin:.2f}%")

# P/B ratio>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

#P/E ratio- TCS
import pandas as pd

def pb_ratio_Tcs(mps,bv):
    # Convert market price per share and book value per share to numeric types (float)
    mps = float(mps)
    bv = float(bv)
    
    # Calculate the P/B ratio
    pb_ratio= (mps / bv) 
    return pb_ratio

# Read data from Excel file
file_path = r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Tata\Tata_Inc.xlsx'  
df = pd.read_excel(file_path, sheet_name='Income statement')

# Iterate over 4 years
for year in range(2, 6):
    # Extract market price per share and book value per share as strings for the current year
    mps_str = df.iloc[82, year]  
    bv_str = df.iloc[85, year]  

    # Calculate the P/E ratio for the current year
    margin = pb_ratio_Tcs(mps_str,bv_str)

    print(f"The Price to book ratio of TCS for {2025 - year} is: {margin:.2f}%")

#P/B ratio- ibm..................................................................................
import pandas as pd

def pb_ratio_ibm(mps,bv):
    # Convert market price per share and book value per share to numeric types (float)
    mps = float(mps)
    bv = float(bv)
    
    # Calculate the P/B ratio
    pb_ratio= (mps / bv) 
    return pb_ratio

# Read data from Excel file
file_path = r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Ibm\ibm_inc.xlsx'  
df = pd.read_excel(file_path, sheet_name='Income statement')

# Iterate over 4 years
for year in range(2, 6):
    # Extract market price per share and book value per share as strings for the current year
    mps_str = df.iloc[87, year]  
    bv_str = df.iloc[90, year]  

    # Calculate the P/B ratio for the current year
    margin = pb_ratio_ibm(mps_str,bv_str)

    print(f"The Price to book ratio of IBM for {2025 - year} is: {margin:.2f}%")
    
#P/B ratio- capgemini........................................................................
import pandas as pd

def pb_ratio_cap(mps,bv):
    # Convert market price per share and book value per share to numeric types (float)
    mps = float(mps)
    bv = float(bv)
    
    # Calculate the P/B ratio
    pb_ratio= (mps / bv) 
    return pb_ratio

# Read data from Excel file
file_path = r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Capgemini\Cap_Inc.xlsx'  
df = pd.read_excel(file_path, sheet_name='Income statement')

# Iterate over 4 years
for year in range(2, 6):
    # Extract market price per share and book value per share as strings for the current year
    mps_str = df.iloc[92, year]  
    bv_str = df.iloc[95, year]  

    # Calculate the P/B ratio for the current year
    margin = pb_ratio_cap(mps_str,bv_str)

    print(f"The Price to book ratio of Capgemini for {2025 - year} is: {margin:.2f}%")

# Dividend yield>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

# Dividend yield ratio- TCS........................................................................
import pandas as pd

def dividend_yield_Tcs(ds,mps):
    # Convert dividend per share and market price per share to numeric types (float)
    ds = float(ds)
    mps = float(mps)
    
    # Calculate the dividend yield
    dividend_yield= (ds / mps) *100 
    return dividend_yield

# Read data from Excel file
file_path = r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Tata\Tata_Inc.xlsx'  
df = pd.read_excel(file_path, sheet_name='Income statement')

# Iterate over 4 years
for year in range(2, 6):
    # Extract dividend per share and market price per share as strings for the current year
    ds_str = df.iloc[84, year]  
    mps_str = df.iloc[82, year]  

    # Calculate the Dividend yield for the current year
    margin = dividend_yield_Tcs(ds_str,mps_str)

    print(f"The dividend yield of TCS for {2025 - year} is: {margin:.2f}%")

# Dividend yield ratio- ibm..........................................................................
import pandas as pd

def dividend_yield_ibm(ds,mps):
    # Convert dividend per share and market price per share to numeric types (float)
    ds = float(ds)
    mps = float(mps)
    
    # Calculate the dividend yield 
    dividend_yield= (ds / mps) *100
    return dividend_yield

# Read data from Excel file
file_path = r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Ibm\ibm_inc.xlsx'  
df = pd.read_excel(file_path, sheet_name='Income statement')

# Iterate over 4 years
for year in range(2, 6):
    # Extract dividend per share and market price per share as strings for the current year
    ds_str = df.iloc[89, year]  
    mps_str = df.iloc[87, year]  

    # Calculate dividend yield for the current year
    margin = dividend_yield_ibm(ds_str,mps_str)

    print(f"The dividend yield of IBM for {2025 - year} is: {margin:.2f}%")
    
# Dividend yield ratio- capgemini...................................................................
import pandas as pd

def dividend_yield_cap(ds,mps):
    # Convert dividend per share and market price per share to numeric types (float)
    ds = float(ds)
    mps = float(mps)
    
    # Calculate the dividend yield
    dividend_yield= (ds / mps) *100
    return dividend_yield

# Read data from Excel file
file_path = r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Capgemini\Cap_Inc.xlsx'  
df = pd.read_excel(file_path, sheet_name='Income statement')

# Iterate over 4 years
for year in range(2, 6):
    # Extract dividend per share and market price per share as strings for the current year
    ds_str = df.iloc[94, year]  
    mps_str = df.iloc[92, year]  

    # Calculate the dividend yield for the current year
    margin = dividend_yield_cap(ds_str,mps_str)

    print(f"The dividend yield of Capgemini for {2025 - year} is: {margin:.2f}%")

# Inventory turnout ratio>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

#Tcs.........................................................................................

import pandas as pd

def Inventory_turnout_ratio_Tcs(cogs, avg_inventory):
    # Convert cost of goods sold and average inventory cost to numeric types (float)
    cogs = float(cogs)
    avg_inventory = float(avg_inventory)
    
    # Calculate the inventory turnout ratio
    inventory_turnout= (cogs / avg_inventory) 
    return inventory_turnout

# Read data from Excel file
file_path = r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Tata\Tata_Inc.xlsx'  
df = pd.read_excel(file_path, sheet_name='Income statement')

# Iterate over 4 years
for year in range(2, 6):
    # Extract cost of goods sold and average inventory cost as strings for the current year
    cogs_str = df.iloc[92, year]  
    avg_inventory_str = df.iloc[94, year]  

    # Calculate the inventory turnout ratio for the current year
    margin = Inventory_turnout_ratio_Tcs(cogs_str, avg_inventory_str)

    print(f"The Inventory turnout ratio of TCS for {2025 - year} is: {margin:.2f}%")

#IBM......................................................................................

import pandas as pd

def Inventory_turnout_ratio_ibm(cogs, avg_inventory):
    # Convert cost of goods sold and average inventory cost to numeric types (float)
    cogs = float(cogs)
    avg_inventory = float(avg_inventory)
    
    # Calculate the Inventory turnout ratio
    inventory_turnout= (cogs / avg_inventory) 
    return inventory_turnout

# Read data from Excel file
file_path = r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Ibm\ibm_inc.xlsx'  
df = pd.read_excel(file_path, sheet_name='Income statement')

# Iterate over 4 years
for year in range(2, 6):
    # Extract cost of goods sold and average inventory cost as strings for the current year
    cogs_str = df.iloc[97, year]  
    avg_inventory_str = df.iloc[99, year]  

    # Calculate the inventory turnout ratio for the current year
    margin = Inventory_turnout_ratio_ibm(cogs_str, avg_inventory_str)

    print(f"The Inventory turnout ratio of IBM for {2025 - year} is: {margin:.2f}%")

#capgemini (no inventory details avaliable, hence cannot calculate).............................................

# Asset Turonver ratio>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

#TCS..........................................................................................................

import pandas as pd

def asset_turnover_ratio_Tcs(net_sales, avg_total_assets):
    # Convert net sales and average total assets to numeric types (float)
    net_sales = float(net_sales)
    avg_total_assets = float(avg_total_assets)
    
    # Calculate the asset turnover ratio
    asset_turnover_ratio= (net_sales / avg_total_assets) 
    return asset_turnover_ratio

# Read data from Excel file
file_path = r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Tata\Tata_Inc.xlsx'  
df = pd.read_excel(file_path, sheet_name='Income statement')

# Iterate over 4 years
for year in range(2, 6):
    # Extract net sales and average total assets as strings for the current year
    net_sales_str = df.iloc[0, year]  
    avg_total_assets = df.iloc[95, year]

    # Calculate asset turnover ratio for the current year
    margin = asset_turnover_ratio_Tcs(net_sales_str, avg_total_assets)

    print(f"The Asset Turnover ratio of TCS for {2025 - year} is: {margin:.2f}%")

#IBM....................................................................................................

import pandas as pd

def asset_turnover_ratio_ibm(net_sales, avg_total_assets):
    # Convert net sales and average total assets to numeric types (float)
    net_sales = float(net_sales)
    avg_total_assets = float(avg_total_assets)
    
    # Calculate the asset turnover ratio
    asset_turnover_ratio= (net_sales / avg_total_assets) 
    return asset_turnover_ratio

# Read data from Excel file
file_path = r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Ibm\ibm_inc.xlsx'  
df = pd.read_excel(file_path, sheet_name='Income statement')

# Iterate over 4 years
for year in range(2, 6):
    # Extract net sales and average total assets as strings for the current year
    net_sales_str = df.iloc[0, year]  
    avg_total_assets = df.iloc[100, year]

    # Calculate asset turnover ratio for the current year
    margin = asset_turnover_ratio_ibm(net_sales_str, avg_total_assets)

    print(f"The Asset Turnover ratio of IBM for {2025 - year} is: {margin:.2f}%")

# Capgemini...........................................................................................

import pandas as pd

def asset_turnover_ratio_cap(net_sales, avg_total_assets):
    # Convert net sales and average total assets to numeric types (float)
    net_sales = float(net_sales)
    avg_total_assets = float(avg_total_assets)
    
    # Calculate the asset turnover ratio
    asset_turnover_ratio= (net_sales / avg_total_assets) 
    return asset_turnover_ratio

# Read data from Excel file
file_path = r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Capgemini\Cap_Inc.xlsx'  
df = pd.read_excel(file_path, sheet_name='Income statement')

# Iterate over 4 years
for year in range(2, 6):
    # Extract net sales and average total assets as strings for the current year
    net_sales_str = df.iloc[0, year]  
    avg_total_assets = df.iloc[104, year]

    # Calculate the asset turnover ratio for the current year
    margin = asset_turnover_ratio_cap(net_sales_str, avg_total_assets)

    print(f"The Asset Turnover ratio of Capgemini for {2025 - year} is: {margin:.2f}%")

#Plotting company's graphs...................................................................................
#..................................................................................................

# Financial metrices for Tcs...............................

import pandas as pd
import matplotlib.pyplot as plt

# Read values from Excel file
file_path = r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Tata\Tata_Inc.xlsx'
df = pd.read_excel(file_path, header=None)

# Extract elements data 
elements_data = df.iloc[102:117, 1:5]

# Define labels for the elements
element_labels = ['Net profit margin', 'Gross profit margin', 'ROA', 'ROE', 'Current ratio', 'Quick ratio','Debt to Equity Ratio','Debt Ratio','Revenue Growth rate','Earnings ratio','P/E ratio','P/B ratio','Divident yield','Inventory turnover ratio','Asset turnover ratio'
]

# Plot the graph
plt.figure(figsize=(10, 6))
for i, row in enumerate(elements_data.index):
    plt.plot(elements_data.loc[row], marker='o', label=element_labels[i])

plt.title('Time Series Graph')
plt.xlabel('Year')
plt.ylabel('Financial metrices of TCS')
plt.xticks(range(5), [2019, 2020, 2021, 2022, 2023])  # Set x-axis labels manually
plt.xlim(left=1)
plt.legend()
plt.grid(True)

# Financial metrices for IBM....................................................................

import pandas as pd
import matplotlib.pyplot as plt

# Read values from Excel file
file_path = r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Ibm\ibm_inc.xlsx'
df = pd.read_excel(file_path, header=None)

# Extract elements data 
elements_data = df.iloc[110:124, 1:5]

# Define labels for the elements
element_labels = ['Net profit margin', 'Gross profit margin', 'ROA', 'ROE', 'Current ratio', 'Quick ratio','Debt to Equity Ratio','Debt Ratio','Revenue Growth rate','Earnings ratio','P/E ratio','P/B ratio','Divident yield','Inventory turnover ratio','Asset turnover ratio'
]

# Plot the graph
plt.figure(figsize=(10, 6))
for i, row in enumerate(elements_data.index):
    plt.plot(elements_data.loc[row], marker='o', label=element_labels[i])

plt.title('Time Series Graph')
plt.xlabel('Year')
plt.ylabel('Financial metrices of IBM')
plt.xticks(range(5), [2019, 2020, 2021, 2022, 2023])  # Set x-axis labels manually
plt.xlim(left=1)
plt.legend()
plt.grid(True)

# Financial metrices for Capgemini....................................................................

import pandas as pd
import matplotlib.pyplot as plt

# Read values from Excel file
file_path = r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Capgemini\Cap_Inc.xlsx'
df = pd.read_excel(file_path, header=None)

# Extract elements data 
elements_data = df.iloc[112:125, 1:5]

# Define labels for the elements
element_labels = ['Net profit margin', 'Gross profit margin', 'ROA', 'ROE', 'Current ratio', 'Quick ratio','Debt to Equity Ratio','Debt Ratio','Revenue Growth rate','Earnings ratio','P/E ratio','P/B ratio','Divident yield','Inventory turnover ratio','Asset turnover ratio'
]

# Plot the graph
plt.figure(figsize=(10, 6))
for i, row in enumerate(elements_data.index):
    plt.plot(elements_data.loc[row], marker='o', label=element_labels[i])

plt.title('Time Series Graph')
plt.xlabel('Year')
plt.ylabel('Financial metrices of Capgemini')
plt.xticks(range(5), [2019, 2020, 2021, 2022, 2023])  # Set x-axis labels manually
plt.xlim(left=1)
plt.legend()
plt.grid(True)

#Plotting comparative graphs..........................................................................................
#.........................................................................................................................

#Graphs for comparing revenue growth of all 3 companies...............................................................

import pandas as pd
import matplotlib.pyplot as plt

file_paths = [
    r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Tata\Tata_Inc.xlsx',
    r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Ibm\ibm_inc.xlsx',
    r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Capgemini\Cap_Inc.xlsx'
]

# Initialize empty lists to store elements data and labels
elements_data_list = []
element_labels = []

# Read data from each Excel file and store elements data and labels
for i, file_path in enumerate(file_paths):
    df = pd.read_excel(file_path, header=None)
    if i == 0:
        row_index = 110  # Read 1st row from the first Excel file
    elif i == 1:
        row_index = 117  # Read 2nd row from the second Excel file
    else:
        row_index = 119  # Read 7th row from the third Excel file
    elements_data_list.append(df.iloc[row_index, 1:5])  # Extract data from the specified row
    element_labels.append(f'File {i+1}')

element_labels = ['TCS', 'IBM', 'Capgemini']
# Plot the graph
plt.figure(figsize=(10, 6))
for i, elements_data in enumerate(elements_data_list):
    plt.plot(elements_data, marker='o', label=element_labels[i])
    for x, y in zip(range(1, len(elements_data) + 1), elements_data):
       plt.text(x, y, f'{y:.2f}', ha='right', va='bottom', fontsize=10)

        
plt.title('Time Series Graph')
plt.xlabel('Year')
plt.ylabel('Revenue Growth')
plt.xticks(range(5), [2019, 2020, 2021, 2022, 2023])  # Set x-axis labels manually
plt.xlim(left=1)
plt.legend()
plt.grid(True)
plt.tight_layout()
plt.show()

# Graph for comparing gross profit of the 3 companies.................................................

# Read values from Excel files
import pandas as pd
import matplotlib.pyplot as plt

file_paths = [
    r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Tata\Tata_Inc.xlsx',
    r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Ibm\ibm_inc.xlsx',
    r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Capgemini\Cap_Inc.xlsx'
]

# Initialize empty lists to store elements data and labels
elements_data_list = []
element_labels = []

# Read data from each Excel file and store elements data and labels
for i, file_path in enumerate(file_paths):
    df = pd.read_excel(file_path, header=None)
    if i == 0:
        row_index = 103  # Read 1st row from the first Excel file
    elif i == 1:
        row_index = 110  # Read 2nd row from the second Excel file
    else:
        row_index = 112  # Read 7th row from the third Excel file
    elements_data_list.append(df.iloc[row_index, 1:5])  # Extract data from the specified row
    element_labels.append(f'File {i+1}')

element_labels = ['TCS', 'IBM', 'Capgemini']
# Plot the graph
plt.figure(figsize=(10, 6))
for i, elements_data in enumerate(elements_data_list):
    plt.plot(elements_data, marker='o', label=element_labels[i])
    for x, y in zip(range(1, len(elements_data) + 1), elements_data):
       plt.text(x, y, f'{y:.2f}', ha='right', va='bottom', fontsize=10)

        
plt.title('Time Series Graph')
plt.xlabel('Year')
plt.ylabel('Gross Profit')
plt.xticks(range(5), [2019, 2020, 2021, 2022, 2023])  # Set x-axis labels manually
plt.xlim(left=1)
plt.legend()
plt.grid(True)
plt.tight_layout()
plt.show()

# Graph for comparing Divident yield of the 3 companies..............................................

# Read values from Excel files
import pandas as pd
import matplotlib.pyplot as plt

file_paths = [
    r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Tata\Tata_Inc.xlsx',
    r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Ibm\ibm_inc.xlsx',
    r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Capgemini\Cap_Inc.xlsx'
]

# Initialize empty lists to store elements data and labels
elements_data_list = []
element_labels = []

# Read data from each Excel file and store elements data and labels
for i, file_path in enumerate(file_paths):
    df = pd.read_excel(file_path, header=None)
    if i == 0:
        row_index = 114  # Read 1st row from the first Excel file
    elif i == 1:
        row_index = 121  # Read 2nd row from the second Excel file
    else:
        row_index = 123  # Read 7th row from the third Excel file
    elements_data_list.append(df.iloc[row_index, 1:5])  # Extract data from the specified row
    element_labels.append(f'File {i+1}')

element_labels = ['TCS', 'IBM', 'Capgemini']
# Plot the graph
plt.figure(figsize=(10, 6))
for i, elements_data in enumerate(elements_data_list):
    plt.plot(elements_data, marker='o', label=element_labels[i])
    for x, y in zip(range(1, len(elements_data) + 1), elements_data):
       plt.text(x, y, f'{y:.2f}', ha='right', va='bottom', fontsize=10)

        
plt.title('Time Series Graph')
plt.xlabel('Year')
plt.ylabel('Dividend yield')
plt.xticks(range(5), [2019, 2020, 2021, 2022, 2023])  # Set x-axis labels manually
plt.xlim(left=1)
plt.legend()
plt.grid(True)
plt.tight_layout()
plt.show()

# Part 2: Forecasting future values######################################################################
#########################################################################################################
#########################################################################################################

# TCS.......................................................................................................


import pandas as pd
from sklearn.linear_model import LinearRegression
from sklearn.impute import SimpleImputer

def read_data(file_path, sheet_name, start_row, end_row, start_col, end_col):
    """
    Read data from specific rows and columns in an Excel file.
    Args:
        file_path (str): Path to the Excel file.
        sheet_name (str): Name of the sheet containing the data.
        start_row (int): Row number to start reading from.
        end_row (int): Row number to end reading at.
        start_col (int): Column number to start reading from.
        end_col (int): Column number to end reading at.
    Returns:
        pandas.DataFrame: DataFrame containing the read data.
    """
    # Read data from Excel file and convert non-numeric values to NaN
    data = pd.read_excel(file_path, sheet_name=sheet_name, header=None, skiprows=start_row-1, nrows=end_row-start_row+1, usecols=range(start_col-1, end_col))
    return data

def predict_future_values(data, years_to_predict):
    """
    Predict future values using linear regression.
    Args:
        data (pandas.DataFrame): DataFrame containing the data.
        years_to_predict (int): Number of years to predict.
    Returns:
        pandas.DataFrame: DataFrame containing the predicted values.
    """
    # Extract years and elements
    years = data.columns.values
    elements = data.values

    # Initialize DataFrame for predicted values
    predicted_values = pd.DataFrame(index=range(years[-1] + 1, years[-1] + 1 + years_to_predict))

    # Perform linear regression for each element
    for i, element in enumerate(data.index):
        # Extract values for the element and handle NaN values
        values = elements[i].reshape(-1, 1)

        # Check if values are numeric
        if pd.api.types.is_numeric_dtype(values.flatten()):
            # Initialize imputer to handle missing values
            imputer = SimpleImputer(strategy='mean')
            values_imputed = imputer.fit_transform(values)

            # Initialize linear regression model
            model = LinearRegression()

            # Fit the model
            model.fit([[year] for year in years], values_imputed)

            # Predict future values
            future_years = [year for year in range(years[-1] + 1, years[-1] + 1 + years_to_predict)]
            future_values = model.predict([[year] for year in future_years])

            # Add predicted values to DataFrame
            predicted_values[element] = future_values.flatten()
        else:
            # If the values are not numeric, fill with NaNs
            predicted_values[element] = [None] * years_to_predict

    return predicted_values

# Usage
file_path = r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Tata\Tata_Inc.xlsx'  
sheet_name = 'Income statement'  
start_row = 103  
end_row = 117  
start_col = 2  
end_col = 5  
years_to_predict = 2  

# Read data from Excel file
data = read_data(file_path, sheet_name, start_row, end_row, start_col, end_col)

# Predict future values
predicted_values = predict_future_values(data, years_to_predict)

predicted_values_transposed = predicted_values.transpose()


pd.set_option('display.max_columns', None)

# Print predicted values
print("Predicted Values for the Next 2 Years:")
print(predicted_values)

# IBM.......................................................................................................


import pandas as pd


def read_data(file_path, sheet_name, start_row, end_row, start_col, end_col):
    """
    Read data from specific rows and columns in an Excel file.
    Args:
        file_path (str): Path to the Excel file.
        sheet_name (str): Name of the sheet containing the data.
        start_row (int): Row number to start reading from.import pandas as pd
from sklearn.linear_model import LinearRegression
from sklearn.impute import SimpleImputer

def read_data(file_path, sheet_name, start_row, end_row, start_col, end_col):
    """
    """Read data from specific rows and columns in an Excel file.
    Args:
        file_path (str): Path to the Excel file.
        sheet_name (str): Name of the sheet containing the data.
        start_row (int): Row number to start reading from.
        end_row (int): Row number to end reading at.
        start_col (int): Column number to start reading from.
        end_col (int): Column number to end reading at.
    Returns:
        pandas.DataFrame: DataFrame containing the read data.
    """
    # Read data from Excel file and convert non-numeric values to NaN
    data = pd.read_excel(file_path, sheet_name=sheet_name, header=None, skiprows=start_row-1, nrows=end_row-start_row+1, usecols=range(start_col-1, end_col))
    return data

def predict_future_values(data, years_to_predict):
    """
    Predict future values using linear regression.
    Args:
        data (pandas.DataFrame): DataFrame containing the data.
        years_to_predict (int): Number of years to predict.
    Returns:
        pandas.DataFrame: DataFrame containing the predicted values.
    """
    # Extract years and elements
    years = data.columns.values
    elements = data.values

    # Initialize DataFrame for predicted values
    predicted_values = pd.DataFrame(index=range(years[-1] + 1, years[-1] + 1 + years_to_predict))

    # Perform linear regression for each element
    for i, element in enumerate(data.index):
        # Extract values for the element and handle NaN values
        values = elements[i].reshape(-1, 1)

        # Check if values are numeric
        if pd.api.types.is_numeric_dtype(values.flatten()):
            # Initialize imputer to handle missing values
            imputer = SimpleImputer(strategy='mean')
            values_imputed = imputer.fit_transform(values)

            # Initialize linear regression model
            model = LinearRegression()

            # Fit the model
            model.fit([[year] for year in years], values_imputed)

            # Predict future values
            future_years = [year for year in range(years[-1] + 1, years[-1] + 1 + years_to_predict)]
            future_values = model.predict([[year] for year in future_years])

            # Add predicted values to DataFrame
            predicted_values[element] = future_values.flatten()
        else:
            # If the values are not numeric, fill with NaNs
            predicted_values[element] = [None] * years_to_predict

    return predicted_values

# Usage
file_path = r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Tata\Tata_Inc.xlsx'  
sheet_name = 'Income statement'  
start_row = 103  
end_row = 117  
start_col = 2  
end_col = 5  
years_to_predict = 2  

# Read data from Excel file
data = read_data(file_path, sheet_name, start_row, end_row, start_col, end_col)

# Predict future values
predicted_values = predict_future_values(data, years_to_predict)

predicted_values_transposed = predicted_values.transpose()



pd.set_option('display.max_columns', None)

# Print predicted values
print("Predicted Values for the Next 2 Years:")
print(predicted_values)


# Capgemini.....................................................................................................


import pandas as pd

def read_data(file_path, sheet_name, start_row, end_row, start_col, end_col):
    """
    Read data from specific rows and columns in an Excel file.
    Args:
        file_path (str): Path to the Excel file.
        sheet_name (str): Name of the sheet containing the data.
        start_row (int): Row number to start reading from.
        end_row (int): Row number to end reading at.
        start_col (int): Column number to start reading from.
        end_col (int): Column number to end reading at.
    Returns:
        pandas.DataFrame: DataFrame containing the read data.
    """
    # Read data from Excel file and convert non-numeric values to NaN
    data = pd.read_excel(file_path, sheet_name=sheet_name, header=None, skiprows=start_row-1, nrows=end_row-start_row+1, usecols=range(start_col-1, end_col))
    return data

def predict_future_values(data, years_to_predict):
    """
    Predict future values using linear regression.
    Args:
        data (pandas.DataFrame): DataFrame containing the data.
        years_to_predict (int): Number of years to predict.
    Returns:
        pandas.DataFrame: DataFrame containing the predicted values.
    """
    # Extract years and elements
    years = data.columns.values
    elements = data.values

    # Initialize DataFrame for predicted values
    predicted_values = pd.DataFrame(index=range(years[-1] + 1, years[-1] + 1 + years_to_predict))

    # Perform linear regression for each element
    for i, element in enumerate(data.index):
        # Extract values for the element and handle NaN values
        values = elements[i].reshape(-1, 1)

        # Check if values are numeric
        if pd.api.types.is_numeric_dtype(values.flatten()):
            # Initialize imputer to handle missing values
            imputer = SimpleImputer(strategy='mean')
            values_imputed = imputer.fit_transform(values)

            # Initialize linear regression model
            model = LinearRegression()

            # Fit the model
            model.fit([[year] for year in years], values_imputed)

            # Predict future values
            future_years = [year for year in range(years[-1] + 1, years[-1] + 1 + years_to_predict)]
            future_values = model.predict([[year] for year in future_years])

            # Add predicted values to DataFrame
            predicted_values[element] = future_values.flatten()
        else:
            # If the values are not numeric, fill with NaNs
            predicted_values[element] = [None] * years_to_predict

    return predicted_values

# Usage
file_path = r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Capgemini\Cap_Inc.xlsx'  
sheet_name = 'Income statement'  
start_row = 112  
end_row = 125  
start_col = 2  
end_col = 5  
years_to_predict = 2  # Number of years to predict

# Read data from Excel file
data = read_data(file_path, sheet_name, start_row, end_row, start_col, end_col)

# Predict future values
predicted_values = predict_future_values(data, years_to_predict)

predicted_values_transposed = predicted_values.transpose()

pd.set_option('display.max_columns', None)

# Print predicted values
print("Predicted Values for the Next 2 Years:")
print(predicted_values)

# Graphs for forecasted values.................................................
#.................................................................................

#code to add the forecasted values of TCS.............................................

import pandas as pd
import matplotlib.pyplot as plt

# Read values from Excel file
file_path = r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Tata\Tata_Inc.xlsx'
df = pd.read_excel(file_path, header=None)

# Extract elements data 
elements_data = df.iloc[102:117, 1:7]

# Define labels for the elements
element_labels = ['Net profit margin', 'Gross profit margin', 'ROA', 'ROE', 'Current ratio', 'Quick ratio','Debt to Equity Ratio','Debt Ratio','Revenue Growth rate','Earnings ratio','P/E ratio','P/B ratio','Divident yield','Inventory turnover ratio','Asset turnover ratio'
]

# Check if the value in row 109 is 2024 or 2025
for i in [0,6]:
    if i in [0,3]:
       line_style = ':'
    else:
       line_style = '-'

# Plot the graph
plt.figure(figsize=(10, 6))
for i, row in enumerate(elements_data.index):
    plt.plot(elements_data.loc[row], marker='o', linestyle=line_style, label=element_labels[i])

plt.title('Time Series Graph')
plt.xlabel('Year')
plt.ylabel('Financial metrices of TCS')
plt.xticks(range(7), [2019, 2020, 2021, 2022, 2023, 2024, 2025])  # Set x-axis labels manually

plt.xlim(left=1)
plt.legend()
plt.grid(True)
plt.show()

#code to plot forecasted values for ibm......................................................

import pandas as pd
import matplotlib.pyplot as plt

# Read values from Excel file
file_path = r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Ibm\ibm_inc.xlsx'
df = pd.read_excel(file_path, header=None)

# Extract elements data 
elements_data = df.iloc[109:124, 1:7]

# Define labels for the elements
element_labels = ['Net profit margin', 'Gross profit margin', 'ROA', 'ROE', 'Current ratio', 'Quick ratio','Debt to Equity Ratio','Debt Ratio','Revenue Growth rate','Earnings ratio','P/E ratio','P/B ratio','Divident yield','Inventory turnover ratio','Asset turnover ratio'
]

# Check if the value in row 109 is 2024 or 2025
for i in [0,6]:
    if df.iloc[108, i] in [2024, 2025]:
       line_style = ':'
    else:
       line_style = '-'

# Plot the graph
plt.figure(figsize=(10, 6))
for i, row in enumerate(elements_data.index):
    plt.plot(elements_data.loc[row], marker='o', linestyle=line_style, label=element_labels[i])

plt.title('Time Series Graph')
plt.xlabel('Year')
plt.ylabel('Financial metrices of IBM')
plt.xticks(range(7), [2019, 2020, 2021, 2022, 2023, 2024, 2025])  # Set x-axis labels manually

plt.xlim(left=1)
plt.legend()
plt.grid(True)
plt.show()

#code to plot the forecasted values of Capgemini........................................

import pandas as pd
import matplotlib.pyplot as plt

# Read values from Excel file
file_path = r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Capgemini\Cap_Inc.xlsx'
df = pd.read_excel(file_path, header=None)

# Extract elements data 
elements_data = df.iloc[111:125, 1:7]

# Define labels for the elements
element_labels = ['Net profit margin', 'Gross profit margin', 'ROA', 'ROE', 'Current ratio', 'Quick ratio','Debt to Equity Ratio','Debt Ratio','Revenue Growth rate','Earnings ratio','P/E ratio','P/B ratio','Divident yield','Inventory turnover ratio','Asset turnover ratio'
]

# Check if the value in row 109 is 2024 or 2025
for i in [0,6]:
    if i in [0,3]:
       line_style = ':'
    else:
       line_style = '-'

# Plot the graph
plt.figure(figsize=(10, 6))
for i, row in enumerate(elements_data.index):
    plt.plot(elements_data.loc[row], marker='o', linestyle=line_style, label=element_labels[i])

plt.title('Time Series Graph')
plt.xlabel('Year')
plt.ylabel('Financial metrices of Capgemini')
plt.xticks(range(7), [2019, 2020, 2021, 2022, 2023, 2024, 2025])  # Set x-axis labels manually

plt.xlim(left=1)
plt.legend()
plt.grid(True)
plt.show()

# Part 3: Extracting ESG Factors##########################################################################################
##############################################################################################################

#Tata.........................................................................................................

import pandas as pd
# Read data from Excel file
file_path = r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Tata\Tata_Inc.xlsx'  
df = pd.read_excel(file_path, sheet_name='Income statement')

# Extract esg values from the excel
Env = df.iloc[119, 1]  
soc = df.iloc[120, 1] 
gov = df.iloc[121, 1]
    

print("The Environmental Risk score is ", Env, " Social riskscore is ",soc, " and the governance risk score is ", gov)

#IBM......................................................................................................

import pandas as pd
# Read data from Excel file
file_path = r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Ibm\ibm_inc.xlsx'  
df = pd.read_excel(file_path, sheet_name='Income statement')

# Extract esg values from the excel
Env = df.iloc[126, 1]  
soc = df.iloc[127, 1] 
gov = df.iloc[128, 1]
    
print("The Environmental Risk score is ", Env, " Social riskscore is ",soc, " and the governance risk score is ", gov)

# Capgemini....................................................................................................

import pandas as pd
# Read data from Excel file
file_path = r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Capgemini\Cap_Inc.xlsx'  
df = pd.read_excel(file_path, sheet_name='Income statement')

# Extract esg values from the excel
Env = df.iloc[127, 1]  
soc = df.iloc[128, 1] 
gov = df.iloc[129, 1]
    
print("The Environmental Risk score is ", Env, " Social riskscore is ",soc, " and the governance risk score is ", gov)

#Plotting ESG factors.............................................................................................
#..................................................................................................................

# Plotting graph for Environmental risk score..........................................................
# Read values from Excel files
import pandas as pd
import matplotlib.pyplot as plt

file_paths = [
    r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Tata\Tata_Inc.xlsx',
    r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Ibm\ibm_inc.xlsx',
    r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Capgemini\Cap_Inc.xlsx'
]

# Cell coordinates for each file
cell_coordinates = [
    (120, 1),   # Cell coordinates for the first file (row index, column index)
    (128, 1),   # Cell coordinates for the second file (row index, column index)
    (128, 1)   # Cell coordinates for the third file (row index, column index)
]
 
# Initialize lists to store the values from each Excel file
values = []

# Read data from the specified cells in each Excel file
for file_path, (row_index, col_index) in zip(file_paths, cell_coordinates):
    df = pd.read_excel(file_path, header=None)
    value = df.iloc[row_index, col_index]
    values.append(value)

# Plot the graph
plt.figure(figsize=(10, 6))
plt.plot(values, marker='o', linestyle='-')
plt.title('Comparison of Environmental Risk score')
plt.xlabel('Companies')
plt.ylabel('Environmental risk')
plt.xticks(range(3), ['TCS', 'IBM', 'Capgemini'])  # Set x-axis labels manually
plt.grid(True)
plt.tight_layout()
plt.show()   


# Plotting graph for social risk score..........................................................
# Read values from Excel files
import pandas as pd
import matplotlib.pyplot as plt

file_paths = [
    r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Tata\Tata_Inc.xlsx',
    r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Ibm\ibm_inc.xlsx',
    r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Capgemini\Cap_Inc.xlsx'
]

# Cell coordinates for each file
cell_coordinates = [
    (121, 1),   # Cell coordinates for the first file (row index, column index)
    (129, 1),   # Cell coordinates for the second file (row index, column index)
    (129, 1)   # Cell coordinates for the third file (row index, column index)
]
 
# Initialize lists to store the values from each Excel file
values = []

# Read data from the specified cells in each Excel file
for file_path, (row_index, col_index) in zip(file_paths, cell_coordinates):
    df = pd.read_excel(file_path, header=None)
    value = df.iloc[row_index, col_index]
    values.append(value)

# Plot the graph
plt.figure(figsize=(10, 6))
plt.plot(values, marker='o', linestyle='-')
plt.title('Comparison of Social Risk score of 3 companies')
plt.xlabel('Companies')
plt.ylabel('Social Risk')
plt.xticks(range(3), ['TCS', 'IBM', 'Capgemini'])  # Set x-axis labels manually
plt.grid(True)
plt.tight_layout()
plt.show()   

# Plotting graph for governance risk score........................................................
# Read values from Excel files
import pandas as pd
import matplotlib.pyplot as plt

file_paths = [
    r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Tata\Tata_Inc.xlsx',
    r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Ibm\ibm_inc.xlsx',
    r'C:\Users\Mukul Kapri\Desktop\KDP Projects\Python assignment\Excels\Capgemini\Cap_Inc.xlsx'
]

# Cell coordinates for each file
cell_coordinates = [
    (122, 1),   # Cell coordinates for the first file (row index, column index)
    (130, 1),   # Cell coordinates for the second file (row index, column index)
    (130, 1)   # Cell coordinates for the third file (row index, column index)
]
 
# Initialize lists to store the values from each Excel file
values = []

# Read data from the specified cells in each Excel file
for file_path, (row_index, col_index) in zip(file_paths, cell_coordinates):
    df = pd.read_excel(file_path, header=None)
    value = df.iloc[row_index, col_index]
    values.append(value)

# Plot the graph
plt.figure(figsize=(10, 6))
plt.plot(values, marker='o', linestyle='-')
plt.title('Comparison of Governance risk score of 3 companies')
plt.xlabel('Companies')
plt.ylabel('Governance Risk')
plt.xticks(range(3), ['TCS', 'IBM', 'Capgemini'])  # Set x-axis labels manually
plt.grid(True)
plt.tight_layout()
plt.show()   

############################## Finish ######################################