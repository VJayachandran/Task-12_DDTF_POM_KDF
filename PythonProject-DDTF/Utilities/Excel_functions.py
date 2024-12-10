"""
excel_functions.py

Program : Python program which will Read & Write your Excel file
"""

import pandas as pd
from datetime import datetime

# Create DataFrame
data = {
    'Test ID': [1, 2, 3, 4, 5],
    'Username': ['Admin', 'guvi', 'ramesh', 'dinesh', 'user5'],
    'Password': ['admin123', 'guvi@123', 'password@123', 'password@123', 'pass5'],
    'Date': [datetime.now().strftime('%Y-%m-%d')] * 5,
    'Time of Test': [datetime.now().strftime('%H:%M:%S')] * 5,
    'Name of Tester': ['Tester'] * 5,
    'Test Result': [''] * 5
}

df = pd.DataFrame(data)
df.to_excel('test_data.xlsx', index=False)
