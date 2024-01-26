import pandas as pd
import matplotlib.pyplot as plt
import numpy as np


import createStockData

# Load spreadsheet
xl = pd.ExcelFile('./Data/NVDAData.xlsx')
df = xl.parse('Sheet')

colors = np.where(df['Closing'].diff() > 0, 'green', 'red')

# Find the columns
plt.plot(df['Date'], df['Closing'])
for i in range(1, len(df)):
    plt.plot(df['Date'].iloc[i-1:i+1], df['Closing'].iloc[i-1:i+1], color=colors[i-1])

# Display the plot
plt.show()
