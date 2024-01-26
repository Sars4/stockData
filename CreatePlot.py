import pandas as pd
import matplotlib.pyplot as plt

# Load spreadsheet
xl = pd.ExcelFile('./Data/NVDAData.xlsx')

# Load a sheet into a DataFrame
df = xl.parse('Sheet')

# Assume that 'Column1' and 'Column2' are the columns you want to plot
plt.plot(df['Date'], df['Closing'])

# Display the plot
plt.show()
