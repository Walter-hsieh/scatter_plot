import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import matplotlib.font_manager as fm
import os

# Load the Excel file
file_path = 'data.xlsx'  # Replace with the path to your Excel file
excel_data = pd.ExcelFile(file_path)

# Load the data from the relevant sheet
data = pd.read_excel(file_path, sheet_name='data')

# Dropping rows with missing values in 'dp_inet' and 'D-MAX' columns for the first plot
data_clean_dmax = data.dropna(subset=['dp_inet', 'D-MAX'])

# Extracting the relevant columns for plotting the first plot
x_dmax = data_clean_dmax['dp_inet']
y_dmax = data_clean_dmax['D-MAX']

# Dropping rows with missing values in 'dp_inet' and 'ITRI-PY-17' columns for the second plot
data_clean_itri = data.dropna(subset=['dp_inet', 'ITRI-PY-17'])

# Extracting the relevant columns for plotting the second plot
x_itri = data_clean_itri['dp_inet']
y_itri = data_clean_itri['ITRI-PY-17']

# Path to a font that supports Chinese characters
font_path = 'C:\\Users\\A70459\\AppData\\Local\\Microsoft\\Windows\\Fonts\\Noto Sans CJK Regular.otf'  # Replace with the path to a Chinese-supporting font

# Check if the font file exists
if not os.path.isfile(font_path):
    raise FileNotFoundError(f"The font file at {font_path} does not exist.")

# Load the font
prop = fm.FontProperties(fname=font_path)

# markers
markers = ["o", "v", "s", "p", "x"]

# Creating a combined scatter plot with trendlines for both D-MAX and ITRI-PY-17
plt.figure(figsize=(10, 6))

# Scatter plot and trendline for dp_inet vs. D-MAX
plt.scatter(x_dmax, y_dmax, color='blue', label='D-MAX Data Points', marker=markers[0])
z_dmax = np.polyfit(x_dmax, y_dmax, 1)
p_dmax = np.poly1d(z_dmax)
plt.plot(x_dmax, p_dmax(x_dmax), color='red', linestyle='--', label='D-MAX Trendline')

# Scatter plot and trendline for dp_inet vs. ITRI-PY-17
plt.scatter(x_itri, y_itri, color='green', label='ITRI-PY-17 Data Points', marker=markers[1])
z_itri = np.polyfit(x_itri, y_itri, 1)
p_itri = np.poly1d(z_itri)
plt.plot(x_itri, p_itri(x_itri), color='orange', linestyle='--', label='ITRI-PY-17 Trendline')

plt.xlabel('入口濕度°C',fontproperties=prop, fontsize=14)
plt.ylabel('出口濕度°C', fontproperties=prop, fontsize=14)
plt.title('轉輪除濕性能曲線', fontproperties=prop, fontsize=14)
plt.legend(prop=prop)
plt.grid(True)

# Save the figure with the specified font
plt.savefig('轉輪除濕性能曲線.png', dpi=300, bbox_inches='tight')

# Show the figure
plt.show()
