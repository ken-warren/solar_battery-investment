# Import required libraries
import numpy as np                # Mathematical functions
import pandas as pd               # data manipulation
import matplotlib.pyplot as plt   # data visualization
import seaborn as sns
import openpyxl as pxl            # reading excel file

# Load the data using your own file path
data = pd.read_excel("C:\\Users\\ADMIN\\solar_battery savings\\solar-savings\\Dataset\\2020_solar_Data.xlsx", parse_dates=['date'])

# View the the dataset
print(data.dtypes, data.head())

# Data Cleaning
# 1. Ensure the date is in the correct datetime format
data['date'] = pd.to_datetime(data['date'])

# 2. Check for missing values
missing_vals = data.isnull().sum()
print("Missing values:", missing_vals)

# 3. Check for duplicates
dups = data.duplicated().sum()
print("Duplicates:", dups)

# 4. Drop rows with negative values in 'solar' and 'electricity' columns
data = data[(data['solar'] >= 0) & (data['electricity'] >= 0)]

# Calculate avg solar energy and electricity usage each hour
hourly_avg = data.groupby('hour').agg({'solar': 'mean', 'electricity': 'mean'})
print(hourly_avg)

# Plot the averages
plt.figure(figsize=(14, 7))
sns.lineplot(data=hourly_avg, x='hour', y='solar', label='Average Solar Energy (kWh)') 
sns.lineplot(data=hourly_avg, x='hour', y='electricity', label='Average Electricity Usage (kWh)')
plt.xlabel('Hour')
plt.ylabel('kWh')
plt.title('Hourly Average Solar Electricity Generation & Electricity Usage')
plt.legend()
plt.grid(True)
plt.show()

#  Investigating outliers using IQR
def detect_outliers(data, variable):
    Q1 = data[variable].quantile(0.25)    # lower quartile
    Q3 = data[variable].quantile(0.75)    # upper quartile
    IQR = Q3 - Q1                         # inter-quartile range
    lower_bound = Q1 - 1.5 * IQR
    upper_bound = Q3 + 1.5 * IQR
    outliers = data[(data[variable] < lower_bound) | (data[variable] > upper_bound)]
    return outliers
solar_outliers = detect_outliers(data, 'solar')
electric_outliers = detect_outliers(data, 'electricity')

# Graph representation of outliers
plt.figure(figsize=(14, 7))
sns.boxplot(data=data[['solar', 'electricity']])
plt.title('Boxplot for Solar Energy and Electricity Usage')
plt.grid(True)
plt.show()

# Handle outliers (example: replacing with median)
data.loc[data['solar'] > solar_outliers['solar'].quantile(0.75), 'solar'] = data['solar'].median()
data.loc[data['electricity'] > electric_outliers['electricity'].quantile(0.75), 'electricity'] = data['electricity'].median()

# Verify the corrections
print(data[['solar', 'electricity']].describe())

plt.figure(figsize=(14, 7))
sns.boxplot(data=data[['solar', 'electricity']])
plt.title('Boxplot for Solar Energy and Electricity Usage after Outlier Correction')
plt.grid(True)
plt.show()

# Hourly electricity needed
data['electricity_needed'] = (data['electricity'] - data['solar']).clip(lower=0) #min value subjected to zero
hourly_electricity_needed = data.groupby('hour')['electricity_needed'].sum()
print('Electricity Needed per hour:\n', hourly_electricity_needed)

# Hourly excess solar electricity generated
data['excess_solar'] = (data['solar'] - data['electricity']).clip(lower=0)
hourly_excess_solar = data.groupby('hour')['excess_solar'].sum()
print('Excess Solar Generated per Hour:\n', hourly_excess_solar)

## Cumulative battery charge per hour
# Assumptions: Battery already installed and charge level should:
# 1. Begin at 0 at 2020-01-01 00:00
# 2. Allow increase/decrease depending on hourly results
# 3. subject to max battery charge level
# Max battery charge level = 12.5 kWh

# Initialize the variables
max_charge = 12.5
data['battery_charge'] = 0.0

# Iteration for each hour
for i in range(len('data')):
    if i > 0:   # skip the first hour
        data.at[i, 'battery_charge'] = min(data.at[i-1, 'battery_charge'] + data.at[i-1, 'excess_solar'], max_charge)
        data.at[i, 'battery_charge'] = max(data.at[i, 'battery_charge'] + data.at[i, 'electricity_needed'], 0)

print(data[['hour', 'date', 'battery_charge']].tail())
hourly_battery_charge = data.groupby('hour')['battery_charge'].sum()
print(hourly_battery_charge)

# Calculation of Savings
# Initialize the variables
electricity_price = 0.17  # USD

# Electricity consumption without battery
data['electricity_wout_bat'] = data.apply(lambda row: row['electricity_needed'] if row['battery_charge'] == 0 else 0, axis=1)

# Savings
electricity_saved = (data['electricity_wout_bat'] - data['electricity_w_bat']).clip(lower=0)
saving_cost = electricity_saved * electricity_price
data['savings'] = saving_cost
hourly_savings = data.groupby('hour')['savings'].sum()
total_savings = data['savings'].sum()
print(data[['hour', 'date', 'savings']].head())
print("Hourly Savings for 2020:\n", round(hourly_savings,2))
print("Total Savings for 2020:", round(total_savings,2))

