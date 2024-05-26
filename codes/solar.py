# Import required libraries
import numpy_financial as npf     # Financial functions eg IRR
import pandas as pd               # data manipulation
import matplotlib.pyplot as plt   # data visualization
import seaborn as sns
import openpyxl as pxl            # reading excel file

# Load the data using your own file path
data = pd.read_excel("2020_solar_Data.xlsx", parse_dates=['date'])

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
data['electricity_wout_bat'] = data['electricity_needed']

# Savings
electricity_saved = (data['electricity_wout_bat'] - data['electricity_w_bat']).clip(lower=0)
saving_cost = electricity_saved * electricity_price
data['savings'] = saving_cost
hourly_savings = data.groupby('hour')['savings'].sum()
annual_savings = data['savings'].sum()
print(data[['hour', 'date', 'savings']].head())
print("Hourly Savings for 2020:\n", round(hourly_savings,2))
print("Annual Savings for 2020:", round(annual_savings,2))

# Extract month from the date
data['month'] = data['date'].dt.strftime('%B')

# Order the months for better visualization
month_order = ["January", "February", "March", "April", "May", "June",
               "July", "August", "September", "October", "November", "December"]
data['monthly'] = data['month'].reindex(month_order)

# Monthly tabulation
monthly_solar = data.groupby('month')['solar'].sum()
monthly_electricity = data.groupby('month')['electricity'].sum()
monthly_electricity_wout_bat = data.groupby('month')['electricity_needed'].sum()
monthly_electricity_w_bat = data.groupby('month')['electricity_w_bat'].sum()

monthly_data = pd.DataFrame({
    'Monthly Solar Generation (kWh)': monthly_solar,
    'Monthly Electricity Usage (kWh)': monthly_electricity,
    'Monthly Electricity Purchased (No Battery)': monthly_electricity_wout_bat,
    'Monthly Electricity Purchased (with battery)': monthly_electricity_w_bat
})


#Plot the data
plt.figure(figsize=(14, 7))

# Plot Monthly Solar Generation
plt.subplot(2, 2, 1)
monthly_solar.plot(kind='bar', color='gold')
plt.title('Monthly Solar Generation (2020)')
plt.xlabel('Month')
plt.ylabel('kWh')
plt.grid(True)

# Plot Monthly Electricity Usage
plt.subplot(2, 2, 2)
monthly_electricity.plot(kind='bar', color='lightblue')
plt.title('Monthly Electricity Usage (2020)')
plt.xlabel('Month')
plt.ylabel('kWh')
plt.grid(True)

# Plot Monthly Electricity Purchased (No Battery)
plt.subplot(2, 2, 3)
monthly_electricity_wout_bat.plot(kind='bar', color='lightgreen')
plt.title('Monthly Electricity Purchased (No Battery) (2020)')
plt.xlabel('Month')
plt.ylabel('kWh')
plt.grid(True)

# Plot Monthly Electricity Purchased (with battery)
plt.subplot(2, 2, 4)
monthly_electricity_w_bat.plot(kind='bar', color='blue')
plt.title('Monthly Electricity Purchased (with battery) (2020)')
plt.xlabel('Month')
plt.ylabel('kWh')
plt.grid(True)

plt.tight_layout()
plt.show()

# Calculation of future projections
# Annual Savings
savings_2022 = sum(data['savings'])
print('The projected annual savings for 2022 is:', round(savings_2022, 2))

# Inialize a list for initial investment on battery
cost_savings1 = [-7000]

# initialize the values
r1 = 0.04              # price increase rate (4% p.a)
years = range(0, 20)

# Cost savings per year
for i in years:
    yearly_savings = savings_2022 * (1 + r1) ** i
    
    # append to the initial investment
    cost_savings1.append(yearly_savings)
    
print(cost_savings1)
# NPV calculation for scenario 1

# Initialize the values
dr = 0.06         # discount rate of 6% p.a
npv_vals1 = []

for i,c in enumerate(cost_savings):
    discount_val = c / (1 + dr) **i
    npv_vals1.append(discount_val)
    
npv1 = sum(npv_vals1)
print('Total NPV for scenario 1:', round(npv1, 2))

# Initialize the values 
r2 = 0.0025
cost_savings2 = [-7000]

# Yearly cost savings in scenario 2
for i in years:
    
    # The first years increase is 4% p.a
    if i == 0: 
        yearly_savings = savings_2022 * (1 + r1) **i
    else:
        yearly_savings = savings_2022 * (1 + (r1 + ((i) * r2))) **i
        cost_savings2.append(yearly_savings)
print(cost_savings2)

# NPV for Scenario 2
# initialize the values
npv_vals2 = []

for i,c in enumerate(cost_savings2):
    discount_vals = c / (1 + dr) **i
    npv_vals2.append(discount_vals)
    
npv2 = sum(npv_vals2)
print('Total NPV for scenario 2:', round(npv2, 2))

# Calculating IRR

#Scenario 1: 4% p.a
IRR1 = round(npf.irr(npv_vals1), 2)

# Scenario 2: 4% p.a + 0.25% p.a
IRR2 = round(npf.irr(npv_vals2), 2)

print('IRR for scenario 1:', IRR1)
print('IRR for scenario 2:', IRR2)