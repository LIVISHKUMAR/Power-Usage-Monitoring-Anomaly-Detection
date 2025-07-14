import pandas as pd
import matplotlib.pyplot as plt

# Load your Excel data
df = pd.read_excel("power_data.xlsx")

# Convert Date Time to datetime and sort
df['Date Time'] = pd.to_datetime(df['Date Time'])
df = df.sort_values('Date Time')

# Mark ON/OFF based on power threshold
df['Status'] = df['Power'].apply(lambda x: 'ON' if x >= 21 else 'OFF')
df['Previous_Status'] = df['Status'].shift(1)
df['Transition'] = df['Previous_Status'] + ' → ' + df['Status']

# Filter transitions (ON → OFF for closing, OFF → ON for opening)
transitions = df[df['Transition'].isin(['OFF → ON', 'ON → OFF'])].copy()
transitions['Date'] = transitions['Date Time'].dt.date
transitions['Hour'] = transitions['Date Time'].dt.hour

# Closing Time: ON → OFF between 12:00 AM to 07:00 AM
closing_df = transitions[
    (transitions['Transition'] == 'ON → OFF') & (transitions['Hour'] < 7)
].copy()
closing_df['Business Date'] = closing_df['Date Time'].dt.date

# Opening Time: OFF → ON after 07:00 AM
opening_df = transitions[
    (transitions['Transition'] == 'OFF → ON') & (transitions['Hour'] >= 7)
].copy()
opening_df['Business Date'] = opening_df['Date Time'].dt.date

# Group by business date to get first opening and closing times
closing_grouped = closing_df.groupby('Business Date')['Date Time'].first().reset_index(name='Closing Time')
opening_grouped = opening_df.groupby('Business Date')['Date Time'].first().reset_index(name='Opening Time')

# Merge both times into summary
summary = pd.merge(closing_grouped, opening_grouped, on='Business Date', how='inner')
summary['Date'] = summary['Business Date']
summary['Closing Time'] = pd.to_datetime(summary['Closing Time']).dt.strftime('%I:%M %p')
summary['Opening Time'] = pd.to_datetime(summary['Opening Time']).dt.strftime('%I:%M %p')

# Final Output Format
final_output = summary[['Closing Time', 'Date', 'Opening Time']]

# Save to Excel
final_output.to_excel('shop_timing_summary.xlsx', index=False)

# Print for confirmation
print(final_output)

# Optional: Plot
df_drops = df[df['Power'] < 21][['Date Time', 'Power']]
plt.figure(figsize=(14, 6))
plt.plot(df['Date Time'], df['Power'], label='Power', color='blue')
plt.scatter(df_drops['Date Time'], df_drops['Power'], color='red', label='Power < 21', s=15)
plt.title('Power Meter Readings with Drops')
plt.xlabel('Date Time')
plt.ylabel('Power')
plt.legend()
plt.grid(True)
plt.tight_layout()
plt.savefig('power_readings_with_drops.png')
plt.show()
