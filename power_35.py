import pandas as pd
import matplotlib.pyplot as plt

# Load your Excel file
file_path = 'power_data.xlsx'  # Update this to your actual file name

# Read the Excel file
df = pd.read_excel(file_path)

# Convert to datetime
df['Date Time'] = pd.to_datetime(df['Date Time'])

# Sort by time
df = df.sort_values('Date Time')

# Filter where power is less than 35
low_power = df[df['Power'] < 35].copy()

# Add formatted date/time column
low_power['Drop_Detected_At'] = low_power['Date Time'].dt.strftime('%d-%b-%Y %I:%M:%S %p')

# Display (optional)
print("⚠️ Power readings below 35:\n")
print(low_power[['Drop_Detected_At', 'Power']])

# Save to Excel
low_power[['Drop_Detected_At', 'Date Time', 'Power']].to_excel("power_drops_below_35.xlsx", index=False)

# Optional plot
plt.figure(figsize=(12, 5))
plt.plot(df['Date Time'], df['Power'], label='Power')
plt.scatter(low_power['Date Time'], low_power['Power'], color='red', label='Below 35')
plt.title('Power < 35 Drop Detection')
plt.xlabel('Time')
plt.ylabel('Power')
plt.legend()
plt.grid(True)
plt.xticks(rotation=45)
plt.tight_layout()
plt.show()
