import pandas as pd
import matplotlib.pyplot as plt

# === File Paths ===
input_file = 'power_data.xlsx'
output_file = 'daily_open_close.xlsx'
threshold = 21

# === Load Data (Skip first row if invalid) ===
df = pd.read_excel(input_file)

# Remove any row where "Date Time" is not a datetime
df = df[pd.to_datetime(df['Date Time'], errors='coerce').notna()]
df['Date Time'] = pd.to_datetime(df['Date Time'])

# === Continue as before ===
df = df.sort_values('Date Time')
df['Date'] = df['Date Time'].dt.date
df['Time'] = df['Date Time'].dt.time

# === Get daily opening/closing times ===
records = []

for date, group in df.groupby('Date'):
    above_threshold = group[group['Power'] >= threshold]
    if not above_threshold.empty:
        opening_time = above_threshold.iloc[0]['Date Time'].strftime('%I:%M %p')
        closing_time = above_threshold.iloc[-1]['Date Time'].strftime('%I:%M %p')
        records.append({
            'Closing Time': closing_time,
            'Date': date,
            'Opening Time': opening_time
        })

daily_df = pd.DataFrame(records)
daily_df.to_excel(output_file, index=False)
print(f"Excel saved to: {output_file}")

# === Plot power drops only (<21) ===
drop_df = df[df['Power'] < threshold]

plt.figure(figsize=(14, 6))
plt.plot(df['Date Time'], df['Power'], label='Power', color='blue')
plt.scatter(drop_df['Date Time'], drop_df['Power'], color='red', label='Drop < 21', s=40)
plt.axhline(threshold, color='gray', linestyle='--', label='Threshold = 21')

plt.title('Power Drops (Below 21)')
plt.xlabel('Date Time')
plt.ylabel('Power Value')
plt.legend()
plt.grid(True)
plt.tight_layout()
plt.show()
