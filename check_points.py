import pandas as pd

df = pd.read_csv('powerlifting_results_processed.csv')

print("IPF GL Points Statistics:")
print(f"Average Points: {df['Points'].mean():.2f}")
print(f"Highest Points: {df['Points'].max():.2f}")
print(f"Lowest Points: {df['Points'].min():.2f}")
print()

print("Top 10 performances by IPF GL Points:")
top_10 = df.nlargest(10, 'Points')[['Name', 'Club', 'Sex', 'Division', 'BodyweightKg', 'TotalKg', 'Points', 'Event']]
print(top_10.to_string(index=False))
print()

print("Points breakdown by Event:")
event_stats = df.groupby('Event')['Points'].agg(['count', 'mean', 'max', 'min']).round(2)
print(event_stats)
print()

print("Points breakdown by Sex:")
sex_stats = df.groupby('Sex')['Points'].agg(['count', 'mean', 'max', 'min']).round(2)
print(sex_stats) 