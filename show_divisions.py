import pandas as pd

df = pd.read_csv('powerlifting_results_processed.csv')
print('Unique divisions:')
for division in sorted(df['Division'].unique()):
    print(f"  - {division}")

print(f'\nTotal records: {len(df)}')
print(f'Total unique divisions: {len(df["Division"].unique())}')

# Show breakdown by sex
print(f'\nBreakdown by sex:')
print(df['Sex'].value_counts())

# Show breakdown by event
print(f'\nBreakdown by event:')
print(df['Event'].value_counts()) 