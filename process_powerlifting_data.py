import pandas as pd
import numpy as np
import math

def calculate_ipf_gl_points(bodyweight, total, sex, event):
    """Calculate IPF GL Points using official IPF formula: 100/(A-B*exp(-C*bodyweight))"""
    if pd.isna(bodyweight) or pd.isna(total) or total == 0:
        return 0
    
    # Official IPF GL Coefficients from IPF_GL_Coefficients-2020.pdf
    if event == 'SBD':  # Classic Powerlifting
        if sex == 'M':  # Men's Classic Powerlifting
            A = 1199.72839
            B = 1025.18162
            C = 0.00921
        else:  # Women's Classic Powerlifting
            A = 610.32796
            B = 1045.59282
            C = 0.03048
    else:  # Bench Only (Classic Bench Press)
        if sex == 'M':  # Men's Classic Bench Press
            A = 320.98041
            B = 281.40258
            C = 0.01008
        else:  # Women's Classic Bench Press
            A = 142.40398
            B = 442.52671
            C = 0.04724
    
    try:
        # IPF GL Points formula: 100/(A-B*exp(-C*bodyweight))
        ipf_gl_points = 100 / (A - B * math.exp(-C * bodyweight)) * total
        
        return round(ipf_gl_points, 2)
    except:
        return 0

def process_powerlifting_data():
    # Read the detailed competition results (OPL format)
    # Skip first 4 rows (lines 1-4), use line 5 as header
    df_detailed = pd.read_csv('bjelovar/3-bjelovar-record-breakers.opl (1).csv', skiprows=4)
    
    # Check columns in detailed file
    print(f"Columns in detailed file: {df_detailed.columns.tolist()}")
    
    # Read the nominations file to get club information
    # Skip the first 2 rows and use row 3 as header
    df_nominations = pd.read_csv('bjelovar/Bjelovar-record-breakers-finalne-nominacije-2-1-3-1-1-1.csv', 
                                sep=';', encoding='utf-8', skiprows=2)
    
    # Create a mapping of names to clubs from the nominations file
    # Clean up the nominations data first
    # Check if the expected columns exist
    expected_cols = ['IME', 'PREZIME', 'KLUB']
    available_cols = df_nominations.columns.tolist()
    print(f"Available columns in nominations file: {available_cols}")
    
    # Find the correct column indices/names
    name_col = None
    surname_col = None
    club_col = None
    
    for col in available_cols:
        if col == 'IME':
            name_col = col
        elif col == 'PREZIME':  
            surname_col = col
        elif col == 'KLUB':
            club_col = col
    
    print(f"Using columns - Name: {name_col}, Surname: {surname_col}, Club: {club_col}")
    
    # Filter and clean nominations data
    df_nominations_clean = df_nominations.copy()
    if name_col and surname_col and club_col:
        df_nominations_clean = df_nominations_clean.dropna(subset=[name_col, surname_col, club_col])
        df_nominations_clean = df_nominations_clean[df_nominations_clean[name_col].notna()]
        df_nominations_clean = df_nominations_clean[df_nominations_clean[name_col] != '']
    
    # Create full names and club mapping
    club_mapping = {}
    if name_col and surname_col and club_col:
        for _, row in df_nominations_clean.iterrows():
            if pd.notna(row[name_col]) and pd.notna(row[surname_col]) and pd.notna(row[club_col]):
                full_name = f"{row[name_col]} {row[surname_col]}"
                club_mapping[full_name] = row[club_col]
    else:
        print("Could not find required columns for club mapping")
    
    # Filter out divisions starting with "Best"
    df_filtered = df_detailed[~df_detailed['Division'].str.startswith('Best', na=False)]
    
    # Remove rows where Place is NaN or empty (header rows, etc.)
    df_filtered = df_filtered.dropna(subset=['Place'])
    df_filtered = df_filtered[df_filtered['Place'] != '']
    
    # Create the output dataframe with requested columns
    output_data = []
    
    for _, row in df_filtered.iterrows():
        # Get club from mapping, or use empty string if not found
        club = club_mapping.get(row['Name'], '')
        
        # Calculate IPF GL Points
        ipf_points = calculate_ipf_gl_points(
            bodyweight=pd.to_numeric(row['BodyweightKg'], errors='coerce'),
            total=pd.to_numeric(row['TotalKg'], errors='coerce'),
            sex=row['Sex'],
            event=row['Event']
        )
        
        output_data.append({
            'Place': row['Place'],
            'Name': row['Name'],
            'Club': club,
            'Sex': row['Sex'],
            'BirthYear': row['BirthYear'],
            'Division': row['Division'],
            'BodyweightKg': row['BodyweightKg'],
            'WeightClassKg': row['WeightClassKg'],
            'Squat1Kg': row['Squat1Kg'],
            'Squat2Kg': row['Squat2Kg'],
            'Squat3Kg': row['Squat3Kg'],
            'Best3SquatKg': row['Best3SquatKg'],
            'Bench1Kg': row['Bench1Kg'],
            'Bench2Kg': row['Bench2Kg'],
            'Bench3Kg': row['Bench3Kg'],
            'Best3BenchKg': row['Best3BenchKg'],
            'Deadlift1Kg': row['Deadlift1Kg'],
            'Deadlift2Kg': row['Deadlift2Kg'],
            'Deadlift3Kg': row['Deadlift3Kg'],
            'Best3DeadliftKg': row['Best3DeadliftKg'],
            'TotalKg': row['TotalKg'],
            'Points': ipf_points,
            'Event': row['Event']
        })
    
    # Create output dataframe
    output_df = pd.DataFrame(output_data)
    
    # Save to CSV
    output_df.to_csv('powerlifting_results_processed.csv', index=False, encoding='utf-8')
    
    print(f"Processed {len(output_df)} records")
    print(f"Found club information for {sum(1 for club in output_df['Club'] if club != '')} athletes")
    print("Output saved to 'powerlifting_results_processed.csv'")
    
    # Display some sample data
    print("\nFirst 5 records:")
    print(output_df.head())
    
    # Show unique divisions
    print(f"\nUnique divisions (after filtering 'Best' divisions):")
    print(output_df['Division'].unique())

if __name__ == "__main__":
    process_powerlifting_data() 