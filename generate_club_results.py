import pandas as pd
import numpy as np

def is_integer_placement(place_str):
    """Check if placement is an integer (not DQ, DD, G, etc.)"""
    try:
        int(place_str)
        return True
    except (ValueError, TypeError):
        return False

def generate_club_results():
    # Read the processed data
    df = pd.read_csv('powerlifting_results_processed.csv')
    
    # Apply filters
    print("Applying filters...")
    
    # Rule 1: Only lifters with integer placements
    df_filtered = df[df['Place'].apply(is_integer_placement)]
    print(f"After integer placement filter: {len(df_filtered)} records")
    
    # Rule 2: No guest lifters (exclude divisions containing "Guest")
    df_filtered = df_filtered[~df_filtered['Division'].str.contains('Guest', na=False)]
    print(f"After excluding guest lifters: {len(df_filtered)} records")
    
    # Rule 3: Only consider lifters with Points > 0 (valid calculations)
    df_filtered = df_filtered[df_filtered['Points'] > 0]
    print(f"After excluding zero points: {len(df_filtered)} records")
    
    # Group and process for each category
    categories = [
        ('M', 'SBD', 'Male_Powerlifting'),
        ('F', 'SBD', 'Female_Powerlifting'),
        ('M', 'B', 'Male_Bench_Only'),
        ('F', 'B', 'Female_Bench_Only')
    ]
    
    for sex, event, filename in categories:
        print(f"\nProcessing {filename}...")
        
        # Filter for specific category
        category_data = df_filtered[(df_filtered['Sex'] == sex) & (df_filtered['Event'] == event)]
        print(f"Records in {filename}: {len(category_data)}")
        
        if len(category_data) == 0:
            print(f"No data for {filename}, creating empty file...")
            empty_df = pd.DataFrame(columns=['Club', 'Rank', 'Name', 'Division', 'BodyweightKg', 'TotalKg', 'Points'])
            empty_df.to_csv(f'{filename}.csv', index=False)
            continue
        
        # Group by club and get top 5 lifters per club by GL Points
        club_results = []
        
        for club, club_data in category_data.groupby('Club'):
            # Sort by GL Points (descending) and take top 5
            top_lifters = club_data.nlargest(5, 'Points')
            
            for rank, (_, lifter) in enumerate(top_lifters.iterrows(), 1):
                club_results.append({
                    'Club': club,
                    'Rank': rank,
                    'Name': lifter['Name'],
                    'Division': lifter['Division'],
                    'BodyweightKg': lifter['BodyweightKg'],
                    'TotalKg': lifter['TotalKg'],
                    'Points': lifter['Points']
                })
        
        # Create DataFrame and save
        result_df = pd.DataFrame(club_results)
        
        if not result_df.empty:
            # Sort by club name, then by rank
            result_df = result_df.sort_values(['Club', 'Rank'])
            
            print(f"Generated {len(result_df)} club entries for {filename}")
            print(f"Number of clubs: {result_df['Club'].nunique()}")
            
            # Show top performers for this category
            if not result_df.empty:
                top_performers = result_df[result_df['Rank'] == 1].nlargest(3, 'Points')
                print(f"Top 3 performers in {filename}:")
                for _, performer in top_performers.iterrows():
                    print(f"  {performer['Name']} ({performer['Club']}) - {performer['Points']:.2f} points")
        
        result_df.to_csv(f'{filename}.csv', index=False)
        print(f"Saved {filename}.csv")
    
    print(f"\nGenerated 4 club result files successfully!")
    
    # Summary statistics
    print("\nSummary:")
    for sex, event, filename in categories:
        try:
            df_check = pd.read_csv(f'{filename}.csv')
            clubs_count = df_check['Club'].nunique() if not df_check.empty else 0
            lifters_count = len(df_check)
            print(f"{filename}: {clubs_count} clubs, {lifters_count} lifters")
        except:
            print(f"{filename}: File error")

if __name__ == "__main__":
    generate_club_results() 