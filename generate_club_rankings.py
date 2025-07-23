import pandas as pd

def generate_club_rankings():
    """Generate club ranking tables based on summed GL Points from top lifters"""
    
    categories = [
        ('Male_Powerlifting.csv', 'Male_Powerlifting_Ranking.csv'),
        ('Female_Powerlifting.csv', 'Female_Powerlifting_Ranking.csv'),
        ('Male_Bench_Only.csv', 'Male_Bench_Only_Ranking.csv'),
        ('Female_Bench_Only.csv', 'Female_Bench_Only_Ranking.csv')
    ]
    
    for input_file, output_file in categories:
        print(f"\nProcessing {input_file}...")
        
        try:
            # Read the club results file
            df = pd.read_csv(input_file)
            
            if df.empty:
                print(f"No data in {input_file}, creating empty ranking...")
                empty_df = pd.DataFrame(columns=['Place', 'Club', 'Points'])
                empty_df.to_csv(output_file, index=False)
                continue
            
            # Sum points for each club
            club_totals = df.groupby('Club')['Points'].sum().reset_index()
            
            # Sort by total points (descending) and add ranking
            club_totals = club_totals.sort_values('Points', ascending=False).reset_index(drop=True)
            club_totals['Place'] = range(1, len(club_totals) + 1)
            
            # Reorder columns to match required format: Place, Club, Points
            club_rankings = club_totals[['Place', 'Club', 'Points']].copy()
            
            # Round points to 2 decimal places
            club_rankings['Points'] = club_rankings['Points'].round(2)
            
            # Save the ranking
            club_rankings.to_csv(output_file, index=False)
            
            print(f"Generated {output_file}")
            print(f"Number of clubs: {len(club_rankings)}")
            
            # Show top 3 clubs
            top_3 = club_rankings.head(3)
            print("Top 3 clubs:")
            for _, row in top_3.iterrows():
                print(f"  {row['Place']}. {row['Club']} - {row['Points']:.2f} points")
                
        except Exception as e:
            print(f"Error processing {input_file}: {e}")
    
    print(f"\nGenerated 4 club ranking files successfully!")
    
    # Summary
    print(f"\nClub Rankings Summary:")
    for input_file, output_file in categories:
        try:
            df = pd.read_csv(output_file)
            if not df.empty:
                winner = df.iloc[0]
                print(f"{output_file}: {winner['Club']} wins with {winner['Points']:.2f} points ({len(df)} clubs total)")
            else:
                print(f"{output_file}: No clubs")
        except:
            print(f"{output_file}: File error")

if __name__ == "__main__":
    generate_club_rankings() 