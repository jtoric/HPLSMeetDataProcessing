import pandas as pd
import numpy as np
import math
from data_loader import load_results, load_clubs

def calculate_ipf_gl_points(bodyweight, total, sex, event):
    """
    Calculate IPF GL Points using official IPF formula: 100/(A-B*exp(-C*bodyweight))
    
    NOTE: Ova funkcija se koristi samo kao fallback ako Points nisu dostupni u podacima.
    U većini slučajeva, Points se uzimaju direktno iz rezultata (kolona "Points" ili "Goodlift").
    """
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

def process_powerlifting_data(input_dir='input'):
    """
    Obrađuje podatke o natjecanju iz input/ foldera.
    
    Args:
        input_dir: Putanja do input foldera (default: 'input')
    """
    # Učitaj rezultate koristeći standardizirani loader
    df_detailed = load_results(input_dir)
    
    # Učitaj podatke o klubovima koristeći standardizirani loader
    club_mapping, birthyear_mapping = load_clubs(input_dir)
    
    # Provjeri da li postoji kolona Division
    if 'Division' not in df_detailed.columns:
        raise ValueError("Datoteka rezultata ne sadrži kolonu 'Division'")
    
    # Filter out divisions starting with "Best"
    df_filtered = df_detailed[~df_detailed['Division'].str.startswith('Best', na=False)]
    
    # Remove rows where Place is NaN or empty (header rows, etc.)
    if 'Place' in df_filtered.columns:
        df_filtered = df_filtered.dropna(subset=['Place'])
        df_filtered = df_filtered[df_filtered['Place'] != '']

    # Exclude NS (No Show) entries entirely from results
    def is_ns_row(row):
        place = str(row.get('Place', '')).strip().upper()
        total = str(row.get('TotalKg', '')).strip().upper()
        return place == 'NS' or total == 'NS'
    ns_mask = df_filtered.apply(is_ns_row, axis=1)
    if ns_mask.any():
        removed_ns = df_filtered.loc[ns_mask, 'Name'].tolist()
        removed_count = len(removed_ns)
        try:
            preview = ', '.join(removed_ns[:10])
            more_text = f' ... (+{removed_count-10} vise)' if removed_count > 10 else ''
            print(f"Uklonjeni NS zapisi: {preview}{more_text}")
        except UnicodeEncodeError:
            # Fallback ako ima problema s encodingom
            print(f"Uklonjeno {removed_count} NS zapisa")
    df_filtered = df_filtered[~ns_mask]
    
    # Create the output dataframe with requested columns
    output_data = []
    
    # Helper funkcija za sigurno dohvaćanje kolone
    def safe_get(row, col, default=''):
        if col in row.index:
            val = row[col]
            return val if pd.notna(val) else default
        return default
    
    for _, row in df_filtered.iterrows():
        # Get club from mapping using normalized name
        normalized_name = str(row['Name']).strip().lower()
        club = club_mapping.get(normalized_name, '')
        
        # Ako nema kluba u mapiranju, pokušaj iz kolone Team (OPL format)
        if not club and 'Team' in row.index and pd.notna(row['Team']):
            club = str(row['Team']).strip()
        
        # Dohvati godinu rođenja
        birth_year_val = birthyear_mapping.get(normalized_name, np.nan)
        if pd.isna(birth_year_val):
            # Pokušaj iz kolone BirthYear ako postoji
            if 'BirthYear' in row.index:
                birth_year_val = pd.to_numeric(row['BirthYear'], errors='coerce')
        
        # Dohvati IPF GL Points iz podataka (ako postoje)
        # OPL format koristi "Points", standardni CSV format koristi "Goodlift"
        ipf_points = None
        
        # Pokušaj prvo "Points" (OPL format)
        if 'Points' in row.index and pd.notna(row['Points']):
            ipf_points = pd.to_numeric(row['Points'], errors='coerce')
        
        # Ako nema Points, pokušaj "Goodlift" (standardni CSV format)
        if (ipf_points is None or pd.isna(ipf_points)) and 'Goodlift' in row.index and pd.notna(row['Goodlift']):
            ipf_points = pd.to_numeric(row['Goodlift'], errors='coerce')
        
        # Fallback: izračunaj ako nema u podacima (rijetko će se dogoditi)
        if ipf_points is None or pd.isna(ipf_points) or ipf_points == 0:
            bodyweight = pd.to_numeric(safe_get(row, 'BodyweightKg'), errors='coerce')
            total = pd.to_numeric(safe_get(row, 'TotalKg'), errors='coerce')
            ipf_points = calculate_ipf_gl_points(
                bodyweight=bodyweight,
                total=total,
                sex=row['Sex'],
                event=row['Event']
            )
            # Log ako se koristi fallback (može biti znak problema s podacima)
            try:
                print(f"Napomena: Points izracunati za {row['Name']} (nema u podacima)")
            except:
                pass
        
        # Zaokruži na 2 decimale
        if pd.notna(ipf_points):
            ipf_points = round(float(ipf_points), 2)
        else:
            ipf_points = 0
        
        # Dodaj Equipment ako postoji u podacima
        # Normaliziraj Equipment: Sleeves, Raw, Wraps = Raw; sve ostalo = Equipped
        equipment = safe_get(row, 'Equipment', 'Raw')
        division_str = str(row.get('Division', '')).lower()
        
        # Provjeri da li Division sadrži "-EQ" sufiks (Equipped)
        if '-eq' in division_str or 'equipped' in division_str:
            equipment = 'Equipped'
        elif not equipment or equipment == '':
            equipment = 'Raw'
        else:
            equipment_lower = str(equipment).lower().strip()
            # Sleeves, Raw, Wraps su Raw format
            if equipment_lower in ['sleeves', 'raw', 'wraps', 'straps']:
                equipment = 'Raw'
            else:
                # Single-ply, Multi-ply, Unlimited, itd. su Equipped
                equipment = 'Equipped'
        
        output_data.append({
            'Place': safe_get(row, 'Place'),
            'Name': row['Name'],
            'Club': club,
            'Sex': row['Sex'],
            'BirthYear': birth_year_val if not pd.isna(birth_year_val) else np.nan,
            'Division': row['Division'],
            'BodyweightKg': safe_get(row, 'BodyweightKg'),
            'WeightClassKg': safe_get(row, 'WeightClassKg'),
            'Squat1Kg': safe_get(row, 'Squat1Kg'),
            'Squat2Kg': safe_get(row, 'Squat2Kg'),
            'Squat3Kg': safe_get(row, 'Squat3Kg'),
            'Best3SquatKg': safe_get(row, 'Best3SquatKg'),
            'Bench1Kg': safe_get(row, 'Bench1Kg'),
            'Bench2Kg': safe_get(row, 'Bench2Kg'),
            'Bench3Kg': safe_get(row, 'Bench3Kg'),
            'Best3BenchKg': safe_get(row, 'Best3BenchKg'),
            'Deadlift1Kg': safe_get(row, 'Deadlift1Kg'),
            'Deadlift2Kg': safe_get(row, 'Deadlift2Kg'),
            'Deadlift3Kg': safe_get(row, 'Deadlift3Kg'),
            'Best3DeadliftKg': safe_get(row, 'Best3DeadliftKg'),
            'TotalKg': safe_get(row, 'TotalKg'),
            'Points': ipf_points,
            'Event': row['Event'],
            'Equipment': equipment
        })
    
    # Create output dataframe
    output_df = pd.DataFrame(output_data)
    
    # Enforce that all competitors have a club
    missing_club_mask = output_df['Club'].isna() | (output_df['Club'].astype(str).str.strip() == '')
    if missing_club_mask.any():
        missing_names = output_df.loc[missing_club_mask, 'Name'].tolist()
        preview = ', '.join(missing_names[:10])
        more = '' if len(missing_names) <= 10 else f" ... (+{len(missing_names) - 10} more)"
        raise ValueError(f"Natjecatelji bez kluba: {preview}{more}. Dodajte klubove u '{input_dir}/klubovi.csv'.")

    # Save to CSV
    output_df.to_csv('powerlifting_results_processed.csv', index=False, encoding='utf-8')
    
    print(f"Processed {len(output_df)} records")
    print(f"Found club information for {sum(1 for club in output_df['Club'] if club != '')} athletes")
    print("Output saved to 'powerlifting_results_processed.csv'")
    
    # Display some sample data (skip if encoding issues)
    try:
        print("\nFirst 5 records:")
        print(output_df.head())
    except UnicodeEncodeError:
        print("\nFirst 5 records: (skipped due to encoding)")
    
    # Show unique divisions
    try:
        print(f"\nUnique divisions (after filtering 'Best' divisions):")
        print(output_df['Division'].unique())
    except UnicodeEncodeError:
        print(f"\nUnique divisions count: {output_df['Division'].nunique()}")

if __name__ == "__main__":
    process_powerlifting_data() 