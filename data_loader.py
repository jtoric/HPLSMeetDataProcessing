"""
Standardizirani modul za učitavanje podataka o natjecanjima.

Ovaj modul automatski detektira format rezultata (.csv ili .opl.csv) i učitava ih
u standardizirani format za daljnju obradu.
"""

import pandas as pd
import os
import glob
from pathlib import Path


def detect_results_file(input_dir='input'):
    """
    Automatski pronalazi datoteku s rezultatima u input/ folderu.
    
    Args:
        input_dir: Putanja do input foldera (default: 'input')
    
    Returns:
        tuple: (putanja_datoteke, format) gdje je format 'csv' ili 'opl'
    
    Raises:
        FileNotFoundError: Ako nema datoteka s rezultatima
        ValueError: Ako ima više datoteka s rezultatima
    """
    input_path = Path(input_dir)
    
    # Pronađi sve CSV datoteke osim klubovi.csv
    csv_files = list(input_path.glob('*.csv'))
    csv_files = [f for f in csv_files if f.name != 'klubovi.csv']
    
    if not csv_files:
        raise FileNotFoundError(
            f"Nema datoteka s rezultatima u '{input_dir}' folderu. "
            f"Očekivane datoteke: *.csv ili *.opl.csv"
        )
    
    if len(csv_files) > 1:
        # Ako ima više datoteka, pokušaj pronaći .opl.csv prvo
        opl_files = [f for f in csv_files if f.name.endswith('.opl.csv')]
        if opl_files:
            return str(opl_files[0]), 'opl'
        
        # Ako nema .opl.csv, uzmi prvu .csv datoteku
        regular_csv = [f for f in csv_files if not f.name.endswith('.opl.csv')]
        if regular_csv:
            return str(regular_csv[0]), 'csv'
        
        # Fallback: uzmi prvu datoteku
        return str(csv_files[0]), 'csv'
    
    # Jedna datoteka - provjeri format
    file_path = csv_files[0]
    if file_path.name.endswith('.opl.csv'):
        return str(file_path), 'opl'
    else:
        return str(file_path), 'csv'


def load_results_csv(file_path):
    """
    Učitava standardni CSV format rezultata.
    
    Args:
        file_path: Putanja do CSV datoteke
    
    Returns:
        pd.DataFrame: DataFrame s rezultatima
    """
    df = pd.read_csv(file_path)
    
    # Provjeri da li ima potrebne kolone
    required_cols = ['Name', 'Sex', 'Event']
    missing_cols = [col for col in required_cols if col not in df.columns]
    
    if missing_cols:
        raise ValueError(
            f"CSV datoteka ne sadrži potrebne kolone: {missing_cols}. "
            f"Pronađene kolone: {df.columns.tolist()}"
        )
    
    return df


def load_results_opl(file_path):
    """
    Učitava OPL format rezultata (.opl.csv).
    
    OPL format obično ima 4-5 metadata redova prije headera.
    
    Args:
        file_path: Putanja do OPL CSV datoteke
    
    Returns:
        pd.DataFrame: DataFrame s rezultatima
    """
    # OPL format obično ima header na redu 6 (indeks 5, skiprows=5)
    # Pokušaj sa skiprows=5 prvo (najčešći slučaj)
    df = pd.read_csv(file_path, skiprows=5)
    
    # Provjeri da li ima potrebne kolone
    required_cols = ['Name', 'Sex', 'Event']
    has_required = all(col in df.columns for col in required_cols)
    
    if not has_required:
        # Pokušaj sa skiprows=4
        df = pd.read_csv(file_path, skiprows=4)
        has_required = all(col in df.columns for col in required_cols)
    
    if not has_required:
        # Pokušaj sa skiprows=6
        df = pd.read_csv(file_path, skiprows=6)
        has_required = all(col in df.columns for col in required_cols)
    
    if not has_required:
        # Pokušaj bez skiprows (ako je već čist CSV)
        df = pd.read_csv(file_path)
        has_required = all(col in df.columns for col in required_cols)
    
    if not has_required:
        raise ValueError(
            f"OPL datoteka ne sadrži potrebne kolone. "
            f"Pronadjene kolone: {df.columns.tolist()}"
        )
    
    return df


def load_results(input_dir='input'):
    """
    Glavna funkcija za učitavanje rezultata - automatski detektira format.
    
    Args:
        input_dir: Putanja do input foldera (default: 'input')
    
    Returns:
        pd.DataFrame: DataFrame s rezultatima
    """
    file_path, file_format = detect_results_file(input_dir)
    
    print(f"Pronadjena datoteka rezultata: {file_path}")
    print(f"Format: {file_format.upper()}")
    
    if file_format == 'opl':
        df = load_results_opl(file_path)
    else:
        df = load_results_csv(file_path)
    
    print(f"Ucitano {len(df)} zapisa")
    print(f"Kolone: {', '.join(df.columns.tolist()[:10])}{'...' if len(df.columns) > 10 else ''}")
    
    return df


def load_clubs(input_dir='input', clubs_file='klubovi.csv'):
    """
    Učitava podatke o klubovima iz input/klubovi.csv.
    
    Očekivana struktura:
    - Red 1-2: Naslov/prazni redovi
    - Red 3: Header (KATEGORIJA, IME, PREZIME, GODIŠTE, KLUB, TOTAL)
    - Red 4+: Podaci
    
    Args:
        input_dir: Putanja do input foldera (default: 'input')
        clubs_file: Naziv datoteke s klubovima (default: 'klubovi.csv')
    
    Returns:
        tuple: (club_mapping, birthyear_mapping) gdje su:
            - club_mapping: dict {normalized_name: club_name}
            - birthyear_mapping: dict {normalized_name: birth_year}
    """
    file_path = Path(input_dir) / clubs_file
    
    if not file_path.exists():
        raise FileNotFoundError(
            f"Datoteka s klubovima nije pronađena: {file_path}"
        )
    
    # Učitaj datoteku - preskoči prva 2 reda, koristi red 3 kao header
    df = pd.read_csv(file_path, skiprows=2, encoding='utf-8')
    
    # Pronađi kolone
    name_col = None
    surname_col = None
    club_col = None
    birthyear_col = None
    
    for col in df.columns:
        col_upper = str(col).upper().strip()
        if col_upper == 'IME':
            name_col = col
        elif col_upper == 'PREZIME':
            surname_col = col
        elif col_upper == 'KLUB':
            club_col = col
        elif col_upper in ['GODIŠTE', 'GODISTE', 'BIRTHYEAR']:
            birthyear_col = col
    
    if not name_col or not surname_col or not club_col:
        raise ValueError(
            f"Datoteka s klubovima ne sadrži potrebne kolone. "
            f"Pronađene kolone: {df.columns.tolist()}. "
            f"Očekivane: IME, PREZIME, KLUB"
        )
    
    # Kreiraj mapiranja
    club_mapping = {}
    birthyear_mapping = {}
    
    for _, row in df.iterrows():
        try:
            ime = str(row[name_col]).strip() if pd.notna(row[name_col]) else ''
            prezime = str(row[surname_col]).strip() if pd.notna(row[surname_col]) else ''
            
            # Preskoči header redove i prazne redove
            if not ime or not prezime or ime.upper() == 'IME' or prezime.upper() == 'PREZIME':
                continue
            
            full_name = f"{ime} {prezime}".strip()
            normalized_name = full_name.lower()
            
            # Dodaj klub
            if pd.notna(row[club_col]) and str(row[club_col]).strip():
                club_mapping[normalized_name] = str(row[club_col]).strip()
            
            # Dodaj godinu rođenja
            if birthyear_col and pd.notna(row[birthyear_col]):
                try:
                    year = pd.to_numeric(row[birthyear_col], errors='coerce')
                    if not pd.isna(year):
                        birthyear_mapping[normalized_name] = int(year)
                except:
                    pass
        except Exception as e:
            # Preskoči redove s greškama
            continue
    
    print(f"Ucitano {len(club_mapping)} mapiranja klubova")
    print(f"Ucitano {len(birthyear_mapping)} mapiranja godina rodjenja")
    
    return club_mapping, birthyear_mapping


if __name__ == "__main__":
    # Test
    import sys
    import io
    
    # Postavi UTF-8 encoding za stdout
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    
    print("Testiranje ucitavanja podataka...")
    print("\n1. Ucitavanje rezultata:")
    df = load_results()
    print(f"\n2. Ucitavanje klubova:")
    club_map, birth_map = load_clubs()
    print(f"\n3. Test mapiranja:")
    if club_map:
        sample_name = list(club_map.keys())[0]
        print(f"   Primjer: '{sample_name}' -> Klub: {club_map[sample_name]}")

