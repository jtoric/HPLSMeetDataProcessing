#!/usr/bin/env python3
"""
Bjelovar Record Breakers - Complete Data Processing Pipeline

This script processes powerlifting competition data from two CSV files and generates
a comprehensive Excel report with results, club rankings, and statistics.

Input files:
- bjelovar/3-bjelovar-record-breakers.opl (1).csv (competition results)
- bjelovar/Bjelovar-record-breakers-finalne-nominacije-2-1-3-1-1-1.csv (club nominations)

Output:
- Bjelovar_Record_Breakers_Rezultati.xlsx (final Excel report)

Intermediate files created (preserved):
- powerlifting_results_processed.csv
- Male_Powerlifting.csv, Female_Powerlifting.csv
- Male_Bench_Only.csv, Female_Bench_Only.csv  
- Male_Powerlifting_Ranking.csv, Female_Powerlifting_Ranking.csv
- Male_Bench_Only_Ranking.csv, Female_Bench_Only_Ranking.csv
"""

import os
import sys

def check_input_files():
    """Check if the required input files exist"""
    required_files = [
        "bjelovar/3-bjelovar-record-breakers.opl (1).csv",
        "bjelovar/Bjelovar-record-breakers-finalne-nominacije-2-1-3-1-1-1.csv"
    ]
    
    missing_files = []
    for file in required_files:
        if not os.path.exists(file):
            missing_files.append(file)
    
    if missing_files:
        print("❌ Missing required input files:")
        for file in missing_files:
            print(f"   - {file}")
        print("\nPlease make sure the input CSV files are in the correct location.")
        return False
    
    print("✅ All required input files found.")
    return True

def run_data_processing():
    """Step 1: Process raw powerlifting data and calculate IPF GL points"""
    print("\n" + "="*60)
    print("KORAK 1: Obrada osnovnih podataka o natjecanju")
    print("="*60)
    
    try:
        from process_powerlifting_data import process_powerlifting_data
        process_powerlifting_data()
        print("✅ Osnovni podaci uspješno obrađeni!")
    except Exception as e:
        print(f"❌ Greška u obradi osnovnih podataka: {e}")
        return False
    
    return True

def run_club_results_generation():
    """Step 2: Generate club results (top 5 per club per category)"""
    print("\n" + "="*60)
    print("KORAK 2: Generiranje rezultata klubova")
    print("="*60)
    
    try:
        from generate_club_results import generate_club_results
        generate_club_results()
        print("✅ Rezultati klubova uspješno generirani!")
    except Exception as e:
        print(f"❌ Greška u generiranju rezultata klubova: {e}")
        return False
    
    return True

def run_club_rankings_generation():
    """Step 3: Generate club rankings (summed points)"""
    print("\n" + "="*60)
    print("KORAK 3: Generiranje rang lista klubova")
    print("="*60)
    
    try:
        from generate_club_rankings import generate_club_rankings
        generate_club_rankings()
        print("✅ Rang liste klubova uspješno generirane!")
    except Exception as e:
        print(f"❌ Greška u generiranju rang lista klubova: {e}")
        return False
    
    return True

def run_excel_report_creation():
    """Step 4: Create the final Excel report"""
    print("\n" + "="*60)
    print("KORAK 4: Kreiranje Excel izvještaja")
    print("="*60)
    
    try:
        from create_excel_report import create_pretty_excel
        filename = create_pretty_excel()
        print(f"✅ Excel izvještaj uspješno kreiran: {filename}")
    except Exception as e:
        print(f"❌ Greška u kreiranju Excel izvještaja: {e}")
        return False
    
    return True

def main():
    """Main processing pipeline"""
    print("🏋️‍♂️ BJELOVAR RECORD BREAKERS - OBRADA PODATAKA 🇭🇷")
    print("=" * 60)
    
    # Check input files
    if not check_input_files():
        sys.exit(1)
    
    # Run processing pipeline
    steps = [
        ("Obrada osnovnih podataka", run_data_processing),
        ("Generiranje rezultata klubova", run_club_results_generation), 
        ("Generiranje rang lista klubova", run_club_rankings_generation),
        ("Kreiranje Excel izvještaja", run_excel_report_creation)
    ]
    
    for step_name, step_function in steps:
        if not step_function():
            print(f"\n❌ Pipeline prekinut na koraku: {step_name}")
            sys.exit(1)
    
    print("\n" + "="*60)
    print("🎉 SVI KORACI USPJEŠNO ZAVRŠENI!")
    print("="*60)
    print("📁 Kreirane datoteke:")
    print("   • powerlifting_results_processed.csv")
    print("   • Male_Powerlifting.csv, Female_Powerlifting.csv")
    print("   • Male_Bench_Only.csv, Female_Bench_Only.csv")
    print("   • Male_Powerlifting_Ranking.csv, Female_Powerlifting_Ranking.csv") 
    print("   • Male_Bench_Only_Ranking.csv, Female_Bench_Only_Ranking.csv")
    print("   • bjelovar/Bjelovar_Record_Breakers_Rezultati.xlsx")
    print("\n🏆 Gotovo! Excel izvještaj je spreman za korištenje.")

if __name__ == "__main__":
    main() 