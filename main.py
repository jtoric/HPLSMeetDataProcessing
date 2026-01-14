#!/usr/bin/env python3
"""
HPLS - Powerlifting Results Processing Pipeline

This script processes powerlifting competition data and generates
comprehensive Excel and PDF reports with results and club rankings.

Input files (in input/ folder):
- rezultati.csv or rezultati.opl.csv (competition results)
- klubovi.csv (club nominations with lifter data)

Output:
- powerlifting_results_processed.csv (processed data)
- rezultati.xlsx (Excel report with rankings and statistics)
- rezultati.pdf (PDF report with results by category)
"""

import os
import sys
import pandas as pd

def check_input_files(input_dir='input'):
    """Check if the required input files exist"""
    from data_loader import detect_results_file
    
    # Provjeri da li postoji input folder
    if not os.path.exists(input_dir):
        print(f"[GRESKA] Input folder '{input_dir}' ne postoji.")
        return False
    
    # Provjeri da li postoji klubovi.csv
    clubs_file = os.path.join(input_dir, 'klubovi.csv')
    if not os.path.exists(clubs_file):
        print(f"[GRESKA] Datoteka s klubovima nije pronadjena: {clubs_file}")
        return False
    
    # Provjeri da li postoji datoteka s rezultatima
    try:
        results_file, _ = detect_results_file(input_dir)
        print(f"[OK] Pronadjena datoteka rezultata: {results_file}")
    except FileNotFoundError as e:
        print(f"[GRESKA] {e}")
        return False
    
    print(f"[OK] Datoteka s klubovima: {clubs_file}")
    print("[OK] Svi potrebni input fajlovi su pronadeni.")
    return True

def run_data_processing(input_dir='input'):
    """Step 1: Process raw powerlifting data and calculate IPF GL points"""
    print("\n" + "="*60)
    print("KORAK 1: Obrada osnovnih podataka o natjecanju")
    print("="*60)
    
    try:
        from process_powerlifting_data import process_powerlifting_data
        process_powerlifting_data(input_dir)
        print("[OK] Osnovni podaci uspjesno obradeni!")
    except Exception as e:
        print(f"[GRESKA] Greska u obradi osnovnih podataka: {e}")
        import traceback
        traceback.print_exc()
        return False
    
    return True


def run_excel_report_creation():
    """
    Step 2: Create formatted Excel report.
    
    Creates rezultati.xlsx with:
    - Individual results (by sex and event)
    - Club rankings (Raw and Equipped separated)
    - Statistics (Top 5 performers by category)
    """
    print("\n" + "="*60)
    print("KORAK 2: Kreiranje Excel izvjestaja")
    print("="*60)
    
    try:
        from create_excel_report import create_pretty_excel
        
        # Create single Excel with all results (Raw and Equipped)
        print("\nKreiranje Excel izvjestaja...")
        create_pretty_excel(equipment_filter=None, output_filename='rezultati.xlsx')
        print("[OK] Excel izvjestaj kreiran: rezultati.xlsx")
        
        return True
        
    except Exception as e:
        print(f"\n[GRESKA] Greska u kreiranju Excel izvjestaja: {e}")
        return False


def run_pdf_report_creation():
    """
    Step 3: Create PDF report with results.
    
    Creates rezultati.pdf with:
    - Male Powerlifting results by category
    - Female Powerlifting results by category
    - Male Bench Only results by category
    - Female Bench Only results by category
    """
    print("\n" + "="*60)
    print("KORAK 3: Kreiranje PDF izvjestaja")
    print("="*60)
    
    try:
        from create_pdf_report import create_pdf_report
        
        print("\nKreiranje PDF izvjestaja...")
        create_pdf_report(output_filename='rezultati.pdf')
        print("[OK] PDF izvjestaj kreiran: rezultati.pdf")
        
        return True
        
    except Exception as e:
        print(f"\n[GRESKA] Greska u kreiranju PDF izvjestaja: {e}")
        import traceback
        traceback.print_exc()
        return False

def main(input_dir='input'):
    """
    Main processing pipeline.
    
    Executes all steps in sequence:
    1. Data processing (load, map, calculate)
    2. Excel report generation
    """
    print("HPLS - OBRADA PODATAKA")
    print("=" * 60)
    print(f"Koristi se input folder: {input_dir}")
    print("=" * 60)
    
    # Check input files
    if not check_input_files(input_dir):
        sys.exit(1)
    
    # Run processing pipeline
    if not run_data_processing(input_dir):
        print("\n[GRESKA] Pipeline prekinut na koraku: Obrada podataka")
        sys.exit(1)
    
    if not run_excel_report_creation():
        print("\n[GRESKA] Pipeline prekinut na koraku: Kreiranje Excel izvjestaja")
        sys.exit(1)
    
    if not run_pdf_report_creation():
        print("\n[GRESKA] Pipeline prekinut na koraku: Kreiranje PDF izvjestaja")
        sys.exit(1)
    
    print("\n" + "="*60)
    print("SVI KORACI USPJESNO ZAVRSENI!")
    print("="*60)
    print("Kreirane datoteke:")
    print("   - powerlifting_results_processed.csv (obradeni podaci)")
    print("   - rezultati.xlsx (finalni Excel izvjestaj)")
    print("   - rezultati.pdf (finalni PDF izvjestaj)")
    print("\nGotovo! Izvjestaji su spremni za koristenje.")

if __name__ == "__main__":
    main() 