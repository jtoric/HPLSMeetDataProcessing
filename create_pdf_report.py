"""
PDF Report Generator for Powerlifting Competition Results

Creates formatted PDF reports from processed powerlifting data.
Results are organized by: Male Powerlifting, Female Powerlifting,
Male Bench Only, Female Bench Only.

Each section contains results grouped by division (Kadeti, Juniori, Seniori, Veterani)
and weight class (from lightest to heaviest).
"""

import os
from typing import Tuple, List, Any

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm, cm
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, 
    Paragraph, Spacer, PageBreak
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


# =============================================================================
# CONSTANTS
# =============================================================================

# Division ordering (age categories)
DIVISION_ORDER = {
    'Sub-Junior': 1,  # Kadeti
    'Junior': 2,      # Juniori
    'Open': 3,        # Seniori
    'Master I': 4,    # Veterani 1
    'Master II': 5,   # Veterani 2
    'Master III': 6,  # Veterani 3
    'Master IV': 7    # Veterani 4
}

# Division translations (English -> Croatian)
DIVISION_TRANSLATIONS = {
    'Sub-Junior': 'KADETI',
    'Junior': 'JUNIORI',
    'Open': 'SENIORI',
    'Master I': 'VETERANI 1',
    'Master II': 'VETERANI 2',
    'Master III': 'VETERANI 3',
    'Master IV': 'VETERANI 4'
}

# Report sections configuration
REPORT_SECTIONS = [
    ('M', 'SBD', 'MUŠKI POWERLIFTING'),
    ('F', 'SBD', 'ŽENSKI POWERLIFTING'),
    ('M', 'B', 'MUŠKI POTISAK S KLUPE'),
    ('F', 'B', 'ŽENSKI POTISAK S KLUPE')
]

# Table column configuration
POWERLIFTING_HEADERS = [
    'Plasman', 'Ime i prezime', 'Klub', 'Godište', 'Težina',
    'Čučanj', 'Potisak s klupe', 'Mrtvo dizanje', 'Total', 'GL bodovi'
]
POWERLIFTING_WIDTHS = [14*mm, 45*mm, 38*mm, 14*mm, 14*mm, 16*mm, 22*mm, 20*mm, 16*mm, 18*mm]

BENCH_HEADERS = [
    'Plasman', 'Ime i prezime', 'Klub', 'Godište', 'Težina',
    'Potisak 1', 'Potisak 2', 'Potisak 3', 'Najbolji potisak', 'GL bodovi'
]
BENCH_WIDTHS = [14*mm, 45*mm, 38*mm, 14*mm, 14*mm, 18*mm, 18*mm, 18*mm, 24*mm, 18*mm]

# Color scheme
COLORS = {
    'header_dark': colors.HexColor('#1a1a2e'),
    'header_medium': colors.HexColor('#16213e'),
    'header_light': colors.HexColor('#0f3460'),
    'row_alt': colors.HexColor('#f0f0f0'),
    'grid': colors.HexColor('#cccccc'),
    'white': colors.white
}

# Font paths for UTF-8 support
FONT_PATHS = [
    'C:/Windows/Fonts/arial.ttf',
    'C:/Windows/Fonts/calibri.ttf',
    'C:/Windows/Fonts/tahoma.ttf',
    'C:/Windows/Fonts/verdana.ttf',
    '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',
    '/usr/share/fonts/TTF/DejaVuSans.ttf',
]


# =============================================================================
# FONT REGISTRATION
# =============================================================================

def register_fonts() -> Tuple[str, str]:
    """
    Register fonts that support Croatian characters (čćšđž).
    
    Returns:
        Tuple of (regular_font_name, bold_font_name)
    """
    for font_path in FONT_PATHS:
        if not os.path.exists(font_path):
            continue
        
        try:
            pdfmetrics.registerFont(TTFont('CustomFont', font_path))
            
            # Try to find bold variant
            bold_path = font_path.replace('.ttf', 'bd.ttf').replace('.TTF', 'bd.TTF')
            if not os.path.exists(bold_path):
                bold_path = font_path.replace('.ttf', '-Bold.ttf')
            
            if os.path.exists(bold_path):
                pdfmetrics.registerFont(TTFont('CustomFont-Bold', bold_path))
            else:
                pdfmetrics.registerFont(TTFont('CustomFont-Bold', font_path))
            
            return 'CustomFont', 'CustomFont-Bold'
        except Exception:
            continue
    
    # Fallback to Helvetica (no Croatian chars support)
    return 'Helvetica', 'Helvetica-Bold'


# =============================================================================
# DATA PARSING UTILITIES
# =============================================================================

def get_division_type(division_name) -> str:
    """
    Extract standardized division type from various input formats.
    
    Handles formats like: 'Junior', 'Juniors', 'Kadet', 'Master 1', 'Masters 2', etc.
    """
    if pd.isna(division_name):
        return 'Open'
    
    text = str(division_name).strip()
    text_lower = text.lower()
    
    # Check Masters (must check IV before I to avoid false matches)
    if 'master iv' in text_lower or 'masters 4' in text_lower or text == 'Master 4':
        return 'Master IV'
    if 'master iii' in text_lower or 'masters 3' in text_lower or text == 'Master 3':
        return 'Master III'
    if 'master ii' in text_lower or 'masters 2' in text_lower or text == 'Master 2':
        return 'Master II'
    if 'master i' in text_lower or 'masters 1' in text_lower or text == 'Master 1':
        return 'Master I'
    
    # Check Sub-Junior/Kadet
    if 'sub-junior' in text_lower or 'sub-juniors' in text_lower or text_lower == 'kadet':
        return 'Sub-Junior'
    
    # Check Junior (but not Sub-Junior)
    if 'junior' in text_lower and 'sub-junior' not in text_lower:
        return 'Junior'
    
    # Check Open/Guest
    if 'open' in text_lower or 'guest' in text_lower:
        return 'Open'
    
    return 'Open'


def parse_weight_class(weight_class) -> float:
    """
    Parse weight class string to numeric value for sorting.
    
    Handles formats like: '83', '84+', '120+', '69 kg'
    """
    if pd.isna(weight_class):
        return 999.0
    
    wc_str = str(weight_class).replace('+', '').replace('kg', '').strip()
    try:
        return float(wc_str)
    except ValueError:
        return 999.0


def parse_place(place) -> int:
    """
    Parse place/rank to numeric value for sorting.
    
    Handles: 1, 2, 3, 'DQ', 'DNS', None
    """
    if pd.isna(place):
        return 999
    try:
        return int(float(place))
    except (ValueError, TypeError):
        return 998  # DQ/DNS go to end but before empty


def format_number(value, decimals: int = 1) -> str:
    """
    Format numeric value for display with Croatian decimal separator.
    
    Args:
        value: Number to format (can be NaN)
        decimals: Number of decimal places (0, 1, or 2)
    
    Returns:
        Formatted string with comma as decimal separator
    """
    if pd.isna(value):
        return ''
    
    try:
        num = float(str(value).replace(',', '.'))
        if decimals == 0:
            return str(int(num))
        elif decimals == 1:
            return f"{num:.1f}".replace('.', ',')
        else:
            return f"{num:.2f}".replace('.', ',')
    except (ValueError, TypeError):
        return str(value)


def format_place(place_value) -> str:
    """Format place/rank for display."""
    if pd.isna(place_value):
        return ''
    
    if isinstance(place_value, (int, float)):
        try:
            return str(int(place_value))
        except (ValueError, TypeError):
            pass
    
    return str(place_value)


def format_year(year_value) -> str:
    """Format birth year for display."""
    if pd.isna(year_value):
        return ''
    
    try:
        return str(int(year_value))
    except (ValueError, TypeError):
        return str(year_value)


# =============================================================================
# TABLE BUILDING
# =============================================================================

def create_table_styles(font_name: str, font_bold: str) -> TableStyle:
    """Create standard table styling for result tables."""
    return TableStyle([
        # Weight class header row (row 0)
        ('SPAN', (0, 0), (-1, 0)),
        ('BACKGROUND', (0, 0), (-1, 0), COLORS['header_dark']),
        ('TEXTCOLOR', (0, 0), (-1, 0), COLORS['white']),
        ('FONTNAME', (0, 0), (-1, 0), font_bold),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        
        # Column headers row (row 1)
        ('BACKGROUND', (0, 1), (-1, 1), COLORS['header_medium']),
        ('TEXTCOLOR', (0, 1), (-1, 1), COLORS['white']),
        ('FONTNAME', (0, 1), (-1, 1), font_bold),
        ('FONTSIZE', (0, 1), (-1, 1), 7),
        ('ALIGN', (0, 1), (-1, 1), 'CENTER'),
        
        # Data rows (row 2+)
        ('FONTNAME', (0, 2), (-1, -1), font_name),
        ('FONTSIZE', (0, 2), (-1, -1), 8),
        ('ALIGN', (0, 2), (0, -1), 'CENTER'),   # Place column
        ('ALIGN', (3, 2), (-1, -1), 'CENTER'),  # Numeric columns
        
        # Alternating row colors
        ('ROWBACKGROUNDS', (0, 2), (-1, -1), [COLORS['white'], COLORS['row_alt']]),
        
        # Grid and borders
        ('GRID', (0, 0), (-1, -1), 0.5, COLORS['grid']),
        ('BOX', (0, 0), (-1, -1), 1, COLORS['header_dark']),
        
        # Cell padding
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ('LEFTPADDING', (0, 0), (-1, -1), 3),
        ('RIGHTPADDING', (0, 0), (-1, -1), 3),
    ])


def build_result_row(row: pd.Series, event: str) -> List[str]:
    """
    Build a single data row for the results table.
    
    Args:
        row: DataFrame row with lifter data
        event: 'SBD' for powerlifting, 'B' for bench only
    
    Returns:
        List of formatted cell values
    """
    common_cols = [
        format_place(row['Place']),
        str(row['Name']) if pd.notna(row['Name']) else '',
        str(row['Club']) if pd.notna(row['Club']) else '',
        format_year(row['BirthYear']),
        format_number(row['BodyweightKg']),
    ]
    
    if event == 'SBD':
        return common_cols + [
            format_number(row['Best3SquatKg']),
            format_number(row['Best3BenchKg']),
            format_number(row['Best3DeadliftKg']),
            format_number(row['TotalKg']),
            format_number(row['Points'], decimals=2)
        ]
    else:
        return common_cols + [
            format_number(row['Bench1Kg']),
            format_number(row['Bench2Kg']),
            format_number(row['Bench3Kg']),
            format_number(row['Best3BenchKg']),
            format_number(row['Points'], decimals=2)
        ]


def create_weight_class_table(
    wc_df: pd.DataFrame, 
    weight_class: Any, 
    event: str,
    font_name: str,
    font_bold: str
) -> Table:
    """
    Create a formatted table for a single weight class.
    
    Args:
        wc_df: DataFrame with lifters in this weight class
        weight_class: Weight class value (e.g., '83', '84+')
        event: 'SBD' for powerlifting, 'B' for bench only
        font_name: Regular font name
        font_bold: Bold font name
    
    Returns:
        Styled reportlab Table
    """
    # Select headers and widths based on event
    if event == 'SBD':
        headers = POWERLIFTING_HEADERS
        col_widths = POWERLIFTING_WIDTHS
    else:
        headers = BENCH_HEADERS
        col_widths = BENCH_WIDTHS
    
    # Build table data
    table_data = []
    
    # Weight class header row
    wc_header = f"{weight_class} kg"
    table_data.append([wc_header] + [''] * (len(headers) - 1))
    
    # Column headers row
    table_data.append(headers)
    
    # Data rows
    for _, row in wc_df.iterrows():
        table_data.append(build_result_row(row, event))
    
    # Create and style table
    table = Table(table_data, colWidths=col_widths)
    table.setStyle(create_table_styles(font_name, font_bold))
    
    return table


# =============================================================================
# PARAGRAPH STYLES
# =============================================================================

def create_paragraph_styles(font_bold: str) -> dict:
    """Create custom paragraph styles for section headers."""
    styles = getSampleStyleSheet()
    
    return {
        'title': ParagraphStyle(
            'CustomTitle',
            parent=styles['Heading1'],
            fontName=font_bold,
            fontSize=18,
            spaceAfter=10,
            alignment=1,  # Center
            textColor=COLORS['header_dark']
        ),
        'section': ParagraphStyle(
            'SectionTitle',
            parent=styles['Heading2'],
            fontName=font_bold,
            fontSize=14,
            spaceBefore=15,
            spaceAfter=8,
            textColor=COLORS['header_medium']
        ),
        'category': ParagraphStyle(
            'CategoryTitle',
            parent=styles['Heading3'],
            fontName=font_bold,
            fontSize=11,
            spaceBefore=10,
            spaceAfter=5,
            textColor=COLORS['header_light']
        )
    }


# =============================================================================
# DATA PREPARATION
# =============================================================================

def prepare_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """
    Add sorting columns to the dataframe.
    
    Adds: DivisionType, DivisionOrder, WeightClassSort, PlaceSort
    """
    df = df.copy()
    df['DivisionType'] = df['Division'].apply(get_division_type)
    df['DivisionOrder'] = df['DivisionType'].map(DIVISION_ORDER).fillna(99)
    df['WeightClassSort'] = df['WeightClassKg'].apply(parse_weight_class)
    df['PlaceSort'] = df['Place'].apply(parse_place)
    return df


# =============================================================================
# MAIN REPORT GENERATION
# =============================================================================

def create_pdf_report(output_filename: str = 'rezultati.pdf') -> None:
    """
    Create PDF report with competition results.
    
    Generates a landscape A4 PDF with results organized by:
    1. Sex and Event (Male/Female, Powerlifting/Bench)
    2. Equipment (Raw/Equipped)
    3. Division (Kadeti → Juniori → Seniori → Veterani)
    4. Weight Class (lightest → heaviest)
    
    Args:
        output_filename: Output PDF file path
    """
    # Register UTF-8 compatible fonts
    font_name, font_bold = register_fonts()
    
    # Load and prepare data
    df = pd.read_csv('powerlifting_results_processed.csv', encoding='utf-8')
    df = prepare_dataframe(df)
    
    # Create PDF document
    doc = SimpleDocTemplate(
        output_filename,
        pagesize=landscape(A4),
        leftMargin=1*cm,
        rightMargin=1*cm,
        topMargin=1*cm,
        bottomMargin=1*cm,
        title='Rezultati',
        author='HPLS',
        subject='Rezultati natjecanja'
    )
    
    # Create styles
    para_styles = create_paragraph_styles(font_bold)
    
    # Build document elements
    elements = []
    first_section = True
    
    for sex, event, title in REPORT_SECTIONS:
        section_elements = build_section(
            df, sex, event, title,
            para_styles, font_name, font_bold,
            add_page_break=(not first_section)
        )
        
        if section_elements:
            elements.extend(section_elements)
            first_section = False
    
    # Build PDF
    doc.build(elements)
    print(f"PDF izvještaj kreiran: {output_filename}")


def build_section(
    df: pd.DataFrame,
    sex: str,
    event: str,
    title: str,
    para_styles: dict,
    font_name: str,
    font_bold: str,
    add_page_break: bool = False
) -> List:
    """
    Build all elements for a report section (e.g., Male Powerlifting).
    
    Returns:
        List of reportlab elements, or empty list if no data
    """
    # Filter data for this section
    section_df = df[(df['Sex'] == sex) & (df['Event'] == event)]
    
    if section_df.empty:
        return []
    
    elements = []
    
    # Page break before section (except first)
    if add_page_break:
        elements.append(PageBreak())
    
    # Section title
    elements.append(Paragraph(title, para_styles['title']))
    elements.append(Spacer(1, 5*mm))
    
    # Check for equipped lifters
    has_equipped = (
        'Equipment' in section_df.columns and 
        (section_df['Equipment'] == 'Equipped').any()
    )
    
    # Process by equipment type
    equipment_types = ['Raw', 'Equipped'] if has_equipped else ['Raw']
    
    for equipment in equipment_types:
        if 'Equipment' in section_df.columns:
            equip_df = section_df[section_df['Equipment'] == equipment]
        else:
            equip_df = section_df
        
        if equip_df.empty:
            continue
        
        # Equipment subtitle (only if both Raw and Equipped exist)
        if has_equipped:
            equip_title = "RAW" if equipment == 'Raw' else "EQUIPPED"
            elements.append(Paragraph(equip_title, para_styles['section']))
        
        # Sort data
        equip_df = equip_df.sort_values(['DivisionOrder', 'WeightClassSort', 'PlaceSort'])
        
        # Build tables for each division
        division_elements = build_division_tables(
            equip_df, event, para_styles, font_name, font_bold
        )
        elements.extend(division_elements)
    
    return elements


def build_division_tables(
    equip_df: pd.DataFrame,
    event: str,
    para_styles: dict,
    font_name: str,
    font_bold: str
) -> List:
    """
    Build tables for all divisions within an equipment category.
    
    Returns:
        List of reportlab elements (paragraphs, tables, spacers)
    """
    elements = []
    
    # Get unique divisions in correct order
    divisions = sorted(
        equip_df['DivisionType'].unique(),
        key=lambda x: DIVISION_ORDER.get(x, 99)
    )
    
    for division_type in divisions:
        div_df = equip_df[equip_df['DivisionType'] == division_type]
        
        # Division header
        div_title = DIVISION_TRANSLATIONS.get(division_type, division_type.upper())
        elements.append(Paragraph(div_title, para_styles['category']))
        
        # Get weight classes sorted by numeric value
        weight_classes = (
            div_df[['WeightClassKg', 'WeightClassSort']]
            .drop_duplicates()
            .sort_values('WeightClassSort')
        )
        
        # Build table for each weight class
        for _, wc_row in weight_classes.iterrows():
            weight_class = wc_row['WeightClassKg']
            wc_df = div_df[div_df['WeightClassKg'] == weight_class].sort_values('PlaceSort')
            
            table = create_weight_class_table(
                wc_df, weight_class, event, font_name, font_bold
            )
            elements.append(table)
            elements.append(Spacer(1, 3*mm))
    
    return elements


# =============================================================================
# ENTRY POINT
# =============================================================================

if __name__ == '__main__':
    create_pdf_report()
