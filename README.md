# HPLS Powerlifting Data Processing System

Standardizirani sistem za obradu rezultata powerlifting natjecanja i generiranje Excel izvještaja.

## 🎯 Značajke

- **Automatska detekcija formata** - podržava `.csv` i `.opl.csv` formate
- **Mapiranje klubova** - automatsko spajanje natjecatelja s njihovim klubovima
- **GL Points** - koristi postojeće bodove iz rezultata (fallback na kalkulaciju ako nedostaju)
- **Raw/Equipped odvajanje** - odvojeni rangovi i Top 5 za Raw i Equipped natjecatelje
- **Club Rankings** - rang klubova baziran na top-5 natjecatelja po klubu
- **Formatiran Excel izvještaj** - profesionalno formatiran s bojama medalja, kategorijama i statistikom

## 📁 Struktura Projekta

```
obradarezultata/
├── main.py                              # Glavni script - pokreće cijeli pipeline
├── data_loader.py                       # Učitavanje podataka (automatska detekcija formata)
├── process_powerlifting_data.py         # Obrada podataka i mapiranje klubova
├── create_excel_report.py               # Generiranje Excel izvještaja
├── input/                               # INPUT folder
│   ├── klubovi.csv                      # Podaci o klubovima (obavezno)
│   └── rezultati.csv ili .opl.csv       # Rezultati natjecanja (jedan format)
├── powerlifting_results_processed.csv   # Obrađeni podaci (izlaz)
└── rezultati.xlsx                       # Finalni Excel izvještaj (izlaz)
```

## 🚀 Instalacija

1. **Instaliraj Python dependencies:**
```bash
pip install -r requirements.txt
```

2. **Pripremi input datoteke:**
   - Stavi `klubovi.csv` u `input/` folder
   - Stavi `rezultati.csv` ILI `rezultati.opl.csv` u `input/` folder

## 📊 Input Formati

### 1. klubovi.csv (obavezno)

Format klubova može biti bilo koji, ali mora sadržavati:
- Ime i prezime natjecatelja
- Naziv kluba
- Godište (opcionalno)

Primjer:
```
,KATEGORIJA,IME,PREZIME,GODIŠTE,KLUB,TOTAL
,ŽENE,,,,,
,JUNIOR,,,,,
,57,,,,,
,,Matea,Kucljak,2003,Galacticos,267.5
```

### 2a. rezultati.csv (Standard OpenPowerlifting format)

Standardni CSV format s kolonama:
- `Name`, `Sex`, `Event`, `Equipment`, `Division`, `BodyweightKg`
- `WeightClassKg`, `Best3SquatKg`, `Best3BenchKg`, `Best3DeadliftKg`
- `TotalKg`, `Goodlift` (GL Points), itd.

### 2b. rezultati.opl.csv (OpenLifter format)

OPL format s metadata linijama na početku:
```
OPL Format v1 (OpenLifter 1.4),...
Federation,Date,MeetCountry,...
HPLS,'2025-12-18,Croatia,...
Place,Name,Sex,Country,Equipment,Division,...
1,Matea Kucljak,F,Croatia,Sleeves,Junior,...
```

## 🔧 Korištenje

### Jednostavno pokretanje:
```bash
python main.py
```

Pipeline se sastoji od **2 koraka**:

1. **Obrada podataka**
   - Učitavanje rezultata i klubova
   - Mapiranje natjecatelja na klubove
   - Normalizacija Equipment tipova (Raw/Equipped)
   - Ekstrakcija ili kalkulacija GL Points
   - Generira: `powerlifting_results_processed.csv`

2. **Generiranje Excel izvještaja**
   - Individualni rezultati po kategorijama
   - Rang klubova (Raw i Equipped odvojeno)
   - Top 5 statistika (Raw i Equipped odvojeno)
   - Generira: `rezultati.xlsx`

## 📈 Excel Izvještaj - Sadržaj

### 1. Muški Powerlifting
Svi muški powerlifting rezultati sortirani po:
- Kategorija (Kadeti → Juniori → Seniori → Veterani)
- Težinska klasa
- Mjesto

**Headerovi kategorija:**
- ═══ **KADETI KATEGORIJA** ═══
- ═══ **JUNIORI KATEGORIJA** ═══
- ═══ **SENIORI KATEGORIJA** ═══
- ═══ **VETERANI 1/2/3 KATEGORIJA** ═══

### 2. Ženski Powerlifting
Isti format kao muški powerlifting.

### 3. Muški Potisak s klupe
Svi muški bench only rezultati (isti format).

### 4. Ženski Potisak s klupe
Svi ženski bench only rezultati (isti format).

### 5. Rang Klubova

**Format za svaku kategoriju:**

```
Muški Powerlifting Rang Klubova

Mjesto | Klub              | Bodovi
1      | Štanga            | 468.23  (🥇 zlatna)
2      | Galacticos        | 464.68  (🥈 srebrna)
3      | Gumeni medvjedići | 449.49  (🥉 brončana)
...

EQUIPPED (samo ako postoji)
Mjesto | Klub         | Bodovi
1      | Power Crew   | 77.89
...
```

**Pravila:**
- **Top-5 natjecatelja** po klubu se uzimaju u obzir
- **Raw rang** - prikazuje se BEZ dodatnog naslova (podrazumijeva se)
- **Equipped rang** - prikazuje se samo ako postoje Equipped natjecatelji

### 6. Statistika

**Top 5 po kategorijama:**
- Top 5 Muški/Ženski Powerlifting (ukupno)
- Top 5 po divizijama (Kadeti, Juniori, Seniori, Veterani)
- Top 5 Muški/Ženski Potisak s klupe (ukupno)
- Top 5 po divizijama

**Raw i Equipped odvajanje:**
- Raw Top 5 - prikazuje se bez dodatnog naslova
- Equipped Top 5 - prikazuje se s "- EQUIPPED" oznakom (narančasta boja)

**Bojenje medalja:**
- 🥇 1. mjesto - zlatna
- 🥈 2. mjesto - srebrna  
- 🥉 3. mjesto - brončana

## 🔍 Equipment Types

Sistem automatski normalizira equipment tipove:

**Raw:**
- `Sleeves`
- `Raw`
- `Wraps`
- `Straps`

**Equipped:**
- `Single-ply`
- `Multi-ply`
- `Unlimited`
- Sve Division s `-EQ` sufiksom (npr. `Junior-EQ`)

## ⚙️ Kategorije (Divisions)

Sistem prepoznaje sljedeće kategorije:

| Input Naziv | Prepoznato kao | Hrvatski naziv |
|-------------|---------------|----------------|
| Kadet, Sub-Junior, Sub-Juniors | Sub-Junior | Kadeti |
| Junior, Juniors | Junior | Juniori |
| Open, Open-OSI | Open | Seniori |
| Master 1, Master I, Masters 1 | Master I | Veterani 1 |
| Master 2, Master II, Masters 2 | Master II | Veterani 2 |
| Master 3, Master III, Masters 3 | Master III | Veterani 3 |
| Master 4, Master IV, Masters 4 | Master IV | Veterani 4 |

## 🎨 Stilovi u Excel-u

- **Header boja:** Tamno plava (#1F4E79)
- **Granice:** Svijetlo sive (#D9D9D9)
- **Font:** Arial
- **Equipped naslovi:** Narančasta (#C65911)
- **Auto-fit kolone:** Automatski prilagođena širina

## 📋 Primjer Output-a

```
============================================================
SVI KORACI USPJESNO ZAVRSENI!
============================================================
Kreirane datoteke:
   - powerlifting_results_processed.csv (obradeni podaci)
   - rezultati.xlsx (finalni izvjestaj)

Gotovo! Excel izvjestaj je spreman za koristenje.
```

## 🐛 Troubleshooting

### Greška: "Datoteka s klubovima nije pronadjena"
- Provjeri da postoji `input/klubovi.csv`
- Provjeri da je datoteka pravilno nazvana

### Greška: "Natjecatelji bez kluba"
- Dodaj nedostajuće natjecatelje u `input/klubovi.csv`
- Provjeri da se ime i prezime točno poklapaju

### Greška: "Permission denied: rezultati.xlsx"
- Zatvori Excel datoteku ako je otvorena
- Pokreni ponovno

### Encoding problemi (čćšđž)
- Sistem koristi UTF-8 encoding
- Svi CSV fajlovi moraju biti u UTF-8 formatu

## 📝 Napomene

- **GL Points prioritet:** Koristi postojeće Points iz rezultata; kalkulira samo ako nedostaju
- **NS (No Show) zapisi:** Automatski se isključuju iz rezultata
- **Guest natjecatelji:** Isključeni iz club rankings-a
- **Prazna mjesta:** Prikazuju se samo natjecatelji s validnim rezultatima (TotalKg > 0)

## 🔄 Workflow

1. Dobij rezultate natjecanja (`.csv` ili `.opl.csv`)
2. Kreiraj `input/klubovi.csv` s podacima o klubovima
3. Pokreni `python main.py`
4. Otvori `rezultati.xlsx`
5. Gotovo! ✨

## 📚 Dodatne Informacije

- Python 3.8+
- Dependencies: `pandas`, `openpyxl`, `numpy`
- Testiran na Windows 10/11
- Unicode support za hrvatska slova (čćšđž)

---

**Razvio:** HPLS Data Processing Team  
**Verzija:** 2.0 (Standardizirana)  
**Datum:** 2025
