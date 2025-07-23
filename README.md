# Bjelovar Record Breakers - Data Processing System 🏋️‍♂️🇭🇷

Kompletni sustav za obradu podataka o natjecanju u powerliftingu s automatskim generiranjem Excel izvještaja.

## 📋 Pregled

Ovaj sustav obrađuje podatke iz dva CSV fajla i generira profesionalni Excel izvještaj s:
- Rezultatima natjecatelja (po kategorijama i težinskim klasama)
- Rang listama klubova
- Statistikama natjecanja
- IPF GL bodovima za sve natjecatelje

## 🚀 Kako koristiti

### Jednostavno pokretanje (preporučeno)
```bash
python main.py
```

**To je sve!** Skripta će automatski:
1. Provjeriti postojanje ulaznih datoteka
2. Obraditi osnovne podatke
3. Generirati rezultate klubova
4. Kreirati rang liste klubova  
5. Stvoriti konačni Excel izvještaj

### Ulazne datoteke
Stavite sljedeće datoteke u `bjelovar/` mapu:
- `3-bjelovar-record-breakers.opl (1).csv` - rezultati natjecanja
- `Bjelovar-record-breakers-finalne-nominacije-2-1-3-1-1-1.csv` - nominacije klubova

### Izlazne datoteke

#### Glavni izvještaj
- **`bjelovar/Bjelovar_Record_Breakers_Rezultati.xlsx`** - konačni Excel izvještaj

#### Međusobni CSV-ovi (zadržani)
- `powerlifting_results_processed.csv` - obrađeni osnovni podaci
- `Male_Powerlifting.csv` / `Female_Powerlifting.csv` - rezultati klubova (klasično)
- `Male_Bench_Only.csv` / `Female_Bench_Only.csv` - rezultati klubova (potisak)
- `*_Ranking.csv` - rang liste klubova za sve kategorije

## 📊 Excel izvještaj sadrži

### Stranice rezultata
- **Muški Powerlifting** - muški natjecatelji (čučanj, potisak, mrtvo dizanje)
- **Ženski Powerlifting** - ženske natjecateljice (čučanj, potisak, mrtvo dizanje)  
- **Muški Potisak s klupe** - muški natjecatelji (samo potisak)
- **Ženski Potisak s klupe** - ženske natjecateljice (samo potisak)

### Dodatne stranice
- **Rang Klubova** - rang liste klubova po kategorijama
- **Statistika** - opća statistika natjecanja i top 5 performanse

### Značajke formatiranja
- 🥇🥈🥉 Medalje bojanje (zlato, srebro, bronca)
- 📊 Profesionalna color shema
- 🇭🇷 Potpuna hrvatska lokalizacija
- 📋 Vizualno odvajanje kategorija
- 📈 Auto-fit kolumne za optimalno čitanje

## 🛠️ Tehnički detalji

### Kategorije/Uzrasti
- **Kadeti** (Sub-Junior)
- **Juniori** (Junior)  
- **Seniori** (Open)
- **Veterani 1-4** (Master I-IV)
- **Gost** (Guest) - neoficijalni rezultati

### IPF GL bodovi
Koristi službene IPF GL koeficijente (2020) za:
- Muška/ženska klasična powerlifting
- Muški/ženski klasični potisak s klupe

### Sortiranje
- Kategorije: Kadeti → Juniori → Seniori → Veterani 1-4
- Težinske klase: od najlakših prema najtežim (uključuje superheavy +)
- Plasmani: numerički (1, 2, 3, ..., 10, 11)

## 📁 Struktura projekta
```
obradarezultata/
├── main.py                          # 🎯 GLAVNI SCRIPT
├── bjelovar/
│   ├── 3-bjelovar-record-breakers.opl (1).csv
│   └── Bjelovar-record-breakers-finalne-nominacije-2-1-3-1-1-1.csv  
├── process_powerlifting_data.py     # Korak 1: Osnovna obrada
├── generate_club_results.py         # Korak 2: Rezultati klubova
├── generate_club_rankings.py        # Korak 3: Rang liste
├── create_excel_report.py           # Korak 4: Excel izvještaj
└── README.md                        # Dokumentacija
```

## 🔧 Instalacija

### Preuzimanje
```bash
git clone https://github.com/YOUR_USERNAME/bjelovar-record-breakers.git
cd bjelovar-record-breakers
```

### Zavisnosti
```bash
pip install -r requirements.txt
```

**Ili ručno:**
```bash
pip install pandas openpyxl
```

## ✅ Testiranje
Za testiranje sustava s postojećim podacima:
```bash
python main.py
```

Skripta će prikazati napredak kroz sve korake i obavijestiti o uspješnom završetku.

---
*Sustav razvijen za Bjelovar Record Breakers natjecanje u powerliftingu* 🏋️‍♂️ 