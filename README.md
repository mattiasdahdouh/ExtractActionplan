# Extrahera Handlingsplaner

Ett lokalt Python-program med grafiskt gränssnitt (GUI) som extraherar data från Word-dokument och exporterar till Excel och CSV.

## Vad programmet gör

Programmet läser `.docx`-filer från en vald mapp och extraherar tabellen på första sidan. Tabellen förväntas ha **7 rader** där kolumn A innehåller kategorin och kolumn B innehåller datan. Resultatet sparas som en sammanställd Excel- och CSV-fil.

**Exempel på fält som extraheras:**
| Kategori | Exempel |
|---|---|
| Butik | Nära Korvgubben |
| Nuvarande handlare | Bertil Bertilsson |
| Kandidat | Bertil Bertilsson Junior |
| Kontakt ICA/Affärspartner | Anna Andersson |
| Tidpunkt för generationskifte | Inom 5 år |
| Grundkravprofil | 3 |
| Omsättning | 50 Msek |

## Krav

- **Python 3** — ladda ner från [python.org](https://www.python.org/downloads/). Kryssa i **"Add Python to PATH"** under installationen.
- **Paket** — installeras via terminal:

```
pip install python-docx openpyxl
```

`tkinter` och `csv` ingår i Python och behöver inte installeras separat.

## Starta programmet

**Alternativ 1** — dubbelklicka på `starta.bat`

**Alternativ 2** — kör via terminal:

```
py extrahera_handlingsplan.py
```

## Användning

1. Klicka på **Bläddra** bredvid _Källmapp_ och välj mappen som innehåller dina Word-dokument
2. Klicka på **Bläddra** bredvid _Utdatamapp_ och välj var de exporterade filerna ska sparas
3. Klicka på **Starta** — statusraden och progressbaren visar hur långt programmet har kommit
4. När det är klart visas en dialog med sökvägen till de skapade filerna

## Utdata

Två filer skapas i den valda utdatamappen:

| Fil                                    | Beskrivning                                                |
| -------------------------------------- | ---------------------------------------------------------- |
| `handlingsplaner_sammanstallning.xlsx` | Formaterad Excel-fil med rubrikrad och cellformatering     |
| `handlingsplaner_sammanstallning.csv`  | Semikolonseparerad CSV-fil, UTF-8 (öppnas korrekt i Excel) |

Om en fil med samma namn redan finns i utdatamappen skrivs den **inte** över. Istället skapas en ny fil med suffixet `(1)`, `(2)` osv.

## Filstruktur

```
ExtractActionplan/
├── extrahera_handlingsplan.py   # Huvudprogrammet
├── starta.bat                   # Snabbstart för Windows
└── README.md
```
