# WedstrijdCalculator - Handleiding

Lokale Python-tool voor zeilwedstrijdscoring via het **Portsmouth Yardstick (PY)** systeem.

## Vereisten

Python 3.10 of hoger is vereist. Installeer de benodigde packages eenmalig:

```
pip install pandas openpyxl
```

## Bestanden

| Bestand | Beschrijving |
|---|---|
| `wedstrijd_calculator.py` | Hoofdscript |
| `voorbeeld_invoer.csv` | Voorbeeldinvoer met 6 deelnemers en 5 reeksen |
| `extra_boottypes.csv` | Optioneel: voeg eigen boottypes/PY-waarden toe |
| `wedstrijd_resultaten.xlsx` | Wordt aangemaakt na uitvoering |

## Gebruik

### Interactief menu (aanbevolen)

```
python wedstrijd_calculator.py
```

Volg het menu om een bestand te laden of demo-data te starten.

### Directe verwerking via de commandoregel

```
python wedstrijd_calculator.py --invoer voorbeeld_invoer.csv
python wedstrijd_calculator.py --invoer voorbeeld_invoer.xlsx
python wedstrijd_calculator.py --demo
```

### Opties

| Optie | Beschrijving |
|---|---|
| `--invoer` / `-i` | Pad naar CSV- of Excel-invoerbestand |
| `--uitvoer` / `-o` | Naam van het Excel-uitvoerbestand (standaard: `wedstrijd_resultaten.xlsx`) |
| `--schrap` | Schrapresultaat toepassen - slechtste reeks niet meetellen (standaard aan) |
| `--geen-schrap` | Geen schrapresultaat toepassen |
| `--extra-py` | CSV-bestand met extra boottype/PY-waarden |
| `--demo` | Voer uit met ingebouwde voorbeelddata |

### Voorbeeld met extra PY-bestand

```
python wedstrijd_calculator.py --invoer mijn_wedstrijd.csv --extra-py extra_boottypes.csv
```

## Invoerformaat (CSV of Excel)

Verplichte kolommen (hoofdletterongevoelig):

| Kolom | Beschrijving |
|---|---|
| `naam` | Naam van de deelnemer |
| `boottype` | Type boot (moet overeenkomen met een PY-waarde) |
| `reeks` | Reeksnummer (1, 2, 3, ...) |
| `minuten` | Verlopen tijd: minuten |
| `seconden` | Verlopen tijd: seconden |

Optionele kolom:

| Kolom | Beschrijving |
|---|---|
| `py` | Manuele PY-waarde (overschrijft de interne tabel) |

## PY-waarden aanpassen

Open `wedstrijd_calculator.py` en zoek de sectie `BOAT_PY` bovenaan het bestand:

```python
BOAT_PY: dict[str, float] = {
    "Laser":        1100,
    "Laser Radial": 1147,
    "RS Aero 7":    1065,
    "RS Neo":       1176,
    # Voeg hier toe:
    # "MijnBoot":   1050,
}
```

Of gebruik het bestand `extra_boottypes.csv` om boottypes toe te voegen zonder de code te wijzigen.

## PY-formule

```
gecorrigeerde_tijd = verlopen_tijd_seconden * 1000 / PY
```

Een lagere gecorrigeerde tijd is beter. De rang per reeks wordt bepaald op basis van deze gecorrigeerde tijd.

## Puntensysteem

- 1e plaats = 1 punt
- 2e plaats = 2 punten
- enzovoort

Laagste totaalpunten wint. Met schrapresultaat wordt de slechtste reeks per deelnemer niet meegeteld.

## Excel-uitvoer

Het resultaatbestand bevat drie tabbladen:

- **Klassement** - Totaalklassement met punten per reeks, som, slechtste en eindstand
- **Detail per reeks** - Per reeks: naam, boottype, PY, tijd, gecorrigeerde tijd en rang
- **Posities per reeks** - Compacte tabel met positie per reeks per deelnemer
