# Voetbalkamp Schema Generator

Streamlit app voor het genereren van een spelschema voor een voetbalkamp (of ander sportkamp). Het schema is gebaseerd op een **Latin Square**, waardoor elk team elk spel precies 1x speelt en in elke ronde precies 1x aan de beurt is.

## Features

- **Latin Square algoritme** met geoptimaliseerde tegenstander-diversiteit
- **Duo-spellen** (2 teams tegen elkaar) en **solo-spellen** (1 team per spel)
- Dynamische tijdslots op basis van starttijd, speelduur en wisseltijd
- Aanpasbare team- en spelnamen
- Spelbeschrijvingen per spel
- Seed-parameter voor reproduceerbare maar gevarieerde schema's
- **Excel-export** met meerdere tabbladen:
  - Overzichtsschema
  - Per spel (met scorekolom)
  - Per team (met spel en tegenstander)
- Verificatie van het Latin Square (met som-controle per rij en kolom)

## Vereisten

```
pip install streamlit pandas openpyxl
```

## Gebruik

```bash
streamlit run latin_square_app.py
```

De app opent in je browser. Configureer in de sidebar:

1. **Aantal teams** en teamnamen
2. **Aantal duo- en solo-spellen** (duo x 2 + solo = aantal teams)
3. **Tijdsinstellingen** (starttijd, speelduur, wisseltijd)
4. **Variatie (seed)** voor een ander schema
5. Klik op **Genereer schema**

## Hoe werkt het?

Het schema gebruikt een [Latin Square](https://nl.wikipedia.org/wiki/Latijns_vierkant): een n x n rooster waarin elk getal (team) precies 1x per rij en 1x per kolom voorkomt. Kolommen worden gegroepeerd in paren voor duo-spellen en singles voor solo-spellen. Een optimalisatiestap zorgt ervoor dat teams zoveel mogelijk verschillende tegenstanders treffen.
