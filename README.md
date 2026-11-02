# CLEANVEST Fund Analytics

Explorative Auswertung von CLEANVEST-Fondsdaten. Can Kuyucu hat die Analysen im Rahmen eines Projekts bei ESG Plus geplant und als Streamlit-Applikation umgesetzt, sodass sie sich auf einen neuen Datenauszug erneut anwenden lassen.

## Daten

Grundlage ist ein Auszug der Finanzprodukt- und Nachhaltigkeitsdaten von CLEANVEST, einschließlich der Nachhaltigkeitsbewertungen der Fonds. Ein Fonds kann mehrere Tranchen haben, etwa je eine in EUR und in einer anderen Währung. Die Bewertung ist über die Tranchen eines Fonds dieselbe. Die Applikation liest deshalb zwei Tabellen:

- `only_unique_funds` — ein Datensatz pro Fonds
- `all_fund_tranches` — ein Datensatz pro Tranche

Enthalten sind unter anderem Fondsname, ISIN, Fondstyp, Kapitalanlagegesellschaft, Domizil, Volumen und Datum sowie der CLEANVEST Score, die Einzelratings der Nachhaltigkeitskriterien und die jeweiligen Anteile im Fonds.

Die Auszüge sind nicht öffentlich verfügbar und liegen diesem Repository nicht bei. Ohne sie lässt sich die Applikation nicht starten. Ein berechtigter Auszug kommt als CSV in den Ordner `input/`. Die Dateinamen tragen das Datum des Abzugs und enthalten einen der beiden Namensteile oben. Volumina in Fremdwährung werden mit Kursen der Europäischen Zentralbank (Frankfurter-API) nach EUR umgerechnet. Abgefragt werden dafür nur Währung und Datum.

## Auswertungen

- **Überblick.** Anzahl der Fonds und Tranchen, Aufteilung nach Fondstyp, Einordnung nach Score und Gesamtvolumen.
- **Score-Verteilung.** Wie sich die bewerteten Fonds auf die Score-Bereiche verteilen.
- **Kriterien.** Je Nachhaltigkeitskriterium die Verteilung der Ratings und der Anteile im Fonds.
- **Weitere Analysen.** Zusätzliche Schnitte über den Auszug, etwa die Fonds mit den höchsten und niedrigsten Scores.
- **Kapitalanlagegesellschaft.** Anzahl, durchschnittlicher Score und Volumen je Gesellschaft.
- **Domizil.** Dieselben Kennzahlen je Land, mit einer Karte.
- **Fondsvergleich und Fondssuche.** Einzelne Fonds gegenüberstellen oder über die ISIN öffnen.
- **Monatsvergleich.** Veränderung der Scores gegenüber einem früheren Auszug.
- **Assistent.** Lokaler Chat über Ollama oder ein lokales Sprachmodell (Qwen über transformers). Er beantwortet Fragen zu den Eckdaten des geladenen Auszugs: Fondsanzahl, durchschnittlicher Score, die fünf höchsten und fünf niedrigsten Scores, die stärksten und schwächsten Kriterien. Die Auswertung läuft auf dem eigenen Rechner.

## Start

Python 3.13 und ein CLEANVEST-Auszug im Ordner `input/`.

```bash
poetry install
poetry run streamlit run main.py
```

## Abbildungen

<!--
![Überblick](screenshots/01-ueberblick.png)
![Score-Verteilung](screenshots/02-score-verteilung.png)
![Kriterien](screenshots/03-kriterien.png)
![KAG](screenshots/04-kag.png)
![Domizil](screenshots/05-domizil.png)
![Monatsvergleich](screenshots/06-monatsvergleich.png)
![KI-Assistent](screenshots/07-ki-assistent.png)
-->
