# Handleiding voor de leraar — Les 14: Uitnodigen met de agenda

<!-- Beginvragen (stap 1 van CONTEXT.md), beantwoord door Jonas op 16-11-2026 ("Maak de lessen van deze week die ik nog
     nodig heb"; telkens het voorstel gekozen):
     klassen 3ORLO, 4ORLOa, 4ORLOb · les 14 van het jaarplan (module 3), uur 1 van W11: 4ORLOb di 17-11 (8ste uur),
     3ORLO di 17-11 (9de uur), 4ORLOa wo 18-11 (4de uur) · vorige les: les 13 (Phishing of echt?), gegeven in alle drie
     de klassen → lesstart = ophalen uit les 13 · lokaal 18, Windows 11 · in Google Agenda (schoolaccount) een
     NovaDepot-vergadering maken met lokaal, tijd, beschrijving en genodigden (Jonas + een medeleerling) en de uitnodiging
     versturen · werkdocument CO4_Uitnodigen-met-de-agenda (reeks Communicatie) met een screenshot en korte vragen ·
     20 punten, deadline vr 20-11-2026 20.00 · context NovaDepot. 
     Getest op 16-11-2026 met demoleerling3 (Jonas: "Test het?"): Google Agenda werkt voor leerlingen (afspraak maken,
     gasten uitnodigen, Sturen); Gmail niet. Geen ruimtes: het lokaal komt bij Locatie. Dus: de uitnodiging in de echte
     Google Agenda, met Jonas en de duopartner als gasten; de e-mail (E1) in een nagebootst venster, zoals les 12. -->

**Vak:** Toegepaste Informatica
**Doelgroep:** 3ORLO, 4ORLOa en 4ORLOb — 2de graad Organisatie en logistiek, arbeidsmarktgerichte finaliteit
**Lesduur:** 1 × 50 minuten: 10 minuten instructie + 40 minuten keuzewerktijd
**Context:** NovaDepot, *het overleg over de infosessie* — module 3 van het jaarplan, les 14 (reeks *Communicatie 4*)
**Toestel:** **lokaal 18, Windows 11** — alles in Chrome, in de **echte Google Agenda** van het schoolaccount (geen Gmail)
**Lesdag:** 4ORLOb di 17 november (8ste uur) · 3ORLO di 17 november (9de uur) · 4ORLOa wo 18 november (4de uur)
**Werkvorm:** individueel een afspraak maken; de **duopartner** is een gast en antwoordt op de uitnodiging
**Kernleerplandoel:** `BK2_02.10` (`BK2_02.10.02`, `BK2_02.10.03`) en `BV2_04.01` (toepassen) · ondersteunend `BK2_02.04`
**Deadline:** vrijdag 20 november 2026, 20.00 uur · 20 punten
**Pages:** https://jonasdaltongent.github.io/Agenda-NovaDepot/ · dia's: …/presentatie.html · repository: `jonasdaltongent/Agenda-NovaDepot`

Lotte Maes (personeelsdienst) vraagt Tom Peeters een overleg vast te leggen voor de infosessie van de jobstudenten (dinsdag 1 december 2026, 10.00–10.45):
*"Nodig mij en Sara uit. Ik leid het overleg."* De leerling kiest een zaal, maakt in Google Agenda de afspraak (titel, datum, uur, locatie, gasten, beschrijving),
klikt op **Sturen**, antwoordt **Ja** op de uitnodiging van de duopartner en typt over wat hij invulde in het werkdocument. Gasten: **jij** (Lotte) en de **duopartner**
(Sara). Codewoord van de zelftest: **OVERLEG**. Opbouw en bronnen: `lesvoorbereiding.md`; de uitleg komt uit `_vakinhoud/digitale-communicatie.md`.

---

## 1. Inhoud van het pakket

```text
W11 - Les 14 - ORLO - Uitnodigen met de agenda/
├── index.html                   # de lespagina: route · één stap · checklist (7 stappen + Extra)
├── presentatie.html             # 10 dia's: 1 Zo start je · 2 lesstart (les 13) · 3 de vraag van Lotte · 4 zes dingen · 5–6 demo · 7 controle · 8 stap 1 samen · 9 kloktijden · 10 exitvraag
├── css/style.css, css/slides.css
├── js/script.js, js/slides.js
├── assets/                      # logo's, lettertypes, mascotte, het filmpje Inleveren (geen schermafbeeldingen)
├── werkdocument/
│   ├── CO4_Uitnodigen-met-de-agenda.docx   # het werkdocument dat de leerling INLEVERT (kopie per leerling)
│   ├── maak_werkdocumenten.py              # maakt het opnieuw (python-docx, Pillow)
│   └── huisstijl.py, go-dalton-gent.png
├── lesvoorbereiding.md
├── dalton-lesfiche.html         # openen, Kopieer de fiche, plakken in je planner
├── classroom.json               # de opdracht voor _tools/zet_opdracht_klaar.py
├── lesdoelen.json               # leerplandoelen voor je jaaroverzicht
└── README.md                    # deze handleiding
```

Geen zip en geen bronbestanden.

### 1b. Hoe de lespagina werkt

Zoals les 11–13: links de **route** (*Het overleg voorbereiden*: stap 1–3 · *Het overleg vastleggen*: stap 4–7), midden **één stap** met *Toon stapjes* per
handeling (de les begint in *Kort*, zonder knop), rechts de **checklist** met 16 taken. Geen *Hulp nodig?*, geen AI-hulp.

- **Startpagina** met vier beelden (de opdracht in Classroom → deze lespagina → Google Agenda → je werkdocument CO4), *Dit maak je vandaag* (het skelet van de
  afspraak: verklapt niets), *Na deze les kan je* en *Vast?*.
- **Stap 1** Klaarzetten: werkdocument openen, het kader lezen, **Google Agenda** in een tweede tabblad (link `calendar.google.com`; de twee eerste vensters: **OK**).
- **Stap 2** de zes onderdelen (deel 1): infokader met de zes onderdelen, zonder de letters.
- **Stap 3** de zaal kiezen (deel 2a); de feiten (drie lokalen) staan in het kader van het werkdocument.
- **Stap 4** de afspraak maken, met tekeningen van Google Agenda in *Toon stapjes* en een infokader met de **bouwstenen van de beschrijving** (de feiten erbij).
- **Stap 5** schermafbeelding (2c), **Opslaan › Sturen**, vraag 2d, overtypen (2b).
- **Stap 6** antwoorden: de uitnodiging van de duopartner, **Ja**, de twee vragen van deel 3.
- **Stap 7** *Inleveren*: zelftest van drie vragen, de groene regel met het codewoord, dan *Afwerken en inleveren* (codewoord, exitvraag, controletabel, inleveren, het filmpje).
- **Extra** (telt niet mee): *Tijdstip zoeken*; opnieuw inleveren met **Ongedaan maken**.
- `localStorage` bewaart vinkjes en de laatste stap (voorvoegsel `novadepot_agenda_v1_`), met een wisknop. Geen login, geen tracking, geen externe scripts.

---

## 2. Klaarzetten

### Stap 1 — Publiceren via GitHub Pages

Repository `jonasdaltongent/Agenda-NovaDepot`, Pages op `main`, map `/ (root)`; adres hierboven. Dat adres staat op dia 9, in `classroom.json` en in `lesdoelen.json`.

### Stap 2 — De opdracht in Classroom, met de koppeling

`classroom.json`: **Communicatie 4 — Uitnodigen met de agenda**, onderwerp *Digitale competenties*, 3ORLO, 4ORLOa en 4ORLOb, 20 punten, deadline vrijdag 20 november 2026,
20.00 uur. Bijlagen: de link naar de lespagina en `CO4_Uitnodigen-met-de-agenda` (**Een kopie maken voor elke leerling**). Geen zip.

**Stand:** {{CLASSROOM}}

Instructietekst:

```text
1. Open de lespagina (link) en je werkdocument CO4_Uitnodigen-met-de-agenda.
2. Je maakt de afspraak in Google Agenda en nodigt je duopartner en je leraar uit. Het adres van je leraar staat op het bord.
3. Klik op Inleveren, ten laatste vrijdag 20 november 2026 om 20.00 uur.
```

Inplannen: een half uur vóór elke les: di 17-11 om 14.05 (4ORLOb) en 14.55 (3ORLO), wo 18-11 om 10.45 (4ORLOa).

### Stap 3 — Met de hand, als de koppeling niet werkt

1. Upload `werkdocument/CO4_Uitnodigen-met-de-agenda.docx` **in Drive zelf** (**Nieuw** › **Bestand uploaden**, met *Uploads converteren* aan) en voeg het toe met **Bijvoegen** › **Drive** › **Een kopie maken voor elke leerling**.
2. Voeg de lespagina toe met **Link**.

### Afvinklijst vóór de les

- [ ] De lespagina staat online en het adres op dia 9 klopt.
- [ ] De opdracht staat in de drie klassen, met twee bijlagen (de link en het werkdocument).
- [ ] **Schrijf het adres van de leraar op het bord.** Het staat nergens anders (niet op de lespagina, de dia's, in het werkdocument of in `classroom.json`).
- [ ] De duo's staan vast (naast elkaar; een oneven aantal: één trio).
- [ ] Les 13 is in die klas gegeven (dia 2 haalt er een vraag uit op).
- [ ] Je bent **aangemeld in Google Agenda** met je schoolaccount, en de demo staat klaar (*Overleg nieuwe rekken magazijn*, do 3 december 2026, 14.00–14.30, Magazijnkantoor). Voor het antwoorden als gast heb je een tweede venster of account nodig.
- [ ] Een testleerling krijgt een eigen kopie van het werkdocument, met de eigen naam in de titel.
- [ ] Kijk in de app *Na de les* of er over les 11–13 iets bij te sturen is.
- [ ] De Dalton-lesfiche staat in je planner (open `dalton-lesfiche.html`, klik op **Kopieer de fiche**, plak).
- [ ] `presentatie.html` opent op de beamer; `N` toont je notities, `F` is volledig scherm.

### 2b. Nagelezen klikpaden

| Handeling | Klikpad / naam | Bron |
|---|---|---|
| Werkdocument openen | in de opdracht rechts bij **Jouw werk** op de naam klikken | `_afspraken/klikpaden.md` |
| Een afspraak maken | links **Maken** › **Afspraak** › in het kleine venster **Meer opties**; daar **Titel toevoegen**, de datum en het uur, **Locatie toevoegen**, **Beschrijving toevoegen**, rechts **Gasten toevoegen**, **Opslaan** | `_afspraken/klikpaden.md` (rijen *Agenda (leerling)…* en *Agenda: Meer opties*; demoleerling3, 16-11-2026); [Agenda 72143](https://support.google.com/calendar/answer/72143?hl=nl) |
| De uitnodiging versturen | **Opslaan** › venster *Wil je uitnodigingsmails sturen naar Google Agenda-gasten?* met **Terug naar bewerken** · **Niet sturen** · **Sturen** | `_afspraken/klikpaden.md` (*Agenda: de uitnodiging versturen*; demoleerling3, 16-11-2026) |
| Een afspraak weer bewerken (na *Niet sturen*) | op de afspraak klikken › **Afspraak bewerken** › **Opslaan** › **Sturen** | [Agenda 37141](https://support.google.com/calendar/answer/37141?hl=nl) (lees: *Klik op een afspraak › Afspraak bewerken*); het opnieuw sturen is **niet getest** |
| Antwoorden op een uitnodiging | in Google Agenda op de afspraak klikken › **Ja** · **Nee** · **Misschien** | [Agenda 37135](https://support.google.com/calendar/answer/37135?hl=nl), *Reageren op een uitnodiging in Google Agenda* (gelezen 16-11-2026); **niet getest** met een leerlingaccount |
| Wie komt? | bij je eigen afspraak, bij **Gasten** | [Agenda 37141](https://support.google.com/calendar/answer/37141?hl=nl) (*Als je wilt zien wie op je afspraak heeft gereageerd, klik je op de afspraak*); hoe het eruitziet is **niet getest** |
| Tijdstip zoeken (Extra) | in *Meer opties* het tabblad **Tijdstip zoeken** | `_afspraken/klikpaden.md` (*Agenda: Meer opties*) |
| Schermafbeelding | **Windows + Shift + S**, plakken met **Ctrl + V** | `_afspraken/toestel.md` |
| Inleveren | **Inleveren** › venster **Je werk inleveren?** › **Inleveren** › status **Ingeleverd** | [Classroom 6020285](https://support.google.com/edu/classroom/answer/6020285?hl=nl); het venster gezien met demoleerling3 (15-10-2026) |
| Opnieuw inleveren (Extra) | **Ongedaan maken**, daarna weer **Inleveren** | zelfde bron en proef |

---

## 3. Het verloop van de les

| Fase | Tijd | Wat |
|---|---|---|
| **Instructie** | **10'** | Dia 1 terwijl de computers opstarten (1'). Dia 2 lesstart (2'): een uitnodiging van een onbekende (les 13), met vingers. Dia 3 (1'): de vraag van Lotte en het eindproduct. Dia 4 (1'): de zes dingen in een uitnodiging (plaats-datum-tijd en duur voor). Dia 5–6 (3'): **demo in Google Agenda, hardop denkend**: *Maken › Afspraak › Meer opties*, titel, datum, uur, *Locatie toevoegen*, *Gasten toevoegen*, beschrijving, *Opslaan › Sturen*, en als gast **Ja**. Dia 7 controle (1'): *Opslaan › Sturen*. |
| **Samen** | **3'** | Dia 8: de duo's zitten naast elkaar; iedereen doet stap 1; duim op. |
| **Keuzewerktijd** | **tot het einde** | Dia 9 blijft staan, met **kloktijden** (`_tools/zet_lesmomenten.py`); stap 2 tot 7, inleveren inbegrepen. Dia 10 in de laatste minuut. |

Stap 1–7: 3 + 4 + 2 + 12 + 7 + 7 + 5 = **40 minuten**.

**Eerste rondgang:** heeft **elke** leerling het eigen werkdocument open en Google Agenda in een tweede tabblad? **Tweede rondgang, rond stap 5:** staat de afspraak op
**dinsdag 1 december 2026** (niet vandaag)? Zit de zaal bij *Locatie*? Klikte iedereen op **Sturen**? Kijk zelf in je agenda of de uitnodigingen binnenkomen. Zeg aan het einde
mondeling dat wie niet klaar is, toch inlevert. Minimumroute: deel 1, 2a, 2b (de afspraak verstuurd en overgetypt) en deel 3 vraag 2. **Duopartner afwezig:** nodig alleen
jou uit; in stap 6 antwoordt de leerling op de uitnodiging van een ander duo (jij wijst aan).

**Je krijgt veel uitnodigingen** (± 60 voor dinsdag 1 december 2026 in je eigen agenda): dat is de bedoeling, zo zie je wie verstuurde. Ze mogen er daarna weer uit.

---

## 4. Verbetersleutel

- **Codewoord** bovenaan *Deel 4 — Klaar?*: **OVERLEG**. Ontbreekt het, dan is de zelftest niet (juist) gemaakt: −2 (een tikfout telt als juist).

De leerling levert alleen het werkdocument in (20 punten: 4 · 13 · 2 · 1). De afspraak zelf zie je ook in je eigen agenda (di 1-12-2026).

| Onderdeel | Verwacht |
|---|---|
| Deel 1 (4) | wie het overleg leidt **B** · de reden en het doel **E** · de agenda **A** · wat je doet als je niet kan komen **D** (1 per letter; **C** en **F** staan al ingevuld) |
| Deel 2a (2) | **vergaderzaal 1** (1) + een reden uit het kader: plaats voor drie, scherm, rustig, naast het onthaal (1). *Omdat het een vergaderzaal is* = ½ voor de reden. Niet goed: de kantine (druk, te groot), het kantoor van Lotte (te klein) |
| Deel 2b (9) | **titel** met *overleg* en *infosessie* (1) · **dinsdag 1 december 2026** (1) · **10.00 tot 10.45** (1) · zaal bij **Locatie** (1) · **leraar en duopartner** als gast (1) · **beschrijving**: wie leidt (Lotte Maes) (1), het doel (1), de vier agendapunten (1; ½ bij twee of drie), wat je doet als je niet kan komen (*antwoord met Nee*) (1) |
| Deel 2c, 2d (2) | een **schermafbeelding** van de eigen afspraak (1) · de drie knoppen *Terug naar bewerken · Niet sturen · Sturen* en dat de leerling op *Sturen* klikte (1) |
| Deel 3 (2) | vraag 1: wie leidt volgens de uitnodiging van de duopartner (**Lotte Maes**; staat er niets, dan *dat staat er niet*) (1) · vraag 2: wie **Ja** zei en wie nog niet antwoordde, uit het eigen scherm (geen vast antwoord, wel echt) (1) |
| Deel 4 (1) | de exitvraag: Sara antwoordt **Nee**, en bij jouw afspraak staat dat ze niet komt (of je krijgt een melding) (1) |

**Hoe de punten verdeeld worden** (de controletabel onderaan het werkdocument heeft dezelfde criteria, samen 20; § 22 van de lesvoorbereiding):
4 · 2 · 3 · 2 · 4 · 2 · 2 · 1. Een datum zonder jaartal of met de verkeerde dag (*dinsdag 2 december*) = 0 voor de datum; een uur *10–11* = 0 voor het uur (de duur is 45
minuten). Een beschrijving die de feiten letterlijk uit het kader overneemt, is goed (*Mag je overnemen*). **Wat niet telt:** spelling, hoofdletters in de titel, of de
Meet-link er nog bij staat, de volgorde van de punten in de beschrijving.

### Essentiële fouten — geef hier altijd feedback op

- Opgeslagen met **Niet sturen**: de gasten weten van niets (2d).
- Geen gasten, of een e-mailadres van buiten de school.
- De zaal in de **titel** in plaats van bij *Locatie*; de agendapunten in de titel.
- Het verkeerde jaar of de verkeerde dag (de vergadering valt in het verleden of op een andere dag dan Lotte vroeg).

Verbeteren en opnieuw inleveren mag. Les 15 (E1) herhaalt de agenda-uitnodiging.

---

## 5. Schermafbeeldingen (optioneel)

Geen: alle beelden zijn tekeningen in HTML en CSS (de opdracht in Classroom, Google Agenda, de tabellen van het werkdocument), behalve het filmpje *Inleveren* uit het
lespakket. De tekening van deel 1 in het werkdocument (de afspraak met de letters A–F) is met Pillow getekend. Zie `assets/screenshots/LEESMIJ.md`.

---

## 6. Het materiaal opnieuw maken

```bash
python3 "werkdocument/maak_werkdocumenten.py"
```

Maakt `CO4_Uitnodigen-met-de-agenda.docx` opnieuw. Vereist `python-docx` en `Pillow`. Het kader (*Wat Lotte vraagt*), de onderdelen met hun letters en de criteria staan bovenaan het
script; wijzig je de vakken, de letters of de punten, doe het ook in § 14 en § 22 van `lesvoorbereiding.md` en op dia 4.

---

## 7. Leerplandoelen in je jaaroverzicht

`lesdoelen.json`: `BK2_02.10`, `BK2_02.10.02`, `BK2_02.10.03` en `BV2_04.01` (kern), `BK2_02.04` (ondersteunend). De repository krijgt de `pre-push` hook uit `_sjablonen/pre-push`.

---

## 8. Wat nog moet blijken in de klas

- **Alles in de echte Google Agenda.** Getest met demoleerling3 (16-11-2026): een afspraak maken, gasten toevoegen en **Sturen** werken; Gmail niet. **Niet getest:**
  - of een leerling de **duopartner** kan uitnodigen door de **naam** te typen (de lijst met suggesties) of dat het volledige adres nodig is. De lespagina zegt: *typ de naam en kies uit de lijst*. Lukt dat niet, zeg dan het adres hardop of schrijf het formaat op het bord;
  - hoe een gast de uitnodiging **ontvangt** (staat ze meteen in de week van 1 december, of moet de gast eerst op een melding klikken?) en of **Ja · Nee · Misschien** bij de afspraak staat (de help zegt van wel, zonder Gmail);
  - hoe de organisator bij **Gasten** ziet wie **Ja** zei (deel 3, vraag 2 vraagt het); kijk de eerste rondgang na;
  - of **opnieuw sturen** na *Niet sturen* werkt zoals de rode waarschuwing zegt (*Afspraak bewerken › Opslaan › Sturen*);
  - of **Tijdstip zoeken** (Extra) de agenda van de duopartner toont: alleen als die zijn agenda deelt of in dezelfde school zit.
- **De adressen van de gasten.** Het adres van de leraar staat op het bord; het adres van de duopartner zoekt de leerling zelf. Kijk of dat vlot gaat.
- **Haalbaarheid in 40 minuten.** De afspraak (stap 4) is 12 minuten voor titel, datum, uur, locatie, gasten en beschrijving; stap 5 en 6 samen 14. De minimumroute is de uitweg.
- **De datum.** Google Agenda zet *vandaag* en het uur van nu: veel leerlingen vergeten de datum aan te passen (dinsdag 1 december 2026, niet 2025). Daarop staat de tip bij de datum en de rode waarschuwing bij *Niet sturen*.
- **Jouw agenda.** Reken op ± 60 uitnodigingen; kijk of dat de agenda onbruikbaar maakt (Google bundelt ze niet). Wis ze daarna met de hand.
- **Dia 4 en 6** zijn nieuwe soorten tekeningen (de afspraak met zes onderdelen; het venster *Sturen* en *Ja*); de notities lezen ze de eerste keer hardop voor.
- De **werkdocument-tekening** van deel 1 is niet in Google Documenten bekeken (geen pdf-export vooraf): kijk na in de kopie van een testleerling of de afbeelding en de tabel op één bladzijde staan.
- De lespagina en dia's zijn gemeten met scripts (geen horizontaal scrollen op 683, 1024 en 1440 px; elke dia binnen de rand; zelftest, codewoord, checklist) maar niet door een
  mens op een Windows-toestel: kijk de tekeningen na.
