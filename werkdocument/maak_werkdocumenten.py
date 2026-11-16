#!/usr/bin/env python3
"""
maak_werkdocumenten.py
Genereert het werkdocument voor les 14 "Uitnodigen met de agenda" (NovaDepot, ORLO, Communicatie 4).

  CO4_Uitnodigen-met-de-agenda.docx   het werkdocument dat de leerling INLEVERT (een kopie per leerling)

Geen bronbestanden en geen zip. De leerling maakt de afspraak in de ECHTE Google Agenda (schoolaccount) en typt in dit
document over wat hij invulde, plakt een schermafbeelding en beantwoordt een paar vragen. De tekening van deel 1 (een
afspraak met zes gemerkte stukjes) wordt hier met Pillow getekend.

Gebruik:  python3 maak_werkdocumenten.py
Vereist:  python-docx, Pillow (huisstijl.py en go-dalton-gent.png staan ernaast)

LET OP bij aanpassen:
 - Alle namen en lokalen zijn verzonnen. Het e-mailadres van de leraar staat NIET in dit document (het staat op het bord).
 - Deel 1: de letters staan in een andere volgorde dan de onderdelen (zodat A-B-C-D-E-F niet vanzelf juist is):
   plaats-datum-tijd = C (voorgedaan) · duur = F (voorgedaan) · wie leidt = B · reden en doel = E · agenda = A ·
   afwezig melden = D. Wijzig je de tekening, wijzig dan ook ONDERDELEN, § 14 en § 22 van de lesvoorbereiding en dia 4.
 - Elk criterium van § 22 heeft zijn eigen vak (zodat de verbetertool per criterium punten kan geven). De punten staan per
   deel (4 · 13 · 2 · 1) en samen in de controletabel (CRITERIA, samen 20). Verander je ze, doe het op beide plaatsen.
 - Het codewoord (OVERLEG) staat bovenaan deel 4. De lespagina toont het pas als de zelftest juist is.
"""
import io
import os

from PIL import Image, ImageDraw, ImageFont

import huisstijl as hs   # de huisstijl van het werkdocument (uit _sjablonen/lespakket/werkdocument/)

HERE = os.path.dirname(os.path.abspath(__file__))
UIT = os.path.join(HERE, "CO4_Uitnodigen-met-de-agenda.docx")

# Het kader bovenaan: wat Lotte vraagt (lesvoorbereiding § 4)
VRAAG = ("Lotte Maes van de personeelsdienst bij NovaDepot vraagt jou: “Tom, wil jij in de agenda een overleg vastleggen "
         "voor de infosessie van de jobstudenten? Nodig mij en Sara uit. Ik leid het overleg. Alles wat je nodig hebt, "
         "staat hieronder.”")
FEITEN = [
    ("Dag en uur:", "dinsdag 1 december 2026, om 10.00 uur"),
    ("Duur:", "45 minuten (tot 10.45 uur)"),
    ("Plaats:", "die kies je zelf, uit deze drie lokalen van NovaDepot:"),
    ("", "NovaDepot – vergaderzaal 1: 8 plaatsen, met een scherm, gelijkvloers naast het onthaal"),
    ("", "NovaDepot – de kantine: 40 plaatsen, druk tussen 12 en 13 uur"),
    ("", "NovaDepot – het kantoor van Lotte: 2 plaatsen"),
    ("Wie leidt het overleg:", "Lotte Maes, personeelsdienst"),
    ("Reden en doel:", "de infosessie voor de jobstudenten van maandag 21 december voorbereiden"),
    ("De agenda (de punten):", "1 de onthaalbrochure overlopen · 2 de veiligheidsregels in het magazijn (Sara) · "
                               "3 wie onthaalt de jobstudenten om 7.30 uur · 4 rondvraag"),
    ("Kan je niet komen?", "antwoord dan met Nee op de uitnodiging"),
    ("Gasten:", "Lotte (in de les: je leraar) en Sara Claes, preventiedienst (in de les: je duopartner)"),
]

# Deel 1: onderdeel · letter. Rij 1 en 2 zijn al ingevuld (voorgedaan, in het grijs).
ONDERDELEN = [
    ("plaats, datum en tijdstip", "C"),
    ("de duur", "F"),
    ("wie het overleg leidt", ""),
    ("de reden en het doel", ""),
    ("de agenda: de punten die besproken worden", ""),
    ("wat je doet als je niet kan komen", ""),
]

# De criteria van lesvoorbereiding § 22 (samen 20): 4 · 2 · 3 · 2 · 4 · 2 · 2 · 1
CRITERIA = [
    ("Deel 1: ik typte de juiste letter bij wie het overleg leidt, het doel, de agenda en wat je doet als je niet kan "
     "komen.", 4),
    ("Deel 2a: ik koos vergaderzaal 1 en gaf een reden uit het kader.", 2),
    ("Deel 2b: de titel bevat overleg en infosessie; de datum is dinsdag 1 december 2026; het uur is 10.00 tot 10.45.", 3),
    ("Deel 2b: de zaal staat bij Locatie, en de gasten zijn mijn leraar en mijn duopartner.", 2),
    ("Deel 2b: de beschrijving zegt wie het overleg leidt, het doel, de vier agendapunten en wat je doet als je niet kan "
     "komen.", 4),
    ("Deel 2c en 2d: ik plakte een schermafbeelding van mijn afspraak, en ik schreef de drie knoppen en waarop ik "
     "klikte.", 2),
    ("Deel 3: ik beantwoordde de twee vragen over de uitnodiging van mijn duopartner en over mijn eigen afspraak.", 2),
    ("Deel 4: de exitvraag: ik schreef wat Sara doet en wat ik dan zie.", 1),
]


# ---------------------------------------------------------------------------------------------------------------
# De tekening van deel 1: een afspraak in Google Agenda met zes gemerkte stukjes (C F B E A D, van boven naar onder)
# ---------------------------------------------------------------------------------------------------------------
def _lettertype(vet=False, grootte=30):
    for pad in (("/System/Library/Fonts/Supplemental/Arial Bold.ttf" if vet else "/System/Library/Fonts/Supplemental/Arial.ttf"),
                ("C:/Windows/Fonts/arialbd.ttf" if vet else "C:/Windows/Fonts/arial.ttf"),
                ("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if vet else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf")):
        if os.path.exists(pad):
            return ImageFont.truetype(pad, grootte)
    return ImageFont.load_default()


def afspraak_tekening():
    """Geeft een BytesIO met een PNG van de afspraak *Overleg nieuwe rekken magazijn*. De letters staan op de stukjes."""
    b, h = 1500, 740
    img = Image.new("RGB", (b, h), "#FFFFFF")
    d = ImageDraw.Draw(img)
    kl, gr, vet = "#202124", "#5F6368", True
    f_titel, f_tekst, f_klein, f_let = _lettertype(vet, 40), _lettertype(False, 31), _lettertype(False, 27), _lettertype(vet, 30)
    d.rounded_rectangle((2, 2, b - 3, h - 3), radius=22, outline="#9AA0A6", width=3)
    d.rounded_rectangle((2, 2, b - 3, 90), radius=22, fill="#F2F6FC")
    d.rectangle((2, 60, b - 3, 90), fill="#F2F6FC")
    d.line((3, 90, b - 4, 90), fill="#DADCE0", width=2)
    d.ellipse((34, 28, 64, 58), fill="#0B8043")
    d.text((86, 18), "Overleg nieuwe rekken magazijn", font=f_titel, fill=kl)

    def badge(x, y, letter):
        d.ellipse((x, y, x + 46, y + 46), fill="#3C51A0")
        w = d.textlength(letter, font=f_let)
        d.text((x + 23 - w / 2, y + 5), letter, font=f_let, fill="#FFFFFF")

    # C: plaats, datum en tijdstip
    y = 116
    d.rounded_rectangle((26, y, b - 26, y + 112), radius=14, outline="#3C51A0", width=3)
    badge(44, y + 33, "C")
    d.text((120, y + 12), "donderdag 3 december 2026 · 14.00 – 14.30", font=f_tekst, fill=kl)
    d.text((120, y + 60), "Magazijnkantoor", font=f_tekst, fill=kl)
    # gasten (zonder letter)
    y = 250
    d.text((120, y), "Gasten:  Ahmed (aankoop), Lotte", font=f_tekst, fill=gr)
    # de beschrijving, vijf regels, elk een stukje
    y = 300
    d.text((44, y), "Beschrijving", font=f_klein, fill=gr)
    regels = [
        ("F", "Duur: 30 minuten."),
        ("B", "Ahmed leidt het overleg."),
        ("E", "Doel: beslissen welke nieuwe rekken we bestellen."),
        ("A", "Agenda: 1 de oude rekken bekijken · 2 twee offertes vergelijken · 3 beslissen en bestellen."),
        ("D", "Kan je niet komen? Antwoord dan met Nee."),
    ]
    y = 344
    for letter, tekst in regels:
        d.rounded_rectangle((26, y, b - 26, y + 62), radius=12, outline="#3C51A0", width=3)
        badge(44, y + 8, letter)
        d.text((120, y + 14), tekst, font=f_tekst, fill=kl)
        y += 74
    buf = io.BytesIO()
    img.save(buf, "PNG")
    buf.seek(0)
    return buf


def nieuwe_bladzijde(doc):
    """Een bladzijdeskeer vóór een deel met puntenvak: page_break_before in een tabelcel (huisstijl.deel met punten)
    negeert Google Documenten. Een kleine lege alinea die zelf op een nieuwe bladzijde begint, doet het wel."""
    p = doc.add_paragraph()
    p.paragraph_format.page_break_before = True
    hs.zonder_ruimte(p)
    p.paragraph_format.line_spacing = hs.Pt(2)
    r = p.add_run("")
    r.font.size = hs.Pt(1)
    return p


def werkdocument():
    """In de huisstijl van 12-10-2026 (huisstijl.py): alleen wat de leerling invult, en de feiten die hij
    daarvoor nodig heeft. Hoe hij werkt, staat op de lespagina."""
    doc = hs.nieuw_document()
    doc.core_properties.title = "CO4 Uitnodigen met de agenda - werkdocument"
    doc.core_properties.author = "Toegepaste Informatica"
    hs.kopblok(doc, "UITNODIGEN MET DE AGENDA", punten=20)

    hs.kader(doc, "Wat Lotte vraagt", [VRAAG] + FEITEN)

    # ---- Deel 1 ----
    hs.deel(doc, 1, "Wat staat er in een uitnodiging?", punten=4)
    hs.tekst(doc, "Dit is de afspraak Overleg nieuwe rekken magazijn. Zes stukjes hebben een letter.", klein=True,
             cursief=True)
    p = doc.add_paragraph()
    p.alignment = hs.WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(afspraak_tekening(), width=hs.Cm(hs.BREEDTE))
    hs.ruimte(doc, 4)
    hs.tabel(doc, ["Een uitnodiging zegt …", "Letter"], ONDERDELEN, [hs.BREEDTE - 3.0, 3.0], invul=(1,),
             rijhoogte=1.0, voorbeeld_rij=None)
    # de eerste twee rijen zijn voorgedaan: in het grijs
    t = doc.tables[-1]
    for r in (1, 2):
        for k in (0, 1):
            for par in t.rows[r].cells[k].paragraphs:
                for run in par.runs:
                    run.font.color.rgb = hs.GRIJS
                    run.italic = (k == 1)

    # ---- Deel 2 ----
    nieuwe_bladzijde(doc)
    hs.deel(doc, 2, "Het overleg vastleggen", punten=13)
    hs.vraag(doc, "2a. Welke zaal kies je? Typ het lokaal, en één zin waarom het past.")
    hs.tabel(doc, ["Wat", "Jouw antwoord"], [("Het lokaal", ""), ("Waarom het past", "")], [4.2, hs.BREEDTE - 4.2],
             invul=(1,), rijhoogte=1.4, vet=(0,))
    hs.vraag(doc, "2b. Typ over wat je in Google Agenda invulde, één vak per veld.")
    hs.tabel(doc, ["Veld in Google Agenda", "Wat ik invulde"],
             [("Titel", ""), ("Datum", ""), ("Van … tot …", ""), ("Locatie", ""),
              ("Gasten", ""), ("Beschrijving", "")],
             [4.2, hs.BREEDTE - 4.2], invul=(1,), rijhoogte=1.2, vet=(0,))
    t = doc.tables[-1]
    hs.hoogte(t.rows[5], 1.6)    # Gasten: twee voornamen, geen adressen
    hs.hoogte(t.rows[6], 5.2)    # Beschrijving: plakken
    hs.tekst(doc, "Bij Gasten typ je alleen de voornaam van je duopartner en het woord leraar, geen e-mailadressen.",
             klein=True, cursief=True)
    nieuwe_bladzijde(doc)
    hs.vraag(doc, "2c. Plak hier een schermafbeelding van je afspraak.")
    hs.antwoordvak(doc, 6.5)
    hs.tekst(doc, "Lukt het niet? Typ dan hierboven in het vak wat er in je afspraak staat.", klein=True, cursief=True)
    hs.vraag(doc, "2d. Na Opslaan vroeg Google Agenda of je uitnodigingsmails wilde sturen. Welke drie knoppen zag je, "
                  "en op welke klikte je?")
    hs.antwoordvak(doc, 1.6)

    # ---- Deel 3 ----
    hs.deel(doc, 3, "Antwoorden", punten=2)
    hs.vraag(doc, "1. Open de uitnodiging van je duopartner en klik op Ja. Wie leidt het overleg volgens die uitnodiging? "
                  "(Staat er niets? Schrijf dat.)")
    hs.antwoordvak(doc, 1.4)
    hs.vraag(doc, "2. Open je eigen afspraak. Wat zie je bij Gasten? Schrijf wie Ja zei en wie nog niet antwoordde.")
    hs.antwoordvak(doc, 1.8)

    # ---- Deel 4 ----
    nieuwe_bladzijde(doc)
    hs.deel(doc, 4, "Klaar?", punten=1)
    # eerst het codewoord van de zelftest (15-10-2026): de lespagina toont het pas als de zelftest juist is; zo ziet
    # Jonas wie de zelftest echt maakte (_afspraken/lespagina.md, De laatste stap)
    hs.vraag(doc, "Het codewoord van de zelftest (stap 7 op de lespagina):")
    hs.antwoordvak(doc, 0.9)
    hs.vraag(doc, "Sara kan op 1 december niet komen. Wat doet ze, en wat zie jij dan in je agenda? Schrijf één zin.")
    hs.antwoordvak(doc, 1.6)
    hs.controletabel(doc, CRITERIA)

    hs.deel(doc, None, "Extra — niet verplicht")
    hs.vraag(doc, "Open je afspraak en klik op Tijdstip zoeken. Is je duopartner op dinsdag 1 december om 10 uur vrij? "
                  "Schrijf wat je zag.")
    hs.antwoordvak(doc, 2.2)

    doc.save(UIT)


if __name__ == "__main__":
    werkdocument()
    print("Klaar.")
    print("  werkdocument om in te leveren : " + os.path.basename(UIT))
