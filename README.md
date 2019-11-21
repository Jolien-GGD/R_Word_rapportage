# Word rapportages vullen met R
Script waarin met behulp van de officer package een Word document wordt gevuld. In het Word template kan gewerkt worden met bijvoorbeeld 
plaatjes in de koptekst, paginanummering, opmaak, en tekst in kolommen. Tekst komt in de vorm van:
“Van de jongeren op uw school geeft perc_roken aan regelmatig te roken”. 

De officer package zoekt de positie van perc_roken op, en vervangt het woord met een berekend getal, met “TE LAGE N”, of “N = 0”. Na het runnen van het script dient het rapport nog nagelezen te worden om zinnen met deze laatste 2 typen vervangingen handmatig aan te passen (in verband met enkelvoud/meervoud) en om de plaatjes op de juiste plek te krijgen. Theoretisch gezien moet er een manier zijn om dit geautomatiseerd te doen, maar mijn Word template was een moeilijk geval, omdat hij gecreeerd was vanuit een PDF bestand.


