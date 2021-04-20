
### Syntax Word-template vullen met cijfers van ZJHM 
#
# najaar 2018
#
# In deze syntax wordt het Word bestand dat als template dient voor de schoolrapportages voor Zeg Jij Het Maar 
# geopend en worden vooraf aangegeven stukjes tekst (altijd een underscore bevattend) vervangen met de juiste cijfers.
# De cijfers worden berekend vanuit een SPSS databestand, waarin de ruwe data uit Business Objects al is omgezet naar
# indicatoren. 
#
# Het script bestaat uit meerdere delen
# 1. Initialisatie:
#    - Openen van packages
#    - Aanmaken van subfuncties
#    - Word-template openen
#    - Data openen (csv en sav bestand)
# 2. Per hoofdstuk in het rapport het uitrekenen en wegschrijven van de data
#    - Hoofdstuk 1: Inleiding
#    - Hoofdstuk 2: Leefomstandigheden
#    - Hoofdstuk 3: Gezondheid
#    - Hoofdstuk 4: Voeding en Bewegen
#    - Hoofdstuk 5: Genotmiddelen 
#    - Hoofdstuk 6: Seksualiteit
#    - Hoofdstuk 7: Oproepindicaties
# 3. Word document opslaan


#########################
# Deel 1: Initialisatie #
#########################

# Gooi eventueel bestaande variabelen weg.
rm(list=ls())

# Clear console
# cat("\014")


# Set working directory. Kopieer pad vanuit Verkenner en voeg bij elke backslash een tweede backslash (Windows is stom)
setwd("Y:\\Sector Strategie en Ontwikkeling\\EBG\\Epi\\60 Afdelingen\\03. JGZ\\05. Screening VO\\05. ZJHM\\schoolrapportages\\dataverwerking")

# Laad benodigde extra functionaliteit (packages)
library(haven) # Voor inlezen SPSS bestanden
library(officer) # Voor het openen en schrijven naar Word
library(stringr) # Om de labels in de grafieken op meerdere regels te krijgen
library(flextable) # Voor het maken van tabellen 
library(dplyr) # Voor bijna alles
library(ggplot2) # Voor het maken van plots
library(grid) # Nog meer plots
library(gridExtra) # All them plots?
library(gtable) # Voor het combineren van plots tot 1 figuur
library(labelled)
library(tidyr)
library(expss)



# Maak niet automatisch factoren van tekstvariabelen
options(stringsAsFactors = FALSE)


#############
#### Subfuncties
#############

# R's afrondfunctie maakt gebruik van een rare standaard: getal wordt afgerond naar het eerste even getal (0.5 wordt 0, 1.5 wordt 2). 
# Onderstaande functie zorgt ervoor dat 0.5 altijd naar boven wordt afgerond. Deze functie wordt verderop in deze syntax gebruikt.
round2 = function(x) {
  # Eerst bepalen of het 1 decimaal moet zijn, of 0 (kleiner dan 3 is 1 decimaal)
  if (x > 3) {n = 0}
  else {n = 1}
  # Afronden
  scale<-10^n; trunc(x*scale+sign(x)*0.5)/scale}

maak_grafiek <- function(datavariabele) {
  ggplot() + 
    geom_col(data = datavariabele, aes(x = Sortering, y = Percentage), position = "dodge", fill ="#35B8B2", width = 0.5) +
    scale_x_discrete(label = str_wrap(datavariabele$Label, width = 15)) + 
    scale_y_continuous(limits = c(0,100), expand = c(0,0), breaks = seq(0, 100, 10)) +
    xlab("") + 
    ylab("%") +
    geom_text(data = datavariabele, aes(x = Sortering, y = Percentage, label = round(Percentage)), colour = "#002A5C", position = position_stack(vjust = 0.5)) + 
    theme(panel.border = element_blank(),
          panel.grid.major.x = element_blank(),
          panel.grid.major.y = element_line(size=.1, color="grey"),
          panel.grid.minor = element_blank(),
          panel.background = element_blank(),
          axis.title.x = element_text(face = "bold"),
          axis.title.y = element_text(angle = 0, vjust = 1, color = "#002A5C"),
          axis.text = element_text(colour = "#002A5C"))
}


# Voor testdoeleinden
# maak_grafiek(df_Leefomstandigheden)
# maak_grafiek(df_Seksualiteit)
# maak_grafiek(df_Drugs)

######################
# Open Word template #
######################

# Open de Word template 
Rapportage_template <- read_docx(path = "Y:\\Projecten\\Schoolrapportages\\Rapportages\\Template ZJHM\\20181114_Template ZJHM.docx")
#Rapportage_template <- read_docx(path = "Y:\\Sector Strategie en Ontwikkeling\\EBG\\Epi\\60 Afdelingen\\03. JGZ\\05. Screening VO\\05. rapportage obv screening\\schoolrapportages\\dataverwerking\\Rapportage_Compaen_klas3_20181127.docx")


#############
# Open data #
#############

# Open de csv met aanwezigheidsinformatie
df_aanwezig <- read.csv("Data uit BO\\Zuiderzee College\\klas 1\\ZJHM_aanwezigheidsinfo.csv", sep = ";")

# Open het SPSS bestand met de antwoorden op de vragenlijst
data_ZJHM_totaal <- expss::read_spss('databestanden in SPSS\\ZJHM_Zuiderzee_klas1.sav')

# Waarschuwing "Long string missing values record found (record type 7, subtype 22), but ignored" kan genegeerd worden. Komt omdat SPSS bepaalde info niet vrijgeeft

# Maak een dataframe met alleen leerlingen die digitaal hebben meegedaan
data_ZJHM_digitaal <- data_ZJHM_totaal[is.na(data_ZJHM_totaal$Ingevuldop) == FALSE,]

### leerjaar
leerjaar <- c("1")

## Splice dataframes om alleen klas x te hebben.
df_aanwezig <- df_aanwezig[df_aanwezig$Jaar == as.numeric(leerjaar),]
data_ZJHM_digitaal <- data_ZJHM_digitaal[data_ZJHM_digitaal$Jaar == as.numeric(leerjaar),]


##########################
# Hoofdstuk 1: Inleiding #
##########################

# leerjaar
# aant_leerlingen (template code): Aantal leerlingen dat vragenlijst digitaal heeft ingevuld
# aant_jongens: aantal jongens digitaal ingevuld
# aant_meisjes: aantal meisjes digitaal ingevuld

# etniciteit
# leeftijd
# opleidingsniveau

# Vervang in Word
body_replace_all_text(Rapportage_template, "leer_jaar", as.character(leerjaar), fixed=TRUE, only_at_cursor = FALSE)

### Aantal leerlingen dat vragenlijst heeft ingevuld (uit variabele df_aanwezig)
# (Oorspronkelijk gedaan om zowel digitaal als schriftelijk te tellen, want niet iedere jongere deed digitaal mee. Later gekozen om alleen voor digitaal te gaan)

# Als Ingevuldop een waarde heeft, heeft jongere vragenlijst online ingevuld. 
# Als deze leeg is, maar Status staat wel op "Verschenen" dan heeft jongere vragenlijst op papier ingevuld (er waren problemen met inlogcodes).
# Als op papier is ingevuld, staan de antwoorden niet in KD+ (en worden daardoor niet meegenomen in de analyse)
# Vanwege veranderende waarden van de ingeplande afspraken komen dossiernummers dubbel in het bestand... Daarom werken met length(unique())

aant_leerlingen <- length(unique(df_aanwezig$Dossiernummer)) # Aantal unieke dossiers
aant_ingevuld_all <- length(unique(df_aanwezig$Dossiernummer[df_aanwezig$Status == "VERSCHENEN"])) # aantal jongeren met status Verschenen (online of op papier ingevuld)
aant_ingevuld_online <- length(unique(df_aanwezig$Dossiernummer[df_aanwezig$Ingevuldop != ""])) # aantal jongeren online ingevuld

aant_jongens <- length(unique(df_aanwezig$Dossiernummer[df_aanwezig$Geslacht == "MAN" & df_aanwezig$Ingevuldop != ""])) 
aant_meisjes <- length(unique(df_aanwezig$Dossiernummer[df_aanwezig$Geslacht == "VROUW" & df_aanwezig$Ingevuldop != ""]))

# Vervang in Word
body_replace_all_text(Rapportage_template, "aant_leerlingen", as.character(aant_ingevuld_online), fixed=TRUE, only_at_cursor = FALSE)
body_replace_all_text(Rapportage_template, "aant_jongens", as.character(aant_jongens), fixed=TRUE, only_at_cursor = FALSE)
body_replace_all_text(Rapportage_template, "aant_meisjes", as.character(aant_meisjes), fixed=TRUE, only_at_cursor = FALSE)


perc_jongens <- aant_jongens / (aant_jongens + aant_meisjes) * 100
perc_jongens <- round2(perc_jongens) # rond af 

perc_meisjes <- aant_meisjes / (aant_jongens + aant_meisjes) * 100
perc_meisjes <- round2(perc_meisjes) # rond af 

### Opbouw van code om percentages uit te rekenen:
# (count van antwoord waar je in geinteresseerd bent / count van alle antwoorden) * 100.
#
# [dataframe]$[variabelenaam] == (is gelijk aan) [waarde waarvan je percentage wilt weten] geeft een vector met allemaal nullen en enen. 
# Nul als de cel in het databestand niet gelijk is aan [waarde waarvan je percentage wilt weten], 1 als deze wel gelijk is.
# Als je alle enen optelt (met sum() ), dan krijg je het aantal respondenten dat dat specifieke antwoord hebt gegeven. De na.rm = TRUE staat erin om
# R alle missings te laten negeren. Iets optellen waar een missing in staat geeft namelijk altijd een missing (weergegeven als NA) als eindresultaat.
#
# Ik heb ervoor gekozen om het deel met alle antwoorden ook als sum te doen (ipv het aantal regels uit de dataframe, oftewel alle respondenten) te nemen,
# omdat er op die manier flexibel omgesprongen kan worden met het meenemen of uitsluiten van de missings. 
###

## Op naar percentages uitrekenen!

### etniciteit
niet_westers_1 <- (sum(data_ZJHM_digitaal$etni2cat == 1, na.rm = TRUE) / sum(data_ZJHM_digitaal$etni2cat %in% c(0, 1), na.rm = TRUE)) *100 # c(0, 1, 9) om ook de missings mee te laten tellen. Gebruik c(0, 1) als je de missings uit wilt sluiten.
niet_westers <- round2(niet_westers_1) # rond af 

westers_1 <- (sum(data_ZJHM_digitaal$etni2cat == 0, na.rm = TRUE) / sum(data_ZJHM_digitaal$etni2cat %in% c(0, 1), na.rm = TRUE)) *100 # c(0, 1, 9) om ook de missings mee te laten tellen. Gebruik c(0, 1) als je de missings uit wilt sluiten.
westers <- round2(westers_1) # rond af 

### leeftijd
perc_12_1 <- (sum(data_ZJHM_digitaal$LEEFTIJD < 13, na.rm = TRUE) / length(data_ZJHM_digitaal$LEEFTIJD)) *100 
perc_12 <- round2(perc_12_1) # rond af 

perc_13_1  <- (sum(data_ZJHM_digitaal$LEEFTIJD == 13, na.rm = TRUE) / length(data_ZJHM_digitaal$LEEFTIJD)) *100 
perc_13 <- round2(perc_13_1) # rond af 

perc_14_1  <- (sum(data_ZJHM_digitaal$LEEFTIJD == 14, na.rm = TRUE) / length(data_ZJHM_digitaal$LEEFTIJD)) *100 
perc_14 <- round2(perc_14_1) # rond af 

perc_15_1  <- (sum(data_ZJHM_digitaal$LEEFTIJD > 14, na.rm = TRUE) / length(data_ZJHM_digitaal$LEEFTIJD)) *100 
perc_15 <- round2(perc_15_1) # rond af 

### opleidingsniveau
perc_praktijk_1 <- (sum(data_ZJHM_digitaal$Opleiding_ZJHM_nieuw == 1, na.rm = TRUE) / sum(data_ZJHM_digitaal$Opleiding_ZJHM_nieuw %in% c(1, 2, 3, 4, 5, 6, 7, 8), na.rm = TRUE)) *100 
perc_praktijk <- round2(perc_praktijk_1)

perc_schakel_1 <- (sum(data_ZJHM_digitaal$Opleiding_ZJHM_nieuw == 2, na.rm = TRUE) / sum(data_ZJHM_digitaal$Opleiding_ZJHM_nieuw %in% c(1, 2, 3, 4, 5, 6, 7, 8), na.rm = TRUE)) *100 
perc_schakel <- round2(perc_schakel_1)

perc_vmbo_1 <- (sum(data_ZJHM_digitaal$Opleiding_ZJHM_nieuw == 3, na.rm = TRUE) / sum(data_ZJHM_digitaal$Opleiding_ZJHM_nieuw %in% c(1, 2, 3, 4, 5, 6, 7, 8), na.rm = TRUE)) *100 
perc_vmbo <- round2(perc_vmbo_1)

perc_vmbohavo_1 <- (sum(data_ZJHM_digitaal$Opleiding_ZJHM_nieuw == 4, na.rm = TRUE) / sum(data_ZJHM_digitaal$Opleiding_ZJHM_nieuw %in% c(1, 2, 3, 4, 5, 6, 7, 8), na.rm = TRUE)) *100 
perc_vmbohavo <- round2(perc_vmbohavo_1)

perc_havo_hoger_1 <- (sum(data_ZJHM_digitaal$Opleiding_ZJHM_nieuw %in% c(5, 6, 7, 8), na.rm = TRUE) / sum(data_ZJHM_digitaal$Opleiding_ZJHM_nieuw %in% c(1, 2, 3, 4, 5, 6, 7, 8), na.rm = TRUE)) *100 
perc_havo_hoger <- round2(perc_havo_hoger_1)

###
# Tabel maken #
###

Tbl1_aant_leerling <- c("Aantal leerlingen*", aant_ingevuld_online)
Tbl1_perc_meisjes <- c("Meisje (%)", perc_meisjes)
Tbl1_perc_jongens <- c("Jongen (%)", perc_jongens)
Tbl1_perc_Nietwest <- c("Niet-westerse achtergrond (%)", niet_westers)
Tbl1_perc_West <- c("Westerse achtergrond (%)", westers)
Tbl1_lftjonger12 <- c("<= 12 jaar (%)", perc_12)
Tbl1_lft13 <- c("13 jaar (%)", perc_13)
Tbl1_lft14 <- c("14 jaar (%)", perc_14)
Tbl1_lft15 <- c("=> 15 jaar (%)", perc_15)
Tbl1_praktijk <- c("Praktijkonderwijs (%)", perc_praktijk)
Tbl1_schakel <- c("Schakelklas (%)", perc_schakel)
Tbl1_vmbo <- c("Vmbo (%)", perc_vmbo)
Tbl1_vmbohavo <- c("Vmbo / Havo (%)", perc_vmbohavo)
Tbl1_havo_hoger<- c("Havo / Vwo / Gymnasium (%)", perc_havo_hoger)



# Voeg eerste regels samen tot tabel
Tabel1 <- as.data.frame(rbind(Tbl1_aant_leerling, Tbl1_perc_meisjes, Tbl1_perc_jongens, Tbl1_perc_Nietwest, Tbl1_perc_West))
                        
# Voeg leeftijden toe aan tabel, mits percentage niet 0 is
if (perc_12 > 0) {Tabel1 <- rbind(Tabel1, Tbl1_lftjonger12)} 
if (perc_13 > 0) {Tabel1 <- rbind(Tabel1, Tbl1_lft13)} 
if (perc_14 > 0) {Tabel1 <- rbind(Tabel1, Tbl1_lft14)} 
if (perc_15 > 0) {Tabel1 <- rbind(Tabel1, Tbl1_lft15)} 


# Voeg schoolniveau toe aan tabel, mits percentage niet 0 is.
if (perc_praktijk > 0) {Tabel1 <- rbind(Tabel1, Tbl1_praktijk)} 
if (perc_schakel > 0) {Tabel1 <- rbind(Tabel1, Tbl1_schakel)} 
if (perc_vmbo > 0) {Tabel1 <- rbind(Tabel1, Tbl1_vmbo)} 
if (perc_vmbohavo > 0) {Tabel1 <- rbind(Tabel1, Tbl1_vmbohavo)} 
if (perc_havo_hoger > 0) {Tabel1 <- rbind(Tabel1, Tbl1_havo_hoger)} 

colnames(Tabel1) <- c("Kenmerk", "Uw_school")

myft <- flextable(Tabel1)

# Specificeer font type en  grootte
myft <- fontsize(myft, size = 10, part = "all")
myft <- font(myft, fontname = "Arial", part = "all")

# Verander header namen
myft <- set_header_labels(myft, Kenmerk = "" )
myft <- set_header_labels(myft, Uw_school = "Uw school" )

# Verander header kleuren
myft <- bg(myft, bg = "#002659", part = "header")
myft <- color(myft, color = "white", part = "header")
myft <- color(myft, color = "#002A5C", part = "body")

# Verander kleuren van sommige rijen
bg(myft, i = c(2, 3, 6:8), bg = "#DAEEF3", part = "body")

# Verander breedte van de kolommen 
# dim_pretty(myft)
# autofit(myft)
myft <- width(myft, width = 1.5)

# Pas alignment aan van tekst in de kolommen
myft <- align(myft, j = "Kenmerk", align = "left", part = "all" )
myft <- align(myft, j = "Uw_school", align = "center", part = "all" )

# Om in word bestand te zetten: body_add_flextable() 
cursor_reach(Rapportage_template, keyword = "Tabel 1. Kenmerken van de leerlingen") 
# body_remove(my_doc)
body_add_flextable(Rapportage_template, myft, pos = "after")



###################################
# Hoofdstuk 2: Leefomstandigheden #
###################################

# [code in Word template], [variabelenaam van indicator]

# perc_schoolbeleving, Schoolbeleving2CAT
# perc_cijfers, Tevredenheid_cijfers2CAT 
# perc_peercontact, Plezier_peercontact2CAT
# perc_gepest, Frequentie_gepest
# perc_genegeerd, Geestelijk_mishandeld
# perc_lichmish, Lichamelijk_mishandeld
# perc_pestfilmpjes, seksfilmfoto_verspreid
# perc_opkomen, Opkomen2CAT
# meeste_wonen, MBGSK321
# perc_scheiden, GESCHEIDEN
# perc_thuissfeer, Thuissfeer2CAT
# perc_thuissfeer_praten (niet meer in rapportage template)
# perc_thuissfeer_ander (niet meer in rapportage template)
# perc_thuissfeer_niet (niet meer in rapportage template)
# perc_ouders_praten, Ouders_praten2CAT_ZJHM
# perc_ingr_geb
# perc_schuld50, Schulden2CAT
# perc_geldgezin, OnvoldoendeGeldNvt
# perc_geldeten ( => n < 5 probleem?)


############## Begin snippet
# Vul in de 2 stukjes hieronder de variabelen in om te kijken wat de relative frequencies zijn, en de labels die bij de waarden horen.
# Optie van missings excluden zit er (nu?) niet in, vandaar dat de uiteindelijke variabelen op een andere rekenmanier gevuld worden.

# data_ZJHM_digitaal %>%
#  count(FBBMZ3S1) %>%
#  mutate(prop = prop.table(n)) %>%
#  arrange(desc(prop))

# print_labels(data_ZJHM_digitaal$Opkomen2CAT)
############## Einde snippet

cursor_reach(Rapportage_template, keyword = "perc_schoolbeleving")

### Percentage dat 6 of hoger geeft op de vraag "Hoe vind je het op school"
perc_schoolbeleving_1 <- (sum(data_ZJHM_digitaal$Schoolbeleving2CAT == 2, na.rm = TRUE) / sum(data_ZJHM_digitaal$Schoolbeleving2CAT %in% c(1, 2, 3), na.rm = TRUE)) *100 # c(1, 2, 3) om ook de missings mee te laten tellen. Gebruik c(1, 2) als je de missings uit wilt sluiten.
perc_schoolbeleving <- round2(perc_schoolbeleving_1) # rond af 
# Vervang in Word
# cursor_reach(Rapportage_template, keyword = "perc_schoolbeleving")
body_replace_all_text(Rapportage_template, "perc_schoolbeleving", as.character(perc_schoolbeleving), fixed=TRUE, only_at_cursor = FALSE)


### Percentage dat 6 of hoger geeft op de vraag "Hoe tevreden ben je met je cijfers"
perc_cijfers_1 <- (sum(data_ZJHM_digitaal$Tevredenheid_cijfers2CAT == 2, na.rm = TRUE) / sum(data_ZJHM_digitaal$Tevredenheid_cijfers2CAT %in% c(1, 2, 3), na.rm = TRUE)) *100 # c(1, 2, 3) om ook de missings mee te laten tellen. Gebruik c(1, 2) als je de missings uit wilt sluiten.
perc_cijfers <- round2(perc_cijfers_1) # rond af 
# Vervang in Word
# cursor_reach(Rapportage_template, keyword = "perc_cijfers")
body_replace_all_text(Rapportage_template, "perc_cijfers", as.character(perc_cijfers), fixed=TRUE, only_at_cursor = FALSE)

### Percentage dat plezier heeft in contact met leeftijdsgenoten
perc_peercontact_1 <- (sum(data_ZJHM_digitaal$Plezier_peercontact2CAT == 2, na.rm = TRUE) / sum(data_ZJHM_digitaal$Plezier_peercontact2CAT %in% c(1, 2, 3), na.rm = TRUE)) *100 # c(1, 2, 3) om ook de missings mee te laten tellen. Gebruik c(1, 2) als je de missings uit wilt sluiten.
perc_peercontact <- round2(perc_peercontact_1) # rond af 
# Vervang in Word
# cursor_reach(Rapportage_template, keyword = "perc_peercontact")
body_replace_all_text(Rapportage_template, "perc_peercontact", as.character(perc_peercontact), fixed=TRUE, only_at_cursor = FALSE)

### Percentage dat voor zichzelf opkomt ( 6 of hoger) 
perc_opkomen_1 <- (sum(data_ZJHM_digitaal$Opkomen2CAT == 2, na.rm = TRUE) / sum(data_ZJHM_digitaal$Opkomen2CAT %in% c(1, 2), na.rm = TRUE)) *100 # c(1, 2, 3) om ook de missings mee te laten tellen. Gebruik c(1, 2) als je de missings uit wilt sluiten.
perc_opkomen <- round2(perc_opkomen_1) # rond af 
# Vervang in Word
# cursor_reach(Rapportage_template, keyword = "perc_opkomen")
body_replace_all_text(Rapportage_template, "perc_opkomen", as.character(perc_opkomen), fixed=TRUE, only_at_cursor = FALSE)

### Percentage dat vaker dan 1 keer per week gepest wordt 
perc_gepest_1 <- (sum(data_ZJHM_digitaal$Frequentie_gepest == 4, na.rm = TRUE) / sum(data_ZJHM_digitaal$Frequentie_gepest %in% c(1, 2, 3, 4), na.rm = TRUE)) *100 # c(1, 2, 3) om ook de missings mee te laten tellen. Gebruik c(1, 2) als je de missings uit wilt sluiten.
perc_gepest <- round2(perc_gepest_1) # rond af 
# Check of N < 5 of gelijk aan 0. Zo ja, overschrijf. Daarna: Vervang in Word
cursor_reach(Rapportage_template, keyword = "perc_gepest")
if (sum(data_ZJHM_digitaal$Frequentie_gepest == 4, na.rm = TRUE) < 5) {perc_gepest <- "TE LAGE N"} 
if (sum(data_ZJHM_digitaal$Frequentie_gepest == 4, na.rm = TRUE) == 0) {perc_gepest <- "N = 0"} 
body_replace_all_text(Rapportage_template, "perc_gepest", as.character(perc_gepest), fixed=TRUE, only_at_cursor = FALSE)


### Ik word genegeerd, uitgescholden of buitengesloten thuis, op de sportschool, Whatsapp of op school en heb daar last van
perc_genegeerd_1 <- (sum(data_ZJHM_digitaal$Geestelijk_mishandeld %in% c(3,4), na.rm = TRUE) / sum(data_ZJHM_digitaal$Geestelijk_mishandeld %in% c(1, 2, 3, 4, 5), na.rm = TRUE)) *100 # c(1, 2, 3) om ook de missings mee te laten tellen. Gebruik c(1, 2) als je de missings uit wilt sluiten.
perc_genegeerd <- round2(perc_genegeerd_1) # rond af 
# Vervang in Word
# cursor_reach(Rapportage_template, keyword = "perc_genegeerd")
if (sum(data_ZJHM_digitaal$Geestelijk_mishandeld %in% c(3,4), na.rm = TRUE) < 5) {perc_genegeerd <- "TE LAGE N"} 
if (sum(data_ZJHM_digitaal$Geestelijk_mishandeld %in% c(3,4), na.rm = TRUE) == 0) {perc_genegeerd <- "N = 0"} 
body_replace_all_text(Rapportage_template, "perc_genegeerd", as.character(perc_genegeerd), fixed=FALSE, only_at_cursor = FALSE)


### Ik word geschopt, geslagen of op een andere manier mishandeld (en heb daar nog last van)
perc_lichmish_1 <- (sum(data_ZJHM_digitaal$Lichamelijk_mishandeld %in% c(3,4), na.rm = TRUE) / sum(data_ZJHM_digitaal$Lichamelijk_mishandeld %in% c(1, 2, 3, 4), na.rm = TRUE)) *100 # c(1, 2, 3) om ook de missings mee te laten tellen. Gebruik c(1, 2) als je de missings uit wilt sluiten.
perc_lichmish <- round2(perc_lichmish_1) # rond af 
# Check of N < 5 of gelijk aan 0. Zo ja, overschrijf. Daarna: Vervang in Word
# cursor_reach(Rapportage_template, keyword = "perc_lichmish")
if (sum(data_ZJHM_digitaal$Lichamelijk_mishandeld %in% c(3,4), na.rm = TRUE) < 5) {perc_lichmish <- "TE LAGE N"} 
if (sum(data_ZJHM_digitaal$Lichamelijk_mishandeld %in% c(3,4), na.rm = TRUE) == 0) {perc_lichmish <- "N = 0"} 
body_replace_all_text(Rapportage_template, "perc_lichmish", as.character(perc_lichmish), fixed=TRUE, only_at_cursor = FALSE)

### Percentage waarbij tegen wil in filmpjes of foto's zijn verspreid (vraag gaat over seksueel materiaal, template tekst is hier niet duidelijk in)
perc_pestfilmpjes_1 <- (sum(data_ZJHM_digitaal$Seksfilmfoto_verspreid == 1, na.rm = TRUE) / sum(data_ZJHM_digitaal$Seksfilmfoto_verspreid %in% c(0, 1), na.rm = TRUE)) *100 # c(1, 2, 3) om ook de missings mee te laten tellen. Gebruik c(1, 2) als je de missings uit wilt sluiten.
perc_pestfilmpjes <- round2(perc_pestfilmpjes_1) # rond af 
# Check of N < 5 of gelijk aan 0. Zo ja, overschrijf. Daarna: Vervang in Word
# cursor_reach(Rapportage_template, keyword = "perc_pestfilmpjes")
if (sum(data_ZJHM_digitaal$Seksfilmfoto_verspreid == 1, na.rm = TRUE) < 5) {perc_pestfilmpjes <- "TE LAGE N"} 
if (sum(data_ZJHM_digitaal$Seksfilmfoto_verspreid == 1, na.rm = TRUE) == 0) {perc_pestfilmpjes <- "N = 0"} 
body_replace_all_text(Rapportage_template, "perc_pestfilmpjes", as.character(perc_pestfilmpjes), fixed=TRUE, only_at_cursor = FALSE)


#################################### 
### Grafieken Hoofdstuk 2 deel 1 ###
####################################

### Maak dataframe voor plots. 
Percentages <- c(as.character(perc_schoolbeleving),
                 as.character(perc_cijfers), 
                 as.character(perc_peercontact),
                 as.character(perc_opkomen),
                 as.character(perc_gepest), # Later omzetten naar 1 - perc
                 as.character(perc_genegeerd), # Later omzetten naar 1 - perc
                 as.character(perc_lichmish), # Later omzetten naar 1 - perc
                 as.character(perc_pestfilmpjes)# Later omzetten naar 1 - perc
                 ) 

Labels <- c("Geeft school een 6 of hoger",
            "Tevreden over cijfers",
            "Tevreden over contact met leeftijdsgenoten",
            "Komt voor zichzelf op",
            "Afgelopen 3 maanden minder dan wekelijks of niet gepest",
            "Wordt niet genegeerd, uitgescholden of buitengesloten",
            "Wordt niet geschopt, geslagen of op een andere manier mishandeld",
            "Geen ongewilde seksuele of naaktfoto's of -filmpjes verspreid"
            )


# Voeg alle percentages en de labels van de vraag samen in 1 dataframe
df_Leefomstandigheden <- data.frame(Sortering = as.character(1:length(Percentages)), Percentage = Percentages, Label = Labels)
# Verwijder alle percentages waarvan de N te laag was
df_Leefomstandigheden <- df_Leefomstandigheden[!df_Leefomstandigheden$Percentage %in% c("TE LAGE N", "N = 0", "niemand", "Niemand", "geen", "Geen"),]
# Percentages terug van tekst naar nummers
df_Leefomstandigheden$Percentage <- as.numeric(df_Leefomstandigheden$Percentage)

# Zet sommige percentages om naar 1 - percentage.
for (i in 1:nrow(df_Leefomstandigheden)) {
  if (df_Leefomstandigheden[i,3] %in% c("Afgelopen 3 maanden minder dan wekelijks of niet gepest",
                                         "Wordt niet genegeerd, uitgescholden of buitengesloten",
                                         "Wordt niet geschopt, geslagen of op een andere manier mishandeld",
                                         "Geen ongewilde seksuele of naaktfoto's of -filmpjes verspreid")) {
    df_Leefomstandigheden[i,2] <- 100 - df_Leefomstandigheden[i,2]
  }
}

### Maak plots
grafiek_Leefomstandigheden <- maak_grafiek(df_Leefomstandigheden)

### Zet in Word
cursor_reach(Rapportage_template, keyword = "Grafiekgebied_Leefomstandigheden")
body_remove(Rapportage_template)
body_add_gg(Rapportage_template, grafiek_Leefomstandigheden, width = 7, height = 4)


### Woonsituatie meeste dagen van de week
df_meeste_wonen <- data_ZJHM_digitaal %>%
  count(MBGSK321) %>%
  mutate(prop = prop.table(n)) %>%
  arrange(desc(prop))

# Regel werkt niet meer...
# meeste_wonen <- val_label(data_ZJHM_digitaal$MBGSK321, as.vector(df_meeste_wonen$MBGSK321)[1])

# Alternatief
meeste_wonen <- attr(attr(df_meeste_wonen$MBGSK321, "labels"),"names")[as.vector(df_meeste_wonen$MBGSK321)[1]]


if (meeste_wonen == "Bij mijn ouders (samen)") {meeste_wonen_tekst = c("bij hun beide ouders")
} else if (meeste_wonen == "Helft bij mijn ene ouder en helft bij mijn andere ouder") {meeste_wonen_tekst = c("afwisselend de ene en de andere ouder")
} else if (meeste_wonen == "Alleen bij mijn moeder") {meeste_wonen_tekst = c("bij hun moeder")
} else {meeste_wonen_tekst = c("NIET GESPECIFICEERD. VRAAG NA BIJ JOLIEN")
}

# Vervang in Word
# cursor_reach(Rapportage_template, keyword = "Thuissituatie")
body_replace_all_text(Rapportage_template, "meeste_wonen", meeste_wonen_tekst, fixed=TRUE, only_at_cursor = FALSE)


### Percentage waarvan ouders gescheiden zijn
perc_gescheiden_1 <- (sum(data_ZJHM_digitaal$GESCHEIDEN == 1, na.rm = TRUE) / sum(data_ZJHM_digitaal$GESCHEIDEN %in% c(1, 2), na.rm = TRUE)) *100 
perc_gescheiden <- round2(perc_gescheiden_1) # rond af 
# Vervang in Word
body_replace_all_text(Rapportage_template, "perc_scheiden", as.character(perc_gescheiden), fixed=TRUE, only_at_cursor = FALSE)

### Percentage dat thuissfeer voldoende geeft
perc_thuissfeer_1 <- (sum(data_ZJHM_digitaal$Thuissfeer2CAT == 2, na.rm = TRUE) / sum(data_ZJHM_digitaal$Thuissfeer2CAT %in% c(1, 2), na.rm = TRUE)) *100 
perc_thuissfeer <- round2(perc_thuissfeer_1) # rond af 
# Vervang in Word
body_replace_all_text(Rapportage_template, "perc_thuissfeer", as.character(perc_thuissfeer), fixed=TRUE, only_at_cursor = FALSE)

### Percentage dat voldoende met hun ouders kunnen praten
cursor_reach(Rapportage_template, keyword = "perc_ouderspraten")
perc_ouders_praten_1 <- (sum(data_ZJHM_digitaal$Ouders_praten2CAT_ZJHM == 2, na.rm = TRUE) / sum(data_ZJHM_digitaal$Ouders_praten2CAT_ZJHM %in% c(1, 2), na.rm = TRUE)) *100 
perc_ouders_praten  <- round2(perc_ouders_praten_1) # rond af 
# Vervang in Word
body_replace_all_text(Rapportage_template, "perc_ouderspraten", as.character(perc_ouders_praten), fixed=TRUE, only_at_cursor = FALSE)

### Ingrijpende gebeurtenissen
df_ingrijpende_gebeurtenis <- data.frame(Omschrijving = c("echtscheiding van de ouders", "een lichamelijke of psychische ziekte van iemand uit het gezin ", "verslaving van iemand uit het gezin of de vriendenkring", "incest", "het overlijden van een dierbare", "discriminatie van zichzelf of familie", "seksueel misbruik", "een gewelddadige gebeurtenis", "een overige gebeurtenis"), 
          Aantal = c(sum(data_ZJHM_digitaal$Ingr_geb_ZJHM_Scheiding == 1, na.rm = TRUE),
                    sum(data_ZJHM_digitaal$Ingr_geb_ZJHM_Ziekte == 1, na.rm = TRUE),
                    sum(data_ZJHM_digitaal$Ingr_geb_ZJHM_Verslaving == 1, na.rm = TRUE),
                    sum(data_ZJHM_digitaal$Ingr_geb_ZJHM_Incest == 1, na.rm = TRUE),
                    sum(data_ZJHM_digitaal$Ingr_geb_ZJHM_Overlijden == 1, na.rm = TRUE),
                    sum(data_ZJHM_digitaal$Ingr_geb_ZJHM_Discriminatie == 1, na.rm = TRUE),
                    sum(data_ZJHM_digitaal$Ingr_geb_ZJHM_Misbruik == 1, na.rm = TRUE),
                    sum(data_ZJHM_digitaal$Ingr_geb_ZJHM_Geweld == 1, na.rm = TRUE),
                    sum(data_ZJHM_digitaal$Ingr_geb_ZJHM_Anders == 1, na.rm = TRUE)),
          Percentage = c((sum(data_ZJHM_digitaal$Ingr_geb_ZJHM_Scheiding == 1, na.rm = TRUE) / sum(data_ZJHM_digitaal$Ingr_geb_ZJHM_Scheiding %in% c(1, 0), na.rm = TRUE)) *100 ,
                         (sum(data_ZJHM_digitaal$Ingr_geb_ZJHM_Ziekte == 1, na.rm = TRUE) / sum(data_ZJHM_digitaal$Ingr_geb_ZJHM_Ziekte %in% c(1, 0), na.rm = TRUE)) *100 ,
                         (sum(data_ZJHM_digitaal$Ingr_geb_ZJHM_Verslaving == 1, na.rm = TRUE) / sum(data_ZJHM_digitaal$Ingr_geb_ZJHM_Verslaving %in% c(1, 0), na.rm = TRUE)) *100 ,
                         (sum(data_ZJHM_digitaal$Ingr_geb_ZJHM_Incest == 1, na.rm = TRUE) / sum(data_ZJHM_digitaal$Ingr_geb_ZJHM_Incest %in% c(1, 0), na.rm = TRUE)) *100 ,
                         (sum(data_ZJHM_digitaal$Ingr_geb_ZJHM_Overlijden == 1, na.rm = TRUE) / sum(data_ZJHM_digitaal$Ingr_geb_ZJHM_Overlijden %in% c(1, 0), na.rm = TRUE)) *100 ,
                         (sum(data_ZJHM_digitaal$Ingr_geb_ZJHM_Discriminatie == 1, na.rm = TRUE) / sum(data_ZJHM_digitaal$Ingr_geb_ZJHM_Discriminatie %in% c(1, 0), na.rm = TRUE)) *100 ,
                         (sum(data_ZJHM_digitaal$Ingr_geb_ZJHM_Misbruik == 1, na.rm = TRUE) / sum(data_ZJHM_digitaal$Ingr_geb_ZJHM_Misbruik %in% c(1, 0), na.rm = TRUE)) *100 ,
                         (sum(data_ZJHM_digitaal$Ingr_geb_ZJHM_Geweld == 1, na.rm = TRUE) / sum(data_ZJHM_digitaal$Ingr_geb_ZJHM_Geweld %in% c(1, 0), na.rm = TRUE)) *100 ,
                         (sum(data_ZJHM_digitaal$Ingr_geb_ZJHM_Anders == 1, na.rm = TRUE) / sum(data_ZJHM_digitaal$Ingr_geb_ZJHM_Anders %in% c(1, 0), na.rm = TRUE)) *100
))

df_ingrijpende_gebeurtenis <- arrange(df_ingrijpende_gebeurtenis, desc(Percentage))

if (df_ingrijpende_gebeurtenis$Aantal[1] > 4) {gebeurtenis_1 <- df_ingrijpende_gebeurtenis$Omschrijving[1]
} else {gebeurtenis_1 = "TE LAGE N"}
if (df_ingrijpende_gebeurtenis$Aantal[2] > 4) {gebeurtenis_2 <- df_ingrijpende_gebeurtenis$Omschrijving[2]
} else {gebeurtenis_2 = "TE LAGE N"}
if (df_ingrijpende_gebeurtenis$Aantal[3] > 4) {gebeurtenis_3 <- df_ingrijpende_gebeurtenis$Omschrijving[3]
} else {gebeurtenis_3 = "TE LAGE N"}

# Vervang in Word
cursor_reach(Rapportage_template, keyword = "gebeurtenis_1")
body_replace_all_text(Rapportage_template, "gebeurtenis_1", as.character(gebeurtenis_1), fixed=TRUE, only_at_cursor = FALSE)
body_replace_all_text(Rapportage_template, "gebeurtenis_2", as.character(gebeurtenis_2), fixed=TRUE, only_at_cursor = FALSE)
body_replace_all_text(Rapportage_template, "gebeurtenis_3", as.character(gebeurtenis_3), fixed=TRUE, only_at_cursor = FALSE)


### Percentage dat meer dan 50 euro schuld heeft
perc_schuld50_1 <- (sum(data_ZJHM_digitaal$Schulden2CAT == 2, na.rm = TRUE) / sum(data_ZJHM_digitaal$Schulden2CAT %in% c(1, 2), na.rm = TRUE)) *100 
perc_schuld50  <- round2(perc_schuld50_1) # rond af 
if (sum(data_ZJHM_digitaal$Schulden2CAT == 2, na.rm = TRUE) < 5) {perc_schuld50 = "TE LAGE N"}
if (sum(data_ZJHM_digitaal$Schulden2CAT == 2, na.rm = TRUE) == 0) {perc_schuld50 = "N = 0"}
# Vervang in Word
body_replace_all_text(Rapportage_template, "perc_schuld50", as.character(perc_schuld50), fixed=TRUE, only_at_cursor = FALSE)

### Genoeg geld voor eten, sport, kleding en uitjes.
cursor_reach(Rapportage_template, keyword = "perc_geldgezin")
perc_voldoende_geld_1 <- (sum(data_ZJHM_digitaal$OnvoldoendeGeldNvt == 1, na.rm = TRUE) / sum(data_ZJHM_digitaal$OnvoldoendeGeldNvt %in% c(0, 1), na.rm = TRUE)) *100 
perc_voldoende_geld  <- round2(perc_voldoende_geld_1) # rond af 
# Vervang in Word
body_replace_all_text(Rapportage_template, "perc_geldgezin", as.character(perc_voldoende_geld), fixed=TRUE, only_at_cursor = FALSE)

#################################### 
### Grafieken Hoofdstuk 2 deel 2 ###
####################################

### Maak dataframe voor plots.
Percentages2 <- c(as.character(perc_gescheiden),
                 as.character(perc_thuissfeer), 
                 as.character(perc_ouders_praten),
                 as.character(perc_schuld50),
                 as.character(perc_voldoende_geld)) 

Labels2 <- c("Heeft gescheiden ouders",
            "Sfeer thuis is voldoende",
            "Kan voldoende praten met ouders",
            "Geen schuld hoger dan 50 euro",
            "Voldoende geld in gezin voor eten, sport, kleding en uitjes")

# Voeg alle percentages en de labels van de vraag samen in 1 dataframe
df_ThuissitFinancien <- data.frame(Sortering = as.character(1:length(Percentages2)), Percentage = Percentages2, Label = Labels2)
# Verwijder alle percentages waarvan de N te laag was
df_ThuissitFinancien <- df_ThuissitFinancien[!df_ThuissitFinancien$Percentage %in% c("TE LAGE N", "N = 0", "niemand", "Niemand", "geen", "Geen"),]
# Percentages terug van tekst naar nummers
df_ThuissitFinancien$Percentage <- as.numeric(df_ThuissitFinancien$Percentage)

# Zet sommige percentages om naar 1 - percentage.
for (i in 1:nrow(df_ThuissitFinancien)) {
  if (df_ThuissitFinancien[i,3] %in% c("Geen schuld hoger dan 50 euro")) {
    df_ThuissitFinancien[i,2] <- 100 - df_ThuissitFinancien[i,2]
  }
}

### Maak plots
grafiek_df_ThuissitFinancien <- maak_grafiek(df_ThuissitFinancien)

### Zet in Word
cursor_reach(Rapportage_template, keyword = "Grafiekgebied_ThuissitFinancien")
body_remove(Rapportage_template)
body_add_gg(Rapportage_template, grafiek_df_ThuissitFinancien, width = 7, height = 4)


###########################
# Hoofdstuk 3: Gezondheid #
###########################

# perc_ziekte_probleem: MH zegt in overzicht_codes.xlsx dat deze niet in de schoolrapportage moeten

# perc_slaapvallen
# perc_verslapen
# min_slaap
# perc_overdag_slaap
# perc_MHI5
# perc_MHI5_meisjes
# perc_MHI5_jongens

### Percentage dat aangeeft (bijna) altijd binnen een uur in slaap te vallen
perc_slaapvallen_1 <- (sum(data_ZJHM_digitaal$BinnenUurInslapen2CAT == 2, na.rm = TRUE) / sum(data_ZJHM_digitaal$BinnenUurInslapen2CAT %in% c(1, 2), na.rm = TRUE)) *100 
perc_slaapvallen <- round2(perc_slaapvallen_1) # rond af 
# Vervang in Word
body_replace_all_text(Rapportage_template, "perc_slaapvallen", as.character(perc_slaapvallen), fixed=TRUE, only_at_cursor = FALSE)

### Percentage dat zich vaak verslaapt
perc_verslapen_1 <- (sum(data_ZJHM_digitaal$Verslapen == 3, na.rm = TRUE) / sum(data_ZJHM_digitaal$Verslapen %in% c(1, 2, 3), na.rm = TRUE)) *100 
perc_verslapen <- round2(perc_verslapen_1) # rond af 
# Vervang in Word
if (sum(data_ZJHM_digitaal$Verslapen == 3, na.rm = TRUE) < 5) {perc_verslapen <- "TE LAGE N"} 
if (sum(data_ZJHM_digitaal$Verslapen == 3, na.rm = TRUE) == 0) {perc_verslapen <- "N = 0"} 
body_replace_all_text(Rapportage_template, "perc_verslapen", as.character(perc_verslapen), fixed=TRUE, only_at_cursor = FALSE)


### Aantal uren slaap per nacht
df_slapen <- data_ZJHM_digitaal %>%
  count(GemiddeldeSlaaptijd) %>%
  mutate(prop = prop.table(n)) %>%
  arrange(desc(prop))

# Regel werkt niet meer...
# gemiddeld_slaap <- val_label(data_ZJHM_digitaal$GemiddeldeSlaaptijd, df_slapen[1,1,1])

# Alternatief
gemiddeld_slaap <- attr(attr(data_ZJHM_digitaal$GemiddeldeSlaaptijd, "labels"),"names")[as.vector(df_slapen$GemiddeldeSlaaptijd)[1]]


# Vervang in Word
body_replace_all_text(Rapportage_template, "min_slaap", as.character(gemiddeld_slaap), fixed=TRUE, only_at_cursor = FALSE)

### Percentage dat overdag "wel eens of vaak" slaapt
perc_overdag_slaap_1 <- (sum(data_ZJHM_digitaal$Verslapen %in% c(2,3), na.rm = TRUE) / sum(data_ZJHM_digitaal$Verslapen %in% c(1, 2, 3), na.rm = TRUE)) *100 
perc_overdag_slaap <- round2(perc_overdag_slaap_1) # rond af 
# Vervang in Word
body_replace_all_text(Rapportage_template, "perc_overdag_slaap", as.character(perc_overdag_slaap), fixed=TRUE, only_at_cursor = FALSE)

# MHI-5
perc_MHI5_1 <- (sum(data_ZJHM_digitaal$PsyGez == 1, na.rm = TRUE) / sum(data_ZJHM_digitaal$PsyGez %in% c(1, 2), na.rm = TRUE)) *100 
perc_MHI5 <- round2(perc_MHI5_1) # rond af 
# Vervang in Word
cursor_reach(Rapportage_template, keyword = "perc_MHI5")
body_replace_all_text(Rapportage_template, "perc_MHI5", as.character(perc_MHI5), fixed=TRUE, only_at_cursor = FALSE)


## MHI-5 per geslacht
# df_MHI5_geslacht <- data_ZJHM_digitaal %>%
#   group_by(GESLACHT) %>%
#   count(PsyGez) %>%
#   mutate(prop = prop.table(n)*100)
# 
# # Psychisch gezonde jongens
# perc_MHI5_jongens_1 <- df_MHI5_geslacht[df_MHI5_geslacht$GESLACHT == 1 & df_MHI5_geslacht$PsyGez == 2,4]
# perc_MHI5_jongens <- round2(perc_MHI5_jongens_1) # rond af 
# Voeg check <5 in
# body_replace_all_text(Rapportage_template, "percMHI5_jongens", as.character(perc_MHI5_jongens), fixed=TRUE, only_at_cursor = FALSE)
# 
# # Psychisch gezonde jongens
# perc_MHI5_meisjes_1 <- df_MHI5_geslacht[df_MHI5_geslacht$GESLACHT == 2 & df_MHI5_geslacht$PsyGez == 2,4]
# perc_MHI5_meisjes <- round2(perc_MHI5_meisjes_1) # rond af 
# Voeg check <5 in
# body_replace_all_text(Rapportage_template, "percMHI5_meisjes", as.character(perc_MHI5_meisjes), fixed=TRUE, only_at_cursor = FALSE)

# => MHI-5 per geslacht staat niet meer in template. 

############################# 
### Grafieken Hoofdstuk 3 ###
############################# 

### Maak dataframe voor plots.
Percentages3 <- c(as.character(perc_slaapvallen),
                  as.character(perc_verslapen), 
                  as.character(perc_overdag_slaap),
                  as.character(perc_MHI5)) 

Labels3 <- c("Valt (bijna) altijd binnen een uur in slaap",
             "Verslaapt zich niet vaak",
             "Slaapt niet overdag",
             "Psychisch gezond")

# Voeg alle percentages en de labels van de vraag samen in 1 dataframe
df_LichGezondheid <- data.frame(Sortering = as.character(1:length(Percentages3)), Percentage = Percentages3, Label = Labels3)
# Verwijder alle percentages waarvan de N te laag was
df_LichGezondheid <- df_LichGezondheid[!df_LichGezondheid$Percentage %in% c("TE LAGE N", "N = 0", "niemand", "Niemand", "geen", "Geen"),]
# Percentages terug van tekst naar nummers
df_LichGezondheid$Percentage <- as.numeric(df_LichGezondheid$Percentage)

# Zet sommige percentages om naar 1 - percentage.
for (i in 1:nrow(df_LichGezondheid)) {
  if (df_LichGezondheid[i,3] %in% c("Verslaapt zich niet vaak", "Slaapt niet overdag", "Psychisch gezond")) {
    df_LichGezondheid[i,2] <- 100 - df_LichGezondheid[i,2]
  }
}

### Maak plots
grafiek_df_LichGezondheid <- maak_grafiek(df_LichGezondheid)

### Zet in Word
cursor_reach(Rapportage_template, keyword = "Grafiekgebied_LichGezondheid")
body_remove(Rapportage_template)
body_add_gg(Rapportage_template, grafiek_df_LichGezondheid, width = 7, height = 4)



###################################
# Hoofdstuk 4: Voeding en bewegen #
###################################

# perc_overgewicht
# perc_obesitas
# perc_ondergewicht
# perc_probleemeten
# perc_eten_meisjes_regio (Suggestie voor DBC eigen getallen, wrs n < 5 probleem)
# perc_eten_jongens_regio (Suggestie voor DBC eigen getallen, wrs n < 5 probleem)
# perc_beweegnorm

cursor_reach(Rapportage_template, keyword = "perc_overgewicht")

### Percentage met overgewicht
perc_overgewicht_1 <- (sum(data_ZJHM_digitaal$FBBMZ3S1 %in% c(4,5), na.rm = TRUE) / sum(data_ZJHM_digitaal$FBBMZ3S1 %in% c(1, 2, 3, 4, 5), na.rm = TRUE)) *100 
perc_overgewicht <- round2(perc_overgewicht_1) # rond af 
# Vervang in Word
if (sum(data_ZJHM_digitaal$FBBMZ3S1 %in% c(4,5), na.rm = TRUE) < 5) {perc_overgewicht <- "TE LAGE N"}
if (sum(data_ZJHM_digitaal$FBBMZ3S1 %in% c(4,5), na.rm = TRUE) == 0) {perc_overgewicht <- "N = 0"}
body_replace_all_text(Rapportage_template, "perc_overgewicht", as.character(perc_overgewicht), fixed=TRUE, only_at_cursor = FALSE)

### Percentage met obesitas: Niet meer als aparte categorie, maar samengevoegd met overgewicht.
# perc_obesitas_1 <- (sum(data_ZJHM_digitaal$FBBMZ3S1 == 5, na.rm = TRUE) / sum(data_ZJHM_digitaal$FBBMZ3S1 %in% c(1, 2, 3, 4, 5), na.rm = TRUE)) *100 
# perc_obesitas <- round2(perc_obesitas_1) # rond af 
# # Check of N < 5 of gelijk aan 0. Zo ja, overschrijf. Daarna: Vervang in Word
# if (sum(data_ZJHM_digitaal$FBBMZ3S1 == 5, na.rm = TRUE) < 5) {perc_obesitas <- "TE LAGE N"}
# if (sum(data_ZJHM_digitaal$FBBMZ3S1 == 5, na.rm = TRUE) == 0) {perc_obesitas <- "N = 0"}
# body_replace_all_text(Rapportage_template, "perc_obesitas", as.character(perc_obesitas), fixed=TRUE, only_at_cursor = FALSE)

### Percentage met ondergewicht
perc_ondergewicht_1 <- (sum(data_ZJHM_digitaal$FBBMZ3S1 %in% c(1, 2), na.rm = TRUE) / sum(data_ZJHM_digitaal$FBBMZ3S1 %in% c(1, 2, 3, 4, 5), na.rm = TRUE)) *100 
perc_ondergewicht <- round2(perc_ondergewicht_1) # rond af 
# Check of N < 5 of gelijk aan 0. Zo ja, overschrijf. Daarna: Vervang in Word
if (sum(data_ZJHM_digitaal$FBBMZ3S1 %in% c(1, 2), na.rm = TRUE) < 5) {perc_ondergewicht <- "TE LAGE N"} 
if (sum(data_ZJHM_digitaal$FBBMZ3S1 %in% c(1, 2), na.rm = TRUE) == 0) {perc_ondergewicht <- "N = 0"} 
body_replace_all_text(Rapportage_template, "perc_ondergewicht", as.character(perc_ondergewicht), fixed=TRUE, only_at_cursor = FALSE)

cursor_reach(Rapportage_template, keyword = "perc_probleemeten")

### Percentage dat aangeeft "wel eens" of "vaak" problemen te hebben met eten
perc_probleemeten_1 <- (sum(data_ZJHM_digitaal$Eetproblemen2CAT == 2, na.rm = TRUE) / sum(data_ZJHM_digitaal$Eetproblemen2CAT %in% c(1, 2), na.rm = TRUE)) *100 
perc_probleemeten <- round2(perc_probleemeten_1) # rond af 
# Check of N < 5 of gelijk aan 0. Zo ja, overschrijf. Daarna: Vervang in Word
if (sum(data_ZJHM_digitaal$Eetproblemen2CAT == 2, na.rm = TRUE) < 5) {perc_probleemeten <- "TE LAGE N"} 
if (sum(data_ZJHM_digitaal$Eetproblemen2CAT == 2, na.rm = TRUE) == 0) {perc_probleemeten <- "N = 0"} 
body_replace_all_text(Rapportage_template, "perc_probleemeten", as.character(perc_probleemeten), fixed=TRUE, only_at_cursor = FALSE)

### Percentage eetproblemen per geslacht
tibble_eetproblemen <- data_ZJHM_digitaal %>%
  filter(Eetproblemen2CAT %in% c(1,2)) %>%
  group_by(GESLACHT) %>%
  count(Eetproblemen2CAT) %>%
  mutate(prop = prop.table(n)) 

# Resultaat van hierboven komt als tibble, en ik weet niet hoe ik tibbles moet subsetten. Daarom gecast naar dataframe.
df_eetproblemen <- as.data.frame(tibble_eetproblemen)


# Percentage probleemeten jongens
perc_probleemeten_jongens <- round2(df_eetproblemen[df_eetproblemen$GESLACHT == 1 & df_eetproblemen$Eetproblemen == 2, 4]*100) # Kolom 4 heeft proporties
# Check of N < 5 of gelijk aan 0. Zo ja, overschrijf. Daarna: Vervang in Word
if (sum(data_ZJHM_digitaal$Eetproblemen2CAT == 2 & data_ZJHM_digitaal$GESLACHT == 1, na.rm = TRUE) < 5) {perc_probleemeten_jongens <- "TE LAGE N"} 
if (sum(data_ZJHM_digitaal$Eetproblemen2CAT == 2 & data_ZJHM_digitaal$GESLACHT == 1, na.rm = TRUE) == 0) {perc_probleemeten_jongens <- "N = 0"} 
body_replace_all_text(Rapportage_template, "perc_etenjongens", as.character(perc_probleemeten_jongens), fixed=TRUE, only_at_cursor = FALSE)

# Percentage probleemeten meisjes
perc_probleemeten_meisjes <- round2(df_eetproblemen[df_eetproblemen$GESLACHT == 2 & df_eetproblemen$Eetproblemen == 2, 4]*100)
# Check of N < 5 of gelijk aan 0. Zo ja, overschrijf. Daarna: Vervang in Word
if (sum(data_ZJHM_digitaal$Eetproblemen2CAT == 2 & data_ZJHM_digitaal$GESLACHT == 2, na.rm = TRUE) < 5) {perc_probleemeten_meisjes <- "TE LAGE N"} 
if (sum(data_ZJHM_digitaal$Eetproblemen2CAT == 2 & data_ZJHM_digitaal$GESLACHT == 2, na.rm = TRUE) == 0) {perc_probleemeten_meisjes <- "N = 0"} 
body_replace_all_text(Rapportage_template, "perc_etenmeisjes", as.character(perc_probleemeten_meisjes), fixed=TRUE, only_at_cursor = FALSE)


cursor_reach(Rapportage_template, keyword = "perc_beweegnorm")

### Percentage dat minimaal 1 uur per dag beweegt
perc_beweegnorm_1 <- (sum(data_ZJHM_digitaal$DagelijkseBeweging2CAT == 2, na.rm = TRUE) / sum(data_ZJHM_digitaal$DagelijkseBeweging2CAT %in% c(1, 2), na.rm = TRUE)) *100 
perc_beweegnorm <- round2(perc_beweegnorm_1) # rond af 
# Vervang in Word
body_replace_all_text(Rapportage_template, "perc_beweegnorm", as.character(perc_beweegnorm), fixed=TRUE, only_at_cursor = FALSE)


############################# 
### Grafieken Hoofdstuk 4 ###
############################# 

### Maak dataframe voor plots. 
Percentages4 <- c(as.character(perc_overgewicht), 
                  as.character(perc_ondergewicht),
                 as.character(perc_probleemeten),
                 as.character(perc_beweegnorm)) 

Labels4 <- c("Overgewicht",
            "Ondergewicht",
            "Probleemeten",
            "Beweegt minder dan 1 uur per dag")

# Voeg alle percentages en de labels van de vraag samen in 1 dataframe
df_VoedingBewegen <- data.frame(Sortering = as.character(1:length(Percentages4)), Percentage = Percentages4, Label = Labels4)
# Verwijder alle percentages waarvan de N te laag was
df_VoedingBewegen <- df_VoedingBewegen[!df_VoedingBewegen$Percentage %in% c("TE LAGE N", "N = 0", "niemand", "Niemand", "geen", "Geen"),]
# Percentages terug van tekst naar nummers
df_VoedingBewegen$Percentage <- as.numeric(df_VoedingBewegen$Percentage)

# Zet sommige percentages om naar 1 - percentage.
for (i in 1:nrow(df_VoedingBewegen)) {
  if (df_VoedingBewegen[i,3] %in% c("Beweegt minder dan 1 uur per dag")) {
    df_VoedingBewegen[i,2] <- 100 - df_VoedingBewegen[i,2]
  }
}

### Maak plots
grafiek_VoedingBewegen <- maak_grafiek(df_VoedingBewegen)

### Zet in Word
cursor_reach(Rapportage_template, keyword = "Grafiekgebied_Voeding_bewegen")
body_remove(Rapportage_template)
body_add_gg(Rapportage_template, grafiek_VoedingBewegen, width = 7, height = 3)


##############################
# Hoofdstuk 5: Genotmiddelen #
##############################

# perc_geen_alcohol
# perc_4wekenalcohol
# perc_5drankjes
# perc_roken_wekelijks => Op dit moment problemen met uitrekenen
# perc_rookpauze
# perc_weleenswiet
# perc_recentwiet
# perc_wietpauze
# perc_lachgas => Wel analyseren, niet in rapportage?
# perc_lachgas_recent => wel analyseren, niet in rapportage?
# perc_lachgas_pauze => wel analyseren, niet in rapportage?
# perc_drugsoverig
# perc_internet
# perc_gamen
# nietgenoegtijd
# perc_genoegtijd
# moeilijk_stoppen
# perc_stoppen

cursor_reach(Rapportage_template, keyword = "perc_geen_alcohol")

### Percentage dat geen alcohol drinkt
perc_geen_alcohol_1 <- (sum(data_ZJHM_digitaal$WeleensAlcohol2CAT == 0, na.rm = TRUE) / sum(data_ZJHM_digitaal$WeleensAlcohol2CAT %in% c(0, 1, 88), na.rm = TRUE)) *100 
perc_geen_alcohol <- round2(perc_geen_alcohol_1) # rond af 
# Check of N < 5 of gelijk aan 0. Zo ja, overschrijf. Daarna: Vervang in Word
if (sum(data_ZJHM_digitaal$WeleensAlcohol2CAT == 0, na.rm = TRUE) < 5) {perc_geen_alcohol <- "TE LAGE N"} 
if (sum(data_ZJHM_digitaal$WeleensAlcohol2CAT == 0, na.rm = TRUE) == 0) {perc_geen_alcohol <- "N = 0"} 
body_replace_all_text(Rapportage_template, "perc_geen_alcohol", as.character(perc_geen_alcohol), fixed=TRUE, only_at_cursor = FALSE)

### Percentage dat in de afgelopen 4 weken alcohol heeft gedronken
perc_4wekenalcohol_1 <- (sum(data_ZJHM_digitaal$DagenRecentAlcohol2CAT == 1, na.rm = TRUE) / sum(data_ZJHM_digitaal$DagenRecentAlcohol2CAT %in% c(0, 1, 88), na.rm = TRUE)) *100 
perc_4wekenalcohol <- round2(perc_4wekenalcohol_1) # rond af 
# Check of N < 5 of gelijk aan 0. Zo ja, overschrijf. Daarna: Vervang in Word
if (sum(data_ZJHM_digitaal$DagenRecentAlcohol2CAT == 1, na.rm = TRUE) < 5) {perc_4wekenalcohol <- "TE LAGE N"} 
if (sum(data_ZJHM_digitaal$DagenRecentAlcohol2CAT == 1, na.rm = TRUE) == 0) {perc_4wekenalcohol <- "N = 0"} 
body_replace_all_text(Rapportage_template, "perc_4wekenalcohol", as.character(perc_4wekenalcohol), fixed=TRUE, only_at_cursor = FALSE)


### Percentage dat 5 of meer drankjes heeft gedronken bij 1 gelegenheid
perc_5drankjes_1 <- (sum(data_ZJHM_digitaal$BingeDrinken2CAT == 1, na.rm = TRUE) / sum(data_ZJHM_digitaal$BingeDrinken2CAT %in% c(0, 1, 88), na.rm = TRUE)) *100 
perc_5drankjes <- round2(perc_5drankjes_1) # rond af 
# Check of N < 5 of gelijk aan 0. Zo ja, overschrijf. Daarna: Vervang in Word
if (sum(data_ZJHM_digitaal$BingeDrinken2CAT == 1, na.rm = TRUE) < 5) {perc_5drankjes <- "TE LAGE N"} 
if (sum(data_ZJHM_digitaal$BingeDrinken2CAT == 1, na.rm = TRUE) == 0) {perc_5drankjes <- "N = 0"} 
body_replace_all_text(Rapportage_template, "perc_5drankjes", as.character(perc_5drankjes), fixed=TRUE, only_at_cursor = FALSE)

cursor_reach(Rapportage_template, keyword = "perc_roken")

### Percentage dat wekelijks rookt
perc_roken_wekelijks_1 <- (sum(data_ZJHM_digitaal$HoeVaakRoken2CAT == 1, na.rm = TRUE) / sum(data_ZJHM_digitaal$HoeVaakRoken2CAT %in% c(0, 1, 88), na.rm = TRUE)) *100 
perc_roken_wekelijks <- round2(perc_roken_wekelijks_1) # rond af 
# Check of N < 5 of gelijk aan 0. Zo ja, overschrijf. Daarna: Vervang in Word
if (sum(data_ZJHM_digitaal$HoeVaakRoken2CAT == 1, na.rm = TRUE) < 5) {perc_roken_wekelijks <- "TE LAGE N"} 
if (sum(data_ZJHM_digitaal$HoeVaakRoken2CAT == 1, na.rm = TRUE) == 0) {perc_roken_wekelijks <- "N = 0"} 
body_replace_all_text(Rapportage_template, "perc_roken", as.character(perc_roken_wekelijks), fixed=TRUE, only_at_cursor = FALSE)

### Percentage dat rookt in de pauzes
perc_rookpauze_1 <- (sum(data_ZJHM_digitaal$RokenInPauze == 1, na.rm = TRUE) / sum(data_ZJHM_digitaal$RokenInPauze %in% c(0, 1, 88), na.rm = TRUE)) *100 
perc_rookpauze <- round2(perc_rookpauze_1) # rond af 
# Check of N < 5 of gelijk aan 0. Zo ja, overschrijf. Daarna: Vervang in Word
if (sum(data_ZJHM_digitaal$RokenInPauze == 1, na.rm = TRUE) < 5) {perc_rookpauze <- "TE LAGE N"} 
if (sum(data_ZJHM_digitaal$RokenInPauze == 1, na.rm = TRUE) == 0) {perc_rookpauze <- "N = 0"} 
body_replace_all_text(Rapportage_template, "perc_rookpauze", as.character(perc_rookpauze), fixed=TRUE, only_at_cursor = FALSE)

### Percentage dat recent wiet heeft gerookt
perc_recentwiet_1 <- (sum(data_ZJHM_digitaal$RecentHasjWiet2CAT == 1, na.rm = TRUE) / sum(data_ZJHM_digitaal$RecentHasjWiet2CAT %in% c(0, 1, 88), na.rm = TRUE)) *100 
perc_recentwiet <- round2(perc_recentwiet_1) # rond af 
# Check of N < 5 of gelijk aan 0. Zo ja, overschrijf. Daarna: Vervang in Word
if (sum(data_ZJHM_digitaal$RecentHasjWiet2CAT == 1, na.rm = TRUE) < 5) {perc_recentwiet <- "TE LAGE N"} 
if (sum(data_ZJHM_digitaal$RecentHasjWiet2CAT == 1, na.rm = TRUE) == 0) {perc_recentwiet <- "N = 0"} 
body_replace_all_text(Rapportage_template, "perc_recentwiet", as.character(perc_recentwiet), fixed=TRUE, only_at_cursor = FALSE)

### Percentage dat blowt in de pauzes
perc_wietpauze_1 <- (sum(data_ZJHM_digitaal$BlowenInPauze == 1, na.rm = TRUE) / sum(data_ZJHM_digitaal$BlowenInPauze %in% c(0, 1, 88), na.rm = TRUE)) *100 
perc_wietpauze <- round2(perc_wietpauze_1) # rond af 
# Check of N < 5 of gelijk aan 0. Zo ja, overschrijf. Daarna: Vervang in Word
if (sum(data_ZJHM_digitaal$BlowenInPauze == 1, na.rm = TRUE) < 5) {perc_wietpauze <- "TE LAGE N"} 
if (sum(data_ZJHM_digitaal$BlowenInPauze == 1, na.rm = TRUE) == 0) {perc_wietpauze <- "N = 0"} 
body_replace_all_text(Rapportage_template, "perc_wietpauze", as.character(perc_wietpauze), fixed=TRUE, only_at_cursor = FALSE)


cursor_reach(Rapportage_template, keyword = "perc_lachgas")

### Percentage dat ooit lachgas heeft gebruikt
perc_lachgas_1 <- (sum(data_ZJHM_digitaal$Ooitlachgas == 1, na.rm = TRUE) / sum(data_ZJHM_digitaal$Ooitlachgas %in% c(0, 1, 88), na.rm = TRUE)) *100 
perc_lachgas <- round2(perc_lachgas_1) # rond af 
# Check of N < 5 of gelijk aan 0. Zo ja, overschrijf. Daarna: Vervang in Word
if (sum(data_ZJHM_digitaal$Ooitlachgas == 1, na.rm = TRUE) < 5) {perc_lachgas <- "TE LAGE N"} 
if (sum(data_ZJHM_digitaal$Ooitlachgas == 1, na.rm = TRUE) == 0) {perc_lachgas <- "N = 0"} 
body_replace_all_text(Rapportage_template, "perc_lachgas", as.character(perc_lachgas), fixed=TRUE, only_at_cursor = FALSE)

### Percentage dat recent lachgas heeft gebruikt
perc_lachgas_recent_1 <- (sum(data_ZJHM_digitaal$HoevaakRecentLachgas %in% c(2, 3, 4, 5, 6), na.rm = TRUE) / sum(data_ZJHM_digitaal$HoevaakRecentLachgas %in% c(1, 2, 3, 4, 5, 6, 88), na.rm = TRUE)) *100 
perc_lachgas_recent <- round2(perc_lachgas_recent_1) # rond af 
# Check of N < 5 of gelijk aan 0. Zo ja, overschrijf. Daarna: Vervang in Word
if (sum(data_ZJHM_digitaal$HoevaakRecentLachgas %in% c(2, 3, 4, 5, 6), na.rm = TRUE) < 5) {perc_lachgas_recent <- "TE LAGE N"} 
if (sum(data_ZJHM_digitaal$HoevaakRecentLachgas %in% c(2, 3, 4, 5, 6), na.rm = TRUE) == 0) {perc_lachgas_recent <- "N = 0"} 
body_replace_all_text(Rapportage_template, "perclachgas_recent", as.character(perc_lachgas_recent), fixed=TRUE, only_at_cursor = FALSE)

### Percentage dat lachgas gebruikt in de pauzes
perc_lachgas_pauze_1 <- (sum(data_ZJHM_digitaal$LachgasInPauze == 1, na.rm = TRUE) / sum(data_ZJHM_digitaal$LachgasInPauze %in% c(0, 1, 88), na.rm = TRUE)) *100 
perc_lachgas_pauze <- round2(perc_lachgas_pauze_1) # rond af 
# Check of N < 5 of gelijk aan 0. Zo ja, overschrijf. Daarna: Vervang in Word
if (sum(data_ZJHM_digitaal$LachgasInPauze == 1, na.rm = TRUE) < 5) {perc_lachgas_pauze <- "TE LAGE N"} 
if (sum(data_ZJHM_digitaal$LachgasInPauze == 1, na.rm = TRUE) == 0) {perc_lachgas_pauze <- "N = 0"} 
body_replace_all_text(Rapportage_template, "perclachgas_pauze", as.character(perc_lachgas_pauze), fixed=TRUE, only_at_cursor = FALSE)


cursor_reach(Rapportage_template, keyword = "perc_drugsoverig")

### Percentage dat ooit overige drugs heeft gebruikt
perc_drugsoverig_1 <- (sum(data_ZJHM_digitaal$DrugsOverig_ooit == 1, na.rm = TRUE) / sum(data_ZJHM_digitaal$DrugsOverig_ooit %in% c(0, 1, 88), na.rm = TRUE)) *100 
perc_drugsoverig <- round2(perc_drugsoverig_1) # rond af 
# Check of N < 5 of gelijk aan 0. Zo ja, overschrijf. Daarna: Vervang in Word
if (sum(data_ZJHM_digitaal$DrugsOverig_ooit == 1, na.rm = TRUE) < 5) {perc_drugsoverig <- "TE LAGE N"} 
if (sum(data_ZJHM_digitaal$DrugsOverig_ooit == 1, na.rm = TRUE) == 0) {perc_drugsoverig <- "N = 0"} 
body_replace_all_text(Rapportage_template, "perc_drugsoverig", as.character(perc_drugsoverig), fixed=TRUE, only_at_cursor = FALSE)


#################################### 
### Grafieken Hoofdstuk 5 deel 1 ###
#################################### 

### Maak dataframe voor plots. 
Percentages5_1 <- c(as.character(perc_geen_alcohol),
                    as.character(perc_4wekenalcohol), 
                    as.character(perc_5drankjes),
                    as.character(perc_roken_wekelijks),
                    as.character(perc_recentwiet),
                    as.character(perc_lachgas),
                    as.character(perc_lachgas_recent),
                    as.character(perc_drugsoverig)) 

Labels5_1 <- c("Drinkt wel eens alcohol",
               "Dronk in de afgelopen 4 weken alcohol",
               "Dronk recent 5 of meer drankjes",
               "Rookt wekelijks",
               "Rookte recent wiet",
               "Heeft ooit lachgas gebruikt",
               "Gebruikte recent lachgas",
               "Heeft ooit harddrugs gebruikt")

# Voeg alle percentages en de labels van de vraag samen in 1 dataframe
df_Drugs <- data.frame(Sortering = as.character(1:length(Percentages5_1)), Percentage = Percentages5_1, Label = Labels5_1)
# Verwijder alle percentages waarvan de N te laag was
df_Drugs <- df_Drugs[!df_Drugs$Percentage %in% c("TE LAGE N", "N = 0", "niemand", "Niemand", "geen", "Geen"),]
# Percentages terug van tekst naar nummers
df_Drugs$Percentage <- as.numeric(df_Drugs$Percentage)

# Zet sommige percentages om naar 1 - percentage.
for (i in 1:nrow(df_Drugs)) {
  if (df_Drugs[i,3] %in% c("Drinkt wel eens alcohol")) {
    df_Drugs[i,2] <- 100 - df_Drugs[i,2]
  }
}

### Maak plots
grafiek_Drugs <- maak_grafiek(df_Drugs)

### Zet in Word
cursor_reach(Rapportage_template, keyword = "Grafiekgebied_Drugs")
body_remove(Rapportage_template)
body_add_gg(Rapportage_template, grafiek_Drugs, width = 7, height = 4)


###
### Percentages deel 2


cursor_reach(Rapportage_template, keyword = "perc_internet")


### Percentage dat dagelijks minimaal 3 uur in de vrije tijd besteed aan internet en mobiele telefonie
perc_internet_1 <- (sum(data_ZJHM_digitaal$Frequentie_3uurInternet == 1, na.rm = TRUE) / sum(data_ZJHM_digitaal$Frequentie_3uurInternet %in% c(0, 1), na.rm = TRUE)) *100 
perc_internet <- round2(perc_internet_1) # rond af 
# Check of N < 5 of gelijk aan 0. Zo ja, overschrijf. Daarna: Vervang in Word
if (sum(data_ZJHM_digitaal$Frequentie_3uurInternet == 1, na.rm = TRUE) < 5) {perc_internet <- "TE LAGE N"} 
if (sum(data_ZJHM_digitaal$Frequentie_3uurInternet == 1, na.rm = TRUE) == 0) {perc_internet <- "N = 0"} 
body_replace_all_text(Rapportage_template, "perc_internet", as.character(perc_internet), fixed=TRUE, only_at_cursor = FALSE)

### Percentage dat dagelijks minimaal 3 uur in de vrije tijd besteed aan gamen
perc_gamen_1 <- (sum(data_ZJHM_digitaal$Frequentie_3uurgamen == 1, na.rm = TRUE) / sum(data_ZJHM_digitaal$Frequentie_3uurgamen %in% c(0, 1), na.rm = TRUE)) *100 
perc_gamen <- round2(perc_gamen_1) # rond af 
# Check of N < 5 of gelijk aan 0. Zo ja, overschrijf. Daarna: Vervang in Word
if (sum(data_ZJHM_digitaal$Frequentie_3uurgamen == 1, na.rm = TRUE) < 5) {perc_gamen <- "TE LAGE N"} 
if (sum(data_ZJHM_digitaal$Frequentie_3uurgamen == 1, na.rm = TRUE) == 0) {perc_gamen <- "N = 0"} 
body_replace_all_text(Rapportage_template, "perc_gamen", as.character(perc_gamen), fixed=TRUE, only_at_cursor = FALSE)


### Percentage dat vindt dat er genoeg tijd overblijft
perc_genoegtijd_1 <- (sum(data_ZJHM_digitaal$Genoegtijdnaastinternetengamen2CAT == 2, na.rm = TRUE) / sum(data_ZJHM_digitaal$Genoegtijdnaastinternetengamen2CAT %in% c(1, 2), na.rm = TRUE)) *100 
perc_genoegtijd <- round2(perc_genoegtijd_1) # rond af 
# Check of N < 5 of gelijk aan 0. Zo ja, overschrijf. Daarna: Vervang in Word
if (sum(data_ZJHM_digitaal$Genoegtijdnaastinternetengamen2CAT == 2, na.rm = TRUE) < 5) {perc_genoegtijd <- "TE LAGE N"} 
if (sum(data_ZJHM_digitaal$Genoegtijdnaastinternetengamen2CAT == 2, na.rm = TRUE) == 0) {perc_genoegtijd <- "N = 0"} 
body_replace_all_text(Rapportage_template, "perc_genoegtijd", as.character(perc_genoegtijd), fixed=TRUE, only_at_cursor = FALSE)



### Percentage dat op tijd kan stoppen 
perc_stoppen_1 <- (sum(data_ZJHM_digitaal$Optijdstoppen2CAT == 2, na.rm = TRUE) / sum(data_ZJHM_digitaal$Optijdstoppen2CAT %in% c(1, 2), na.rm = TRUE)) *100 
perc_stoppen <- round2(perc_stoppen_1) # rond af 
# Check of N < 5 of gelijk aan 0. Zo ja, overschrijf. Daarna: Vervang in Word
if (sum(data_ZJHM_digitaal$Optijdstoppen2CAT == 2, na.rm = TRUE) < 5) {perc_stoppen <- "TE LAGE N"} 
if (sum(data_ZJHM_digitaal$Optijdstoppen2CAT == 2, na.rm = TRUE) == 0) {perc_stoppen <- "N = 0"} 
body_replace_all_text(Rapportage_template, "perc_stoppen", as.character(perc_stoppen), fixed=TRUE, only_at_cursor = FALSE)


#################################### 
### Grafieken Hoofdstuk 5 deel 2 ###
#################################### 

### Maak dataframe voor plots. 
Percentages5_2 <- c(as.character(perc_internet),
                 as.character(perc_gamen), 
                 as.character(perc_genoegtijd),
                 as.character(perc_stoppen)) 

Labels5_2 <- c("Minstens 3 uur per dag op telefoon/internet",
            "Gamet minstens 3 uur per dag",
            "Heeft hierdoor niet genoeg tijd voor andere dingen",
            "Slaapt hierdoor minder uren")

# Voeg alle percentages en de labels van de vraag samen in 1 dataframe
df_InternetGamen <- data.frame(Sortering = as.character(1:length(Percentages5_2)), Percentage = Percentages5_2, Label = Labels5_2)
# Verwijder alle percentages waarvan de N te laag was
df_InternetGamen <- df_InternetGamen[!df_InternetGamen$Percentage %in% c("TE LAGE N", "N = 0", "niemand", "Niemand", "geen", "Geen"),]
# Percentages terug van tekst naar nummers
df_InternetGamen$Percentage <- as.numeric(df_InternetGamen$Percentage)

# Zet sommige percentages om naar 1 - percentage.
for (i in 1:nrow(df_InternetGamen)) {
  if (df_InternetGamen[i,3] %in% c("Heeft hierdoor niet genoeg tijd voor andere dingen",
                                   "Slaapt hierdoor minder uren")) {
    df_InternetGamen[i,2] <- 100 - df_InternetGamen[i,2]
  }
}

### Maak plots
grafiek_InternetGamen <- maak_grafiek(df_InternetGamen)

### Zet in Word
cursor_reach(Rapportage_template, keyword = "Grafiekgebied_InternetGamen")
body_remove(Rapportage_template)
body_add_gg(Rapportage_template, grafiek_InternetGamen, width = 7, height = 4)



#############################
# Hoofdstuk 6: Seksualiteit #
#############################

# perc_seksgehad
# perc_condoom
# perc_ongewenst

cursor_reach(Rapportage_template, keyword = "perc_seksgehad")

### Percentage dat weleens seks heeft gehad
perc_seksgehad_1 <- (sum(data_ZJHM_digitaal$Frequentie_Seksgehad == 1, na.rm = TRUE) / sum(data_ZJHM_digitaal$Frequentie_Seksgehad %in% c(0, 1, 88), na.rm = TRUE)) *100 
perc_seksgehad <- round2(perc_seksgehad_1) # rond af 
# Check of N < 5 of gelijk aan 0. Zo ja, overschrijf. Daarna: Vervang in Word
if (sum(data_ZJHM_digitaal$Frequentie_Seksgehad == 1, na.rm = TRUE) < 5) {perc_seksgehad <- "TE LAGE N"} 
if (sum(data_ZJHM_digitaal$Frequentie_Seksgehad == 1, na.rm = TRUE) == 0) {perc_seksgehad <- "N = 0"} 
body_replace_all_text(Rapportage_template, "perc_seksgehad", as.character(perc_seksgehad), fixed=TRUE, only_at_cursor = FALSE)


### Percentage dat altijd een condoom gebruikt (noemer nu niet alle leerlingen, maar alleen jongeren die seksueel actief zijn)
perc_condoom_1 <- (sum(data_ZJHM_digitaal$Frequentie_Condoomgebruik == 1, na.rm = TRUE) / sum(data_ZJHM_digitaal$Frequentie_Seksgehad == 1, na.rm = TRUE)) *100 
perc_condoom <- round2(perc_condoom_1) # rond af 
# Check of N < 5 of gelijk aan 0. Zo ja, overschrijf. Daarna: Vervang in Word
if (sum(data_ZJHM_digitaal$Frequentie_Condoomgebruik == 1, na.rm = TRUE) < 5) {perc_condoom <- "TE LAGE N"} 
if (sum(data_ZJHM_digitaal$Frequentie_Condoomgebruik == 1, na.rm = TRUE) == 0) {perc_condoom <- "N = 0"} 
# Niet meer in de tekst van de template, is vervangen door percentage onveilige seks. Variabelen nog wel nodig om grafiek te kunnen maken.
# body_replace_all_text(Rapportage_template, "perc_condoom", as.character(perc_condoom), fixed=TRUE, only_at_cursor = FALSE)

### Percentage dat onveilige seks heeft (noemer nu niet alle leerlingen, maar alleen jongeren die seksueel actief zijn)
perc_seksonveilig_1 <- (sum(data_ZJHM_digitaal$Frequentie_Condoomgebruik %in% c(2, 3, 4), na.rm = TRUE) / sum(data_ZJHM_digitaal$Frequentie_Seksgehad == 1, na.rm = TRUE)) *100 
perc_seksonveilig <- round2(perc_seksonveilig_1) # rond af 
# Check of N < 5 of gelijk aan 0. Zo ja, overschrijf. Daarna: Vervang in Word
if (sum(data_ZJHM_digitaal$Frequentie_Condoomgebruik %in% c(2, 3, 4), na.rm = TRUE) < 5) {perc_seksonveilig <- "TE LAGE N"} 
if (sum(data_ZJHM_digitaal$Frequentie_Condoomgebruik %in% c(2, 3, 4), na.rm = TRUE) == 0) {perc_seksonveilig <- "N = 0"} 
body_replace_all_text(Rapportage_template, "perc_seksonveilig", as.character(perc_seksonveilig), fixed=TRUE, only_at_cursor = FALSE)


### Percentage dat wel eens tegen de wil in op een intieme manier is aangeraakt of hiertoe is gedwongen
perc_ongewenst_1 <- (sum(data_ZJHM_digitaal$Frequentie_OngewenstIntiem == 1, na.rm = TRUE) / sum(data_ZJHM_digitaal$Frequentie_OngewenstIntiem %in% c(0, 1, 2), na.rm = TRUE)) *100 
perc_ongewenst <- round2(perc_ongewenst_1) # rond af 
# Check of N < 5 of gelijk aan 0. Zo ja, overschrijf. Daarna: Vervang in Word
if (sum(data_ZJHM_digitaal$Frequentie_OngewenstIntiem == 1, na.rm = TRUE) < 5) {perc_ongewenst <- "TE LAGE N"} 
if (sum(data_ZJHM_digitaal$Frequentie_OngewenstIntiem == 1, na.rm = TRUE) == 0) {perc_ongewenst <- "N = 0"} 
body_replace_all_text(Rapportage_template, "perc_ongewenst", as.character(perc_ongewenst), fixed=TRUE, only_at_cursor = FALSE)

############################# 
### Grafieken Hoofdstuk 6 ###
############################# 

### Maak dataframe voor plots. 
Percentages6 <- c(as.character(perc_seksgehad),
                    as.character(perc_condoom), 
                    as.character(perc_ongewenst)) 

Labels6 <- c("Heeft geslachtsgemeenschap gehad",
               "Gebruikt altijd een condoom",
               "Heeft weleens een ongewenste seksuele ervaring gehad")

# Voeg alle percentages en de labels van de vraag samen in 1 dataframe
df_Seksualiteit <- data.frame(Sortering = as.character(1:length(Percentages6)), Percentage = Percentages6, Label = Labels6)
# Verwijder alle percentages waarvan de N te laag was
df_Seksualiteit <- df_Seksualiteit[!df_Seksualiteit$Percentage %in% c("TE LAGE N", "N = 0", "niemand", "Niemand", "geen", "Geen"),]
# Percentages terug van tekst naar nummers
df_Seksualiteit$Percentage <- as.numeric(df_Seksualiteit$Percentage)

### Maak plots
grafiek_Seksualiteit <- maak_grafiek(df_Seksualiteit)

### Zet in Word
cursor_reach(Rapportage_template, keyword = "Grafiekgebied_Seksualiteit")
body_remove(Rapportage_template)
body_add_gg(Rapportage_template, grafiek_Seksualiteit, width = 7, height = 4)



#################################
# Hoofdstuk 7: Oproepindicaties #
#################################

# oproep_percentage
# meest_oproep
# top_5_oproep
# perc_gesprek
# perc_chat
# perc_info

## Aantal dat is opgeroepen
aant_oproep <- sum(data_ZJHM_digitaal$Oproepindicatie == 1, na.rm = TRUE)


cursor_reach(Rapportage_template, keyword = "aant_oproep")
body_replace_all_text(Rapportage_template, "aant_oproep", as.character(aant_oproep), fixed=TRUE, only_at_cursor = FALSE)

################
## Meest opgeroepen voor:

# Eerst Dossiernummer naar numeric.
data_ZJHM_digitaal$Dossiernummer <- as.numeric(data_ZJHM_digitaal$DossierNummer)
# MHI naar 2cat met oproepgrens van < 60 .
data_ZJHM_digitaal$MHI_Oproep[data_ZJHM_digitaal$MHIscore < 60] <- "1"

# Herstructureer dataframe naar 1 kolom met vragen en 1 kolom met antwoorden (+ Dossiernummer, voor zekerheid) (long in naam omdat het nu in "long format" staat)
data_long <- data_ZJHM_digitaal %>%
  select(Dossiernummer, MBGSK321, Contact_gesch_ouders, Contact_lich_gezondh,
         Info_ongesteld, Vragen_lich_relaties_seks, Vragen_soa_zwanger, Vragen_gevoelens, Frequentie_OngewenstIntiem,
         Seksvoorgeld, Hulp_gedrag_leerproblemen, Geestelijk_mishandeld, Lichamelijk_mishandeld,
         MHI_Oproep, Automutilatie, Zelfmoordgedachten, Contact_lichaam_zelfbeeld, Contact_sfeer_thuis,
         Vragen_genotmiddelen, Hulp_ingrijp_gebeurt, ZorgenThuissituatie, Overige_vragen) %>%
  mutate_all(funs(as.character)) %>%
  gather(vraag, antwoord, MBGSK321:Overige_vragen, factor_key=TRUE)

# Voeg vraagtekst of uitleg toe bij de variabelenamen
data_long$Label[data_long$vraag == "MBGSK321"] <- "Bij wie woon je de meeste dagen van de week?"
data_long$Label[data_long$vraag == "Contact_gesch_ouders"] <- "Wil je hulp of informatie over omgaan met gescheiden ouders?"
data_long$Label[data_long$vraag == "Contact_lich_gezondh"] <- "Heb je zorgen of vragen over je lichamelijke gezondheid? Wil je hierover met ons in gesprek?"
data_long$Label[data_long$vraag == "Info_ongesteld"] <- "Wil je informatie over ongesteldheid?"
data_long$Label[data_long$vraag == "Vragen_lich_relaties_seks"] <- "Ik heb vragen over mijn lichaam, liefde, relaties en/of seks"
data_long$Label[data_long$vraag == "Vragen_soa_zwanger"] <- "Ik heb vragen over geslachtsziektes (soa's), voorbehoedsmiddelen, zwangerschap of zwangerschapstest"
data_long$Label[data_long$vraag == "Vragen_gevoelens"] <- "Ik heb vragen over seksuele gevoelens of gedachten (zoals bijvoorbeeld homo/bi)"
data_long$Label[data_long$vraag == "Frequentie_OngewenstIntiem"] <- "Heeft iemand wel eens tegen je wil in op een intieme manier aangeraakt of je hiertoe gedwongen?"
data_long$Label[data_long$vraag == "Seksvoorgeld"] <- "Ben je wel eens verder gegaan dan zoenen voor geld of andere dingen (bijvoorbeeld cadeautjes, drankjes of een slaapplaats)?"
data_long$Label[data_long$vraag == "Hulp_gedrag_leerproblemen"] <- "Heb je een aandoening, ziekte , leer- en/of gedragsproblemen waardoor jij niet je huiswerk kunt doen of minder vaak op school bent? Wil je hulp?"
data_long$Label[data_long$vraag == "Geestelijk_mishandeld"] <- "Ik word genegeerd, uitgescholden of buitengesloten thuis, op de sportschool, WhatsApp of op school en heb daar last van"
data_long$Label[data_long$vraag == "Lichamelijk_mishandeld"] <- "Ik word geschopt, geslagen of op een andere manier mishandeld (en heb daar nog last van).Welk antwoord past het beste bij jou?"
data_long$Label[data_long$vraag == "MHI_Oproep"] <- "MHI-5 score"
data_long$Label[data_long$vraag == "Automutilatie"] <- "Heb je jezelf in de afgelopen 6 maanden wel eens expres pijn gedaan (bijvoorbeeld snijden, krassen, branden, krabben, bijten of stoten)?"
data_long$Label[data_long$vraag == "Zelfmoordgedachten"] <- "Heb je in de laatste 12 maanden er wel eens serieus over gedacht een eind te maken aan je leven?"
data_long$Label[data_long$vraag == "Contact_lichaam_zelfbeeld"] <- "Heb je vragen, of wil je informatie over je lengte, gewicht of zelfbeeld?"
data_long$Label[data_long$vraag == "Contact_sfeer_thuis"] <- "Wil je met iemand over de sfeer thuis praten?"
data_long$Label[data_long$vraag == "Vragen_genotmiddelen"] <- "Heb je zorgen over of vragen voor de verpleegkundige/arts over alcohol, roken, drugs, internetgebruik, gamen, of schulden?"
data_long$Label[data_long$vraag == "Hulp_ingrijp_gebeurt"] <- "Welke gebeurtenissen heb je meegemaakt, waar je nu nog veel mee bezig bent? Wil je hulp?"
data_long$Label[data_long$vraag == "ZorgenThuissituatie"] <- "Door het zorgen voor iemand in mijn gezin (bijv. oppassen, huishouden of verzorgen van een gezinslid) of doordat ik mij zorgen maak om mijn thuissituatie, kom ik niet toe aan andere dingen zoals hobby's en vrienden"
data_long$Label[data_long$vraag == "Overige_vragen"] <- "Heb jij andere vragen, zorgen of problemen waarvoor je uitgenodigd wilt worden door de verpleegkundige of arts?"
  
# Flaggen welke combi van vraag en antwoord ertoe leidt dat jongere wordt opgeroepen.
data_long$Flag[data_long$vraag == "MBGSK321" & data_long$antwoord == "Ik woon op mezelf"] <- 1
data_long$Flag[data_long$vraag == "Contact_gesch_ouders" & data_long$antwoord == "Ja, ik wil een persoonlijk gesprek met een arts/verpleegkundige"] <- 1
data_long$Flag[data_long$vraag == "Contact_lich_gezondh" & data_long$antwoord == "Ja, ik wil dit graag persoonlijk bespreken met een verpleegkundige/arts"] <- 1
data_long$Flag[data_long$vraag == "Info_ongesteld" & data_long$antwoord == "Ik heb vragen voor de verpleegkundige/arts en wil dit persoonlijk bespreken"] <- 1
data_long$Flag[data_long$vraag == "Vragen_lich_relaties_seks" & data_long$antwoord == "Ik heb vragen voor de verpleegkundige/arts en wil dit persoonlijk bespreken"] <- 1
data_long$Flag[data_long$vraag == "Vragen_soa_zwanger" & data_long$antwoord == "Ik heb vragen voor de verpleegkundige/arts en wil dit persoonlijk bespreken"] <- 1
data_long$Flag[data_long$vraag == "Vragen_gevoelens" & data_long$antwoord == "Ik wil dit graag persoonlijk bespreken met een verpleegkundige/arts"] <- 1
data_long$Flag[data_long$vraag == "Frequentie_OngewenstIntiem" & data_long$antwoord == "Ja"] <- 1
data_long$Flag[data_long$vraag == "Seksvoorgeld" & data_long$antwoord == "Regelmatig"] <- 1
data_long$Flag[data_long$vraag == "Hulp_gedrag_leerproblemen" & data_long$antwoord == "Ja, ik wil een persoonlijk gesprek met een jeugdarts of jeugdverpleegkundige"] <- 1
data_long$Flag[data_long$vraag == "Geestelijk_mishandeld" & data_long$antwoord == "Vaak"] <- 1
data_long$Flag[data_long$vraag == "Lichamelijk_mishandeld" & data_long$antwoord == "Vaak"] <- 1
data_long$Flag[data_long$vraag == "MHI_Oproep" & data_long$antwoord == "1"] <- 1
data_long$Flag[data_long$vraag == "Automutilatie" & data_long$antwoord == "Vaak"] <- 1
data_long$Flag[data_long$vraag == "Zelfmoordgedachten" & data_long$antwoord %in% c("Vaak", "Heel vaak")] <- 1
data_long$Flag[data_long$vraag == "Contact_lichaam_zelfbeeld " & data_long$antwoord == "Ja, ik wil een persoonlijk gesprek hierover met een verpleegkundige of arts"] <- 1
data_long$Flag[data_long$vraag == "Contact_sfeer_thuis " & data_long$antwoord == "Ik wil graag een persoonlijk gesprek met de verpleegkundige/arts"] <- 1
data_long$Flag[data_long$vraag == "Vragen_genotmiddelen " & data_long$antwoord == "Ja"] <- 1
data_long$Flag[data_long$vraag == "Hulp_ingrijp_gebeurt " & data_long$antwoord == "Ja, ik wil dit graag persoonlijk bespreken met een verpleegkundige/arts"] <- 1
data_long$Flag[data_long$vraag == "ZorgenThuissituatie" & data_long$antwoord %in% c("Meerdere keren per week", "(Bijna) Elke dag")] <- 1
data_long$Flag[data_long$vraag == "Overige_vragen " & data_long$antwoord == "Ja"] <- 1

# Gooi alle regels weg die geen oproepindicatie hebben
data_long <- data_long[data_long$Flag == 1 & is.na(data_long$Flag) == FALSE,]

# Tel op  hoe vaak elke indicatie voorkomt. 
oproep_aantalPerVraag <- data_long %>%
  group_by(vraag, Label) %>%
  count(vraag)

oproep_aantalPerVraag <- as.data.frame(oproep_aantalPerVraag)  
oproep_aantalPerVraag <- oproep_aantalPerVraag %>%
                          arrange(desc(n))

meeste_oproep <- oproep_aantalPerVraag[1,2] 

if (meeste_oproep == "MHI-5 score") {meeste_oproep = "de Mental Health Inventory-5 (MHI-5) vragenlijst. Deze vragenlijst screent op psychische problematiek, zoals angst en depressie. Alle jongeren die naar aanleiding van de vragenlijst zorgen oproepen over hun geestelijke gezondheid, worden uitgenodigd op gesprek"}

body_replace_all_text(Rapportage_template, "meeste_oproep", as.character(meeste_oproep), fixed=TRUE, only_at_cursor = FALSE)

##
########## Einde meeste_oproep

# top 5 oproep
top_5_oproep <- oproep_aantalPerVraag[1:5,2:3] # Alleen eerste 5 regels en eerste kolom weg


# Nog grafiek van maken...? Tabel? 
oproep_tabel <- flextable(top_5_oproep)

# Verander header namen
oproep_tabel <- set_header_labels(oproep_tabel, Label = "Oproepindicatie" )
oproep_tabel <- set_header_labels(oproep_tabel, n = "Aantal leerlingen" )
# Verander header kleuren
oproep_tabel <- bg(oproep_tabel, bg = "#002659", part = "header")
oproep_tabel <- color(oproep_tabel, color = "#FFFFFF", part = "header")
oproep_tabel <- color(oproep_tabel, color = "#002A5C", part = "body")
# Verander kleuren van sommige rijen
bg(oproep_tabel, i = c(2, 4), bg = "#DAEEF3", part = "body")
# Verander breedte van de kolommen 
# dim_pretty(myft)
# autofit(myft)
oproep_tabel <- width(oproep_tabel, width = 4)
# Pas alignment aan van tekst in de kolommen
oproep_tabel <- align(oproep_tabel, j = "Label", align = "left", part = "all" )
oproep_tabel <- align(oproep_tabel, j = "n", align = "center", part = "all" )


# Om in word bestand te zetten: body_add_flextable() 
cursor_reach(Rapportage_template, keyword = "Tabel 2. Meest voorkomende oproepindicaties. Een opgeroepen jongere kan meerdere oproepindicaties hebben.") 
# body_remove(my_doc)
body_add_flextable(Rapportage_template, oproep_tabel, pos = "after")


### Percentage dat heeft aangegeven een persoonlijk gesprek te willen met de jeugdverpleegkundige/arts
perc_gesprek_1 <- (sum(data_ZJHM_digitaal$Wil_gesprek == 1, na.rm = TRUE) / aant_ingevuld_online) *100 
perc_gesprek <- round2(perc_gesprek_1) # rond af 
# Check of N < 5 of gelijk aan 0. Zo ja, overschrijf. Daarna: Vervang in Word
if (sum(data_ZJHM_digitaal$Wil_gesprek == 1, na.rm = TRUE) < 5) {perc_gesprek <- "TE LAGE N"} 
if (sum(data_ZJHM_digitaal$Wil_gesprek == 1, na.rm = TRUE) == 0) {perc_gesprek <- "N = 0"} 
body_replace_all_text(Rapportage_template, "perc_gesprek", as.character(perc_gesprek), fixed=TRUE, only_at_cursor = FALSE)

### Percentage dat wil chatten via JouwGGD/Sense.info
perc_chat_1 <- (sum(data_ZJHM_digitaal$Wil_chatten %in% c(1, 2), na.rm = TRUE) / aant_ingevuld_online) *100 
perc_chat <- round2(perc_chat_1) # rond af 
# Check of N < 5 of gelijk aan 0. Zo ja, overschrijf. Daarna: Vervang in Word
if (sum(data_ZJHM_digitaal$Wil_chatten %in% c(1, 2), na.rm = TRUE) < 5) {perc_chat <- "TE LAGE N"} 
if (sum(data_ZJHM_digitaal$Wil_chatten %in% c(1, 2), na.rm = TRUE) == 0) {perc_chat <- "N = 0"} 
body_replace_all_text(Rapportage_template, "perc_chat", as.character(perc_chat), fixed=TRUE, only_at_cursor = FALSE)

### Percentage dat informatie heeft opgevraagd
perc_info_1 <- (sum(data_ZJHM_digitaal$Wil_informatie == 1, na.rm = TRUE) / aant_ingevuld_online) *100 
perc_info <- round2(perc_info_1) # rond af 
# Check of N < 5 of gelijk aan 0. Zo ja, overschrijf. Daarna: Vervang in Word
if (sum(data_ZJHM_digitaal$Wil_informatie == 1, na.rm = TRUE) < 5) {perc_info <- "TE LAGE N"} 
if (sum(data_ZJHM_digitaal$Wil_informatie == 1, na.rm = TRUE) == 0) {perc_info <- "N = 0"} 
body_replace_all_text(Rapportage_template, "perc_info", as.character(perc_info), fixed=TRUE, only_at_cursor = FALSE)




##########################
# Sla het wordbestand op #
##########################

print(Rapportage_template, target = "Rapportage_Zuiderzee_klas1.docx")



