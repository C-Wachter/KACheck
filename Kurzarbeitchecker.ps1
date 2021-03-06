﻿function check_legal_entity ($ein) {
# Alle Gesellschaften außer AIT, AITs, UCaC, USaS aussortieren
# by Clemens Wachter 2020-05-11
Write-Host "1. Prüfung auf falsche Gesellschaften"
$raus = $ein | Where-Object {$_."Legal Entity" -ne "AIT"} | Where-Object {$_."Legal Entity" -ne "AITs"} | Where-Object {$_."Legal Entity" -ne "UCaC"} | Where-Object {$_."Legal Entity" -ne "USaS"} 
$raus | Format-Table -Property "DAS-ID","Name, Vorname","Legal Entity"
#return $raus
}

function check_Auszubildende ($ein) {
#Alles außer Angestellte und ÜT aussortieren
# by Clemens Wachter 2020-05-11
write-Host "2. Prüfung auf Azubis, Praktikanten etc."
$raus = $ein | Where-Object {$_."Mitarbeiterkreis" -ne "Angestellte"} | Where-Object {$_."Mitarbeiterkreis" -ne "ÜT"} 
$raus | Format-Table -Property "DAS-ID","Name, Vorname","Mitarbeiterkreis"
#return $raus
}

function check_ATZ ($ein) {
#Alle ATZ, ER aussortieren ohne Rücksicht auf Beginn (da Aufhebungsvertrag)
# by Clemens Wachter 2020-05-11
write-Host "3. Prüfung auf ATZ, ER, ER+ etc."
$raus = $ein | Where-Object {$_."Freistellungs-grund" -ne ""} | Where-Object {$_."Freistellungs-grund" -ne "Arbeitsphase TimeOut"} 
$raus | Format-Table -Property "DAS-ID","Name, Vorname","Freistellungs-grund"
#return $raus
}

function check_gekAV ($ein) {
#Alle MA mit Austrittsgrund und kein Normal Retirement (also gekündigte Arbeitsvertrag)
# by Clemens Wachter 2020-05-11
write-Host "4. Prüfung auf gekündigtes Arbeitsverhältnis"
$raus = $ein | Where-Object {$_."Austrittgrund" -ne ""} | Where-Object {$_."Austrittgrund".trim() -ne "Normal Retirement"} 
$raus | Format-Table -Property "DAS-ID","Name, Vorname","Austrittgrund"
#return $raus
}

function check_Trafo ($ein) {
#Prüfung auf Transformation
# by Clemens Wachter 2020-05-11
write-Host "5. Prüfung auf TRAFO"
$raus = $ein | Where-Object {$_."TRAFO Scope" -ne ""} 
$raus | Format-Table -Property "DAS-ID","Name, Vorname","TRAFO Scope"
#return $raus
}

function check_Begruendung ($ein) {
#Leere oder unzureichende Begründungen
# by Clemens Wachter 2020-05-11
write-Host "6. Fehlende oder unzureichende Begründungen"
$raus = $ein | Where-Object {
$_."Begründung" -eq "" -OR 
($_."Begründung").Contains("Slide Deck") -OR
($_."Begründung").Contains("Gleichmäßige Verteilung Arbeitslast") -OR
($_."Begründung").StartsWith("Unterauslastung")
} 
$raus | Format-Table -Property "DAS-ID","Name, Vorname","Begründung"
#return $raus
}

function check_negativ_personen ($ein) {
#Bekannte Personen, die KEINE KA leisten können
# by Clemens Wachter 2020-05-11
write-Host "7. Negativ-Liste"
$negativ=Import-Csv -Path C:\Users\A172169\Documents\BR\Kurzarbeit\negativ-personen.csv -Delimiter ";"
$treffer=Compare-Object -ReferenceObject $liste -DifferenceObject $negativ -Property 'DAS-ID' -Excludedifferent -IncludeEqual
$treffliste=@()
$treffer | ForEach-Object {
    $suche=$_."DAS-ID"
    $treffliste=$treffliste+ ($negativ | Where-Object {$_."DAS-ID" -eq $suche})
}
$treffliste | Format-Table -Property "DAS-ID","Name","Begruendung"
}



# Main
# Kurzarbeitchecker V0.1.1
# by Clemens Wachter 2020-05-11
# Konstanten definieren
$arbeitsverzeichnis="C:\temp\KA"
$quelle="liste_ein.csv"
$kurzarbeit="liste_KA.csv"
$negativperson="negativ-personen.csv"
$negativabteilung="negativ-abteilung.csv"
$positivperson="positiv-personen.csv"
$positivabteilung="positiv-abteilung.csv"

#***!!! $woche anpassen auf aktuelle Woche !!!***
$woche="KW20"

# Importiere die volle Liste als CSV
$listeein=Import-Csv -Path ($arbeitsverzeichnis + "\" + $quelle) -Delimiter ";" -Encoding UTF7

# Filtern nach Kurzarbeit in dieser Woche und exportieren
$listeein | Where-Object {$_.$woche -ne ""} | export-csv -Path ($arbeitsverzeichnis + "\" + $kurzarbeit ) -NoTypeInformation -Encoding UTF8

# Importiere die volle Liste als CSV
$listeka=Import-Csv -Path ($arbeitsverzeichnis + "\" + $kurzarbeit ) -Delimiter ","

#Checks
check_legal_entity $listeka
check_Auszubildende $listeka
check_ATZ $listeka
check_gekAV $listeka
check_Trafo $listeka
check_Begruendung $listeka
