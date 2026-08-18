ICAO-ACN 1.0 - methode legacy ACN/PCN (FAA)
https://www.airporttech.tc.faa.gov/Products/Airport-Safety-Papers-Publications/Airport-Safety-Detail/icao-acn-10

  ICAO-ACN.exe               le programme (GUI WPF, .NET 4.0 Client Profile)
  ICAO-ACN.chm               sa documentation - a ouvrir sous Windows
  ICAO-ACN-installer.zip     le paquet complet tel que telecharge (setup.exe + MSI)

ATTENTION - ce n'est PAS l'equivalent d'ICAO-ACR.

ICAO-ACR livrait : une DLL appelable (ACRClassLib.dll), une doc d'API decrivant
ses signatures, le code source complet, et un cas de test chiffre.

ICAO-ACN ne livre QUE le programme graphique et son aide. Pas de bibliotheque,
pas de doc d'API, pas de source. Il n'y a donc aucune interface supportee pour
appeler le calcul ACN depuis un autre programme.

Methode implementee par ce programme (d'apres la page FAA) :
  - ACN souple : methode CBR, valeurs standard CBR 15 / 10 / 6 / 3  (A/B/C/D)
  - ACN rigide : methode PCA, valeurs standard k 150 / 80 / 40 / 20 MN/m3 (A/B/C/D)
  - facteurs alpha selon la lettre d'Etat OACI du 16 octobre 2007
  - base sur COMFAA 3.0 (AC 150/5335-5C) ; portage WPF de COMFAA 3.1.1
  - "n'est pas une norme officielle FAA ou OACI a ce jour" (mention FAA)

A noter aussi :
  - ACRClassLib.dll ne calcule PAS l'ACN. Ses fonctions CalculateACN* sont des
    noms herites qui appellent toutes CalculateACR (voir YYY.vb / YYY2.vb).
  - aircraft.xml ne contient aucune donnee ACN (que des champs ACR).
  - Le programme sait exporter sa bibliotheque avions vers
    "ICAO-ACN Aircraft All.txt" / "ICAO-ACN Aircraft Data.txt".
