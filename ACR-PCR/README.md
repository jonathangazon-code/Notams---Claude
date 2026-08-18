# ACR/PCR et ACN/PCN — vérification des masses maximales

Petit exe autonome : on saisit le code publié d'une piste et il affiche, pour **B737-400,
B737-800, B747-400F et B747-400ERF**, leur classification et la **masse maximale admissible**
sur cette chaussée.

La fenêtre est en deux moitiés, même présentation sur 4 lignes :

- **en haut, PCR** (ex. `690/R/B/W/T`) — méthode actuelle, en vigueur depuis le 28 nov. 2024
- **en bas, PCN** (ex. `80/R/B/W/T`) — méthode legacy, encore publiée par beaucoup d'AIP

Les deux moitiés sont la même classe (`RatingSection`) : elles ne peuvent pas diverger de
présentation. Ce qui change est derrière l'interface `IRatingEngine`.

La FAA publie le moteur de calcul ACR officiel mais **sans interface** — d'où cet outil.

**Sans rapport avec Dispatch Watch.** Application distincte, solution distincte : ne pas ajouter
`AcrTool.csproj` à `Dispatch Watch.sln`.

## Construire

Ouvrir `ACR-PCR.sln` dans SharpDevelop, **F8**. Un seul projet, `AcrTool` (C#), qui référence
`vendor/faa-lib/ACRClassLib.dll` — le moteur de calcul ACR de la FAA, en binaire.

`aircraft.xml` est copié à côté de l'exe à la compilation ; l'exe le lit au démarrage.

### Pourquoi le code VB de la FAA n'est pas compilé

Une première version référençait les projets VB `ACClassLib` / `ICAOModels` de la FAA. Ça ne
compile pas dans SharpDevelop, pour deux raisons indépendantes :

- ils sont en `ToolsVersion="15.0"` (Visual Studio 2017), que MSBuild 4.0 ne charge pas ;
- ils utilisent `NameOf` (36 occurrences), syntaxe VB14 que le compilateur VB fourni refuse.

Ils ne servent à rien ici : `AcrTool/AircraftLibrary.cs` lit `aircraft.xml` directement, et les
cinq valeurs utilisées sont lues telles quelles depuis ce même XML par `clsAC.InitACLib`, sans
aucune transformation (`clsAC.vb`, lignes 222-274) :

| valeur | élément XML |
|---|---|
| pression pneus (psi) | `Cp/us` |
| MTOW (lb) | `_GrossWeight/us` |
| part de masse sur une paire de bogies | `MgPercentPCN` |
| nombre de roues | nombre de `WheelCoordinates` |
| coordonnées de roues (pouces) | `WheelCoordinates X/Y us`, écrites en base 1 |

La géométrie dérivée de `modAC.vb` (points d'évaluation, cas particuliers par type de train)
n'est jamais utilisée par cet outil, donc rien n'en est porté. Le source VB reste dans
`vendor/faa-source-reference/` à titre de référence, hors build.

### Si la compilation échoue sur la version du framework

`ACRClassLib.dll` cible **.NET 4.6.1**, et le projet est aligné dessus — donc, contrairement à
Dispatch Watch, il ne peut pas être en 4.0. Si SharpDevelop dit ne pas trouver le framework
4.6.1, installer le **.NET Framework 4.8 Developer Pack** (gratuit, Microsoft) : il fournit les
assemblies de référence pour toutes les versions 4.x.

## Les deux méthodes ne sont pas interchangeables

Même forme de code à 5 composantes, mais des sens différents :

| | ACN/PCN (legacy) | ACR/PCR (depuis 28/11/2024) |
|---|---|---|
| échelle | — | ~10× l'ACN |
| sous-couche souple | CBR 15 / 10 / 6 / 3 | module E |
| sous-couche rigide | k = 150 / 80 / 40 / 20 MN/m³ | module E |
| surcharge, souple | +10 % | +10 % |
| surcharge, rigide | **+5 %** | **+10 %** |
| pneus W / X / Y / Z | illimité / 1,75 / 1,25 / 0,50 MPa | *identique* |

Les **catégories de pression pneus sont les mêmes dans les deux systèmes**. La comparaison
avant/après du 28 novembre 2024 publiée par l'OACI donne un tableau identique de chaque côté.
Une version antérieure de ce projet donnait 1,50 et 1,00 MPa pour X et Y en ACN : ce sont les
catégories **d'avant 2008**, révisées par l'OACI bien avant le passage à l'ACR, et fausses pour
tout PCN publié aujourd'hui — elles auraient refusé des avions que la chaussée accepte.

Ce qui diffère réellement, côté tolérance, c'est la **surcharge sur chaussée rigide** : l'Annexe 14
dit désormais « for flexible and rigid pavements ... ACR not exceeding 10 per cent above the
reported PCR », alors que la méthode legacy distinguait les deux (10 % souple, 5 % rigide).
`PavementCode.OverloadFactor()` applique la règle propre à chaque méthode.

## D'où viennent les chiffres ACN

**Ils ne sont pas calculés.** Contrairement à l'ACR, il n'existe aucune bibliothèque appelable
pour la méthode legacy : ICAO-ACN 1.0 ne livre qu'un programme graphique, sans API, sans source
et sans point d'entrée documenté (détails dans `vendor/faa-acn/README.txt`).

Les valeurs viennent des tables constructeur fournies dans **`ACN.xlsx`**, à la racine du projet.
`AcrTool/acn-data.xml` est **généré depuis ce classeur**, pas recopié à la main — si le classeur
change, il faut le régénérer.

- masses = colonne **F du classeur, en kg** (`weightKg`) ; elles coïncident avec les masses max
  que la moitié ACR lit dans la bibliothèque FAA **à moins d'une livre près** sur les quatre
  avions, donc les deux moitiés sont sur la même base
- les colonnes C et E du classeur (MTOW/MLW opérateur, équivalents en lb) ne servent pas au calcul
- les lignes « PCN » du classeur (L6, L15, L24, L33) sont des valeurs de test, pas des données
- les pressions pneus ne figurent pas dans le classeur : elles viennent de la bibliothèque FAA
  (`aircraft.xml`, `Cp/us`) — à confirmer contre la même source que les tables ACN

### La formule

Celle du classeur, à l'identique — interpolation linéaire en masse :

```
masse = masseMax − (ACNmax − PCN) / (ACNmax − ACNmin) × (masseMax − masseMin)
```

Vérifiée en rejouant les 32 combinaisons (4 avions × 8 sous-couches) contre le classeur :
**aucun écart**.

Deux différences volontaires, aux bords de la plage publiée, où le classeur extrapole :

- `PCN` supérieur à l'ACN à masse max → l'outil renvoie la masse max et affiche « pas de
  limitation chaussée », au lieu d'une masse extrapolée au-delà du certifié
- `PCN` inférieur à l'ACN à vide → l'outil affiche « inutilisable » : si l'avion ne passe pas
  même vide, il ne passe pas

L'inversion est exacte — on inverse la droite du bon segment, sans recherche, contrairement au
côté ACR où le moteur est une boîte noire et où il faut bisecter.

### Formules cassées dans le classeur

Les lignes de résultat du classeur (L8, L17, L26, L35) pointent vers des cellules vides : les
trois premiers blocs référencent `F40/F41`, `F49/F50`, `F58/F59` — 37 lignes trop bas, la feuille
s'arrêtant ligne 35 — et affichent donc **0** ; le quatrième référence `F13/F14`, c'est-à-dire les
masses du 737-800, et sort des valeurs fausses. Les bonnes références sont `F3/F4`, `F12/F13`,
`F21/F22`, `F30/F31`. L'outil applique la formule correcte ; le classeur, lui, reste à corriger.

Un avion avec moins de deux points publiés est traité comme « pas de donnée » : la moitié basse
le dit explicitement plutôt que d'afficher une réponse inventée.

## Démarrage et réactivité

Le premier jet mettait une vingtaine de secondes à ouvrir. Trois causes, toutes corrigées :

- **`aircraft.xml` était parcouru en entier** : `Descendants()` visitait les 87 737 éléments du
  fichier en testant les attributs de chacun, pour n'en retenir que 411. Remplacé par un parcours
  des seuls enfants de `<Airplanes>`.
- **Le fichier était parsé deux fois** : `Version()` rechargeait les 1,9 Mo juste pour lire un
  attribut de la racine. La version est maintenant capturée pendant le parsing unique.
- **Le solveur LEAF était appelé ~90 fois**, et chaque appel coûtait plus cher que nécessaire.
  Quatre changements :
  - `AcrEngine` **mémoïse** ses résultats par (avion, type de chaussée, masse) — la même masse
    était redemandée plusieurs fois par rendu, et à chaque changement d'unité ou de tolérance
  - la bisection est remplacée par une **fausse position** : l'ACR étant quasi linéaire en masse,
    interpoler entre les bornes tombe tout de suite près de la réponse. Mesuré contre une
    référence fine sur des courbes du linéaire au fortement convexe : **~8 résolutions au lieu
    de 21**, à précision égale (~50 lb). Le critère d'arrêt porte sur le déplacement de
    l'estimation, pas sur la largeur de l'encadrement, ce qui neutralise le blocage classique
    de la méthode ; l'encadrement est préservé et le nombre d'itérations plafonné.
  - le tableau **`SW()`** est désormais passé en chaussée souple (`StrainGrid`), comme le fait le
    driver FAA : il restreint la grille d'évaluation des contraintes aux roues dont la coordonnée
    latérale est au-dessus de la moyenne, soit un côté d'un train symétrique. La doc FAA est
    explicite — inclure toutes les roues « would take much longer » pour une différence de
    résultat « insignificant ». Ne rien passer, comme avant, revenait à payer ce prix.
  - le chemin rigide n'utilise pas `SW` : il charge déjà un seul bogie, et le driver FAA passe
    lui aussi la surcharge sans.

### Pourquoi les avions ne sont pas calculés en parallèle

`ACRClassLib` range son état de travail dans des **modules VB** (`gICAOCodeIndex`,
`gPavementType`, `gStrainTarget`…), qui sont statiques et partagés par toutes les instances.
La bibliothèque n'est donc pas réentrante : deux appels concurrents à `CalculateACR` se
corrompraient mutuellement, **sans erreur visible**. Répartir les quatre avions sur plusieurs
threads est l'optimisation évidente qu'il ne faut surtout pas faire ici.

Pour la même raison, toute entrée dans la bibliothèque passe par `AcrEngine.SyncRoot` : le calcul
tourne sur un worker pendant que la colonne « check a weight » de la grille peut en déclencher un
autre depuis le thread d'interface, et le self-test également.

Le calcul tourne en outre **hors du thread d'interface**, avec une barre de progression pendant
l'évaluation. Un clic pendant qu'un calcul tourne ne s'empile pas : le calcul en cours finit,
puis se relance une fois avec l'état courant des contrôles.

## Vérifier — à faire avant tout usage

**Bouton « Self-test ».** Il rejoue l'exemple chiffré publié par la FAA (doc *User Information for
ICAO-ACR*, §2.4) : train 2D-400, chaussée rigide, unités US, 400 000 lb, 47,5 % sur le train,
4 roues, 200 psi. Valeurs attendues :

```
subgrade D = 894.672058    C = 817.8752    B = 744.0264    A = 641.6398
```

Ce test valide d'un coup l'interop, l'ordre des arguments, les unités et l'indexation inversée des
sous-couches. **Tant qu'il ne passe pas, aucun chiffre affiché par l'outil ne veut dire quoi que ce soit.**

Ensuite, recoupement terrain : prendre une piste dont le PCR est connu et vérifier que la masse max
sortie est cohérente avec ce qui s'y opère réellement. Pour un contrôle indépendant, le programme
ICAO-ACR de la FAA (installeur `SetupACRClassDriver.msi` sur le site FAA) calcule les mêmes valeurs
avec sa propre interface.

## Comment le calcul est fait

Tout suit le programme FAA `ICAO-ACR` (`ACRClassDriver/Form1_ICAO.vb`), délibérément copié plutôt
que redérivé, pour que les résultats soient comparables à la référence :

- **Unités** : la bibliothèque travaille en US (`libGL` en lb, `libCP` en psi, `libTX/libTY` en
  pouces), donc `Metric = false` partout. La conversion en kg est uniquement à l'affichage.
- **Coordonnées de roues** : tableaux 1-based, indices 1..n, l'emplacement 0 reste inutilisé.
- **Sous-couches inversées** : `ACRdata.libACR()` va de 1 à 4 pour **D, C, B, A** — dans cet ordre.
- **Chaussée souple** : toutes les roues du train principal.
  - 737 : `wheels = libNWheels` (4), `percent_gw = libMGpcntPCN × 2`.
  - 747 : quatre bogies principaux (2 voilure + 2 fuselage), fournis comme **deux trains** à la
    surcharge à 13 paramètres, exactement comme le driver FAA le fait pour le 747 et l'A380.
- **Chaussée rigide** : le bogie le plus contraignant seulement — `wheels = libNWheels / 2`,
  `percent_gw = libMGpcntPCN`. Pour le 747, voilure et fuselage sont évalués tous les deux et
  l'ACR le plus élevé est retenu, plutôt que de deviner lequel est le plus contraignant.
- **Masse max** : par bisection sur la masse (l'ACR croît de façon monotone), plafonnée à la MTOW.

Le pourcentage par bogie n'est **pas** corrigé à la main : `SetPCN_for_AC` du driver a l'air de le
faire, mais il écrit `libMGpcnt` alors que tous les chemins ACR lisent `libMGpcntPCN`, et la ligne
qui écrirait `libMGpcntPCN` est commentée. La bibliothèque porte déjà la bonne valeur
(`MgPercentPCN = 0.2333` pour le 747-400, soit exactement 93,32/100/4).

### Les 747 cargo

`aircraft.xml` n'a pas d'entrée `-400F` / `-400ERF`. Les versions cargo partagent la géométrie de
train et la pression pneus des `-400` / `-400ER`, et les masses de la bibliothèque sont déjà celles
du cargo (877 000 lb et 913 000 lb). L'entrée utilisée est affichée dans la colonne
« Library entry » pour que la substitution soit visible et vérifiable, pas implicite.

## Pression des pneus

Contrôle **indépendant** de l'ACR : un avion peut passer en ACR et être refusé sur la pression.
La lettre du code PCR donne la limite — `W` illimité, `X` ≤ 1,75 MPa (254 psi), `Y` ≤ 1,25 MPa
(181 psi), `Z` ≤ 0,50 MPa (73 psi).

## Surcharge — case à cocher, jamais par défaut

Chaque moitié a une case **« Allow ICAO overload (+10% flex / +5% rigid) »**, décochée au
démarrage. Cochée, la limite comparée devient :

- **PCR × 1,10** en souple comme en rigide
- **PCN × 1,10** en souple, **× 1,05** en rigide ou composite

Rien n'est appliqué en silence :

- la ligne d'information passe en orange et affiche **les deux** valeurs — la limite effective
  utilisée *et* le code publié
- chaque verdict concerné est suffixé `(overload +10%)` ou `(overload +5%)`
- la colonne *Margin* est calculée contre la limite effective

Ces critères s'accompagnent de conditions que l'outil **ne peut pas vérifier** : les mouvements
en surcharge doivent rester de l'ordre de 5 % des mouvements annuels au plus, aucun n'est
acceptable sur une chaussée présentant des signes de dégradation ou en période de dégel, et
au-delà de la tolérance il faut une analyse de dommage cumulé (CDF). La décision reste celle de
l'exploitant d'aérodrome.

La tolérance peut faire basculer un cas de « inutilisable » à utilisable — par exemple le
B737-400 en souple sous-couche C avec PCN 17,6 : rien en strict, 35 454 kg avec la tolérance,
parce que la limite relevée repasse au-dessus de l'ACN à vide.

## Provenance

Fichiers sous `vendor/`, téléchargés depuis la FAA
([page ICAO-ACR](https://www.airporttech.tc.faa.gov/Products/Airport-Safety-Papers-Publications/Airport-Safety-Detail/ICAO-ACR-15)).
Travail du gouvernement des États-Unis, domaine public.

| Fichier | Provenance |
|---|---|
| `faa-lib/ACRClassLib.dll` | ICAO-ACR, build .NET 4.6.1 du 2020-12-14 |
| `faa-lib/aircraft.xml` | bibliothèque avions FAA, `LibraryVersion 1.2.4`, 411 appareils |
| `docs/User Information for ICAO-ACR.pdf` | doc d'API officielle (contient le cas de test) |
| `faa-source-reference/` | source VB, référence hors build (voir plus haut) |

`aircraft.xml` est copié à côté de l'exe à la compilation. Pour le mettre à jour, remplacer
`vendor/faa-lib/aircraft.xml` par la version courante et recompiler — la version chargée est
affichée en pied de fenêtre.
