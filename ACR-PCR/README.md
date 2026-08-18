# ACR / PCR — vérification des masses maximales

Petit exe autonome : on saisit le **PCR publié** d'une piste (ex. `690/R/B/W/T`) et il affiche,
pour **B737-400, B737-800, B747-400F et B747-400ERF**, leur ACR et la **masse maximale admissible**
sur cette chaussée.

Depuis le 28 nov. 2024 l'OACI a remplacé ACN/PCN par ACR/PCR. La règle d'exploitation est
`ACR ≤ PCR`. La FAA publie le moteur de calcul officiel mais **sans interface** — d'où cet outil,
qui est le front-end manquant.

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

## Surcharge

L'outil affiche `ACR ≤ PCR` et rien de plus. Exploiter au-delà du PCR publié est une décision de
l'exploitant d'aérodrome, pas un calcul : **aucune tolérance de surcharge n'est appliquée**.
Aide à la planification, à recouper avec l'AIP.

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
