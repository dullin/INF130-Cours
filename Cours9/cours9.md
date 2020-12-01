---
title: Objets d'Excel
author: Hugo Leblanc
date: Cours 8
pandocomatic_:
  use-template: 
  - presentation
  - presentation-handout
...

Evironnement d'Excel
====================


-   L'environnement d'Excel est un hiérarichie d'objets.

-   Au plus haut niveau, nous avons le classeur (.xlsm).

-   Le classeur contient des feuilles de calculs, des modules et des
    formulaires.

-   Chaque feuille de calculs contiennent des cellules.

-   Une plage est une collection de plusieurs cellules.

Contrôles de formulaire
=======================


-   On doit avoir accès à l'onglet Développeur (on l'ajoute dans les
    options d'Excel) pour avoir accès au contrôles.

-   Les contrôles peuvents être ajoutés à des feuilles de calcul ou des
    formulaires

Contrôles de formulaire
========================

-   Les cases à cocher permettent d'indiquer si une option est activée
    au moyen d'un crochet.

-   Les cases d'options permettent de choisir une option parmi un groupe
    d'options mutuellement exclusives. Pour les utiliser, il faut tout
    d'abord ajouter un groupe d'options.


Contrôles de formulaire
========================

-   Les compteurs et les barres de défilement permettent d'augmenter ou
    de réduire une valeur affichée.

-   Les zones de liste et les zones combinées déroulantes permettent de
    choisir un élément dans une liste.

-   Les boutons permettent l'exécution de macros ou de programmes VBA.

Les objets d'Excel
==================


-   Un objet est un entité qui contient à la fois des propriétés
    (variables) et des méthodes (sous-programmes).

-   Des centaines d'objets sont diponible dans Excel mais les plus
    important sont les suivants:

    -   Application -- l'application d'Excel (et tous les classeurs
        ouvert)

    -   Workbook -- Un classeur avec ses feuilles

    -   Worksheet -- Une feuille de calcul

    -   Range -- Une sélection de cellule d'une feuille

-   Les objets peuvent être utilisé en tant que paramètre de
    sous-programmes.

Accéder au contenu d'une feuille
=================================

-   Il est possible d'accéder au contenu des cellules dans une feuille à
    l'aide des propriétés Range et Cells.

-   Cells(ligne ,colonne) permet de faire réréfence à une cellule
    particulière.

Exemple d'accès avec range
==========================
-   Range permet de sélectionner un plage de cellule voici quelque
    exemple d'utilisation:

    -   Range(\"A1\")

    -   Range(\"A1:E5\")

    -   Ragne(\"CELLULE\_NOMME\")

    -   Range(Cells(2, 2), Cells(10, 5))
