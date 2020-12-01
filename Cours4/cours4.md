---
title: Portées
author: Hugo Leblanc
date: Cours 4
pandocomatic_:
  use-template: 
  - presentation
  - presentation-handout
...

Sous-programmes
===============

Durée de vie des variables
--------------------------

-   Tout ce qui se passe à l'intérieur des fonctions est détruit après
    l'appel de la fonction.

-   Toutes déclaration de variables à l'intérieur d'une fonction est
    détruite après l'appel de la fonction.

-   Seul le retour est renvoyé.

-   Le mot clé [Static]{.alert} permet de conserver un variable vivante
    entre ses appels de sous-programmes.

Portées des éléments
--------------------

-   La portée d'un élément indique sa visibilité par rapport aux autres
    modules.

-   Si un élément est visible, il peut être appelé (utilisé) à
    l'intérieur de sa portée.

-   La portée d'une procédure inclue ses déclarations de variables avec
    [Dim]{.alert}.

-   La portée privée ([Private]{.alert}) d'un module inclue ses
    déclaration de variables et de sous-programme privées.

Portées (suite)
-----------------------

-   La portée publique ([Public]{.alert}) d'un module inclue toutes les
    variables et sous-programmes publiques d'un projet (classeur).

-   Les variables de portée globale sont [INTERDITES]{.alert} dans le
    cours, sauf avis contraire.

Passage par référence - ByRef
-----------------------------

Passage par référence - ByRef

-   Le passage par référence reçoit une variable à la place d'une valeur
    comme dans le passage par valeur.

-   La référence est donc lié entre l'appellant du sous-programme et le
    sous-programme lui-même.

-   Une modification d'une variable passé en référence restera modifié
    après l'exécution du sous-programme.

-   Les noms entre le paramètre passée et le nom du paramètre n'a pas
    d'importance.

Validation
----------

-   La validation permet de regarder et valider des informations avant
    de continuer dans un programme

-   La validation est surtout utilisée quand nous recevons des
    informations de l'extérieur (avec InputBox)

-   La validation permet de générer une erreur avant d'avoir des
    résultats erronés.

-   Le mot-clé [End]{.alert} permet de terminer l'exécution de notre
    programme prématurément.

Test
----

-   Le test d'une fonction permet de s'assurer de son comportement avant
    de passer à la création d'élément plus complexes.

-   Les tests peuvent être dynamique avec des saisis et affichages ou
    statique par rapport à des valeurs et réponse déjà connus.

-   Pour de grand projet, les tests statiques sont utilisé pour
    s'assurer du bon fonctionnement au fur et à mesure de la conception.

Exercices
-----------

-   Écrivez une fonction qui retourne le nombre de fois que la fonction
    à été appellé. Utilisez le mot-clé Static.

-   Écrivez une procédure qui reçoit une variable par référence et un
    nombre. La procédure additionne le nombre à la variable donnée.

-   Écrivez une procédure qui reçoit une valeur entière et arrète
    l'exécution du programme si celle-ci n'est pas positive.
