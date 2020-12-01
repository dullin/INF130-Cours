---
title: Fonctions
author: Hugo Leblanc
date: Cours 3
pandocomatic_:
  use-template: 
  - presentation
  - presentation-handout
...

Sous-programmes
===============

Sous-programmes
---------------

Sous-programmes

-   Un sous-programme est un bloc de code réalisant une tâche précise.

-   Il existe deux type de sous-programmes:

    -   Prodédures (Sub)

    -   Fonctions (Function)

-   La différence entre une procédure et une fonction est qu'une
    fonction retourne une valeur tandis qu'une procédure ne retourne
    aucune valeur.

Procédures
----------

Procédures

-   La procédure permet d'executé une série d'instruction répondant à
    une tache précise.

-   On peut appeler (exécuter) une procédure déjà existante avec
    l'utilisation d'un Call suivit du nom de la procédure.

-   La procédures peut recevoir des paramètres d'entrées. Nous allons
    voir plus loin comment créer des sous-programmes avec des
    paramètres.

Fonctions
---------

Fonctions

-   Les fonctions retourne une valeurs après son exécution.

-   Le type de valeur de retour doit être défini à l'énoncé de la
    fonction.

-   La valeur de retour doit être assigné durant l'exécution de la
    fonction.

-   La nom de la valeur de retour est le nom de la fonction.

-   Puisque la fonction retourne une valeur, elle peut être utilisé
    durant une expression.

##

~~~VB
Function retourne_5() As Integer
    retourne_5 = 5
End Function


Sub test_fcn()
    Dim x As Integer
    x = retourne_5() + 8
    Call MsgBox(x)
End Sub
~~~

Exercices
---------

-   Écrivez une fonction qui trouve l'aire d'un triangle à partir de sa
    base et sa hauteur. Saisir la base et la hauteur.

-   Écrivez une fonction qui détermine si un nombre est impaire.\
    Indice : utilisez Mod. Saisir le nombre à tester.

Paramètres
----------

-   Les sous-programmes peuvent recevoir des paramètres d'entrées.

-   Les paramètres sont les informations critiques dont les
    sous-programmes ont besoins pour bien fonctionner.

-   Chaque paramètre est typé et devient une variable utilisable durant
    l'exécution du sous-programmes.

-----------

~~~VB
Sub affiche_param(ByVal entree As Integer)
    Call MsgBox(entree)
End Sub

Function retourne_param(ByVal entree As Integer) As Integer
    retourne_param = entree
End Function
~~~

Exercices
---------

-   Écrivez une fonction qui trouve l'aire d'un triangle à partir de sa
    base et sa hauteur.

-   Écrivez une fonction qui détermine si un nombre est impaire.\
    Indice : utilisez Mod.

Règles des sous-programmes
==========================

Durée de vie des variables
--------------------------


-   Tout ce qui se passe à l'intérieur des fonctions est détruit après
    l'appel de la fonction.

-   Toutes déclaration de variables à l'intérieur d'une fonction est
    détruite après l'appel de la fonction.

-   Seul le retour est renvoyé.

Passage par valeur - ByVal
--------------------------


-   Les paramètres et les retours sont renommé pour la durée du
    sous-programme.

-   Seul leur valeurs seront transféré entre le sous-programme et
    l'appelant.

-   Les noms des paramètres n'ont aucune incidence.

-   L'ordre des paramètre est ce qui sera considéré.
