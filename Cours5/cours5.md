---
title: Chaines
author: Hugo Leblanc
date: Cours 5
pandocomatic_:
  use-template: 
  - presentation
  - presentation-handout
...

Chaine de caractères
====================

Table ASCII
-----------

-   Chaques caractères utillisées dans une chaine est représenté dans la
    table de conversion ASCII.

-   Deux fonctions de VBA nous permets de faire les conversion d'une à
    l'autre:

    -   Chr retourne le caractères associé à un code ASCII.

    -   Asc retourne le code ASCII associé à un caractère donné.

Exercices
----------

-   Écrivez une fonction qui convertit un caractère représentant une
    lettre minuscule en lettre majuscule.

Opérateurs sur le chaines
-------------------------


-   [=]{.alert} permet l'assignation d'une chaine

-   Opérateurs de comparaisons permettent de faire une comparaison par
    rapport à la valeur de la table ASCII entre deux chaines.

-   [&]{.alert} permet la concaténation entre deux valeurs qui donne une
    chaine

-   [+]{.alert} permet la concaténation qui donne une valeur numérique
    si une des deux expressions est numérique.

Exercices
----------

-   Écrivez une fonction qui retourne True si un caractère représente un
    chiffre et False dans tous les autres cas.

-   Écrivez une fonction qui retourne True si un caractère représente
    une lettre majuscule et False dans tous les autres cas.

-   Écrivez une fonction qui retourne True si un caractère représente
    une lettre minuscule et False dans tous les autres cas.

-   Écrivez une fonction qui retourne True si un caractère représente
    une lettre et False dans tous les autres cas.

Exercices
---------

-   Écrivez une fonction nommée lpad qui reçoit une chaîne de caractères
    et ajoute n blancs au début de la chaîne. À titre d'exemple, l'appel
    lpad(\"allo\", 3) retourne \" allo\".

Fonctions sur les chaines
-------------------------

-   [Len]{.alert} -- Donne le nombres de caractères dans la chaine.

-   [UCase]{.alert} / [LCase]{.alert} -- Convertie en majuscules ou
    minuscule un chaine.

-   [Left]{.alert} / [Right]{.alert} -- Extrait une partie de la chaine
    à partir de la gauche ou la droite.

-   [Mid]{.alert} -- Extrait une partie de la chaine à partir d'un
    caractère définie.

-   [Trim]{.alert} -- Enlève les blancs à gauche et à droite de la
    chaine.

-   [InStr]{.alert} -- Cherche une chaine dans une autre.

Exercices
---------

-   Écrivez une fonction nommée ltrim qui reçoit une chaîne de
    caractères et retourne celle-ci sans les blancs se trouvant au
    début.

-   Écrivez une fonction qui reçoit une chaîne de caractères et un
    caractère. Elle retourne le nombre d'occurrences de ce caractère
    dans la chaîne (le nombre de fois que ce caractère se retrouve dans
    la chaîne).

Exercices
---------

-   Écrivez une fonction qui reçoit deux chaînes de caractères; la
    première contient une phrase alors que la seconde contient une liste
    de caractères à conserver. La fonction parcourt la phrase et à
    chaque fois qu'elle trouve un caractère qui n'est pas dans la
    seconde chaîne et qui n'est pas un blanc, elle le remplace par une
    étoile. Elle retourne la phrase obtenue. À titre d'exemple, l'appel
    phrase\_censuree(\"vive le vent\", \"eit \") retourne la chaîne
    \"\*i\*e \*e \*e\*t\".
