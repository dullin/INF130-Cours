---
title: Types définis
author: Hugo Leblanc
date: Cours 6
pandocomatic_:
  use-template: 
  - presentation
  - presentation-handout
...

Types définis
============

Description
-----------

Description

-   Un type défini est une nouvelle structure de données (comme les
    tableaux ou les chaines de caractères).

-   Sous un seul identificateur, des champs (avec des noms de notre
    propre crus) vont chacun contenir des valeurs.

Création
--------

-   La déclaration d'un type défini se fait à l'extérieur des
    sous-programmes. Le type sera disponible pour tous les
    sous-programmes.

-   Chaque champs aura un type spécifique.

-   Les champs peuvent aussi être des tableaux.

~~~VB
Type mon_type
    champ1 As Integer
    champ2() As Double
    champ3 As String
End Type
~~~

Utilisation des types définis
-----------------------------

-   On utilise un type défini comme un autre type connu (Integer,
    String, etc.).

-   On doit le déclaré pour en faire sont utilisation.

-   L'accès au champs se fait avec l'opérateur point « . ».

~~~VB
Sub test_type()
    Dim type1 as mon_type

    type1.champ1 = 5
    ReDim type1.champ2(5)
    type1.champ2(4) = 4.5
    type1.champ3 = "allo"
End Sub
~~~
