---
title: Tableaux
author: Hugo Leblanc
date: Cours 6
pandocomatic_:
  use-template: 
  - presentation
  - presentation-handout
...

Tableaux
========

Description
-----------

-   Les tableaux sont des collections de plusieurs éléments du même type
    avec un seul identificateur.

-   Les tableaux nous permettent de travailler avec de grands
    échantillons de données.

-   Nous utiliserons la configuration Option Base 1 dans nos module pour
    que les tableau commence à la case 1 (et non 0).

Création de tableaux
--------------------

-   La déclaration de tableau se fait comme une variable mais en
    ajoutant des parenthèses après le nom avec le nombre de cases
    voulus.

~~~VB
Dim monTableau(5) As Integer
Dim AutreTableau(1 To 10) As Double
Dim DernierTableau(2 To 4) As String
~~~

-   La numérotation des cases commence à 1 par défaut (grâce à Option
    Base 1).

Tableaux de multiples dimensions
---------------------------------

-   On peut avoir des tableaux de plusieurs dimensions. On délimite
    chaque dimension avec un virgule.

-   Par convention, les dimensions d'une tableau 2D est représenté en
    (ligne,colonne).

~~~VB
Dim tableau2D(5, 5) As Integer
Dim AutreTableau2D(1 To 5, 2 To 4) As Double
~~~

Utilisation des cases du tableau
--------------------------------

-   On invoque la case d'un tableau en utilisant un indice entre
    parenthèses.

-   On utilise cette méthode pour assigner de nouvelle valeur ou aller
    consulter les valeurs déjà présentes.

-   Les tableaux de plusieurs dimensions demandes le bon nombres
    d'indice (1 pour chaque dimensions).

Exemple d'utilisation d'un tableau
----------------------------------

~~~VB
monTableau(3) = 4
If monTableau(3) = 4 Then
'Autre instructions

AutreTableau2D(2,3) = 2.43
~~~

Allocation dynamique
--------------------

-   Un tableau dynamique ne contient pas de taille à sa définition.

-   Il faut ensuite le redimensionner avant de l'utiliser avec Redim.

-   On peut ensuite le redimensionner à volonté.

-   On redimensionne avec le mot clé Preserve pour ne pas perdre
    l'information déjà contenu dans le tableau.

Passage de tableaux
-------------------

-   On passe des tableaux en paramètre sans indiquer leur taille. On ne
    fait qu'ajouter les parenthèses après l'identificateur pour indiquer
    que la variable sera un tableau.

-   Les tableaux sont obligatoirement passés par référence. Il faut donc
    faire attention aux modifications de tableaux durant l'exécution de
    sous-programme.

-   Pour savoir la taille d'un tableau, les fonction LBound et UBound
    permettent de déterminer respectivement l'indice inférieur et
    supérieur d'un tableau.

Retour de tableaux
------------------

Retour de tableaux Pour retourner un tableau, trois éléments doivent
être mis en places:

-   Le type de retour doit indiquer que le retour est un tableau (en
    ajoutant des parenthèses).

-   Créez un tableau de retour qui va contenir le tableau qu'on veut
    retourner (on ne peux pas utiliser la variable de retour comme un
    tableau directement).

-   Assignez la tableau de retour à la variable de retour.

Exemple de retour de tableau
----------------------------
~~~VB
Function retourTableau() As Integer()
    Dim retour() As Integer
    ReDim retour(5)
    retourTableau = retour
End Function
~~~

Destruction de tableaux dynamique
---------------------------------

-   Le mot clé [Erase]{.alert} permet de détruire un tableau dynamique
    déjà existant.

-   Faire attention, Ubound/Lbound sur un tableau vide génère une
    erreur.

~~~VB
Sub TestErase()
    Dim tab() As Integer
    ReDim tab(5)
    tab(3) = 5
    
    Erase tab
End Sub
~~~

Exercices
------------

* Écrivez une procédure qui saisit un nombre et et rempli un tableau de n cases avec différentes saisies pour chaque cases.
* Écrivez une procédure qui reçoit un tableau d'entier et affiche les valeurs du tableau dans une boîte de dialogue.
* Écrivez une fonction qui reçoit un entier n. La fonction retourne un tableau de n cases avec les puissances de 2 comme valeurs du tableau.