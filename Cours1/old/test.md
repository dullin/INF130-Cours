---
author:
- Hugo Leblanc
subtitle: 'Semaine 1 -- Programmation, variables, entrées/sorties, conditionnelles'
title: 'INF130 : Ordinateurs et programmation'
---

# Présentation du cours


## Présentation personnelle

Présentation personnelle

Hugo Leblanc

hugo.leblanc\@etsmtl.ca

Baccalauréat en génie électrique de l'ÉTS en 2012

14 ans de programmation scolarisée

Spécialisation en systèmes embarqués

# Présentation du cours

Présentation du cours

-   Plan de cours

    -   Pondération

    -   Date de remise

    -   Politique de plagiat

-   Moodle ena.etsmtl.ca

    -   Centre de toures les intéractions du cours

    -   Notes de cours, exerciecs, remise, messages

-   Slack

    -   Plate-forme de communication à l'extérieur du cours
    
-   Structure de travail du cours

    -   1000 erreurs à faire avant la fin de la session

    -   Très facile de perdre le contrôle de la matière

    -   Les travaux pratiques sont des préparations aux examens

Cours
=====

Objectifs de la semaine

-   Éditeur VBA

-   Modules

-   Commentaires

-   Procédures

-   Variables, types et assignation

-   Saisie et affichage

-   Opérateurs

-   Structure de contrôles conditionnelles (if)

Éditeur VBA
-----------

Éditeur VBA

-   Chaque application de la suite Microsoft Office contient un éditeur
    VBA qui nous permet de faire de la programmation à l'intérieur de
    l'application.

-   Dans le cours, nous utiliserons Excel comme application de base pour
    tous nos travaux.

-   Pour ouvrir l'éditeur, appuyer sur [Alt-F11]{.alert} à l'intérieur
    d'Excel

-   Un fichier contenant des programmes est un fichier avec des macros
    et doit avoir l'extension correspondante à la sauvegarde
    ([.xlsm]{.alert})

-   Un seul éditeur peux traiter plusieurs document Excel en même temps

Modules
-------

Modules

-   Un module est un regroupement de sous-programme

-   Plusieurs modules peuvent être dans le même fichier Excel

-   Les modules sont utilisés pour découper un grand programme en plus
    petites parties indépendantes

-   On ajoute un nouveau module par le menu [Insert -\> Module]{.alert}

Commentaires
------------

Commentaires

-   Le caractère [']{.alert} (apostrophe) précède tous commentaires

-   Le reste de la ligne après le ' ne sera pas considéré par VBA durant
    l'exécution de code

-   Le commentaires sont primordiaux à la programmation

-   Les commentaires sont utilisés en en-tête de modules, fonction et
    programme ainsi qu'à l'intérieur d'un programme pour aider à
    comprendre l'intention des instructions.

Procédures
----------

Procédures (sous-programmes)

-   Un sous-programmes est un amalgam d'instructions qui seront
    exécutées ensemble.

-   Le premier type de sous programme que nous allons voir est la
    procédure.

-   On exécute une procédure en appuyant sur le bouton « play » ou avec
    le raccourci F5.

<!-- -->

    'Les parties en bleu sont
    'essentielles à la syntaxes.
    'Le nom de la procédure est
    'à la discretion du programmeur.
    Sub NomDeProcedure()
        'Instructions
    End Sub
            

Variables, types, Assignation
-----------------------------

Variables - Déclaration

-   Une variable est la combinaison d'un espace mémoire réservé, un
    identificateur, une valeur et un type.

-   Une variable doit être déclaré avant de pouvoir être utilisé.

-   La déclaration est sous la forme suivante:

        Dim nomVariable As Integer
                    

-   Dim est le mot réservé pour la déclaration d'espace.

-   Le nom de la variable est à votre discretion.

-   Le dernier mot est le type de la variable.

-   Pour nous aider avec les déclarations nous utiliserons la
    configuration Option Explicit au début de tout nos modules.

-   Les déclarations se font au début des sous-programme.

Variables - Types

-   Un variable doit être définie par un type

-   Le type indique quel genre de donnée peut exister dans la variable

-   Les types de base sont:

    -   [Integer]{.alert} : Les nombres entiers de -32 768 à 32 767.

    -   [Long]{.alert} : Les nombres entiers de -2 147 483 648 à 2 147
        483 647.

    -   [Double]{.alert} : Les nombres réels (avec une certaine marge
        d'erreur) de $-1.79769313486232 x 10^308$ à
        $1.79769313486232 x 10^308$.

    -   [String]{.alert} : Les chaines de caractères, donc du texte. Par
        exemple \"Allo!\" ou \"Miam, miam, miam. Les bons gros
        légumes.\" . Les chaines de caractères sont toujours délimiter
        de double guillmets.

    -   [Boolean]{.alert} : Vrai ou faux (True et False sous VBA).

Variables - Assignation

-   On assigne une valeur à une variable l'opérateur =

        Dim nomVariable As Integer
        nomVariable = 10
                    

-   L'assignation va seulement prendre le type de valeur que la variable
    peut contenir (ne pas mettre du texte dans une variable de type
    Integer)

-   Les variables ont une valeur par nulle défault.

-   On utilise les valeur dans les variables en invoquant leur nom dans
    des équations

        nomVariable = nomVariable + 5
                    

Affichage et saisie
-------------------

Affichage - MsgBox

-   L'affichage se fait avec un appel à MsgBox. Ce sous-programme doit
    recevoir une chaine de caractères qui sera affiché par la suite.

        MsgBox "Allo monde!"
        Call MsgBox("Allo monde!")
                    

-   On peut aussi envoyer une valeur numérique qui sera convertie en
    chaine avant d'être affiché.

        MsgBox 345
                    

-   MsgBox est plus versatile que les exemples plus haut mais il faudra
    attendre un plus grande connaissance des appel de sous-programme
    avant d'approfondir le sujet.

Exercice -- Salaire hebdomadaire

-   Écrivez un sous-programme avec trois variables (taux horaire, nombre
    d'heures travaillées et salaire hebdomadaire.

-   Calculez le salaire à partir du taux horaire et du nombre d'heures
    travaillées.

-   Affichez-le salaire dans une boite de dialogue par la suite.

Saisie - InputBox

-   La saisie de texte ou de valeur se fait avec un appel au
    sous-programme InputBox.

-   On assigne le retour (la valeur saisie) dans un variable lors de
    l'appel.

~~~VB
Dim reponse As String
reponse = InputBox("Entrez quelque chose:")
~~~

-   Le type du retour est une chaine de caractères. Pour saisir un
    nombre, on convertie la valeur avant de l'assigner.

        Dim nombre As Integer
        nombre = Val(InputBox("Entrez quelque chose:"))
                    

Exercice -- Salaire hebdomadaire avec InputBox Ajoutez à l'exercice
précédent une saisie du taux horaire et du nombre d'heures travaillées
avec des InputBox

Opérateurs
----------

Opérateurs

-   Les opérateurs arithmétique /, +, -, \* et \^ sont disponibles pour
    les opérations numérique.

-   L'opérateur \\permet de faire la division entière. La division
    entière coupe la partie fractionnaire).

-   L'opérateur Mod permet de faire le module d'un nombre. Le module est
    le restant après un division.

-   Les opérateurs relationnels \<, \<=, \>, \>=, =, \<\> (différent)
    permet de comparer deux valeurs numériques.

-   L'opérateur & peut faire la concaténation (coller) deux chaines de
    caractères ensemble.

Exercice -- Salaire hebdomadaire avec concaténation Ajoutez à l'exercice
précédent un affichage avec une phrase complète dans le MsgBox.

Structure de contrôle conditionnel
----------------------------------

Structure de contrôle conditionnelle - if

-   La structure confitionnelle nous permet de prendre des décisions
    durant l'exécurtion de nos scripts

-   La décision à prendre doit être fait sur une expression booléenne
    (vrai ou faux)

    \smallskip
        If nombre > 10 Then
            MsgBox "Plus que 10!"
        End If

    \hfill
        If nombre > 10 Then
            MsgBox "Plus que 10!"
        ElseIf nombre > 20 Then
            MsgBox "Plus que 20!"
        Else
            MsgBox "Moins..."
        End If

-   Un seul bloc d'instruction est exécuté.

-   Le elseif peut être répété au besoin.

Exercice -- Salaire hebdomadaire avec temps supplémentaire Ajoutez à
l'exercice précédant un calcul pour le temps supplémentaire.
