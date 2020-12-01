---
title: Boucles
author: Hugo Leblanc
date: Cours 2
pandocomatic_:
  use-template: 
  - presentation
  - presentation-handout
...
# Présentation du code

## Commentaires
- En-tête
    - Nom
    - Description
    - Auteur
- Intérieur des blocs de code
    - Description des actions
    - Un commentaire par action

## Espacement

- Indentation horizontale
- Espacement vertical (saut de ligne)

## Noms significatifs

- Noms significatifs au contexte
- Les noms des sous-programment sont en PascalCase
- Les noms des variables sont en camelCase 

## Constantes

- Les constantes remplacent l'utilisation des valeurs numériques statiques dans le code.

- Par convention, elles sont préfixé du mot `const`.

- Les déclarations se font au début du module (en dehors des sous-programmes)

- On assigne directement la valeur à la constante durant la déclaration.

~~~VB
Const constNomConstante As Integer = 50
~~~

# Structures de contrôles itératives

## Boucle - While

- Les boucles répètent un bloc de code.

- La boucle while répète le bloc de code tant que la condition est respecté.

- Le while est habituellement utilisé quand on ne connait pas le nombre d'itérations à faire.

~~~VB
While x < 10
    'Instructions
Wend
~~~

## Exercices

- Saisit un nombre à l’utilisateur et recommence la saisit tant que le nombren’est pas 0.

- Calcule la somme des nombres de 1 a n. On saisit n.

- Procédure qui saisit un nombre à l’utilisateur et va afficher le nombrefactoriel de la saisit (5!=1x2x3x4x5).

## Boucles - For

- La boucle for intègre la configuration d'un compteur à même la boucle.

- On utilise le for quand l'on connait le nombre d'itérations à faire.

~~~VB
For i = 1 To 10
    MsgBox i
Next i

For i = 1 To 10 Step 2
    MsgBox i
Next i

For i = 5 To 1 Step -1
    MsgBox i
Next i
~~~

## Exercices

- Calcule la somme des nombre de 1 à n. On saisit n. Utilisez un for.
- Compte le nombre de 0 entrez au clavier sur 10 essais.

# Opérateurs

## Opérateurs logiques

- Les opérateurs logiques opèrent sur des valeurs booléennes
- La conjonction ET : And
- La disjonction OU : Or
- La négation NON : Not

a | b | a And b | a Or b | Not a
--|---|---------|--------|------
F | F | F | F | V
F | V | F | V | V
V | F | F | V | F
V | V | V | V | F

## Exercices

- Procédure qui saisit l’âge de l’utilisateur et qui indique si ce dernier a droitau tarif réduit. Le tarif réduit est disponible pour les personnes d’âge mineur(<18) ou d’âge d’or (>60).

# Sous-programmes

## Sous-programmes
- Un sous-programme est un bloc de code réalisant une tâche précise.
- La fonction retourne (se transformer en) une valeur.
- Un sous-programme qui ne retourne rien est nommée une procédure.
- Un sous-programme peut avoir des paramètres. Ceux-ci dictent l’utilisation du sous-programme.

~~~VB
variable = NomFonction(paramètre1, paramètre2)
~~~

