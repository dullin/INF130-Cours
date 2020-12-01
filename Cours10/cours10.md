---
title: INF130 - Événements et formulaire
date: Cours 10
pandocomatic_:
    use-template: handout
---

Événements d'Excel
==================


-   Les objets d'Excel peuvent aussi avoir des événements reliés avec du
    code.

-   Un événement est une action pris par l'utilisateur ou l'application
    qui est reconnu par l'objet. (Ex: Sélectionné une feuille, Ouvrir
    Excel, Clicker sur un bouton)

-   La liste d'événement des différents objets est disponible dans des
    menus déroulant dans l'éditeur VBA.

Formulaires
===========


-   Un formulaire est créé au même niveau qu'un module.

-   Le formulaire aura deux aspect: visuel et code.

-   On peut modifier l'aspect visuel pour y ajouter des contrôles de
    notre choix.

-   On peut lié les contrôles à du code avec les événements des objets.

Gestion du formulaires
=======================

-   Chaque contrôles nous donne un nouvel objet avec des événements.

-   On doit géré les événements dans la partie code du formulaire.

-   On utilize les methodes show et hide de l'objet du formulaire pour
    faire afficher et disparaitre le formulaire à partir de nos module.

-   Les formulaires sont une exception à la règle habituelle sur les
    variables globales. Nous les utiliserons pour garder en mémoire
    l'information contenu dans le formulaire.
