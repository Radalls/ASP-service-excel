# ExcelService

ExcelService est un service ASP.net Core MVC EF6 permettant de convertir des types de modèles de données en template Excel remplissable pour insérer des données. 

Il permet également d'évaluer et de valider les données reçues d'un template et de les convertir en données de base de données.

## Installation

Il est avant tout primordial d'ajouter le package ClosedXML (via nuget) aux packages du projet. Le module se base sur les outils de ce package pour fonctionner.

Le service `ExcelService.cs` ainsi que les fichiers d'extensions `ExcelExtensions.cs` et `CommonsExtensions.cs` doivent obligatoirement être implémenté dans votre projet pour en faire usage.

Il est également proposé un ensemble d'attributs customisés dans le dossier `Attributes` qui permettent d'affiner le comportement du module par rapport à des modèles de données
dont certaines propriétés seraient marquées de ces attributs customisés.

Le contenu des dossiers `Controllers` et `Models` n'est présent qu'à titre d'exemple est n'est pas nécessaire au fonctionnement du module.

## Usage

Le code fourni n'a pas besoin d'être modifié pour fonctionner, hormis pour les notions suivantes :

- les références `namespace` manquantes
- les références `using` manquantes
- les références au contexte de base de donnée (probablement) non compatible.

L'utilisation du module est présentée dans les Controllers donnés en exemple : `FooController.cs` et `BarController.cs`.

Il suffit, dans une méthode d'action d'un Controller, d'invoquer une instance de service `ExcelService` et d'appeler la méthode de service souhaitée en lui donnant les paramètres dont elle a besoin.

Un exemple de d'interaction côté Vue est disponible dans le fichier `Views/Index.cshtml`. C'est une vue partielle minimaliste qui propose un formulaire d'intéraction avec le service.

Le code est entièrement documenté pour comprendre son fonctionnement dans le détail.

## Auteur

Léo SALLARD, 2021
