# ExcelService

ExcelService est un service ASP.net Core MVC EF6 permettant de convertir des types de mod�les de donn�es en template Excel remplissable pour ins�rer des donn�es. 

Il permet �galement d'�valuer et de valider les donn�es re�ues d'un template et de les convertir en donn�es de base de donn�es.

## Installation

Il est avant tout primordial d'ajouter le package ClosedXML (via nuget) aux packages du projet. Le module se base sur les outils de ce package pour fonctionner.

Le service `ExcelService.cs` ainsi que les fichiers d'extensions `ExcelExtensions.cs` et `CommonsExtensions.cs` doivent obligatoirement �tre impl�ment� dans votre projet pour en faire usage.

Il est �galement propos� un ensemble d'attributs customis�s dans le dossier `Attributes` qui permettent d'affiner le comportement du module par rapport � des mod�les de donn�es
dont certaines propri�t�s seraient marqu�es de ces attributs customis�s.

Le contenu des dossiers `Controllers` et `Models` n'est pr�sent qu'� titre d'exemple est n'est pas n�cessaire au fonctionnement du module.

## Usage

Le code fourni n'a pas besoin d'�tre modifi� pour fonctionner, hormis pour les notions suivantes :

- les r�f�rences `namespace` manquantes
- les r�f�rences `using` manquantes
- les r�f�rences au contexte de base de donn�e (probablement) non compatible.

L'utilisation du module est pr�sent�e dans les Controllers donn�s en exemple : `CompanyController.cs` et `AccountantController.cs`.

Il suffit, dans une m�thode d'action d'un Controller, d'invoquer une instance de service `ExcelService` et d'appeler la m�thode de service souhait�e en lui donnant les param�tres dont elle a besoin.

Un exemple de d'interaction c�t� Vue est disponible dans le fichier `Views/Index.cshtml`. C'est une vue partielle minimaliste qui propose un formulaire d'int�raction avec le service.

Le code est enti�rement document� pour comprendre son fonctionnement dans le d�tail.

## Auteur

L�o SALLARD (Radalls), 2021