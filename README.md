![oriflame_app](https://github.com/user-attachments/assets/a0887200-9828-4a9f-b05c-61d359f2dc67)


#  Oriflame Products Scraper

Un outil de web scraping avancé pour extraire les données produits du site Oriflame Maroc avec Selenium et Python.

##  Description

**Oriflame Products Scraper** est un script Python robuste conçu pour extraire automatiquement les informations produits du site e-commerce Oriflame Maroc. L'outil navigue intelligemment à travers les catégories, charge dynamiquement le contenu et exporte les données structurées dans un format Excel exploitable.

##  Fonctionnalités

###  Navigation Intelligente
- **Détection automatique** des catégories produits
- **Mapping sémantique** des catégories français→anglais
- **Navigation complète** à travers toutes les sections

###  Gestion du Contenu Dynamique
- **Chargement progressif** avec détection du bouton "Charger plus"
- **Attentes intelligentes** pour le rendu JavaScript
- **Limite de sécurité** pour éviter les boucles infinies

###  Extraction de Données Complète
- **Liens produits** complets
- **Noms et marques** des produits
- **Prix** en devise locale
- **Évaluations** clients
- **Gestion robuste** des données manquantes

###  Export et Formatage
- **Export Excel** structuré
- **DataFrame Pandas** pour l'analyse
- **Formatage propre** des données

##  Architecture Technique

### Stack Technologique
```python
# Core Technologies
Selenium WebDriver - Navigation et interaction navigateur
Pandas - Manipulation et export des données
OpenPyXL - Génération de fichiers Excel
ChromeDriver - Automatisation Chrome

# Gestion des Attentes
WebDriverWait - Attentes explicites
Expected Conditions - Conditions de chargement

### Structure du Code :
get_categories()      
get_products()        
main()

### Gestion d'Erreurs
Try/Except robuste sur chaque opération

Timeouts configurables pour les chargements

Fallbacks pour les données optionnelles

### Installation
Prérequis Système
Python 3.8+

Google Chrome installé
ChromeDriver compatible

###  Processus Automatisé
Lancement du navigateur en mode headless

Navigation vers Oriflame Maroc

Détection des catégories principales

Parcours de chaque catégorie

Chargement progressif des produits

Extraction des données structurées

Export Excel automatique
