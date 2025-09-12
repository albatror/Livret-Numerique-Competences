# Livret Numérique des Compétences

Ce projet est un script avec interface graphique permettant de générer automatiquement un livret numérique des compétences pour chaque élève, sous la forme d’un fichier PowerPoint personnalisé.

## Fonctionnalités principales

- Importation d’une liste de compétences à partir d’un fichier .txt structuré (domaines et sous-domaines).
- Génération automatique d’un PowerPoint contenant :
  - La fiche de présentation de l’élève (nom, prénom, date de naissance).
  - Les informations sur les différentes sections fréquentées (Toute Petite Section, Petite Section, Moyenne Section, Grande Section), avec :
    - Année scolaire
    - École
    - Enseignant(s)
  - Les compétences acquises pour chaque domaine et sous-domaine durant l’année.
  - Ajout de la photo de l’élève et d’illustrations sur les pages.
  - Une page par domaine, avec mise en page automatique (auto-scaling des zones de texte et d’image).
- Interface utilisateur pour :
  - Prévisualiser les pages du livret
  - Sélectionner, ajouter ou supprimer des compétences
  - Gérer les photos à intégrer

## Public visé

- Enseignant(e)s de maternelle ou primaire
- Écoles souhaitant automatiser la création de livrets personnalisés pour chaque élève

## Installation & Utilisation

1. Clonez ce dépôt :
   ```bash
   git clone https://github.com/albatror/Livret-Numerique-Competences.git
   ```
2. Placez votre fichier de compétences au format texte dans le répertoire du projet.
3. Lancez le script principal (voir documentation interne pour le nom exact du fichier à exécuter).
4. Suivez l’interface pour saisir les informations de l’élève, choisir les compétences et générer le livret PowerPoint.

## Dépendances

- Python 3.x
- Bibliothèques utilisées : (à compléter selon vos besoins, ex: `tkinter`, `python-pptx`, etc.)

## Exemple de fichier de compétences

```
Domaine 1
  Sous-domaine 1.1
  Sous-domaine 1.2
Domaine 2
  Sous-domaine 2.1
...
```

## Licence

Projet open source sous licence MIT.