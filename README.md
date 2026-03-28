# Livret Numérique des Compétences (Version C)

Ce projet est une application en C avec interface graphique permettant de générer automatiquement un livret numérique des compétences pour chaque élève, sous la forme d’un fichier PowerPoint personnalisé.

## Fonctionnalités principales

- Importation d’une liste de compétences à partir d’un fichier .csv structuré.
- Génération d'un PowerPoint personnalisé.
- Interface utilisateur GTK.

## Installation & Utilisation

### Linux
#### Prérequis
- GCC
- GTK+ 3.0
- libzip
- libxml2

#### Compilation
```bash
make
```

#### Utilisation
```bash
./LivretCompetences
```

### Windows
#### Prérequis
- [Visual Studio 2022](https://visualstudio.microsoft.com/)
- [vcpkg](https://vcpkg.io/) (recommandé pour gérer les dépendances)

#### Installation des dépendances avec vcpkg
```powershell
vcpkg install gtk:x64-windows libzip:x64-windows libxml2:x64-windows
vcpkg integrate install
```

#### Compilation
1. Ouvrez `LivretCompetences.sln` dans Visual Studio.
2. Sélectionnez la configuration `Release` et la plateforme `x64`.
3. Cliquez sur `Générer la solution`.

## Licence

Projet open source sous licence MIT.
