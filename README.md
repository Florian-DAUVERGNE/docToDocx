# docToDocx
Ce projet est une application Windows qui convertit des fichiers .doc en .docx en utilisant win32com.client pour interagir avec Microsoft Word. L'interface utilisateur (Tkinter) permet de sélectionner un dossier et de convertir tous les fichiers .doc qu'il contient. Elle affiche la progression et permet d'arrêter la conversion à tout moment.

# DOC to DOCX Converter

## Description
Une application Windows qui convertit les fichiers .doc en .docx en utilisant `win32com.client` pour interagir avec Microsoft Word. L'interface utilisateur avec Tkinter permet de sélectionner un dossier, de convertir les fichiers, d'afficher la progression et d'arrêter la conversion.

## Prérequis
- Windows
- Microsoft Word
- Python 3.11
- Bibliothèques: `pywin32`, `tkinter`

## Installation
1. Clonez le repository:
    ```sh
    [git clone https://github.com/Florian-DAUVERGNE/docToDocx.git]
    cd docToDocx
    ```
2. Installez les dépendances:
    ```sh
    pip install pywin32
    ```

## Utilisation
1. Exécutez le script:
    ```sh
    python doc_to_docx_converter.py
    ```
2. Sélectionnez un dossier et lancez la conversion.

## Contribuer
1. Forkez le projet.
2. Créez une branche (`git checkout -b ma-nouvelle-fonctionnalite`).
3. Commitez vos modifications (`git commit -am 'Ajout d'une nouvelle fonctionnalité'`).
4. Pushez (`git push origin ma-nouvelle-fonctionnalite`).
5. Ouvrez une Pull Request.

## Licence
MIT. Voir [LICENSE](LICENSE) pour plus de détails.
