# 📄 Convertisseur Universel vers PDF

Un outil de bureau simple, rapide et doté d'une interface graphique moderne pour convertir massivement vos fichiers Word, Excel et Texte en PDF, le tout localement sans dépendre d'une connexion internet.

## ✨ Fonctionnalités Principales

* **Multi-Formats :** Prend en charge les fichiers `.txt`, `.docx` (Word) et `.xlsx` (Excel).
* **Interface Moderne :** Une interface utilisateur élégante avec mode sombre, développée avec `customtkinter`.
* **Structure Conservée :** Ne pollue pas vos dossiers d'origine. Le logiciel crée automatiquement un dossier miroir (ex: `MonDossier_PDF`) contenant toute l'arborescence convertie.
* **100% Privé et Hors Ligne :** Aucune donnée n'est envoyée sur le cloud. La conversion se fait entièrement sur votre machine.

## 🛠️ Prérequis et Installation (Mode Développeur)

Si vous souhaitez faire tourner le code source directement en Python, vous aurez besoin de **Python 3.x** installé sur votre machine.

1.  Clonez ou téléchargez ce dépôt.
2.  Ouvrez un terminal dans le dossier du projet.
3.  Installez les dépendances requises à l'aide de `pip` :

```bash
pip install customtkinter python-docx openpyxl reportlab
