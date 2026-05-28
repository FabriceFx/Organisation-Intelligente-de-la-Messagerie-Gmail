# Organisation intelligente de la messagerie Gmail


[🇫🇷 Version Française](#-version-française) | [🇬🇧 English Version](#-english-version)

![License MIT](https://img.shields.io/badge/License-MIT-blue.svg)
![Platform](https://img.shields.io/badge/Platform-Google%20Apps%20Script-green)
![Runtime](https://img.shields.io/badge/Google%20Apps%20Script-V8-green)
![Author](https://img.shields.io/badge/Auteur-Fabrice%20Faucheux-orange)

## 🇫🇷 Version Française


## 📋 Description

Ce projet propose une solution automatisée pour structurer la boîte de réception Gmail (Google Workspace). Il classifie dynamiquement les emails entrants en fonction de deux critères critiques :
1.  **Le contexte du destinataire** : Direct (À) ou Copie (Cc).
2.  **L'état de lecture** : Lu ou Non Lu.

Le script assure une **gestion d'état exclusive** (un email ne peut pas avoir deux libellés contradictoires) et applique un **code couleur visuel** pour une identification immédiate des priorités.

## ✨ Fonctionnalités clés

* **Classification Contextuelle** : Distinction automatique entre les messages qui vous sont adressés directement et ceux en copie.
* **Code Couleur sémantique** :
    * 🔴 **Rouge** : Message direct non lu (Urgent).
    * ⚪ **Gris** : Message lu (Archivé/Traité).
    * 🔵 **Bleu** : Copie non lue (Information).
* **Performance V8** : Utilisation stricte des opérations par lots (*Batch Operations*) pour respecter les quotas de l'API Google.
* **Automatisation** : Système de déclencheur (*Trigger*) intégré pour une exécution en arrière-plan toutes les 10 minutes.

## 🛠 Prérequis technique (Important)

Pour que la coloration des libellés fonctionne, vous devez activer le **Service Avancé Gmail**.

1.  Dans l'éditeur Apps Script.
2.  Dans la colonne de gauche, cliquez sur le `+` à côté de **Services**.
3.  Sélectionnez **Gmail API**.
4.  Cliquez sur **Ajouter**.

## 🚀 Installation

1.  Ouvrez [Google Apps Script](https://script.google.com).
2.  Créez un **Nouveau projet**.
3.  Copiez le contenu du fichier `Code.js` de ce dépôt dans l'éditeur.
4.  Sauvegardez (`Ctrl + S`).
5.  **Première exécution** : Lancez manuellement la fonction `installerAutomatisation`.
    * Cela va créer le déclencheur horaire.
    * Cela va tenter d'appliquer les couleurs aux libellés.
    * Acceptez les demandes d'autorisation.

## ⚙️ Configuration

Le comportement du script est centralisé dans l'objet constante `CONFIGURATION` en début de fichier :

```javascript
const CONFIGURATION = {
  LIBELLES: {
    // Noms des libellés générés
    PERSO_NON_LU: "# 1.Moi (Non Lu)",
    // ...
  },
  TRIGGER: {
    // Fréquence d'exécution
    FREQUENCE_MINUTES: 10
  }
  // ...
};


---
## 🇬🇧 English Version

> English translation coming soon.

---
<p align="center"><a href="https://faucheux.bzh" target="_blank" style="color: inherit; text-decoration: none;">&lt;&gt; par Fabrice Faucheux</a></p>