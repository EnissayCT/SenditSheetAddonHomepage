# Sendit pour Google Sheets

Add-on **Editor** Google Sheets : créer des colis Sendit, synchroniser les statuts et imprimer les étiquettes 10x10, depuis la feuille ouverte.

Même rôle que les plugins Shopify / WooCommerce, pour les marchands qui travaillent dans Sheets.

## Fonctionnalités (v1.1)

- Création d'un colis depuis la ligne sélectionnée, un formulaire, ou une sélection de lignes
- Suivi des statuts à la demande (ignore livré / retourné / annulé / refusé)
- Étiquettes PDF 10x10
- Feuille modèle (villes Sendit, téléphone en texte, listes de statuts)
- Clés API et ville de ramassage par utilisateur (PropertiesService)

La synchronisation n'est **pas** automatique en v1 (déclencheurs horaires incompatibles avec un add-on Marketplace tant que l'ID de chaque feuille n'est pas stocké).

## Prérequis

- Compte Sendit : [app.sendit.ma](https://app.sendit.ma)
- Public Key + Secret Key (Intégration API)
- Google Sheets (ordinateur)

## Utilisation

1. Extensions → **Sendit pour Google Sheets** → **Configurer les clés API**
2. Coller les clés → **Charger les villes** → choisir la ville de ramassage → Enregistrer
3. **Créer une feuille modèle**
4. Remplir les lignes (statut `non traité`) → **Ouvrir Sendit** → créer / synchroniser / étiquettes

Colonnes du modèle : Code, Statut, Nom, Téléphone, Adresse, Ville, Montant, Produits, Référence, Commentaire, Autoriser_Ouverture, Autoriser_Essai.

Le téléphone doit rester du **texte** (ex. `0612345678`). La ville doit être le **nom** de la liste, pas un tarif.

## Pages à héberger (Marketplace)

Publiez ces fichiers en HTTPS (idéalement sur `sendit.ma`) et utilisez les URLs dans OAuth + fiche Marketplace :

| Fichier | Usage |
|---------|--------|
| `privacy.html` | Politique de confidentialité |
| `terms.html` | Conditions d'utilisation |
| `support.html` | Support |
| `index.html` | Site développeur / listing |

## Développement local

1. Google Sheet → Extensions → Apps Script
2. Copier `Code.gs`, `appsscript.json`, `Sidebar.html`, `Config.html`, `Welcome.html`
3. Exécuter `onOpen` → autoriser → recharger la feuille

API : `https://app.sendit.ma/api/v1`

## Publication

Editor add-on uniquement (pas de bloc `addOns` Workspace / Cards). Scopes :

- `spreadsheets.currentonly`
- `script.container.ui`
- `script.external_request`

Whitelist : `https://app.sendit.ma/`

Guide pas à pas : voir la conversation de revue / checklist GCP + OAuth + Marketplace SDK.

## Fichiers du projet

```
Code.gs              Logique
Sidebar.html         Barre latérale
Config.html          Clés API
Welcome.html         Accueil à l'installation
appsscript.json      Manifest Editor add-on
privacy.html         Confidentialité (à héberger)
terms.html           CGU (à héberger)
support.html         Support (à héberger)
index.html           Landing (à héberger)
```

Version: 1.1.0
