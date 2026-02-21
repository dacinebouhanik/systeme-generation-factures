# 📊 Système automatisé de génération de factures clients

## 🎯 Description du projet

Ce projet a été réalisé dans le cadre d’un travail personnel en Python.

L’objectif est d’automatiser la génération de factures clients à partir de plusieurs fichiers Excel contenant des feuilles de commande.

Le programme permet de :

- Lire plusieurs fichiers Excel (.xlsx)
- Vérifier que leur structure est conforme
- Extraire les données des clients
- Regrouper les commandes par client
- Trier les commandes par date
- Générer automatiquement une facture Excel par client

---

## ⚙️ Fonctionnement

Le programme suit une logique de type ETL :

1. **Extraction** : lecture des fichiers Excel présents dans un dossier
2. **Transformation** : nettoyage des données (remplacement des valeurs vides, regroupement par client, tri par date)
3. **Chargement** : création automatique d’un fichier facture par client

---

## 📁 Structure du projet
