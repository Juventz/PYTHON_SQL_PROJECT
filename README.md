# Projet d'Analyse des Données de Vente

## Description

Ce projet permet d'extraire, transformer et analyser des données de ventes provenant d'un fichier Excel et de les stocker dans une base de données PostgreSQL via Docker. Une analyse visuelle est réalisée à l'aide de Matplotlib.

## Structure du projet


├── docker-compose.yml
├── main.py
├── Makefile
├── output.png
├── requirements.txt
├── Test_Données.xlsx

## Prérequis

Python 3.x

Docker et Docker Compose

Fichier .env contenant les variables de connexion PostgreSQL :

```env```
- `POSTGRES_HOST=localhost`
- `POSTGRES_DB=nom_de_la_base`
- `POSTGRES_USER=utilisateur`
- `POSTGRES_PASSWORD=mot_de_passe`
- `POSTGRES_PORT=5432`

## Installation

Cloner le dépôt

git clone <URL_du_repository>
cd <nom_du_dossier>

Installer les dépendances Python

pip install -r requirements.txt

Démarrer PostgreSQL avec Docker

make up

Charger les données dans PostgreSQL

python main.py

## Utilisation du Makefile

Le Makefile permet de gérer facilement le conteneur PostgreSQL avec Docker.

Démarrer PostgreSQL : make up

Arrêter PostgreSQL : make down

Supprimer PostgreSQL et les volumes : make clean

Supprimer l'image PostgreSQL : make rmi

Afficher les logs du conteneur : make logs

Se connecter à PostgreSQL : make psql

Reconstruire et relancer : make rebuild

Lister les conteneurs actifs : make ps

Lister les images Docker : make images

Lister les volumes Docker : make volumes

## Fonctionnalités principales

Chargement des données depuis Excel : Extraction de plusieurs feuilles et nettoyage des colonnes.

Stockage dans PostgreSQL : Importation des données dans des tables.

Analyse et visualisation : Génération de graphiques avec Matplotlib pour l'analyse des ventes par pays et année.

Requêtes SQL : Exécution de requêtes SQL sur PostgreSQL et affichage des résultats.

Résultat attendu

Le script main.py génère un fichier output.png contenant des visualisations des ventes.
