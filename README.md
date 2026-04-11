# Comparateur de balances N / N-1

Cette application de bureau permet de charger deux balances comptables et de calculer automatiquement :

- `SOLDE N`
- `SOLDE N-1`
- `VARIATION (ABS)` = `SOLDE N - SOLDE N-1`
- `VARIATION (%)` = `(VARIATION (ABS) / SOLDE N-1) * 100`

## Formats pris en charge

- `CSV` et `TXT` : support natif
- `XLSX` / `XLS` : supporte si `pandas` et `openpyxl` sont installes

## Colonnes attendues

Chaque balance doit contenir au minimum les informations suivantes :

- `COMPTE`
- `LIBELLE`
- `SOLDE`

Des variantes courantes sont reconnues automatiquement pour les en-tetes, par exemple `Numero de compte`, `Intitule`, `Montant`, `Balance`, etc.

## Lancer l'application

```powershell
python app.py
```

## Lancer la version web Streamlit

```powershell
streamlit run streamlit_app.py
```

Ensuite, ouvrez dans le navigateur l'adresse affichee par Streamlit, en general :

`http://localhost:8501`

## Deploiement sur Streamlit Community Cloud

Les fichiers utiles pour le deploiement sont :

- `streamlit_app.py`
- `balance_core.py`
- `requirements.txt`

Etapes minimales :

1. creer un depot GitHub ;
2. y envoyer ces fichiers ;
3. aller sur `https://share.streamlit.io/` ;
4. cliquer sur `Create app` ;
5. choisir le depot, la branche et le fichier `streamlit_app.py` ;
6. dans `Advanced settings`, choisir Python `3.12` si besoin ;
7. lancer le deploiement.

## Publication web

Si vous voulez la rendre accessible a d'autres utilisateurs via une adresse web, vous pouvez deployer `streamlit_app.py` sur la meme plateforme que votre application d'intangibilite.

## Export

Le resultat peut etre exporte en `Excel (.xlsx)` ou en `CSV` avec les colonnes :

`COMPTE ; LIBELLE ; SOLDE N ; SOLDE N-1 ; VARIATION (ABS) ; VARIATION (%)`

## Fusionner plusieurs petits logiciels

Oui, c'est tout a fait possible. La bonne approche consiste generalement a :

1. conserver un noyau commun pour les fonctions de calcul ;
2. ajouter un menu ou un tableau de bord avec un bouton par module ;
3. partager les imports/exports, la base clients ou les parametres entre les modules.

Autrement dit, vos petits logiciels peuvent devenir progressivement une seule application avec plusieurs ecrans.
