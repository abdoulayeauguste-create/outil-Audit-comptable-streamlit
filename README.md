# Outils comptables Streamlit

Cette application regroupe deux modules :

- `Amortissements des immobilisations`
- `Comparateur de balances N / N-1`

## Module amortissements

Le module amortissements permet de saisir ou d'importer une liste d'immobilisations avec les colonnes suivantes :

- `REFERENCE`
- `DESIGNATION`
- `VALEUR ORIGINE`
- `DATE ACQUISITION`
- `DUREE (ANS)`

Pour chaque immobilisation, l'application calcule automatiquement :

- l'annuite complete ;
- le prorata de l'annee de reference ;
- l'annuite de l'annee ;
- l'amortissement cumule ;
- la `VNC` de fin d'annee ;
- le statut : annuite complete, incomplete ou nulle.

Les cas suivants sont couverts automatiquement :

- acquisition pendant l'annee de reference ;
- acquisition au 1er janvier avec annuite complete ;
- fin d'amortissement pendant l'annee de reference ;
- immobilisation encore en cours d'amortissement avec annuite complete ;
- immobilisation deja totalement amortie.

Deux modes de prorata sont proposes :

- `Mensuel (mois entame compte)`
- `Journalier (exact au jour)`

Un exemple precharge reprend vos 5 cas typiques sur l'annee `2025`.

## Module balances

Le module balances permet de charger deux balances comptables et de calculer automatiquement :

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

## Rendre l'application accessible a distance

Pour un acces web a distance, l'option la plus simple est :

1. creer un depot GitHub ;
2. y envoyer `streamlit_app.py`, `amortissement_core.py`, `balance_core.py` et `requirements.txt` ;
3. deployer sur Streamlit Community Cloud.

Pour un usage interne en entreprise, vous pouvez aussi deployer Streamlit sur un serveur ou une machine virtuelle et publier l'URL via :

- `Streamlit Community Cloud` pour un partage simple ;
- `Render`, `Railway` ou `Azure Web App` pour plus de controle ;
- un serveur `Windows` ou `Linux` avec `Nginx` en reverse proxy.

## Deploiement sur Streamlit Community Cloud

Les fichiers utiles pour le deploiement sont :

- `streamlit_app.py`
- `amortissement_core.py`
- `balance_core.py`
- `requirements.txt`
- `runtime.txt`
- `.streamlit/config.toml`
- `sample_immobilisations.csv`

Etapes minimales :

1. creer un depot GitHub ;
2. y envoyer ces fichiers ;
3. aller sur `https://share.streamlit.io/` ;
4. cliquer sur `Create app` ;
5. choisir le depot, la branche et le fichier `streamlit_app.py` ;
6. dans `Advanced settings`, choisir Python `3.12` si besoin ;
7. lancer le deploiement.

## Depot pret pour Streamlit

Le projet est maintenant structure pour un deploiement direct :

- `streamlit_app.py` : point d'entree de l'application ;
- `requirements.txt` : dependances Python ;
- `runtime.txt` : version Python ciblee ;
- `.streamlit/config.toml` : configuration Streamlit adaptee au cloud ;
- `sample_immobilisations.csv` : fichier d'exemple a importer pour tester l'application des le premier lancement.

## Procedure de publication conseillee

1. creer un depot GitHub public ou prive ;
2. envoyer le contenu du dossier dans ce depot ;
3. verifier que le fichier principal est bien a la racine : `streamlit_app.py` ;
4. ouvrir [Streamlit Community Cloud](https://share.streamlit.io/) ;
5. cliquer sur `Create app` ;
6. selectionner le depot GitHub, la branche et `streamlit_app.py` comme `Main file path` ;
7. deployer ;
8. tester l'onglet `Amortissements` avec `sample_immobilisations.csv`.

## Recommandation de publication

Pour une premiere mise en ligne, `Streamlit Community Cloud` est le meilleur choix :

- mise en service rapide ;
- URL web partageable ;
- mises a jour simples a chaque `git push` ;
- suffisant pour un prototype ou un outil metier leger.

## Export

Les resultats peuvent etre exportes en `Excel (.xlsx)` ou en `CSV`.

## Vision produit

Vos petits logiciels peuvent etre regroupes progressivement dans une seule application avec plusieurs onglets ou modules. La bonne approche consiste generalement a :

1. conserver un noyau commun pour les fonctions de calcul ;
2. ajouter un menu ou un tableau de bord avec un bouton par module ;
3. partager les imports/exports, la base clients ou les parametres entre les modules.
