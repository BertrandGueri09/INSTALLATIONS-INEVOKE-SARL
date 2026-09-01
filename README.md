# Dashboard — Suivi des puissances installées par client · INEVOKE SARL

Tableau de bord Streamlit centré sur le **suivi des puissances déjà installées, par client**,
alimenté par le formulaire **KoboToolbox** INEVOKE. Identité visuelle INEVOKE conservée.

## Ce que fait cette version

- **Filtre principal par client** et exclusion automatique des installations encore
  « en cours d'installation » (on ne suit que les puissances **réellement posées**).
- **8 KPI orientés client** : nombre de clients, installations, puissance PV installée (kWc/MWc),
  puissance moyenne par client, onduleurs (kVA), stockage (kWh), production réelle, disponibilité.
- **Onglet « Par client »** : classement des clients par puissance installée, part de chaque
  client dans le parc, synthèse chiffrée (sites, PV, onduleurs, stockage, prod., dispo.),
  arborescence client → sites, export CSV.
- **Puissances & équipements** : PV vs onduleurs et stockage par client, nuage de
  dimensionnement (kVA/kWc) coloré par client.
- **Performance énergétique** : production estimée vs réelle par client, ratio de performance,
  productible spécifique (kWh/kWc/jour) par installation.
- **Carte** (OpenStreetMap / MapLibre) : chaque installation dimensionnée par sa puissance,
  colorée par statut, info-bulle avec client, ville et puissances.
- **Rapport par installation** : tableau détaillé groupé par client, export **Excel (3 feuilles)** et CSV.
- **Fiche client** : puissances agrégées du client, puissance par site, détail et export.

## Installation

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Connexion KoboToolbox

Copier `.streamlit/secrets.toml.example` vers `.streamlit/secrets.toml` et renseigner
`KOBO_API_URL`, `KOBO_ASSET_UID`, `KOBO_API_TOKEN`, `KOBO_FORM_LINK`.
Sans Secrets, l'application tourne en **mode démonstration** ; dès que Kobo renvoie des
données, elles remplacent automatiquement les données fictives.

Le champ **`client`** du XLSForm est la clé de tous les regroupements. Les puissances
totales et le stockage sont recalculés côté application si Kobo ne renvoie pas les
champs `calculate`, pour des indicateurs toujours cohérents.

## Adaptation au formulaire en ligne

Ce tableau de bord est calé sur le formulaire **« Suivi des installations solaires déjà installées — INEVOKE SARL »**. Comme le XLSForm utilise des groupes, l'application retire automatiquement les préfixes de groupe renvoyés par l'API Kobo (ex. `identification/nom_site` → `nom_site`) et décode les codes de choix en libellés lisibles (`san_pedro` → San-Pédro, `on_grid` → On-grid, etc.). Il suffit donc de renseigner l'`ASSET_UID` et le token du projet dans les Secrets pour que les données réelles s'affichent, regroupées par client.

## Collecte intégrée au dashboard

L'onglet **« Collecte & saisie »** embarque directement le formulaire Enketo (« Suivi des installations solaires déjà installées — INEVOKE SARL ») dans le tableau de bord : les agents peuvent enregistrer une nouvelle installation ou mettre à jour une fiche sans quitter l'application (le formulaire fonctionne aussi hors-ligne). Après une soumission, cliquer sur **« Forcer l'actualisation »** dans la barre latérale recharge les indicateurs. Un bouton **« Ouvrir en plein écran »** sert de repli si le navigateur bloque l'affichage intégré. Le lien du formulaire est modifiable via le Secret `KOBO_FORM_LINK`.

## Remise à zéro / choix de la source

Dans la barre latérale, **« Source des données »** permet de choisir, tant que Kobo n'est pas connecté, entre :

- **Tableau vide (démarrer à zéro)** — par défaut : tous les compteurs à zéro, aucun graphique fictif. Idéal pour démarrer la collecte réelle proprement.
- **Données de démonstration** — jeu fictif pour visualiser le rendu.

Dès que les Secrets Kobo sont renseignés, ce choix est ignoré et les vraies données s'affichent automatiquement.
