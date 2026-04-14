# Stock Movement Report — Format Sage 100

**Odoo 18 | v18.0.1.2.0 | LGPL-3**

Rapport de mouvements de stock valorisés reproduisant fidèlement le format Sage 100, avec calcul du CMUP (Coût Moyen Unitaire Pondéré).

## Fonctionnalités

- **Wizard de filtres** : Période (obligatoire), Produit(s), Dépôt, Lot/N° de série, Par dépôt (toggle)
- **PDF QWeb** : Mise en page paysage A4 identique au rapport Sage
  - En-tête : Société, Dépôt, Période
  - Report d'ouverture (stock initial)
  - Détail des mouvements par produit
  - Sous-totaux par produit, Total dépôt, "A reporter"
- **Export Excel** : Même structure avec formatage couleur professionnel
- **CMUP dynamique** : Calculé via `stock.valuation.layer` (méthode AVCO), recalculé à chaque entrée de stock

## Colonnes du rapport

| Colonne | Description |
|---------|-------------|
| Date mouv. | Date du mouvement de stock |
| Type mouv. | REC (Réception), BL (Bon de livraison), INT (Interne), FAB (Fabrication), INV (Inventaire), RET (Retour) |
| N° de pièce | Référence du picking ou du mouvement |
| Référence / Tiers | Nom du partenaire (client/fournisseur) |
| +/- | Quantité entrée (+) ou sortie (-) |
| Solde | Stock courant (balance glissante) |
| P.R. unitaire | CMUP au moment du mouvement |
| Stock permanent | Valeur du stock = Solde × CMUP |

## Capture de référence (Sage 100)

![Sage 100](https://raw.githubusercontent.com/Lalaina710/stock_movement_report/main/static/description/sage_reference.png)

## Dépendances

- `stock`
- `stock_account`
- `xlsxwriter` (Python, pour l'export Excel)

## Installation

1. Copier le module dans le dossier `addons`
2. Mettre à jour la liste des applications
3. Installer "Mouvements de Stock (Format Sage)"

## Accès

**Menu** : Inventaire → Rapports → Mouvements de stock (Sage)

**Droits** : Utilisateur Stock (`stock.group_stock_user`)

## Source de données

- `stock.move` (mouvements validés, état `done`)
- `stock.valuation.layer` (coût unitaire par mouvement, méthode AVCO)
- `stock.move.line` (pour le filtrage par lot)

## Changelog

### v18.0.1.2.0
- **feat:** Toggle "Par dépôt" dans le wizard — activé = rapport groupé par dépôt (défaut), désactivé = tous dépôts combinés en un seul bloc

### v18.0.1.1.0
- **fix:** CMUP d'ouverture inclut désormais les layers de revalorisation (LEFT JOIN au lieu de INNER JOIN)
- **fix:** Filtrage `company_id` ajouté aux requêtes SQL (support multi-société)
- **fix:** Conversion timezone utilisateur pour les bornes de dates (important UTC+3 Madagascar)

### v18.0.1.0.0
- Version initiale — wizard, PDF QWeb, export Excel, CMUP via `stock.valuation.layer`

## Performance

- Requêtes SQL pour le calcul du stock d'ouverture et du CMUP historique
- Prefetch bulk des `stock.valuation.layer` (évite le N+1)
- Cache ORM natif pour les données statiques

## Auteur

SOPROMER — Développé pour la migration Sage 100 → Odoo 18
