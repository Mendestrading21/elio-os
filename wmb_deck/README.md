# Deck WMB — WallStreet Market Brief

Présentation premium (~46 pages, 16:9) combinant le business plan, l'analyse
du site et la charte de design WMB (banque privée : sobre, blanc dominant,
accents or/navy parcimonieux, data-viz lisibles).

## Livrables
- `WMB_Presentation.pdf` — version PDF vectorielle autonome (prête à présenter/partager).
- `WMB_Presentation.pptx` — version PowerPoint éditable, graphiques natifs.

## Régénération
```bash
pip install reportlab PyMuPDF python-pptx
python3 generate_wmb_pdf.py     # -> WMB_Presentation.pdf
python3 generate_wmb_deck.py    # -> WMB_Presentation.pptx
```

## Structure (6 sections)
1. La vision — résumé exécutif, vision/mission, valeurs, problème, solution
2. L'offre produit — entonnoir, modules, scoreboard signature, breaking
3. Marché & opportunité — TAM/SAM/SOM, tendances, concurrence, positionnement, moats
4. Modèle économique — abonnement, économie unitaire, funnel, conversion, canaux
5. Produit, technologie & site — stack, pipeline IA, CMS, identité, forces/faiblesses
6. Finances & risques — conformité, équipe, projections, scénarios, KPIs, risques, feuille de route

Chaque fonction de slide prend ses données en paramètres : le deck est
régénérable avec d'autres chiffres. Constantes de design centralisées (THEME).

Note : « Playfair Display » étant absent de l'environnement, le PDF embarque un
serif élégant de substitution (Liberation Serif) ; à l'ouverture du PPTX,
PowerPoint utilisera Playfair Display s'il est installé.
