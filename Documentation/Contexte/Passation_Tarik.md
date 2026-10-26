# Untitled

# Projet 15 — Passation au membre 4

Tu  construire le rapport Power BI à partir de l’export de résultats. La soutenance a lieu le vendredi 23 octobre ; la copie maître des deux classeurs reste chez Tarik.

## Ce que tu reçois

| Fichier | Rôle |
| --- | --- |
| `Gestion_Logistique.xlsm` | Application de gestion : fiches, saisie, contrôles, clôture des périodes |
| `Decisionnel_Logistique.xlsm` | Application décisionnelle : import, calcul des coûts, 19 écrans d’analyse et de décision |

## Avant de commencer

Les chemins enregistrés sont ceux du PC de Tarik : tant que tu ne les corriges pas, l’actualisation affiche « L’application de gestion est introuvable ».

1. Copie les deux classeurs dans un même dossier.
2. Dans Excel : *Fichier → Options → Centre de gestion de la confidentialité → Paramètres → Paramètres des macros*, active les macros et coche **Accès approuvé au modèle d’objet du projet VBA**. Sans cette case, les générateurs de formulaires ne peuvent pas s’exécuter.
3. Ouvre le décisionnel, affiche la feuille masquée `Param_Dec` (clic droit sur un onglet → *Afficher*) et corrige la colonne *Valeur* :
    - **D-01** : chemin complet de Gestion_Logistique`.xlsm` ;
    - **D-02** : dossier des exports (crée un sous-dossier `Exports`) ;
    - **D-03** : chemin du futur rapport, `…\Projet15_Rapport.pbix`.
4. Masque de nouveau `Param_Dec`, puis enregistre.

Règles de travail jusqu’à la soutenance :

- Ne modifie aucun écran existant : tes deux tâches s’ajoutent au code, elles ne le remplacent pas.
- **Ctrl+S après chaque modification.** Une nuit de corrections a déjà été perdue par une fermeture sans enregistrer.
- Fais une copie datée des classeurs avant chaque grosse manipulation.

## Tâche  — Power BI

Power BI se branche sur un fichier Excel que le décisionnel produit lui-même : l’application calcule, Power BI dessine. Le cahier des charges réserve les graphiques à Power BI, les écrans VBA n’affichent que des chiffres.

**1. Produire l’export, sur le trimestre.** Dans le décisionnel : *Actualiser les données* → coche janvier, février et mars 2025 → *Importer* → *Calculer les coûts* → *Lancer le calcul complet* (environ 80 secondes) → *Rapport et exports* → **Export Power BI**. Le fichier `Resultats_PowerBI_AAAA-MM-JJ_HHMM.xlsx` arrive dans le dossier D-02. Comme son nom change à chaque export, copie le dernier sous un nom fixe, `Resultats_PowerBI.xlsx`, et branche Power BI sur ce nom : une actualisation suffira ensuite.

**2. Les tables.** Les feuilles `Calc_*` sont les faits calculés ; `Clients`, `Produits`, `Vehicules`, `Zones_Stockage`, `Periodes` et `Types_Couts` sont les référentiels.

**Piège :** dans `Types_Couts`, les en-têtes sont en **ligne 5**, sous un titre et une phrase d’explication. Dans Power Query : *Supprimer les lignes du haut* (4), puis *Utiliser la première ligne pour les en-têtes*.

**3. Les relations**, toutes de plusieurs vers un :

| Table de faits | Table de référence | Clé |
| --- | --- | --- |
| `Calc_Lignes` | `Calc_Commandes` | `N_commande` |
| `Calc_Commandes` | `Clients` | `Code_client` |
| `Calc_Lignes` | `Produits` | `Code_produit` |
| `Calc_Stock` | `Produits` | `Code_produit` |
| `Calc_Missions` | `Vehicules` | `Code_vehicule` |
| `Calc_Vehicules` | `Vehicules` | `Code_vehicule` |
| `Calc_Zones` | `Zones_Stockage` | `Code_zone` → `Code_zone_stockage` |
| `Calc_Depenses` | `Types_Couts` | `Code_type` |

**4. Les mesures de base**, définies comme dans l’application pour que les chiffres concordent :

```
Cout global = SUM ( Calc_Depenses[Montant_periode] )
Chiffre d affaires = SUM ( Calc_Commandes[CA_HT] )
Cout sur CA = DIVIDE ( [Cout global], [Chiffre d affaires] )
Cout par commande = DIVIDE ( [Cout global], COUNTROWS ( Calc_Commandes ) )
```

**5. Les pages à construire**, celles que les écrans ne dessinent pas : évolution mensuelle du coût face au budget ; répartition du coût par étape de la chaîne (`Calc_Depenses[Etape]`) ; Pareto des clients par coût logistique ; coût au kilomètre et remplissage par camion ; occupation et sous-activité des zones de stockage.

**6. Le contrôle de concordance**, à faire avant tout graphique : sur mars 2025 seul, Power BI doit retrouver **3 100 428 MAD** de coût global, **13,6 %** de coût sur chiffre d’affaires et **5 699 MAD** par commande, les trois chiffres de l’accueil. Sur le trimestre, le coût global attendu est **8 668 774 MAD**. Un écart signifie une relation ou une mesure fausse, pas une erreur de l’application : `Calc_Controles` y montre un écart de conservation de 0,00.

**7. L’enregistrement.** Enregistre le rapport sous `Projet15_Rapport.pbix`, au chemin indiqué en D-03 : le bouton *Ouvrir le rapport Power BI* de l’accueil l’ouvrira directement.

## Pièges VBA déjà rencontrés

Chacun de ces pièges a coûté au moins une erreur pendant le développement ; les reconnaître fait gagner beaucoup de temps.

- **VBA ignore la casse.** Une constante `TEXTE` et un paramètre `texte` sont le même nom : c’est ce qui a fait échouer la première version des générateurs (erreur 13).
- **Erreur 75** quand on recrée un UserForm du même nom dans la même session : supprime le formulaire, enregistre, ferme Excel, rouvre, puis relance le générateur.
- **Rien après `End Sub` sauf des commentaires**, et les déclarations de module tout en haut. Un bloc collé au mauvais endroit peut effacer l’en-tête d’une procédure : vérifie la liste déroulante des procédures après chaque collage.
- **`IIf` évalue ses deux branches** : une division par zéro dans la branche non retenue déclenche quand même l’erreur.
- **La recherche de colonne par morceau de nom** retient la première colonne qui contient le mot : « tarif » trouve `Zone_tarifaire`, « budget » trouve `Poste_budgetaire`. Pour un nom court, utilise `ColNomExact`.
- **`Split("", " ")(0)`** déclenche l’erreur 9 : couper une chaîne vide renvoie un tableau vide, pas une case vide.
- **Les fichiers `.bas` sont encodés en Windows-1252** : un autre encodage transforme les accents en points d’interrogation.
- **Le message « Une erreur est survenue (réf. mmjj-hhmm) »** renvoie à la feuille masquée `Journal`. Pour lire la dernière ligne, dans la fenêtre Exécution :

```
?ThisWorkbook.Worksheets("Journal").Cells(ProchaineLigneVide("Journal","A") - 1, 5).Value
```

## Ce qu’il faut renvoyer à Tarik

Renvoie des résultats, pas des classeurs : Tarik appliquera les générateurs sur la copie maître en dix minutes, ce qui évite de fusionner deux versions la veille de la soutenance.

- [ ]  Le fichier `Projet15_Rapport.pbix`, avec ses pages et ses relations.
- [ ]  Les trois chiffres du contrôle de concordance de mars, lus dans Power BI.
- [ ]  Tout message d’erreur rencontré, avec la ligne du journal correspondante.

## État de l’application et limites connues

Le cœur est validé : sur mars 2025, le calcul répartit 3 100 428 MAD avec un écart de conservation de 0,00 et un rapprochement exact avec le grand livre. Les écrans Synthèse, Indicateurs, Analyse par catégorie, Transport et flotte, Stockage, Clients et produits, Coût d’une commande, Scénarios et Nouveau produit ont été testés et corrigés le 22 octobre. Comparer les fournisseurs, Comparer et décider, Alertes, Suivi d’une alerte et Rapports restent à tester.

Limites connues, à ne pas prendre pour des erreurs dans Power BI :

- **Frais d’approche, retard des fournisseurs et livraisons à la date prévue** sont à n.d. : les données source ne sont pas renseignées. La règle de l’application est d’afficher n.d., jamais zéro.
- **Le contrôle des factures de transport** ne porte que sur la messagerie : la grille d’affrètement est comparée au code de zone, alors qu’elle est écrite en noms de villes.
- **Nouveau produit** n’utilise ni le champ *Clients visés* ni le volume : le transport y est estimé au coût moyen par kilo de l’entreprise, ce qui rend l’estimation prudente.
- **Les indicateurs comparent un mois au mois précédent** seulement quand un seul mois est importé ; sur un trimestre, la colonne de comparaison est à n.d.