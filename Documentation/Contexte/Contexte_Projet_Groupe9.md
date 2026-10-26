# Projet 15 — Analyse des coûts logistiques — Dossier complet du groupe 9

**À quoi sert ce fichier.** Il contient tout le contexte du projet : l'entreprise, le problème, l'architecture, la répartition du travail entre les quatre membres, la structure de la base de données, les formules de calcul, la spécification des 41 écrans élément par élément, le code déjà écrit et les tests de recette. Un membre du groupe qui donne ce fichier à Claude dans son propre compte obtient un assistant qui connaît le projet aussi bien que celui qui l'a écrit.

Module de programmation VBA · EMINES – UM6P · Groupe 9 · Version du 10/10/2026 · Présentation le 23/10/2026

---

## Sommaire

- 0. Mode d'emploi : démarrer avec Claude
- 1. Le projet en bref
- 2. L'entreprise, le problème et l'objectif
- 3. Périmètre et architecture de la solution
- 4. Répartition du travail entre les quatre membres
- 5. Règles communes de développement
- 6. Partie 1 — Socle commun, accès et paramétrage
- 7. Partie 2 — Opérations, contrôles et clôture
- 8. Partie 3 — Calcul et analyse des coûts
- 9. Partie 4 — Décision, alertes, restitution et Power BI
- 10. Structures des tables à créer
- 11. La base Atlantis Distribution 2025 : les 43 feuilles
- 12. Paramètres et compteurs de la feuille Ref_Data
- 13. Les 20 formules de calcul
- 14. Spécification des écrans, élément par élément
- 15. Le socle commun : code déjà écrit et testé
- 16. Pièges rencontrés et décisions prises
- 17. Calendrier et intégration
- 18. Priorités, règles de gestion, utilisateurs et contraintes
- 19. Prompt de démarrage, à copier dans Claude
- 20. Ce qu'il faut savoir pour l'oral du 23/10

---

## 0. Mode d'emploi : démarrer avec Claude

1. Ouvre une nouvelle conversation sur claude.ai.
2. Joins ce fichier `.md` à ton premier message, ainsi que le classeur `Projet15_Couts_Logistiques.xlsm` si tu l'as.
3. Écris le prompt de démarrage de ta partie, donné à la section 14.
4. Garde la même conversation pour toute ta partie : Claude se souviendra de ton code au fur et à mesure.

Ce que Claude peut faire et ne peut pas faire :

- il ne peut pas exécuter de VBA ni ouvrir Excel : c'est toi qui colles le code dans l'éditeur et qui lances les tests ;
- il peut en revanche lire le classeur `.xlsm` si tu le joins, et calculer lui-même les résultats attendus de tes tests sur les vraies données ;
- envoie-lui une capture d'écran quand un message d'erreur apparaît : il lit les images.

Méthode conseillée, celle utilisée pour la partie 1 : procédure par procédure. Claude donne l'énoncé d'une procédure (nom, paramètres, comportement) et les résultats chiffrés attendus, tu écris ou colles le code, tu lances les tests dans la fenêtre Exécution (Ctrl+G), tu renvoies les résultats, il corrige. C'est plus long que de recevoir le code tout fait, mais c'est la seule façon de pouvoir défendre son code à l'oral du 23/10.

---

## 1. Le projet en bref

| Élément | Contenu |
|---------------------|--------------------------------------------------------------------------|
| Sujet | Projet VBA n°15 : analyse des coûts logistiques |
| Cadre | Module de programmation VBA, EMINES – UM6P |
| Maître d'œuvre | Groupe 9, quatre membres |
| Outils | Microsoft Excel et VBA (Windows), Power BI Desktop |
| Données | Base « Atlantis Distribution 2025 », 43 feuilles, environ 200 000 lignes |
| Livrables | Application transactionnelle, application décisionnelle, rapport Power BI, notice |
| Envoi du cahier des charges | 06/10/2026 |
| Présentation | 23/10/2026, examen final |

Les quatre fichiers de la solution :

| Fichier | Rôle |
|---------------------------------------------|---------------------------------------------------------------|
| `Projet15_Couts_Logistiques.xlsm` | Application de gestion (volet transactionnel), construite dans la base Atlantis |
| `Projet15_Decisionnel.xlsm` | Application décisionnelle ; lit le fichier précédent en lecture seule |
| `BDD_Decisionnel_AAAA-MM-JJ_hhmm.xlsx` | Fichier de résultats, en valeurs seules, écrit par l'écran D19 |
| `Rapport_Couts_Logistiques.pbix` | Rapport Power BI, lit le fichier de résultats |

---

## 2. L'entreprise, le problème et l'objectif

### Description du client

Le client est Atlantis Distribution SA, une entreprise fictive créée pour le projet. Distributeur-grossiste de produits de grande consommation au Maroc, elle achète des marchandises auprès de 35 fournisseurs, dont 14 situés à l'étranger qui livrent par conteneur maritime ou par camion, les stocke, puis les livre à 158 clients : 21 grandes et moyennes surfaces, 40 grossistes, 42 supérettes, 33 cafés, hôtels et restaurants, 10 sites de restauration collective et 12 distributeurs à l'export, en Afrique de l'Ouest et en Europe. Son catalogue compte 140 produits répartis en sept familles, de l'épicerie sèche aux surgelés en passant par les produits laitiers, les boissons et les produits d'entretien.

Son organisation repose sur un entrepôt central à Casablanca, qui réapprovisionne par navettes deux plateformes régionales situées à Tanger et à Agadir. Les trois sites offrent ensemble 2 098 emplacements de palettes, répartis en zones à température ambiante, en froid positif pour les produits frais et en froid négatif pour les surgelés. L'entreprise emploie 98 salariés logistiques, exploite sa propre flotte de 27 véhicules, dont 11 frigorifiques, et fait appel à 18 prestataires : transporteurs, transitaires, compagnies maritimes, assureurs et garages. En 2025, elle a traité 6 474 commandes comportant 44 004 lignes de produits et réalisé 3 923 missions de transport, pour un chiffre d'affaires de 245,4 millions de dirhams. Sa logistique lui a coûté 33,8 millions de dirhams, soit 13,8 % de ce chiffre d'affaires, alors que son budget en prévoyait 32,8 millions.

### Contexte général du projet

Le projet concerne la Direction logistique, et plus précisément le contrôle de gestion logistique, qui suit les charges, les compare au budget et doit en expliquer les écarts. C'est ce service qui utilisera la solution au quotidien, avec les responsables opérationnels chargés d'enregistrer les dépenses et les opérations.

Le coût étudié couvre toute la chaîne, de l'arrivée des marchandises à la livraison des clients, en quatre étapes auxquelles s'ajoutent des frais de structure. L'approvisionnement regroupe le fret, les frais portuaires, le transit et l'assurance des produits importés ; il a coûté 4,8 millions de dirhams en 2025. Le stockage comprend les loyers, l'énergie, la manutention, l'entretien et le gardiennage des trois sites ; il a coûté 5,8 millions. Le traitement des commandes recouvre leur saisie, leur préparation et leur emballage ; il a coûté 2,8 millions. Le transport et la livraison, assurés par la flotte de l'entreprise, par des transporteurs et par des prestataires à l'export, représentent la plus grande part avec 19,2 millions, soit 57 % du total, auxquels s'ajoutent 1,1 million de frais de direction et de systèmes. Le transport est aussi le coût le plus difficile à attribuer à un client, car un même camion livre plusieurs clients au cours d'une même tournée.

### Objectif du projet

Le projet doit doter la Direction logistique d'un outil complet de gestion et d'analyse de ses coûts, organisé en deux applications. La première, dite transactionnelle, sert à enregistrer et à tenir à jour toutes les informations : dépenses, événements, commentaires, commandes, missions et fiches de référence. La seconde, dite décisionnelle, exploite ces informations pour calculer automatiquement les coûts, les analyser à l'aide d'indicateurs, les répartir jusqu'au client et au produit, et simuler l'effet d'une décision avant de la prendre. Ses résultats alimentent un rapport Power BI.

Cet outil poursuit quatre objectifs. Le premier est de structurer l'information, en enregistrant chaque dépense une seule fois et en la rattachant à une catégorie, à un fournisseur et à l'objet qu'elle concerne. Le deuxième est de calculer les coûts sans intervention manuelle, sur n'importe quelle période. Le troisième est de les expliquer, grâce aux indicateurs, aux comparaisons avec le budget et aux alertes. Le quatrième est d'aider à la décision, en mesurant à l'avance l'effet d'un choix comme la sous-traitance d'une région ou le renouvellement d'un camion.

La réussite du projet sera jugée sur des critères vérifiables. Toutes les commandes livrées en 2025 devront recevoir un coût complet. La somme des coûts répartis devra être égale au total des dépenses, soit 33 796 058 dirhams, et les coûts calculés devront correspondre au grand livre comptable à moins de 0,5 % près pour chaque poste. Le calcul de l'année ne devra pas dépasser une minute, et les six anomalies volontairement introduites dans la base devront être détectées. Enfin, une nouvelle catégorie de coût devra être utilisable sans modifier le programme, et Power BI devra afficher les mêmes chiffres que l'application décisionnelle.

### Problème posé

L'entreprise connaît le total de ses coûts logistiques, mais elle ne sait pas dire quels camions, quels sites, quels clients ni quels produits les consomment. Sa comptabilité classe en effet les charges par nature (carburant, maintenance, assurances, loyers, salaires, factures des transporteurs), ce qui indique combien a été dépensé, mais pas pour qui ni pour quoi. Or deux clients qui génèrent le même chiffre d'affaires peuvent coûter très différemment selon leur éloignement, la taille et la fréquence de leurs commandes, ou la part de produits frais et surgelés qu'ils achètent. L'entreprise ne peut donc ni expliquer son dépassement de budget, ni repérer les clients ou les activités qui coûtent plus qu'ils ne rapportent, ni évaluer à l'avance l'effet d'une décision.

Ce calcul est aujourd'hui impossible pour quatre raisons. Aucun outil n'enregistre les dépenses de manière structurée, en les rattachant à une catégorie, à l'objet concerné et à l'événement qui les a provoquées. Les données sont dispersées entre plusieurs sources, la gestion commerciale, l'entrepôt, le transport, la flotte et la comptabilité, qui ne sont reliées que par des codes. Une grande partie des charges est commune à plusieurs activités, comme une tournée qui livre plusieurs clients ou un entrepôt qui abrite tous les produits, ce qui exige des règles de partage claires. Enfin, les flux ne se correspondent pas un à un : une commande peut être livrée en plusieurs fois, annulée ou enlevée par le client lui-même, et une même facture de transporteur couvre souvent plusieurs expéditions. La question à traiter est donc la suivante : **combien coûte réellement la logistique, par catégorie, par camion, par site, par commande, par client et par produit, sur une période donnée, et quelles décisions permettraient de réduire ce coût ?**

### Description de l'existant

Le suivi actuel repose sur le grand livre analytique, qui enregistre les charges par mois, par compte et par centre de coûts, et sur le budget annuel, détaillé par poste et par site. Leur comparaison montre que le réalisé 2025 dépasse le budget de 957 058 dirhams, soit 2,9 %, sans permettre de dire d'où vient cet écart. Elle montre seulement que trois postes concentrent 66 % de la dépense : le transport assuré par la flotte de l'entreprise, avec 12,0 millions de dirhams, le stockage, avec 5,8 millions, et le transport sous-traité en national, avec 4,5 millions. Les données opérationnelles, comme les commandes, les missions, les stocks, le carburant, la maintenance ou les factures des transporteurs, existent sous forme d'extractions séparées, rassemblées pour le projet dans un classeur de 43 feuilles.

Cet existant présente plusieurs limites. Il n'existe aucun moyen de saisir, de rechercher ou de corriger une dépense, ni d'enregistrer un événement comme une panne ou un litige, ni d'ajouter une remarque sur un fournisseur ou un véhicule. Les catégories d'analyse sont figées par le plan comptable : il est impossible, par exemple, de suivre séparément la maintenance préventive et la maintenance curative. L'analyse se fait uniquement par mois civil, alors que le contrôleur de gestion a souvent besoin d'étudier une période quelconque. Aucun coût n'est calculé par camion, par commande, par client ou par produit, les factures des transporteurs ne sont pas comparées aux tarifs convenus, et les erreurs de saisie, comme un poids de produit manquant ou un ticket de carburant saisi deux fois, ne sont pas détectées ligne par ligne. Enfin, les coûts unitaires utiles à la décision, comme le coût d'un kilomètre ou d'un emplacement de palette, ne sont jamais calculés, ce qui rend toute simulation impossible.

---

## 3. Périmètre et architecture de la solution

### Périmètre

La solution couvre l'ensemble des coûts logistiques de l'entreprise. Elle traite le transport réalisé par la flotte de l'entreprise, qu'il s'agisse des tournées de livraison, des navettes entre sites ou des enlèvements chez les fournisseurs, le transport confié à des prestataires en messagerie ou en affrètement de camions complets, et le transport à l'export, par route et par mer. Elle couvre aussi le stockage sur les trois sites et dans les trois régimes de température, le traitement des commandes et l'approvisionnement des produits importés. Sur ce périmètre, elle assure la gestion des données, le calcul et la répartition des coûts, les contrôles, la synthèse, les indicateurs, les scénarios et les rapports. Tous les montants sont exprimés en dirhams hors taxes.

La solution mesure et explique les coûts, mais n'optimise pas les opérations elles-mêmes : l'optimisation des tournées, du chargement des camions et la gestion de la flotte relèvent d'autres projets du module, et la solution utilise donc les tournées et les chargements tels qu'ils ont eu lieu. Sont également exclus la comptabilité générale, le paiement des factures, la facturation des clients, les droits de douane et la TVA, la prévision de la demande et la connexion directe à un logiciel de gestion d'entreprise.

### Organisation générale de la solution

Conformément aux recommandations de l'encadrement, la solution comprend deux applications Excel/VBA complémentaires et un rapport Power BI, et les informations circulent toujours dans le même sens, de la saisie vers l'analyse. L'application transactionnelle est un outil de gestion classique, avec un menu et des formulaires, dans lequel les utilisateurs enregistrent, modifient, recherchent et suppriment les informations de l'entreprise ; elle contrôle la qualité des données, mais ne calcule aucun indicateur, et permet de clôturer une période lorsque ses données sont vérifiées. L'application décisionnelle récupère les périodes clôturées, sans pouvoir les modifier, calcule automatiquement les coûts, produit la synthèse et les indicateurs, déclenche les alertes et permet de simuler des scénarios.

Le rapport Power BI lit le fichier de résultats produit par l'application décisionnelle et présente les coûts sous forme de pages graphiques interactives. Il additionne et filtre des montants déjà calculés, sans refaire les calculs, ce qui garantit que Power BI et Excel affichent toujours les mêmes chiffres.

*[Figure 1 — Organisation de la solution]*

Cette organisation sépare clairement les rôles : ceux qui saisissent les données travaillent dans la première application, ceux qui analysent travaillent dans la seconde et dans Power BI. Elle protège aussi les données, puisqu'aucun calcul ni aucune simulation ne peut altérer les informations enregistrées. Trois principes s'appliquent enfin à toute la solution. Rien n'est figé dans le programme : toutes les listes de choix, comme les types de coûts, de documents ou d'événements, sont complétées par l'utilisateur et prises en compte immédiatement. Les coûts unitaires, les taux et les seuils sont regroupés dans un écran de paramètres modifiable sans programmation. Chaque information n'est saisie qu'une seule fois, puis réutilisée partout où elle est nécessaire.

Le plan de navigation ci-après (figure 2) montre l'ensemble des écrans de la solution et leur enchaînement. Les écrans de l'application de gestion portent un code qui commence par T, ceux de l'application décisionnelle un code qui commence par D, et les pages du rapport Power BI un code qui commence par P ; ces codes sont repris dans tout le document. Le parcours type d'une période, à droite du plan, résume l'ordre dans lequel les écrans sont utilisés, de la saisie des dépenses jusqu'à la présentation des résultats dans Power BI.

*[Figure 2 — Plan de navigation]*

---

## 4. Répartition du travail entre les quatre membres

Chaque partie correspond à des écrans précis du cahier des charges et à un fichier. Les parties 1 et 3 tiennent les fichiers maîtres et font l'intégration ; les parties 2 et 4 travaillent dans une copie et livrent leurs modules.

| Partie | Responsable | Écrans | Fichier de travail | Fournit aux autres | Reçoit des autres |
| --- | --- | --- | --- | --- | --- |
| 1. Socle commun, accès et paramétrage | Tarik EL BARAKA | T1, T2, T3, T4, T5, T6, T13, T16, T17, T18 | Projet15\_Couts\_Logistiques.xlsm (maître gestion) | Modules du socle, tables Types\_Couts et Listes, paramètres, menu | Modules et formulaires de la partie 2 à intégrer |
| 2. Opérations, contrôles et clôture | Membre 2 | T7, T8, T9, T10, T11, T12, T14, T15 et reprise des dépenses 2025 | Copie de Projet15\_Couts\_Logistiques.xlsm | Tables Depenses, Evenements, Commentaires, Clotures, Resultats\_Controles | Socle, Types\_Couts et Listes (partie 1) |
| 3. Calcul et analyse des coûts | Membre 3 | D1 à D12 | Projet15\_Decisionnel.xlsm (maître décisionnel) | Feuilles de résultats Calc\_\*, Indicateurs, fonction CalculerCouts | Périodes clôturées et table Depenses (partie 2), socle (partie 1) |
| 4. Décision, alertes, restitution et Power BI | Membre 4 | D13 à D19, P1 à P4 | Copie de Projet15\_Decisionnel.xlsm et Rapport\_Couts\_Logistiques.pbix | Scénarios, alertes, rapports PDF, fichier BDD\_Decisionnel, rapport Power BI | Résultats Calc\_\* et CalculerCouts (partie 3), socle (partie 1) |

La charge est comparable : 8 à 12 écrans par partie. La partie 3 porte le moteur de calcul, la plus difficile ; ses écrans D11 et D12 sont donc classés souhaitables.

---

## 5. Règles communes de développement

Ces règles permettent à quatre personnes de coder en même temps sans s'attendre ni s'écraser.

1. **Un dossier, les mêmes noms.** Chacun travaille dans Documents\\Projet15 et garde les noms de fichiers, de feuilles, de tables et de colonnes de ce document. Un changement de nom se décide à quatre.
2. **Personne ne modifie le fichier d'un autre.** La partie 1 tient le fichier maître de gestion, la partie 3 celui de l'application décisionnelle. Les parties 2 et 4 travaillent dans une copie et livrent leur code au responsable du fichier maître.
3. **Le code circule par fichiers.** Dans l'éditeur VBA, clic droit sur un module ou un formulaire → Exporter un fichier (.bas, .frm) ; le destinataire fait Fichier → Importer un fichier. Une feuille passe d'un classeur à l'autre par clic droit sur l'onglet → Déplacer ou copier.
4. **Le socle commun appartient à la partie 1.** Les modules modUtils, modParametres, modMessages et modJournal sont importés tels quels par les autres. Pour une nouvelle fonction commune, on la demande à Tarik au lieu de la coder chacun de son côté.
5. **Règles de code.** Option Explicit dans chaque module ; un en-tête (objet, entrées, sortie) avant chaque procédure ; un Dim par variable ; aucun coût, taux ou seuil écrit en dur, tout passe par LireParametre ; les codes gardent le format de la base (P0001, CL0001, VH01, FR001, TR01, EMP001).
6. **Tables Excel.** Toutes les feuilles de données sont des tableaux nommés T\_…. Pour supprimer un enregistrement, on supprime la ligne entière, jamais son seul contenu : sinon la ligne vide reste dans le tableau et DerniereLigne la compte.
7. **Tests chiffrés.** Chaque procédure est livrée avec ses tests dans la fenêtre Exécution. Une partie n'est terminée que si ses tests de recette passent sur la base complète.
8. **Sauvegarde avant chaque intégration.** Ctrl+Z n'annule jamais ce qu'une macro a écrit : on garde une copie datée du fichier maître avant d'y importer quoi que ce soit.

---

## 6. Partie 1 — Socle commun, accès et paramétrage

La partie 1 fournit les fonctions que toutes les autres parties appellent, puis les écrans qui configurent l'application de gestion. Elle est déjà commencée : DerniereLigne et ChercherLigne sont terminées et testées, LireParametre est en cours.

**Écrans et priorités**

| Code | Écran | Priorité | Contenu |
| --- | --- | --- | --- |
| — | Socle commun | Indispensable | Modules modUtils, modParametres, modMessages, modJournal |
| T2 | Menu principal | Indispensable | Feuille Menu à boutons, bandeau, feuilles de données masquées |
| T5 | Types de coûts | Indispensable | Table Types\_Couts initialisée depuis le grand livre, environ 60 types |
| T6 | Listes de choix | Indispensable | Table Listes : documents, événements, prestations, véhicules, commandes, fonctions, segments, devises |
| T3, T4 | Fiches de référence | Indispensable | Un formulaire modèle (camions), puis les 8 autres fiches |
| T13 | Paramètres de calcul | Important | Ajout des seuils S-10 à S-14 dans Ref\_Data |
| T18 | Messages | Important | Règles appliquées par modMessages dans les deux applications |
| T1 | Connexion | Souhaitable | Trois profils : administrateur, analyste, consultation |
| T16 | Utilisateurs et journal | Souhaitable | Tables Utilisateurs et Journal |
| T17 | Sauvegardes et tests | Souhaitable | Sauvegarde datée, vidage des opérations, données fictives |

**Fonctions du socle commun**

| Fonction | Module | Rôle | État |
| --- | --- | --- | --- |
| DerniereLigne(nomFeuille, \[colonne\], \[wb\]) | modUtils | Dernière ligne remplie, y compris dans un autre classeur | Terminée |
| ChercherLigne(nomFeuille, colonne, valeur, \[wb\]) | modUtils | Ligne où se trouve une valeur, 0 si elle est absente | Terminée |
| LireParametre(code) | modParametres | Valeur d'un paramètre de Ref\_Data | En cours |
| EcrireParametre(code, valeur) | modParametres | Modifie un paramètre, réservé à l'administrateur, avec bornes | À faire |
| ProchainNumero(compteur) | modUtils | Lit et incrémente un compteur de T\_Compteurs | À faire |
| FormatNumero(prefixe, numero, chiffres) | modUtils | Fabrique un code : DEP-2025-00001, EVT-0001, TC-01 | À faire |
| MsgInfo, MsgErreur, Confirmer | modMessages | Messages uniformes de l'écran T18 | À faire |
| GererErreur(procedure) | modMessages | Message compréhensible au lieu de la fenêtre technique d'Excel | À faire |
| Journaliser(action, objet, detail) | modJournal | Écrit une ligne dans T\_Journal | À faire |
| EstNombrePositif, EstDateValide | modUtils | Contrôles de saisie communs | À faire |

**Livraisons aux autres parties**

- 11/10 : les modules du socle, en fichiers .bas, aux trois autres membres.
- 12/10 : les tables Types\_Couts et Listes remplies ; la partie 2 en a besoin pour la saisie des dépenses et la reprise.
- À chaque intégration : import dans le fichier maître des modules et formulaires de la partie 2.

**Tests de recette**

- LireParametre("A\_ADV\_REF") + 4 × LireParametre("B\_PREP\_REF") = 312,12, la formule 5 du cahier des charges.
- ProchainNumero sur le compteur des commandes renvoie 256475 et l'enregistre comme dernier numéro ; celui des missions renvoie 603924.
- Un type de coût créé dans T5 apparaît dans la liste de T7 sans modifier le programme.
- La fiche du produit P0140, sans poids, ne peut pas être enregistrée tant que le poids n'est pas saisi (anomalie A1).
- Avec le profil consultation, tous les boutons de modification sont grisés.

---

## 7. Partie 2 — Opérations, contrôles et clôture

La partie 2 enregistre l'activité : les dépenses, les événements, les commentaires, les commandes et les missions. Elle contrôle ensuite les données et clôture les périodes, les seules que l'application décisionnelle peut importer. Elle reprend aussi les 4 169 écritures du grand livre 2025 comme dépenses de départ.

**Écrans et priorités**

| Code | Écran | Priorité | Contenu |
| --- | --- | --- | --- |
| — | Reprise des dépenses 2025 | Indispensable | Module modReprise : une dépense par écriture de la feuille Charges |
| T7 | Saisie d'une dépense | Indispensable | Formulaire qui s'adapte au type de coût choisi ; contrôle des doublons ; avoirs |
| T8 | Recherche des dépenses | Indispensable | Recherche multicritère avec total des montants trouvés |
| T9 | Registre des événements | Indispensable | Pannes, sinistres, litiges, avec durée d'immobilisation et dépenses liées |
| T10 | Remarques et commentaires | Indispensable | Commentaires datés sur n'importe quel enregistrement |
| T11 | Commande client | Important | En-tête et lignes ; poids, volume, palettes et montant calculés |
| T12 | Mission de transport | Important | Véhicule ou transporteur, kilomètres, commandes chargées, contrôle de capacité |
| T14 | Contrôle des données | Important | Codes inconnus, dates impossibles, doublons, valeurs manquantes |
| T15 | Clôture de période | Important | Refusée tant qu'une anomalie critique reste ; sauvegarde datée |
| — | Dépenses récurrentes | Souhaitable | Génération automatique d'un loyer ou d'un contrat mensuel |

**Règle de la reprise, à respecter strictement.** Chaque écriture du grand livre (feuille Charges) devient une dépense : son type vient du couple compte et nature de charge, son objet est le centre de coût, sa date le mois de l'écriture, et son origine indique le numéro d'écriture. Les tables détaillées (Carburant, Maintenance, Frais\_Transport\_Ext, Frais\_Appro) ne sont pas reprises comme dépenses : elles servent seulement de clés pour rattacher les coûts aux camions et aux missions. Ainsi aucune charge n'est comptée deux fois.

**Modules et formulaires** : modReprise, modDepenses, modEvenements, modCommandes, modControles, modCloture ; frmDepense, frmRechercheDepenses, frmEvenements, frmCommentaires, frmCommande, frmMission, frmControles, frmCloture.

**Ce que la partie 2 reçoit et livre**

- Reçoit de la partie 1 le socle commun (11/10) et les tables Types\_Couts et Listes (12/10). En attendant, elle crée quelques types de test avec la même structure.
- Livre à la partie 3, le 14/10, la table Depenses remplie par la reprise ; dès le 11/10, les tables vides Depenses et Clotures avec leurs en-têtes, pour que la partie 3 puisse tester son import.

**Tests de recette**

- Reprise : 4 169 dépenses pour un total de 33 796 057,86 MAD, dont 7 304 638,65 MAD de salaires bruts (compte 6171).
- Une facture avec le même fournisseur, le même numéro et le même montant qu'une dépense existante déclenche l'alerte de doublon.
- Une assurance de 12 000 MAD couvrant du 01/01 au 31/12 est enregistrée avec sa période de couverture.
- T14 détecte les six anomalies de la base : A1 produit P0140 sans poids, A2 client CL0020 sans distance, A3 mission 601904 aux compteurs inversés, A4 ticket de carburant saisi deux fois (1 354,94 MAD d'écart), A5 commande 252590 livrée avant d'être commandée, A6 facture TR01-202503 comptée deux fois.
- La clôture est refusée tant qu'une anomalie critique reste ouverte ; une fois acceptée, une copie datée du classeur est créée.
- Une mission dont le chargement dépasse la charge utile du véhicule est refusée.

---

## 8. Partie 3 — Calcul et analyse des coûts

La partie 3 crée l'application décisionnelle Projet15\_Decisionnel.xlsm. Elle importe les périodes clôturées, calcule automatiquement tous les coûts et les présente dans les écrans d'analyse. Son moteur de calcul est utilisé ensuite par la partie 4 pour les scénarios et par Power BI.

**Écrans et priorités**

| Code | Écran | Priorité | Contenu |
| --- | --- | --- | --- |
| D2 | Actualisation des données | Indispensable | Import en lecture seule des périodes clôturées ; refus d'une période non clôturée |
| D3 | Calcul et contrôles de justesse | Indispensable | Conversion en dirhams, étalement, charges calculées, répartition, contrôle de conservation |
| D4 | Synthèse de période | Indispensable | Coût global « de telle période à telle période », par catégorie et par objet |
| D5 | Indicateurs de performance | Indispensable | Indicateurs par famille, comparés à la référence et à l'objectif |
| D1 | Accueil | Indispensable | Chiffres clés et boutons ; la zone des alertes est remplie par la partie 4 |
| D6 | Analyse par catégorie | Important | Regroupement par type, famille, étape, compte, site, mois |
| D7 | Transport et flotte | Important | Coût complet des camions, coût des missions, contrôle des factures |
| D8 | Stockage | Important | Occupation, détention, sous-activité, rotation, couverture |
| D9 | Clients et produits | Important | Coût complet et marge après logistique, concentration du coût |
| D10 | Coût d'une commande | Important | Jalons, composantes, missions et lignes d'une commande |
| D11 | Statistiques | Souhaitable | Moyenne, médiane, écart type, distribution |
| D12 | Méthodes de calcul | Souhaitable | Fiches de définition et formules |

**Modules** : modUtils et modParametres (copies du socle), modImport, modCalcul, modSynthese, modIndicateurs, modAnalyses, modStats.

**Fonctions que la partie 3 met à disposition**

- ImporterPeriodes() : copie les tables de la gestion dans les feuilles Imp\_… ; refuse toute période absente de la table Clotures.
- CalculerCouts(dateDebut, dateFin) : remplit les feuilles de résultats Calc\_… décrites plus bas. La partie 4 l'appelle pour chaque scénario, sur une copie des données.
- CoutPeriode(dateDebut, dateFin, categories, objet) : renvoie le coût d'une sélection, utilisé par D4.

**Ce que la partie 3 reçoit et livre**

- Reçoit le socle (11/10) et la table Depenses remplie (14/10). En attendant, elle teste son import sur les tables vides livrées le 11/10 et sur quelques lignes saisies à la main.
- Livre à la partie 4, le 15/10, la structure figée des feuilles Calc\_… et la fonction CalculerCouts.

**Tests de recette**

- Le coût logistique total de 2025 vaut 33 796 058 MAD ; avec un chiffre d'affaires de 245,4 millions, le coût rapporté au CA vaut 13,8 %.
- Contrôle de conservation : coûts répartis + sous-activité + frais de structure = total des dépenses, écart nul.
- Chaque poste correspond au grand livre à moins de 0,5 % près.
- Une assurance annuelle de 12 000 MAD compte pour 1 019 MAD en mars (12 000 × 31 / 365).
- Le traitement d'une commande de 4 lignes coûte 312,12 MAD, et 405,76 MAD si elle est urgente.
- Écart au budget 2025 : 33 796 058 − 32 839 000 = +957 058 MAD, soit +2,9 %.
- Le calcul de l'année entière dure moins d'une minute.

---

## 9. Partie 4 — Décision, alertes, restitution et Power BI

La partie 4 transforme les résultats en décisions et en restitutions : scénarios, alertes et leur suivi, rapports PDF, export des résultats et rapport Power BI. Elle peut commencer tout de suite par Power BI et l'export, sur un fichier de test à la structure figée, puis se brancher sur les vrais résultats dès que la partie 3 les livre.

**Écrans et priorités**

| Code | Écran | Priorité | Contenu |
| --- | --- | --- | --- |
| D19 (export) | Export vers Power BI | Indispensable | Écrit les feuilles Calc\_… en valeurs dans BDD\_Decisionnel\_AAAA-MM-JJ\_hhmm.xlsx |
| P1 | Page Synthèse | Indispensable | Coût total, coût / CA, coût par commande, écart au budget, évolution mensuelle |
| P2 | Page Catégories de coûts | Indispensable | Répartition par type, famille et étape, curseur de simulation de prix |
| D13 | Scénarios | Indispensable | Premiers leviers : variation du prix d'une catégorie, sous-traitance d'une zone |
| D16 | Comparaison et décision | Important | Scénarios côte à côte, contraintes, décision justifiée |
| D17 | Alertes | Important | Les 14 seuils S-01 à S-14, gravité et priorité |
| D18 | Suivi d'une alerte | Important | Statut, action, responsable, échéance |
| D19 (rapports) | Rapports PDF | Important | Rapport mensuel, rapport d'analyse, rapport de scénario, avec vérification de concordance |
| P3 | Page Flotte | Important | Coût et indicateurs par camion |
| P4 | Autres pages | Souhaitable | Entrepôts, prestataires, clients, produits, événements, scénarios |
| D14 | Nouveau produit | Souhaitable | Coût logistique estimé, marge, besoin de stockage |
| D15 | Comparaison des fournisseurs | Souhaitable | Coût rendu, délais, fiabilité |

**Modules** : modScenarios, modAlertes, modRapports, modExportPBI, plus les copies du socle.

**Ce que la partie 4 reçoit et livre**

- Reçoit le socle (11/10), puis de la partie 3 la structure des feuilles Calc\_… et la fonction CalculerCouts (15/10). Avant cette date, elle travaille sur un fichier de test qui respecte exactement la structure de la section suivante.
- Livre à la partie 3, à chaque intégration, ses modules et formulaires ; elle remplit la zone des alertes de l'écran d'accueil D1.

**Tests de recette**

- Scénario « gasoil +15 % » sur la flotte : l'écart est de 2 374 109,43 × 0,15 = +356 116 MAD sur 2025, et les données réelles restent inchangées après le calcul.
- Chaque seuil S-01 à S-14 déclenche une alerte sur un cas construit pour le dépasser, par exemple une zone occupée à plus de 90 %.
- L'export vers Power BI est refusé tant qu'une alerte critique, comme un écart avec le grand livre, n'est pas traitée.
- Le rapport PDF passe la vérification de concordance : chaque montant est égal à celui de l'application.
- Power BI affiche 33 796 058 MAD pour 2025, le même total qu'Excel.

---

## 10. Structures des tables à créer

Ces structures servent de contrat entre les parties : chacun peut créer des données de test conformes sans attendre les autres. Les colonnes sont listées dans l'ordre, en ligne 1 de chaque feuille, et chaque plage est un tableau Excel nommé.

| Tableau (feuille) | Responsable | Colonnes, dans l'ordre |
| --- | --- | --- |
| T\_TypesCouts (Types\_Couts) | Partie 1 | Code\_type (TC-01) · Libelle · Famille · Etape (Approvisionnement, Stockage, Traitement, Transport, Structure) · Compte · Nature (Fixe, Variable) · Obtention (Facture, Quantite\_prix, Calcule) · Rattachement (Direct, Partage, Structure) · Type\_objet · Cle\_partage · Repartition (Ponctuelle, Etalee) · Actif |
| T\_Listes (Listes) | Partie 1 | Liste · Code · Libelle · Attribut\_1 (gravité par défaut, taux de change…) · Attribut\_2 (immobilisation…) · Actif |
| T\_Utilisateurs (Utilisateurs) | Partie 1 | Identifiant · Nom · Profil (Administrateur, Analyste, Consultation) · Mot\_de\_passe · Actif · Derniere\_connexion |
| T\_Journal (Journal) | Partie 1 | Date\_heure · Utilisateur · Action · Objet · Detail |
| T\_Depenses (Depenses) | Partie 2 | N\_depense (DEP-2025-00001) · Statut (Saisie, Validee, Annulee) · Date\_piece · Code\_type · Code\_tiers · Type\_objet · Code\_objet · N\_evenement · Type\_document · N\_document · Fichier · Debut\_couverture · Fin\_couverture · Quantite · Prix\_unitaire · Montant\_devise · Devise · Taux\_change · Montant\_HT\_MAD · Taux\_TVA · Commentaire · Origine (Saisie, ou Reprise GL n° …) · Cree\_par · Cree\_le |
| T\_Evenements (Evenements) | Partie 2 | N\_evenement (EVT-0001) · Date\_heure · Type · Gravite (Mineure, Majeure, Critique) · Type\_objet · Code\_objet · Description · Debut\_immobilisation · Fin\_immobilisation · Jours\_immobilisation · Statut (Ouvert, Clos) · Date\_cloture |
| T\_Commentaires (Commentaires) | Partie 2 | N\_commentaire · Table\_cible · Id\_cible · Date\_heure · Utilisateur · Texte |
| T\_ResultatsControles (Resultats\_Controles) | Partie 2 | N\_anomalie · Date\_controle · Controle · Feuille · Identifiant · Description · Gravite · Etat (A\_traiter, Corrigee) |
| T\_Clotures (Clotures) | Partie 2 | Periode (AAAA-MM) · Date\_cloture · Utilisateur · Sauvegarde |
| Calc\_Depenses | Partie 3 | Mois · N\_depense · Code\_type · Libelle\_type · Famille · Etape · Compte · Type\_objet · Code\_objet · Site · Montant\_MAD · Nature\_montant (Enregistre, Calcule) |
| Calc\_Missions | Partie 3 | N\_mission · Date · Mode (Interne, Sous\_traite) · Code\_vehicule · Code\_prestataire · Km · Cout\_fixe · Cout\_km · Frais · Cout\_total · Taux\_remplissage |
| Calc\_Commandes | Partie 3 | N\_commande · Date · Code\_client · Segment · Site · CA\_HT · Cout\_transport · Cout\_traitement · Cout\_stockage · Cout\_complet · Marge\_apres\_logistique |
| Calc\_Lignes | Partie 3 | N\_commande · N\_ligne · Code\_produit · Code\_client · Colis · Cout\_transport · Cout\_traitement · Cout\_stockage · Cout\_total |
| Calc\_Stockage | Partie 3 | Mois · Site · Zone · Code\_produit · Palettes\_moyennes · Cout\_occupation · Cout\_detention · Cout\_sous\_activite |
| Indicateurs | Partie 3 | Mois · Code\_indicateur · Perimetre · Valeur · Objectif · Etat |

La partie 1 ajoute aussi dans Ref\_Data les seuils qui manquent (SEUIL\_RETARD\_LIVRAISON 1 jour, SEUIL\_RETARD\_FOURN 3 jours, MARGE\_SECURITE\_STOCK 3 jours, SEUIL\_TAUX\_PERTES 0,5 %, SEUIL\_MARGE\_MIN 5 %) et, dans T\_Compteurs, les compteurs des dépenses, des événements et des types de coûts.

---

## 11. La base Atlantis Distribution 2025 : les 43 feuilles

Le classeur fourni avec le projet contient 43 feuilles et environ 200 000 lignes, issues d'une simulation jour par jour de l'année 2025. Chaque plage est aussi un tableau Excel nommé `T_…`, directement lisible par Power BI ; les en-têtes sont en ligne 1 et les données commencent en ligne 2. Les quatre premières feuilles sont la documentation : LISEZMOI, Dictionnaire (description de chaque colonne), Anomalies (les six anomalies volontaires) et Controles (15 contrôles de cohérence, tous au statut OK).

| Feuille | Tableau | Lignes de données | Colonnes |
|---------------|---------------------|------|---------------------------------------------------|
| LISEZMOI | — | 107 | Base de données logistique – ATLANTIS DISTRIBUTION SA (société fictive) – Exercice 2025 |
| Dictionnaire | T_Dictionnaire | 423 | Feuille · Colonne · Type · Description · Cle |
| Anomalies | T_Anomalies | 6 | Code · Table · Identifiant · Anomalie · Controle_a_mettre_en_place |
| Controles | T_Controles | 17 | N° · Domaine · Contrôle · Valeur A · Valeur B · Écart A − B · Écart attendu · Statut · Commentaire |
| Ref_Data | T_Parametres + T_Compteurs | 42 | Code_parametre · Libelle · Valeur · Unite · Commentaire · Compteur · Derniere_valeur · Prochaine_valeur |
| Calendrier | T_Calendrier | 365 | Date · Jour · Semaine_ISO · Mois · Libelle_mois · Trimestre · Jour_ouvre · Ferie · Periode_commerciale |
| Taux_Change | T_Taux_Change | 24 | Mois · Devise · Taux_MAD |
| Prix_Carburant | T_Prix_Carburant | 12 | Mois · Prix_gasoil_TTC · Taux_TVA · Prix_gasoil_HT |
| Sites | T_Sites | 3 | Code_site · Nom · Ville · Type_site · Latitude · Longitude · Surface_totale_m2 · Surface_communs_m2 · Loyer_HT_m2_mois · Date_ouverture · Responsable · Horaires |
| Zones_Stockage | T_Zones_Stockage | 9 | Code_zone_stockage · Site · Regime · Libelle · Temperature_consigne · Capacite_emplacements_palettes · Surface_m2 · Centre_de_cout |
| Zones_Livraison | T_Zones_Livraison | 19 | Code_zone · Libelle · Site · Mode_livraison · Jours_tournee · Type_tournee · PTAC_max_t |
| Familles | T_Familles | 7 | Code_famille · Libelle · Regime_dominant · Taux_obsolescence_annuel |
| Produits | T_Produits | 140 | Code_produit · Designation · Code_famille · Sous_famille · Regime_temperature · Unite_vente_UVC · UVC_par_colis · Poids_brut_colis_kg · Longueur_colis_cm · Largeur_colis_cm · Hauteur_colis_cm · Volume_colis_m3 · Colis_par_couche · Couches_par_palette · Colis_par_palette · Poids_palette_kg · Code_fournisseur · Pays_origine · Devise_achat · Prix_achat_colis_devise · Taux_droit_douane · Cout_standard_colis_MAD · Prix_vente_tarif_HT_MAD · Hausse_tarif_01_07 · DLC_DLUO_jours · Code_SH · Produit_export · Date_creation · Statut · Date_arret |
| Fournisseurs | T_Fournisseurs | 35 | Code_fournisseur · Raison_sociale · Pays · Ville · Devise · Incoterm · Mode_approvisionnement · Port_chargement · Delai_appro_jours · Periodicite_commande · Sites_livres · Taux_service_contractuel · Conditions_paiement · Contact |
| Clients | T_Clients | 158 | Code_client · Raison_sociale · Segment · Libelle_segment · Canal_commande · Ville · Region · Pays · Latitude · Longitude · Site_rattachement · Code_zone · Distance_site_km · Part_autoroute · Zone_tarifaire_messagerie · Mode_livraison · Creneau_debut · Creneau_fin · RDV_obligatoire · Hayon_requis · Minimum_commande_HT · Franco_port_HT · Frais_port_factures_HT · Remise_negociee · Incoterm · Lieu_incoterm · Devise_facturation · ICE · Date_creation · Statut |
| Vehicules | T_Vehicules | 27 | Code_vehicule · Immatriculation · Type_vehicule · Site_affectation · PTAC_t · Charge_utile_kg · Volume_utile_m3 · Palettes_max · Frigorifique · Hayon · Consommation_L_100km · Capacite_reservoir_L · Classe_peage · Valeur_acquisition_HT · Date_mise_en_service · Duree_amortissement_ans · Assurance_annuelle_HT · Taxe_essieu_vignette_annuelle · Km_compteur_01_01_2025 · Date_sortie_flotte · Statut |
| Personnel | T_Personnel | 98 | Matricule · Nom_complet · Fonction · Site · Centre_de_cout · Salaire_brut_mensuel · Date_entree · Date_sortie · Type_contrat · Vehicule_habituel |
| Prestataires | T_Prestataires | 18 | Code_prestataire · Raison_sociale · Activite · Pays · Devise · Mode_facturation |
| Grille_Messagerie | T_Grille_Messagerie | 90 | Code_prestataire · Sites_origine · Zone_tarifaire · Poids_min_kg · Poids_max_kg · Type_tarif · Tarif · Minimum_perception · Date_debut · Date_fin |
| Grille_Affretement | T_Grille_Affretement | 254 | Code_prestataire · Site_origine · Destination · Type_vehicule · Regime · Forfait_HT_MAD · Date_debut · Date_fin |
| Tarifs_International | T_Tarifs_International | 25 | Code_prestataire · Origine · Destination · Mode · Devise · Tarif · Unite · Date_debut · Date_fin |
| Surcharge_Gazole | T_Surcharge_Gazole | 72 | Mois · Code_prestataire · Taux_surcharge |
| Natures_Frais | T_Natures_Frais | 30 | Code_nature · Libelle · Categorie_analytique · Inclus_cout_logistique · Compte_CGNC |
| Centres_Couts | T_Centres_Couts | 26 | Code_centre · Libelle · Site · Poste_budgetaire · Inducteur_suggere |
| Commandes | T_Commandes | 6474 | N_commande · Date_commande · Heure_commande · Code_client · Site_expedition · Canal_commande · Type_commande · Mode_livraison · Date_expedition_prevue · Date_livraison_prevue · Date_expedition_reelle · Date_livraison_effective · Statut · Motif_annulation · Incoterm · Devise_facturation · Taux_change_facturation · Frais_port_factures_HT · Duree_saisie_ADV_min · Duree_preparation_min |
| Lignes_Commandes | T_Lignes_Commandes | 44004 | N_commande · N_ligne · Code_produit · Qte_commandee · Qte_expediee · Qte_retournee · Qte_livree_nette · Prix_unitaire_HT · Remise · Montant_net_HT · Promotion · Motif_ecart |
| Missions | T_Missions | 3923 | N_mission · Type_mission · Mode_execution · Site_depart · Code_zone · Date_depart · Heure_depart · Date_retour · Heure_retour · Code_vehicule · Matricule_chauffeur · Matricule_aide · Code_prestataire · Type_vehicule_affrete · Regime_temperature · Km_compteur_depart · Km_compteur_retour · Distance_estimee_km · Peages_HT · Nuitees_hors_base · Nb_arrets · N_transfert · Destination_principale · Statut · Commentaire |
| Arrets | T_Arrets | 7072 | N_mission · Rang · Type_arret · Type_document · N_document · Partie_commande · Code_tiers · Ville · Km_depuis_arret_precedent · Date_arrivee · Heure_arrivee · Heure_depart · Statut · Motif · Tentative · Dans_creneau |
| Frais_Transport_Ext | T_Frais_Transport_Ext | 2384 | N_ligne_frais · N_mission · N_commande · Code_prestataire · N_facture · Date_facture · Code_nature · Libelle · Quantite · Unite · Montant_devise · Devise · Taux_change · Montant_MAD_HT |
| Transferts | T_Transferts | 400 | N_transfert · Date_depart · Date_arrivee · Site_origine · Site_destination · Regime · N_mission · Statut |
| Lignes_Transferts | T_Lignes_Transferts | 14463 | N_transfert · Code_produit · Qte_colis |
| Retours | T_Retours | 150 | N_retour · Date_retour_entrepot · Site · N_commande · Code_client · Code_produit · Quantite_colis · Motif · Origine · Decision · N_mission_retour · Avoir_emis · Montant_avoir_HT |
| Cmd_Fournisseurs | T_Cmd_Fournisseurs | 2399 | N_commande_fournisseur · Date_commande · Code_fournisseur · Site_livraison · Incoterm · Mode_transport · Type_conteneur · Nb_conteneurs · Port_chargement · Port_dechargement · Devise · Date_depart_ETD · Date_arrivee_prevue_ETA · Date_arrivee_reelle · Date_mainlevee_douane · Date_reception_entrepot · Circuit_douane · N_DUM · N_mission_enlevement · Statut |
| Lignes_Cmd_Fourn | T_Lignes_Cmd_Fourn | 7398 | N_commande_fournisseur · N_ligne · Code_produit · Qte_commandee · Qte_recue · Qte_non_conforme · Prix_unitaire_devise · Motif_ecart |
| Frais_Appro | T_Frais_Appro | 1385 | N_ligne_frais · N_commande_fournisseur · Code_nature · Libelle · Categorie_analytique · Inclus_cout_logistique · Code_prestataire · N_facture · Date_facture · Quantite · Unite · Montant_devise · Devise · Taux_change · Montant_MAD_HT |
| Stock_Initial | T_Stock_Initial | 400 | Site · Code_produit · Stock_initial_colis |
| Mouvements_Stock | T_Mouvements_Stock | 79960 | N_mouvement · Date · Site · Code_produit · Type_mouvement · Quantite_colis · Type_document · N_document |
| Stocks_Hebdo | T_Stocks_Hebdo | 21200 | Date_photo · Site · Code_produit · Stock_colis |
| Inventaires | T_Inventaires | 490 | Date_inventaire · Type_inventaire · Site · Code_produit · Stock_theorique · Stock_compte · Ecart |
| Carburant | T_Carburant | 1360 | N_ticket · Date · Heure · Code_vehicule · N_carte_carburant · Station · Litres · Prix_litre_TTC · Montant_TTC · TVA · Montant_HT · Km_compteur |
| Maintenance | T_Maintenance | 184 | N_intervention · Code_vehicule · Date_debut · Date_fin · Type_intervention · Description · Km_compteur · Code_prestataire · Cout_pieces_HT · Cout_main_oeuvre_HT · Cout_total_HT · Jours_immobilisation |
| Charges | T_Charges | 4169 | N_ecriture · Mois · Compte_CGNC · Libelle_compte · Centre_de_cout · Nature_charge · Montant_HT_MAD · Origine |
| Budget | T_Budget | 204 | Mois · Poste_budgetaire · Site · Montant_budget_HT_MAD |

---

## 12. Paramètres et compteurs de la feuille Ref_Data

Ces valeurs sont lues par `LireParametre` ; aucune ne doit être écrite en dur dans le code. Les deux dernières lignes (TAUX_STD_EUR et TAUX_STD_USD) ont été ajoutées par le groupe, elles manquaient dans la base d'origine.

| Code | Libellé | Valeur | Unité |
|------------------------|----------------------------------------------------------|------------|------------------------|
| DEVISE | Devise de référence | MAD | — |
| PERIODE_DEBUT | Début de la période d'analyse | 01/01/2025 | date |
| PERIODE_FIN | Fin de la période d'analyse | 31/12/2025 | date |
| TVA | Taux de TVA standard | 0,2 | % |
| TAUX_CAPITAL | Coût du capital immobilisé (r) | 0,075 | %/an |
| TAUX_ASSUR_STOCK | Taux d'assurance du stock | 0,004 | %/an |
| TAUX_OBSOL_DEFAUT | Taux d'obsolescence par défaut | 0,02 | %/an |
| CHARGES_SOCIALES | Taux de charges patronales (CNSS, AMO, TFP) | 0,2109 | % |
| MAJ_URGENCE | Majoration du coût de traitement d'une commande urgente (μ) | 0,3 | % |
| RHO_TAXABLE | Facteur de conversion poids / volume (poids taxable) | 250 | kg/m³ |
| VOL_PALETTE | Volume utile d'une palette de préparation | 1,25 | m³ |
| POIDS_PALETTE_EQ | Poids équivalent d'une palette de préparation | 850 | kg |
| NUITEE_CHAUFFEUR | Indemnité de nuitée chauffeur | 250 | MAD/nuit |
| NUITEE_AIDE | Indemnité de nuitée aide-livreur | 200 | MAD/nuit |
| INTERIM_CHAUFFEUR | Coût journalier chauffeur intérimaire | 480 | MAD/jour |
| INTERIM_PREPA | Coût journalier préparateur intérimaire | 230 | MAD/jour |
| PRIX_KWH | Prix de l'électricité HT | 1,25 | MAD/kWh |
| TOLL_C1 | Tarif péage indicatif classe 1 (VUL) | 0,28 | MAD/km autoroute |
| TOLL_C2 | Tarif péage indicatif classe 2 (porteurs 2 essieux) | 0,46 | MAD/km autoroute |
| TOLL_C3 | Tarif péage indicatif classe 3 (≥ 3 essieux) | 0,68 | MAD/km autoroute |
| A_ADV_REF | Coût administratif par commande (a) – valeur de référence 2025 | 126,6 | MAD/commande |
| B_PREP_REF | Coût de préparation par ligne (b) – valeur de référence 2025 | 46,38 | MAD/ligne |
| CMV_REF | Coût d'un mouvement d'entrée en stock (c_mv) – valeur de référence 2025 | 20,72 | MAD/mouvement |
| E_CAS_AMB | Coût mensuel d'un emplacement palette CAS AMB (e) – référence 2025 | 159,77 | MAD/emplacement/mois |
| E_CAS_FRA | Coût mensuel d'un emplacement palette CAS FRA (e) – référence 2025 | 352,57 | MAD/emplacement/mois |
| E_CAS_SUR | Coût mensuel d'un emplacement palette CAS SUR (e) – référence 2025 | 382,18 | MAD/emplacement/mois |
| E_TNG_AMB | Coût mensuel d'un emplacement palette TNG AMB (e) – référence 2025 | 360,25 | MAD/emplacement/mois |
| E_TNG_FRA | Coût mensuel d'un emplacement palette TNG FRA (e) – référence 2025 | 542,91 | MAD/emplacement/mois |
| E_TNG_SUR | Coût mensuel d'un emplacement palette TNG SUR (e) – référence 2025 | 604,73 | MAD/emplacement/mois |
| E_AGA_AMB | Coût mensuel d'un emplacement palette AGA AMB (e) – référence 2025 | 284,53 | MAD/emplacement/mois |
| E_AGA_FRA | Coût mensuel d'un emplacement palette AGA FRA (e) – référence 2025 | 464,53 | MAD/emplacement/mois |
| E_AGA_SUR | Coût mensuel d'un emplacement palette AGA SUR (e) – référence 2025 | 532,81 | MAD/emplacement/mois |
| SEUIL_REMPLISSAGE | Seuil d'alerte : taux de remplissage minimal d'une tournée | 0,6 | % |
| SEUIL_PI_CLIENT | Seuil d'alerte : coût logistique / CA client maximal | 0,12 | % |
| SEUIL_COUVERTURE | Seuil d'alerte : couverture de stock maximale | 90 | jours |
| SEUIL_DORMANT | Seuil d'alerte : jours sans sortie (stock dormant) | 60 | jours |
| SEUIL_OCCUPATION | Seuil d'alerte : taux d'occupation maximal d'une zone | 0,9 | % |
| SEUIL_ECART_BUDGET | Seuil d'alerte : écart au budget mensuel | 0,1 | % |
| SEUIL_ECART_FACTURE | Seuil d'alerte : écart facture transporteur / grille | 0,03 | % |
| SEUIL_TAUX_SERVICE | Objectif de taux de service (quantités) | 0,97 | % |
| TAUX_STD_EUR | Taux standard de valorisation EUR | 10,58 | MAD/EUR |
| TAUX_STD_USD | Taux standard de valorisation USD | 9,45 | MAD/USD |

**Compteurs (table T_Compteurs, colonnes H à J de Ref_Data)**

| Compteur | Dernière valeur | Prochaine valeur |
|-----------------------------------|------------|------------|
| Dernier N° de commande client | 256474 | 256475 |
| Dernier N° de mission | 603923 | 603924 |
| Dernier N° de transfert | 700400 | 700401 |
| Dernier N° de commande fournisseur | 802399 | 802400 |
| Dernier N° de retour | 900150 | 900151 |
| Dernier N° de mouvement de stock | 79960 | 79961 |
| Dernier code produit | P0140 | P0141 |
| Dernier code client | CL0158 | CL0159 |

---

## 13. Les 20 formules de calcul

Les formules ci-dessous sont celles qu'applique l'application décisionnelle et qu'affiche l'écran des méthodes de calcul (D12). Chaque symbole est défini juste après la formule, et chaque exemple peut être refait à partir des valeurs des captures citées. Les montants sont hors taxes ; le coût d'achat des marchandises, les droits de douane et la TVA récupérable ne sont jamais comptés dans les coûts logistiques.

**Coûts**

**1. Montant d'une dépense rattaché à une période**

$$M_p = M \times \frac{j_p}{j_c}$$

où $M$ est le montant de la dépense, $j_c$ le nombre de jours couverts par la dépense et $j_p$ le nombre de ces jours situés dans la période analysée. Une dépense ponctuelle est rattachée en entier au jour de sa pièce.

*Exemple (T7, D3)* : une assurance annuelle de 12 000 MAD compte en mars pour 12 000 × 31 / 365 = 1 019 MAD.

**2. Coût logistique total d'une période**

$$C = \sum_{k} M_{p,k}$$

où $M_{p,k}$ est le montant rattaché à la période de la dépense $k$, qu'elle relève du transport, du traitement des commandes, du stockage, de l'approvisionnement, de la structure ou des pertes ; chaque dépense n'est comptée qu'une fois.

*Exemple (D1, P1)* : pour l'exercice 2025, $C$ = 33 796 058 MAD.

**3. Coût par kilomètre d'un camion**

$$c_{km} = \frac{A + \sum D_v}{K}$$

où $A$ est l'amortissement du camion sur la période (sa valeur d'achat divisée par sa durée d'amortissement, au prorata de la période), $\sum D_v$ la somme de ses dépenses de la période (carburant, maintenance, assurance, chauffeur, péages…) et $K$ le nombre de kilomètres qu'il a parcourus.

*Exemple (D4, D7)* : le camion VEH-07 a coûté 58 214 MAD en mars pour 11 856 km, soit 58 214 / 11 856 = 4,91 MAD par kilomètre.

**4. Part du coût d'une mission imputée à une commande**

$$C_{m,i} = C_m \times \frac{P_i}{\sum_{j} P_j}$$

où $C_m$ est le coût de la mission, $P_i$ le poids taxable de la commande $i$, c'est-à-dire le plus grand de son poids réel et de son volume multiplié par le coefficient volumique fixé dans les paramètres (T13), et $\sum_j P_j$ le poids taxable de toutes les commandes transportées par la mission.

*Exemple (D7, D10)* : la mission 603921 a coûté 1 082,40 MAD ; la commande 256475 y pèse 430,1 kg sur 1 804 kg, et reçoit 1 082,40 × 430,1 / 1 804 = 258,06 MAD, soit 23,84 % du coût de la mission.

**5. Coût de traitement d'une commande**

$$C_t = (F + n \times L) \times (1 + u)$$

où $F$ est le coût fixe par commande (126,60 MAD), $L$ le coût par ligne (46,38 MAD), $n$ le nombre de lignes de la commande et $u$ la majoration d'urgence, égale à 30 % pour une commande urgente et à 0 sinon.

*Exemple (D10)* : la commande 256475, normale, compte 4 lignes : 126,60 + 4 × 46,38 = 312,12 MAD.

**6. Coût de stockage d'un produit**

$$C_s = \bar{Q} \times c_e \times m + \bar{V} \times \tau \times \frac{j}{365}$$

où $\bar{Q}$ est le nombre d'emplacements occupés en moyenne par le produit, $c_e$ le coût mensuel d'un emplacement de sa zone, $m$ le nombre de mois de la période, $\bar{V}$ la valeur moyenne de son stock, $\tau$ le taux annuel de détention (coût du capital, assurance et obsolescence de sa famille) et $j$ le nombre de jours de la période. Le premier terme est le coût d'occupation, le second le coût de détention.

*Exemple (D8, D12)* : pour le lait UHT en mars, 38,2 × 352,10 × 1 + 182 000 × 14,3 % × 31 / 365 = 13 450 + 2 210 = 15 660 MAD.

**7. Coût d'un emplacement et coût de sous-activité d'une zone**

$$c_e = \frac{K_z}{N_z} \qquad S_z = c_e \times (N_z - \bar{Q}_z)$$

où $K_z$ représente les charges mensuelles de la zone (loyer, énergie, amortissements et part des charges du site), $N_z$ sa capacité en emplacements et $\bar{Q}_z$ le nombre d'emplacements occupés en moyenne ; $S_z$ est le coût des emplacements payés mais vides, présenté à part.

*Exemple (D8)* : pour la zone des produits frais de Casablanca, 83 096 / 236 = 352,10 MAD par emplacement, et $S_z$ = 352,10 × (236 − 214) = 7 746 MAD.

**8. Frais d'approche par colis importé**

$$a = \frac{F_r + P_o + T_r + A_t}{N_c}$$

où $F_r$, $P_o$, $T_r$ et $A_t$ sont respectivement le fret, les frais portuaires, les frais de transit et l'assurance transport de la période, et $N_c$ le nombre de colis importés reçus ; les droits de douane sont présentés à part.

*Exemple (D5, D14)* : en mars, 148 224 / 12 800 = 11,58 MAD par colis.

**9. Coût complet d'une commande**

$$C_c = C_{m,i} + C_t + C_{s,i}$$

où $C_{m,i}$ est la part de transport (formule 4), $C_t$ le coût de traitement (formule 5) et $C_{s,i}$ le coût de stockage des produits de la commande, au prorata des colis sortis. La sous-activité et les frais de structure ne sont pas imputés aux commandes.

*Exemple (D10)* : commande 256475 : 258,06 + 312,12 + 96,85 = 667,03 MAD, soit 667,03 / 135 = 4,94 MAD par colis.

**Indicateurs**

**10. Coût unitaire**

$$c = \frac{C_P}{Q}$$

où $C_P$ est le coût du périmètre analysé et $Q$ la quantité de référence, toujours affichée avec le résultat : commande livrée, ligne, colis, tonne, kilomètre ou tonne-kilomètre.

*Exemple (D1, D5)* : en janvier et février, 5 482 960 / 1 046 = 5 242 MAD par commande livrée.

**11. Coût rapporté au chiffre d'affaires**

$$r = \frac{C}{\mathit{CA}}$$

où $C$ est le coût logistique et $\mathit{CA}$ le chiffre d'affaires hors taxes du même périmètre et de la même période.

*Exemple (D1, D5, D9)* : en 2025, 33 796 058 / 245 400 000 = 13,8 %.

**12. Taux de remplissage d'un véhicule**

$$\rho = \max\left(\frac{P}{P_{\max}}\,;\ \frac{V}{V_{\max}}\,;\ \frac{N}{N_{\max}}\right)$$

où $P$, $V$ et $N$ sont le poids, le volume et le nombre de palettes chargés, et $P_{\max}$, $V_{\max}$ et $N_{\max}$ les capacités correspondantes du véhicule, inscrites sur sa fiche (T3) ; la contrainte la plus serrée l'emporte.

*Exemple (T12, D7)* : 1,65 t sur 3,5 t, 13 m³ sur 30 m³ et 7 palettes sur 12 donnent max(47 % ; 43 % ; 58 %) = 58 %.

**13. Rotation et couverture du stock**

$$R = \frac{S}{\bar{Q}_s} \qquad \mathit{Cv} = \frac{\bar{Q}_s}{S / j}$$

où $S$ représente les sorties de la période en colis, $\bar{Q}_s$ le stock moyen en colis et $j$ le nombre de jours de la période ; la couverture $\mathit{Cv}$ s'exprime en jours de sorties.

*Exemple (D8)* : pour le lait UHT en mars, $R$ = 3 740 / 2 650 = 1,4 et $\mathit{Cv}$ = 2 650 / (3 740 / 31) = 22 jours.

**14. Occupation moyenne et occupation de pointe d'une zone**

$$o = \frac{\bar{Q}_z}{N_z} \qquad o_{\max} = \frac{Q_{z,\max}}{N_z}$$

où $\bar{Q}_z$ est le nombre d'emplacements occupés en moyenne, $Q_{z,\max}$ le plus haut nombre d'emplacements occupés relevé dans la période et $N_z$ la capacité de la zone.

*Exemple (D8, D17)* : pour la zone des produits frais de Casablanca en mars, 214 / 236 = 91 % en moyenne et 229 / 236 = 97 % en pointe, pour un seuil de 90 %.

**15. Délai, retard et ponctualité**

$$D = t_l - t_c \qquad R_t = \max(0\,;\ t_l - t_s) \qquad \pi = \frac{n(R_t = 0)}{n}$$

où $t_c$, $t_l$ et $t_s$ sont les dates de commande, de livraison réelle et de livraison souhaitée, $n$ le nombre de commandes livrées et $n(R_t = 0)$ le nombre de celles livrées sans retard.

*Exemple (D5, D10)* : commande 256475 : $D$ = 05/03 − 04/03 = 1 jour et $R_t$ = max(0 ; 05/03 − 05/03) = 0 jour.

**16. Taux de pertes**

$$\tau_p = \frac{V_p}{V_s}$$

où $V_p$ est la valeur des pertes constatées (casse, produits périmés, écarts d'inventaire et retours non revendables) et $V_s$ la valeur d'achat des sorties de la période.

*Exemple (D8, T13)* : en mars, 18 420 / 16 960 000 = 0,11 %, sous le seuil de 0,5 %.

**17. Marge après logistique**

$$\mu = \mathit{CA} - A_v - C_c \qquad \text{taux} = \frac{\mu}{\mathit{CA}}$$

où $\mathit{CA}$ est le chiffre d'affaires hors taxes, $A_v$ le coût d'achat des produits vendus et $C_c$ le coût logistique complet imputé (formule 9).

*Exemple (D9, D10)* : commande 256475 : 15 545,00 − 12 746,90 − 667,03 = 2 131,07 MAD, soit 13,7 % du montant.

**Écarts et statistiques**

**18. Écart au budget**

$$E = C - B \qquad e = \frac{E}{B}$$

où $C$ est le coût réel et $B$ le budget de la même période et du même périmètre.

*Exemple (D1, D4, P1)* : en 2025, 33 796 058 − 32 839 000 = 957 058 MAD, soit +2,9 %.

**19. Effet volume et effet coût unitaire**

$$\Delta C = (Q_1 - Q_0) \times c_0 + (c_1 - c_0) \times Q_1$$

où $Q_0$ et $Q_1$ sont l'activité de la période de référence et de la période analysée (kilomètres, commandes, colis…), et $c_0$ et $c_1$ les coûts unitaires correspondants. Le premier terme est l'effet volume, le second l'effet coût unitaire ; leur somme est égale à l'écart total.

*Exemple (D4)* : flotte de Casablanca, mars comparé à février : (124 140 − 122 230) × 4,6859 = 8 950 MAD et (4,7200 − 4,6859) × 124 140 ≈ 4 230 MAD, soit un écart total de 13 180 MAD.

**20. Coefficient de variation**

$$\mathit{CV} = \frac{\sigma}{\bar{x}}$$

où $\sigma$ est l'écart type et $\bar{x}$ la moyenne de l'indicateur sur la population étudiée ; plus il est élevé, plus les valeurs sont dispersées.

*Exemple (D11)* : coût par commande de janvier à mars : 4 579 / 5 011 = 0,91.

---

## 14. Spécification des écrans, élément par élément

Les 41 écrans et pages de la solution — T1 à T18 pour l'application de gestion, D1 à D19 pour l'application décisionnelle, P1 à P4 pour Power BI — sont décrits ici un par un : code, titre, manière d'y accéder, profils autorisés, rôle, puis la définition de chacun de leurs éléments. Les numéros entre parenthèses sont les repères portés sur les captures de la maquette (fichier `Cahier_des_charges_et_maquette_Projet15.pdf`, 75 pages).

### Volet transactionnel : l'application de gestion

Le volet transactionnel est la partie de la solution qui enregistre l'activité. Il prend la forme d'une application de gestion, outil quotidien des personnes qui saisissent les informations : le contrôleur de gestion, les assistants de l'administration des ventes, le responsable transport ou le gestionnaire de la flotte. Elle doit leur permettre de créer, consulter, modifier, rechercher et supprimer, simplement et sans risque d'erreur, toutes les informations dont l'analyse des coûts aura besoin : fiches de référence, types de coûts, dépenses, événements, commentaires, commandes et missions. Elle contrôle la qualité de ces informations et permet de clôturer une période lorsqu'elles sont vérifiées, mais ne calcule aucun indicateur. Ses fonctions sont regroupées en huit domaines, présentés ci-dessous avec les écrans qui les réalisent (T1 à T18).

### A. Connexion, accueil et navigation

L'utilisateur s'identifie par un nom et un mot de passe ; son profil détermine les fonctions auxquelles il a accès. Il arrive ensuite directement sur un menu principal qui regroupe les fonctions par domaine : fiches de référence, opérations courantes, contrôles et clôture, administration. Les feuilles de données restent masquées pour éviter toute modification accidentelle. Chaque écran propose un retour au menu, et un bandeau rappelle l'exercice en cours, le profil connecté, le nombre d'anomalies à traiter et la date de la dernière clôture. Un bouton permet enfin d'ouvrir directement l'application décisionnelle.

#### T1 — Connexion

*Accès : ouverture du classeur. Profils : tous.*

Premier écran affiché à l'ouverture de l'application transactionnelle. Il identifie l'utilisateur et détermine son profil, qui fixe ensuite les fonctions accessibles dans le menu. Tant que la connexion n'est pas réussie, aucune feuille n'est visible.

Éléments de l'écran :

- (1) **Identité de l'application** (bandeau) : Rappelle le nom de l'entreprise, le rôle de l'application et l'exercice en cours, pour que l'utilisateur sache dans quel outil il se trouve.
- (2) **Identifiant** (zone de saisie) : Nom de connexion de l'utilisateur. Champ obligatoire, non sensible à la casse ; les espaces en début et en fin sont ignorés.
- (3) **Mot de passe** (zone de saisie masquée) : Les caractères sont remplacés par des points à la frappe. Champ obligatoire ; la touche Entrée lance la connexion.
- (4) **Profil reconnu** (zone calculée) : Affiche le profil associé à l'identifiant (administrateur, analyste ou consultation) dès que celui-ci est reconnu. Non modifiable.
- (5) **Message d'erreur** (message) : Apparaît en rouge si l'identifiant ou le mot de passe est faux, avec le nombre de tentatives. Après trois échecs, l'application se ferme.
- (6) **Se connecter** (bouton principal) : Vérifie le couple identifiant et mot de passe, inscrit la connexion dans le journal puis ouvre le menu principal avec les droits du profil.
- (7) **Quitter** (bouton) : Ferme le classeur sans rien enregistrer.


#### T2 — Menu principal

*Accès : après la connexion, ou bouton « Menu » de chaque écran. Profils : tous (boutons filtrés selon le profil).*

Feuille d'accueil de l'application transactionnelle. Elle regroupe toutes les fonctions par domaine sous forme de boutons et affiche en permanence l'état de l'application. Les feuilles de données restent masquées : toute action passe par ce menu.

Éléments de l'écran :

- (1) **Bandeau d'état** (bandeau) : Affiche la société, l'exercice, le profil connecté, le nombre d'anomalies restant à traiter et la date de la dernière clôture. Il est mis à jour à chaque retour au menu.
- (2) **Fiches de référence** (groupe de boutons) : Un bouton par fiche : chacun ouvre le formulaire de gestion correspondant (écrans T3 et T4).
- (3) **Opérations courantes** (groupe de boutons) : Accès à la saisie et à la recherche des dépenses (T7, T8), aux événements (T9), aux commandes (T11), aux missions (T12) et aux commandes fournisseurs.
- (4) **Paramétrage** (groupe de boutons) : Accès aux types de coûts (T5), aux autres listes de choix (T6) et aux paramètres de calcul (T13). Ce dernier bouton est grisé pour un profil autre qu'administrateur.
- (5) **Contrôles et clôture** (groupe de boutons) : Lance le contrôle des données (T14) et ouvre la clôture de période (T15), réservée à l'administrateur.
- (6) **Administration** (groupe de boutons) : Gestion des utilisateurs, du journal, des sauvegardes et de la réinitialisation (T16). Boutons actifs pour l'administrateur seulement.
- (7) **Ouvrir l'application décisionnelle** (bouton) : Ouvre le second classeur, consacré au calcul et à l'analyse des coûts, sans fermer celui-ci.
- (8) **Aide** (bouton) : Ouvre la feuille d'aide : rôle de chaque bouton, parcours type et signification des alertes.
- (9) **Se déconnecter** (bouton) : Enregistre le classeur, inscrit la déconnexion dans le journal et revient à l'écran de connexion.


### B. Fiches de référence

L'application permet de créer, consulter, modifier, rechercher et supprimer les fiches des fournisseurs, des prestataires, des camions, des entrepôts et de leurs zones, du personnel, des produits, des clients et des zones de livraison, ainsi que les grilles tarifaires des transporteurs. Chaque fiche contient les informations utiles au calcul des coûts : la fiche d'un fournisseur indique son pays, sa devise et ses conditions de paiement ; celle d'un camion, ses capacités, son mode de détention, sa valeur et sa consommation de référence ; celle d'un entrepôt, ses zones et leur capacité en palettes ; celle d'un salarié, sa fonction, son site et son coût mensuel ; celle d'un client, sa zone de livraison, sa distance et son segment. Toutes les fiches suivent les mêmes règles : code attribué automatiquement, contrôle des champs obligatoires avec un message clair, recherche et tri, et désactivation plutôt que suppression d'un élément qui possède déjà un historique. Chaque fiche peut être affichée en détail, avec les dépenses, les événements et les commentaires qui la concernent. Toutes utilisent le même formulaire, présenté avec l'exemple des camions ; seuls leurs champs changent.

#### T3 — Fiche de référence (exemple : les camions)

*Accès : Menu › Fiches de référence › Camions. Profils : administrateur, analyste (consultation seule pour le profil consultation).*

Toutes les fiches de référence suivent ce même modèle : une recherche, une liste des éléments existants et une fiche détaillée. L'exemple montre la fiche d'un camion ; les champs propres aux autres fiches sont décrits en T4. Une fiche sélectionnée dans la liste s'affiche en dessous.

Éléments de l'écran :

- (1) **Zone de recherche** (zone de saisie et liste) : Texte recherché dans le champ choisi (code, immatriculation, type, site ou tous les champs). La liste se met à jour au clic sur « Rechercher ».
- (2) **Filtre de statut** (liste déroulante) : Affiche les éléments actifs, inactifs ou tous. Par défaut, seuls les éléments actifs sont listés.
- (3) **Liste des camions** (liste à colonnes) : Résultats de la recherche. Un clic sur un en-tête trie la colonne ; un clic sur une ligne affiche la fiche en dessous.
- (4) **Onglets** (onglets) : « Fiche » : les champs de l'élément. « Historique » : ses dépenses et ses événements, avec leur total. « Commentaires » : les remarques datées (écran T10).
- (5) **Code** (zone calculée) : Attribué automatiquement à la création selon le compteur de la table (VEH-01, VEH-02…). Jamais modifiable, il garantit l'absence de doublon.
- (6) **Identification** (cadre) : Immatriculation, type de véhicule et site d'affectation, tous obligatoires et choisis dans des listes que l'utilisateur peut compléter (écran T6).
- (7) **Capacités** (cadre) : Charge utile, volume et nombre de palettes ; nombres positifs. Elles servent au contrôle de chargement des missions et au taux de remplissage.
- (8) **Détention et coûts** (cadre) : Mode de détention, valeur d'achat, date de mise en service, durée d'amortissement et consommation de référence : données utilisées pour le coût complet du camion.
- (9) **Situation** (cadre) : Statut (actif, immobilisé, cédé) et chauffeur attitré. Un camion cédé n'apparaît plus dans les listes de saisie.
- (10) **Message** (message) : Confirme l'enregistrement ou indique en rouge le champ à corriger (champ vide, nombre négatif, date impossible, immatriculation déjà utilisée).
- (11) **Nouveau** (bouton) : Vide la fiche, attribue le prochain code et place le curseur sur le premier champ.
- (12) **Enregistrer** (bouton principal) : Contrôle tous les champs puis écrit la fiche dans la table ; l'action est inscrite au journal.
- (13) **Supprimer** (bouton) : Après confirmation, supprime un élément jamais utilisé. S'il possède déjà des dépenses ou des missions, la suppression est refusée et la désactivation est proposée.
- (14) **Désactiver** (bouton) : Retire l'élément des listes de saisie en conservant ses données et son historique.
- (15) **Annuler** (bouton) : Abandonne les modifications en cours et réaffiche la fiche telle qu'elle est enregistrée.
- (16) **Menu** (bouton) : Ferme le formulaire et revient au menu principal.


#### T4 — Champs des autres fiches de référence

Les fiches ci-dessous utilisent exactement le formulaire décrit en T3 (recherche, liste, onglets, boutons et règles communes). Seuls les champs de la fiche changent. Les champs marqués d'un astérisque sont obligatoires ; les listes de choix proviennent des listes de paramétrage (T6) et peuvent être complétées par l'utilisateur.

| Fiche | Champs de la fiche | Contrôles et règles particulières |
|---|---|---|
| Fournisseurs | Code (automatique), raison sociale*, pays*, ville, devise de facturation*, type de marchandises ou de services, conditions de paiement, délai d'approvisionnement, contact, statut. | La devise doit exister dans la table des devises ; un fournisseur étranger impose une devise et un pays. Un fournisseur ayant des dépenses ne peut être que désactivé. |
| Prestataires | Code, raison sociale*, nature de la prestation* (transporteur, transitaire, compagnie maritime, assureur, garage, gardiennage…), date de début et de fin de contrat, montant annuel du contrat, grille tarifaire associée, contact, statut. | La date de fin de contrat ne peut précéder la date de début. Un transporteur peut être rattaché à une grille tarifaire pour le contrôle de ses factures. |
| Entrepôts et zones | Entrepôt : code, nom*, ville*, surface (m²), mode de détention (propriété ou location), loyer mensuel. Zones de l'entrepôt : code, régime de température* (ambiant, froid positif, froid négatif), capacité en emplacements de palettes*. | La capacité d'une zone est un entier positif. La somme des capacités des zones donne la capacité de l'entrepôt, affichée sans saisie. |
| Personnel | Matricule, nom et prénom*, fonction*, site*, centre de coûts*, type de contrat, coût mensuel chargé*, date d'entrée*, date de sortie. | La date de sortie, si elle existe, doit suivre la date d'entrée. Ces dates permettent de calculer la masse salariale de n'importe quelle période. |
| Produits | Code, désignation*, famille*, régime de température*, poids par colis (kg)*, volume par colis (m³)*, colis par palette*, coût d'achat, prix de vente, fournisseur principal, statut. | Poids, volume et palettisation doivent être strictement positifs : un produit sans poids est une anomalie critique (contrôle T14). |
| Clients | Code, raison sociale*, segment* (grande surface, grossiste, supérette, café-hôtel-restaurant, restauration collective, export…), ville*, zone de livraison*, distance au site de livraison (km)*, site livreur, montant minimal de commande, conditions de paiement, statut. | La distance est obligatoire et positive : un client sans distance est une anomalie critique. Le site livreur est proposé d'après la zone de livraison. |
| Zones de livraison | Code, nom*, site de départ*, mode de transport* (flotte interne, sous-traité, export), distance de référence (km), jours de tournée. | Le mode de transport oriente le choix entre camion de l'entreprise et transporteur lors de la saisie des missions. |
| Grilles tarifaires | Transporteur*, type de prestation* (messagerie, affrètement, location), zone ou liaison*, tranche de poids, prix convenu*, minimum facturé, date de début et de fin de validité*. | Deux tarifs d'un même transporteur ne peuvent se chevaucher sur la même zone et la même tranche. Le tarif en vigueur à la date d'une facture sert à la contrôler. |
| Sites | Code, nom*, ville*, rôle (entrepôt central ou plateforme régionale), responsable. | Liste courte, utilisée par toutes les autres fiches ; un site utilisé ne peut pas être supprimé. |

Règles communes à toutes les fiches : code attribué automatiquement ; contrôle des champs avant enregistrement avec un message qui désigne le champ à corriger ; mise à jour immédiate des listes déroulantes ; recherche et tri ; désactivation au lieu de la suppression dès qu'un élément possède un historique ; consultation de l'historique et des commentaires dans les onglets de la fiche.


### C. Types de coûts et listes de paramétrage

La liste des types de coûts est la pièce maîtresse du paramétrage. Elle contient au départ toutes les catégories issues du plan comptable, comme le carburant, les péages, la maintenance préventive et curative, les pneumatiques, les assurances, les loyers, l'énergie, le gardiennage, le fret, les frais portuaires, la messagerie ou l'affrètement, et l'utilisateur peut y ajouter à tout moment de nouvelles catégories, aussitôt disponibles dans la saisie, les calculs et Power BI. Pour chaque type, il indique sa famille, l'étape de la chaîne à laquelle il appartient et son compte comptable, puis trois caractéristiques : comment le montant est obtenu (saisi depuis une facture, calculé à partir d'une quantité et d'un prix, ou calculé par l'application), à quoi la dépense se rattache (un objet précis, plusieurs objets selon une clé de partage, ou les frais de structure) et si elle est ponctuelle ou doit être étalée sur une période. Sur le même modèle, l'application gère les types de documents, les types d'événements, les types de prestations, de véhicules et de commandes, les fonctions du personnel et les devises.

#### T5 — Types de coûts

*Accès : Menu › Paramétrage › Types de coûts. Profils : administrateur (modification), analyste (consultation).*

Référentiel des catégories dans lesquelles chaque dépense est classée. La liste est ouverte : l'utilisateur crée autant de types que nécessaire. Pour chaque type, trois caractéristiques indiquent à l'application comment traiter les dépenses correspondantes, sans aucune modification du programme.

Éléments de l'écran :

- (1) **Filtre par famille** (liste déroulante) : Restreint la liste à une famille (flotte, stockage, approvisionnement, personnel…) ; la case « Inactifs » affiche aussi les types désactivés.
- (2) **Liste des types** (liste à colonnes) : Types existants, créés au départ à partir du plan comptable de la base. Un clic affiche le type à droite.
- (3) **Code** (zone calculée) : Attribué automatiquement (TC-01, TC-02…) ; il identifie le type dans les dépenses et les résultats.
- (4) **Libellé** (zone de saisie) : Nom affiché dans les listes et dans Power BI. Obligatoire et unique.
- (5) **Famille** (liste déroulante) : Regroupement d'analyse utilisé par la synthèse et les filtres. Obligatoire.
- (6) **Étape de la chaîne** (liste déroulante) : Approvisionnement, stockage, traitement des commandes, transport et livraison ou structure. Permet d'agréger les coûts par étape.
- (7) **Compte comptable** (liste déroulante) : Compte du grand livre correspondant ; il sert au rapprochement des coûts calculés avec la comptabilité.
- (8) **Nature** (boutons d'option) : Charge fixe ou variable ; utilisée dans les scénarios de simulation.
- (9) **Obtention du montant** (boutons d'option) : Saisi d'après une facture, obtenu par une quantité multipliée par un prix (le formulaire de dépense demande alors les deux), ou calculé par l'application (amortissement, masse salariale).
- (10) **Rattachement** (boutons d'option) : La dépense concerne un objet précis, est partagée entre plusieurs objets selon une clé, ou reste une charge de structure.
- (11) **Objet et clé** (listes déroulantes) : Type d'objet exigé à la saisie (camion, entrepôt, zone, commande…) si le rattachement est direct ; clé de partage (capacité, kilomètres, nombre de commandes…) s'il est partagé.
- (12) **Répartition dans le temps** (boutons d'option) : Ponctuelle : rattachée au jour de la pièce. Étalée : le formulaire de dépense demande les dates de couverture et le coût est réparti sur les jours.
- (13) **Type actif** (case à cocher) : Un type décoché n'est plus proposé à la saisie mais reste dans l'historique et les calculs passés.
- (14) **Utilisation** (zone calculée) : Nombre de dépenses enregistrées avec ce type ; un type utilisé peut seulement être désactivé.
- (15) **Information** (message) : Rappelle que tout nouveau type est disponible immédiatement partout, sans intervention sur le programme.
- (16) **Nouveau** (bouton) : Prépare un type vide avec le prochain code.
- (17) **Enregistrer** (bouton principal) : Contrôle la cohérence des caractéristiques (par exemple, une clé est exigée si le rattachement est partagé) puis enregistre.
- (18) **Désactiver** (bouton) : Désactive le type sélectionné après confirmation.


#### T6 — Listes de choix

*Accès : Menu › Paramétrage › Listes de choix. Profils : administrateur (modification), analyste (consultation).*

Un même écran gère toutes les listes de choix utilisées dans les formulaires, autres que les types de coûts. Aucune liste n'est figée dans le programme : un élément ajouté ici apparaît immédiatement dans les listes déroulantes correspondantes. L'exemple montre la liste des types d'événements.

Éléments de l'écran :

- (1) **Choix de la liste** (liste) : Types de documents (facture, avoir, ticket de carburant, contrat, devis…), d'événements, de prestations, de véhicules et de commandes, fonctions du personnel, segments de clientèle, familles de produits, devises. Un clic affiche le contenu à droite.
- (2) **Contenu de la liste** (liste à colonnes) : Éléments de la liste choisie. Les colonnes dépendent de la liste : gravité par défaut pour les événements, taux de change mensuel pour les devises, régime de température pour les familles.
- (3) **Libellé** (zone de saisie) : Nom affiché dans les formulaires. Obligatoire et unique dans la liste.
- (4) **Gravité par défaut** (liste déroulante) : Propre aux types d'événements : gravité proposée automatiquement à la saisie d'un événement, que l'utilisateur peut modifier.
- (5) **Immobilisation** (case à cocher) : Propre aux types d'événements : si elle est cochée, la saisie de l'événement demande les dates d'immobilisation du camion ou de la zone.
- (6) **Actif** (case à cocher) : Un élément inactif n'est plus proposé à la saisie mais reste lisible dans l'historique.
- (7) **Message** (message) : Signale qu'un élément déjà utilisé ne peut pas être supprimé et propose sa désactivation.
- (8) **Ajouter** (bouton) : Crée un nouvel élément dans la liste choisie avec le prochain code.
- (9) **Enregistrer** (bouton principal) : Contrôle l'unicité du libellé puis met à jour la liste et toutes les listes déroulantes qui l'utilisent.
- (10) **Désactiver** (bouton) : Désactive l'élément sélectionné après confirmation.


### D. Gestion des dépenses

C'est le cœur du volet transactionnel. Pour ajouter une dépense, l'utilisateur indique sa date, son type de coût, le fournisseur ou le prestataire, l'objet concerné (un camion, un entrepôt, une commande), le document justificatif et le montant ; le formulaire s'adapte au type choisi, par exemple en demandant les litres et le prix au litre pour le carburant, ou les dates de couverture pour une assurance. Chaque dépense porte un statut, saisie, validée ou annulée. Elle peut être modifiée, annulée ou supprimée après confirmation, seul l'administrateur pouvant la supprimer définitivement, et elle peut être retrouvée grâce à une recherche multicritère (période, type, fournisseur, objet, montant, numéro de document) qui affiche le total des montants trouvés. L'application refuse les factures en double, accepte les avoirs, permet de joindre le justificatif numérisé, de dupliquer une dépense et de générer automatiquement les dépenses récurrentes comme un loyer. Les dépenses de 2025 de la base sont reprises au démarrage, sans jamais compter deux fois une même charge.

#### T7 — Saisie d'une dépense

*Accès : Menu › Opérations courantes › Saisir une dépense, ou double-clic dans la recherche (T8). Profils : administrateur, analyste.*

Cœur du volet transactionnel : enregistrer, modifier, annuler ou supprimer toute dépense logistique. Le formulaire s'adapte au type de coût choisi : l'exemple d'une assurance annuelle fait apparaître la période couverte, tandis que la zone quantité × prix reste inactive.

Éléments de l'écran :

- (1) **Numéro et statut** (zone calculée) : Numéro attribué automatiquement à l'enregistrement ; le statut (saisie, validée, annulée) indique où en est la dépense. Non modifiables.
- (2) **Date de la pièce** (zone de saisie) : Date figurant sur le justificatif. Obligatoire ; doit appartenir à une période non clôturée.
- (3) **Type de coût** (liste déroulante) : Choisi dans le référentiel T5. Son choix adapte le formulaire : champs quantité et prix, période couverte, objet exigé.
- (4) **Famille et étape** (zone calculée) : Rappel de la famille et de l'étape de la chaîne du type choisi, pour vérifier le classement.
- (5) **Fournisseur ou prestataire** (liste déroulante) : Tiers qui a facturé la dépense ; recherche par les premières lettres. Obligatoire.
- (6) **Objet concerné** (listes déroulantes) : Nature de l'objet (camion, entrepôt, zone, commande…) puis l'objet ; obligatoire si le type exige un rattachement direct.
- (7) **Événement lié** (liste déroulante) : Facultatif : panne, sinistre ou litige qui a provoqué la dépense, pour en connaître le coût réel (T9).
- (8) **Document** (liste et zone de saisie) : Type de justificatif (facture, avoir, ticket, contrat…) et son numéro, obligatoire. Un avoir est saisi en montant négatif et relié à la facture d'origine.
- (9) **Fichier numérisé** (zone et bouton) : Chemin du justificatif scanné ; le bouton « Ouvrir » l'affiche.
- (10) **Période couverte** (zones de saisie) : Visible seulement si le type est étalé dans le temps. La date de fin doit suivre la date de début ; le coût sera réparti sur ces jours.
- (11) **Quantité et prix unitaire** (zones de saisie) : Actives seulement si le montant s'obtient par quantité × prix (litres de carburant, heures…) ; le montant est alors calculé.
- (12) **Montant, devise et taux** (zones de saisie) : Montant hors taxes, positif sauf pour un avoir ; pour une devise étrangère, le taux du mois est proposé automatiquement.
- (13) **Montants calculés** (zones calculées) : Montant TTC et montant hors taxes converti en dirhams, calculés sans saisie à partir du montant, de la TVA et du taux.
- (14) **Commentaire** (zone de texte) : Remarque libre, consultable dans la recherche et dans l'historique de l'objet.
- (15) **Contrôle de doublon** (message) : Avertit si une dépense porte déjà le même fournisseur, le même numéro de document et le même montant ; l'enregistrement exige alors une confirmation.
- (16) **Enregistrer** (bouton principal) : Contrôle les champs puis enregistre ; à la création, le numéro est attribué et l'action est inscrite au journal.
- (17) **Dupliquer** (bouton) : Crée une nouvelle dépense pré-remplie à partir de celle affichée, pour saisir plus vite une dépense semblable.
- (18) **Annuler la dépense** (bouton) : Passe la dépense au statut « annulée » : elle reste dans l'historique mais n'est plus comptée dans les coûts.
- (19) **Supprimer** (bouton) : Suppression définitive après confirmation, réservée à l'administrateur.
- (20) **Fermer** (bouton) : Ferme le formulaire ; une confirmation est demandée s'il reste des modifications non enregistrées.


#### T8 — Recherche des dépenses

*Accès : Menu › Opérations courantes › Rechercher les dépenses. Profils : tous (ouverture en lecture seule pour le profil consultation).*

Retrouver rapidement n'importe quelle dépense en combinant plusieurs critères, puis l'ouvrir pour la consulter, la modifier ou l'annuler. L'exemple montre les dépenses du premier trimestre 2025 pour le camion VEH-07.

Éléments de l'écran :

- (1) **Période** (zones de saisie) : Dates de début et de fin de la recherche, sur la date de la pièce. Par défaut : l'exercice en cours.
- (2) **Type et famille** (listes déroulantes) : Restreignent la recherche à un type de coût ou à une famille entière.
- (3) **Fournisseur ou prestataire** (liste déroulante) : Restreint la recherche aux dépenses d'un tiers.
- (4) **Objet** (listes déroulantes) : Nature puis identifiant de l'objet : un camion, un entrepôt, une zone, une commande…
- (5) **Fourchette de montants** (zones de saisie) : Montant hors taxes minimal et maximal ; l'une ou l'autre borne peut rester vide.
- (6) **Numéro de document** (zone de saisie) : Recherche sur tout ou partie du numéro du justificatif.
- (7) **Mot du commentaire** (zone de saisie) : Recherche un mot dans les commentaires des dépenses.
- (8) **Statut** (cases à cocher) : Inclut ou exclut les dépenses saisies, validées ou annulées.
- (9) **Rechercher** (bouton principal) : Applique tous les critères renseignés ; « Effacer les critères » remet le formulaire à zéro.
- (10) **Résultats** (liste à colonnes) : Dépenses trouvées, triables par un clic sur l'en-tête. Un double-clic ouvre la dépense dans le formulaire T7.
- (11) **Total** (zone calculée) : Nombre de dépenses trouvées et somme de leurs montants hors taxes, recalculés à chaque recherche.
- (12) **Montant réparti** (case à cocher) : Affiche, pour les dépenses étalées, la seule part qui concerne la période recherchée (par exemple 1 019 MAD en mars pour une assurance annuelle de 12 000 MAD).
- (13) **Ouvrir** (bouton principal) : Ouvre la dépense sélectionnée dans le formulaire de saisie.
- (14) **Nouvelle dépense** (bouton) : Ouvre un formulaire de saisie vide.
- (15) **Exporter la liste** (bouton) : Copie les résultats dans une nouvelle feuille Excel, sans macros, pour les transmettre.
- (16) **Dépenses récurrentes** (bouton) : Ouvre la liste des modèles de dépenses mensuelles (loyer, abonnement, contrat d'entretien…) et génère, pour le mois choisi, les dépenses correspondantes au statut « saisie », après confirmation.


### E. Événements, remarques et commentaires

L'application tient un registre des événements qui affectent l'activité, comme une panne, un accident, un sinistre, un retard ou un litige, avec leur date, leur gravité, l'objet concerné et la durée d'immobilisation éventuelle ; un événement est clôturé lorsqu'il est réglé. Il peut être relié aux dépenses qu'il a provoquées, ce qui donne le coût réel de chaque incident, et sa durée permet de mesurer la disponibilité de la flotte. Une remarque ou un commentaire, daté et signé du profil de son auteur, peut être ajouté sur n'importe quel enregistrement.

#### T9 — Registre des événements

*Accès : Menu › Opérations courantes › Événements, ou onglet « Historique » d'une fiche. Profils : administrateur, analyste.*

Enregistrer les faits qui affectent l'activité (panne, accident, sinistre, retard, litige, rupture de stock, contrôle…), les relier aux dépenses qu'ils ont provoquées et suivre leur résolution. Les immobilisations alimentent l'indicateur de disponibilité de la flotte dans l'application décisionnelle.

Éléments de l'écran :

- (1) **Filtres** (zones et listes) : Période, type, gravité et statut des événements affichés dans le registre.
- (2) **Registre** (liste à colonnes) : Événements correspondant aux filtres, avec leur gravité en couleur (rouge critique, orange majeure, bleu mineure). Un clic affiche l'événement en dessous.
- (3) **Date et heure** (zone de saisie) : Moment où l'événement s'est produit. Obligatoire ; le numéro est attribué automatiquement.
- (4) **Type** (liste déroulante) : Choisi dans la liste des types d'événements (T6). Il propose la gravité par défaut et, s'il y a lieu, les dates d'immobilisation.
- (5) **Gravité** (liste déroulante) : Mineure, majeure ou critique ; proposée par le type, modifiable par l'utilisateur.
- (6) **Objet concerné** (listes déroulantes) : Camion, zone de stockage, commande, prestataire ou fournisseur concerné. Obligatoire.
- (7) **Description** (zone de texte) : Récit des faits et des mesures prises. Obligatoire.
- (8) **Immobilisation** (zones de saisie) : Dates de début et de fin si l'objet a été arrêté ; la durée en jours est calculée et sert au taux de disponibilité.
- (9) **Statut** (zone calculée) : Ouvert tant que l'événement n'est pas réglé ; la date de clôture s'affiche lorsqu'il l'est.
- (10) **Dépenses liées** (liste à colonnes) : Dépenses rattachées à l'événement, par exemple les factures de réparation et la location d'un véhicule de remplacement.
- (11) **Lier une dépense** (bouton) : Ouvre la recherche des dépenses (T8) pour en rattacher une ; « Délier » retire le lien sans supprimer la dépense.
- (12) **Coût de l'événement** (zone calculée) : Somme des dépenses liées : coût réel de l'incident.
- (13) **Nouveau** (bouton) : Prépare un nouvel événement à la date du jour.
- (14) **Enregistrer** (bouton principal) : Contrôle les champs (la fin d'immobilisation doit suivre son début) puis enregistre.
- (15) **Clôturer l'événement** (bouton) : Marque l'événement comme réglé à la date du jour, après confirmation.
- (16) **Commentaires** (bouton) : Ouvre les remarques de l'événement (T10) ; le nombre entre parenthèses indique combien il en existe.


#### T10 — Remarques et commentaires

*Accès : bouton « Commentaires » d'une fiche, d'une dépense, d'un événement, d'une commande ou d'une mission. Profils : administrateur, analyste (lecture seule pour la consultation).*

Fenêtre commune qui permet d'ajouter une remarque datée et signée sur n'importe quel enregistrement de l'application, et de consulter toutes celles qui existent. Les commentaires sont conservés dans une table unique et apparaissent aussi dans l'onglet « Commentaires » des fiches.

Éléments de l'écran :

- (1) **Enregistrement commenté** (bandeau) : Rappelle la nature et l'identifiant de l'enregistrement auquel les commentaires sont rattachés.
- (2) **Commentaires existants** (liste à colonnes) : Remarques classées de la plus ancienne à la plus récente, avec leur date, leur heure et le profil de leur auteur.
- (3) **Nouveau commentaire** (zone de texte) : Texte libre de la remarque à ajouter ; il ne peut pas être vide.
- (4) **Ajouter** (bouton principal) : Enregistre la remarque avec la date, l'heure et l'identifiant de l'utilisateur connecté.
- (5) **Supprimer** (bouton) : Supprime le commentaire sélectionné après confirmation ; autorisé pour son auteur et pour l'administrateur seulement.
- (6) **Fermer** (bouton) : Ferme la fenêtre et met à jour le nombre de commentaires affiché sur le bouton d'origine.


### F. Commandes et missions

L'application permet de saisir une commande client avec ses lignes, en affichant automatiquement son poids, son volume, son nombre de palettes et son montant, puis de suivre son statut de la saisie à la livraison ; une commande peut être annulée ou livrée en plusieurs fois. Elle permet aussi de saisir une mission de transport, avec le véhicule ou le transporteur, le chauffeur, les kilomètres parcourus, les péages et les commandes chargées, en vérifiant que la capacité du véhicule n'est pas dépassée, ainsi que les commandes passées aux fournisseurs.

#### T11 — Saisie d'une commande client

*Accès : Menu › Opérations courantes › Commandes clients. Profils : administrateur, analyste.*

Saisir une commande avec son en-tête et ses lignes, puis suivre son statut jusqu'à la livraison. Les totaux logistiques se calculent pendant la saisie à partir des fiches produits. Le formulaire des commandes fournisseurs suit le même modèle (fournisseur, dates, lignes et coût de passation de la commande).

Éléments de l'écran :

- (1) **N° de commande** (zone calculée) : Numéro suivant du compteur des commandes (250001, 250002…), attribué à l'enregistrement.
- (2) **Client** (liste déroulante) : Recherche par code ou par nom parmi les clients actifs. Obligatoire.
- (3) **Rappel du client** (zone calculée) : Segment, zone de livraison et distance du client, repris de sa fiche ; ils serviront à la répartition du transport.
- (4) **Dates** (zones de saisie) : Date de commande et date de livraison souhaitée, obligatoires ; la seconde ne peut précéder la première.
- (5) **Site expéditeur** (liste déroulante) : Site qui prépare et expédie la commande ; proposé d'après la zone du client.
- (6) **Type de commande** (liste déroulante) : Normale, urgente, enlèvement par le client ou export. Une commande urgente est majorée dans le coût de traitement ; un enlèvement ne supporte aucun transport.
- (7) **Statut** (zone calculée) : Saisie, préparée, expédiée, livrée, livrée partiellement ou annulée ; il évolue avec les missions de transport.
- (8) **Produit** (liste déroulante) : Produit à ajouter, recherché par code ou par désignation parmi les produits actifs.
- (9) **Quantité** (zone de saisie) : Nombre de colis, entier strictement positif. « Ajouter la ligne » l'insère dans la liste ; « Retirer » supprime la ligne sélectionnée.
- (10) **Lignes de la commande** (liste à colonnes) : Une ligne par produit ; poids, volume, palettes et montant sont calculés à partir de la fiche produit.
- (11) **Totaux** (zones calculées) : Nombre de lignes, colis, poids, volume, palettes, poids taxable et montant, recalculés à chaque ligne ajoutée ou retirée. Le poids taxable tient compte de l'encombrement.
- (12) **Contrôle** (message) : Indique si la commande atteint le montant minimal convenu avec le client ; en dessous, un avertissement est affiché sans bloquer.
- (13) **Enregistrer** (bouton principal) : Exige au moins une ligne, enregistre l'en-tête et les lignes et attribue le numéro.
- (14) **Annuler la commande** (bouton) : Passe la commande au statut « annulée » après confirmation ; elle ne supportera alors que son coût administratif.
- (15) **Commandes du client** (bouton) : Affiche la liste des commandes du client sur une période choisie, avec leur statut et leur montant.


#### T12 — Saisie d'une mission de transport

*Accès : Menu › Opérations courantes › Missions de transport. Profils : administrateur, analyste.*

Enregistrer un transport réalisé par un camion de l'entreprise ou par un transporteur, avec ses kilomètres, ses frais et les commandes chargées. Ces informations permettront de calculer le coût de la mission et de le répartir entre les commandes transportées.

Éléments de l'écran :

- (1) **N° de mission** (zone calculée) : Numéro suivant du compteur des missions (600001, 600002…). La date et le site de départ sont obligatoires.
- (2) **Type de mission** (liste déroulante) : Tournée de livraison, navette entre sites, enlèvement chez un fournisseur, livraison express ou export.
- (3) **Moyen de transport** (boutons d'option) : Flotte de l'entreprise ou transporteur ; le choix active soit le véhicule et le chauffeur, soit le transporteur.
- (4) **Véhicule** (liste déroulante) : Camions actifs du site de départ ; un camion immobilisé ce jour-là par un événement n'est pas proposé.
- (5) **Chauffeur** (liste déroulante) : Salariés de fonction chauffeur ; le chauffeur attitré du véhicule est proposé par défaut.
- (6) **Transporteur** (liste déroulante) : Prestataire de transport, actif seulement pour une mission sous-traitée ; son coût viendra de ses factures.
- (7) **Compteurs** (zones de saisie) : Relevés kilométriques au départ et au retour ; le retour doit être supérieur au départ, et la distance est calculée.
- (8) **Péages** (zone de saisie) : Montant des péages de la mission, positif ou nul.
- (9) **Frais de mission** (zones et cases) : Nuitées hors base, présence d'un aide-livreur et recours à un chauffeur intérimaire, qui s'ajoutent au coût de la mission.
- (10) **Commandes non affectées** (liste à colonnes) : Commandes du site prêtes à livrer et non encore chargées sur une mission.
- (11) **Charger / retirer** (boutons) : Déplacent la commande sélectionnée d'une liste à l'autre ; les jauges se mettent à jour aussitôt.
- (12) **Commandes chargées** (liste à colonnes) : Commandes transportées par la mission ; une commande peut être livrée en plusieurs missions.
- (13) **Remplissage** (jauges) : Poids, volume et palettes chargés rapportés à la capacité du véhicule ; vert sous 75 %, orange au-delà, rouge si la capacité est dépassée.
- (14) **Contrôle de capacité** (message) : Signale un chargement qui dépasse la capacité ; l'enregistrement est alors refusé.
- (15) **Enregistrer** (bouton principal) : Contrôle les champs puis enregistre la mission et ses commandes, dont le statut passe à « expédiée ».
- (16) **Clôturer la mission** (bouton) : Exige le compteur de retour, puis passe les commandes chargées au statut « livrée ».


### G. Paramètres, contrôles et clôture

Un écran unique regroupe les paramètres utilisés par les calculs, comme les coûts unitaires, les taux et les seuils d'alerte, modifiables par l'administrateur seulement. À la demande, l'application contrôle l'ensemble des données et signale les incohérences, comme des codes inconnus, des dates impossibles, des valeurs manquantes ou des doublons, dont les six anomalies volontaires de la base. Chaque anomalie est classée comme critique, majeure ou simple information, et un double-clic ouvre l'enregistrement à corriger. Une période ne peut être clôturée que lorsque toutes les anomalies critiques sont corrigées, ce qui garantit que l'analyse porte toujours sur des données justes.

#### T13 — Paramètres de calcul

*Accès : Menu › Paramétrage › Paramètres. Profils : administrateur (modification) ; lecture seule pour les autres profils.*

Écran unique qui regroupe toutes les valeurs utilisées par les calculs et les alertes. Aucune de ces valeurs n'est écrite dans le programme : changer un tarif, un taux ou un seuil se fait ici, sans aucune modification du code, et s'applique au calcul suivant.

Éléments de l'écran :

- (1) **Onglets** (onglets) : « Coûts unitaires » : coûts de référence (traitement d'une commande et d'une ligne, emplacement de palette, kilomètre…). « Taux et coefficients » : coût du capital, majoration des urgences, coefficient du poids taxable. « Seuils d'alerte » : seuils des alertes. « Général » : exercice, règle de partage du transport par défaut, emplacement de l'application décisionnelle.
- (2) **Liste des paramètres** (liste à colonnes) : Code, libellé, valeur, unité et rôle de chaque paramètre de l'onglet. Les seuils couvrent les coûts, le budget, les factures, le stockage, le service, les retards, le risque de rupture, les pertes, la marge et le rapprochement comptable. Un clic affiche le paramètre en dessous.
- (3) **Valeur** (zone de saisie) : Nouvelle valeur du paramètre, dans l'unité indiquée. Seule zone modifiable de l'écran.
- (4) **Bornes admises** (zone calculée) : Intervalle de valeurs acceptées pour ce paramètre ; il sert au contrôle de saisie.
- (5) **Dernière modification** (zone calculée) : Date et auteur du dernier changement ; chaque modification est aussi inscrite au journal.
- (6) **Contrôle de saisie** (message) : Refuse une valeur non numérique ou hors bornes, en indiquant la plage attendue.
- (7) **Enregistrer** (bouton principal) : Enregistre la nouvelle valeur après confirmation ; elle sera utilisée par l'application décisionnelle au prochain calcul.
- (8) **Annuler** (bouton) : Rétablit la valeur enregistrée.


#### T14 — Contrôle des données et anomalies

*Accès : Menu › Contrôles et clôture › Contrôler les données. Profils : administrateur, analyste.*

Passer en revue toutes les données d'une période et lister les incohérences ligne par ligne, avant la clôture. Le contrôle détecte notamment les six anomalies volontaires de la base : produit sans poids, client sans distance, compteurs inversés, ticket de carburant en double, livraison antérieure à la commande et ligne de facture en double.

Éléments de l'écran :

- (1) **Période contrôlée** (liste déroulante) : Mois ou période à contrôler ; par défaut, la première période non clôturée.
- (2) **Périmètre** (liste déroulante) : Contrôle toutes les feuilles ou une seule (produits, clients, missions, dépenses…).
- (3) **Lancer le contrôle** (bouton principal) : Exécute tous les contrôles : codes inconnus, enregistrements rattachés à rien, dates impossibles, doublons, valeurs manquantes ou négatives, capacités dépassées.
- (4) **Dernier contrôle** (zone calculée) : Date, heure et durée du dernier contrôle, pour savoir si la liste est à jour.
- (5) **Compteurs** (indicateurs) : Nombre d'anomalies à traiter par gravité et nombre d'anomalies corrigées. Le nombre restant s'affiche aussi dans le bandeau du menu.
- (6) **Liste des anomalies** (liste à colonnes) : Contrôle concerné, feuille, enregistrement, description, gravité en couleur et état. Un double-clic ouvre directement l'enregistrement à corriger.
- (7) **Filtres d'affichage** (listes déroulantes) : Affichent les anomalies à traiter, corrigées ou toutes, et une gravité donnée.
- (8) **Ouvrir l'enregistrement** (bouton principal) : Ouvre le formulaire de l'enregistrement en cause (produit, client, mission, dépense…) pour le corriger.
- (9) **Marquer comme corrigée** (bouton) : Relance le contrôle sur la ligne choisie et la marque corrigée si l'anomalie a disparu ; une anomalie encore présente reste à traiter.
- (10) **Exporter** (bouton) : Copie la liste dans une nouvelle feuille pour la transmettre aux responsables concernés.
- (11) **Clôturer la période** (bouton) : Ouvre l'écran de clôture (T15) ; actif pour l'administrateur seulement.


#### T15 — Clôture de période

*Accès : Menu › Contrôles et clôture › Clôturer une période, ou bouton de l'écran T14. Profils : administrateur.*

Valider définitivement les données d'une période pour les rendre disponibles à l'application décisionnelle. La clôture est refusée tant qu'une anomalie critique subsiste : l'analyse ne porte ainsi jamais sur des données erronées. Une période clôturée ne peut plus recevoir de saisie.

Éléments de l'écran :

- (1) **Période à clôturer** (liste déroulante) : Première période non clôturée ; les périodes doivent être clôturées dans l'ordre chronologique.
- (2) **Conditions de clôture** (liste de vérifications) : État de chaque condition : contrôle récent, anomalies critiques (bloquantes), dépenses non validées, missions sans compteur de retour, événements ouverts (non bloquants). Vert : satisfaite ; rouge : bloquante ; orange ou bleu : à vérifier.
- (3) **Résumé de la période** (indicateurs) : Nombre de dépenses et montant total, commandes et missions de la période, pour vérifier la complétude avant de clôturer.
- (4) **Motif du refus** (message) : Explique pourquoi la clôture n'est pas possible et renvoie vers la liste des anomalies.
- (5) **Sauvegarde automatique** (message) : Rappelle qu'une copie datée du classeur est enregistrée juste avant la clôture.
- (6) **Clôturer la période** (bouton principal) : Grisé tant qu'une condition bloquante n'est pas satisfaite. Sinon, demande confirmation, crée la sauvegarde, verrouille la période et l'inscrit au journal.
- (7) **Voir les anomalies** (bouton) : Ouvre l'écran de contrôle (T14) filtré sur la période.
- (8) **Périodes clôturées** (liste à colonnes) : Historique des clôtures : période, date, auteur et nom de la sauvegarde. L'application décisionnelle ne peut importer que ces périodes.


### H. Administration et sécurité

L'accès se fait par profil, les actions importantes sont inscrites dans un journal, les feuilles sont protégées et une sauvegarde datée est créée avant chaque clôture. L'administrateur gère les comptes des utilisateurs, peut vider les opérations en conservant les fiches de référence et créer des données fictives pour les tests. En cas d'erreur, l'application affiche un message compréhensible plutôt que la fenêtre technique d'Excel ; tous les messages suivent les mêmes règles de présentation.

#### T16 — Administration : utilisateurs et journal

*Accès : Menu › Administration › Utilisateurs et journal. Profils : administrateur.*

Gérer les comptes d'accès et leurs profils, puis consulter la trace de toutes les actions importantes : connexions, créations, modifications, annulations, suppressions, clôtures et changements de paramètres. Le journal ne peut être ni modifié ni effacé depuis l'application, et les feuilles de données restent protégées par mot de passe.

Éléments de l'écran :

- (1) **Onglets** (onglets) : « Utilisateurs » (affiché), « Journal des actions » (aperçu en dessous) et « Sauvegardes et réinitialisation » (écran T17).
- (2) **Liste des utilisateurs** (liste à colonnes) : Comptes existants avec leur profil, leur état et leur dernière connexion. Un clic affiche le compte à droite.
- (3) **Profil** (liste déroulante) : Administrateur, analyste ou consultation ; il fixe les boutons actifs dans les deux applications.
- (4) **Compte actif** (case à cocher) : Un compte désactivé ne peut plus se connecter mais reste visible dans le journal.
- (5) **Réinitialiser le mot de passe** (bouton) : Attribue un mot de passe provisoire que l'utilisateur devra changer à sa prochaine connexion.
- (6) **Nouveau** (bouton) : Prépare un nouveau compte ; l'identifiant doit être unique.
- (7) **Enregistrer** (bouton principal) : Enregistre le compte et inscrit l'opération au journal.
- (8) **Journal des actions** (liste à colonnes) : Date et heure, utilisateur, action et objet de chaque opération importante, du plus récent au plus ancien.
- (9) **Filtres du journal** (zones et listes) : Restreignent le journal à une période, un utilisateur ou un type d'action.
- (10) **Exporter le journal** (bouton) : Copie les lignes filtrées dans une nouvelle feuille, par exemple pour un contrôle.


#### T17 — Administration : sauvegardes, réinitialisation et tests

*Accès : Menu › Administration › onglet « Sauvegardes et réinitialisation ». Profils : administrateur.*

Protéger le travail par des copies datées du classeur, remettre l'application à zéro pour un nouvel exercice ou une démonstration, et créer rapidement des données fictives pour tester les contrôles et les calculs.

Éléments de l'écran :

- (1) **Dossier des sauvegardes** (zone calculée) : Emplacement où sont enregistrées les copies ; il est défini dans les paramètres généraux.
- (2) **Liste des sauvegardes** (liste à colonnes) : Copies existantes avec leur date, leur motif (clôture, réinitialisation, sauvegarde manuelle) et leur taille.
- (3) **Sauvegarder maintenant** (bouton principal) : Enregistre une copie datée du classeur dans le dossier des sauvegardes.
- (4) **Ouvrir le dossier** (bouton) : Ouvre le dossier des sauvegardes dans l'explorateur de fichiers, pour restaurer une copie si nécessaire.
- (5) **Options de réinitialisation** (cases à cocher) : Choisissent ce qui est vidé (les opérations) et ce qui est conservé (fiches, listes, paramètres).
- (6) **Avertissement** (message) : Rappelle que l'opération est irréversible et qu'une sauvegarde est faite au préalable.
- (7) **Réinitialiser** (bouton) : Demande une double confirmation, crée la sauvegarde, vide les opérations choisies et inscrit l'action au journal.
- (8) **Volume de test** (zone de saisie) : Nombre de dépenses fictives à créer et période de leur date ; entier positif.
- (9) **Anomalies volontaires** (case à cocher) : Ajoute quelques erreurs connues (doublon, objet manquant, date impossible) pour vérifier qu'elles sont détectées par le contrôle T14.
- (10) **Générer** (bouton) : Crée les données fictives à partir des fiches existantes et affiche le nombre d'enregistrements créés.


#### T18 — Messages, confirmations et erreurs

*Accès : affichés par tous les écrans des deux applications. Profils : tous.*

Les deux applications communiquent avec l'utilisateur par un petit nombre de messages types, toujours rédigés en français, qui expliquent la situation et la suite possible. Aucune fenêtre technique d'Excel n'est jamais présentée à l'utilisateur.

Éléments de l'écran :

- (1) **Confirmation** (message à deux choix) : Demandée avant toute suppression, annulation, clôture ou réinitialisation. Le bouton par défaut est « Non », pour éviter une validation par erreur.
- (2) **Avertissement** (message à deux choix) : Signale une situation inhabituelle mais possible, comme un doublon de facture ; l'utilisateur peut poursuivre en connaissance de cause.
- (3) **Saisie incomplète** (message d'information) : Indique précisément le champ manquant ou invalide et la règle à respecter ; le champ concerné est mis en évidence dans le formulaire.
- (4) **Suppression impossible** (message à deux choix) : Explique pourquoi un élément utilisé ne peut pas être supprimé et propose directement sa désactivation.
- (5) **Accès refusé** (message d'erreur) : Affiché si une fonction n'est pas autorisée pour le profil connecté ; en pratique, les boutons correspondants sont déjà grisés.
- (6) **Erreur maîtrisée** (message d'erreur) : Remplace toute erreur imprévue : l'action est abandonnée sans rien enregistrer, l'affichage est rétabli et le détail est inscrit au journal avec une référence.

### Volet décisionnel : l'application de calcul et d'analyse

Le volet décisionnel est la partie de la solution qui transforme les données enregistrées en informations utiles à la décision. Il prend la forme d'une application de calcul et d'analyse, destinée au contrôleur de gestion et à la direction logistique. Elle ne saisit aucune opération : elle lit les périodes clôturées de l'application de gestion, calcule automatiquement les coûts sur n'importe quelle période, les répartit jusqu'au client et au produit, les explique par des indicateurs et des statistiques, déclenche les alertes, simule des décisions et produit les rapports ainsi que le fichier de résultats lu par Power BI. Ses fonctions sont regroupées en sept domaines, présentés ci-dessous avec les écrans qui les réalisent (D1 à D19). Elles suivent une même logique d'analyse, qui va du constat à la décision.

| Étape de l'analyse | Écrans | Résultat |
| --- | --- | --- |
| 1. Poser la question | D4 | Période, objet et catégories choisis. |
| 2. Repérer les contributeurs | D4, D6, D9, D11 | Répartition, concentration et dispersion du coût. |
| 3. Comparer à périmètre cohérent | D4, D5, D11 | Écarts, effet volume et effet coût unitaire. |
| 4. Comprendre l'origine | D7, D8, D10, D12 | Opérations, composantes et méthode de calcul. |
| 5. Traiter les anomalies | D17, D18 | Action, responsable et justification. |
| 6. Simuler des options | D13, D14, D15 | Scénarios chiffrés sur une copie des données. |
| 7. Décider et documenter | D16, D19, P1 à P4 | Décision justifiée, rapport et Power BI. |

### A. Accueil et récupération des données

L'écran d'accueil rappelle l'état des données importées, affiche les chiffres clés de la période et les alertes à traiter en premier, et donne accès à toutes les fonctions. Un bouton d'actualisation importe en lecture seule les données des périodes clôturées de l'application de gestion, qui ne peut donc jamais être modifiée par l'analyse. Une période non clôturée est refusée, afin que l'analyse porte toujours sur des données vérifiées, et la date de la dernière importation reste affichée.

#### D1 — Accueil de l'application décisionnelle

*Accès : ouverture du classeur ou bouton du menu de l'application de gestion. Profils : administrateur, analyste, consultation (lecture seule).*

Feuille d'accueil du volet décisionnel. Elle indique sur quelles données l'utilisateur travaille, résume les chiffres clés de la période importée, montre les alertes à traiter en premier et donne accès à toutes les analyses. La connexion utilise les mêmes comptes et profils que l'application de gestion.

Éléments de l'écran :

- (1) **Bandeau d'état** (bandeau) : Périodes importées, date de la dernière importation, date et durée du dernier calcul, nombre d'alertes à examiner. Il rappelle en permanence la fraîcheur des données analysées.
- (2) **Chiffres clés** (indicateurs) : Coût global, coût rapporté au chiffre d'affaires, coût par commande livrée et écart au budget sur toute la période importée ; un clic sur un chiffre ouvre l'analyse détaillée correspondante.
- (3) **Données** (groupe de boutons) : Actualiser les données ouvre l'import des périodes clôturées (D2) ; Calculer les coûts lance le calcul complet et ses contrôles (D3).
- (4) **Analyse** (groupe de boutons) : Synthèse de période (D4), indicateurs (D5), analyse par catégorie (D6), transport et flotte (D7), stockage (D8), clients et produits (D9), coût d'une commande (D10), statistiques (D11), méthodes de calcul et définitions (D12).
- (5) **Décision** (groupe de boutons) : Scénarios (D13), simulation d'un nouveau produit (D14), comparaison des fournisseurs (D15), comparaison des scénarios et décision (D16).
- (6) **Suivi et restitution** (groupe de boutons) : Alertes (D17) et leur suivi (D18), rapports et exports (D19), ouverture du rapport Power BI.
- (7) **Retour à l'application de gestion** (bouton) : Active le classeur transactionnel s'il est ouvert, ou l'ouvre.
- (8) **Aide** (bouton) : Ouvre la feuille d'aide : parcours type, définition des indicateurs et des alertes.
- (9) **Quitter** (bouton) : Enregistre et ferme le classeur décisionnel.
- (10) **À traiter en premier** (tableau) : Les alertes de priorité la plus élevée encore à examiner ou en cours, avec leur objet ; un double-clic ouvre directement leur suivi (D18).


#### D2 — Actualisation des données

*Accès : Accueil › Données › Actualiser les données. Profils : administrateur, analyste.*

Récupérer dans l'application décisionnelle les données des périodes clôturées de l'application de gestion. Le classeur de gestion est ouvert en lecture seule, copié puis refermé : l'analyse ne peut jamais modifier les données saisies, et elle ne porte que sur des périodes vérifiées.

Éléments de l'écran :

- (1) **Emplacement de l'application de gestion** (zone et bouton) : Chemin du classeur transactionnel, lu dans les paramètres ; « Parcourir » permet de le corriger si le fichier a été déplacé.
- (2) **Dernière clôture** (zone calculée) : Dernière période clôturée dans l'application de gestion, avec sa date de clôture.
- (3) **Périodes disponibles** (liste à colonnes) : Périodes de l'exercice avec leur état. Seules les périodes clôturées peuvent être cochées ; la colonne « Importée » indique celles déjà présentes.
- (4) **Importer** (bouton principal) : Copie, pour les périodes cochées, les fiches de référence, les types de coûts, les dépenses, les commandes, les missions, les événements et les paramètres.
- (5) **Déroulement de l'import** (liste à colonnes) : Étapes de l'import avec le nombre de lignes copiées et leur état ; une étape en échec est signalée en rouge avec sa cause.
- (6) **Période refusée** (message) : Explique qu'une période non clôturée ne peut pas être importée, pour que l'analyse porte toujours sur des données contrôlées.
- (7) **Compte rendu** (message) : Durée et heure de l'import ; ces informations s'affichent aussi dans le bandeau de l'accueil.
- (8) **Calculer maintenant** (bouton) : Enchaîne directement sur le calcul des coûts (D3).


### B. Calcul automatique des coûts et contrôles de justesse

L'application calcule le coût de chaque catégorie et le coût global sur la période choisie, comme l'a demandé l'encadrement. Elle convertit les montants en dirhams, étale les dépenses qui couvrent une période sur les jours concernés (une assurance annuelle de 12 000 dirhams ne compte ainsi que pour environ 1 019 dirhams en mars), calcule les charges sans facture comme l'amortissement des camions, le coût de détention du stock ou la masse salariale, et partage les charges communes selon la clé associée à leur type. Elle recalcule aussi, à partir des charges réelles, les coûts unitaires de référence et les propose pour mettre à jour les paramètres. Deux contrôles garantissent la justesse du résultat : la somme des coûts répartis, de la sous-activité et des frais de structure doit être égale au total des dépenses, et chaque poste doit correspondre au grand livre à moins de 0,5 % près.

#### D3 — Calcul des coûts et contrôles de justesse

*Accès : Accueil › Données › Calculer les coûts, ou après une actualisation. Profils : administrateur, analyste.*

Lancer en une seule action le calcul automatique de tous les coûts, puis vérifier qu'il est juste : aucune charge perdue ni comptée deux fois, et des coûts conformes au grand livre. Les résultats de toutes les analyses proviennent de ce calcul.

Éléments de l'écran :

- (1) **Période** (liste déroulante) : Toute la période importée ou un mois donné ; un calcul partiel est plus rapide mais les cumuls annuels restent ceux du dernier calcul complet.
- (2) **Partage du transport** (boutons d'option) : Règle de répartition du coût d'une mission entre ses commandes : poids, poids taxable ou poids multiplié par la distance. La valeur par défaut vient des paramètres.
- (3) **Lancer le calcul complet** (bouton principal) : Exécute les étapes dans l'ordre, après une sauvegarde automatique des résultats précédents.
- (4) **Étapes du calcul** (liste à colonnes) : Chaque étape avec sa durée et son état ; une étape en échec arrête le calcul et affiche sa cause.
- (5) **Progression** (jauge) : Avancement du calcul en cours.
- (6) **Contrôle de conservation** (indicateur) : Compare le total des dépenses de la période à la somme des coûts répartis, de la sous-activité et des frais de structure. L'écart doit être nul à l'arrondi près.
- (7) **Rapprochement avec le grand livre** (liste à colonnes) : Compare, poste par poste, les coûts calculés aux charges comptables. Un écart supérieur à 0,5 % crée une alerte critique qui bloque l'export vers Power BI.
- (8) **Coûts unitaires recalculés** (message) : Coûts de référence obtenus à partir des charges réelles (traitement d'une commande, préparation d'une ligne, emplacement, kilomètre…).
- (9) **Proposer la mise à jour** (bouton) : Affiche les nouveaux coûts unitaires à côté des valeurs actuelles des paramètres ; l'administrateur décide de les reporter dans l'application de gestion.
- (10) **Durée totale** (zone calculée) : Temps total du calcul ; l'objectif est de rester sous une minute pour l'année complète.


### C. Synthèse de période et indicateurs de performance

L'écran de synthèse répond à la demande d'analyser les coûts « de telle période à telle période ». L'utilisateur choisit une date de début, une date de fin, des catégories et éventuellement un objet, et l'application affiche le coût global, sa répartition par catégorie, son évolution mensuelle, les dix objets les plus coûteux et sa comparaison avec le budget et avec une période de référence. Il peut ainsi obtenir le coût de n'importe quelle combinaison de catégories, pour n'importe quel objet et sur n'importe quelle période, par exemple le coût « carburant, maintenance et assurance » d'un camion de mars à juin. L'écart avec la période de référence est séparé en deux parts : celle qui vient du volume d'activité et celle qui vient du coût unitaire.

Les indicateurs de performance couvrent le coût global (part de chaque catégorie, coût rapporté au chiffre d'affaires, coût par commande et par tonne livrée), le transport et la flotte (coût par kilomètre, taux de remplissage, consommation réelle et attendue, disponibilité, part de la maintenance curative, part de la sous-traitance), l'approvisionnement (frais d'approche, retard des fournisseurs), le stockage (coût par palette, occupation, rotation, couverture, pertes), les commandes et le service (délai de traitement, livraisons à l'heure, taux de service, part des commandes urgentes), les prestataires (écart entre factures et tarifs), les clients et les produits (marge après logistique) et les événements (nombre et coût des incidents). Les ratios sont calculés sur les totaux de la sélection, et non comme des moyennes de ratios.

#### D4 — Synthèse de période

*Accès : Accueil › Analyse › Synthèse de période. Profils : tous.*

Écran central de la demande « de telle période à telle période » : l'utilisateur choisit librement des dates, des catégories de coûts et un objet, et obtient le coût global, sa répartition, son évolution, les objets les plus coûteux et la comparaison avec une référence et avec le budget. Exemple : carburant, maintenance et assurance de la flotte de Casablanca en mars 2025.

Éléments de l'écran :

- (1) **Période analysée** (zones de saisie) : Dates de début et de fin, libres : un mois, un trimestre glissant, la période du Ramadan… La fin doit suivre le début et rester dans les périodes importées.
- (2) **Catégories** (liste à cases) : Types de coûts inclus dans l'analyse, regroupés par famille ; tous par défaut. Tout nouveau type créé dans l'application de gestion y apparaît après l'import.
- (3) **Objet analysé** (listes déroulantes) : Nature (entreprise entière, site, camion, entrepôt, prestataire, client, produit, segment) puis valeur. Les coûts affichés se limitent à cet objet.
- (4) **Comparer à** (liste déroulante) : Période de référence : période précédente de même durée, même période de l'année précédente, ou période choisie. Si les deux périodes n'ont pas la même durée, un avertissement s'affiche et l'écart est aussi donné en coût par jour.
- (5) **Actualiser** (bouton principal) : Recalcule l'ensemble de la synthèse pour la sélection ; les graphiques et tableaux se mettent à jour automatiquement.
- (6) **Indicateurs clés** (indicateurs) : Coût global de la sélection, écart à la période de référence, indicateur principal de l'objet (ici le coût par kilomètre) et écart au budget.
- (7) **Répartition par catégorie** (graphique) : Coût de chaque catégorie sélectionnée, du plus élevé au plus faible.
- (8) **Évolution mensuelle** (graphique) : Coût de chaque mois de la période, précédé des deux mois antérieurs pour donner la tendance, et budget mensuel en pointillés, pour situer les dérives.
- (9) **Principaux contributeurs** (tableau) : Les dix objets du niveau inférieur les plus coûteux (ici les camions du site) avec leur coût, leur part du total et leur indicateur principal ; les autres sont regroupés dans une ligne « Autres » pour que le total reste exact. Un double-clic ouvre le détail de l'objet.
- (10) **Comparaison au budget** (tableau) : Réel, budget et écart en pourcentage par catégorie ou par poste ; un écart supérieur à 10 % est coloré.
- (11) **Exporter en PDF** (bouton) : Produit le rapport de l'analyse affichée, avec sa sélection, ses méthodes et ses limites (voir D19).
- (12) **Indicateurs détaillés** (bouton) : Ouvre le tableau complet des indicateurs de performance (D5) pour la même sélection.
- (13) **Décomposition de l'écart** (zone calculée) : Sépare l'écart avec la période de référence en deux parts : celle qui vient de la variation de l'activité (effet volume, ici les kilomètres) et celle qui vient de la variation du coût unitaire (effet coût unitaire). Elle indique si le coût monte parce qu'on travaille plus ou parce que chaque unité coûte plus cher.


#### D5 — Indicateurs de performance

*Accès : Accueil › Analyse › Indicateurs, ou bouton de la synthèse (D4). Profils : tous.*

Tableau complet des indicateurs de performance de la sélection, regroupés par famille, avec la valeur de la période, celle de la période de référence, l'écart et l'objectif. Tous les ratios sont calculés à partir des totaux de la sélection, et non comme des moyennes de ratios.

Éléments de l'écran :

- (1) **Rappel de la sélection** (bandeau) : Période, catégories, objet et période de comparaison utilisés ; ce sont ceux de la synthèse (D4).
- (2) **Modifier la sélection** (bouton) : Revient à la synthèse pour changer les dates, les catégories ou l'objet.
- (3) **Tableau des indicateurs** (tableau) : Une ligne par indicateur, regroupées par famille : coût global (coût total, coût / chiffre d'affaires, coût par commande et par tonne) ; transport et flotte (coût par km et par tonne-kilomètre, remplissage, consommation, disponibilité, maintenance curative, sous-traitance) ; approvisionnement (frais d'approche par colis, retard des fournisseurs) ; stockage (coût par palette, occupation, rotation, couverture, sous-activité, pertes et écarts d'inventaire) ; commandes et service (délai de traitement, livraisons à la date prévue, taux de service, urgences, petites commandes) ; prestataires, clients et produits (écart aux tarifs, concentration, coût / CA, marge après logistique) ; événements (nombre, coût, immobilisations). Colonnes : valeur, référence, écart, objectif, état et tendance.
- (4) **Filtre de famille** (liste déroulante) : Affiche une seule famille d'indicateurs.
- (5) **Code couleur** (légende) : État de chaque indicateur par rapport à son objectif ou à son seuil d'alerte, lu dans les paramètres.
- (6) **Définition** (bouton) : Ouvre la fiche de l'indicateur sélectionné dans l'écran des méthodes de calcul (D12) : définition, unité, composantes incluses et données utilisées.
- (7) **Exporter** (bouton) : Copie le tableau dans une nouvelle feuille ou l'ajoute au rapport mensuel.


### D. Analyse détaillée des coûts

Les coûts peuvent être analysés par type, par famille, par étape de la chaîne, par compte, par site, par fournisseur ou par mois. Le coût d'une mission réalisée par la flotte réunit une part fixe propre au véhicule (amortissement, assurance, salaire du chauffeur), une part liée aux kilomètres parcourus (carburant et entretien) et les frais de la mission comme les péages. Celui d'une mission sous-traitée est celui de ses factures, comparées aux tarifs convenus : toute facture qui dépasse le tarif de plus de 3 % est signalée. Ce coût est ensuite partagé entre les commandes transportées selon une règle choisie par l'utilisateur : le poids, l'encombrement, ou le poids et la distance. Le coût de stockage d'un produit tient compte des emplacements qu'il occupe et de la valeur immobilisée en stock ; le coût mensuel d'un emplacement varie ainsi de 159,77 dirhams en zone ambiante à Casablanca à 604,73 dirhams en zone surgelée à Tanger, et les emplacements vides sont présentés à part comme coût de sous-activité. Le coût de traitement d'une commande comprend un coût fixe par commande et un coût par ligne, de 126,60 et 46,38 dirhams en référence 2025, majoré de 30 % pour une commande urgente.

Tous ces coûts sont répartis jusqu'à la ligne de commande, seul niveau où l'on connaît à la fois le client et le produit. On obtient ainsi le coût complet de chaque camion, commande, client, produit, segment de clientèle et site, ainsi que la marge qui reste après la logistique. Les clients et les produits sont classés selon leur poids dans le coût total, ce qui fait ressortir le petit nombre de clients qui concentre la plus grande partie de la dépense, et le coût d'une commande peut être détaillé jusqu'à ses jalons, ses missions et ses lignes de produits.

#### D6 — Analyse des coûts par catégorie

*Accès : Accueil › Analyse › Analyse par catégorie. Profils : tous.*

Tableau croisé qui permet de regrouper les coûts selon n'importe quel axe : type de coût, famille, étape de la chaîne, compte comptable, centre de coûts, site, fournisseur ou mois. Il répond à la demande de calculer les coûts « à partir des catégories » et permet de remonter jusqu'aux dépenses qui composent chaque montant.

Éléments de l'écran :

- (1) **Regroupement en lignes** (liste déroulante) : Axe des lignes : type de coût, famille › type, étape de la chaîne, compte comptable, centre de coûts, site, fournisseur ou prestataire, objet.
- (2) **Regroupement en colonnes** (liste déroulante) : Axe des colonnes : mois, trimestre, site, étape, ou aucun pour un simple total.
- (3) **Mesure** (liste déroulante) : Montant hors taxes, part du total, écart au budget, ou montant par unité d'activité (par commande, par tonne, par kilomètre).
- (4) **Période** (zones de saisie) : Dates de début et de fin de l'analyse ; les dépenses étalées ne comptent que pour leur part de la période.
- (5) **Tableau croisé** (tableau) : Montants par ligne et par colonne, avec les sous-totaux par famille et le total général. Les familles se développent ou se réduisent.
- (6) **Graphique** (graphique) : Répartition du total de la période selon un second axe, ici l'étape de la chaîne logistique.
- (7) **Voir les dépenses** (bouton) : Affiche la liste des dépenses qui composent la cellule sélectionnée (par exemple la maintenance curative de mars), comme dans la recherche T8.
- (8) **Développer ou réduire** (bouton) : Affiche ou masque le détail des types sous chaque famille.
- (9) **Exporter** (bouton) : Copie le tableau affiché dans une feuille sans macros.


#### D7 — Transport et flotte

*Accès : Accueil › Analyse › Transport et flotte. Profils : tous.*

Analyser le coût du transport : coût complet et indicateurs de chaque camion, coût de chaque mission et de sa répartition sur les commandes, et contrôle des factures des transporteurs par rapport aux tarifs convenus.

Éléments de l'écran :

- (1) **Onglets** (onglets) : « Camions » (affiché) : coût et indicateurs par véhicule. « Missions » : coût de chaque mission, distance, remplissage et commandes transportées. « Modes et délais » : comparaison des modes de transport (aperçu en 8). « Factures » : chaque ligne facturée comparée au tarif en vigueur.
- (2) **Sélection** (zones et liste) : Période et site analysés.
- (3) **Tableau des camions** (tableau) : Kilomètres, coût complet, coût par kilomètre, remplissage moyen, consommation par rapport à la référence, disponibilité et part de la maintenance curative. Les valeurs au-delà d'un seuil sont en rouge.
- (4) **Détail du coût complet** (graphique) : Composition du coût du camion sélectionné : amortissement, carburant, maintenance, assurance (pour sa part de la période), chauffeur, péages. Elle permet de comparer les véhicules entre eux.
- (5) **Factures au-dessus du tarif** (tableau) : Lignes de factures dont le montant dépasse de plus de 3 % le prix issu de la grille tarifaire ; chacune crée une alerte.
- (6) **Règle de partage** (liste déroulante) : Règle utilisée pour répartir le coût des missions sur les commandes ; la changer ici relance le calcul de répartition pour la période.
- (7) **Fiche du camion** (bouton) : Ouvre l'historique complet du camion : missions, dépenses, événements et immobilisations.
- (8) **Modes et délais** (tableau) : Pour chaque mode (flotte interne, messagerie, affrètement, export) : tonnage, coût par tonne et par colis, délai moyen entre expédition et livraison et part des livraisons à la date prévue. Il permet de comparer coût et délai avant de choisir un mode.


#### D8 — Stockage

*Accès : Accueil › Analyse › Stockage. Profils : tous.*

Analyser le coût du stockage par site, par zone et par produit. Le coût d'un produit réunit l'occupation de ses emplacements et la détention de sa valeur ; les emplacements payés mais vides sont présentés à part comme coût de sous-activité, pour ne pas alourdir injustement le coût des produits.

Éléments de l'écran :

- (1) **Sélection** (listes déroulantes) : Période, site et régime de température analysés.
- (2) **Zones de stockage** (tableau) : Capacité et occupation moyenne en emplacements, taux d'occupation moyen (le taux de pointe, plus haut relevé de la période, s'affiche au survol), coût mensuel d'un emplacement (de 159,77 MAD en ambiant à Casablanca à 604,73 MAD en surgelé à Tanger), coût d'occupation et coût de sous-activité de chaque zone.
- (3) **Produits** (tableau) : Pour chaque produit : palettes occupées en moyenne, coût d'occupation, coût de détention, coût total, rotation (sorties du mois / stock moyen) et couverture (stock moyen / sorties moyennes par jour). Les valeurs anormales sont en rouge.
- (4) **Chiffres clés** (indicateurs) : Coût total du stockage, coût de sous-activité, taux d'occupation moyen, nombre de produits dormants et valeur des pertes constatées (casse, produits périmés, écarts d'inventaire et retours non revendables), calculée à partir des inventaires et des retours.
- (5) **Filtres produits** (cases à cocher) : Limitent la liste aux produits dormants (sans sortie depuis 60 jours) ou couverts pour plus de 90 jours ; les seuils viennent des paramètres.
- (6) **Alerte** (message) : Signale une zone dont l'occupation moyenne ou de pointe dépasse le seuil ; le cas figure aussi dans la liste des alertes (D17).
- (7) **Exporter** (bouton) : Copie les tableaux affichés dans une feuille sans macros.


#### D9 — Clients et produits

*Accès : Accueil › Analyse › Clients et produits. Profils : tous.*

Coût logistique complet par client, par produit, par segment et par commande, obtenu en répartissant tous les coûts jusqu'à la ligne de commande. Il montre quels clients et quels produits concentrent la dépense et lesquels coûtent cher au regard de leur chiffre d'affaires.

Éléments de l'écran :

- (1) **Onglets** (onglets) : « Clients » (affiché), « Produits », « Segments » de clientèle et « Commandes » : même tableau au niveau choisi ; l'onglet Commandes détaille le coût de chaque commande.
- (2) **Sélection** (listes déroulantes) : Période et segment de clientèle analysés.
- (3) **Classement** (tableau) : Clients triés du plus coûteux au moins coûteux : segment (GMS : grandes et moyennes surfaces, CHR : cafés, hôtels et restaurants…), chiffre d'affaires, coûts de transport, de traitement et de stockage, coût complet, coût rapporté au CA et marge après logistique. Un coût / CA supérieur à 12 % est en rouge.
- (4) **Concentration du coût** (graphique) : Part cumulée du coût en fonction de la part des clients classés ; elle fait ressortir le petit nombre de clients qui concentre l'essentiel de la dépense.
- (5) **Comparer deux clients** (tableau) : Met côte à côte deux clients choisis dans le classement, par exemple deux clients au chiffre d'affaires voisin dont le coût logistique diffère fortement.
- (6) **Commandes du client** (bouton) : Affiche les commandes du client sélectionné avec le coût de chacune, ligne par ligne.
- (7) **Exporter** (bouton) : Copie le classement dans une feuille sans macros.


#### D10 — Coût détaillé d'une commande

*Accès : Accueil › Analyse › Coût d'une commande, ou double-clic sur une commande dans D9. Profils : tous.*

Descendre du total jusqu'à l'opération qui l'explique : pour une commande, l'écran montre ses jalons et ses délais, le coût de chaque composante, la mission qui l'a transportée avec la part de coût qui lui revient, et le coût de chaque ligne de produit. Exemple : la commande 256475 de l'Épicerie Al Baraka.

Éléments de l'écran :

- (1) **Recherche** (zone et liste) : Numéro de commande, ou client puis choix parmi ses commandes de la période.
- (2) **En-tête** (bandeau) : Client, segment, zone et distance, site expéditeur, type de commande, statut et montant ; repris de l'application de gestion.
- (3) **Jalons et délais** (indicateurs) : Dates de commande, de préparation, d'expédition, de livraison et date souhaitée ; délai total et retard éventuel, qui alimentent les indicateurs de service.
- (4) **Coût par composante** (graphique) : Transport, traitement et stockage imputés à la commande, coût complet, part du montant et marge après logistique.
- (5) **Coût par ligne** (tableau) : Pour chaque produit : quantité, montant, part de transport (selon le poids taxable), part de traitement (coût fixe partagé et coût par ligne), part de stockage, coût total et coût par colis. La ligne Total est égale au coût de la commande.
- (6) **Missions et partenaires** (tableau) : Mission(s) qui ont transporté la commande, véhicule ou transporteur, chauffeur, coût de la mission, clé de partage et part attribuée. Une commande livrée en plusieurs fois a une ligne par mission.
- (7) **Ouvrir dans l'application de gestion** (bouton) : Ouvre la commande dans son formulaire de saisie (T11), en lecture seule si la période est clôturée.
- (8) **Voir la mission** (bouton) : Ouvre la mission sélectionnée dans l'analyse du transport (D7), avec toutes les commandes qui la partagent.
- (9) **Méthode de calcul** (bouton) : Ouvre la fiche « coût complet d'une commande » de l'écran des méthodes (D12).
- (10) **Exporter le rapport** (bouton) : Produit le rapport de la commande, avec les règles de répartition appliquées.


### E. Statistiques, méthodes et formules de calcul

Pour compléter les totaux, l'écran des statistiques décrit la manière dont un coût se répartit entre les commandes, les missions, les clients ou les camions : moyenne, médiane, dispersion, distribution et valeurs extrêmes, comparées d'un groupe à l'autre, car une moyenne seule peut cacher quelques cas très coûteux. L'écran des méthodes de calcul explique chaque chiffre affiché : pour chaque coût, indicateur ou règle, il donne sa définition en langage courant, son unité, ses composantes, ses exclusions, les paramètres appliqués et un exemple. Les formules correspondantes sont réunies à la suite de cet écran, chacune avec un exemple que l'on peut refaire à partir des valeurs des captures.

#### D11 — Statistiques et dispersion

*Accès : Accueil › Analyse › Statistiques. Profils : tous.*

Compléter les totaux par une lecture statistique : comment un coût se répartit entre les commandes, les missions, les clients ou les camions, quelle est sa dispersion, et en quoi les groupes diffèrent. Une moyenne seule peut cacher quelques cas extrêmes qui expliquent l'essentiel de la dépense.

Éléments de l'écran :

- (1) **Sélection** (listes déroulantes) : Population étudiée (commandes, lignes, missions, clients, produits, camions, dossiers d'importation), indicateur (coût par commande, par colis, par kilomètre, par tonne, frais d'approche…), critère de regroupement (segment, site, mode de transport, fournisseur, famille) et période.
- (2) **Statistiques descriptives** (indicateurs) : Effectif, moyenne, médiane, écart type, minimum, maximum et coefficient de variation de l'indicateur sur la population. Une médiane bien inférieure à la moyenne signale quelques valeurs très élevées. Le coût d'une commande est son coût complet imputé (D10) : sous-activité et frais de structure, non imputés, n'y entrent pas.
- (3) **Distribution** (graphique) : Nombre d'éléments par tranche de valeur : il montre où se situe la majorité des cas et la longueur de la traîne des cas coûteux.
- (4) **Comparaison des groupes** (tableau) : Effectif, moyenne, médiane, écart type et part du coût total de chaque groupe, pour comparer les segments, les sites ou les modes sur une base homogène.
- (5) **Valeurs extrêmes** (tableau) : Éléments les plus éloignés de la médiane ; un double-clic ouvre leur détail (par exemple le coût d'une commande, D10).
- (6) **Avertissement** (message) : Signale les groupes à effectif trop faible pour tirer une conclusion, et les comparaisons de périodes de durées différentes.
- (7) **Ouvrir la valeur choisie** (bouton) : Ouvre l'écran de détail de l'élément sélectionné.
- (8) **Exporter** (bouton) : Ajoute les statistiques et la distribution au rapport de l'analyse.


#### D12 — Méthodes de calcul et définitions

*Accès : Accueil › Analyse › Méthodes et définitions, ou bouton « Définition » des analyses. Profils : tous.*

Expliquer chaque chiffre affiché : pour chaque coût, indicateur ou règle, une fiche donne sa définition en langage courant, son unité, ce qu'il inclut et exclut, les données et paramètres utilisés, la nature du montant et un exemple chiffré. Les paramètres affichés sont ceux réellement appliqués au dernier calcul.

Éléments de l'écran :

- (1) **Liste des termes** (liste) : Coûts, indicateurs et règles de calcul, regroupés par nature ; un clic affiche la fiche.
- (2) **Définition** (texte) : Ce que mesure le terme, en une ou deux phrases ; la formule exacte de chaque terme est récapitulée, avec un exemple recalculable, dans les formules qui suivent cet écran.
- (3) **Unité et période** (texte) : Unité d'expression (dirhams, jours, pourcentage, par colis…) et période sur laquelle il est calculé.
- (4) **Composantes incluses** (texte) : Éléments qui entrent dans le calcul, chacun compté une seule fois.
- (5) **Exclusions** (texte) : Éléments volontairement laissés de côté, comme la sous-activité, le coût d'achat, les droits de douane ou la TVA récupérable, pour éviter les doubles comptes.
- (6) **Données utilisées** (texte) : Feuilles et informations de la base qui alimentent le calcul.
- (7) **Paramètres appliqués** (texte) : Taux, coûts unitaires et seuils utilisés, avec leur valeur au dernier calcul ; ils se modifient dans les paramètres (T13).
- (8) **Nature du montant** (texte) : Enregistré (saisi), calculé à partir de montants enregistrés, estimé (coûts de référence, par exemple pour un nouveau produit) ou simulé (scénario).
- (9) **Exemple** (texte) : Application à un cas réel de la base, pour vérifier la compréhension.
- (10) **Écrans concernés** (texte) : Écrans et pages Power BI où le terme apparaît.
- (11) **Recherche** (zone de saisie) : Trouve un terme par un mot de son nom ou de sa définition.
- (12) **Imprimer la fiche** (bouton) : Imprime la fiche affichée ; « Exporter toutes les définitions » produit un glossaire complet, joint aux rapports.

### F. Scénarios et aide à la décision

L'utilisateur peut créer, enregistrer et comparer côte à côte des scénarios, toujours calculés sur une copie des données. Il peut faire varier le prix de n'importe quelle catégorie ou n'importe quel paramètre, par exemple une hausse de 15 % du gasoil, tester la sous-traitance d'une zone ou son retour à la flotte interne, ou changer la règle de partage du transport et mesurer l'effet sur le coût de chaque client. Des scénarios plus avancés calculent la quantité de commande la plus économique, le nombre de voyages à partir duquel un camion de l'entreprise devient plus rentable qu'un transporteur, l'effet d'un montant minimal de commande, du renouvellement d'un camion ou d'un meilleur équilibre entre maintenance préventive et curative, ainsi que la sensibilité du coût total aux principaux paramètres.

L'application permet aussi d'évaluer l'introduction d'un nouveau produit, en estimant son coût logistique, sa marge et son besoin de stockage avant de l'intégrer, et de comparer des fournisseurs sur leur coût rendu, leurs délais et leur fiabilité. Pour décider, l'utilisateur fixe un objectif et des contraintes, par exemple un délai maximal ou une marge minimale : l'application indique les scénarios qui les respectent, sans présenter le meilleur scénario testé comme un optimum, et la décision, retenue, à approfondir ou rejetée, est enregistrée avec sa justification. Un résultat défavorable déclenche un examen, jamais la suppression automatique d'un produit ou d'un partenaire.

#### D13 — Scénarios et simulations

*Accès : Accueil › Décision › Scénarios. Profils : administrateur, analyste (consultation : lecture seule).*

Construire des scénarios d'optimisation en combinant un ou plusieurs leviers, les calculer sur une copie des données, en mesurer l'effet, puis consigner la décision prise et sa justification. Exemple : remplacer les 114 navettes affrétées en 2025 (1 155 202 MAD) par une troisième semi-remorque de l'entreprise.

Éléments de l'écran :

- (1) **Liste des scénarios** (liste à colonnes) : Scénarios enregistrés, écart de coût obtenu et décision : en étude, à approfondir, retenu ou rejeté. Un clic affiche le scénario.
- (2) **Nom** (zone de saisie) : Nom court, obligatoire et unique ; il apparaît dans les comparaisons, les rapports et Power BI.
- (3) **Description** (zone de texte) : Hypothèse étudiée, en quelques mots.
- (4) **Référence** (liste déroulante) : Période réelle à partir de laquelle le scénario est simulé.
- (5) **Volumes** (boutons d'option) : Volumes constants pour isoler l'effet du levier, ou modifiés (en pourcentage) ; la différence de volume est alors affichée dans toutes les comparaisons.
- (6) **Levier** (liste déroulante) : Prix d'une catégorie ; paramètre (coût du capital, urgences…) ; sous-traitance d'une zone ou reprise par la flotte ; règle de partage ; tarifs ou quantités par envoi d'un transporteur ; quantité économique de commande ; seuil flotte / sous-traitance ; montant minimal de commande ; renouvellement d'un camion ; équilibre maintenance préventive / curative ; réduction du stock (couverture visée) ; retrait d'un produit (seuls les coûts réellement évitables disparaissent, la perte de chiffre d'affaires et de marge est chiffrée à part) ; ajout d'un produit (D14) ; changement, réduction ou arrêt d'un fournisseur (D15).
- (7) **Leviers du scénario** (liste à colonnes) : Leviers retenus et leurs réglages ; un scénario peut en combiner plusieurs.
- (8) **Calculer le scénario** (bouton principal) : Copie les données, applique les leviers, relance le calcul des coûts sur la copie et affiche le résultat.
- (9) **Enregistrer** (bouton) : Conserve le scénario, ses hypothèses, ses résultats et sa décision, pour les comparer et les transmettre à Power BI.
- (10) **Comparer et décider** (bouton) : Ouvre la comparaison des scénarios avec la situation réelle et les contraintes (D16).
- (11) **Nouveau** (bouton) : Crée un scénario vide.
- (12) **Supprimer** (bouton) : Supprime le scénario après confirmation ; un scénario retenu ne peut qu'être archivé.
- (13) **Résultat** (zone calculée) : Coût total simulé et écart avec la référence, en dirhams et en pourcentage.
- (14) **Statut de décision** (liste déroulante) : En étude, à approfondir, retenu ou rejeté. Un résultat défavorable déclenche un examen, jamais une suppression automatique d'un produit ou d'un partenaire.
- (15) **Justification** (zone de texte) : Motif de la décision et conditions éventuelles ; obligatoire pour retenir ou rejeter. L'auteur et la date sont enregistrés.
- (16) **Garantie** (message) : Les simulations ne modifient jamais les données ni les résultats réels.


#### D14 — Simulation d'un nouveau produit

*Accès : Accueil › Décision › Nouveau produit, ou levier « ajout d'un produit » d'un scénario. Profils : administrateur, analyste.*

Évaluer, avant de le référencer, ce que coûterait logistiquement un nouveau produit : coût par colis de chaque étape, marge après logistique, place nécessaire dans l'entrepôt et sensibilité au volume vendu. Les estimations reposent sur les coûts unitaires réels de 2025 de la base.

Éléments de l'écran :

- (1) **Désignation, famille et régime** (zones et listes) : Identité du produit envisagé ; la famille et le régime déterminent la zone de stockage et le taux de détention appliqués.
- (2) **Fournisseur et origine** (listes déroulantes) : Fournisseur pressenti et origine locale ou importée ; pour un import, les frais d'approche par colis sont estimés d'après les frais d'importation réels du fournisseur ou du pays.
- (3) **Caractéristiques du colis** (zones de saisie) : Poids, volume et colis par palette, obligatoires et positifs ; ils déterminent le poids taxable, la place occupée et la part du transport.
- (4) **Prix d'achat et de vente** (zones de saisie) : Par colis, en dirhams hors taxes ; ils servent à la marge et au coût de détention.
- (5) **Volume et stock** (zones et liste) : Ventes mensuelles prévues, couverture de stock visée et site de stockage.
- (6) **Clients visés** (liste déroulante) : Répartition prévue des ventes par segment ; elle détermine le coût moyen de traitement et de livraison par colis.
- (7) **Coût et marge par colis** (tableau) : Coût d'approvisionnement, de stockage, de traitement et de transport par colis, coût logistique, coût complet, marge brute et marge après logistique.
- (8) **Besoin de stockage** (indicateur) : Palettes occupées en moyenne et effet sur le taux d'occupation de la zone ; un dépassement du seuil de 90 % est signalé.
- (9) **Effet du volume** (tableau) : Marge après logistique pour plusieurs volumes de vente, pour savoir à partir de quel volume le produit atteint l'objectif de marge.
- (10) **Nature des montants** (message) : Rappelle que les montants sont estimés et non enregistrés.
- (11) **Calculer** (bouton principal) : Recalcule les estimations après toute modification.
- (12) **Enregistrer comme scénario** (bouton) : Crée un scénario « ajout d'un produit » (D13) pour le comparer et consigner la décision.
- (13) **Comparer à un produit existant** (bouton) : Affiche côte à côte le produit envisagé et un produit réel de la même famille.
- (14) **Limite** (message) : La simulation ne crée rien dans les données ; la fiche produit n'est créée que dans l'application de gestion, si le produit est retenu.


#### D15 — Comparaison des fournisseurs

*Accès : Accueil › Décision › Comparer les fournisseurs. Profils : tous (simulation : administrateur, analyste).*

Comparer les fournisseurs d'un produit ou d'une famille sur leur coût rendu, c'est-à-dire le prix d'achat augmenté des frais d'approche et du coût de stockage qu'impose la taille de leurs lots, mais aussi sur leurs délais, leur fiabilité et la dépendance de l'entreprise à leur égard, avant de changer, de réduire ou d'arrêter une relation.

Éléments de l'écran :

- (1) **Sélection** (listes et zone) : Produit ou famille, période d'observation et besoin annuel à couvrir.
- (2) **Contraintes** (zones de saisie) : Délai maximal et conformité minimale acceptés ; d'autres contraintes (quantité minimale, part maximale d'un seul fournisseur) peuvent être ajoutées.
- (3) **Tableau comparatif** (tableau) : Pour chaque fournisseur : prix d'achat, frais d'approche (fret, port, transit, assurance), coût de stockage induit par la quantité minimale, coût rendu par colis, délai moyen et retard moyen observés sur ses commandes, taux de livraisons conformes, et respect des contraintes.
- (4) **Coût rendu** (graphique) : Comparaison visuelle du coût rendu par colis.
- (5) **Dépendance** (texte calculé) : Part des achats assurée par chaque fournisseur et capacité de remplacement : délai et stock de sécurité nécessaires pour changer de fournisseur.
- (6) **Arbitrage** (message) : Signale le fournisseur le moins cher qui ne respecte pas les contraintes, pour éviter un choix fondé sur le seul prix.
- (7) **Simuler un changement** (bouton principal) : Crée un scénario (D13) qui transfère tout ou partie des achats vers le fournisseur choisi et en calcule l'effet sur les coûts, le stock et le service.
- (8) **Simuler une réduction ou un arrêt** (bouton) : Crée un scénario mesurant l'effet d'une réduction ou d'un arrêt de la relation : coûts de transition, délais et risque de rupture.
- (9) **Exporter** (bouton) : Ajoute la comparaison au rapport de décision.


#### D16 — Comparaison des scénarios et décision

*Accès : Accueil › Décision › Comparer et décider, ou bouton de l'écran D13. Profils : tous (décision : administrateur).*

Mettre côte à côte la situation réelle et plusieurs scénarios sur les mêmes indicateurs de coût, de délai, de marge et de service, vérifier lesquels respectent les contraintes fixées, repérer les paramètres les plus influents, puis consigner la décision. L'outil compare des options testées ; il ne prétend pas trouver un optimum global.

Éléments de l'écran :

- (1) **Scénarios comparés** (liste à cases) : Scénarios enregistrés à afficher, jusqu'à quatre à la fois.
- (2) **Objectif et contraintes** (cadre) : Objectif poursuivi (minimiser le coût total ou le coût par commande, maximiser la marge) et contraintes à respecter : délai moyen maximal, marge minimale, taux de service minimal. Les comparaisons se font à volumes constants, ou en affichant la différence de volume.
- (3) **Tableau comparatif** (tableau) : Coût total, écart et part due au volume, coût du transport, coût / CA, délai moyen, marge après logistique et taux de service pour le réel et chaque scénario ; ligne de respect des contraintes et décision enregistrée.
- (4) **Coût total par scénario** (graphique) : Comparaison visuelle du coût réel et des coûts simulés.
- (5) **Sensibilité** (graphique) : Effet sur le coût total d'une baisse puis d'une hausse de 10 % de chaque paramètre principal, du plus influent au moins influent.
- (6) **Conclusion** (message) : Désigne le scénario le moins coûteux parmi ceux qui respectent les contraintes, en précisant qu'il s'agit du meilleur des scénarios testés et non d'un optimum démontré.
- (7) **Boutons de décision** (boutons) : Retenir, à approfondir ou rejeter le scénario sélectionné ; une justification est demandée et enregistrée avec l'auteur et la date.
- (8) **Analyse de sensibilité** (bouton) : Recalcule le coût total en faisant varier un à un les paramètres choisis, sur une copie des données.
- (9) **Rapport de décision** (bouton) : Produit le rapport du scénario (D19) : hypothèses, référence, résultats, contraintes, décision et limites.


### G. Alertes, suivi et restitution

L'application signale automatiquement les situations anormales selon des seuils paramétrables : budget mensuel dépassé de plus de 10 %, facture supérieure au tarif de plus de 3 %, zone de stockage occupée à plus de 90 %, produit sans sortie depuis 60 jours, stock couvrant plus de 90 jours, client dont la logistique dépasse 12 % du chiffre d'affaires, tournée remplie à moins de 60 %, taux de service inférieur à 97 %, retard de livraison ou de fournisseur, risque de rupture, taux de pertes supérieur à 0,5 % des sorties ou marge après logistique inférieure à 5 %. Chaque alerte indique la mesure observée, le seuil, l'objet concerné et sa priorité, fondée sur l'ampleur du dépassement, le montant en jeu et l'urgence ; elle distingue une perte constatée d'un risque. Elle suit un cycle de traitement, à examiner, en cours, traitée ou écartée avec justification, avec une action, un responsable et une échéance, sans que l'observation d'origine soit jamais effacée. Une alerte critique, comme un écart avec le grand livre, bloque l'export vers Power BI jusqu'à sa résolution.

L'application produit trois types de rapports exportables en PDF : le rapport mensuel, le rapport de l'analyse affichée et le rapport de scénario et de décision. Chacun reprend la sélection, les résultats sous forme d'indicateurs, de graphiques, de statistiques et de tableaux, les méthodes de calcul, les alertes, la décision et les limites, et ses montants sont vérifiés contre ceux de l'application avant l'export. Elle prépare enfin le fichier de résultats destiné à Power BI, archive les indicateurs de chaque mois et peut exporter ses résultats dans un classeur sans macros.

#### D17 — Alertes

*Accès : Accueil › Suivi et restitution › Alertes. Profils : tous.*

Registre de toutes les situations anormales détectées au dernier calcul. Chaque alerte indique la règle franchie, la valeur mesurée et le seuil, sa priorité, sa nature (perte constatée, risque ou anomalie de données) et son statut de suivi, du constat jusqu'au traitement.

Éléments de l'écran :

- (1) **Filtres** (listes déroulantes) : Période, domaine, priorité et statut (ouvertes, traitées, écartées ou toutes).
- (2) **Matrice des priorités** (tableau) : Nombre d'alertes par domaine et par priorité, et nombre restant à examiner ; un clic sur une case filtre la liste.
- (3) **Liste des alertes** (tableau) : Priorité, règle déclenchée, objet, valeur mesurée, seuil, nature et statut (à examiner, en cours, traitée, écartée avec justification). Règles : coût par km d'un camion, budget mensuel dépassé de plus de 10 %, facture au-dessus du tarif de plus de 3 %, zone occupée à plus de 90 %, produit sans sortie depuis 60 jours ou couvert plus de 90 jours, risque de rupture, retard de livraison ou de fournisseur, pertes, client au-delà de 12 % du CA, marge faible ou négative, tournée remplie à moins de 60 %, taux de service sous 97 %, écart avec le grand livre au-delà de 0,5 %, incohérence de données.
- (4) **Blocage de l'export** (message) : Une alerte de rapprochement non traitée bloque l'export vers Power BI, pour qu'aucun chiffre faux ne soit présenté.
- (5) **Suivre l'alerte** (bouton principal) : Ouvre la fiche de suivi de l'alerte sélectionnée (D18).
- (6) **Voir l'analyse** (bouton) : Ouvre l'analyse correspondante, déjà filtrée sur l'objet (camion, zone, client, facture, fournisseur…).
- (7) **Recalculer** (bouton) : Compare de nouveau les résultats aux seuils, par exemple après une modification des paramètres.
- (8) **Exporter** (bouton) : Ajoute le registre au rapport, avec la règle, la priorité et le statut de chaque alerte.
- (9) **Origine des seuils** (information) : Les seuils se modifient dans les paramètres de l'application de gestion (T13), puis sont importés avec les données.
- (10) **Calcul de la priorité** (message) : La priorité repose sur des critères affichés : ampleur du dépassement, montant concerné et urgence. Aucune probabilité de perte n'est présentée sans méthode définie.


#### D18 — Suivi d'une alerte

*Accès : double-clic sur une alerte (D17 ou accueil D1). Profils : administrateur, analyste (consultation : lecture seule).*

Traiter une alerte de bout en bout : comprendre la règle déclenchée et les critères de sa priorité, qualifier sa nature, décider d'une action avec un responsable et une date, puis la clore ou l'écarter avec justification. L'historique garde la trace de chaque étape.

Éléments de l'écran :

- (1) **Identification** (bandeau) : Règle, objet, période concernée, date et heure de détection.
- (2) **Constat** (zones calculées) : Règle déclenchée exprimée en mots, valeur mesurée, seuil et écart en dirhams ; non modifiables.
- (3) **Priorité** (zone calculée) : Priorité et critères qui la justifient : ampleur du dépassement rapportée au seuil, montant concerné et urgence (par exemple un blocage d'export).
- (4) **Nature** (boutons d'option) : Perte constatée (coût déjà subi), risque (situation qui peut entraîner un coût) ou anomalie de données (erreur à corriger à la source).
- (5) **Statut** (boutons d'option) : À examiner, en cours, traitée, ou écartée avec justification. Le passage à « traitée » ou « écartée » exige une action ou une justification.
- (6) **Responsable** (liste déroulante) : Utilisateur chargé du suivi ; l'alerte apparaît dans sa liste à l'accueil.
- (7) **Date de suivi** (zone de saisie) : Échéance de la prochaine vérification ; une alerte dont la date est dépassée remonte en tête de liste.
- (8) **Action décidée** (zone de texte) : Mesure corrective prévue ou réalisée.
- (9) **Justification** (zone de texte) : Obligatoire pour écarter une alerte : motif pour lequel la situation est acceptée.
- (10) **Historique du suivi** (liste à colonnes) : Chaque changement de statut avec sa date, son auteur et son commentaire ; il ne peut pas être modifié.
- (11) **Enregistrer** (bouton principal) : Enregistre le suivi et l'inscrit à l'historique.
- (12) **Voir l'analyse** (bouton) : Ouvre l'écran d'analyse correspondant, filtré sur l'objet.
- (13) **Ouvrir dans l'application de gestion** (bouton) : Ouvre l'enregistrement en cause (dépense, facture, fiche) pour le corriger à la source.
- (14) **Règle de clôture** (message) : Traiter une alerte n'efface jamais le constat ; une situation qui persiste rouvre l'alerte automatiquement.


#### D19 — Rapports et exports

*Accès : Accueil › Suivi et restitution › Rapport et exports, ou bouton « Exporter le rapport » des analyses. Profils : administrateur, analyste (consultation : PDF seulement).*

Produire des rapports PDF fidèles à ce que montre l'application, pour un mois, pour l'analyse affichée ou pour un scénario, avec leurs méthodes et leurs limites ; alimenter Power BI ; transmettre les résultats sans macros ; archiver les indicateurs. L'export ne modifie ni les données ni les filtres de travail.

Éléments de l'écran :

- (1) **Type de rapport** (boutons d'option) : Rapport mensuel ; rapport de l'analyse affichée (il reprend la sélection et les filtres de l'écran d'où il est lancé) ; rapport de scénario et de décision.
- (2) **Mois ou objet** (liste déroulante) : Mois du rapport mensuel, ou scénario choisi pour un rapport de décision.
- (3) **Niveau de détail** (liste déroulante) : Synthèse sur une ou deux pages, ou rapport détaillé avec les tableaux complets.
- (4) **Sections** (cases à cocher) : Identification (titre, date de génération, période, référence) ; périmètre (filtres actifs, objets, unités) ; résultats (indicateurs, graphiques, tableaux) ; méthodes (règles de calcul, clés de répartition, taux, tarifs et seuils appliqués) ; alertes et leur suivi ; décision (hypothèses, écarts, contraintes, justification) ; limites (résultats indisponibles, montants estimés, hypothèses influentes).
- (5) **Aperçu** (aperçu) : Rendu du rapport avant export : titres lisibles, unités visibles, tableaux paginés et graphiques non coupés.
- (6) **Générer le rapport** (bouton principal) : Construit le rapport à partir des derniers résultats et de la sélection choisie.
- (7) **Exporter en PDF** (bouton) : Enregistre le rapport sous un nom qui indique son type, sa période et sa date de génération.
- (8) **Vérifier la concordance** (bouton) : Compare chaque montant du rapport à la valeur affichée par l'application pour la même sélection.
- (9) **Résultat de la vérification** (message) : Nombre de montants contrôlés et concordance ; tout écart est listé avant l'envoi du rapport.
- (10) **Exporter vers Power BI** (bouton) : Écrit les tables de résultats, en valeurs seules, dans un fichier daté que lit le rapport Power BI. Refusé si une alerte de rapprochement n'est pas traitée.
- (11) **Ouvrir le rapport Power BI** (bouton) : Lance Power BI Desktop, qui s'actualise à partir du dernier fichier de résultats.
- (12) **Exporter sans macros** (bouton) : Copie les résultats dans un classeur ordinaire, facile à transmettre.
- (13) **Archiver les indicateurs** (bouton) : Ajoute les indicateurs du mois à l'historique des courbes d'évolution ; un mois déjà archivé est remplacé après confirmation.
- (14) **Derniers exports** (liste à colonnes) : Nature, nom de fichier et date des derniers exports.

### Rapport Power BI

Le rapport Power BI répond à la demande de relier Excel à Power BI et de l'alimenter avec les données de l'application. Il se connecte au fichier de résultats de l'application décisionnelle, vérifie les données lors de leur chargement, puis les organise en un modèle simple : d'un côté les faits mesurés (dépenses réparties, lignes de commande avec leur coût complet, missions, stocks mensuels, événements et résultats des scénarios), de l'autre les axes d'analyse (calendrier, catégories de coûts, camions, sites, entrepôts et zones, prestataires et fournisseurs, personnel, clients et produits). Chaque table de faits est reliée aux axes par leurs codes. Power BI additionne et filtre des montants déjà calculés, ce qui garantit qu'il affiche les mêmes chiffres qu'Excel.

Toutes les pages partagent les mêmes filtres : période, site, catégorie de coût et objet analysé. Les pages de synthèse et de catégories de coûts constituent le minimum attendu ; s'y ajoutent des pages consacrées à la flotte, aux entrepôts, aux prestataires et fournisseurs, aux clients, aux produits, aux événements et à la comparaison des scénarios. Un curseur de simulation permet aussi d'illustrer en direct l'effet d'une variation de prix simple. Chaque page peut être exportée en PDF depuis Power BI Desktop.

#### P1 — Rapport Power BI : page Synthèse

*Accès : Power BI Desktop › Projet15_Rapport, ou bouton « Ouvrir le rapport Power BI ». Profils : tous (consultation).*

Première page du rapport, destinée à la direction et à la présentation en réunion. Elle répond à la demande de relier Excel à Power BI : elle lit le fichier de résultats de l'application décisionnelle et affiche les mêmes chiffres, sous une forme graphique et interactive.

Éléments de l'écran :

- (1) **Filtres communs** (segments) : Période, site, famille ou type de coût et objet analysé. Ils sont partagés par toutes les pages : une sélection faite ici s'applique partout.
- (2) **Cartes d'indicateurs** (cartes) : Coût logistique de la sélection, coût rapporté au chiffre d'affaires, écart au budget et coût par commande livrée ; les ratios sont calculés à partir des totaux de la sélection.
- (3) **Coût mensuel** (graphique en courbes) : Coût réel de chaque mois et budget correspondant, pour repérer les mois en dérive.
- (4) **Répartition par étape** (graphique en anneau) : Part de l'approvisionnement, du stockage, du traitement des commandes, du transport et de la structure dans le coût de la sélection.
- (5) **Postes budgétaires** (graphique en barres) : Postes budgétaires classés par montant ; un clic sur une barre filtre les autres visuels de la page.
- (6) **Écart au budget** (tableau) : Réel, budget et écart par poste budgétaire, avec mise en couleur des écarts importants.
- (7) **Navigation** (onglets de pages) : Accès aux neuf pages du rapport : synthèse et catégories (minimum attendu), puis flotte, entrepôts, prestataires, clients, produits, événements et scénarios (pages P2 à P4).
- (8) **Source et actualisation** (information) : Nom du fichier de résultats lu et date de la dernière actualisation ; le bouton « Actualiser » de Power BI relit le fichier le plus récent.


#### P2 — Rapport Power BI : page Catégories de coûts

*Accès : onglet « Catégories » du rapport. Profils : tous (consultation).*

Détail de la dépense par catégorie de coût, par étape de la chaîne et par compte comptable sur la période choisie. Avec la page Synthèse, elle constitue le minimum attendu du rapport. Power BI n'y refait aucun calcul de répartition : il additionne et filtre des montants déjà calculés par l'application décisionnelle.

Éléments de l'écran :

- (1) **Matrice des catégories** (matrice) : Coût par famille et par type de coût en lignes et par mois en colonnes ; les familles se développent pour afficher leurs types, comme dans l'analyse D6.
- (2) **Coût par étape** (graphique en barres) : Montant de chaque étape de la chaîne logistique sur la sélection.
- (3) **Coût par compte** (graphique en barres) : Principaux comptes comptables, utiles pour le dialogue avec la comptabilité.
- (4) **Détail des dépenses** (tableau) : Dépenses qui composent la cellule ou la barre sélectionnée ; un clic dans un visuel filtre ce tableau.
- (5) **Curseur de simulation** (paramètre de simulation) : Fait varier en direct le prix d'une catégorie (ici le carburant) entre −20 % et +20 % et affiche le coût simulé. Il illustre une variation simple ; les scénarios complets sont calculés dans l'application décisionnelle (D13).


#### P3 — Rapport Power BI : page Flotte

*Accès : onglet « Flotte » du rapport. Profils : tous (consultation).*

Exemple des pages thématiques du rapport : la page Flotte compare les camions entre eux. Les autres pages thématiques suivent la même construction (filtres communs, cartes, graphiques, tableau) et sont décrites en P4.

Éléments de l'écran :

- (1) **Cartes de la flotte** (cartes) : Coût complet de la flotte, coût par kilomètre, taux de disponibilité et remplissage moyen, avec leur objectif ou leur seuil.
- (2) **Coût par kilomètre** (graphique en barres) : Camions classés par coût au kilomètre ; un clic sur un camion filtre toute la page sur ce véhicule.
- (3) **Composition du coût** (graphique en colonnes) : Carburant, amortissement, maintenance curative et préventive, assurance : composition du coût complet de la sélection.
- (4) **Immobilisations** (tableau) : Jours d'immobilisation, nombre de pannes et coût des incidents par camion, issus du registre des événements.
- (5) **Interne et sous-traité** (graphique en barres) : Part du transport assurée par la flotte de l'entreprise et par les transporteurs.


#### P4 — Rapport Power BI : autres pages et organisation des données

Les pages ci-dessous complètent le rapport. Toutes partagent les filtres de période, de site, de catégorie et d'objet de la page Synthèse, et toutes lisent le même fichier de résultats, si bien que leurs chiffres sont toujours identiques à ceux de l'application décisionnelle.

| Page | Contenu et visuels | Questions auxquelles elle répond |
|---|---|---|
| Entrepôts | Cartes : coût du stockage, sous-activité, occupation moyenne. Graphiques : coût par palette et par mois par zone, occupation par zone avec seuil de 90 %, coût de sous-activité par site. Tableau : produits dormants. | Quelles zones sont saturées ou sous-utilisées ? Combien coûtent les emplacements vides ? Quels produits immobilisent de la place sans se vendre ? |
| Prestataires | Cartes : montant facturé, nombre de factures, écart moyen aux tarifs. Graphiques : montant par transporteur et par fournisseur, écart facture / tarif par transporteur. Tableau : factures au-dessus du tarif. | Quels prestataires pèsent le plus ? Lesquels facturent au-dessus des tarifs convenus ? La dépense est-elle concentrée sur peu de prestataires ? |
| Clients | Cartes : coût logistique, coût / CA, marge après logistique. Graphiques : courbe de concentration du coût, coût / CA par segment, nuage chiffre d'affaires / coût logistique par client. Tableau : clients au-delà de 12 % du CA. | Quels clients concentrent le coût ? Lesquels coûtent cher au regard de leur chiffre d'affaires ? Quel segment est le moins rentable après logistique ? |
| Produits | Graphiques : coût logistique par famille, coût par colis livré, part du stockage dans le coût de chaque famille. Tableau : produits les plus coûteux avec leur rotation et leur couverture. | Quelles familles et quels produits coûtent le plus à stocker et à livrer ? Les produits frais et surgelés justifient-ils leur surcoût ? |
| Événements | Cartes : nombre d'événements, coût total, jours d'immobilisation. Graphiques : événements par type et par gravité, coût par type, évolution mensuelle. Tableau : événements ouverts. | Quels incidents coûtent le plus ? Quels camions ou quelles zones sont le plus souvent immobilisés ? |
| Scénarios | Graphique : coût total du réel et de chaque scénario enregistré. Tableau : écarts par étape et par indicateur, comme dans l'écran D16. | Quelle option réduit le plus le coût ? Avec quel effet sur les autres indicateurs ? |

Organisation des données dans Power BI. Le fichier de résultats contient deux sortes de tables. Les tables de faits regroupent ce qui est mesuré : les dépenses réparties, les lignes de commande avec leur coût complet, les missions, les stocks mensuels par zone, les événements et les résultats des scénarios. Les tables d'axes regroupent ce selon quoi on analyse : le calendrier, les catégories de coûts, les camions, les sites, entrepôts et zones, les prestataires et fournisseurs, le personnel, les clients et les produits. Chaque table de faits est reliée aux axes par leurs codes, ce qui permet de filtrer toutes les pages avec les mêmes segments. Les données sont vérifiées au chargement (types, codes manquants), et toutes les mesures sont de simples sommes ou des ratios de sommes.

---

## 15. Le socle commun : code déjà écrit et testé

Ce code est dans le module `modUtils` du fichier maître de gestion. La partie 1 le livre aux trois autres membres en fichier `.bas` ; personne ne le réécrit de son côté. Pour l'importer : dans l'éditeur VBA, Fichier → Importer un fichier.

### 12.1 DerniereLigne — terminée

```vba
'------------------------------------------------------------------
' DerniereLigne
' Objet   : numéro de la dernière ligne non vide d'une colonne
' Entrées : nomFeuille - nom de l'onglet
'           colonne    - lettre de la colonne lue ("A" par défaut)
'           wb         - classeur lu (ce classeur si on ne le donne pas)
' Sortie  : numéro de ligne (fin du tableau Excel si la feuille en contient un)
'------------------------------------------------------------------
Public Function DerniereLigne(nomFeuille As String, Optional colonne As String = "A", _
                              Optional wb As Workbook) As Long
    Dim ws As Worksheet
    If wb Is Nothing Then Set wb = ThisWorkbook
    Set ws = wb.Worksheets(nomFeuille)
    DerniereLigne = ws.Cells(ws.Rows.Count, colonne).End(xlUp).Row
End Function
```

Le paramètre `wb` existe pour l'application décisionnelle : elle appelle cette fonction sur le classeur de gestion qu'elle vient d'ouvrir, et non sur le sien.

Résultats vérifiés sur la base : Commandes 6475, Lignes_Commandes 44005, Missions 3924, Produits 141, Clients 159, Vehicules 28, Ref_Data 43, Ref_Data colonne H 9.

### 12.2 ChercherLigne — terminée

```vba
'------------------------------------------------------------------
' ChercherLigne
' Objet   : numéro de la ligne où se trouve une valeur dans une colonne
' Entrées : nomFeuille - nom de l'onglet
'           colonne    - lettre de la colonne où chercher
'           valeur     - valeur cherchée (texte ou nombre)
'           wb         - classeur lu (ce classeur si on ne le donne pas)
' Sortie  : numéro de ligne, ou 0 si la valeur n'existe pas
'------------------------------------------------------------------
Public Function ChercherLigne(nomFeuille As String, colonne As String, valeur As Variant, _
                              Optional wb As Workbook) As Long
    Dim ws As Worksheet
    Dim derniere As Long
    Dim i As Long
    If wb Is Nothing Then Set wb = ThisWorkbook
    Set ws = wb.Worksheets(nomFeuille)
    derniere = DerniereLigne(nomFeuille, colonne, wb)
    For i = 2 To derniere
        If CStr(ws.Cells(i, colonne).Value) = CStr(valeur) Then
            ChercherLigne = i
            Exit Function
        End If
    Next i
    ChercherLigne = 0
End Function
```

Résultats vérifiés : P0001 → 2, P0140 → 141, CL0020 → 21, VH01 → 2, mission 601904 → 1905, commande 252590 → 2591 (que le numéro soit donné en nombre ou en texte), P9999 → 0.

### 12.3 LireParametre — en cours (module modParametres)

```vba
'------------------------------------------------------------------
' LireParametre
' Objet   : valeur d'un paramètre de la table T_Parametres (Ref_Data)
' Entrée  : code - code du paramètre, par exemple "MAJ_URGENCE"
' Sortie  : valeur du paramètre (nombre, texte ou date) ;
'           vide si le code n'existe pas (un message prévient)
'------------------------------------------------------------------
Public Function LireParametre(code As String) As Variant
    Dim ligne As Long
    ligne = ChercherLigne("Ref_Data", "A", code)
    If ligne = 0 Then
        MsgBox "Paramètre introuvable : " & code, vbExclamation, "Paramètres"
        LireParametre = Empty
    Else
        LireParametre = ThisWorkbook.Worksheets("Ref_Data").Range("C" & ligne).Value
    End If
End Function
```

Test de validation, qui est aussi la formule 5 du cahier des charges :

```
?LireParametre("A_ADV_REF") + 4 * LireParametre("B_PREP_REF")
  → 312,12
?(LireParametre("A_ADV_REF") + 4 * LireParametre("B_PREP_REF")) * (1 + LireParametre("MAJ_URGENCE"))
  → 405,756
```

### 12.4 Reste du socle, à écrire par la partie 1

| Fonction | Module | Rôle |
|-----------------------------------------|----------------|-----------------------------------------------------------------|
| `EcrireParametre(code, valeur)` | modParametres | Modifie un paramètre après contrôle des bornes ; administrateur seulement ; journalise |
| `ProchainNumero(compteur)` | modUtils | Lit la prochaine valeur dans T_Compteurs, l'incrémente et la renvoie |
| `FormatNumero(prefixe, numero, chiffres)` | modUtils | Fabrique DEP-2025-00001, EVT-0001, TC-01 |
| `MsgInfo`, `MsgErreur`, `Confirmer` | modMessages | Messages uniformes décrits par l'écran T18 |
| `GererErreur(procedure)` | modMessages | Message compréhensible au lieu de la fenêtre technique d'Excel |
| `Journaliser(action, objet, detail)` | modJournal | Écrit une ligne dans T_Journal |
| `EstNombrePositif`, `EstDateValide` | modUtils | Contrôles de saisie communs à tous les formulaires |

---

## 16. Pièges rencontrés et décisions prises

Ces points ont coûté du temps à la partie 1. Ils valent pour tout le monde.

1. **`End(xlUp)` et les tableaux Excel.** Toutes les feuilles de la base sont des tableaux nommés `T_…`. `End(xlUp)` s'arrête à la dernière ligne du tableau, même si la cellule y est vide. `DerniereLigne("Produits", "H")` renvoie donc 141 et non 140, alors que la cellule H141 est bien vide (c'est l'anomalie A1). Conséquences : on lit toujours la colonne A, celle du code, qui n'est jamais vide ; et pour supprimer un enregistrement, on supprime la ligne entière, jamais son seul contenu, sinon la ligne vide reste dans le tableau.
2. **`CStr` avant de comparer.** Le numéro de commande 252590 est un nombre dans la feuille ; un numéro tapé dans un formulaire est un texte. Sans `CStr` des deux côtés, la comparaison échoue et la recherche renvoie 0.
3. **Tester une fonction.** Le bouton Exécuter (triangle vert) ne lance que les macros sans paramètres ; il ouvre une fenêtre « Macros » vide pour une fonction qui attend des arguments. On teste dans la fenêtre Exécution (Ctrl+G) en tapant `?NomFonction(arguments)` puis Entrée.
4. **Ctrl+Z n'annule rien de ce qu'une macro a écrit.** On garde une copie datée du classeur avant tout test qui écrit, et avant toute intégration.
5. **Enregistrer le code.** Écrire dans l'éditeur VBA ne suffit pas : il faut Ctrl+S dans Excel, en gardant le format « prenant en charge les macros ».
6. **Double extension.** Un fichier enregistré depuis Excel en tapant soi-même « .xlsm » donne `nom.xlsm.xlsm`. Afficher les extensions dans l'Explorateur et vérifier.

---

## 17. Calendrier et intégration

Les fonctions indispensables doivent être terminées le 15/10 dans les quatre parties ; le code est gelé le 20/10 au soir pour laisser deux jours de recette.

| Date | Jalon | Qui |
| --- | --- | --- |
| 11/10 | Socle commun livré en fichiers .bas ; tables vides Depenses et Clotures créées avec leurs en-têtes | Parties 1 et 2 |
| 12/10 | Tables Types\_Couts et Listes remplies ; Projet15\_Decisionnel.xlsm créé avec ses feuilles | Parties 1 et 3 |
| 14/10 | Reprise des dépenses 2025 terminée (4 169 dépenses) ; import D2 opérationnel | Parties 2 et 3 |
| 15/10 | Indispensables terminés ; première intégration (gestion par la partie 1, décisionnel par la partie 3) ; structure Calc\_… et CalculerCouts livrées à la partie 4 | Tous |
| 18/10 | Importants terminés ; deuxième intégration ; Power BI branché sur le vrai fichier de résultats | Tous |
| 20/10 | Souhaitables selon le temps ; gel du code le soir | Tous |
| 21–22/10 | Recette sur la base complète avec les tests de chaque partie, notice d'utilisation, répétition | Tous |
| 23/10 | Présentation à l'examen final | Tous |

---

## 18. Priorités, règles de gestion, utilisateurs et contraintes

### Priorités de réalisation

Les fonctions sont classées en quatre niveaux selon la méthode MoSCoW : indispensable, important, souhaitable et exclu. Les fonctions indispensables seront terminées en premier, de sorte que la solution réponde à toutes les demandes de l'encadrement même si le temps venait à manquer. Elles comprennent le menu, les fiches de référence, les listes de types, la gestion complète des dépenses, les événements et commentaires, l'import, le calcul par catégorie et le coût global, la synthèse et les indicateurs, la liaison avec Power BI et les premiers scénarios.

Le niveau important regroupe les commandes et les missions, le calcul détaillé du transport, du stockage et du traitement, la répartition jusqu'au client et au produit, le coût détaillé d'une commande, les contrôles et la clôture, les alertes et leur suivi, et les rapports PDF. Le niveau souhaitable regroupe la connexion par profil, le journal, les sauvegardes, les dépenses récurrentes, les statistiques, la simulation d'un nouveau produit, la comparaison des fournisseurs et les scénarios avancés. Sont exclus de cette version l'échange de fichiers avec un logiciel de gestion d'entreprise, l'envoi automatique des rapports par messagerie et l'historique des paramètres.


### Données et règles de gestion

### Données d'entrée

La solution s'appuie sur la base « Atlantis Distribution 2025 », un classeur Excel de 43 feuilles et d'environ 200 000 lignes. Cette base est cohérente, car elle est issue d'une simulation jour par jour de l'année : les stocks, les mouvements, les commandes, les missions et les charges se correspondent, comme le vérifient les 15 contrôles intégrés au classeur. L'application transactionnelle sera développée directement dans ce classeur.

La base contient les paramètres (coûts unitaires, seuils, calendrier, taux de change, prix du gasoil), les fiches de référence et les grilles tarifaires des transporteurs, les données de ventes et de transport (6 474 commandes et leurs 44 004 lignes, 3 923 missions, 2 384 lignes de factures des transporteurs, navettes et retours), les données d'approvisionnement et de stock (commandes aux fournisseurs, frais d'importation, mouvements et relevés de stock) et les données de coûts (grand livre de 33,8 millions de dirhams, budget, pleins de carburant et interventions de maintenance). Quatre feuilles la documentent : une notice, un dictionnaire des colonnes, la liste des anomalies volontaires et les contrôles de cohérence. L'application transactionnelle y ajoute quatre tables : les dépenses, les types de coûts, les événements et les commentaires. La base comprend enfin des cas réels que la solution doit traiter correctement, comme les ruptures de stock, les livraisons partielles, les commandes annulées, les retours, les commandes enlevées par le client ou les missions sous-traitées.

### Règles de gestion

Tous les montants sont exprimés en dirhams hors taxes, et une dépense en devise est convertie au taux enregistré. Une dépense qui couvre une période est étalée au prorata des jours, alors qu'une dépense ponctuelle est rattachée au jour de sa pièce. Une dépense est imputée directement à l'objet qu'elle concerne lorsqu'il est connu ; sinon, elle est partagée selon une clé adaptée à sa nature, comme la capacité pour un entrepôt, les kilomètres pour la flotte ou le nombre de commandes pour l'administration des ventes, et les charges qui ne concernent aucune activité précise restent des frais de structure.

Le coût d'une mission est partagé entre les commandes transportées selon la règle choisie, et une commande livrée en plusieurs fois reçoit sa part de chaque mission. Une commande annulée ne supporte que son coût administratif, et une commande enlevée par le client ne supporte aucun coût de transport. Le coût d'un client, d'un produit ou d'un segment est la somme des coûts de ses lignes de commande. Le total des coûts répartis et non répartis doit toujours être égal au total des dépenses de la période, les ratios sont calculés à partir des totaux de la sélection, et tous les seuils d'alerte restent des paramètres modifiables.

### Interfaçage avec d'autres systèmes

La solution est autonome et n'échange que des fichiers, toujours dans le même sens. La base Atlantis 2025 est intégrée à l'application transactionnelle, que seules les saisies des formulaires modifient. L'application décisionnelle lit l'application transactionnelle en lecture seule, en retrouvant son emplacement grâce à un paramètre, produit un fichier de résultats daté que Power BI actualise à la demande, et exporte ses rapports en PDF. Elle n'est reliée directement à aucun logiciel de gestion, d'entrepôt ou de comptabilité.

### Utilisateurs

Trois profils sont prévus, chacun avec des droits adaptés à son rôle. L'accès se fait par un formulaire de connexion avec un mot de passe masqué, et les boutons qui ne correspondent pas aux droits du profil connecté sont désactivés. L'administrateur, en principe le contrôleur de gestion, dispose de tous les droits : paramètres, fiches et types de coûts, suppressions définitives, clôture, réinitialisation et mots de passe. L'analyste, qui correspond aux personnes chargées de la saisie et de l'analyse au quotidien, enregistre les dépenses, les événements, les commentaires, les commandes et les missions, importe les données, lance les calculs, construit des scénarios et exporte les rapports. Le profil de consultation, destiné à la direction, permet de consulter les fiches, la synthèse, les rapports et Power BI sans rien modifier.

### Contraintes

### Coûts

Le projet ne dispose d'aucun budget et sera réalisé sans achat de logiciel ni de licence. Il utilise Microsoft Excel 2016 ou une version plus récente, avec son éditeur VBA, et Power BI Desktop, qui est gratuit, sur les ordinateurs des membres du groupe sous Windows et sur le poste de la salle de présentation. Le temps disponible correspond aux séances 4, 5 et 6 du module et au travail personnel du groupe.

### Délais

Le cahier des charges sera envoyé aux professeurs le 6 octobre 2026. Les fonctions indispensables devront être opérationnelles à l'issue de la séance 5, puis la solution complète sera vérifiée sur l'ensemble de la base avant d'être présentée et démontrée lors de l'examen final du 23 octobre 2026.

### Autres contraintes

Les applications seront réalisées avec les méthodes vues en cours : feuille de menu à boutons, feuille de paramètres, formulaires de saisie et procédures VBA déclenchées par les actions de l'utilisateur. Les montants sont exprimés en dirhams hors taxes et les comptes du plan comptable marocain sont repris tels qu'ils figurent dans la base. Les données étant fictives, les classeurs ne contiennent aucune donnée personnelle réelle et peuvent être diffusés sans restriction.

La solution devra aussi respecter plusieurs exigences de qualité. Elle devra être rapide, avec un calcul de l'année en moins d'une minute et des formulaires qui s'ouvrent en moins de deux secondes, et simple d'emploi, avec toute fonction accessible en deux clics depuis le menu et des messages en français. Elle devra être fiable, en donnant les mêmes résultats d'un calcul à l'autre et les mêmes chiffres dans Power BI, robuste, sans aucune erreur non maîtrisée même sur une base vidée de ses opérations, et évolutive, puisque les coûts, les seuils et les listes doivent rester modifiables sans programmation. Elle est conçue pour Excel sous Windows, Power BI Desktop ne fonctionnant que sur ce système.

---

## 19. Prompt de démarrage, à copier dans Claude

Chacun colle le prompt de sa partie dans une nouvelle conversation, en joignant ce fichier `.md` et, si possible, le classeur `Projet15_Couts_Logistiques.xlsm`.

### Partie 1 — Socle commun, accès et paramétrage

> Je suis étudiant à EMINES (UM6P), en groupe de quatre sur le projet VBA n°15 « Analyse des coûts logistiques ». Le dossier complet du projet est dans le fichier joint. Je prends la partie 1 : le socle commun, l'accès et le paramétrage, c'est-à-dire les modules modUtils, modParametres, modMessages et modJournal, puis les écrans T1, T2, T3, T4, T5, T6, T13, T16, T17 et T18.
>
> DerniereLigne et ChercherLigne sont terminées et testées ; LireParametre est écrite et doit être testée. Je travaille sur Excel sous Windows, dans Documents\Projet15\Projet15_Couts_Logistiques.xlsm. Je suis débutant en VBA.
>
> Travaillons procédure par procédure : donne-moi l'énoncé d'une seule procédure à la fois, avec son code à compléter ou à coller et les résultats chiffrés attendus calculés sur la base Atlantis ; je te renvoie ce que j'obtiens dans la fenêtre Exécution et tu corriges. Ne me donne jamais plusieurs procédures d'un coup. Commence par vérifier LireParametre, puis enchaîne sur ProchainNumero et FormatNumero.

### Partie 2 — Opérations, contrôles et clôture

> Je suis étudiant à EMINES (UM6P), en groupe de quatre sur le projet VBA n°15 « Analyse des coûts logistiques ». Le dossier complet du projet est dans le fichier joint. Je prends la partie 2 : les opérations, les contrôles et la clôture, c'est-à-dire la reprise des dépenses 2025 et les écrans T7, T8, T9, T10, T11, T12, T14 et T15.
>
> Je travaille sur Excel sous Windows, dans une copie de Projet15_Couts_Logistiques.xlsm ; je recevrai le socle commun (modUtils, modParametres) de mon camarade de la partie 1 et je ne le réécris pas. Les structures de mes tables (T_Depenses, T_Evenements, T_Commentaires, T_ResultatsControles, T_Clotures) sont fixées dans le dossier joint et je ne dois pas en changer les colonnes. Je suis débutant en VBA.
>
> Travaillons procédure par procédure : donne-moi l'énoncé d'une seule procédure à la fois, avec son code à compléter ou à coller et les résultats chiffrés attendus ; je te renvoie ce que j'obtiens et tu corriges. Commence par créer les feuilles et les tableaux de mes tables, puis par le module modReprise qui transforme les 4 169 écritures de la feuille Charges en dépenses, en respectant la règle anti-double-comptage du dossier.

### Partie 3 — Calcul et analyse des coûts

> Je suis étudiant à EMINES (UM6P), en groupe de quatre sur le projet VBA n°15 « Analyse des coûts logistiques ». Le dossier complet du projet est dans le fichier joint. Je prends la partie 3 : le calcul et l'analyse des coûts, c'est-à-dire l'application décisionnelle Projet15_Decisionnel.xlsm et les écrans D1 à D12.
>
> Je travaille sur Excel sous Windows. Je recevrai le socle commun de mon camarade de la partie 1 et la table des dépenses de celui de la partie 2 ; je ne réécris ni l'un ni l'autre. Les 20 formules de calcul et la structure des feuilles de résultats Calc_… sont fixées dans le dossier joint. Je suis débutant en VBA.
>
> Travaillons procédure par procédure : donne-moi l'énoncé d'une seule procédure à la fois, avec son code à compléter ou à coller et les résultats chiffrés attendus, que tu calcules sur la base Atlantis ; je te renvoie ce que j'obtiens et tu corriges. Commence par créer le classeur décisionnel avec ses feuilles, puis par ImporterPeriodes.

### Partie 4 — Décision, alertes, restitution et Power BI

> Je suis étudiant à EMINES (UM6P), en groupe de quatre sur le projet VBA n°15 « Analyse des coûts logistiques ». Le dossier complet du projet est dans le fichier joint. Je prends la partie 4 : la décision, les alertes, la restitution et Power BI, c'est-à-dire les écrans D13 à D19 et les pages P1 à P4.
>
> Je travaille sur Excel et Power BI Desktop sous Windows. Je recevrai de mon camarade de la partie 3 la structure des feuilles Calc_… et la fonction CalculerCouts ; en attendant je travaille sur un fichier de test qui respecte exactement cette structure, donnée dans le dossier joint. Je suis débutant en VBA et en Power BI.
>
> Commence par me faire fabriquer le fichier de test aux bonnes colonnes, puis l'export vers Power BI et la page P1. Donne-moi une seule étape à la fois, avec les résultats chiffrés attendus ; je te renvoie ce que j'obtiens et tu corriges. Pour Power BI, donne-moi les requêtes Power Query et les mesures DAX une par une, avec le résultat attendu de chacune.

---

## 20. Ce qu'il faut savoir pour l'oral du 23/10

Chacun doit pouvoir expliquer son propre code. Questions probables, à préparer par toutes les parties :

- Pourquoi `ThisWorkbook` et non `ActiveWorkbook` ? Que se passerait-il sans le paramètre `wb` quand l'application décisionnelle lit le classeur de gestion ?
- Pourquoi `ws.Rows.Count` plutôt que 1048576 ?
- Pourquoi convertir en texte avec `CStr` avant de comparer deux valeurs ?
- Pourquoi aucun coût unitaire ni aucun seuil n'est écrit dans le code ?
- Pourquoi désactive-t-on un élément au lieu de le supprimer dès qu'il possède un historique ?
- Pourquoi l'application décisionnelle refuse-t-elle une période non clôturée ?
- Comment garantit-on que Power BI affiche les mêmes chiffres qu'Excel ?
- Pourquoi la somme des coûts répartis doit-elle être égale au total des dépenses, et que fait-on des emplacements vides et des frais de direction ?

Résultats à connaître par cœur : coût logistique 2025 de 33 796 058 MAD, soit 13,8 % d'un chiffre d'affaires de 245,4 millions ; budget dépassé de 957 058 MAD, soit 2,9 % ; 6 474 commandes, 44 004 lignes, 3 923 missions ; traitement d'une commande de 4 lignes à 312,12 MAD ; six anomalies volontaires à détecter.
