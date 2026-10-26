# Test du clic unique — 22 octobre 2026

Deux clics sur le bouton existant du UserForm ont été vérifiés sans intervention sur le bouton Actualiser de Power BI.

| Test | Début | Résultat | Export lu dans Power BI |
|---|---|---|---|
| Rapport fermé | 2026-10-22T15:56:47.5336116+00:00 | Ouverture, actualisation et sauvegarde réussies | Resultats_PowerBI_2026-10-22_1556.xlsx |
| Rapport déjà ouvert | 2026-10-22T16:00:07.0379262+00:00 | Même fenêtre réutilisée, actualisation et sauvegarde réussies | Resultats_PowerBI_2026-10-22_1600.xlsx |

Le même processus Power BI (37364) a été conservé entre les deux passages. Le second clic n’a pas ouvert de doublon.

À chaque passage, le nom du fichier et l’heure d’actualisation du modèle ont changé, et le coût global, le CA, les commandes et le budget concordent avec Excel au centime. Les traces complètes figurent dans Test_Clic_Rapport_Ferme.json et Test_Clic_Rapport_Ouvert.json.

Le calcul natif a aussi été relancé lorsque l’import de 14:53 était plus récent que le calcul enregistré dans l’état de l’application. Les contrôles de conservation ont été vérifiés avant l’export. Le VBA n’a pas été modifié ; les résultats et traces du calcul/export ont été actualisés normalement en mémoire.

Les chemins C3/C4 ont été réglés dans le décisionnel ouvert. Aucun enregistrement du décisionnel ou du transactionnel n’a été lancé par l’agent. Enregistrer le décisionnel reste nécessaire pour conserver ces réglages.
