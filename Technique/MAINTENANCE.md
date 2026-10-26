# Maintenance de l’intégration

Les deux XLSM sont des copies binaires des dernières sources enregistrées. Le code VBA n’a pas été modifié.
Le script de configuration utilise Excel COM pour modifier Param_Dec C2, C3 et C4 dans le seul classeur du dossier courant ; il laisse l’enregistrement à l’utilisateur.
Le raccourci est créé sur le PC destinataire et appelle Windows PowerShell sans option de contournement de politique.
Le lanceur reprend l’automatisation validée, avec une vérification supplémentaire du chemin complet du classeur. Il lit l’état, refuse les périodes non clôturées, recalcule si l’import est plus récent, vérifie les contrôles, appelle modCalcul.ExporterPowerBI et génère les références Budget/Parametres depuis la mémoire Excel.
Le PBIT sélectionne le dernier export natif daté et contrôle Calc_Controles avant chargement. Son paramètre DossierExports est vide pour la première installation. Le PBIX fourni reste un instantané de référence jusqu’à sa recréation locale.
L’actualisation et l’enregistrement passent par les boutons natifs Power BI via UIAutomation. Les mesures principales sont lues par ADOMD sur le modèle local ; cette dépendance explique l’attente d’une installation classique dans Program Files. Une autre version de l’interface peut nécessiter une adaptation des identifiants des contrôles.
Les scripts ne modifient aucune politique Windows/Office, ne débloquent aucun fichier automatiquement, n’enregistrent pas les XLSM et n’envoient rien à un service distant.
Le dossier doit être accessible en écriture et le bureau interactif déverrouillé. Les applications doivent être ouvertes dans le même compte/session Windows, sans élévation différente. Aucun Python requis sur le PC du destinataire.
Après déplacement : refaire le PBIT, enregistrer le PBIX au chemin exact, fermer Power BI, reconfigurer les connexions et enregistrer le décisionnel.
Les preuves historiques portent sur le PC d’Ayoub ; les nouveaux tests créent leurs traces directement dans Rapports.


## Correctif des contrôles du lanceur — 22 octobre 2026

Le diagnostic reçu du PC ACH porte sur juin 2025 : les six états Excel sont OK, aucun rapprochement ouvert, avec un écart de 782,99 MAD (0,0282 % du total ; 0,0515 % du transport). Le lanceur imposait à tort une limite absolue supplémentaire de 0,01 MAD à ces contrôles internes. Il respecte maintenant la validation native d’Excel et conserve les écarts dans le journal et les visuels. Les états différents de OK, les contrôles absents et les alertes de rapprochement ouvertes restent bloquants. La comparaison des quatre montants Excel/Power BI au centime est inchangée. Aucune donnée ou code VBA n’est modifié. La validation complète de l’actualisation reste à effectuer sur le PC ACH.
