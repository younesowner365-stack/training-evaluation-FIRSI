# FIRSI V8.3 — version pilote

- Consultations réutilisables et séparées des évaluations.
- RH/Admin : brouillon, formations, publication/clôture, tableau de bord nominatif, export Excel.
- Collaborateur : réponses libres, modifiables jusqu’à clôture.
- Contrôles d’accès serveur et jeton CSRF sur toutes les écritures du nouveau module.
- SECRET_KEY robuste obligatoire en production, suppression du mot de passe admin de secours prédictible.

## Limites avant mise en production
- Ne pas déployer avant tests intégration et audit sécurité des routes historiques.
- Les consultations de décembre ne sont pas précréées : création par RH depuis le module.
- PDF de consultation et questionnaires personnalisables non encore intégrés.
- Les anciens formulaires ne bénéficient pas encore de la protection CSRF.

## Correctif de finalisation — verrouillage des consultations
- Une réponse enregistrée devient immédiatement non modifiable (y compris après retour à l'espace personnel).
- Déverrouillage ponctuel réservé au rôle ADMIN, avec CSRF et journal d'audit.
- Après nouvelle validation, la réponse est de nouveau verrouillée.
- Migration additive `modification_autorisee` : les réponses existantes sont conservées et verrouillées par défaut ; aucune réécriture des choix/commentaires.
- Test de non-régression sur le cycle réponse → refus → autorisation → modification → nouveau verrouillage.
