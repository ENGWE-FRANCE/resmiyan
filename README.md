# ResMiyan — outil de réservation MIYAN

## Architecture

```
reservation.html  ──┐
                     ├──► Google Apps Script (Web App, /exec)  ──► Google Sheet
admin.html        ──┘        Code.gs : doGet / doPost               "Reservations" + "Config"
                              - calcul des créneaux disponibles
                              - verrou (LockService) anti double-booking
                              - email de confirmation + lien d'annulation
```

- **Google Sheet** = base de données. Deux onglets :
  - `Reservations` : une ligne par réservation (id, date, heure, service, couverts, contact, statut, token d'annulation).
  - `Config` : paramètres clé/valeur (jauge max par créneau, horaires des services, jours fermés, clé admin...). Modifiable sans toucher au code.
- **Apps Script** ([apps-script/Code.gs](apps-script/Code.gs)) = backend unique, exposé en Web App. Il sert les deux pages HTML et c'est lui qui décide si un créneau est encore disponible (avec un verrou pour éviter que deux réservations simultanées dépassent la jauge).
- **reservation.html** = page client (déjà prototypée par Julien), inchangée : elle appelle l'API pour lister les créneaux puis poster la réservation.
- **admin.html** = tableau de bord interne qui se rafraîchit automatiquement (toutes les 8 s) et met en évidence les nouvelles réservations — c'est la "page qui se met à jour" demandée dans l'échange.

## Mise en place (une fois)

1. Créer un Google Sheet dédié, ex. "MIYAN — Réservations".
2. Dans ce Sheet : `Extensions > Apps Script`. Copier le contenu de [apps-script/Code.gs](apps-script/Code.gs) et de [apps-script/appsscript.json](apps-script/appsscript.json) (via l'éditeur de manifeste).
3. Dans l'éditeur Apps Script, sélectionner la fonction `setupSpreadsheet` puis **Exécuter** (autoriser les accès demandés). Cela crée les onglets `Reservations` et `Config` avec des valeurs par défaut, dont une `adminKey` générée aléatoirement.
4. Ouvrir l'onglet `Config` et ajuster : `email`, `phone`, `capacityPerSlot` (couverts max par créneau de 15 min), `servicesJson` (horaires midi/soir), `closedWeekdays` (jours fermés), et noter la valeur d'`adminKey`.
5. `Déployer > Nouveau déploiement > Application Web` : exécuter en tant que **Moi**, accès **Tout le monde**. Copier l'URL `.../exec`.
6. Coller cette URL dans `const API_URL = '...'` en haut du `<script>` de [reservation.html](reservation.html) **et** de [admin.html](admin.html) — c'est cette étape qui fait disparaître le bandeau "Mode démonstration" (il ne disparaît que si l'URL commence bien par `https://script.google.com/`).
7. Ouvrir `admin.html`, entrer l'`adminKey` notée à l'étape 4 — la clé reste enregistrée sur l'appareil.
8. Tester : réserver depuis `reservation.html`, vérifier que la ligne apparaît dans l'onglet `Reservations` du Sheet **et** dans `admin.html` (rafraîchissement sous 8 s), et que l'email de confirmation arrive.

Tant que l'étape 6 n'est pas faite, les deux pages tournent en local avec de fausses données (`DEMO`/`COLLER_ICI_URL_APPS_SCRIPT`) — c'est normal, c'est ce que vous voyez actuellement.

## Jauge de couverts par créneau

La jauge limite le nombre total de couverts acceptés sur un même créneau de 15 min (`slotMinutes`), service par service (midi / soir). Deux niveaux, dans l'onglet `Config` :

- **Globale** : `capacityPerSlot` (ex. `20`) s'applique à tous les services qui n'ont pas de jauge propre.
- **Par service** (plus précis, recommandé si midi et soir n'ont pas la même capacité) : ajouter `"capacity": N` à chaque service dans `servicesJson`, par ex. :

  ```json
  [
    {"id":"midi","label":"Midi","start":"12:00","end":"14:00","capacity":16},
    {"id":"soir","label":"Soir","start":"19:00","end":"22:00","capacity":28}
  ]
  ```

  Ici midi est limité à 16 couverts par créneau de 15 min, soir à 28, sans toucher au code. Dès que le contact donne le nombre réel de places, c'est cette valeur (`capacity` par service, ou `capacityPerSlot` si une seule jauge suffit) qu'il faut mettre à jour — directement dans la feuille `Config`, aucun redéploiement n'est nécessaire.
  Le calcul : pour un créneau donné, Apps Script additionne les couverts déjà confirmés à cette heure+service, et refuse toute nouvelle réservation qui ferait dépasser cette jauge (vérifié une première fois à l'affichage des créneaux, une seconde fois — sous verrou — au moment d'enregistrer, pour éviter que deux clients ne réservent le dernier covert au même instant).

## Intégration Google / site (étape suivante, côté "my tam")

- `reservation.html` est une page autonome : elle peut être intégrée au site existant via un `<iframe>`, un bouton "Réserver" qui pointe vers son URL d'hébergement, ou copiée dans un composant HTML du CMS.
- Le Google Sheet reste consultable directement par l'équipe si besoin (lecture simple), en complément du tableau de bord `admin.html`.
- Aucune autre intégration Google n'est requise pour le moment ; un export/sync vers Google Calendar est possible ensuite si utile, en ajoutant un appel `CalendarApp` dans `createReservation` côté `Code.gs`.

## Hébergement : GitHub Pages suffit (Netlify optionnel)

`reservation.html` et `admin.html` sont des pages 100 % statiques : tout le code tourne dans le navigateur et appelle directement l'URL Apps Script en `fetch`. Il n'y a pas de serveur à faire tourner, donc pas besoin de Netlify.

- **GitHub Pages** (gratuit, déjà inclus avec le repo) : `Settings > Pages > Deploy from a branch` → `main` / `/ (root)`. Les pages sont alors servies à `https://<compte>.github.io/<repo>/reservation.html` et `.../admin.html`. C'est suffisant pour tout le monde (test, intégration en iframe sur le site, usage interne pour l'admin).
- **Netlify** reste une option si vous voulez un nom de domaine personnalisé plus simple à configurer ou un déploiement auto à chaque push — mais ce n'est pas requis pour que l'outil fonctionne.

Dans les deux cas, le Google Sheet + Apps Script restent la seule "base de données" et le seul backend : GitHub (ou Netlify) ne fait qu'héberger les deux fichiers HTML, il n'est jamais "connecté" au Sheet lui-même — c'est `API_URL` dans le HTML qui fait le lien.

## Sécurité / anti-abus

- Champ piège (`website`) déjà présent côté formulaire client : une soumission qui le remplit est silencieusement ignorée (bot).
- `admin.html` exige la clé stockée dans `Config.adminKey` — à ne partager qu'avec l'équipe.
- Double réservation sur un créneau plein : vérifiée deux fois (affichage + au moment d'écrire, sous verrou) pour éviter une course entre deux clients simultanés.
