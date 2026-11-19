# Statut en réunion (Microsoft Teams et Zoom)

Gladys lit, environ une fois par minute, si vous êtes en réunion ou en appel, dans **Microsoft
Teams** (par Microsoft Graph) et dans **Zoom**, pour qu’une scène allume un voyant devant votre
bureau. Lecture seule : rien n’est modifié sur vos comptes.

Intégration non officielle, sans lien avec Microsoft ni Zoom. Jamais testée avec un vrai appareil
ni un vrai compte : testée sur des réponses calquées sur les documentations officielles.

## Ce que vous obtenez

Un appareil Gladys, **Statut en réunion** (onglet Découverte, dès qu’un service est autorisé), en
lecture seule :

| Fonction                   | Valeur                                                                            |
| -------------------------- | --------------------------------------------------------------------------------- |
| En réunion (Teams ou Zoom) | 1 quand l’un des services le dit ; 0 quand tous les services autorisés disent non |
| En réunion Teams           | 1 / 0                                                                             |
| Présence Teams             | les codes de Microsoft Graph, `availability / activity` (`Busy / InAMeeting`)     |
| En réunion Zoom            | 1 / 0                                                                             |
| Présence Zoom              | le code de Zoom (`In_A_Zoom_Meeting`, `Do_Not_Disturb`…)                          |

« En réunion » : une réunion, un appel ou une présentation (Teams `InAMeeting`, `InACall`,
`InAConferenceCall`, `Presenting` ; Zoom `In_A_Zoom_Meeting`, `On_A_Call`, `Presenting`…). Avec
l’option **Aussi occupé et ne pas déranger**, Occupé, Ne pas déranger et les événements d’agenda
Zoom comptent aussi.

Aucune valeur n’est inventée : une présence absente de la documentation (`PresenceUnknown`, une
nouvelle valeur), ou un service qui n’a pas pu être lu depuis plus de 3 intervalles (3 minutes au
moins), est **inconnue** et rien n’est publié. Gladys n’a pas de valeur « inconnue » : **la dernière
valeur reste affichée** ; lisez l’état de l’intégration (ou les textes « Présence Teams » /
« Présence Zoom » et leur date) avant de vous fier au voyant. « En réunion (Teams ou Zoom) » ne
passe à 0 que si tous les services autorisés ont été lus récemment et disent non : un service dont
la session a pris fin (renouvellement refusé) compte comme inconnu jusqu’à un nouveau Connecter ou
jusqu’à ce que ses clés soient vidées.

Non couverts : les **comptes Microsoft personnels** (Microsoft Graph n’a pas de présence pour eux :
comptes professionnels ou scolaires seulement), Microsoft Teams gratuit, **Google Meet** (pas d’API
de présence), la présence d’autres personnes.

## 1. Microsoft Teams (compte professionnel ou scolaire)

1. Ouvrez le [centre d’administration Microsoft Entra](https://entra.microsoft.com/),
   **Inscriptions d’applications**, **Nouvelle inscription** : comptes d’un annuaire
   d’organisation.
2. Plateforme **Web**, URI de redirection : cliquez une fois sur **Connecter** sur le compte
   Microsoft de la configuration de l’intégration ; l’état en haut de la page affiche alors
   l’adresse à déclarer (en général `https://my.gladysassistant.com/redirect/oauth`). Ou donnez une
   adresse à vous et remplissez **URL de redirection** (voir « Votre propre URL de redirection »).
3. **Autorisations d’API** : Microsoft Graph, **Déléguées** : **Presence.Read** et
   **offline_access**. Votre organisation peut exiger le consentement d’un administrateur.
4. **Certificats et secrets** : un nouveau secret client. Notez sa date de fin : à cette date, le
   renouvellement est refusé et il faut en créer un nouveau, l’enregistrer et cliquer de nouveau
   sur Connecter.
5. Copiez l’**ID d’application (client)** et la **valeur** du secret dans la configuration.
   Tenant : laissez `organizations`, ou donnez l’ID ou le domaine de votre tenant (`contoso.com`).
   `common` et `consumers` sont refusés (les comptes personnels n’ont pas de présence).
6. Enregistrez, cliquez sur **Connecter** et connectez-vous sur la page de Microsoft.

## 2. Zoom

1. Sur le [Zoom App Marketplace](https://marketplace.zoom.us/), **Develop**, **Build App** : une
   **General app**, **gérée par l’utilisateur** (user-managed).
2. **OAuth redirect URL** et **allow list** : l’adresse affichée dans l’état après un premier clic
   sur Connecter (en général `https://my.gladysassistant.com/redirect/oauth`), ou la vôtre.
3. **Scopes** : `user:read:presence_status` (sur une ancienne application à scopes classiques :
   `user:read`), rien d’autre.
4. Copiez le **Client ID** et le **Client Secret** dans la configuration, enregistrez, cliquez sur
   **Connecter**, autorisez l’application sur la page de Zoom.

Ne remplissez que les services que vous utilisez : un service vide n’est jamais lu.

## Votre propre URL de redirection

Si l’URL de redirection de vos applications n’est pas celle de Gladys, saisissez-la dans **URL de
redirection**. Après la page de Microsoft ou de Zoom, copiez l’adresse complète sur laquelle le
navigateur s’est arrêté (avec `code=` et `state=`) et collez-la dans **Terminer l’autorisation**,
dans les 15 minutes. L’adresse doit être l’URL de redirection donnée ; une autorisation ne sert
qu’une fois.

## Lectures, pauses et arrêts

- Une lecture par service et par intervalle (60 s par défaut, une minute au moins), quoi qui la
  lance : bouton Tester, configuration enregistrée, reconnexion de Gladys, redémarrage (l’heure de
  la dernière lecture est gardée dans `/data`).
- **429** (trop de requêtes) : pause aussi longue que le service le demande (`Retry-After`, entre
  1 minute et 6 heures ; 5 minutes sans lui). **5xx** : nouvel essai après 2, 5, 15 puis
  30 minutes. Autres erreurs : jamais plus tôt que l’intervalle, puis 2, 5, 15, 30 minutes. Les
  pauses survivent à un redémarrage.
- **403** sur la présence (ou code Zoom 5411) : la permission ou le type de compte est refusé ;
  nouvel essai après une heure, ou dès un nouveau Connecter.
- Au plus 2 000 requêtes par service et par jour UTC (un filet de sécurité ; le rythme normal est
  d’environ 1 500).
- **Renouvellement refusé** (secret arrivé à sa fin, accès révoqué, jeton de renouvellement
  expiré) : le service s’arrête jusqu’à un nouveau clic sur Connecter ; rien n’est réessayé de
  lui-même, y compris après un redémarrage.
- **Accès refusé avec un nouveau jeton** : le service s’arrête ; **Tester** essaie une fois (pas
  plus tôt que l’intervalle), ou cliquez de nouveau sur Connecter.
- Une réponse d’erreur sans code d’erreur OAuth (une page de proxy, un pare-feu) n’est jamais prise
  pour un refus : elle est réessayée plus tard.

## Vie privée et secrets

- Votre mot de passe se saisit sur la page de Microsoft ou de Zoom seulement.
- Les secrets client sont gardés par Gladys dans sa base, jamais réaffichés, envoyés seulement au
  service qui les a émis.
- Les jetons d’accès et de renouvellement sont gardés par Gladys dans sa base (dans la
  configuration de l’intégration, jamais affichés à l’écran), envoyés seulement au service qui les
  a émis, absents de `/data`, des journaux et des messages d’état. `/data` ne garde que des dates,
  des pauses et des arrêts. Une version qui les gardait dans `/data` les déplace vers Gladys au
  démarrage, puis les retire de `/data`. Les jetons de renouvellement de Zoom et de Microsoft ne
  servent qu’une fois : un renouvellement ne part que si Gladys est connecté et vient d’accepter
  une écriture des jetons actuels ; sinon le jeton actuel reste et le renouvellement est retenté
  une minute plus tard (les journaux disent « Gladys cannot keep the tokens »). Si Gladys refuse
  le nouveau jeton juste après (les journaux disent « tokens not saved in Gladys »), il est
  enregistré de nouveau toutes les 10 secondes à une minute et à chaque reconnexion à Gladys,
  sans autre renouvellement avant. Un arrêt du conteneur attend jusqu’à 8 secondes un
  renouvellement en cours et son enregistrement, et tente une dernière fois d’enregistrer un jeton
  que Gladys avait refusé ; un Connecter pendant l’arrêt est refusé (recliquez une fois
  l’intégration relancée). Seuls un plantage ou un arrêt forcé du conteneur dans les secondes
  mêmes d’un renouvellement, ou un arrêt pendant que Gladys refuse encore le nouveau jeton,
  peuvent le perdre : l’état dit alors que la session a pris fin ; cliquez sur Connecter.
- La présence est gardée en mémoire seulement, ni dans `/data` ni dans les journaux. « En réunion »
  garde son historique dans Gladys (désactivable par fonction) ; les textes de présence n’en ont
  pas.
- L’autorisation utilise PKCE (S256) et un `state` qui ne sert qu’une fois, pendant 15 minutes.

## Dépannage

- **AADSTS50020 / compte personnel** : utilisez un compte professionnel ou scolaire ; les comptes
  personnels n’ont pas de présence.
- **AADSTS65001 / consentement** : un administrateur doit consentir à Presence.Read pour votre
  organisation.
- **invalid_client** : le secret client est faux ou arrivé à sa fin.
- **Zoom 5411** : il manque le scope `user:read:presence_status` à l’application.
- La documentation de Zoom signale qu’un utilisateur connecté sur mobile seulement peut apparaître
  `Offline`.
