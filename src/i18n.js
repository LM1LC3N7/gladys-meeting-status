// -----------------------------------------------------------------------------
// Texts shown in Gladys (connection status, buttons, device and feature
// names), in English and French. No message ever carries a token, a code or
// a secret.
// -----------------------------------------------------------------------------

/** The text of `texts` in `language`, English otherwise. */
export const word = (texts, language) => texts[language] ?? texts.en;

const both = (en, fr) => ({ en, fr });

/** Several messages one after the other, in each language. */
export const join = (...messages) => {
  const list = messages.filter(Boolean);
  return both(list.map((m) => m.en).join(' '), list.map((m) => m.fr).join(' '));
};

/** "45 s" / "25 min" / "3 h", for the pauses. */
export function age(ms) {
  const seconds = Math.max(1, Math.round(ms / 1000));
  if (seconds < 90) {
    return both(`${seconds} s`, `${seconds} s`);
  }
  const minutes = Math.round(seconds / 60);
  if (minutes < 120) {
    return both(`${minutes} min`, `${minutes} min`);
  }
  const hours = Math.round(minutes / 60);
  return both(`${hours} h`, `${hours} h`);
}

export const SERVICE_NAMES = {
  microsoft: both('Microsoft Teams', 'Microsoft Teams'),
  zoom: both('Zoom', 'Zoom'),
};

/** "Zoom: <message>" in each language. */
export const prefixed = (provider, message) =>
  both(
    `${SERVICE_NAMES[provider].en}: ${message.en}`,
    `${SERVICE_NAMES[provider].fr} : ${message.fr}`,
  );

export const DEVICE_NAME = both('Meeting status', 'Statut en réunion');

export const FEATURE_NAMES = {
  in_meeting: both('In a meeting (Teams or Zoom)', 'En réunion (Teams ou Zoom)'),
  microsoft_in_meeting: both('In a Teams meeting', 'En réunion Teams'),
  microsoft_presence: both('Teams presence', 'Présence Teams'),
  zoom_in_meeting: both('In a Zoom meeting', 'En réunion Zoom'),
  zoom_presence: both('Zoom presence', 'Présence Zoom'),
};

export const T = {
  nothing: both(
    'Enter the client id and the client secret of your Microsoft (Entra ID) or Zoom app (see the guide), save, then click Connect.',
    'Saisissez l’identifiant client (client id) et le secret client (client secret) de votre application Microsoft (Entra ID) ou Zoom (voir le guide), enregistrez, puis cliquez sur Connecter.',
  ),
  incomplete: both(
    'the client id, the client secret (or the tenant) is missing or invalid: complete them and save.',
    'l’identifiant client, le secret client (ou le tenant) manque ou est invalide : complétez-les et enregistrez.',
  ),
  signedOut: both(
    'not signed in: click Connect. The integration never signs in by itself.',
    'non connecté : cliquez sur Connecter. L’intégration ne se connecte jamais d’elle-même.',
  ),
  expired: both(
    'the session ended (renewal refused): click Connect to authorize again. Nothing is retried by itself.',
    'la session a pris fin (renouvellement refusé) : cliquez sur Connecter pour autoriser de nouveau. Rien n’est réessayé de lui-même.',
  ),
  unauthorized: both(
    'the access was refused even with a new token: click Connect to authorize again (or "Test" once). Nothing is retried by itself.',
    'l’accès a été refusé même avec un nouveau jeton : cliquez sur Connecter pour autoriser de nouveau (ou une fois sur « Tester »). Rien n’est réessayé de lui-même.',
  ),
  connecting: both('reading the presence…', 'lecture de la présence…'),
  connected: both('connected.', 'connecté.'),
  unknownPresence: both(
    'connected, presence unknown (not listed in the documentation).',
    'connecté, présence inconnue (absente de la documentation).',
  ),
  nextRead: (wait) => both(`Next read in ${wait.en}.`, `Prochaine lecture dans ${wait.fr}.`),
  rateLimited: (wait) =>
    both(
      `too many requests: paused for ${wait.en}.`,
      `trop de requêtes : en pause pendant ${wait.fr}.`,
    ),
  unavailable: (wait) =>
    both(
      `the service is unavailable: next try in ${wait.en}.`,
      `le service est indisponible : nouvel essai dans ${wait.fr}.`,
    ),
  budget: (wait) =>
    both(
      `daily number of requests reached: paused for ${wait.en}.`,
      `nombre de requêtes du jour atteint : en pause pendant ${wait.fr}.`,
    ),
  forbidden: (wait) =>
    both(
      `the presence was refused (work or school account and consent to the permission needed, see the guide): next try in ${wait.en}, or click Connect again.`,
      `la présence a été refusée (compte professionnel ou scolaire et consentement à la permission nécessaires, voir le guide) : nouvel essai dans ${wait.fr}, ou cliquez de nouveau sur Connecter.`,
    ),
  unreachable: (reason, wait) =>
    both(
      `not read (${reason}): next try in ${wait.en}.`,
      `non lu (${reason}) : nouvel essai dans ${wait.fr}.`,
    ),
  renewWaiting: (reason, wait) =>
    both(
      `the token could not be renewed (${reason}): next try in ${wait.en}.`,
      `le jeton n’a pas pu être renouvelé (${reason}) : nouvel essai dans ${wait.fr}.`,
    ),
  storeUnreadable: both(
    'The state file of /data cannot be read: read again every 10 s. Nothing is sent meanwhile.',
    'Le fichier d’état de /data est illisible : relu toutes les 10 s. Rien n’est envoyé en attendant.',
  ),
  configUnreadable: (wait) =>
    both(
      `The configuration could not be read from Gladys: next try in ${wait.en}.`,
      `La configuration n’a pas pu être lue depuis Gladys : nouvel essai dans ${wait.fr}.`,
    ),
  starting: both(
    'Starting, try again in a few seconds.',
    'Démarrage, réessayez dans quelques secondes.',
  ),
  busy: both(
    'A read is running, try again in a few seconds.',
    'Une lecture est en cours, réessayez dans quelques secondes.',
  ),
  test: {
    cooldown: both('Tested less than 30 s ago.', 'Testé il y a moins de 30 s.'),
  },
  signIn: {
    noClient: both(
      'Save the client id and the client secret of this app first (and the tenant for Microsoft).',
      'Enregistrez d’abord l’identifiant et le secret client de cette application (et le tenant pour Microsoft).',
    ),
    unknownField: both('Unknown Connect button.', 'Bouton Connecter inconnu.'),
    noRedirect: both(
      'No return address: fill in "Redirect URI" and save.',
      'Pas d’adresse de retour : remplissez « URL de redirection » et enregistrez.',
    ),
    redirectHint: (address) =>
      both(
        `Redirect URI to register in your app: ${address}`,
        `URL de redirection à déclarer dans votre application : ${address}`,
      ),
    noCode: both(
      'No authorization code in this address: paste the whole address the browser ended on (with code= and state=).',
      'Pas de code d’autorisation dans cette adresse : collez l’adresse complète affichée par le navigateur (avec code= et state=).',
    ),
    unknown: both(
      'This authorization was not started here, was already used or is older than 15 minutes: click Connect again.',
      'Cette autorisation n’a pas été lancée ici, a déjà servi ou date de plus de 15 minutes : cliquez de nouveau sur Connecter.',
    ),
    wrongAddress: both(
      'This address is not the redirect URI of the authorization: paste the address the browser ended on.',
      'Cette adresse n’est pas l’URL de redirection de l’autorisation : collez l’adresse sur laquelle le navigateur s’est arrêté.',
    ),
    cancelled: (code) =>
      both(
        `The authorization was not given (${code}).`,
        `L’autorisation n’a pas été donnée (${code}).`,
      ),
    refused: (code) =>
      both(
        `The service refused the authorization (${code}): check the client id, the client secret (and its end date), the redirect URI, then click Connect again.`,
        `Le service a refusé l’autorisation (${code}) : vérifiez l’identifiant client, le secret client (et sa date de fin), l’URL de redirection, puis cliquez de nouveau sur Connecter.`,
      ),
    slow: both(
      'The service did not answer in time: click Connect again.',
      'Le service n’a pas répondu à temps : cliquez de nouveau sur Connecter.',
    ),
    retryConnect: (reason) =>
      both(
        `The authorization failed (${reason}): click Connect again.`,
        `L’autorisation a échoué (${reason}) : cliquez de nouveau sur Connecter.`,
      ),
    failed: (reason) =>
      both(
        `The authorization failed (${reason}): paste the same address again in a moment.`,
        `L’autorisation a échoué (${reason}) : recollez la même adresse dans un moment.`,
      ),
    noRefresh: both(
      'The service gave no renewal token (Microsoft: the offline_access permission is needed): click Connect again.',
      'Le service n’a pas donné de jeton de renouvellement (Microsoft : la permission offline_access est nécessaire) : cliquez de nouveau sur Connecter.',
    ),
    done: both(
      'Authorized: the presence is being read.',
      'Autorisé : la présence est en cours de lecture.',
    ),
    doneLater: (wait) =>
      both(
        `Authorized: the presence is read in ${wait.en}.`,
        `Autorisé : la présence sera lue dans ${wait.fr}.`,
      ),
  },
};
