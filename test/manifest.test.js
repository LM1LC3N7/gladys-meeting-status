// -----------------------------------------------------------------------------
// The manifest is the contract with Gladys and the store: these tests tie it
// to the code, so a key renamed on one side only fails here.
// -----------------------------------------------------------------------------

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { createFakeGladys } from '@gladysassistant/integration-sdk/testing';
import { DEFAULTS, LANGUAGES } from '../src/config.js';
import { ACTIONS, OAUTH_FIELDS, POLL_SECONDS, RULES } from '../src/constants.js';
import { buildDevice } from '../src/devices.js';

const read = (path) => readFileSync(new URL(`../${path}`, import.meta.url), 'utf8');
const manifest = JSON.parse(read('gladys-assistant-integration.json'));
const handlers = read('src/handlers.js');

function* texts(node, path = '') {
  if (Array.isArray(node)) {
    for (const [i, entry] of node.entries()) {
      yield* texts(entry, `${path}[${i}]`);
    }
  } else if (node && typeof node === 'object') {
    if ('en' in node || 'fr' in node) {
      yield [path, node];
      return;
    }
    for (const [key, value] of Object.entries(node)) {
      yield* texts(value, `${path}.${key}`);
    }
  }
}

test('the store basics; the version is the same everywhere and never pinned in a test', () => {
  assert.equal(manifest.manifest_version, 1);
  assert.equal(manifest.type, 'device');
  assert.match(manifest.version, /^\d+\.\d+\.\d+$/);
  assert.equal(manifest.docker_image, `ghcr.io/lm1lc3n7/gladys-meeting-status:${manifest.version}`);
  assert.equal(manifest.version, JSON.parse(read('package.json')).version);
  const lock = JSON.parse(read('package-lock.json'));
  assert.equal(lock.version, manifest.version);
  assert.equal(lock.packages[''].version, manifest.version);
  assert.deepEqual(manifest.transports, ['cloud']);
  assert.equal(manifest.gladys_version, '>=5.1.0');
  assert.equal(manifest.countries, undefined, 'the whole world');
  assert.deepEqual(manifest.categories, ['services']);
  assert.equal(manifest.keywords, undefined);
  assert.equal(
    manifest.cover_image,
    'https://raw.githubusercontent.com/LM1LC3N7/gladys-meeting-status/main/cover.png',
  );
  for (const lang of ['en', 'fr']) {
    const length = manifest.description[lang].length;
    assert.ok(length >= 10 && length <= 100, `description.${lang}: ${length} characters`);
    const docs = read(`docs/${lang}.md`);
    assert.ok(docs.length >= 300, `docs/${lang}.md is too short for the store`);
    assert.match(docs, /not[\s>]+affiliated[\s>]+with|sans[\s>]+lien[\s>]+avec/i);
    assert.match(
      docs,
      /never[\s>]+tested[\s>]+with[\s>]+a[\s>]+real[\s>]+device|jamais[\s>]+testée?[\s>]+avec[\s>]+un[\s>]+vrai[\s>]+appareil/i,
    );
    assert.match(docs, /personal|personnel/i, 'a personal Microsoft account is not covered');
    assert.match(docs, /Google Meet/);
  }
  for (const file of [
    'README.md',
    'CHANGELOG.md',
    'TODO.md',
    'CLAUDE.md',
    'docs/en.md',
    'docs/fr.md',
  ]) {
    assert.doesNotMatch(read(file), /chiffr|encrypt|hachage/i, file);
  }
  for (const [path, text] of texts(manifest)) {
    for (const lang of ['en', 'fr']) {
      assert.doesNotMatch(text[lang] ?? '', /chiffr|encrypt|hachage/i, path);
    }
  }
  assert.match(read('LICENSE'), /Apache License\s+Version 2\.0/);
  assert.doesNotMatch(read('LICENSE'), /Copyright \d{4} [A-Z][a-z]+ [A-Z]/, 'no name in LICENSE');
  assert.match(read('Dockerfile'), /COPY LICENSE THIRD_PARTY_NOTICES\.md/);
});

test('every text is in English and French', () => {
  for (const [path, text] of texts(manifest)) {
    for (const lang of ['en', 'fr']) {
      assert.ok(text[lang]?.trim(), `${path} needs ${lang}`);
    }
  }
});

test('the configuration matches the code; secrets are secret fields without default', () => {
  const fields = Object.fromEntries(manifest.config_schema.map((f) => [f.key, f]));
  for (const [key, value] of Object.entries(DEFAULTS)) {
    assert.ok(fields[key], `${key} is not in the config_schema`);
    assert.deepEqual(fields[key].default, value, key);
  }
  assert.deepEqual(
    fields.language.options.map((o) => o.value),
    LANGUAGES,
  );
  assert.deepEqual(
    fields.rule.options.map((o) => o.value),
    RULES,
  );
  for (const [provider, key] of Object.entries(OAUTH_FIELDS)) {
    assert.equal(fields[key].type, 'oauth2', key);
    assert.equal(fields[`${provider}_client_id`].type, 'string');
    const secret = fields[`${provider}_client_secret`];
    assert.equal(secret.type, 'secret', 'the secret is never sent back to the browser');
    assert.equal(secret.default, undefined);
    assert.match(
      secret.description.en,
      /Kept by Gladys in its database, never shown again, sent only to/,
    );
    assert.match(
      secret.description.fr,
      /Gardé par Gladys dans sa base, jamais réaffiché, envoyé seulement à/,
    );
  }
  for (const field of manifest.config_schema) {
    assert.ok(!field.key.startsWith('gladys_'), `${field.key}: a gladys_ key is never received`);
    assert.notEqual(field.type, 'password', 'no password: it is typed on the service page only');
    if (field.placeholder) {
      assert.equal(typeof field.placeholder, 'object', `${field.key}: placeholder per language`);
    }
    for (const option of field.options ?? []) {
      for (const lang of ['en', 'fr']) {
        assert.equal(typeof option.label[lang], 'string');
      }
    }
  }
  assert.deepEqual(
    [fields.poll_seconds.min, fields.poll_seconds.max],
    [POLL_SECONDS.min, POLL_SECONDS.max],
  );
  assert.ok(POLL_SECONDS.min >= 60, 'one read a minute at most');
  for (const field of manifest.config_schema.filter((f) => f.type === 'section')) {
    for (const lang of ['en', 'fr']) {
      assert.ok(field.description[lang].length <= 1000, `${field.key}.${lang}`);
    }
  }
  assert.ok(fields.intro.links.every((l) => l.url.startsWith('https://')));
});

test('the actions and their forms are those of the code; read only, no command', () => {
  assert.deepEqual(manifest.actions.map((a) => a.key).sort(), Object.values(ACTIONS).sort());
  for (const name of Object.keys(ACTIONS)) {
    assert.ok(handlers.includes(`onAction(ACTIONS.${name}`), `ACTIONS.${name} has no handler`);
  }
  assert.ok(handlers.includes('onOAuthAuthorizeUrl') && handlers.includes('onOAuthCallback'));
  assert.ok(!handlers.includes('onSetValue'), 'read only');
  for (const action of manifest.actions) {
    assert.ok(action.timeout_seconds >= 5 && action.timeout_seconds <= 120);
  }
  assert.equal(manifest.scene_actions, undefined);
});

test('the published device: read only, no switch, every supported_options valid for the core', () => {
  const gladys = createFakeGladys();
  for (const language of LANGUAGES) {
    const device = buildDevice(gladys, language);
    for (const f of device.features) {
      assert.equal(f.read_only, true);
      assert.ok(!['switch', 'light', 'lock'].includes(f.category));
      assert.ok(f.name.length > 0 && f.name.length <= 100);
      // The core rules (normalizeSupportedOptions): a label is a non-empty trimmed string,
      // a value an integer (or a string on text/select only), unique, no other key.
      const options = f.supported_options ?? [];
      const values = new Set();
      for (const o of options) {
        assert.equal(typeof o.label, 'string');
        assert.equal(o.label, o.label.trim());
        assert.ok(o.label.length > 0);
        assert.ok(
          Number.isInteger(o.value) || (f.type === 'select' && typeof o.value === 'string'),
        );
        assert.ok(!values.has(o.value));
        values.add(o.value);
        assert.deepEqual(Object.keys(o).sort(), ['label', 'value']);
      }
    }
  }
});
