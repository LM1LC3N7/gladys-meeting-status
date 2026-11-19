// -----------------------------------------------------------------------------
// The SDK handlers, wired to the app. Shared by index.js and the end-to-end
// tests, so the tests run the exact wiring of the container.
// -----------------------------------------------------------------------------

import { ACTIONS } from './constants.js';

export function registerHandlers(gladys, app, logger) {
  // Discovery: the device of the integration (never a read).
  gladys.onScanRequest(() => app.scan());
  gladys.onDeviceCreated((device) => app.deviceCreated(device));
  gladys.onDeviceUpdated((device) => app.deviceUpdated(device));
  gladys.onDeviceDeleted((device) => app.deviceDeleted(device));

  // OAuth2: the two Connect buttons of the configuration, then the return to Gladys.
  gladys.onOAuthAuthorizeUrl((key, redirectUri) => app.oauthAuthorizeUrl(key, redirectUri));
  gladys.onOAuthCallback((key, params) => app.oauthCallback(key, params));

  // Configuration screen buttons: the pasted return address, the test.
  gladys.onAction(ACTIONS.FINISH_SIGN_IN, (fields) => app.finishSignIn(fields));
  gladys.onAction(ACTIONS.TEST_CONNECTION, () => app.testConnection());

  gladys.onConfigUpdated((config) => app.configUpdated(config));

  // The stop waits for a renewal in flight and its save (8 s at most): a renewal token serves once.
  gladys.handleShutdown(() => app.shutdown());

  // Connection lifecycle: the first connection starts, a reconnection resends.
  gladys.on('connected', () => {
    app.onConnected(() => gladys.getConfig()).catch((err) => logger?.error('Start failed', err));
  });
}
