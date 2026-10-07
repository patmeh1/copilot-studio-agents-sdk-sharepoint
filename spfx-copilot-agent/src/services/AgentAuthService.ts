import {
  AccountInfo,
  AuthenticationResult,
  InteractionRequiredAuthError,
  LogLevel,
  PublicClientApplication
} from '@azure/msal-browser';

export interface IAgentAuthConfig {
  clientId: string;
  tenantId: string;
  /** UPN of the SharePoint user; used to pick the right cached account and as login_hint. */
  loginHint?: string;
}

/** Thrown when a token cannot be obtained without user interaction (popup). */
export class AuthInteractionRequiredError extends Error {
  public readonly innerError: unknown;

  constructor(innerError: unknown) {
    super('Sign-in required');
    this.name = 'AuthInteractionRequiredError';
    this.innerError = innerError;
    Object.setPrototypeOf(this, AuthInteractionRequiredError.prototype);
  }
}

const clients: { [key: string]: Promise<PublicClientApplication> } = {};
const queues: { [key: string]: Promise<unknown> } = {};

function keyOf(config: IAgentAuthConfig): string {
  return `${config.clientId}|${config.tenantId}`.toLowerCase();
}

/**
 * One MSAL instance per (client, tenant). The redirect URI is the SharePoint origin
 * (https://<tenant>.sharepoint.com), so a single SPA redirect URI covers every site in the tenant.
 */
function getClient(config: IAgentAuthConfig): Promise<PublicClientApplication> {
  const key = keyOf(config);
  if (!clients[key]) {
    const pca = new PublicClientApplication({
      auth: {
        clientId: config.clientId,
        authority: `https://login.microsoftonline.com/${config.tenantId}`,
        redirectUri: window.location.origin,
        navigateToLoginRequestUrl: false
      },
      cache: {
        cacheLocation: 'localStorage'
      },
      system: {
        iframeHashTimeout: 10000,
        loggerOptions: {
          logLevel: LogLevel.Warning,
          piiLoggingEnabled: false,
          loggerCallback: (level: LogLevel, message: string, containsPii: boolean): void => {
            if (containsPii) { return; }
            if (level === LogLevel.Error) { console.error(`[CopilotAgent/MSAL] ${message}`); }
            else if (level === LogLevel.Warning) { console.warn(`[CopilotAgent/MSAL] ${message}`); }
          }
        }
      }
    });
    clients[key] = pca.initialize().then(() => pca);
  }
  return clients[key];
}

/** Serialises token requests per MSAL instance (several web parts can live on one page). */
function enqueue<T>(config: IAgentAuthConfig, work: () => Promise<T>): Promise<T> {
  const key = keyOf(config);
  const previous = queues[key] || Promise.resolve();
  const next = previous.then(work, work);
  queues[key] = next.catch(() => undefined);
  return next;
}

function findAccount(pca: PublicClientApplication, loginHint?: string): AccountInfo | undefined {
  const accounts = pca.getAllAccounts();
  if (loginHint) {
    const hint = loginHint.toLowerCase();
    // Never fall back to a different user's cached account.
    return accounts.filter(a => (a.username || '').toLowerCase() === hint)[0];
  }
  return accounts.length === 1 ? accounts[0] : undefined;
}

/**
 * Silent token acquisition: MSAL cache/refresh token first, then ssoSilent (hidden iframe that
 * re-uses the user's existing Entra ID session, i.e. true SSO with the SharePoint sign-in).
 */
export function acquireTokenSilently(config: IAgentAuthConfig, scopes: string[]): Promise<AuthenticationResult> {
  return enqueue(config, async () => {
    const pca = await getClient(config);
    const account = findAccount(pca, config.loginHint);
    if (account) {
      try {
        return await pca.acquireTokenSilent({ scopes, account });
      } catch (e) {
        console.warn('[CopilotAgent] acquireTokenSilent failed, trying ssoSilent', e);
      }
    }
    try {
      const result = await pca.ssoSilent({ scopes, loginHint: config.loginHint });
      pca.setActiveAccount(result.account);
      return result;
    } catch (e) {
      // Another web part on the page may have completed an interactive sign-in while this attempt was running.
      const signedIn = account ? undefined : findAccount(pca, config.loginHint);
      if (signedIn) {
        try {
          return await pca.acquireTokenSilent({ scopes, account: signedIn });
        } catch (retryError) {
          console.warn('[CopilotAgent] acquireTokenSilent with the newly signed-in account failed', retryError);
        }
      }
      if (e instanceof InteractionRequiredAuthError) {
        throw new AuthInteractionRequiredError(e);
      }
      const code = (e as { errorCode?: string }).errorCode || '';
      // Third-party cookie blocking / iframe timeouts also need an interactive fallback.
      if (/monitor_window_timeout|iframe|login_required|consent_required|interaction_required|no_account/i.test(code)) {
        throw new AuthInteractionRequiredError(e);
      }
      throw e;
    }
  });
}

const SIGNED_IN_EVENT = 'copilot-studio-agent:signed-in';

/**
 * Interactive fallback - must be called from a click handler. The popup is opened right away instead of
 * waiting in the request queue: queued behind another web part's silent attempt, the click's user activation
 * expires and the browser blocks the popup (observed in test TC-03 with two agents on one page).
 */
export function signInWithPopup(config: IAgentAuthConfig, scopes: string[]): Promise<AuthenticationResult> {
  return getClient(config).then(async pca => {
    const result = await pca.acquireTokenPopup({ scopes, loginHint: config.loginHint });
    pca.setActiveAccount(result.account);
    window.dispatchEvent(new CustomEvent<string>(SIGNED_IN_EVENT, { detail: keyOf(config) }));
    return result;
  });
}

/** Notifies other web parts on the page (same app and tenant) that an interactive sign-in completed. */
export function onSignedIn(config: IAgentAuthConfig, callback: () => void): () => void {
  const key = keyOf(config);
  const listener = (e: Event): void => {
    if ((e as CustomEvent<string>).detail === key) { callback(); }
  };
  window.addEventListener(SIGNED_IN_EVENT, listener);
  return () => window.removeEventListener(SIGNED_IN_EVENT, listener);
}
