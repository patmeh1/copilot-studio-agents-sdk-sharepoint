import {
  ConnectionSettings,
  CopilotStudioClient,
  CopilotStudioWebChat
} from '@microsoft/agents-copilotstudio-client';

interface IObserverLike<T> {
  next?: (value: T) => void;
  error?: (error: unknown) => void;
  complete?: () => void;
}

interface IUnsubscribable {
  unsubscribe(): void;
}

/** Minimal Observable shape consumed by Bot Framework Web Chat (avoids rxjs 6/7 type clashes). */
export interface ISubscribable<T> {
  subscribe(observerOrNext?: IObserverLike<T> | ((value: T) => void), error?: (e: unknown) => void, complete?: () => void): IUnsubscribable;
}

/** DirectLine-compatible connection handed to Web Chat. */
export interface IAgentWebChatConnection {
  connectionStatus$: ISubscribable<number>;
  activity$: ISubscribable<unknown>;
  readonly conversationId: string | undefined;
  postActivity(activity: unknown): ISubscribable<string>;
  end(): void;
}

function tapErrors<T>(source: ISubscribable<T>, onError: (error: unknown) => void): ISubscribable<T> {
  return {
    subscribe(observerOrNext?: IObserverLike<T> | ((value: T) => void), error?: (e: unknown) => void, complete?: () => void): IUnsubscribable {
      const observer: IObserverLike<T> = typeof observerOrNext === 'function'
        ? { next: observerOrNext, error, complete }
        : (observerOrNext || {});
      return source.subscribe({
        next: (value: T) => { if (observer.next) { observer.next(value); } },
        error: (err: unknown) => { onError(err); if (observer.error) { observer.error(err); } },
        complete: () => { if (observer.complete) { observer.complete(); } }
      });
    }
  };
}

export interface IAgentTarget {
  /** "Microsoft 365 Agents SDK" connection string copied from Copilot Studio (Channels > App). */
  connectionString?: string;
  /** Alternative to the connection string: environment ID + agent schema name. */
  environmentId?: string;
  agentIdentifier?: string;
  tenantId: string;
  appClientId: string;
}

function clean(value?: string): string | undefined {
  const v = (value || '').trim();
  return v ? v : undefined;
}

export function isTargetConfigured(target: IAgentTarget): boolean {
  return !!clean(target.appClientId) && !!clean(target.tenantId) &&
    (!!clean(target.connectionString) || (!!clean(target.environmentId) && !!clean(target.agentIdentifier)));
}

export function buildConnectionSettings(target: IAgentTarget): ConnectionSettings {
  const directConnectUrl = clean(target.connectionString);
  return new ConnectionSettings({
    appClientId: clean(target.appClientId),
    tenantId: clean(target.tenantId),
    directConnectUrl,
    environmentId: directConnectUrl ? undefined : clean(target.environmentId),
    agentIdentifier: directConnectUrl ? undefined : clean(target.agentIdentifier),
    schemaName: directConnectUrl ? undefined : clean(target.agentIdentifier)
  });
}

/** Token audience for the agent, e.g. https://api.powerplatform.com/.default (cloud-aware). */
export function scopeFor(settings: ConnectionSettings): string {
  return CopilotStudioClient.scopeFromSettings(settings);
}

/**
 * Creates the Agents SDK <-> Bot Framework Web Chat bridge and wraps it so that failures on the
 * activity stream (e.g. HTTP 401/403/404 from Copilot Studio) are reported to the caller.
 */
export function createWebChatConnection(
  settings: ConnectionSettings,
  accessToken: string,
  onError: (error: unknown) => void
): IAgentWebChatConnection {
  const client = new CopilotStudioClient(settings, accessToken);
  const inner = CopilotStudioWebChat.createConnection(client, { showTyping: true });

  return {
    connectionStatus$: inner.connectionStatus$ as unknown as ISubscribable<number>,
    activity$: tapErrors(inner.activity$ as unknown as ISubscribable<unknown>, onError),
    get conversationId(): string | undefined { return inner.conversationId; },
    postActivity: (activity: unknown) =>
      tapErrors(inner.postActivity(activity as Parameters<typeof inner.postActivity>[0]) as unknown as ISubscribable<string>, onError),
    end: () => inner.end()
  };
}

/** Human readable guidance for the most common setup problems. */
export function describeError(error: unknown, appClientId: string): string {
  const e = error as { errorCode?: string; message?: string; status?: number; statusCode?: number };
  const text = `${e && e.errorCode ? e.errorCode + ': ' : ''}${(e && e.message) || String(error)}`;
  const origin = window.location.origin;
  if (/AADSTS50011|redirect_uri/i.test(text)) {
    return `Redirect URI mismatch. Add "${origin}" as a Single-page application redirect URI on Entra app ${appClientId}. (${text})`;
  }
  if (/AADSTS65001|consent_required/i.test(text)) {
    return `Consent is missing. An admin must grant consent for "Power Platform API / CopilotStudio.Copilots.Invoke" on Entra app ${appClientId}. (${text})`;
  }
  if (/AADSTS700016|AADSTS90002|unauthorized_client/i.test(text)) {
    return `Entra app ${appClientId} was not found in this tenant - check the App (client) ID / tenant ID. (${text})`;
  }
  if (/popup_window_error|empty_window_error|popup/i.test(text)) {
    return `The sign-in popup was blocked. Allow popups for ${origin} and select Sign in again. (${text})`;
  }
  if (/\b401\b|Unauthorized/i.test(text)) {
    return `Copilot Studio rejected the token (401). Make sure the agent uses "Authenticate with Microsoft" and the app has CopilotStudio.Copilots.Invoke. (${text})`;
  }
  if (/\b403\b|Forbidden/i.test(text)) {
    return `Access denied (403). The agent must be shared with you (or your security group) and published. (${text})`;
  }
  if (/\b404\b|Not Found/i.test(text)) {
    return `Agent not found (404). Verify the connection string / schema name and that the agent is published. (${text})`;
  }
  return text;
}
