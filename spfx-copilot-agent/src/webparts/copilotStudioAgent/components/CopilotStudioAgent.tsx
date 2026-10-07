import * as React from 'react';
import { ReactWebChat } from 'botframework-webchat';
import { FluentThemeProvider } from 'botframework-webchat-fluent-theme';
import styles from './CopilotStudioAgent.module.scss';
import type { ICopilotStudioAgentProps } from './ICopilotStudioAgentProps';
import {
  acquireTokenSilently,
  AuthInteractionRequiredError,
  IAgentAuthConfig,
  onSignedIn,
  signInWithPopup
} from '../../../services/AgentAuthService';
import {
  buildConnectionSettings,
  createWebChatConnection,
  describeError,
  IAgentTarget,
  IAgentWebChatConnection,
  isTargetConfigured,
  scopeFor
} from '../../../services/AgentConnection';

type Phase = 'unconfigured' | 'signingIn' | 'needsSignIn' | 'connecting' | 'ready' | 'error';

// The Web Chat typings target React 18; SPFx ships React 17 typings, so treat them as plain components.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const WebChat = ReactWebChat as unknown as React.ComponentType<any>;
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const FluentTheme = FluentThemeProvider as unknown as React.ComponentType<any>;

const AgentsSdkChat: React.FC<ICopilotStudioAgentProps> = (props) => {
  const [phase, setPhase] = React.useState<Phase>('signingIn');
  const [connection, setConnection] = React.useState<IAgentWebChatConnection | undefined>(undefined);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const [signedInAs, setSignedInAs] = React.useState<string | undefined>(undefined);
  const [attempt, setAttempt] = React.useState<number>(0);

  const target: IAgentTarget = {
    connectionString: props.connectionString,
    environmentId: props.environmentId,
    agentIdentifier: props.agentIdentifier,
    tenantId: props.tenantId,
    appClientId: props.appClientId
  };
  const configured = isTargetConfigured(target);
  const authConfig: IAgentAuthConfig = {
    clientId: (props.appClientId || '').trim(),
    tenantId: (props.tenantId || '').trim(),
    loginHint: props.userLoginName
  };

  React.useEffect(() => {
    if (!configured) {
      setPhase('unconfigured');
      return undefined;
    }
    let cancelled = false;
    let current: IAgentWebChatConnection | undefined;

    const start = async (): Promise<void> => {
      setError(undefined);
      setConnection(undefined);
      setPhase('signingIn');
      const settings = buildConnectionSettings(target);
      const auth = await acquireTokenSilently(authConfig, [scopeFor(settings)]);
      if (cancelled) { return; }
      setSignedInAs(auth.account ? auth.account.username : undefined);
      setPhase('connecting');
      current = createWebChatConnection(settings, auth.accessToken, (err) => {
        if (!cancelled) {
          setError(describeError(err, authConfig.clientId));
        }
      });
      setConnection(current);
      setPhase('ready');
    };

    start().catch((err) => {
      if (cancelled) { return; }
      if (err instanceof AuthInteractionRequiredError) {
        setPhase('needsSignIn');
      } else {
        setError(describeError(err, authConfig.clientId));
        setPhase('error');
      }
    });

    return () => {
      cancelled = true;
      if (current) {
        try { current.end(); } catch { /* already ended */ }
      }
    };
  }, [configured, props.connectionString, props.environmentId, props.agentIdentifier, props.appClientId, props.tenantId, props.userLoginName, attempt]);

  // One interactive sign-in is enough for every agent web part on the page: the others retry silently.
  React.useEffect(() => {
    if (phase !== 'needsSignIn') {
      return undefined;
    }
    return onSignedIn(authConfig, () => setAttempt(a => a + 1));
  }, [phase, authConfig.clientId, authConfig.tenantId]);

  const onSignIn = (): void => {
    const settings = buildConnectionSettings(target);
    setPhase('signingIn');
    signInWithPopup(authConfig, [scopeFor(settings)])
      .then(() => setAttempt(a => a + 1))
      .catch(err => {
        setError(describeError(err, authConfig.clientId));
        setPhase('error');
      });
  };

  const restart = (): void => setAttempt(a => a + 1);

  const styleOptions = {
    accent: props.accentColor || undefined,
    hideUploadButton: true,
    rootHeight: '100%',
    rootWidth: '100%',
    bubbleBorderRadius: 8,
    bubbleFromUserBorderRadius: 8,
    sendBoxBorderTop: 'solid 1px #e1dfdd',
    botAvatarInitials: 'AI',
    userAvatarInitials: initials(props.userDisplayName)
  };

  const renderBody = (): JSX.Element => {
    switch (phase) {
      case 'unconfigured':
        return (
          <div className={styles.placeholder}>
            <div className={styles.placeholderTitle}>Connect a Copilot Studio agent</div>
            <div>Paste the <b>Microsoft 365 Agents SDK connection string</b> (Copilot Studio &gt; agent &gt; Channels &gt; App) and the Entra app (client) ID in the web part properties.</div>
            {props.isEditMode && <button className={styles.primaryButton} onClick={props.onConfigure}>Configure</button>}
          </div>
        );
      case 'signingIn':
        return <div className={styles.status}><span className={styles.spinner} />Signing you in with Microsoft Entra ID…</div>;
      case 'connecting':
        return <div className={styles.status}><span className={styles.spinner} />Connecting to the agent…</div>;
      case 'needsSignIn':
        return (
          <div className={styles.placeholder}>
            <div className={styles.placeholderTitle}>Sign in to chat with this agent</div>
            <div>Single sign-on could not complete silently (for example because third-party cookies are blocked or consent is pending).</div>
            <button className={styles.primaryButton} onClick={onSignIn}>Sign in with Microsoft</button>
          </div>
        );
      case 'error':
        return (
          <div className={styles.placeholder}>
            <div className={styles.errorText}>{error}</div>
            <button className={styles.primaryButton} onClick={restart}>Try again</button>
          </div>
        );
      default: {
        const chat = <WebChat key={attempt} directLine={connection} styleOptions={styleOptions} locale={props.locale} />;
        return props.useFluentTheme ? <FluentTheme variant="fluent">{chat}</FluentTheme> : chat;
      }
    }
  };

  return (
    <section className={styles.copilotStudioAgent} style={{ height: Math.max(320, props.height || 600) }}>
      {props.showHeader && (
        <header className={styles.header}>
          <div className={styles.headerTitle}>{props.title || 'Agent'}</div>
          {phase === 'ready' && (
            <button className={styles.linkButton} onClick={restart} title="Start a new conversation">New chat</button>
          )}
        </header>
      )}
      {phase === 'ready' && error && (
        <div className={styles.banner} role="alert">
          {error}
          <button className={styles.linkButton} onClick={restart}>Retry</button>
        </div>
      )}
      <div className={styles.body}>{renderBody()}</div>
      <footer className={styles.footer}>
        Microsoft 365 Agents SDK · Microsoft Entra ID SSO{signedInAs ? ` · ${signedInAs}` : ''}
      </footer>
    </section>
  );
};

function initials(name: string): string {
  const parts = (name || '').split(' ').filter(p => !!p);
  return parts.length === 0 ? 'Me' : parts.slice(0, 2).map(p => p.charAt(0).toUpperCase()).join('');
}

const HOSTED_CHAT_URL = /^https:\/\/([a-z0-9-]+\.)*(copilotstudio(\.preview)?\.microsoft\.com|powerva\.microsoft\.com)\//i;

/**
 * Fallback for agents the Agents SDK cannot reach (e.g. GitHub Copilot harness agents):
 * renders the Copilot Studio hosted web chat from the Web app channel embed code. Copilot Studio
 * signs the user in with Microsoft Entra ID inside the frame.
 */
const HostedChatFrame: React.FC<ICopilotStudioAgentProps> = (props) => {
  const url = (props.embedUrl || '').trim();
  const valid = HOSTED_CHAT_URL.test(url);
  return (
    <section className={styles.copilotStudioAgent} style={{ height: Math.max(320, props.height || 600) }}>
      {props.showHeader && (
        <header className={styles.header}>
          <div className={styles.headerTitle}>{props.title || 'Agent'}</div>
        </header>
      )}
      <div className={styles.body}>
        {valid ? (
          <iframe src={url} title={props.title || 'Copilot Studio agent'} className={styles.frame} allow="clipboard-write; microphone" />
        ) : (
          <div className={styles.placeholder}>
            <div className={styles.placeholderTitle}>Connect a Copilot Studio agent (hosted web chat)</div>
            <div>Paste the <b>embed code</b> (or its iframe URL) from Copilot Studio &gt; agent &gt; Channels &gt; App &gt; Embed code.</div>
            {props.isEditMode && <button className={styles.primaryButton} onClick={props.onConfigure}>Configure</button>}
          </div>
        )}
      </div>
      <footer className={styles.footer}>Copilot Studio hosted web chat (App channel embed) · Microsoft Entra ID sign-in by Copilot Studio</footer>
    </section>
  );
};

const CopilotStudioAgent: React.FC<ICopilotStudioAgentProps> = (props) =>
  props.mode === 'embed' ? <HostedChatFrame {...props} /> : <AgentsSdkChat {...props} />;

export default CopilotStudioAgent;
