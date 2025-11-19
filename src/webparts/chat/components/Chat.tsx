import * as React from "react";
import styles from "./Chat.module.scss";
import type { IChatMessage } from "../../../models/IChatMessage";
import { SharePointService } from "../../../services/SharePointService";
import { SettingsService } from "../../../services/SettingsService";
import type { IPageInfo } from "../../../services/SharePointService";
import { SetupService } from "../../../services/SetupService";
import { GraphService } from "../../../services/GraphService";
import { MessageList } from "./MessageList";
import { MessageInput } from "./MessageInput";
import type { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ChatTheme {
  primaryButtonBackground: string;
  primaryButtonText: string;
  surfaceBorderColor: string;
  messageBorderColor: string;
  selfMessageBackground: string;
  mentionBackground: string;
  mentionText: string;
}

export interface IChatProps {
  context: WebPartContext;
  theme?: ChatTheme;
}

type ChatThemeStyle = React.CSSProperties & {
  "--chat-button-bg"?: string;
  "--chat-button-text"?: string;
  "--chat-border-color"?: string;
  "--chat-message-border"?: string;
  "--chat-self-background"?: string;
  "--chat-mention-bg"?: string;
  "--chat-mention-text"?: string;
};

const DEFAULT_THEME: ChatTheme = {
  primaryButtonBackground: "#0078d4",
  primaryButtonText: "#ffffff",
  surfaceBorderColor: "#dddddd",
  messageBorderColor: "#eeeeee",
  selfMessageBackground: "#f3f2f1",
  mentionBackground: "#e8f3ff",
  mentionText: "#004578"
};

const DEFAULT_POLL_INTERVAL_MS = 4000;

export default function Chat({ context, theme }: IChatProps): JSX.Element {
  const resolvedTheme = React.useMemo<ChatTheme>(() => ({
    primaryButtonBackground: theme?.primaryButtonBackground || DEFAULT_THEME.primaryButtonBackground,
    primaryButtonText: theme?.primaryButtonText || DEFAULT_THEME.primaryButtonText,
    surfaceBorderColor: theme?.surfaceBorderColor || DEFAULT_THEME.surfaceBorderColor,
    messageBorderColor: theme?.messageBorderColor || DEFAULT_THEME.messageBorderColor,
    selfMessageBackground: theme?.selfMessageBackground || DEFAULT_THEME.selfMessageBackground,
    mentionBackground: theme?.mentionBackground || DEFAULT_THEME.mentionBackground,
    mentionText: theme?.mentionText || DEFAULT_THEME.mentionText
  }), [theme]);

  const rootThemeStyle = React.useMemo<ChatThemeStyle>(() => ({
    "--chat-button-bg": resolvedTheme.primaryButtonBackground,
    "--chat-button-text": resolvedTheme.primaryButtonText,
    "--chat-border-color": resolvedTheme.surfaceBorderColor,
    "--chat-message-border": resolvedTheme.messageBorderColor,
    "--chat-self-background": resolvedTheme.selfMessageBackground,
    "--chat-mention-bg": resolvedTheme.mentionBackground,
    "--chat-mention-text": resolvedTheme.mentionText
  }), [resolvedTheme]);

  const [messages, setMessages] = React.useState<IChatMessage[]>([]);
  const [loading, setLoading] = React.useState(true);
  const [error, setError] = React.useState<string | null>(null);
  const [pollingReady, setPollingReady] = React.useState(false);
  const [pollIntervalMs, setPollIntervalMs] = React.useState<number>(DEFAULT_POLL_INTERVAL_MS);

  const pageInfoRef = React.useRef<IPageInfo | null>(null);
  const messagesRef = React.useRef<HTMLDivElement | null>(null);
  const lastMessageIdRef = React.useRef(0);

  const registerMessages = React.useCallback((items: IChatMessage[]): void => {
    if (!items.length) {
      return;
    }
    let maxId = lastMessageIdRef.current;
    for (const item of items) {
      if (typeof item.id === "number" && item.id > maxId) {
        maxId = item.id;
      }
    }
    lastMessageIdRef.current = maxId;
  }, []);

  const load = React.useCallback(async (): Promise<void> => {
    try {
      setLoading(true);
      setError(null);
      setPollingReady(false);

      SetupService.init(context);
      await GraphService.init(context);
      await SetupService.ensureList();

      const pageInfo = await SharePointService.getPageInfo(context);
      pageInfoRef.current = pageInfo;

      // Ensure lists exist
      await SettingsService.ensureSettingsList();

      // Load poll interval from settings (page override -> global)
      const configured = await SettingsService.getSetting("PollIntervalMs", pageInfo.pageUniqueId);
      if (configured) {
        const n = parseInt(configured, 10);
        if (!Number.isNaN(n)) {
          const clamped = Math.min(15000, Math.max(2000, n));
          setPollIntervalMs(clamped);
        }
      }

      const items = await SharePointService.getMessages(pageInfo.pageUniqueId);
      registerMessages(items);
      setMessages(items);
      setPollingReady(true);
    } catch (error: unknown) {
      const message = error instanceof Error ? error.message : "Erro ao carregar mensagens";
      setError(message);
    } finally {
      setLoading(false);
    }
  }, [context, registerMessages]);

  React.useEffect(() => {
    load().catch(() => {
      /* erros já tratados dentro de load */
    });
  }, [load]);

  const refreshMessages = React.useCallback(async (): Promise<void> => {
    if (!pageInfoRef.current) {
      return;
    }

    try {
      const latest = await SharePointService.getMessages(pageInfoRef.current.pageUniqueId, {
        afterId: lastMessageIdRef.current
      });
      if (!latest.length) {
        return;
      }
      registerMessages(latest);
      setMessages(prev => {
        if (!prev.length) {
          return [...latest];
        }
        const existingIds = new Set(
          prev
            .map(m => typeof m.id === "number" ? m.id : undefined)
            .filter((id): id is number => typeof id === "number")
        );
        const newItems = latest.filter(m =>
          typeof m.id === "number" ? !existingIds.has(m.id) : true
        );
        return newItems.length ? [...prev, ...newItems] : prev;
      });
    } catch (refreshError) {
      // Evita quebrar a experiência por causa de erros intermitentes durante o polling
      console.warn("Falha ao atualizar mensagens automaticamente", refreshError);
    }
  }, [registerMessages]);

  React.useEffect(() => {
    if (!pollingReady) {
      return;
    }
    if (typeof window === "undefined") {
      return;
    }
    const intervalId = window.setInterval(() => {
      refreshMessages().catch(() => undefined);
    }, pollIntervalMs);

    return () => {
      window.clearInterval(intervalId);
    };
  }, [pollingReady, refreshMessages, pollIntervalMs]);

  React.useEffect(() => {
    const el = messagesRef.current;
    if (el) el.scrollTop = el.scrollHeight;
  }, [messages]);

  const handleMessageSent = (message: IChatMessage): void => {
    registerMessages([message]);
    setMessages(prev => {
      const hasId = typeof message.id === "number";
      const exists = hasId && prev.some(m => m.id === message.id);
      return exists ? prev : [...prev, message];
    });
  };

  const currentUserEmail = context.pageContext.user?.email || "";

  return (
    <div className={styles.chat} style={rootThemeStyle}>
      {error && <div className={styles.themeError}>[!] {error}</div>}

      {loading ? (
        <div style={{ opacity: 0.7 }}>A carregar.</div>
      ) : (
        <div className={styles.messagesContainer} ref={messagesRef}>
          <MessageList context={context} messages={messages} currentUserEmail={currentUserEmail} />
        </div>
      )}

      <div className={styles.inputBar}>
        <MessageInput
          context={context}
          onMessageSent={handleMessageSent}
          pageInfo={pageInfoRef.current || undefined}
        />
      </div>
    </div>
  );
}



