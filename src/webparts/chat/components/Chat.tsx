import * as React from "react";
import styles from "./Chat.module.scss";
import type { IChatMessage } from "../../../models/IChatMessage";
import { SharePointService } from "../../../services/SharePointService";
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

  const pageInfoRef = React.useRef<{ pageName: string; pageUniqueId: string } | null>(null);
  const messagesRef = React.useRef<HTMLDivElement | null>(null);

  const load = React.useCallback(async (): Promise<void> => {
    try {
      setLoading(true);
      setError(null);

      SetupService.init(context);
      await GraphService.init(context);
      await SetupService.ensureList();

      const pageInfo = await SharePointService.getPageInfo(context);
      pageInfoRef.current = pageInfo;

      const items = await SharePointService.getMessages(pageInfo.pageUniqueId);
      setMessages(items);
    } catch (error: unknown) {
      const message = error instanceof Error ? error.message : "Erro ao carregar mensagens";
      setError(message);
    } finally {
      setLoading(false);
    }
  }, [context]);

  React.useEffect(() => {
    load().catch(() => {
      /* erros já tratados dentro de load */
    });
  }, [load]);

  React.useEffect(() => {
    const el = messagesRef.current;
    if (el) el.scrollTop = el.scrollHeight;
  }, [messages]);

  const handleMessageSent = (message: IChatMessage): void => {
    setMessages(prev => [...prev, message]);
  };

  const currentUserEmail = context.pageContext.user?.email || "";

  return (
    <div className={styles.chat} style={rootThemeStyle}>
      {error && <div className={styles.themeError}>⚠️ {error}</div>}

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
