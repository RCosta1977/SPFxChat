import * as React from "react";
import styles from "./Chat.module.scss";
import type { IChatMessage } from "../../../models/IChatMessage";
import { SharePointService } from "../../../services/SharePointService";
import { SetupService } from "../../../services/SetupService";
import { GraphService } from "../../../services/GraphService";
import { MessageList } from "./MessageList";
import { MessageInput } from "./MessageInput";
import type { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IChatProps {
  context: WebPartContext;
}

export default function Chat({ context }: IChatProps): JSX.Element {
  const [messages, setMessages] = React.useState<IChatMessage[]>([]);
  const [loading, setLoading] = React.useState(true);
  const [error, setError] = React.useState<string | null>(null);

  const pageInfoRef = React.useRef<{ pageName: string; pageUniqueId: string } | null>(null);
  const messagesRef = React.useRef<HTMLDivElement | null>(null);

  // carrega mensagens da página
  const load = React.useCallback(async (): Promise<void> => {
    try {
      setLoading(true);
      setError(null);

      SetupService.init(context);
      await GraphService.init(context);
      await SetupService.ensureList();

      const pageInfo = await SharePointService.getPageInfo(context);
      pageInfoRef.current = pageInfo;

      // Se quiseres mais recente no fim, garante que o serviço devolve ascendente
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

  // autoscroll para o fundo quando o array muda
  React.useEffect(() => {
    const el = messagesRef.current;
    if (el) el.scrollTop = el.scrollHeight;
  }, [messages]);

  // usado no MessageInput
  const handleMessageSent = (message: IChatMessage): void => {
    // acrescenta no fim (mantém a ordem cronológica ascendente)
    setMessages(prev => [...prev, message]);
  };

  return (
    <div className={styles.chat}>
      {error && <div style={{ color: "#a4262c" }}>⚠️ {error}</div>}

      {/* HISTÓRICO EM CIMA */}
      {loading ? (
        <div style={{ opacity: 0.7 }}>A carregar…</div>
      ) : (
        <div className={styles.messagesContainer} ref={messagesRef}>
          <MessageList context={context} messages={messages} />
        </div>
      )}

      {/* INPUT EM BAIXO */}
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
