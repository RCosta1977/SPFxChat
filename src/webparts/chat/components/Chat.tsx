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

export default function Chat(props: IChatProps) {
  const { context } = props;
  const [messages, setMessages] = React.useState<IChatMessage[]>([]);
  const [loading, setLoading] = React.useState(true);
  const [error, setError] = React.useState<string | null>(null);
  const pageInfoRef = React.useRef<{ pageName: string; pageUniqueId: string } | null>(null);

  const load = React.useCallback(async () => {
    try {
      setLoading(true);
      setError(null);
      // garante que serviços estão inicializados (defensivo)
      SetupService.init(context);
      await GraphService.init(context);
      await SetupService.ensureList();

      const pageInfo = await SharePointService.getPageInfo(context);
      pageInfoRef.current = pageInfo;
      const items = await SharePointService.getMessages(pageInfo.pageUniqueId);
      setMessages(items);
    } catch (e: any) {
      setError(e?.message || "Erro ao carregar mensagens");
    } finally {
      setLoading(false);
    }
  }, [context]);

  React.useEffect(() => { load(); }, [load]);

  const handleMessageSent = (m: IChatMessage) => {
    // prepend otimista
    setMessages(prev => [m, ...prev]);
  };

  return (
    <div className={styles.chat}>
      {error && <div className={styles.error}>⚠️ {error}</div>}
      <MessageInput
        context={context}
        onMessageSent={handleMessageSent}
        pageInfo={pageInfoRef.current || undefined}
      />
      {loading ? (
        <div className={styles.loading}>A carregar…</div>
      ) : (
        <MessageList context={context} messages={messages} />
      )}
    </div>
  );
}
