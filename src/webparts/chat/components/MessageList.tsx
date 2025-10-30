import * as React from "react";
import type { IChatMessage } from "../../../models/IChatMessage";
import type { WebPartContext } from "@microsoft/sp-webpart-base";
import { AttachmentItem } from "./AttachmentItem";
import { sanitizeRichText } from "../../../utils/richText";
import styles from "./Chat.module.scss";

interface Props {
  context: WebPartContext;
  messages: IChatMessage[];
}

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function buildHtml(text: string): string {
  if (!text) {
    return "";
  }

  const trimmed = text.trim();
  const looksLikeHtml = /<\/?[a-z][\s\S]*>/i.test(trimmed);

  if (looksLikeHtml) {
    return sanitizeRichText(trimmed);
  }

  const escaped = escapeHtml(trimmed).replace(/\r?\n/g, "<br/>");
  return `<p>${escaped}</p>`;
}

export function MessageList({ context: _context, messages }: Props): JSX.Element {
  if (!messages?.length) {
    return <div style={{ opacity: 0.7 }}>Sem mensagens nesta pagina - seja o primeiro a escrever :)</div>;
  }

  return (
    <div>
      {messages.map(msg => (
        <div key={msg.id} style={{ borderBottom: "1px solid #eee", padding: "10px 0" }}>
          <div style={{ fontSize: 12, opacity: 0.8 }}>
            <span style={{ fontWeight: 600 }}>{msg.author.displayName}</span>{" "}
            <span>- {new Date(msg.created).toLocaleString()}</span>
          </div>
          <div
            className={styles.messageContent}
            style={{ marginTop: 4 }}
            dangerouslySetInnerHTML={{ __html: buildHtml(msg.text) }}
          />
          {!!msg.attachments?.length && (
            <div style={{ marginTop: 6 }}>
              {msg.attachments.map(a => (
                <AttachmentItem key={a.serverRelativeUrl} attachment={a} />
              ))}
            </div>
          )}
        </div>
      ))}
    </div>
  );
}

