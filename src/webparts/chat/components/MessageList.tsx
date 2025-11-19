import * as React from "react";
import type { IChatMessage } from "../../../models/IChatMessage";
import type { WebPartContext } from "@microsoft/sp-webpart-base";
import { AttachmentItem } from "./AttachmentItem";
import { sanitizeRichText } from "../../../utils/richText";
import styles from "./Chat.module.scss";

interface Props {
  context: WebPartContext;
  messages: IChatMessage[];
  currentUserEmail?: string;
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
  const containsNumericEntity = /&#(?:x?[0-9a-f]+);/i.test(trimmed);

  if (looksLikeHtml || containsNumericEntity) {
    const sanitized = sanitizeRichText(trimmed);
    const sanitizedHasTags = /<\/?[a-z][\s\S]*>/i.test(sanitized);
    if (sanitizedHasTags) {
      return sanitized;
    }
    const normalized = sanitized.replace(/\r?\n/g, "<br/>");
    return normalized ? `<p>${normalized}</p>` : "";
  }

  const escaped = escapeHtml(trimmed).replace(/\r?\n/g, "<br/>");
  return `<p>${escaped}</p>`;
}

export function MessageList({ context: _context, messages, currentUserEmail }: Props): JSX.Element {
  if (!messages?.length) {
    return <div style={{ opacity: 0.7 }}>Sem mensagens nesta pagina - seja o primeiro a escrever :)</div>;
  }

  return (
    <div>
      {messages.map(msg => {
        const isOwn = currentUserEmail
          ? (msg.author.email || "").toLowerCase() === currentUserEmail.toLowerCase()
          : false;
        const containerClass = isOwn
          ? `${styles.messageItem} ${styles.ownMessage}`
          : styles.messageItem;

        return (
          <div key={msg.id} className={containerClass}>
            <div className={styles.messageMeta}>
              <span className={styles.messageAuthor}>{msg.author.displayName}</span>{" "}
              <span>- {new Date(msg.created).toLocaleString()}</span>
            </div>
            <div
              className={styles.messageContent}
              dangerouslySetInnerHTML={{ __html: buildHtml(msg.text) }}
            />
            {!!msg.attachments?.length && (
              <div className={styles.messageAttachments}>
                {msg.attachments.map(a => (
                  <AttachmentItem key={a.serverRelativeUrl} attachment={a} />
                ))}
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
}
