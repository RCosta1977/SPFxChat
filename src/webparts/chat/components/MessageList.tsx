import * as React from "react";
import type { IChatMessage } from "../../../models/IChatMessage";
import type { WebPartContext } from "@microsoft/sp-webpart-base";
import { AttachmentItem } from "./AttachmentItem";

interface Props {
  context: WebPartContext;
  messages: IChatMessage[];
}

function highlightMentions(text: string): Array<string | JSX.Element> {
  // realça blocos @Nome com <strong>
  const parts: Array<string | JSX.Element> = [];
  const regex = /(^|\s)(@\S+)/g;
  let lastIndex = 0;
  let m: RegExpExecArray | null;
  while ((m = regex.exec(text)) !== null) {
    const start = m.index;
    if (start > lastIndex) parts.push(text.slice(lastIndex, start));
    parts.push(m[1] || " ");
    parts.push(<strong key={start} style={{ fontWeight: 600 }}>{m[2]}</strong>);
    lastIndex = regex.lastIndex;
  }
  if (lastIndex < text.length) parts.push(text.slice(lastIndex));
  return parts;
}

export function MessageList({ context, messages }: Props) {
  if (!messages?.length) {
    return <div style={{ opacity: 0.7 }}>Sem mensagens nesta página — sê tu o primeiro a escrever 🙂</div>;
  }

  return (
    <div>
      {messages.map(msg => (
        <div key={msg.id} style={{ borderBottom: "1px solid #eee", padding: "10px 0" }}>
          <div style={{ fontSize: 12, opacity: 0.8 }}>
            <span style={{ fontWeight: 600 }}>{msg.author.displayName}</span>{" "}
            <span>• {new Date(msg.created).toLocaleString()}</span>
          </div>
          <div style={{ marginTop: 4, whiteSpace: "pre-wrap" }}>
            {highlightMentions(msg.text)}
          </div>
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
