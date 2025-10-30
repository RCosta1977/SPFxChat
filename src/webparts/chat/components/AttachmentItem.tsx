import * as React from "react";
import type { IFileAttachment } from "../../../models/IFileAttachment";

export function AttachmentItem({ attachment }: { attachment: IFileAttachment }): React.ReactElement {
  const sizeKb = Math.max(1, Math.round(attachment.size / 1024));
  return (
    <a
      href={attachment.serverRelativeUrl}
      target="_blank"
      rel="noreferrer"
      style={{
        display: "inline-block",
        fontSize: 12,
        padding: "4px 8px",
        border: "1px solid #ddd",
        borderRadius: 6,
        marginRight: 6,
        textDecoration: "none"
      }}
      title={attachment.name}
    >
      [ficheiro] {attachment.name} - {sizeKb} KB
    </a>
  );
}

