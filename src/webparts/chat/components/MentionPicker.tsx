import * as React from "react";
import type { IUserMention } from "../../../models/IUserMention";

interface Props {
  open: boolean;
  suggestions: IUserMention[];
  onSelect: (m: IUserMention) => void;
}

export function MentionPicker({ open, suggestions, onSelect }: Props) {
  if (!open || suggestions.length === 0) return null;

  return (
    <div
      style={{
        position: "absolute",
        left: 8,
        right: 8,
        top: "100%",
        zIndex: 10,
        background: "white",
        border: "1px solid #ddd",
        borderRadius: 6,
        boxShadow: "0 4px 12px rgba(0,0,0,0.08)",
        marginTop: 4,
        maxHeight: 200,
        overflowY: "auto"
      }}
    >
      {suggestions.map(s => (
        <div
          key={s.email}
          onClick={() => onSelect(s)}
          role="button"
          style={{
            padding: "8px 10px",
            cursor: "pointer"
          }}
          onMouseDown={e => e.preventDefault()} // impedir perder foco do textarea
        >
          <div style={{ fontWeight: 600 }}>{s.displayName}</div>
          <div style={{ fontSize: 12, opacity: 0.8 }}>{s.email}</div>
        </div>
      ))}
    </div>
  );
}
