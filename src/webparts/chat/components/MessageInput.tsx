import * as React from "react";
import { PrimaryButton, DefaultButton, Stack, Label } from "@fluentui/react";
import type { WebPartContext } from "@microsoft/sp-webpart-base";
import type { IChatMessage } from "../../../models/IChatMessage";
import type { IUserMention } from "../../../models/IUserMention";
import type { IFileAttachment } from "../../../models/IFileAttachment";
import { SharePointService } from "../../../services/SharePointService";
import { GraphService } from "../../../services/GraphService";
import { getPageDeepLink } from "../../../utils/pageHelpers";
import { MentionPicker } from "./MentionPicker";

interface Props {
  context: WebPartContext;
  onMessageSent: (m: IChatMessage) => void;
  pageInfo?: { pageName: string; pageUniqueId: string };
}

export function MessageInput({ context, onMessageSent, pageInfo }: Props) {
  const [text, setText] = React.useState("");
  const [members, setMembers] = React.useState<IUserMention[]>([]);
  const [mentions, setMentions] = React.useState<IUserMention[]>([]);
  const [files, setFiles] = React.useState<File[]>([]);
  const [sending, setSending] = React.useState(false);
  const [error, setError] = React.useState<string | null>(null);

  // mention picker state
  const [pickerOpen, setPickerOpen] = React.useState(false);
  const [pickerQuery, setPickerQuery] = React.useState("");
  const textareaRef = React.useRef<HTMLTextAreaElement | null>(null);

  React.useEffect(() => {
    // carregar membros do site (grupo Members)
    SharePointService.getSiteMembers()
      .then(ms => setMembers(ms))
      .catch(() => setMembers([]));
  }, []);

  const onTextChange = (_: any, v?: string) => {
    const value = v ?? "";
    setText(value);

    const caret = textareaRef.current?.selectionStart ?? value.length;
    // detetar a palavra atual onde está o cursor
    const before = value.slice(0, caret);
    const token = before.split(/\s/).pop() || "";
    if (token.startsWith("@")) {
      const q = token.slice(1);
      setPickerQuery(q);
      setPickerOpen(true);
    } else {
      setPickerOpen(false);
      setPickerQuery("");
    }
  };

  const insertMentionAtCaret = (m: IUserMention) => {
    const el = textareaRef.current;
    if (!el) return;
    const caret = el.selectionStart;
    const value = text;
    const before = value.slice(0, caret);
    const after = value.slice(caret);
    const tokens = before.split(/\s/);
    tokens.pop(); // remove o token com '@'
    const beforeNew = tokens.join(" ");
    const display = `@${m.displayName}`;
    const spacer = beforeNew && !beforeNew.endsWith(" ") ? " " : "";
    const newText = `${beforeNew}${spacer}${display} ${after}`;
    setText(newText);
    setPickerOpen(false);
    setPickerQuery("");
    // manter menção única
    setMentions(prev =>
      prev.some(x => x.email.toLowerCase() === m.email.toLowerCase()) ? prev : [...prev, m]
    );
    // focar de novo
    setTimeout(() => {
      el.focus();
      const newPos = (beforeNew + spacer + display + " ").length;
      el.setSelectionRange(newPos, newPos);
    }, 0);
  };

  const onFilesPicked = (e: React.ChangeEvent<HTMLInputElement>) => {
    const selected = Array.from(e.target.files || []);
    const invalid = selected.find(f => f.size > 5 * 1024 * 1024);
    if (invalid) {
      setError(`Ficheiro ${invalid.name} excede 5MB`);
      return;
    }
    setFiles(prev => [...prev, ...selected]);
    // reset input
    e.currentTarget.value = "";
  };

  const removeFile = (name: string) => {
    setFiles(prev => prev.filter(f => f.name !== name));
  };

  const handleSend = async () => {
    if (!text.trim() && files.length === 0) {
      setError("Escreve uma mensagem ou adiciona um ficheiro.");
      return;
    }
    setSending(true);
    setError(null);

    try {
      // info da página
      const info = pageInfo || (await SharePointService.getPageInfo(context));
      // upload dos ficheiros (pasta {NomeDaPágina})
      let uploaded: IFileAttachment[] = [];
      if (files.length) {
        uploaded = await SharePointService.uploadFiles(context, files);
      }

      const currentUser = {
        id: context.pageContext.legacyPageContext?.userId?.toString() || "",
        displayName: context.pageContext.user.displayName,
        email: context.pageContext.user.email || ""
      };

      const message: IChatMessage = {
        text: text.trim(),
        created: new Date().toISOString(),
        author: currentUser,
        mentions: mentions,
        attachments: uploaded,
        pageUniqueId: info.pageUniqueId,
        pageName: info.pageName
      };

      // gravar na lista
      const id = await SharePointService.addMessage(message);
      message.id = id;

      // enviar email aos mencionados (se houver)
      if (mentions.length) {
        const preview = message.text.slice(0, 200);
        const deepLink = getPageDeepLink(context);
        await GraphService.sendMentionEmails(currentUser.displayName, mentions, preview, deepLink);
      }

      // notificar UI
      onMessageSent(message);

      // reset
      setText("");
      setMentions([]);
      setFiles([]);
    } catch (e: any) {
      setError(e?.message || "Falha ao enviar a mensagem");
    } finally {
      setSending(false);
    }
  };

  const filteredSuggestions = React.useMemo(() => {
  if (!pickerOpen) return [];
  const q = pickerQuery.trim().toLowerCase();
  const base = members;
  const filtered = q
    ? base.filter(m =>
        m.displayName.toLowerCase().includes(q) ||
        m.email.toLowerCase().includes(q)
      )
    : base;
  return filtered.slice(0, 8);
}, [pickerOpen, pickerQuery, members]);


  return (
    <Stack tokens={{ childrenGap: 8 }}>
      {error && <div style={{ color: "#a4262c" }}>⚠️ {error}</div>}
      <div style={{ position: "relative" }}>
        <label style={{ display: "block", fontWeight: 600, marginBottom: 4}}>mensagem</label>
        <textarea
          ref={textareaRef}
          value={text}
          onChange={(e) => onTextChange(undefined, e.target.value)}
          placeholder="Escreva aqui... usa @ para mencionar"
          rows={4}
          style={{
            width:"100%",
            boxSizing:"border-box",
            padding:8,
            border:"1px solid #ddd",
            borderRadius: 6,
            resize:"vertical",
            font:"inherit"
          }}
          />
        <MentionPicker
          open={pickerOpen}
            suggestions={filteredSuggestions}
            onSelect={insertMentionAtCaret}
        />
      </div>

      <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="end">
        <div>
          <Label>Anexos (máx. 5MB cada)</Label>
          <input type="file" multiple onChange={onFilesPicked} />
          {!!files.length && (
            <ul style={{ margin: "6px 0" }}>
              {files.map(f => (
                <li key={f.name}>
                  {f.name} ({Math.round(f.size / 1024)} KB)
                  {" "}
                  <DefaultButton
                    text="remover"
                    onClick={() => removeFile(f.name)}
                    styles={{ root: { height: 22, padding: "0 6px", marginLeft: 6 } }}
                  />
                </li>
              ))}
            </ul>
          )}
        </div>

        <div style={{ marginLeft: "auto" }}>
          <PrimaryButton text={sending ? "A enviar…" : "Enviar"} onClick={handleSend} disabled={sending} />
        </div>
      </Stack>
    </Stack>
  );
}
