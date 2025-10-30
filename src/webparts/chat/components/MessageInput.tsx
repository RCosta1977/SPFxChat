import * as React from "react";
import { PrimaryButton, DefaultButton, Stack, Label, IconButton } from "@fluentui/react";
import type { WebPartContext } from "@microsoft/sp-webpart-base";
import type { IChatMessage } from "../../../models/IChatMessage";
import type { IUserMention } from "../../../models/IUserMention";
import type { IFileAttachment } from "../../../models/IFileAttachment";
import { SharePointService } from "../../../services/SharePointService";
import { GraphService } from "../../../services/GraphService";
import { getPageDeepLink } from "../../../utils/pageHelpers";
import { MentionPicker } from "./MentionPicker";
import styles from "./Chat.module.scss";
import { getPlainTextFromHtml, sanitizeRichText } from "../../../utils/richText";

interface Props {
  context: WebPartContext;
  onMessageSent: (m: IChatMessage) => void;
  pageInfo?: { pageName: string; pageUniqueId: string };
}

interface MentionContext {
  node: Text;
  caretOffset: number;
  tokenLength: number;
}

export function MessageInput({ context, onMessageSent, pageInfo }: Props): JSX.Element {
  const [html, setHtml] = React.useState("");
  const [plainText, setPlainText] = React.useState("");
  const [members, setMembers] = React.useState<IUserMention[]>([]);
  const [mentions, setMentions] = React.useState<IUserMention[]>([]);
  const [files, setFiles] = React.useState<File[]>([]);
  const [sending, setSending] = React.useState(false);
  const [error, setError] = React.useState<string | null>(null);

  // mention picker state
  const [pickerOpen, setPickerOpen] = React.useState(false);
  const [pickerQuery, setPickerQuery] = React.useState("");
  const editorRef = React.useRef<HTMLDivElement | null>(null);
  const mentionContextRef = React.useRef<MentionContext | null>(null);

  React.useEffect(() => {
    SharePointService.getSiteMembers()
      .then(ms => setMembers(ms))
      .catch(() => setMembers([]));
  }, []);

  const focusEditor = React.useCallback((): void => {
    const editor = editorRef.current;
    if (!editor) {
      return;
    }
    editor.focus();
    const selection = window.getSelection();
    if (!selection || selection.rangeCount === 0) {
      const range = document.createRange();
      range.selectNodeContents(editor);
      range.collapse(false);
      selection?.addRange(range);
    }
  }, []);

  const detectMentionTrigger = React.useCallback((): void => {
    const selection = window.getSelection();
    if (!selection || selection.rangeCount === 0) {
      mentionContextRef.current = null;
      setPickerOpen(false);
      setPickerQuery("");
      return;
    }

    const range = selection.getRangeAt(0);
    if (range.startContainer.nodeType !== Node.TEXT_NODE) {
      mentionContextRef.current = null;
      setPickerOpen(false);
      setPickerQuery("");
      return;
    }

    const node = range.startContainer as Text;
    const caretOffset = range.startOffset;
    const textUntilCaret = node.data.slice(0, caretOffset);
    const token = textUntilCaret.split(/\s/).pop() ?? "";

    if (token.startsWith("@")) {
      mentionContextRef.current = {
        node,
        caretOffset,
        tokenLength: token.length
      };
      setPickerQuery(token.slice(1));
      setPickerOpen(true);
    } else {
      mentionContextRef.current = null;
      setPickerOpen(false);
      setPickerQuery("");
    }
  }, []);

  const insertMentionAtCaret = (mention: IUserMention): void => {
    const editor = editorRef.current;
    const ctx = mentionContextRef.current;
    if (!editor || !ctx) {
      return;
    }

    const { node, caretOffset, tokenLength } = ctx;
    const tokenStart = caretOffset - tokenLength;
    if (tokenStart < 0) {
      return;
    }

    const tokenNode = node.splitText(tokenStart);
    const afterNode = tokenNode.splitText(tokenLength);
    tokenNode.remove();

    const mentionSpan = document.createElement("span");
    mentionSpan.textContent = `@${mention.displayName}`;
    mentionSpan.setAttribute("data-mention", "true");
    mentionSpan.setAttribute("data-email", mention.email);
    mentionSpan.className = styles.mention;

    tokenNode.parentNode?.insertBefore(mentionSpan, afterNode);
    tokenNode.parentNode?.insertBefore(document.createTextNode(" "), afterNode);

    const selection = window.getSelection();
    if (selection) {
      selection.removeAllRanges();
      const range = document.createRange();
      range.setStartAfter(mentionSpan);
      range.collapse(true);
      selection.addRange(range);
    }

    setHtml(editor.innerHTML);
    setPlainText(editor.innerText ?? "");
    mentionContextRef.current = null;
    setPickerOpen(false);
    setPickerQuery("");
    setMentions(prev =>
      prev.some(x => x.email.toLowerCase() === mention.email.toLowerCase()) ? prev : [...prev, mention]
    );
  };

  const onFilesPicked = (e: React.ChangeEvent<HTMLInputElement>): void => {
    const selected = Array.from(e.target.files || []);
    const invalid = selected.find(f => f.size > 5 * 1024 * 1024);
    if (invalid) {
      setError(`Ficheiro ${invalid.name} excede 5MB`);
      return;
    }
    setFiles(prev => [...prev, ...selected]);
    e.currentTarget.value = "";
  };

  const removeFile = (name: string): void => {
    setFiles(prev => prev.filter(f => f.name !== name));
  };

  const handleSend = async (): Promise<void> => {
    const sanitized = sanitizeRichText(html);
    const plain = getPlainTextFromHtml(sanitized).trim();

    if (!plain && files.length === 0) {
      setError("Escreve uma mensagem ou adiciona um ficheiro.");
      return;
    }

    setSending(true);
    setError(null);

    try {
      const info = pageInfo || (await SharePointService.getPageInfo(context));
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
        text: sanitized,
        created: new Date().toISOString(),
        author: currentUser,
        mentions,
        attachments: uploaded,
        pageUniqueId: info.pageUniqueId,
        pageName: info.pageName
      };

      const id = await SharePointService.addMessage(message);
      message.id = id;

      if (mentions.length) {
        const preview = plain.slice(0, 200);
        const deepLink = getPageDeepLink(context);
        await GraphService.sendMentionEmails(currentUser.displayName, mentions, preview, deepLink);
      }

      onMessageSent(message);

      setHtml("");
      setPlainText("");
      setMentions([]);
      setFiles([]);
      mentionContextRef.current = null;
      if (editorRef.current) {
        editorRef.current.innerHTML = "";
      }
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

  const handleEditorInput = (): void => {
    const editor = editorRef.current;
    if (!editor) {
      return;
    }
    setHtml(editor.innerHTML);
    setPlainText(editor.innerText ?? "");
    detectMentionTrigger();
  };

  const handleEditorKeyUp = (): void => {
    detectMentionTrigger();
  };

  const handleEditorMouseUp = (): void => {
    detectMentionTrigger();
  };

  const applyCommand = (command: string, value?: string): void => {
    focusEditor();
    if (typeof document !== "undefined") {
      document.execCommand(command, false, value);
    }
    handleEditorInput();
  };

  return (
    <Stack tokens={{ childrenGap: 8 }}>
      {error && <div className={styles.errorBanner}>[!] {error}</div>}
      <div>
        <label className={styles.editorLabel}>mensagem</label>
        <Stack horizontal tokens={{ childrenGap: 4 }} styles={{ root: { marginBottom: 6 } }}>
          <IconButton iconProps={{ iconName: "Bold" }} title="Negrito" ariaLabel="Negrito" onMouseDown={e => e.preventDefault()} onClick={() => applyCommand("bold")} />
          <IconButton iconProps={{ iconName: "Italic" }} title="Italico" ariaLabel="Italico" onMouseDown={e => e.preventDefault()} onClick={() => applyCommand("italic")} />
          <IconButton iconProps={{ iconName: "Underline" }} title="Sublinhado" ariaLabel="Sublinhado" onMouseDown={e => e.preventDefault()} onClick={() => applyCommand("underline")} />
          <IconButton iconProps={{ iconName: "BulletedList" }} title="Lista com marcadores" ariaLabel="Lista com marcadores" onMouseDown={e => e.preventDefault()} onClick={() => applyCommand("insertUnorderedList")} />
          <IconButton iconProps={{ iconName: "NumberedList" }} title="Lista numerada" ariaLabel="Lista numerada" onMouseDown={e => e.preventDefault()} onClick={() => applyCommand("insertOrderedList")} />
          <IconButton iconProps={{ iconName: "Link" }} title="Inserir ligacao" ariaLabel="Inserir ligacao" onMouseDown={e => e.preventDefault()} onClick={() => {
            const href = prompt("URL da ligacao:");
            if (href) {
              applyCommand("createLink", href);
            }
          }} />
          <IconButton iconProps={{ iconName: "ClearFormatting" }} title="Limpar formatacao" ariaLabel="Limpar formatacao" onMouseDown={e => e.preventDefault()} onClick={() => applyCommand("removeFormat")} />
        </Stack>
        <div className={styles.richEditorWrapper}>
          {!plainText.trim() && (
            <div className={styles.richEditorPlaceholder}>
              Escreva aqui... usa @ para mencionar
            </div>
          )}
          <div
            ref={editorRef}
            className={styles.richEditor}
            contentEditable
            onInput={handleEditorInput}
            onKeyUp={handleEditorKeyUp}
            onMouseUp={handleEditorMouseUp}
            onBlur={() => setHtml(editorRef.current?.innerHTML ?? "")}
            role="textbox"
            aria-multiline="true"
          />
        </div>
        <MentionPicker
          open={pickerOpen}
          suggestions={filteredSuggestions}
          onSelect={insertMentionAtCaret}
        />
      </div>

      <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="end">
        <div>
          <Label>Anexos (max. 5MB cada)</Label>
          <input type="file" multiple onChange={onFilesPicked} />
          {!!files.length && (
            <ul style={{ margin: "6px 0" }}>
              {files.map(f => (
                <li key={f.name}>
                  {f.name} ({Math.round(f.size / 1024)} KB){" "}
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
          <PrimaryButton text={sending ? "A enviar..." : "Enviar"} onClick={handleSend} disabled={sending} />
        </div>
      </Stack>
    </Stack>
  );
}
