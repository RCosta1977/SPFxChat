import { SetupService } from "./SetupService";
import type { IChatMessage } from "../models/IChatMessage";
import type { IFileAttachment } from "../models/IFileAttachment";
import type { WebPartContext } from "@microsoft/sp-webpart-base";
import "@pnp/sp/webs";
import "@pnp/sp/items";
import "@pnp/sp/lists";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "@pnp/sp/site-groups/web";
import "@pnp/sp/site-users/web";

const ATTACHMENTS_LIBRARY_TITLE = "Anexos dos Chats";

export interface IPageInfo {
  pageName: string;
  pageUniqueId: string;
  folderName: string;
}

interface ISharePointError {
  status?: number;
  message?: string;
}

interface IChatListItem {
  Id: number;
  Message: string;
  MentionsJson: string;
  AttachmentsJson: string;
  Created: string;
  PageUniqueId: string;
  PageName: string;
  Author?: {
    Id?: number;
    Title?: string;
    EMail?: string;
  };
}

function toSharePointError(error: unknown): ISharePointError {
  if (typeof error === "object" && error !== null) {
    return error as ISharePointError;
  }
  return {};
}

function sanitizeFolderName(value: string): string {
  const normalized = value.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  const withoutExtension = normalized.replace(/\.[^/.]+$/, "");
  const cleaned = withoutExtension
    .replace(/[^a-zA-Z0-9\-_. ]+/g, " ")
    .trim()
    .replace(/\s+/g, "-");
  return cleaned || "pagina";
}

function getPathSegment(path: string): string {
  const noQuery = path.split("?")[0];
  const segments = noQuery.split("/").filter(Boolean);
  return segments.length ? segments[segments.length - 1] : "pagina";
}

export class SharePointService {
  private static async ensureAttachmentsLibrary(): Promise<string> {
    const sp = SetupService.sp();
    try {
      const info = await sp.web.lists
        .getByTitle(ATTACHMENTS_LIBRARY_TITLE)
        .select("Title,RootFolder/ServerRelativeUrl")
        .expand("RootFolder")();
      return info.RootFolder.ServerRelativeUrl as string;
    } catch {
      // biblioteca não existe, tentar criar em seguida
    }

    try {
      await sp.web.lists.ensure(ATTACHMENTS_LIBRARY_TITLE, "Biblioteca para anexos dos chats", 101);
      const created = await sp.web.lists
        .getByTitle(ATTACHMENTS_LIBRARY_TITLE)
        .select("RootFolder/ServerRelativeUrl")
        .expand("RootFolder")();
      return created.RootFolder.ServerRelativeUrl as string;
    } catch (err) {
      const spError = toSharePointError(err);
      if (spError.status === 403) {
        throw new Error(
          `Sem permissão para criar a biblioteca "${ATTACHMENTS_LIBRARY_TITLE}". ` +
          `Garante que tens 'Manage Lists' no site (ou pede a um Owner para criá-la).`
        );
      }
      throw err;
    }
  }

  static async getPageInfo(context: WebPartContext): Promise<IPageInfo> {
    const legacy = context?.pageContext?.legacyPageContext as Record<string, unknown> | undefined;
    const listItem = context?.pageContext?.listItem as Record<string, unknown> | undefined;
    const listItemAllFields = (context as unknown as { pageContext?: { listItemAllFields?: Record<string, unknown> } })
      ?.pageContext?.listItemAllFields;

    const uniqueCandidates = [
      listItemAllFields?.UniqueId,
      listItem?.uniqueId,
      legacy?.listItemUniqueId,
      legacy?.uniqueId
    ].filter(Boolean);

    const rawPath =
      (typeof window !== "undefined" ? window.location.pathname + window.location.search : "") ||
      (typeof legacy?.serverRequestPath === "string" ? (legacy.serverRequestPath as string) : "") ||
      context?.pageContext?.web?.serverRelativeUrl ||
      "/";

    const decodedPath = decodeURIComponent(rawPath);
    const pathSegment = getPathSegment(decodedPath);
    const folderName = sanitizeFolderName(pathSegment);

    const displayName =
      (listItem?.title as string | undefined) ||
      (listItemAllFields?.Title as string | undefined) ||
      (legacy?.itemTitle as string | undefined) ||
      (typeof document !== "undefined" ? document.title : undefined) ||
      folderName;

    const pageUniqueId = uniqueCandidates.length ? String(uniqueCandidates[0]) : decodedPath;

    return {
      pageName: displayName,
      pageUniqueId,
      folderName
    };
  }

  static async ensurePageFolder(context: WebPartContext): Promise<string> {
    const sp = SetupService.sp();
    const { folderName } = await this.getPageInfo(context);
    const libraryRoot = await this.ensureAttachmentsLibrary();
    const targetUrl = `${libraryRoot.replace(/\/$/, "")}/${folderName}`;

    try {
      const info = await sp.web.getFolderByServerRelativePath(targetUrl)();
      return info.ServerRelativeUrl;
    } catch (err) {
      const spError = toSharePointError(err);
      if (spError.status && spError.status !== 404) {
        if (spError.status === 403) {
          throw new Error(
            `Sem permissão para aceder à pasta "${targetUrl}". ` +
            `Verifica se tens 'Edit' na biblioteca "${ATTACHMENTS_LIBRARY_TITLE}".`
          );
        }
        throw err;
      }
    }

    try {
      const parent = sp.web.getFolderByServerRelativePath(libraryRoot);
      await parent.folders.addUsingPath(folderName);
    } catch (err) {
      const spError = toSharePointError(err);
      if (!(spError.status === 409 || /already exists/i.test(spError.message || ""))) {
        if (spError.status === 403) {
          throw new Error(
            `Sem permissão para criar a pasta "${folderName}" em "${ATTACHMENTS_LIBRARY_TITLE}". ` +
            `Precisas de 'Edit'.`
          );
        }
        throw err;
      }
    }

    const created = await sp.web.getFolderByServerRelativePath(targetUrl)();
    return created.ServerRelativeUrl;
  }

  static async getSiteMembers(): Promise<Array<{ id: string; displayName: string; email: string }>> {
    const sp = SetupService.sp();
    const groups: number[] = [];

    try {
      const membersGroup = await sp.web.associatedMemberGroup();
      if (membersGroup?.Id) groups.push(membersGroup.Id);
    } catch {
      /* pode não existir */
    }

    try {
      const ownersGroup = await sp.web.associatedOwnerGroup();
      if (ownersGroup?.Id) groups.push(ownersGroup.Id);
    } catch {
      /* pode não existir */
    }

    const results: Array<{ id: string; displayName: string; email: string }> = [];

    for (const groupId of groups) {
      try {
        const users = await sp.web.siteGroups.getById(groupId).users();
        for (const user of users) {
          const email = (user.Email || "").toLowerCase();
          if (!results.some(r => r.email.toLowerCase() === email)) {
            results.push({
              id: user.Id?.toString() || user.LoginName || user.Email || user.Title,
              displayName: user.Title,
              email: user.Email || ""
            });
          }
        }
      } catch {
        // ignora problemas com um grupo específico
      }
    }

    return results;
  }

  static async uploadFiles(context: WebPartContext, files: File[]): Promise<IFileAttachment[]> {
    const sp = SetupService.sp();
    const folderUrl = await this.ensurePageFolder(context);
    const uploads: IFileAttachment[] = [];

    for (const file of files) {
      if (file.size > 5 * 1024 * 1024) {
        throw new Error(`Ficheiro ${file.name} excede 5MB`);
      }
      try {
        await sp.web
          .getFolderByServerRelativePath(folderUrl)
          .files.addUsingPath(file.name, file, { Overwrite: true });

        const serverRelativeUrl = `${folderUrl.replace(/\/$/, "")}/${encodeURIComponent(file.name)}`;
        uploads.push({ name: file.name, serverRelativeUrl, size: file.size });
      } catch (err) {
        const spError = toSharePointError(err);
        if (spError.status === 403) {
          throw new Error(
            `Sem permissão para carregar ficheiros em "${folderUrl}". ` +
            `Verifica se tens 'Contribute/Edit' na biblioteca de documentos.`
          );
        }
        throw err;
      }
    }
    return uploads;
  }

  static async addMessage(msg: IChatMessage): Promise<number> {
    const sp = SetupService.sp();
    const list = sp.web.lists.getByTitle("Chat Messages");

    const result = await list.items.add({
      Title: msg.author.displayName,
      Message: msg.text,
      MentionsJson: JSON.stringify(msg.mentions || []),
      AttachmentsJson: JSON.stringify(msg.attachments || []),
      PageUniqueId: msg.pageUniqueId || "",
      PageName: msg.pageName || ""
    });

    return result.data?.Id as number;
  }

  static async getMessages(pageUniqueId: string, options?: { afterId?: number }): Promise<IChatMessage[]> {
    const sp = SetupService.sp();
    const escapedId = pageUniqueId.replace(/'/g, "''");
    const filterParts = [`PageUniqueId eq '${escapedId}'`];
    const afterId = options?.afterId;
    if (typeof afterId === "number" && afterId > 0) {
      filterParts.push(`Id gt ${afterId}`);
    }

    const items = await sp.web.lists
      .getByTitle("Chat Messages")
      .items.select(
        "Id,Title,Message,MentionsJson,AttachmentsJson,Created,PageUniqueId,PageName,Author/Id,Author/Title,Author/EMail"
      )
      .expand("Author")
      .filter(filterParts.join(" and "))
      .orderBy("Id", true)();

    return items.map((item: IChatListItem) => ({
      id: item.Id,
      text: item.Message,
      created: item.Created,
      author: {
        id: item.Author?.Id?.toString() ?? "",
        displayName: item.Author?.Title ?? "",
        email: item.Author?.EMail ?? ""
      },
      mentions: JSON.parse(item.MentionsJson || "[]"),
      attachments: JSON.parse(item.AttachmentsJson || "[]"),
      pageUniqueId: item.PageUniqueId,
      pageName: item.PageName
    }));
  }
}


