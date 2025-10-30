import { SetupService } from "./SetupService";
import type { IChatMessage } from "../models/IChatMessage";
import type { IFileAttachment } from "../models/IFileAttachment";
import "@pnp/sp/webs";
import "@pnp/sp/items";
import "@pnp/sp/lists";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "@pnp/sp/site-groups/web";
import "@pnp/sp/site-users/web";

const ATTACHMENTS_LIBRARY_TITLE = "Anexos dos Chats";

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
      // não existe, criar
    }

    try {
      await sp.web.lists.ensure(ATTACHMENTS_LIBRARY_TITLE, "Biblioteca para anexos dos chats", 101);
      // ler RootFolder
      const created = await sp.web.lists
        .getByTitle(ATTACHMENTS_LIBRARY_TITLE)
        .select("RootFolder/ServerRelativeUrl")
        .expand("RootFolder")();
      return created.RootFolder.ServerRelativeUrl as string;
    } catch (e: any) {
      if (e?.status === 403) {
        throw new Error(
          `Sem permissão para criar a biblioteca "${ATTACHMENTS_LIBRARY_TITLE}". ` +
        `Garante que tens 'Manage Lists' no site (ou pede a um Owner para criá-la).`
        );
      }
      throw e;
    }
  }
  
  static async getPageInfo(context: any): Promise<{ pageName: string; pageUniqueId: string; }> {
  // 1) Páginas reais (Site Pages): usa título e UniqueId do item
  const li = context?.pageContext?.listItem;
  if (li?.title && li?.uniqueId) {
    return {
      pageName: String(li.title),
      pageUniqueId: String(li.uniqueId)
    };
  }

  // 2) Workbench ou páginas sem listItem: usa URL como chave estável + título amigável
  const path =
    context?.pageContext?.site?.serverRequestPath ||
    (typeof window !== "undefined" ? (window.location.pathname + window.location.search) : "/");

  const isWorkbench = /\/workbench\.aspx/i.test(path);
  const pageName =
    (isWorkbench ? "Workbench" :
     (context?.pageContext?.web?.title || (typeof document !== "undefined" ? document.title : "Página"))) || "Página";

  // usa o caminho completo como "unique id" textual (campo é texto, não tem problema)
  const pageUniqueId = path || pageName;

  return { pageName, pageUniqueId };
}


  

  static async ensurePageFolder(context: any): Promise<string> {
  const sp = SetupService.sp();
  const { pageName } = await this.getPageInfo(context);
  const libRoot = await this.ensureAttachmentsLibrary(); // << biblioteca dedicada
  const targetUrl = `${libRoot.replace(/\/$/, "")}/${pageName}`;

  // 1) Já existe?
  try {
    const info = await sp.web.getFolderByServerRelativePath(targetUrl)(); // IFolderInfo
    return info.ServerRelativeUrl;
  } catch (e: any) {
    if (e?.status && e.status !== 404) {
      if (e.status === 403) {
        throw new Error(
          `Sem permissão para aceder à pasta "${targetUrl}". ` +
          `Verifica se tens 'Edit' na biblioteca "${ATTACHMENTS_LIBRARY_TITLE}".`
        );
      }
      throw e;
    }
  }

  // 2) Criar e ler de novo
  try {
    const parent = sp.web.getFolderByServerRelativePath(libRoot);
    await parent.folders.addUsingPath(pageName);
  } catch (e: any) {
    if (!(e?.status === 409 || /already exists/i.test(e?.message || ""))) {
      if (e?.status === 403) {
        throw new Error(
          `Sem permissão para criar a pasta "${pageName}" em "${ATTACHMENTS_LIBRARY_TITLE}". ` +
          `Precisas de 'Edit'.`
        );
      }
      throw e;
    }
  }

  const created = await sp.web.getFolderByServerRelativePath(targetUrl)();
  return created.ServerRelativeUrl;
}






  static async getSiteMembers(): Promise<Array<{ id: string; displayName: string; email: string }>> {
  const sp = SetupService.sp();
  const groups: number[] = [];

  // Obtém IDs dos grupos associados (Members + Owners), se existirem
  try {
    const mg = await sp.web.associatedMemberGroup();
    if (mg?.Id) groups.push(mg.Id);
  } catch { /* pode não existir */ }

  try {
    const og = await sp.web.associatedOwnerGroup();
    if (og?.Id) groups.push(og.Id);
  } catch { /* pode não existir */ }

  const results: Array<{ id: string; displayName: string; email: string }> = [];

  for (const gid of groups) {
    try {
      const users = await sp.web.siteGroups.getById(gid).users();
      for (const u of users) {
        const email = (u.Email || "").toLowerCase();
        if (!results.some(r => r.email.toLowerCase() === email)) {
          results.push({
            id: u.Id?.toString() || u.LoginName || u.Email || u.Title,
            displayName: u.Title,
            email: u.Email || ""
          });
        }
      }
    } catch {
      // ignora problemas com um grupo específico
    }
  }

  return results;
}



  static async uploadFiles(context: any, files: File[]): Promise<IFileAttachment[]> {
  const sp = SetupService.sp();
  const folderUrl = await this.ensurePageFolder(context);
  const uploads: IFileAttachment[] = [];

  for (const f of files) {
    if (f.size > 5 * 1024 * 1024) throw new Error(`Ficheiro ${f.name} excede 5MB`);
    try {
      await sp.web
        .getFolderByServerRelativePath(folderUrl)
        .files.addUsingPath(f.name, f, { Overwrite: true });

      const serverRelativeUrl = `${folderUrl.replace(/\/$/, "")}/${encodeURIComponent(f.name)}`;
      uploads.push({ name: f.name, serverRelativeUrl, size: f.size });
    } catch (e: any) {
      if (e?.status === 403) {
        throw new Error(
          `Sem permissão para carregar ficheiros em "${folderUrl}".
          Verifica se tens 'Contribute/Edit' na biblioteca de documentos.`
        );
      }
      throw e;
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

  static async getMessages(pageUniqueId: string): Promise<IChatMessage[]> {
    const sp = SetupService.sp();
    const items = await sp.web.lists
    .getByTitle("Chat Messages")
    .items.select(
        "Id,Title,Message,MentionsJson,AttachmentsJson,Created,PageUniqueId,PageName,Author/Id,Author/Title,Author/EMail")
      .expand("Author")
      .filter(`PageUniqueId eq '${pageUniqueId}'`)
      .orderBy("Id", true)(); // false = descendente, true = ascendente

    return items.map((i: any) => ({
      id: i.Id,
      text: i.Message,
      created: i.Created,
      author: {
        id: i.Author?.Id?.toString() ?? "",
        displayName: i.Author?.Title ?? "",
        email: i.Author?.EMail ?? "",
      },
      mentions: JSON.parse(i.MentionsJson || "[]"),
      attachments: JSON.parse(i.AttachmentsJson || "[]"),
      pageUniqueId: i.PageUniqueId,
      pageName: i.PageName,
    }));
  }
}
