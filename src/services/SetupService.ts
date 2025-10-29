import { SPFx, spfi } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/items";
import "@pnp/sp/security";
import { PermissionKind } from "@pnp/sp/security";
import type { WebPartContext } from "@microsoft/sp-webpart-base";

export class SetupService {
  private static _sp: ReturnType<typeof spfi>;

  static init(context: WebPartContext) {
    this._sp = spfi().using(SPFx(context));
  }

  static sp() {
    if (!this._sp) throw new Error("PnPjs not initialized");
    return this._sp;
  }

  static async ensureList(): Promise<void> {
    const sp = this.sp();
    const listTitle = "Chat Messages";
    
    try {
      await sp.web.lists.getByTitle(listTitle)();
      // lista existe
      return;
    } catch {
      // lista não existe, criar
    }

    const canManage = await sp.web.currentUserHasPermissions(PermissionKind.ManageLists);
    if (!canManage) {
      throw new Error(`A lista "${listTitle}" não existe e o utilizador não tem permissões para criá-la. 
      Pede a um Owner para criar a lista ou dá permissão 'Manage Lists' ao teu utilizador.`
    );
    }
    
    const ensure = await sp.web.lists.ensure(listTitle, "", 100); // GenericList
    if (ensure.created) {
      const list = sp.web.lists.getByTitle(listTitle);
      await list.fields.addMultilineText("Message");
      await list.fields.addText("MentionsJson");
      await list.fields.addText("AttachmentsJson");
      await list.fields.addText("PageUniqueId");
      await list.fields.addText("PageName");
    }
  }
}
