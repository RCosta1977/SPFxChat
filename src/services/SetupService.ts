import { SPFx, spfi } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/items";
import "@pnp/sp/security";
import "@pnp/sp/views";
import type { WebPartContext } from "@microsoft/sp-webpart-base";
import type { IView } from "@pnp/sp/views";

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

  const ensure = await sp.web.lists.ensure(listTitle, "", 100);
  const list = sp.web.lists.getByTitle(listTitle);

  if (ensure.created) {
    await list.fields.addMultilineText("Message");
    await list.fields.addText("MentionsJson");
    await list.fields.addText("AttachmentsJson");
    await list.fields.addText("PageUniqueId");
    await list.fields.addText("PageName");
  } else {
    const fields = await list.fields.select("InternalName")();
    const have = new Set(fields.map((f: any) => f.InternalName));
    if (!have.has("Message")) await list.fields.addMultilineText("Message");
    if (!have.has("MentionsJson")) await list.fields.addText("MentionsJson");
    if (!have.has("AttachmentsJson")) await list.fields.addText("AttachmentsJson");
    if (!have.has("PageUniqueId")) await list.fields.addText("PageUniqueId");
    if (!have.has("PageName")) await list.fields.addText("PageName");
  }

  // ❗ Sem parênteses aqui — queremos IView, não IViewInfo
  const defView: IView = list.defaultView;

  // Lê campos atuais da vista
  const current = await defView.fields();      // IViewFieldsResult
  const cols = current.Items as string[];
  const want = ["Message", "PageName", "Created"]; // Created é built-in

  for (const col of want) {
    if (!cols.includes(col)) {
      await defView.fields.add(col);          // adiciona à vista
    }
  }
  // (não é necessário defView.update() após fields.add())
}



}
