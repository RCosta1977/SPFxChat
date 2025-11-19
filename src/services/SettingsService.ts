import { SetupService } from "./SetupService";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/items";
import type { IView } from "@pnp/sp/views";

type SettingItem = {
  Value?: string;
};

export class SettingsService {
  static async ensureSettingsList(): Promise<void> {
    const sp = SetupService.sp();
    const listTitle = "Chat Settings";

    const ensure = await sp.web.lists.ensure(listTitle, "Settings for SPFx Chat webpart", 100);
    const list = sp.web.lists.getByTitle(listTitle);

    if (ensure.created) {
      await list.fields.addText("Value");
      await list.fields.addText("PageUniqueId");
    } else {
      type FieldInfo = { InternalName: string };
      const fields = (await list.fields.select("InternalName")()) as FieldInfo[];
      const have = new Set(fields.map(f => f.InternalName));
      if (!have.has("Value")) await list.fields.addText("Value");
      if (!have.has("PageUniqueId")) await list.fields.addText("PageUniqueId");
    }

    try {
      const defView: IView = list.defaultView;
      const current = await defView.fields();
      const cols = current.Items as string[];
      for (const col of ["Title", "Value", "PageUniqueId"]) {
        if (!cols.includes(col)) {
          await defView.fields.add(col);
        }
      }
    } catch {
      // ignore view updates
    }
  }

  static async getSetting(key: string, pageUniqueId?: string): Promise<string | undefined> {
    const sp = SetupService.sp();
    const list = sp.web.lists.getByTitle("Chat Settings");

    try {
      // 1) Page-scoped override
      if (pageUniqueId) {
        const escapedKey = key.replace(/'/g, "''");
        const escapedId = pageUniqueId.replace(/'/g, "''");
        const items = await list.items
          .select("Id,Title,Value,PageUniqueId")
          .filter(`Title eq '${escapedKey}' and PageUniqueId eq '${escapedId}'`)
          .top(1)();
        if (items && items.length) {
          const candidate = (items as SettingItem[])[0]?.Value;
          return typeof candidate === "string" ? candidate : undefined;
        }
      }

      // 2) Global fallback (no PageUniqueId)
      const escapedKey2 = key.replace(/'/g, "''");
      const items2 = await list.items
        .select("Id,Title,Value,PageUniqueId")
        .filter(`Title eq '${escapedKey2}' and (PageUniqueId eq null or PageUniqueId eq '')`)
        .top(1)();
      if (items2 && items2.length) {
        const candidate = (items2 as SettingItem[])[0]?.Value;
        return typeof candidate === "string" ? candidate : undefined;
      }
      return undefined;
    } catch {
      // If list is missing or permission denied, just return undefined
      return undefined;
    }
  }
}
