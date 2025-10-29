import { AadHttpClient } from "@microsoft/sp-http";
import type { WebPartContext } from "@microsoft/sp-webpart-base";
import type { IUserMention } from "../models/IUserMention";

export class GraphService {
  private static _client: AadHttpClient;

  static async init(context: WebPartContext): Promise<void> {
    this._client = await context.aadHttpClientFactory.getClient("https://graph.microsoft.com");
  }

  static async sendMentionEmails(fromDisplayName: string, mentions: IUserMention[], messagePreview: string, messageUrl: string): Promise<void> {
    if (!mentions?.length) return;
    const toRecipients = mentions.map(m => ({ emailAddress: { address: m.email, name: m.displayName } }));
    const body = {
      message: {
        subject: `[Menção] ${fromDisplayName} mencionou-te no SharePoint`,
        body: {
          contentType: "HTML",
          content: `<p>Foste mencionado numa conversa:</p><blockquote>${messagePreview}</blockquote><p><a href="${messageUrl}">Abrir a conversa</a></p>`
        },
        toRecipients
      },
      saveToSentItems: "false"
    };
    await this._client.post("https://graph.microsoft.com/v1.0/me/sendMail", AadHttpClient.configurations.v1, {
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body)
    });
  }
}
