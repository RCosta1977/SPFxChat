import { AadHttpClient } from "@microsoft/sp-http";
import type { WebPartContext } from "@microsoft/sp-webpart-base";
import type { IUserMention } from "../models/IUserMention";

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

export class GraphService {
  private static _client: AadHttpClient;

  static async init(context: WebPartContext): Promise<void> {
    this._client = await context.aadHttpClientFactory.getClient("https://graph.microsoft.com");
  }

  static async sendMentionEmails(fromDisplayName: string, mentions: IUserMention[], messagePreview: string, messageUrl: string): Promise<void> {
    if (!mentions?.length) {
      return;
    }

    const toRecipients = mentions.map(m => ({ emailAddress: { address: m.email, name: m.displayName } }));
    const safePreview = escapeHtml(messagePreview || "");
    const safeUrl = messageUrl ? encodeURI(messageUrl) : "#";

    const body = {
      message: {
        subject: `[Mencao] ${fromDisplayName} mencionou-te no SharePoint`,
        body: {
          contentType: "HTML",
          content: `<p>Foste mencionado numa conversa:</p><blockquote>${safePreview}</blockquote><p><a href="${safeUrl}">Abrir a conversa</a></p>`
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

