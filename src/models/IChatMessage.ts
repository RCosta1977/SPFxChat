import { IUserMention } from "./IUserMention";
import { IFileAttachment } from "./IFileAttachment";

export interface IChatMessage {
  id?: number;
  text: string;
  created: string;             // ISO
  author: IUserMention;
  mentions: IUserMention[];
  attachments: IFileAttachment[];
  pageUniqueId?: string;
  pageName?: string;
}
