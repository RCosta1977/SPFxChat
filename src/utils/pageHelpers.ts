import type { WebPartContext } from "@microsoft/sp-webpart-base";

export function getPageDeepLink(context: WebPartContext): string {
  const base = context.pageContext?.web?.absoluteUrl || window.location.origin.replace(/\/$/, "");
  const path = context.pageContext?.site?.serverRequestPath || window.location.pathname + window.location.search;
  return base + path;
}
