const allowedTags = new Set([
  "B",
  "STRONG",
  "I",
  "EM",
  "U",
  "P",
  "BR",
  "UL",
  "OL",
  "LI",
  "DIV",
  "SPAN",
  "A"
]);

const allowedAttrsByTag: Record<string, Set<string>> = {
  A: new Set(["href", "target", "rel"]),
  SPAN: new Set(["data-mention", "data-email", "class"])
};

function sanitizeElement(el: Element): void {
  if (!allowedTags.has(el.tagName)) {
    const parent = el.parentNode;
    if (!parent) {
      el.remove();
      return;
    }
    while (el.firstChild) {
      parent.insertBefore(el.firstChild, el);
    }
    parent.removeChild(el);
    return;
  }

  const allowedAttrs = allowedAttrsByTag[el.tagName] || new Set<string>();
  Array.from(el.attributes).forEach(attr => {
    if (!allowedAttrs.has(attr.name.toLowerCase())) {
      el.removeAttribute(attr.name);
    }
  });

  if (el.tagName === "A") {
    const href = el.getAttribute("href") || "";
    const safeHref = /^(https?:|mailto:|#|\/)/i.test(href);
    if (!safeHref) {
      el.removeAttribute("href");
    } else {
      el.setAttribute("target", "_blank");
      el.setAttribute("rel", "noopener noreferrer");
    }
  }

  if (el.tagName === "SPAN" && !el.hasAttribute("data-mention")) {
    el.removeAttribute("class");
  }

  Array.from(el.childNodes).forEach(child => {
    if (child.nodeType === Node.ELEMENT_NODE) {
      sanitizeElement(child as Element);
    } else if (child.nodeType === Node.COMMENT_NODE) {
      child.parentNode?.removeChild(child);
    }
  });
}

export function sanitizeRichText(html: string): string {
  if (!html) {
    return "";
  }
  if (typeof document === "undefined") {
    return html;
  }

  const wrapper = document.createElement("div");
  wrapper.innerHTML = html;

  Array.from(wrapper.childNodes).forEach(node => {
    if (node.nodeType === Node.ELEMENT_NODE) {
      sanitizeElement(node as Element);
    } else if (node.nodeType === Node.COMMENT_NODE) {
      wrapper.removeChild(node);
    }
  });

  return wrapper.innerHTML.trim();
}

export function getPlainTextFromHtml(html: string): string {
  if (!html) {
    return "";
  }
  if (typeof document === "undefined") {
    return html;
  }
  const wrapper = document.createElement("div");
  wrapper.innerHTML = html;
  return wrapper.textContent || "";
}

