/** Inline images live in message HTML, not in the shared-files index. */
export interface MessageImage {
  /** Zero-based index for teams_download_image. */
  index: number;
  url: string;
  alt?: string;
  width?: number;
  height?: number;
  /** Only supported Teams ASM URLs can be downloaded with session credentials. */
  downloadable: boolean;
}

/** Never send Teams credentials to a host or endpoint supplied by arbitrary HTML. */
export function isTeamsImageUrl(value: string): boolean {
  try {
    const url = new URL(value);
    return url.protocol === 'https:' && !url.username && !url.password && !url.port &&
      /^(?:[a-z0-9-]+\.)?asm\.skype\.com$/.test(url.hostname) &&
      /^\/v1\/objects\/[a-zA-Z0-9_-]+\/views\/img[a-zA-Z0-9_-]*$/.test(url.pathname) &&
      !url.search && !url.hash;
  } catch { return false; }
}

function decodeAttribute(value: string): string {
  return value.replace(/&(?:amp|quot|apos|lt|gt|#\d+|#x[\da-f]+);/gi, entity => {
    const named: Record<string, string> = { '&amp;': '&', '&quot;': '"', '&apos;': "'", '&lt;': '<', '&gt;': '>' };
    if (named[entity.toLowerCase()]) return named[entity.toLowerCase()];
    const code = entity[2].toLowerCase() === 'x' ? parseInt(entity.slice(3, -1), 16) : parseInt(entity.slice(2, -1), 10);
    return code > 0 && code <= 0x10ffff ? String.fromCodePoint(code) : entity;
  });
}

export function extractImages(html: string): MessageImage[] {
  const images: MessageImage[] = [];
  const seen = new Set<string>();
  for (const match of html.matchAll(/<img\b(?:[^>"']|"[^"]*"|'[^']*')*>/gi)) {
    const attributes: Record<string, string> = {};
    for (const attribute of match[0].matchAll(/([^\s=<>/]+)\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s>]+))/g)) {
      const key = attribute[1].toLowerCase();
      if (!(key in attributes)) attributes[key] = decodeAttribute(attribute[2] ?? attribute[3] ?? attribute[4]);
    }
    const url = attributes.src;
    if (!url || seen.has(url)) continue;
    // Emoticons are UI decoration, not chat attachments.
    if (attributes.itemtype?.endsWith('/Emoji') || attributes.itemtype?.endsWith('/Emoticon')) continue;
    seen.add(url);
    const dimension = (name: string) => {
      const value = Number(attributes[name]);
      return Number.isFinite(value) && value > 0 ? value : undefined;
    };
    images.push({ index: images.length, url, alt: attributes.alt || undefined,
      width: dimension('width'), height: dimension('height'), downloadable: isTeamsImageUrl(url) });
  }
  return images;
}
