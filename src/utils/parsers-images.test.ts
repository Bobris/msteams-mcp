import { describe, expect, it } from 'vitest';
import { extractImages, isTeamsImageUrl } from './parsers-images.js';

const url = 'https://eu-api.asm.skype.com/v1/objects/0-weu-test/views/imgo';
describe('inline image discovery', () => {
  it('preserves ASM metadata before HTML stripping and deduplicates repeated images', () => {
    const html = `<p>hello</p><img src="${url}" itemtype="http://schema.skype.com/AMSImage" itemscope="png" width="450.25" height="250" alt="a &amp; b"><img src='${url}'>`;
    expect(extractImages(html)).toEqual([{ index: 0, url, width: 450.25, height: 250, alt: 'a & b', downloadable: true }]);
  });
  it('handles attribute ordering, numeric entities, unquoted attributes, and external images', () => {
    expect(extractImages(`<IMG ALT='screen &#x31; &#50;' HEIGHT='bad' SRC=${url}><img src="https://example.org/a?x=1&amp;y=2">`))
      .toEqual([
        { index: 0, url, alt: 'screen 1 2', width: undefined, height: undefined, downloadable: true },
        { index: 1, url: 'https://example.org/a?x=1&y=2', alt: undefined, width: undefined, height: undefined, downloadable: false },
      ]);
  });
  it('ignores emoticons, missing src, and plain text', () => {
    expect(extractImages(`hello<img alt='none'><img itemtype='http://schema.skype.com/Emoji' src='${url}'>`)).toEqual([]);
  });
  it.each([
    'http://eu-api.asm.skype.com/v1/objects/a/views/imgo',
    'https://eu-api.asm.skype.com.evil.org/v1/objects/a/views/imgo',
    'https://evil.org/v1/objects/a/views/imgo',
    'https://eu-api.asm.skype.com@evil.org/v1/objects/a/views/imgo',
    'https://user@eu-api.asm.skype.com/v1/objects/a/views/imgo',
    'https://eu-api.asm.skype.com:8443/v1/objects/a/views/imgo',
    'https://eu-api.asm.skype.com/v1/objects/a/views/imgo?redirect=evil',
    'https://eu-api.asm.skype.com/other',
    'file:///tmp/image', 'data:image/png;base64,abc',
  ])('rejects credential destinations outside the ASM image endpoint: %s', value => {
    expect(isTeamsImageUrl(value)).toBe(false);
  });
});
