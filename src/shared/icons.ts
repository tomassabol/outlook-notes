// Fluent System Icons (the icon set Outlook uses). Paths are extracted at build time and
// rebuilt with createElementNS, so no innerHTML is needed (Outlook may enforce Trusted Types).
import note16 from '@fluentui/svg-icons/icons/note_16_regular.svg?raw';
import note20 from '@fluentui/svg-icons/icons/note_20_regular.svg?raw';
import noteAdd16 from '@fluentui/svg-icons/icons/note_add_16_regular.svg?raw';
import noteAdd20 from '@fluentui/svg-icons/icons/note_add_20_regular.svg?raw';
import tag16 from '@fluentui/svg-icons/icons/tag_16_regular.svg?raw';
import tag20 from '@fluentui/svg-icons/icons/tag_20_regular.svg?raw';
import edit16 from '@fluentui/svg-icons/icons/edit_16_regular.svg?raw';
import chevronDown16 from '@fluentui/svg-icons/icons/chevron_down_16_regular.svg?raw';
import chevronUp16 from '@fluentui/svg-icons/icons/chevron_up_16_regular.svg?raw';
import dismiss20 from '@fluentui/svg-icons/icons/dismiss_20_regular.svg?raw';
import dismiss16 from '@fluentui/svg-icons/icons/dismiss_16_regular.svg?raw';
import add16 from '@fluentui/svg-icons/icons/add_16_regular.svg?raw';
import add20 from '@fluentui/svg-icons/icons/add_20_regular.svg?raw';
import checkmark12 from '@fluentui/svg-icons/icons/checkmark_12_regular.svg?raw';
import delete16 from '@fluentui/svg-icons/icons/delete_16_regular.svg?raw';
import delete20 from '@fluentui/svg-icons/icons/delete_20_regular.svg?raw';
import arrowUp16 from '@fluentui/svg-icons/icons/arrow_up_16_regular.svg?raw';
import arrowDown16 from '@fluentui/svg-icons/icons/arrow_down_16_regular.svg?raw';
import download20 from '@fluentui/svg-icons/icons/arrow_download_20_regular.svg?raw';
import upload20 from '@fluentui/svg-icons/icons/arrow_upload_20_regular.svg?raw';
import search20 from '@fluentui/svg-icons/icons/search_20_regular.svg?raw';
import info16 from '@fluentui/svg-icons/icons/info_16_regular.svg?raw';
import note24 from '@fluentui/svg-icons/icons/note_24_regular.svg?raw';

const RAW = {
  note16, note20, note24, noteAdd16, noteAdd20, tag16, tag20, edit16, chevronDown16, chevronUp16,
  dismiss20, dismiss16, add16, add20, checkmark12, delete16, delete20, arrowUp16, arrowDown16,
  download20, upload20, search20, info16,
};
export type IconName = keyof typeof RAW;

const NS = 'http://www.w3.org/2000/svg';
const parsed = new Map<IconName, { size: string; paths: string[] }>();

function parse(name: IconName) {
  let p = parsed.get(name);
  if (!p) {
    const raw = RAW[name];
    p = {
      size: /viewBox="0 0 (\d+) \d+"/.exec(raw)?.[1] ?? '20',
      paths: [...raw.matchAll(/\sd="([^"]+)"/g)].map((m) => m[1]),
    };
    parsed.set(name, p);
  }
  return p;
}

export function icon(name: IconName, cls = ''): SVGSVGElement {
  const { size, paths } = parse(name);
  const svg = document.createElementNS(NS, 'svg');
  svg.setAttribute('viewBox', `0 0 ${size} ${size}`);
  svg.setAttribute('width', size);
  svg.setAttribute('height', size);
  svg.setAttribute('aria-hidden', 'true');
  svg.setAttribute('focusable', 'false');
  svg.setAttribute('class', `icon ${cls}`.trim());
  for (const d of paths) {
    const path = document.createElementNS(NS, 'path');
    path.setAttribute('d', d);
    svg.append(path);
  }
  return svg;
}
