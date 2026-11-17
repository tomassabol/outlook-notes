export interface Label {
  id: string;
  emoji: string;
  text: string;
  /** #rrggbb */
  color: string;
}

/** Stored per conversation under key `c:<hash(convId)>`. Short field names save sync quota. */
export interface ConvData {
  /** note text */
  n?: string;
  /** label ids */
  l?: string[];
  /** subject snapshot (≤60 chars) */
  s: string;
  /** updated-at, epoch ms */
  u: number;
}

export interface ExportFile {
  app: 'outlook-notes';
  version: 1;
  exportedAt: string;
  labels: Label[];
  conversations: Record<string, ConvData>;
}

export type Theme = 'light' | 'dark';
export type EditorFocus = 'note' | 'labels';

export type Msg = { type: 'open-editor'; focus: EditorFocus };

export const LABEL_COLORS = [
  '#d13438', // red
  '#ca5010', // orange
  '#f2c94c', // yellow
  '#107c41', // green
  '#038387', // teal
  '#0f6cbd', // blue
  '#8764b8', // purple
  '#e3008c', // pink
  '#8e562e', // brown
  '#69797e', // gray
];

export const QUICK_EMOJI = ['🔥', '⭐', '✅', '⏳', '📌', '💰', '❗', '💬', '👀', '🗓️', '🚀', '🤝'];
