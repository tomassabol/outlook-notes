export interface Label {
  id: string;
  emoji: string;
  text: string;
  /** Outlook category palette name, e.g. "Red" (legacy: #rrggbb, mapped to the nearest) */
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

export const QUICK_EMOJI = ['🔥', '⭐', '✅', '⏳', '📌', '💰', '❗', '💬', '👀', '🗓️', '🚀', '🤝'];
