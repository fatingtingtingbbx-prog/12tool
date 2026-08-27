import { get, set } from 'idb-keyval';
import { AppData, CardEntry } from '../types';

export const STORAGE_KEY = 'tavern_vault_data_v1';
export const THEME_KEY = 'tavern_vault_theme';
export const SIDEBAR_KEY = 'tavern_vault_sidebar';

const defaultAppData: AppData = {
  cards: [],
  groups: ['默认'],
  phoneLinks: [],
  themes: [],
  themeCategories: ['默认'],
  beautifications: [],
  beautificationCategories: ['默认'],
  presets: [],
  presetCategories: ['默认'],
  plugins: [],
  pluginCategories: ['默认'],
  normalCards: [],
  normalCardCategories: ['默认'],
  chatMemes: [],
  chatMemeCategories: ['默认'],
};

// In-memory cache for fast synchronous access.
let cachedAppData: AppData | null = null;
let saveQueue: Promise<boolean> = Promise.resolve(true);

// 大数据版本的关键优化：不再每次保存都把整个 AppData（可能包含大量图片/字体）
// 一次性 structured-clone 到 IndexedDB。每个顶层数据集合独立存储，只保存发生变化的集合。
const DATA_FIELDS: (keyof AppData)[] = [
  'cards', 'groups', 'phoneLinks', 'themes', 'themeCategories',
  'beautifications', 'beautificationCategories', 'presets', 'presetCategories',
  'plugins', 'pluginCategories', 'normalCards', 'normalCardCategories',
  'apis', 'apiCategories', 'fonts', 'fontCategories', 'extraStories',
  'extraStoryCategories', 'stickerPacks', 'stickerCategories', 'worldBooks',
  'worldBookCategories', 'chatMemes', 'chatMemeCategories',
];
const fieldKey = (field: keyof AppData) => `${STORAGE_KEY}::${String(field)}`;
const SPLIT_MARKER_KEY = `${STORAGE_KEY}::__split_v1`;

const mergeDefaults = (data: any): AppData => ({
  cards: Array.isArray(data?.cards) ? data.cards : [],
  groups: Array.isArray(data?.groups) && data.groups.length ? data.groups : ['默认'],
  phoneLinks: Array.isArray(data?.phoneLinks) ? data.phoneLinks : [],
  themes: Array.isArray(data?.themes) ? data.themes : [],
  themeCategories: Array.isArray(data?.themeCategories) ? data.themeCategories : ['默认'],
  beautifications: Array.isArray(data?.beautifications) ? data.beautifications : [],
  beautificationCategories: Array.isArray(data?.beautificationCategories) ? data.beautificationCategories : ['默认'],
  presets: Array.isArray(data?.presets) ? data.presets : [],
  presetCategories: Array.isArray(data?.presetCategories) ? data.presetCategories : ['默认'],
  plugins: Array.isArray(data?.plugins) ? data.plugins : [],
  pluginCategories: Array.isArray(data?.pluginCategories) ? data.pluginCategories : ['默认'],
  normalCards: Array.isArray(data?.normalCards) ? data.normalCards : [],
  normalCardCategories: Array.isArray(data?.normalCardCategories) ? data.normalCardCategories : ['默认'],
  apis: Array.isArray(data?.apis) ? data.apis : [],
  apiCategories: Array.isArray(data?.apiCategories) ? data.apiCategories : ['默认'],
  fonts: Array.isArray(data?.fonts) ? data.fonts : [],
  fontCategories: Array.isArray(data?.fontCategories) ? data.fontCategories : ['默认'],
  extraStories: Array.isArray(data?.extraStories) ? data.extraStories : [],
  extraStoryCategories: Array.isArray(data?.extraStoryCategories) ? data.extraStoryCategories : ['默认'],
  stickerPacks: Array.isArray(data?.stickerPacks) ? data.stickerPacks : [],
  stickerCategories: Array.isArray(data?.stickerCategories) ? data.stickerCategories : ['默认'],
  worldBooks: Array.isArray(data?.worldBooks) ? data.worldBooks : [],
  worldBookCategories: Array.isArray(data?.worldBookCategories) ? data.worldBookCategories : ['默认'],
  chatMemes: Array.isArray(data?.chatMemes) ? data.chatMemes : [],
  chatMemeCategories: Array.isArray(data?.chatMemeCategories) ? data.chatMemeCategories : ['默认'],
});

/** Load legacy root data and overlay any new per-field stores. */
export async function loadAppDataAsync(): Promise<AppData> {
  try {
    const splitReady = await get<boolean>(SPLIT_MARKER_KEY);
    let merged: AppData;
    if (splitReady) {
      // 新版本启动时只读取分片数据，避免再次 clone 整个巨型 AppData。
      const values = await Promise.all(DATA_FIELDS.map((field) => get<any>(fieldKey(field))));
      const partial: any = {};
      DATA_FIELDS.forEach((field, index) => {
        if (values[index] !== undefined) partial[field] = values[index];
      });
      merged = mergeDefaults(partial);
    } else {
      // 首次从旧版本迁移时读取一次旧根数据，随后立刻拆分保存；之后启动不再读取巨型根对象。
      const root = await get<AppData>(STORAGE_KEY);
      merged = mergeDefaults(root || {});
      for (const field of DATA_FIELDS) {
        await set(fieldKey(field), (merged as any)[field]);
        // 迁移大数据时主动让出主线程，避免移动端同时 clone 多个大数组。
        await new Promise((resolve) => setTimeout(resolve, 0));
      }
      await set(SPLIT_MARKER_KEY, true);
    }
    cachedAppData = merged;
    return merged;
  } catch (err) {
    console.warn('IndexedDB load fallback to LocalStorage', err);
  }
  const fallback = loadAppDataFromLocalStorage();
  cachedAppData = fallback;
  return fallback;
}

export function loadAppData(): AppData {
  if (cachedAppData) return cachedAppData;
  return loadAppDataFromLocalStorage();
}

function loadAppDataFromLocalStorage(): AppData {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return mergeDefaults({});
    return mergeDefaults(JSON.parse(raw));
  } catch (e) {
    console.error('Failed to load app data from localStorage', e);
    return mergeDefaults({});
  }
}

/**
 * Persist only changed top-level collections. This prevents large cards/fonts/stickers
 * from being cloned again when an unrelated setting is saved.
 */
export function saveAppData(data: AppData): Promise<boolean> {
  const previous = cachedAppData;
  cachedAppData = data;
  const changed = DATA_FIELDS.filter((field) => !previous || (previous as any)[field] !== (data as any)[field]);
  if (changed.length === 0) return Promise.resolve(true);

  // Serialize writes so rapid consecutive confirms cannot race each other.
  const writeBatch = async () => {
    try {
      for (const field of changed) {
        await set(fieldKey(field), (data as any)[field]);
        // 每个集合单独写入，且不同时 clone 多个大数组。
        await new Promise((resolve) => setTimeout(resolve, 0));
      }
      await set(SPLIT_MARKER_KEY, true);
      return true;
    } catch (err) {
      console.error('IndexedDB save error:', err);
      return false;
    }
  };
  saveQueue = saveQueue.then(writeBatch, writeBatch);
  return saveQueue;
}

/**
 * Image file quality compression to WebP Base64 to prevent storage bloating
 */
export async function processImageFile(file: File, maxWidth = 400, maxHeight = 600): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const img = new Image();
      img.onload = () => {
        const canvas = document.createElement('canvas');
        let width = img.width;
        let height = img.height;

        // Scale proportionally
        if (width > maxWidth) {
          height = Math.round((height * maxWidth) / width);
          width = maxWidth;
        }
        if (height > maxHeight) {
          width = Math.round((width * maxHeight) / height);
          height = maxHeight;
        }

        canvas.width = width;
        canvas.height = height;
        const ctx = canvas.getContext('2d');
        ctx?.drawImage(img, 0, 0, width, height);

        // Convert to webp with 0.8 quality
        const compressedBase64 = canvas.toDataURL('image/webp', 0.8);
        resolve(compressedBase64);
      };
      img.onerror = reject;
      img.src = e.target?.result as string;
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

export function getCardDisplayName(card: CardEntry): string {
  if (card.editHistory?.name) return card.editHistory.name;
  if (card.name) return card.name;
  const rd = card.rawData;
  if (rd?.data?.name) return rd.data.name;
  if (rd?.name) return rd.name;
  if (rd?.char_name) return rd.char_name;
  return '未命名角色';
}

export function getCardCreator(card: CardEntry): string {
  if (card.editHistory?.author) return card.editHistory.author;
  if (card.author) return card.author;
  const rd = card.rawData;
  if (rd?.data?.creator) return rd.data.creator;
  if (rd?.creator) return rd.creator;
  if (rd?.data?.extensions?.author) return rd.data.extensions.author;
  return '未知作者';
}

export function getCardDescription(card: CardEntry): string {
  if (card.editHistory?.description !== undefined) return card.editHistory.description;
  const rd = card.rawData;
  if (rd?.data?.description) return rd.data.description;
  if (rd?.description) return rd.description;
  return '';
}

export function getCardPersonality(card: CardEntry): string {
  if (card.editHistory?.personality !== undefined) return card.editHistory.personality;
  const rd = card.rawData;
  if (rd?.data?.personality) return rd.data.personality;
  if (rd?.personality) return rd.personality;
  return '';
}

export function getCardGreeting(card: CardEntry): string {
  if (card.editHistory?.first_mes !== undefined) return card.editHistory.first_mes;
  if (card.editHistory?.greeting !== undefined) return card.editHistory.greeting;
  const rd = card.rawData;
  if (rd?.data?.first_mes) return rd.data.first_mes;
  if (rd?.first_mes) return rd.first_mes;
  if (rd?.greeting) return rd.greeting;
  return '';
}

export function getCardAlternateGreetings(card: CardEntry): string[] {
  if (card.editHistory?.alternate_greetings && Array.isArray(card.editHistory.alternate_greetings)) {
    return card.editHistory.alternate_greetings;
  }
  const rd = card.rawData;
  if (rd?.data?.alternate_greetings) return rd.data.alternate_greetings;
  if (rd?.alternate_greetings) return rd.alternate_greetings;
  return [];
}

export function getCardWorldBook(card: CardEntry): any {
  if (card.editHistory?.character_book) return card.editHistory.character_book;
  const rd = card.rawData;
  if (rd?.data?.character_book) return rd.data.character_book;
  if (rd?.character_book) return rd.character_book;
  if (rd?.extensions?.character_book) return rd.extensions.character_book;
  if (rd?.data?.extensions?.character_book) return rd.data.extensions.character_book;
  return null;
}

export function getCardRegex(card: CardEntry | null): any[] {
  if (!card) return [];
  if (card.editHistory?.regex_scripts && Array.isArray(card.editHistory.regex_scripts)) {
    return card.editHistory.regex_scripts;
  }
  const rd = card.rawData;
  if (!rd) return [];
  if (Array.isArray(rd.data?.extensions?.regex_scripts)) return rd.data.extensions.regex_scripts;
  if (Array.isArray(rd.extensions?.regex_scripts)) return rd.extensions.regex_scripts;
  if (Array.isArray(rd.data?.extensions?.regex)) return rd.data.extensions.regex;
  if (Array.isArray(rd.extensions?.regex)) return rd.extensions.regex;
  if (Array.isArray(rd.data?.regex_scripts)) return rd.data.regex_scripts;
  if (Array.isArray(rd.regex_scripts)) return rd.regex_scripts;
  if (Array.isArray(rd.data?.regexes)) return rd.data.regexes;
  if (Array.isArray(rd.regexes)) return rd.regexes;
  if (Array.isArray(rd.data?.user_regexes)) return rd.data.user_regexes;
  if (Array.isArray(rd.user_regexes)) return rd.user_regexes;
  if (Array.isArray(rd.data?.character_book?.extensions?.regex_scripts)) return rd.data.character_book.extensions.regex_scripts;
  if (Array.isArray(rd.character_book?.extensions?.regex_scripts)) return rd.character_book.extensions.regex_scripts;
  if (rd.findRegex || rd.find_regex || rd.pattern || rd.data?.findRegex || rd.data?.find_regex) {
    return [rd.data || rd];
  }
  return [];
}

export function getCardTags(card: CardEntry): string[] {
  const rd = card.rawData;
  if (rd?.data?.tags) return rd.data.tags;
  if (rd?.tags) return rd.tags;
  return [];
}

/**
 * Returns merged card data by combining rawData and editHistory
 */
export function getCurrentCardData(card: CardEntry): any {
  const data = JSON.parse(JSON.stringify(card.rawData));
  if (!card.edited || !card.editHistory) return data;
  const eh = card.editHistory;

  if (eh.name !== undefined) {
    if (data.data) data.data.name = eh.name; else data.name = eh.name;
  }
  if (eh.description !== undefined) {
    if (data.data) data.data.description = eh.description; else data.description = eh.description;
  }
  if (eh.personality !== undefined) {
    if (data.data) data.data.personality = eh.personality; else data.personality = eh.personality;
  }
  if (eh.scenario !== undefined) {
    if (data.data) data.data.scenario = eh.scenario; else data.scenario = eh.scenario;
  }
  if (eh.system_prompt !== undefined) {
    if (data.data) data.data.system_prompt = eh.system_prompt; else data.system_prompt = eh.system_prompt;
  }
  if (eh.first_mes !== undefined) {
    if (data.data) data.data.first_mes = eh.first_mes; else data.first_mes = eh.first_mes;
  }
  if (eh.alternate_greetings !== undefined) {
    if (data.data) data.data.alternate_greetings = eh.alternate_greetings; else data.alternate_greetings = eh.alternate_greetings;
  }
  if (eh.character_book !== undefined) {
    if (data.data) data.data.character_book = eh.character_book; else data.character_book = eh.character_book;
  }
  if (eh.regex_scripts !== undefined) {
    const target = data.data || data;
    if (!target.extensions) target.extensions = {};
    target.extensions.regex_scripts = eh.regex_scripts;
  }
  if (eh.author !== undefined) {
    if (data.data) {
      if (!data.data.extensions) data.data.extensions = {};
      data.data.extensions.author = eh.author;
    } else {
      data.author = eh.author;
    }
  }
  if (eh.tags !== undefined) {
    if (data.data) data.data.tags = eh.tags; else data.tags = eh.tags;
  }

  return data;
}
