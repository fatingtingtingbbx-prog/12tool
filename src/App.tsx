import React, { useState, useEffect, useRef } from 'react';
import mammoth from 'mammoth';
import { Document, Packer, Paragraph, TextRun } from 'docx';
import { unzipSync, strFromU8 } from 'fflate';
import {
  Menu, X, Sun, Moon, Search, Plus, Trash2, FolderPlus, Edit3, Download, Upload,
  Maximize2, ChevronDown, ChevronUp, Check, Image as ImageIcon, Tag, Folder, CheckSquare, Square,
  MoreHorizontal, RefreshCw, FileText, CheckCircle2, Circle, ArrowRightLeft, Move, Copy, Sliders, ZoomIn, FileCode,
  Save, Dices, QrCode, ArrowUpDown, Sparkles, Link2
} from 'lucide-react';
import { CardEntry, CardAssociation, AppData, PhoneLink, ThemeEntry, PresetEntry, PluginEntry, NormalCardEntry, ApiEntry, ApiKeyItem, FontEntry, ExtraStoryEntry, StickerItem, StickerPackEntry, WorldBookEntry, ChatMemeEntry } from './types';
import {
  loadAppData, loadAppDataAsync, saveAppData, getCardDisplayName, getCardCreator, getCardDescription,
  getCardPersonality, getCardGreeting, getCardAlternateGreetings, getCardWorldBook,
  getCardRegex, getCardTags, getCurrentCardData
} from './utils/storage';
import { parseCardFile } from './utils/cardParser';
import { getPureCardDataForExport, injectPngTextChunk, scaleImage, processImageFile, fileToDataURL } from './utils/pngParser';

export function estimateTokens(text: string | null | undefined): number {
  if (!text) return 0;
  const str = String(text);
  if (!str.trim()) return 0;
  const cjkMatches = str.match(/[\u4e00-\u9fa5\u3040-\u30ff\uac00-\ud7af]/g) || [];
  const cjkCount = cjkMatches.length;
  const nonCjkStr = str.replace(/[\u4e00-\u9fa5\u3040-\u30ff\uac00-\ud7af]/g, ' ');
  const wordMatches = nonCjkStr.match(/\w+|[^\w\s]/g) || [];
  const nonCjkTokenCount = wordMatches.length;
  return Math.ceil(cjkCount * 1.35 + nonCjkTokenCount);
}

export function sortItemList<T>(
  items: T[],
  sortOrder: 'default' | 'az' | 'za' | 'newest' | 'oldest',
  getName: (item: T) => string,
  getCreatedAt?: (item: T) => number
): T[] {
  if (sortOrder === 'default') return items;
  const list = [...items];
  if (sortOrder === 'az') {
    return list.sort((a, b) => getName(a).localeCompare(getName(b), 'zh-CN', { numeric: true, sensitivity: 'base' }));
  }
  if (sortOrder === 'za') {
    return list.sort((a, b) => getName(b).localeCompare(getName(a), 'zh-CN', { numeric: true, sensitivity: 'base' }));
  }
  if (sortOrder === 'newest') {
    return list.sort((a, b) => ((getCreatedAt ? getCreatedAt(b) : 0) - (getCreatedAt ? getCreatedAt(a) : 0)));
  }
  if (sortOrder === 'oldest') {
    return list.sort((a, b) => ((getCreatedAt ? getCreatedAt(a) : 0) - (getCreatedAt ? getCreatedAt(b) : 0)));
  }
  return list;
}

export default function App() {
  // App State
  const [appData, setAppData] = useState<AppData>(() => loadAppData());
  const appDataDirtyRef = useRef(false);
  const appDataHydratedRef = useRef(false);
  const [appDataHydrated, setAppDataHydrated] = useState(false);

  // Hydrate from IndexedDB on startup. If the user has already edited/uploaded
  // something before the async read finishes, never overwrite that newer state.
  useEffect(() => {
    let cancelled = false;
    loadAppDataAsync().then((asyncData) => {
      if (cancelled) return;
      appDataHydratedRef.current = true;
      if (!appDataDirtyRef.current && asyncData) {
        setAppData(asyncData);
      }
      setAppDataHydrated(true);
    });
    return () => { cancelled = true; };
  }, []);
  const [theme, setTheme] = useState<'light' | 'dark'>(() => {
    return (localStorage.getItem('tavern_vault_theme') as 'light' | 'dark') || 'light';
  });
  const [sidebarOpen, setSidebarOpen] = useState(false); // Default collapsed
  const [currentPage, setCurrentPage] = useState('st-cards');

  // Search, Filter & Grouping
  const [searchQuery, setSearchQuery] = useState('');
  const [currentGroup, setCurrentGroup] = useState('全部'); // '全部' shows all cards
  const [batchMode, setBatchMode] = useState(false);
  const [selectedCardIds, setSelectedCardIds] = useState<string[]>([]);
  // 角色卡分页显示：存有成百上千张卡时，一次性把全部卡片渲染进 DOM 会导致页面严重卡顿，
  // 这里只渲染当前可见的一部分，滚动到底部自动加载更多。
  const CARD_PAGE_SIZE = 90;
  const [cardVisibleCount, setCardVisibleCount] = useState(CARD_PAGE_SIZE);
  const cardLoadMoreRef = useRef<HTMLDivElement>(null);

  // Modals State
  const [detailCardId, setDetailCardId] = useState<string | null>(null);
  const [detailTab, setDetailTab] = useState<'overview' | 'personality' | 'greetings' | 'worldbook' | 'regex' | 'qr' | 'raw' | 'extras'>('overview');

  // Sorting States across sections
  const [cardSortOrder, setCardSortOrder] = useState<'default' | 'az' | 'za' | 'newest' | 'oldest'>('default');
  const [themeSortOrder, setThemeSortOrder] = useState<'default' | 'az' | 'za' | 'newest' | 'oldest'>('default');
  const [presetSortOrder, setPresetSortOrder] = useState<'default' | 'az' | 'za' | 'newest' | 'oldest'>('default');
  const [pluginSortOrder, setPluginSortOrder] = useState<'default' | 'az' | 'za' | 'newest' | 'oldest'>('default');
  const [normalCardSortOrder, setNormalCardSortOrder] = useState<'default' | 'az' | 'za' | 'newest' | 'oldest'>('default');
  const [worldBookSortOrder, setWorldBookSortOrder] = useState<'default' | 'az' | 'za' | 'newest' | 'oldest'>('default');
  const [extraStorySortOrder, setExtraStorySortOrder] = useState<'default' | 'az' | 'za' | 'newest' | 'oldest'>('default');
  const [fontSortOrder, setFontSortOrder] = useState<'default' | 'az' | 'za' | 'newest' | 'oldest'>('default');
  const [apiSortOrder, setApiSortOrder] = useState<'default' | 'az' | 'za' | 'newest' | 'oldest'>('default');
  const [beautificationSortOrder, setBeautificationSortOrder] = useState<'default' | 'az' | 'za' | 'newest' | 'oldest'>('default');

  // Random Card Gacha State
  const [gachaCard, setGachaCard] = useState<CardEntry | null>(null);
  const [isGachaSpinning, setIsGachaSpinning] = useState(false);

  // Add Alternate Greeting Modal State
  const [showAddAltGreetingModal, setShowAddAltGreetingModal] = useState(false);
  const [newAltGreetingInputText, setNewAltGreetingInputText] = useState('');
  // ST 角色卡开场白自定义用户名：仅用于界面预览/编辑显示，导出和底层数据始终保留 {{user}}。
  const [customGreetingUsername, setCustomGreetingUsername] = useState('');

  const applyCustomGreetingUsername = (text: string, username = customGreetingUsername) => {
    if (!username.trim()) return text;
    return String(text || '').replace(/\{\{user\}\}/gi, username);
  };

  const restoreCustomGreetingUsername = (text: string, username = customGreetingUsername) => {
    const name = username.trim();
    if (!name) return text;
    return String(text || '').split(name).join('{{user}}');
  };

  useEffect(() => {
    setCustomGreetingUsername('');
  }, [detailCardId]);

  // QR Modal & Input Ref
  const qrFileInputRef = useRef<HTMLInputElement>(null);
  const cardQrFileInputRef = useRef<HTMLInputElement>(null);
  const cardWorldBookImportFileInputRef = useRef<HTMLInputElement>(null);
  const [showQrImportModal, setShowQrImportModal] = useState(false);
  const [qrInputText, setQrInputText] = useState('');
  const [qrSearchQuery, setQrSearchQuery] = useState('');
  const [cardRegexSearchQuery, setCardRegexSearchQuery] = useState('');
  const [cardSectionImportModal, setCardSectionImportModal] = useState<'worldbook' | 'regex' | 'qr' | null>(null);
  const [cardSectionImportMode, setCardSectionImportMode] = useState<'single' | 'all'>('single');
  const [cardSectionImportPendingFile, setCardSectionImportPendingFile] = useState<File | null>(null);
  const [editingCardRegex, setEditingCardRegex] = useState<{ index: number; scriptName: string; findRegex: string; replaceString: string } | null>(null);

  // ST 角色卡关联状态：仅允许作者 + 角色名同时相同的卡片互相关联
  const [associationTargetId, setAssociationTargetId] = useState('');
  const [associationNote, setAssociationNote] = useState('');
  const [associationPrimary, setAssociationPrimary] = useState(false);
  const [editingQrItem, setEditingQrItem] = useState<{ index: number; item: any } | null>(null);
  const [cardWorldBookSearchQuery, setCardWorldBookSearchQuery] = useState('');
  const [jsonWorldBookSearchQuery, setJsonWorldBookSearchQuery] = useState('');
  const [showJsonWorldBookImportPreview, setShowJsonWorldBookImportPreview] = useState(false);
  const [pendingJsonWorldBook, setPendingJsonWorldBook] = useState<any | null>(null);
  const [normalZipPreview, setNormalZipPreview] = useState<{ fileName: string; files: { name: string; size: number; content: string; selected: boolean }[] } | null>(null);
  // 移动端批量上传队列：部分 Android 文件选择器会忽略 multiple，允许连续选择后一次性导入。
  const [batchUploadKind, setBatchUploadKind] = useState<'stickers' | 'worldbook' | 'extras' | 'fonts' | null>(null);
  const [batchUploadFiles, setBatchUploadFiles] = useState<File[]>([]);
  const batchUploadFileInputRef = useRef<HTMLInputElement>(null);
  const [isExportingBackup, setIsExportingBackup] = useState(false);
  const [isImportingBackup, setIsImportingBackup] = useState(false);
  // 当前板块独立导入/导出
  const sectionImportFileInputRef = useRef<HTMLInputElement>(null);
  const [isImportingSection, setIsImportingSection] = useState(false);

  const openBatchUpload = (kind: 'stickers' | 'worldbook' | 'extras' | 'fonts') => {
    setBatchUploadKind(kind);
    setBatchUploadFiles([]);
  };

  const batchUploadAccept: Record<'stickers' | 'worldbook' | 'extras' | 'fonts', string> = {
    stickers: '.docx,.txt,.json',
    worldbook: '.docx,.txt,.json',
    extras: '.docx,.txt',
    fonts: '.ttf,.otf,.woff,.woff2',
  };

  const handleBatchUploadFileSelect = (e: React.ChangeEvent<HTMLInputElement>) => {
    const selected = Array.from(e.target.files || []) as File[];
    if (selected.length) {
      setBatchUploadFiles((prev) => {
        const merged = [...prev];
        for (const file of selected) {
          const exists = merged.some((x) => x.name === file.name && x.size === file.size && x.lastModified === file.lastModified);
          if (!exists) merged.push(file);
        }
        return merged;
      });
    }
    e.target.value = '';
  };

  const removeBatchUploadFile = (index: number) => {
    setBatchUploadFiles((prev) => prev.filter((_, i) => i !== index));
  };

  const confirmBatchUpload = async () => {
    const files = [...batchUploadFiles];
    const kind = batchUploadKind;
    if (!kind || !files.length) return;
    setBatchUploadKind(null);
    setBatchUploadFiles([]);
    if (kind === 'fonts') await processFontFiles(files);
    else if (kind === 'extras') await processExtraStoryFiles(files);
    else if (kind === 'stickers') await processStickerFiles(files);
    else if (kind === 'worldbook') {
      for (const file of files) {
        await processWorldBookFileImport(file);
        if (file.name.toLowerCase().endsWith('.json')) break;
        await new Promise((resolve) => setTimeout(resolve, 0));
      }
    }
  };

  const handleDrawRandomCard = () => {
    if (appData.cards.length === 0) {
      showToast('当前还没有任何角色卡哦！', 'error');
      return;
    }
    setIsGachaSpinning(true);
    setTimeout(() => {
      const randomIndex = Math.floor(Math.random() * appData.cards.length);
      setGachaCard(appData.cards[randomIndex]);
      setIsGachaSpinning(false);
    }, 300);
  };

  const downloadJsonFile = (fileName: string, data: any) => {
    const jsonStr = JSON.stringify(data, null, 2);
    const blob = new Blob([jsonStr], { type: 'application/json;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName.endsWith('.json') ? fileName : `${fileName}.json`;
    a.click();
    URL.revokeObjectURL(url);
  };

  // 导出文件命名统一使用 toolbox-YYYY-MM-DD 样式。
  const formatDateForFileName = (date: Date = new Date()) => {
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, '0');
    const d = String(date.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  };

  /**
   * 将大数组拆成若干个较小的 JSON.stringify 片段，拼成合法的 JSON 数组文本片段列表。
   * 避免在导出成百上千张角色卡（含大量 base64 封面图）时，对整个数组一次性调用
   * JSON.stringify 拼出一个巨大的单一字符串——那样极易触发 "Invalid string length"
   * 或占用双倍以上内存，导致页面卡死甚至崩溃。
   */
  const arrayToJsonParts = (arr: any[] | undefined, chunkSize = 150): string[] => {
    if (!Array.isArray(arr) || arr.length === 0) return ['[]'];
    const parts: string[] = ['['];
    for (let i = 0; i < arr.length; i += chunkSize) {
      const chunk = arr.slice(i, i + chunkSize);
      const chunkJson = JSON.stringify(chunk);
      // 去掉这一小段自身的 [ ]，只保留内部内容，再用逗号和其他片段拼接。
      const inner = chunkJson.slice(1, -1);
      if (i > 0) parts.push(',');
      parts.push(inner);
    }
    parts.push(']');
    return parts;
  };

  // 数组型字段超过这个长度时才分片处理，小分类列表直接整体 stringify 即可。
  const LARGE_ARRAY_THRESHOLD = 150;

  const handleExportFullBackup = async () => {
    if (isExportingBackup) return;
    setIsExportingBackup(true);
    showToast('正在生成完整备份，数据量较大时可能需要几秒钟…', 'info');
    // 让 toast 先绘制出来，避免一开始就同步卡住主线程给人"没反应/崩溃"的错觉。
    await new Promise((resolve) => setTimeout(resolve, 30));

    try {
      // 侧边栏“全部数据”备份不包含字体，避免把体积很大的字体文件一起打包；
      // 字体板块仍可通过顶部的独立“导入/导出”按钮单独备份。
      const keys = (Object.keys(appData).filter((key) => key !== 'fonts' && key !== 'fontCategories')) as (keyof AppData)[];
      const parts: BlobPart[] = ['{'];

      for (let idx = 0; idx < keys.length; idx++) {
        const key = keys[idx];
        const value = (appData as any)[key];
        parts.push(`${JSON.stringify(key)}:`);

        if (Array.isArray(value) && value.length > LARGE_ARRAY_THRESHOLD) {
          // 大数组分片 stringify，避免单次拼出体积巨大的字符串（尤其是含有大量
          // base64 封面图的 cards 字段）。
          parts.push(...arrayToJsonParts(value));
          // 每处理完一个大字段就让出一次主线程，防止长时间无响应。
          await new Promise((resolve) => setTimeout(resolve, 0));
        } else {
          parts.push(JSON.stringify(value));
        }

        if (idx < keys.length - 1) parts.push(',');
      }
      parts.push('}');

      const blob = new Blob(parts, { type: 'application/json;charset=utf-8' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `toolbox-${formatDateForFileName()}.json`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
      showToast('全部数据导出成功（不含字体）！', 'success');
    } catch (err: any) {
      console.error('Export full backup failed', err);
      showToast(
        '导出失败：数据量过大或设备内存不足。可尝试删除部分体积较大的角色卡封面/附件后重试（' +
          (err?.message || String(err)) +
          '）',
        'error'
      );
    } finally {
      setIsExportingBackup(false);
    }
  };

  const normalizeRegexScriptsFromJson = (json: any): any[] => {
    if (Array.isArray(json)) return json;
    if (json && typeof json === 'object') {
      if (Array.isArray(json.regex_scripts)) return json.regex_scripts;
      if (Array.isArray(json.regexes)) return json.regexes;
      if (Array.isArray(json.extensions?.regex_scripts)) return json.extensions.regex_scripts;
      if (Array.isArray(json.data?.extensions?.regex_scripts)) return json.data.extensions.regex_scripts;
      if (json.findRegex || json.find_regex || json.pattern || json.scriptName || json.script_name || json.name) return [json];
    }
    return [];
  };

  const normalizeWorldBookEntriesFromJson = (json: any): { book: any; entries: any[] } => {
    // 兼容常见世界书 JSON：
    // 1) character_book.entries / data.character_book.entries
    // 2) entries 数组
    // 3) SillyTavern 风格的 entries 对象：{ "0": {...}, "1": {...} }
    // 4) data.entries / world_info.entries 等嵌套结构
    // 5) 单条 entry 或直接传入 entry 数组
    const asEntryArray = (value: any): any[] => {
      if (Array.isArray(value)) return value.filter((v) => v && typeof v === 'object');
      if (value && typeof value === 'object') {
        // 世界书常见的 keyed entries 对象，Object.values 会保持 JSON 中的键顺序。
        const values = Object.values(value);
        if (values.length && values.every((v: any) => v && typeof v === 'object')) return values;
      }
      return [];
    };

    const looksLikeEntry = (value: any) => !!(value && typeof value === 'object' && (
      'content' in value || 'comment' in value || 'key' in value || 'keys' in value ||
      'keysecondary' in value || 'uid' in value || 'position' in value || 'constant' in value
    ));

    const findEntries = (root: any): { book: any; entries: any[] } | null => {
      if (!root || typeof root !== 'object') return null;

      const candidates = [
        root?.character_book,
        root?.data?.character_book,
        root?.world_book,
        root?.data?.world_book,
        root?.world_info,
        root?.data?.world_info,
        root,
      ].filter(Boolean);

      for (const candidate of candidates) {
        if (!candidate || typeof candidate !== 'object') continue;
        for (const key of ['entries', 'worldEntries', 'world_entries', 'items']) {
          if (candidate[key] !== undefined) {
            const entries = asEntryArray(candidate[key]);
            if (entries.length) return { book: candidate, entries };
          }
        }
      }

      // 有些导出会把 entries 再包一层 data / book。
      for (const nested of [root?.data, root?.book, root?.data?.book]) {
        if (!nested || typeof nested !== 'object') continue;
        const result = findEntries(nested);
        if (result?.entries?.length) return result;
      }

      if (looksLikeEntry(root)) return { book: { entries: [root] }, entries: [root] };
      return null;
    };

    if (Array.isArray(json)) return { book: { entries: json }, entries: json };

    const found = findEntries(json);
    if (found) return found;

    // 最后的兼容：某些文件直接是 {"0": entry, "1": entry}。
    if (json && typeof json === 'object') {
      const values = asEntryArray(json);
      if (values.length && values.some(looksLikeEntry)) return { book: { entries: values }, entries: values };
    }

    return { book: json && typeof json === 'object' ? json : { entries: [] }, entries: [] };
  };

  const normalizeQrDocument = (json: any) => {
    if (!json || typeof json !== 'object' || Array.isArray(json) || !Array.isArray(json.qrList)) {
      throw new Error('不是有效的 QR JSON：缺少 qrList 数组');
    }
    if (json.version !== undefined && Number(json.version) !== 2) {
      throw new Error(`暂不支持 QR version ${json.version}`);
    }
    return {
      ...json,
      version: Number(json.version || 2),
      name: json.name || '未命名 QR',
      qrList: json.qrList.map((item: any, index: number) => ({
        ...item,
        id: item?.id ?? index + 1,
      })),
      idIndex: Number(json.idIndex || json.qrList.length || 0),
    };
  };

  const saveCardQrDocument = (doc: any) => {
    if (!activeDetailCard) return;
    const normalized = normalizeQrDocument(doc);
    const updatedCards = appData.cards.map((c) =>
      c.id === activeDetailCard.id
        ? { ...c, qrData: normalized, edited: true, updatedAt: Date.now() }
        : c
    );
    updateAppData({ ...appData, cards: updatedCards });
  };

  const addManualQr = () => {
    const current = activeDetailCard?.qrData ? normalizeQrDocument(activeDetailCard.qrData) : { version: 2, name: '', qrList: [], idIndex: 0 };
    const nextId = Math.max(0, ...(current.qrList || []).map((x: any) => Number(x.id) || 0)) + 1;
    const item = {
      id: nextId,
      showLabel: false,
      label: '新 QR',
      title: '',
      message: '',
      contextList: [],
      preventAutoExecute: true,
      isHidden: false,
      executeOnStartup: false,
      executeOnUser: false,
      executeOnAi: false,
      executeOnChatChange: false,
      executeOnGroupMemberDraft: false,
      executeOnNewChat: false,
      executeBeforeGeneration: false,
      automationId: '',
    };
    const next = { ...current, qrList: [...current.qrList, item], idIndex: nextId };
    saveCardQrDocument(next);
    setEditingQrItem({ index: next.qrList.length - 1, item });
  };

  const saveEditedQrItem = () => {
    if (!activeDetailCard || !editingQrItem) return;
    const current = normalizeQrDocument(activeDetailCard.qrData || { version: 2, name: '', qrList: [], idIndex: 0 });
    const qrList = [...current.qrList];
    qrList[editingQrItem.index] = editingQrItem.item;
    saveCardQrDocument({ ...current, qrList });
    setEditingQrItem(null);
    showToast('QR 已保存', 'success');
  };

  const deleteCardQrItem = (index: number) => {
    if (!activeDetailCard) return;
    const current = normalizeQrDocument(activeDetailCard.qrData || { version: 2, name: '', qrList: [], idIndex: 0 });
    if (!current.qrList[index]) return;
    const label = current.qrList[index]?.label || current.qrList[index]?.title || `QR #${index + 1}`;
    const qrList = current.qrList.filter((_: any, i: number) => i !== index);
    saveCardQrDocument({ ...current, qrList });
    if (editingQrItem?.index === index) setEditingQrItem(null);
    showToast(`已删除 QR：${label}`, 'success');
  };

  const saveCardRegexList = (list: any[]) => {
    if (!activeDetailCard) return;
    const updatedCards = appData.cards.map((c) =>
      c.id === activeDetailCard.id
        ? { ...c, editHistory: { ...c.editHistory, regex_scripts: list }, edited: true, updatedAt: Date.now() }
        : c
    );
    updateAppData({ ...appData, cards: updatedCards });
  };

  const saveCardWorldBook = (book: any) => {
    if (!activeDetailCard) return;
    const updatedCards = appData.cards.map((c) =>
      c.id === activeDetailCard.id
        ? { ...c, editHistory: { ...c.editHistory, character_book: book }, edited: true, updatedAt: Date.now() }
        : c
    );
    updateAppData({ ...appData, cards: updatedCards });
  };


  const importQrJsonText = (text: string, mode: 'single' | 'all' = 'all') => {
    try {
      const parsed = normalizeQrDocument(JSON.parse(text));
      if (!activeDetailCard) {
        showToast('请先打开角色卡详情', 'error');
        return;
      }
      const current = activeDetailCard.qrData && typeof activeDetailCard.qrData === 'object'
        ? normalizeQrDocument(activeDetailCard.qrData)
        : null;
      const next = mode === 'all'
        ? parsed
        : {
            ...(current || parsed),
            name: current?.name || parsed.name,
            qrList: [...(current?.qrList || []), parsed.qrList[0]].filter(Boolean),
            idIndex: Math.max(Number(current?.idIndex || 0), Number(parsed.qrList[0]?.id || 0)),
          };
      saveCardQrDocument(next);
      setShowQrImportModal(false);
      setQrInputText('');
      showToast(`QR ${mode === 'all' ? '全部覆盖' : '单条追加'}成功，共 ${next.qrList.length} 条`, 'success');
    } catch (err: any) {
      showToast('QR JSON 导入失败：' + (err?.message || '格式错误'), 'error');
    }
  };

  const handleQrFileImport = async (file: File, mode: 'single' | 'all' = 'all') => {
    try {
      const text = await file.text();
      importQrJsonText(text, mode);
    } catch (err: any) {
      showToast('读取 QR 文件失败：' + (err?.message || '未知错误'), 'error');
    }
  };

  const handleCardSectionFileImport = async (file: File, section: 'worldbook' | 'regex' | 'qr', mode: 'single' | 'all') => {
    try {
      const text = await file.text();
      const json = JSON.parse(text);
      if (section === 'worldbook') {
        const normalized = normalizeWorldBookEntriesFromJson(json);
        if (!normalized.entries.length) throw new Error('未找到世界书 entries');
        const current = JSON.parse(JSON.stringify(getCardWorldBook(activeDetailCard) || { entries: [] }));
        const next = mode === 'all' ? normalized.book : { ...(current || normalized.book), entries: [...(current.entries || []), normalized.entries[0]] };
        saveCardWorldBook(next);
        showToast(`世界书 ${mode === 'all' ? '全部覆盖' : '单条追加'}成功，共 ${(next.entries || []).length} 条`, 'success');
      } else if (section === 'regex') {
        const normalized = normalizeRegexScriptsFromJson(json).map((rx) => {
          const name = rx.scriptName || rx.script_name || rx.name || rx.title || '导入正则';
          const find = rx.findRegex || rx.find_regex || rx.pattern || rx.find || rx.regex || '';
          const replace = rx.replaceString || rx.replace_string || rx.replacement || rx.replace || '';
          return { ...rx, scriptName: name, script_name: name, name, findRegex: find, find_regex: find, pattern: find, replaceString: replace, replace_string: replace, replacement: replace };
        });
        if (!normalized.length) throw new Error('未找到正则条目');
        const current = JSON.parse(JSON.stringify(getCardRegex(activeDetailCard) || []));
        const next = mode === 'all' ? normalized : [...current, normalized[0]];
        saveCardRegexList(next);
        showToast(`正则 ${mode === 'all' ? '全部覆盖' : '单条追加'}成功，共 ${next.length} 条`, 'success');
      } else {
        const parsed = normalizeQrDocument(json);
        const current = activeDetailCard?.qrData && typeof activeDetailCard.qrData === 'object' ? normalizeQrDocument(activeDetailCard.qrData) : null;
        const next = mode === 'all' ? parsed : { ...(current || parsed), name: current?.name || parsed.name, qrList: [...(current?.qrList || []), parsed.qrList[0]].filter(Boolean), idIndex: Math.max(Number(current?.idIndex || 0), Number(parsed.qrList[0]?.id || 0)) };
        saveCardQrDocument(next);
        showToast(`QR ${mode === 'all' ? '全部覆盖' : '单条追加'}成功，共 ${next.qrList.length} 条`, 'success');
      }
    } catch (err: any) {
      showToast(`导入失败：${err?.message || 'JSON 格式错误'}`, 'error');
    }
  };

  const handleCardSectionFileSelected = async (e: React.ChangeEvent<HTMLInputElement>, section: 'worldbook' | 'regex' | 'qr') => {
    const file = e.target.files?.[0];
    if (!file) return;
    await handleCardSectionFileImport(file, section, cardSectionImportMode);
    e.target.value = '';
    setCardSectionImportPendingFile(null);
  };

  // New Group Modal
  const [showNewGroupModal, setShowNewGroupModal] = useState(false);
  const [newGroupName, setNewGroupName] = useState('');

  // Manage Group Modal (Delete / Rename Group)
  const [managingGroup, setManagingGroup] = useState<string | null>(null);
  const [showDeleteGroupConfirm, setShowDeleteGroupConfirm] = useState(false);
  const [deleteCardsWithGroup, setDeleteCardsWithGroup] = useState(false);
  const [renameGroupInput, setRenameGroupInput] = useState('');

  // Batch Move Modal
  const [showBatchMoveModal, setShowBatchMoveModal] = useState(false);
  const [batchTargetGroup, setBatchTargetGroup] = useState('');

  // Fullscreen Viewer Modal
  const [fullscreenData, setFullscreenData] = useState<{ title: string; content: string; type?: 'text' | 'json' } | null>(null);

  // Edit Text Modal (Generic)
  const [editModalData, setEditModalData] = useState<{
    title: string;
    initialValue: string;
    type: 'input' | 'textarea';
    onSave: (val: string) => void;
  } | null>(null);

  // Toast Notification
  const [toasts, setToasts] = useState<{ id: string; msg: string; type: 'success' | 'error' | 'info' }[]>([]);

  // File Inputs Refs
  const uploadFileInputRef = useRef<HTMLInputElement>(null);
  const importBackupInputRef = useRef<HTMLInputElement>(null);
  const updateCardInputRef = useRef<HTMLInputElement>(null);
  const coverUploadInputRef = useRef<HTMLInputElement>(null);
  const authorNoteInputRef = useRef<HTMLInputElement>(null);
  const memoryInputRef = useRef<HTMLInputElement>(null);
  const cardRegexFileInputRef = useRef<HTMLInputElement>(null);
  const beautificationDocumentFileInputRef = useRef<HTMLInputElement>(null);
  const beautificationUpdateInputRef = useRef<HTMLInputElement>(null);

  // Long Press Timer for Group Tabs
  const longPressTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);

  // Batch Selection in Extras Tab
  const [authorNoteBatchMode, setAuthorNoteBatchMode] = useState(false);
  const [selectedAuthorNoteIndices, setSelectedAuthorNoteIndices] = useState<number[]>([]);

  const [memoryBatchMode, setMemoryBatchMode] = useState(false);
  const [selectedMemoryIndices, setSelectedMemoryIndices] = useState<number[]>([]);

  // Mobile Links State
  const [showAddPhoneModal, setShowAddPhoneModal] = useState(false);
  const [phoneForm, setPhoneForm] = useState({ name: '', url: '', contact: '', description: '' });

  const [editingPhoneLink, setEditingPhoneLink] = useState<PhoneLink | null>(null);

  const [phoneSearchQuery, setPhoneSearchQuery] = useState('');
  const [phoneBatchMode, setPhoneBatchMode] = useState(false);
  const [selectedPhoneIds, setSelectedPhoneIds] = useState<string[]>([]);

  // Mobile Links Handlers
  const phoneLinksList = appData.phoneLinks || [];
  const filteredPhoneLinks = phoneLinksList.filter((item) => {
    if (!phoneSearchQuery.trim()) return true;
    const q = phoneSearchQuery.toLowerCase();
    return (
      item.name.toLowerCase().includes(q) ||
      item.url.toLowerCase().includes(q) ||
      (item.contact && item.contact.toLowerCase().includes(q)) ||
      (item.description && item.description.toLowerCase().includes(q))
    );
  });

  const handleAddPhoneLink = () => {
    if (!phoneForm.name.trim()) {
      showToast('请填写小手机名称', 'error');
      return;
    }
    if (!phoneForm.url.trim()) {
      showToast('请填写链接', 'error');
      return;
    }
    const newLink: PhoneLink = {
      id: Date.now().toString(36) + Math.random().toString(36).substring(2, 6),
      name: phoneForm.name.trim(),
      url: phoneForm.url.trim(),
      contact: phoneForm.contact.trim() || undefined,
      description: phoneForm.description.trim() || undefined,
      createdAt: Date.now(),
      updatedAt: Date.now(),
    };

    const updatedLinks = [newLink, ...(appData.phoneLinks || [])];
    updateAppData({ ...appData, phoneLinks: updatedLinks });
    setShowAddPhoneModal(false);
    setPhoneForm({ name: '', url: '', contact: '', description: '' });
    showToast('小手机链接添加成功！', 'success');
  };

  const handleSaveEditedPhoneLink = () => {
    if (!editingPhoneLink) return;
    if (!editingPhoneLink.name.trim()) {
      showToast('请填写小手机名称', 'error');
      return;
    }
    if (!editingPhoneLink.url.trim()) {
      showToast('请填写链接', 'error');
      return;
    }

    const updatedLinks = (appData.phoneLinks || []).map((item) =>
      item.id === editingPhoneLink.id
        ? {
            ...editingPhoneLink,
            name: editingPhoneLink.name.trim(),
            url: editingPhoneLink.url.trim(),
            contact: editingPhoneLink.contact?.trim() || undefined,
            description: editingPhoneLink.description?.trim() || undefined,
            updatedAt: Date.now(),
          }
        : item
    );

    updateAppData({ ...appData, phoneLinks: updatedLinks });
    setEditingPhoneLink(null);
    showToast('保存成功！', 'success');
  };

  const handleDeleteSinglePhoneLink = (id: string) => {
    if (!confirm('确定要删除此小手机链接存档吗？')) return;
    const updatedLinks = (appData.phoneLinks || []).filter((item) => item.id !== id);
    updateAppData({ ...appData, phoneLinks: updatedLinks });
    setEditingPhoneLink(null);
    showToast('已删除存档', 'info');
  };

  const handleBatchDeletePhoneLinks = () => {
    if (selectedPhoneIds.length === 0) return;
    if (!confirm(`确定要删除选中的 ${selectedPhoneIds.length} 个小手机链接吗？`)) return;
    const updatedLinks = (appData.phoneLinks || []).filter(
      (item) => !selectedPhoneIds.includes(item.id)
    );
    updateAppData({ ...appData, phoneLinks: updatedLinks });
    setSelectedPhoneIds([]);
    setPhoneBatchMode(false);
    showToast(`已批量删除 ${selectedPhoneIds.length} 个小手机链接`, 'info');
  };

  // ST Themes State & Refs
  const themeFileInputRef = useRef<HTMLInputElement>(null);
  const themeDocumentFileInputRef = useRef<HTMLInputElement>(null);
  const themeCoverInputRef = useRef<HTMLInputElement>(null);

  const [editingTheme, setEditingTheme] = useState<ThemeEntry | null>(null);
  const [themeDetailTab, setThemeDetailTab] = useState<'info' | 'code'>('info');
  const [isThemeCodeExpanded, setIsThemeCodeExpanded] = useState(false);

  const [themeSearchQuery, setThemeSearchQuery] = useState('');
  const [themeCategoryFilter, setThemeCategoryFilter] = useState('全部分类');
  const [themeBatchMode, setThemeBatchMode] = useState(false);
  const [selectedThemeIds, setSelectedThemeIds] = useState<string[]>([]);
  const [codeSearchQuery, setCodeSearchQuery] = useState('');

  // ST Theme Group Modals & State
  const [showNewThemeGroupModal, setShowNewThemeGroupModal] = useState(false);
  const [newThemeGroupName, setNewThemeGroupName] = useState('');

  const [managingThemeGroup, setManagingThemeGroup] = useState<string | null>(null);
  const [renameThemeGroupInput, setRenameThemeGroupInput] = useState('');
  const [deleteThemesWithGroup, setDeleteThemesWithGroup] = useState(false);

  const [showThemeBatchMoveModal, setShowThemeBatchMoveModal] = useState(false);
  const [batchTargetThemeGroup, setBatchTargetThemeGroup] = useState('');

  // ST Themes Handlers
  const themesList = appData.themes || [];
  const availableThemeCategories = Array.from(
    new Set(['全部分类', '默认', ...(appData.themeCategories || []), ...themesList.map((t) => t.category || '默认')])
  );

  const filteredThemes = themesList.filter((item) => {
    if (themeCategoryFilter !== '全部分类' && (item.category || '默认') !== themeCategoryFilter) {
      return false;
    }
    if (!themeSearchQuery.trim()) return true;
    const q = themeSearchQuery.toLowerCase();
    const rawStr = item.content || item.rawJsonString || (typeof item.jsonData === 'string' ? item.jsonData : JSON.stringify(item.jsonData || {}));
    return (
      item.name.toLowerCase().includes(q) ||
      item.fileName.toLowerCase().includes(q) ||
      (item.author && item.author.toLowerCase().includes(q)) ||
      (item.category && item.category.toLowerCase().includes(q)) ||
      (item.type && item.type.toLowerCase().includes(q)) ||
      (item.source && item.source.toLowerCase().includes(q)) ||
      (item.description && item.description.toLowerCase().includes(q)) ||
      (rawStr && rawStr.toLowerCase().includes(q))
    );
  });

  const handleThemeDocumentFileUpload = async (files: FileList) => {
    const fileArray = Array.from(files);
    if (fileArray.length === 0) return;

    const newThemes: ThemeEntry[] = [];
    const showProgress = fileArray.length > 8;
    if (showProgress) showToast(`开始导入 ${fileArray.length} 个文件…`, 'info');

    // 顺序处理并在每个文件之间让出主线程，避免一次性 forEach(async) 并发解析
    // 大量 docx/图片文件时占用过多内存导致页面卡顿甚至崩溃。
    for (let i = 0; i < fileArray.length; i++) {
      const file = fileArray[i];
      const ext = file.name.split('.').pop()?.toLowerCase() || 'txt';
      const nameClean = file.name.replace(/\.[^/.]+$/, '') || '未命名美化';

      try {
        let contentText = '';
        let parsedJson: any = null;
        let themeType = 'ST主题';

        if (ext === 'docx') {
          themeType = 'ST主题';
          try {
            const arrayBuffer = await file.arrayBuffer();
            const result = await mammoth.extractRawText({ arrayBuffer });
            contentText = result.value || '';
          } catch (err) {
            contentText = '';
          }
        } else {
          contentText = await file.text();
          if (ext === 'json') {
            try {
              parsedJson = JSON.parse(contentText);
              // ST 主题与美化分类完全独立：这里不再根据 JSON 的 type 识别线上/线下美化。
            } catch (e) {
              // keep as string content
            }
          } else if (ext === 'css') {
            themeType = 'CSS样式';
          }
        }

        const themeItem: ThemeEntry = {
          id: Date.now().toString(36) + Math.random().toString(36).substring(2, 6),
          name: nameClean,
          fileName: file.name,
          author: (parsedJson && (parsedJson.author || parsedJson.creator)) || '默认',
          category: themeCategoryFilter !== '全部分类' ? themeCategoryFilter : ((parsedJson && parsedJson.category) || '默认'),
          type: themeType,
          source: (parsedJson && (parsedJson.source || parsedJson.url)) || '',
          description: (parsedJson && (parsedJson.description || parsedJson.comment)) || '',
          coverImage: null,
          jsonData: parsedJson || {},
          rawJsonString: parsedJson ? JSON.stringify(parsedJson, null, 2) : contentText,
          content: contentText,
          fileType: ext,
          createdAt: Date.now(),
          updatedAt: Date.now(),
        };

        newThemes.push(themeItem);
      } catch (e: any) {
        showToast(`读取文件 ${file.name} 失败`, 'error');
      }

      if (showProgress && (i + 1) % 5 === 0) showToast(`正在导入 ${i + 1}/${fileArray.length} 个文件…`, 'info');
      await new Promise((resolve) => setTimeout(resolve, 0));
    }

    if (newThemes.length > 0) {
      updateAppData((prev) => ({ ...prev, themes: [...newThemes, ...(prev.themes || [])] }));
      showToast(`成功导入 ${newThemes.length} 个美化文档！`, 'success');
    }
  };

  const handleThemeFileUpload = (files: FileList) => {
    handleThemeDocumentFileUpload(files);
  };

  const exportThemeDocument = (
    themeItem: Partial<ThemeEntry> & { name: string },
    format: 'docx' | 'txt' | 'json' | 'css'
  ) => {
    try {
      const content = themeItem.content || themeItem.rawJsonString || (typeof themeItem.jsonData === 'string' ? themeItem.jsonData : JSON.stringify(themeItem.jsonData || {}, null, 2)) || '';
      
      if (format === 'docx') {
        const paragraphs = (content || '').split('\n').map((line) => new Paragraph({
          children: [new TextRun({ text: line, size: 22 })],
        }));
        const doc = new Document({
          sections: [{ properties: {}, children: paragraphs }],
        });
        Packer.toBlob(doc).then((blob) => {
          const url = URL.createObjectURL(blob);
          const a = document.createElement('a');
          a.href = url;
          a.download = `${themeItem.name || '美化文档'}.docx`;
          a.click();
          URL.revokeObjectURL(url);
          showToast('已成功导出 .docx 格式文档！', 'success');
        }).catch(() => {
          showToast('导出 .docx 失败', 'error');
        });
        return;
      }

      let mimeType = 'text/plain;charset=utf-8';
      let ext = format;
      if (format === 'json') mimeType = 'application/json;charset=utf-8';
      if (format === 'css') mimeType = 'text/css;charset=utf-8';

      const blob = new Blob([content], { type: mimeType });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `${themeItem.name || '美化文档'}.${ext}`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      showToast(`已成功导出 .${ext} 格式文档！`, 'success');
    } catch (err: any) {
      showToast('导出失败: ' + err.message, 'error');
    }
  };

  const exportThemeAsJson = (themeItem: ThemeEntry) => {
    exportThemeDocument(themeItem, 'json');
  };

  const handleSaveEditedTheme = () => {
    if (!editingTheme) return;
    if (!editingTheme.name.trim()) {
      showToast('请填写美化名称', 'error');
      return;
    }

    let finalJsonData = editingTheme.jsonData;
    let finalRawStr = editingTheme.rawJsonString || editingTheme.content || '';

    if (editingTheme.fileType === 'json' && editingTheme.rawJsonString) {
      try {
        finalJsonData = JSON.parse(editingTheme.rawJsonString);
      } catch (e: any) {
        // retain string
      }
    }

    const updatedThemes = (appData.themes || []).map((t) =>
      t.id === editingTheme.id
        ? {
            ...editingTheme,
            name: editingTheme.name.trim(),
            author: editingTheme.author?.trim() || '默认',
            type: editingTheme.type?.trim() || 'ST主题',
            category: editingTheme.category?.trim() || '默认',
            source: editingTheme.source?.trim() || '',
            description: editingTheme.description?.trim() || '',
            content: editingTheme.content || editingTheme.rawJsonString || '',
            jsonData: finalJsonData,
            rawJsonString: finalRawStr,
            updatedAt: Date.now(),
          }
        : t
    );

    const cat = editingTheme.category?.trim() || '默认';
    const existingCats = appData.themeCategories || ['默认'];
    const newCats = existingCats.includes(cat) ? existingCats : [...existingCats, cat];

    updateAppData({ ...appData, themes: updatedThemes, themeCategories: newCats });
    setEditingTheme(null);
    showToast('美化存档已保存！', 'success');
  };

  const handleDeleteSingleTheme = (id: string) => {
    if (!confirm('确定要删除此主题美化存档吗？')) return;
    const updatedThemes = (appData.themes || []).filter((t) => t.id !== id);
    updateAppData({ ...appData, themes: updatedThemes });
    setEditingTheme(null);
    showToast('已删除主题存档', 'info');
  };

  const handleBatchDeleteThemes = () => {
    if (selectedThemeIds.length === 0) return;
    if (!confirm(`确定要删除选中的 ${selectedThemeIds.length} 个主题美化吗？`)) return;
    const updatedThemes = (appData.themes || []).filter((t) => !selectedThemeIds.includes(t.id));
    updateAppData({ ...appData, themes: updatedThemes });
    setSelectedThemeIds([]);
    setThemeBatchMode(false);
    showToast(`已批量删除 ${selectedThemeIds.length} 个主题`, 'info');
  };

  const handleCreateNewThemeGroup = () => {
    const name = newThemeGroupName.trim();
    if (!name) {
      showToast('请输入分组名称', 'error');
      return;
    }
    const currentCats = appData.themeCategories || ['默认'];
    if (currentCats.includes(name)) {
      showToast('该分组已存在', 'info');
      setThemeCategoryFilter(name);
      setShowNewThemeGroupModal(false);
      setNewThemeGroupName('');
      return;
    }
    const updatedCats = [...currentCats, name];
    updateAppData({ ...appData, themeCategories: updatedCats });
    setThemeCategoryFilter(name);
    setShowNewThemeGroupModal(false);
    setNewThemeGroupName('');
    showToast(`成功新建分组: ${name}`, 'success');
  };

  const handleRenameThemeGroup = () => {
    if (!managingThemeGroup) return;
    const newName = renameThemeGroupInput.trim();
    if (!newName) {
      showToast('请输入新的分组名称', 'error');
      return;
    }
    if (newName === managingThemeGroup) {
      setManagingThemeGroup(null);
      return;
    }

    const currentCats = appData.themeCategories || ['默认'];
    const updatedCats = currentCats.map((c) => (c === managingThemeGroup ? newName : c));

    const updatedThemes = (appData.themes || []).map((t) =>
      (t.category || '默认') === managingThemeGroup ? { ...t, category: newName } : t
    );

    updateAppData({
      ...appData,
      themeCategories: Array.from(new Set(updatedCats)),
      themes: updatedThemes,
    });

    if (themeCategoryFilter === managingThemeGroup) {
      setThemeCategoryFilter(newName);
    }

    setManagingThemeGroup(null);
    showToast('分组重命名成功', 'success');
  };

  const handleConfirmDeleteThemeGroup = () => {
    if (!managingThemeGroup) return;
    const targetGroup = managingThemeGroup;

    let updatedThemes = appData.themes || [];
    if (deleteThemesWithGroup) {
      updatedThemes = updatedThemes.filter((t) => (t.category || '默认') !== targetGroup);
    } else {
      updatedThemes = updatedThemes.map((t) =>
        (t.category || '默认') === targetGroup ? { ...t, category: '默认' } : t
      );
    }

    const currentCats = appData.themeCategories || ['默认'];
    const updatedCats = currentCats.filter((c) => c !== targetGroup);

    updateAppData({
      ...appData,
      themeCategories: updatedCats,
      themes: updatedThemes,
    });

    if (themeCategoryFilter === targetGroup) {
      setThemeCategoryFilter('全部分类');
    }

    setManagingThemeGroup(null);
    setDeleteThemesWithGroup(false);
    showToast(`分组 “${targetGroup}” 已删除`, 'info');
  };

  const handleBatchMoveThemes = () => {
    if (selectedThemeIds.length === 0) return;
    const target = batchTargetThemeGroup.trim();
    if (!target) {
      showToast('请选择或输入目标分组', 'error');
      return;
    }

    const updatedThemes = (appData.themes || []).map((t) =>
      selectedThemeIds.includes(t.id) ? { ...t, category: target } : t
    );

    const existingCats = appData.themeCategories || ['默认'];
    const newCats = existingCats.includes(target) ? existingCats : [...existingCats, target];

    updateAppData({
      ...appData,
      themes: updatedThemes,
      themeCategories: newCats,
    });

    setSelectedThemeIds([]);
    setThemeBatchMode(false);
    setShowThemeBatchMoveModal(false);
    setBatchTargetThemeGroup('');
    showToast(`已将 ${selectedThemeIds.length} 个主题移动到 “${target}”`, 'success');
  };

  // ==================== BEAUTIFICATIONS (美化) STATE & HANDLERS ====================
  const beautificationCoverInputRef = useRef<HTMLInputElement>(null);

  const [editingBeautification, setEditingBeautification] = useState<ThemeEntry | null>(null);
  const [beautificationDetailTab, setBeautificationDetailTab] = useState<'preview' | 'code'>('preview');
  const [isBeautificationCodeExpanded, setIsBeautificationCodeExpanded] = useState(false);

  const [beautificationSearchQuery, setBeautificationSearchQuery] = useState('');
  const [beautificationCategoryFilter, setBeautificationCategoryFilter] = useState('全部分类');
  const [beautificationBatchMode, setBeautificationBatchMode] = useState(false);
  const [selectedBeautificationIds, setSelectedBeautificationIds] = useState<string[]>([]);

  // Beautification Group Modals & State
  const [showNewBeautificationGroupModal, setShowNewBeautificationGroupModal] = useState(false);
  const [newBeautificationGroupName, setNewBeautificationGroupName] = useState('');

  const [managingBeautificationGroup, setManagingBeautificationGroup] = useState<string | null>(null);
  const [renameBeautificationGroupInput, setRenameBeautificationGroupInput] = useState('');
  const [deleteBeautificationsWithGroup, setDeleteBeautificationsWithGroup] = useState(false);

  const [showBeautificationBatchMoveModal, setShowBeautificationBatchMoveModal] = useState(false);
  const [batchTargetBeautificationGroup, setBatchTargetBeautificationGroup] = useState('');

  const beautificationsList = appData.beautifications || [];

  const filteredBeautifications = beautificationsList.filter((item) => {
    if (beautificationCategoryFilter !== '全部分类' && (item.category || '默认') !== beautificationCategoryFilter) {
      return false;
    }
    if (!beautificationSearchQuery.trim()) return true;
    const q = beautificationSearchQuery.toLowerCase();
    const rawStr = item.content || item.rawJsonString || (typeof item.jsonData === 'string' ? item.jsonData : JSON.stringify(item.jsonData || {}));
    return (
      item.name.toLowerCase().includes(q) ||
      item.fileName.toLowerCase().includes(q) ||
      (item.author && item.author.toLowerCase().includes(q)) ||
      (item.category && item.category.toLowerCase().includes(q)) ||
      (item.type && item.type.toLowerCase().includes(q)) ||
      (item.source && item.source.toLowerCase().includes(q)) ||
      (item.description && item.description.toLowerCase().includes(q)) ||
      (rawStr && rawStr.toLowerCase().includes(q))
    );
  });

  const handleBeautificationDocumentFileUpload = async (files: FileList) => {
    const fileArray = Array.from(files);
    if (fileArray.length === 0) return;

    const newBeautifications: ThemeEntry[] = [];
    const showProgress = fileArray.length > 8;
    if (showProgress) showToast(`开始导入 ${fileArray.length} 个文件…`, 'info');

    // 顺序处理并在每个文件之间让出主线程，避免一次性并发解析大量 docx/图片
    // 文件时占用过多内存导致页面卡顿甚至崩溃。
    for (let i = 0; i < fileArray.length; i++) {
      const file = fileArray[i];
      const ext = file.name.split('.').pop()?.toLowerCase() || 'txt';
      const nameClean = file.name.replace(/\.[^/.]+$/, '') || '未命名美化';

      try {
        let contentText = '';
        let parsedJson: any = null;
        let themeType = '线上主题';
        let coverImgUrl: string | null = null;

        if (ext === 'png' || ext === 'jpg' || ext === 'jpeg' || ext === 'webp' || file.type.startsWith('image/')) {
          themeType = 'PNG图片';
          coverImgUrl = await new Promise<string>((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = () => resolve(reader.result as string);
            reader.onerror = (err) => reject(err);
            reader.readAsDataURL(file);
          });
          contentText = coverImgUrl;
        } else if (ext === 'docx') {
          themeType = '线上主题';
          try {
            const arrayBuffer = await file.arrayBuffer();
            const result = await mammoth.extractRawText({ arrayBuffer });
            contentText = result.value || '';
          } catch (err) {
            contentText = '';
          }
        } else {
          contentText = await file.text();
          if (ext === 'json') {
            try {
              parsedJson = JSON.parse(contentText);
              if (parsedJson.type) themeType = parsedJson.type;
            } catch (e) {
              // keep as string content
            }
          } else if (ext === 'css') {
            themeType = 'CSS样式';
          }
        }

        const beautificationItem: ThemeEntry = {
          id: Date.now().toString(36) + Math.random().toString(36).substring(2, 6),
          name: nameClean,
          fileName: file.name,
          author: (parsedJson && (parsedJson.author || parsedJson.creator)) || '默认',
          category: beautificationCategoryFilter !== '全部分类' ? beautificationCategoryFilter : ((parsedJson && parsedJson.category) || '默认'),
          type: themeType,
          source: (parsedJson && (parsedJson.source || parsedJson.url)) || '',
          description: (parsedJson && (parsedJson.description || parsedJson.comment)) || '',
          coverImage: coverImgUrl,
          jsonData: parsedJson || {},
          rawJsonString: parsedJson ? JSON.stringify(parsedJson, null, 2) : contentText,
          content: contentText,
          fileType: ext,
          createdAt: Date.now(),
          updatedAt: Date.now(),
        };

        newBeautifications.push(beautificationItem);
      } catch (e: any) {
        showToast(`读取文件 ${file.name} 失败`, 'error');
      }

      if (showProgress && (i + 1) % 5 === 0) showToast(`正在导入 ${i + 1}/${fileArray.length} 个文件…`, 'info');
      await new Promise((resolve) => setTimeout(resolve, 0));
    }

    if (newBeautifications.length > 0) {
      updateAppData((prev) => ({ ...prev, beautifications: [...newBeautifications, ...(prev.beautifications || [])] }));
      showToast(`成功导入 ${newBeautifications.length} 个美化！`, 'success');
    }
  };

  const exportBeautificationDocument = (beautificationItem: ThemeEntry, format: 'docx' | 'txt' | 'json' | 'css') => {
    try {
      const content = beautificationItem.content || beautificationItem.rawJsonString || (typeof beautificationItem.jsonData === 'string' ? beautificationItem.jsonData : JSON.stringify(beautificationItem.jsonData || {}, null, 2)) || '';
      let mimeType = 'text/plain;charset=utf-8';
      let ext = format;
      if (format === 'json') mimeType = 'application/json;charset=utf-8';
      if (format === 'css') mimeType = 'text/css;charset=utf-8';
      if (format === 'docx') mimeType = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

      const blob = new Blob([content], { type: mimeType });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `${beautificationItem.name || '美化文档'}.${ext}`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      showToast(`已成功导出 .${ext} 格式文档！`, 'success');
    } catch (err: any) {
      showToast('导出失败: ' + err.message, 'error');
    }
  };

  const handleSaveEditedBeautification = () => {
    if (!editingBeautification) return;
    if (!editingBeautification.name.trim()) {
      showToast('请填写美化名称', 'error');
      return;
    }

    let finalJsonData = editingBeautification.jsonData;
    let finalRawStr = editingBeautification.rawJsonString || editingBeautification.content || '';

    if (editingBeautification.fileType === 'json' && editingBeautification.rawJsonString) {
      try {
        finalJsonData = JSON.parse(editingBeautification.rawJsonString);
      } catch (e: any) {
        // retain string
      }
    }

    const updatedBeautifications = (appData.beautifications || []).map((b) =>
      b.id === editingBeautification.id
        ? {
            ...editingBeautification,
            name: editingBeautification.name.trim(),
            author: editingBeautification.author?.trim() || '默认',
            type: editingBeautification.type?.trim() || '线上主题',
            category: editingBeautification.category?.trim() || '默认',
            source: editingBeautification.source?.trim() || '',
            description: editingBeautification.description?.trim() || '',
            content: editingBeautification.content || editingBeautification.rawJsonString || '',
            jsonData: finalJsonData,
            rawJsonString: finalRawStr,
            updatedAt: Date.now(),
          }
        : b
    );

    const cat = editingBeautification.category?.trim() || '默认';
    const existingCats = appData.beautificationCategories || ['默认'];
    const newCats = existingCats.includes(cat) ? existingCats : [...existingCats, cat];

    updateAppData({ ...appData, beautifications: updatedBeautifications, beautificationCategories: newCats });
    setEditingBeautification(null);
    showToast('美化存档已保存！', 'success');
  };

  const handleDeleteSingleBeautification = (id: string) => {
    if (!confirm('确定要删除此美化存档吗？')) return;
    const updatedBeautifications = (appData.beautifications || []).filter((b) => b.id !== id);
    updateAppData({ ...appData, beautifications: updatedBeautifications });
    setEditingBeautification(null);
    showToast('已删除美化存档', 'info');
  };

  const handleBatchDeleteBeautifications = () => {
    if (selectedBeautificationIds.length === 0) return;
    if (!confirm(`确定要删除选中的 ${selectedBeautificationIds.length} 个美化存档吗？`)) return;
    const updatedBeautifications = (appData.beautifications || []).filter((b) => !selectedBeautificationIds.includes(b.id));
    updateAppData({ ...appData, beautifications: updatedBeautifications });
    setSelectedBeautificationIds([]);
    setBeautificationBatchMode(false);
    showToast(`已批量删除 ${selectedBeautificationIds.length} 个美化`, 'info');
  };

  const handleCreateNewBeautificationGroup = () => {
    const name = newBeautificationGroupName.trim();
    if (!name) {
      showToast('请输入分组名称', 'error');
      return;
    }
    const currentCats = appData.beautificationCategories || ['默认'];
    if (currentCats.includes(name)) {
      showToast('该分组已存在', 'info');
      setBeautificationCategoryFilter(name);
      setShowNewBeautificationGroupModal(false);
      setNewBeautificationGroupName('');
      return;
    }
    const updatedCats = [...currentCats, name];
    updateAppData({ ...appData, beautificationCategories: updatedCats });
    setBeautificationCategoryFilter(name);
    setShowNewBeautificationGroupModal(false);
    setNewBeautificationGroupName('');
    showToast(`成功新建分组: ${name}`, 'success');
  };

  const handleRenameBeautificationGroup = () => {
    if (!managingBeautificationGroup) return;
    const newName = renameBeautificationGroupInput.trim();
    if (!newName) {
      showToast('请输入新的分组名称', 'error');
      return;
    }
    if (newName === managingBeautificationGroup) {
      setManagingBeautificationGroup(null);
      return;
    }

    const currentCats = appData.beautificationCategories || ['默认'];
    const updatedCats = currentCats.map((c) => (c === managingBeautificationGroup ? newName : c));

    const updatedBeautifications = (appData.beautifications || []).map((b) =>
      (b.category || '默认') === managingBeautificationGroup ? { ...b, category: newName } : b
    );

    updateAppData({
      ...appData,
      beautificationCategories: Array.from(new Set(updatedCats)),
      beautifications: updatedBeautifications,
    });

    if (beautificationCategoryFilter === managingBeautificationGroup) {
      setBeautificationCategoryFilter(newName);
    }

    setManagingBeautificationGroup(null);
    showToast('分组重命名成功', 'success');
  };

  const handleConfirmDeleteBeautificationGroup = () => {
    if (!managingBeautificationGroup) return;
    const targetGroup = managingBeautificationGroup;

    let updatedBeautifications = appData.beautifications || [];
    if (deleteBeautificationsWithGroup) {
      updatedBeautifications = updatedBeautifications.filter((b) => (b.category || '默认') !== targetGroup);
    } else {
      updatedBeautifications = updatedBeautifications.map((b) =>
        (b.category || '默认') === targetGroup ? { ...b, category: '默认' } : b
      );
    }

    const currentCats = appData.beautificationCategories || ['默认'];
    const updatedCats = currentCats.filter((c) => c !== targetGroup);

    updateAppData({
      ...appData,
      beautificationCategories: updatedCats,
      beautifications: updatedBeautifications,
    });

    if (beautificationCategoryFilter === targetGroup) {
      setBeautificationCategoryFilter('全部分类');
    }

    setManagingBeautificationGroup(null);
    setDeleteBeautificationsWithGroup(false);
    showToast(`分组 “${targetGroup}” 已删除`, 'info');
  };

  const handleBatchMoveBeautifications = () => {
    if (selectedBeautificationIds.length === 0) return;
    const target = batchTargetBeautificationGroup.trim();
    if (!target) {
      showToast('请选择或输入目标分组', 'error');
      return;
    }

    const updatedBeautifications = (appData.beautifications || []).map((b) =>
      selectedBeautificationIds.includes(b.id) ? { ...b, category: target } : b
    );

    const existingCats = appData.beautificationCategories || ['默认'];
    const newCats = existingCats.includes(target) ? existingCats : [...existingCats, target];

    updateAppData({
      ...appData,
      beautifications: updatedBeautifications,
      beautificationCategories: newCats,
    });

    setSelectedBeautificationIds([]);
    setBeautificationBatchMode(false);
    setShowBeautificationBatchMoveModal(false);
    setBatchTargetBeautificationGroup('');
    showToast(`已将 ${selectedBeautificationIds.length} 个美化移动到 “${target}”`, 'success');
  };

  // ==================== ST PRESETS STATE & HANDLERS ====================
  const presetFileInputRef = useRef<HTMLInputElement>(null);
  const presetUpdateInputRef = useRef<HTMLInputElement>(null);
  const presetRegexFileInputRef = useRef<HTMLInputElement>(null);

  const [presetSearchQuery, setPresetSearchQuery] = useState('');
  const [presetCategoryFilter, setPresetCategoryFilter] = useState('全部分组');
  const [presetBatchMode, setPresetBatchMode] = useState(false);
  const [selectedPresetIds, setSelectedPresetIds] = useState<string[]>([]);

  // Group Modals State for Presets
  const [showNewPresetGroupModal, setShowNewPresetGroupModal] = useState(false);
  const [newPresetGroupName, setNewPresetGroupName] = useState('');

  const [managingPresetGroup, setManagingPresetGroup] = useState<string | null>(null);
  const [renamePresetGroupInput, setRenamePresetGroupInput] = useState('');
  const [deletePresetsWithGroup, setDeletePresetsWithGroup] = useState(false);

  const [showPresetBatchMoveModal, setShowPresetBatchMoveModal] = useState(false);
  const [batchTargetPresetGroup, setBatchTargetPresetGroup] = useState('');

  // Preset Detail Modal State
  const [editingPreset, setEditingPreset] = useState<PresetEntry | null>(null);
  const [editingPresetTab, setEditingPresetTab] = useState<'details' | 'content' | 'regex'>('details');
  const [presetEntrySearchQuery, setPresetEntrySearchQuery] = useState('');
  const [expandedPresetEntry, setExpandedPresetEntry] = useState<{ key: string; value: any; index?: number } | null>(null);

  // Preset Entry Management State (Multi-select, Add, Edit, Move Up/Down, Move to other Presets)
  const [presetEntryBatchMode, setPresetEntryBatchMode] = useState(false);
  const [selectedPresetEntryIndices, setSelectedPresetEntryIndices] = useState<number[]>([]);
  const [showAddPresetEntryModal, setShowAddPresetEntryModal] = useState(false);
  const [newPresetEntryForm, setNewPresetEntryForm] = useState({ name: '', content: '' });
  const [editingPresetEntryModal, setEditingPresetEntryModal] = useState<{ index: number; name: string; content: string } | null>(null);
  const [showMovePresetEntryModal, setShowMovePresetEntryModal] = useState(false);
  const [movingEntryIndices, setMovingEntryIndices] = useState<number[]>([]);
  const [targetMovePresetId, setTargetMovePresetId] = useState('');

  // Helper to get structured entries list from preset
  const getPresetEntriesList = (preset: PresetEntry | null): { index: number; name: string; content: string; rawItem: any }[] => {
    if (!preset || !preset.jsonData) return [];
    const data = preset.jsonData;
    const list: { index: number; name: string; content: string; rawItem: any }[] = [];

    if (data && typeof data === 'object' && Array.isArray(data.prompts)) {
      data.prompts.forEach((p: any, idx: number) => {
        const name = p?.name || p?.identifier || p?.role || p?.title || `提示词 #${idx + 1}`;
        const content = p?.content ?? p?.value ?? p?.text ?? (typeof p === 'string' ? p : JSON.stringify(p, null, 2));
        list.push({ index: idx, name, content: decodeUnicodeAndEscapes(content), rawItem: p });
      });
    } else if (Array.isArray(data)) {
      data.forEach((item: any, idx: number) => {
        const name = item?.name || item?.identifier || item?.role || `条目 #${idx + 1}`;
        const content = item?.content ?? item?.value ?? item?.text ?? (typeof item === 'string' ? item : JSON.stringify(item, null, 2));
        list.push({ index: idx, name, content: decodeUnicodeAndEscapes(content), rawItem: item });
      });
    } else if (data && typeof data === 'object') {
      Object.entries(data).forEach(([k, v], idx) => {
        if (k !== 'regex_scripts' && k !== 'regexes' && k !== 'extensions' && k !== 'user_regexes') {
          const content = typeof v === 'object' ? JSON.stringify(v, null, 2) : String(v ?? '');
          list.push({ index: idx, name: k, content: decodeUnicodeAndEscapes(content), rawItem: v });
        }
      });
    } else {
      list.push({ index: 0, name: '内容数据', content: decodeUnicodeAndEscapes(data), rawItem: data });
    }
    return list;
  };

  // Helper to save updated entries list into preset and appData
  const savePresetEntriesList = (preset: PresetEntry, newList: { name: string; content: string; rawItem?: any }[]) => {
    const data = JSON.parse(JSON.stringify(preset.jsonData || {}));
    let newJsonData: any;

    if (data && typeof data === 'object' && Array.isArray(data.prompts)) {
      const updatedPrompts = newList.map((item) => {
        if (item.rawItem && typeof item.rawItem === 'object') {
          return { ...item.rawItem, name: item.name, content: item.content };
        }
        return { name: item.name, content: item.content, role: 'system' };
      });
      newJsonData = { ...data, prompts: updatedPrompts };
    } else if (Array.isArray(data)) {
      newJsonData = newList.map((item) => {
        if (item.rawItem && typeof item.rawItem === 'object') {
          return { ...item.rawItem, name: item.name, content: item.content };
        }
        return { name: item.name, content: item.content };
      });
    } else if (data && typeof data === 'object') {
      const newObj: Record<string, any> = {};
      if (data.regex_scripts) newObj.regex_scripts = data.regex_scripts;
      if (data.extensions) newObj.extensions = data.extensions;
      newList.forEach((item) => {
        newObj[item.name] = item.content;
      });
      newJsonData = newObj;
    } else {
      newJsonData = newList.length === 1 ? newList[0].content : newList;
    }

    const updatedEntry: PresetEntry = {
      ...preset,
      jsonData: newJsonData,
      rawJsonString: JSON.stringify(newJsonData, null, 2),
      updatedAt: Date.now(),
    };

    setEditingPreset(updatedEntry);

    updateAppData((prev) => ({
      ...prev,
      presets: (prev.presets || []).map((p) => (p.id === preset.id ? updatedEntry : p)),
    }));
  };

  // Filtered Presets List
  const presetsList = appData.presets || [];
  const availablePresetCategories = Array.from(
    new Set(['全部分组', '默认', ...(appData.presetCategories || []), ...presetsList.map((p) => p.category || '默认')])
  );

  const filteredPresets = presetsList.filter((item) => {
    if (presetCategoryFilter !== '全部分组' && (item.category || '默认') !== presetCategoryFilter) {
      return false;
    }
    if (!presetSearchQuery.trim()) return true;
    const q = presetSearchQuery.toLowerCase();
    const rawStr = item.rawJsonString || (typeof item.jsonData === 'string' ? item.jsonData : JSON.stringify(item.jsonData));
    return (
      item.name.toLowerCase().includes(q) ||
      (item.fileName && item.fileName.toLowerCase().includes(q)) ||
      (item.author && item.author.toLowerCase().includes(q)) ||
      (item.category && item.category.toLowerCase().includes(q)) ||
      (item.description && item.description.toLowerCase().includes(q)) ||
      (item.source && item.source.toLowerCase().includes(q)) ||
      (rawStr && rawStr.toLowerCase().includes(q))
    );
  });

  const decodeUnicodeAndEscapes = (str: any): string => {
    if (typeof str !== 'string') return typeof str === 'object' ? JSON.stringify(str, null, 2) : String(str ?? '');
    try {
      let s = str.replace(/\\u([0-9a-fA-F]{4})/g, (_, hex) =>
        String.fromCharCode(parseInt(hex, 16))
      );
      s = s.replace(/\\n/g, '\n').replace(/\\r/g, '\r').replace(/\\t/g, '\t');
      return s;
    } catch {
      return str;
    }
  };

  const handlePresetFileUpload = async (files: FileList) => {
    const fileArray = Array.from(files);
    if (fileArray.length === 0) return;

    const newEntries: PresetEntry[] = [];

    for (let i = 0; i < fileArray.length; i++) {
      const file = fileArray[i];
      try {
        const text = await file.text();
        let json: any;
        try {
          json = JSON.parse(text);
        } catch {
          const cleanedText = decodeUnicodeAndEscapes(text);
          json = JSON.parse(cleanedText);
        }
        const nameClean = json.name || json.title || file.name.replace(/\.json$/i, '') || '未命名预设';

        const entry: PresetEntry = {
          id: 'preset_' + Date.now() + '_' + Math.random().toString(36).substring(2, 7),
          name: nameClean,
          fileName: file.name,
          author: json.author || json.creator || json.user || '',
          category: presetCategoryFilter !== '全部分组' ? presetCategoryFilter : '默认',
          source: json.source || json.url || json.dc || '',
          description: json.description || json.notes || json.comment || '',
          jsonData: json,
          rawJsonString: JSON.stringify(json, null, 2),
          createdAt: Date.now(),
          updatedAt: Date.now(),
        };

        newEntries.push(entry);
      } catch (err) {
        console.error('Failed to parse preset JSON file:', file.name, err);
        showToast(`无法解析文件: ${file.name}，请确保是合法的 JSON 格式`, 'error');
      }
    }

    if (newEntries.length > 0) {
      const updated = [...(appData.presets || []), ...newEntries];
      updateAppData({ ...appData, presets: updated });
      showToast(`成功导入 ${newEntries.length} 个预设文件`, 'success');
    }
  };

  const handleCreateNewPresetGroup = () => {
    const name = newPresetGroupName.trim();
    if (!name) {
      showToast('请输入分组名称', 'error');
      return;
    }
    const currentCats = appData.presetCategories || ['默认'];
    if (currentCats.includes(name)) {
      showToast('该分组已存在', 'info');
      setPresetCategoryFilter(name);
      setShowNewPresetGroupModal(false);
      setNewPresetGroupName('');
      return;
    }
    const updatedCats = [...currentCats, name];
    updateAppData({ ...appData, presetCategories: updatedCats });
    setPresetCategoryFilter(name);
    setShowNewPresetGroupModal(false);
    setNewPresetGroupName('');
    showToast(`成功新建预设分组: ${name}`, 'success');
  };

  const handleRenamePresetGroup = () => {
    if (!managingPresetGroup) return;
    const newName = renamePresetGroupInput.trim();
    if (!newName) {
      showToast('请输入新的分组名称', 'error');
      return;
    }
    if (newName === managingPresetGroup) {
      setManagingPresetGroup(null);
      return;
    }

    const currentCats = appData.presetCategories || ['默认'];
    const updatedCats = currentCats.map((c) => (c === managingPresetGroup ? newName : c));

    const updatedPresets = (appData.presets || []).map((p) =>
      (p.category || '默认') === managingPresetGroup ? { ...p, category: newName } : p
    );

    updateAppData({
      ...appData,
      presetCategories: Array.from(new Set(updatedCats)),
      presets: updatedPresets,
    });

    if (presetCategoryFilter === managingPresetGroup) {
      setPresetCategoryFilter(newName);
    }

    setManagingPresetGroup(null);
    showToast('分组重命名成功', 'success');
  };

  const handleConfirmDeletePresetGroup = () => {
    if (!managingPresetGroup) return;
    const targetGroup = managingPresetGroup;

    let updatedPresets = appData.presets || [];
    if (deletePresetsWithGroup) {
      updatedPresets = updatedPresets.filter((p) => (p.category || '默认') !== targetGroup);
    } else {
      updatedPresets = updatedPresets.map((p) =>
        (p.category || '默认') === targetGroup ? { ...p, category: '默认' } : p
      );
    }

    const currentCats = appData.presetCategories || ['默认'];
    const updatedCats = currentCats.filter((c) => c !== targetGroup);

    updateAppData({
      ...appData,
      presetCategories: updatedCats,
      presets: updatedPresets,
    });

    if (presetCategoryFilter === targetGroup) {
      setPresetCategoryFilter('全部分组');
    }

    setManagingPresetGroup(null);
    setDeletePresetsWithGroup(false);
    showToast(`分组 “${targetGroup}” 已删除`, 'info');
  };

  const handleBatchMovePresets = () => {
    if (selectedPresetIds.length === 0) return;
    const target = batchTargetPresetGroup.trim();
    if (!target) {
      showToast('请选择或输入目标分组', 'error');
      return;
    }

    const updatedPresets = (appData.presets || []).map((p) =>
      selectedPresetIds.includes(p.id) ? { ...p, category: target } : p
    );

    const existingCats = appData.presetCategories || ['默认'];
    const newCats = existingCats.includes(target) ? existingCats : [...existingCats, target];

    updateAppData({
      ...appData,
      presets: updatedPresets,
      presetCategories: newCats,
    });

    setSelectedPresetIds([]);
    setPresetBatchMode(false);
    setShowPresetBatchMoveModal(false);
    setBatchTargetPresetGroup('');
    showToast(`已将 ${selectedPresetIds.length} 个预设移动到 “${target}”`, 'success');
  };

  const handleBatchDeletePresets = () => {
    if (selectedPresetIds.length === 0) return;
    const updatedPresets = (appData.presets || []).filter((p) => !selectedPresetIds.includes(p.id));
    updateAppData({ ...appData, presets: updatedPresets });
    setSelectedPresetIds([]);
    setPresetBatchMode(false);
    showToast(`已批量删除 ${selectedPresetIds.length} 个预设`, 'info');
  };

  // Detail Modal Actions
  const handleSaveEditingPreset = async () => {
    if (!editingPreset) return;

    // 以当前编辑面板的最终快照为准，并使用函数式更新，避免快速编辑/保存时旧 appData 覆盖新内容。
    const snapshot: PresetEntry = {
      ...editingPreset,
      jsonData: editingPreset.jsonData,
      updatedAt: Date.now(),
      rawJsonString: JSON.stringify(editingPreset.jsonData || {}, null, 2),
    };

    updateAppData((prev) => ({
      ...prev,
      presets: (prev.presets || []).map((item) => (item.id === snapshot.id ? { ...snapshot } : item)),
    }));

    // 先把编辑态同步为最终快照，保证任何仍在显示的预览内容立即跟随修改；
    // 随后关闭详情面板，外层预设列表直接从更新后的 appData 渲染。
    setEditingPreset(snapshot);
    setEditingPreset(null);
    showToast('预设修改已保存，预览内容已同步更新', 'success');
  };

  const handleExportEditingPreset = () => {
    if (!editingPreset) return;
    const dataStr = JSON.stringify(editingPreset.jsonData || {}, null, 2);
    const blob = new Blob([dataStr], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = editingPreset.fileName || `${editingPreset.name}.json`;
    a.click();
    URL.revokeObjectURL(url);
    showToast('预设文件导出成功', 'success');
  };

  const handleDeleteEditingPreset = () => {
    if (!editingPreset) return;
    const updatedPresets = (appData.presets || []).filter((p) => p.id !== editingPreset.id);
    updateAppData({ ...appData, presets: updatedPresets });
    setEditingPreset(null);
    showToast('预设已删除', 'info');
  };

  const handleUpdateEditingPresetFile = async (file: File) => {
    if (!editingPreset) return;
    try {
      const text = await file.text();
      const json = JSON.parse(text);
      const updatedEntry: PresetEntry = {
        ...editingPreset,
        fileName: file.name,
        jsonData: json,
        rawJsonString: JSON.stringify(json, null, 2),
        updatedAt: Date.now(),
      };
      setEditingPreset(updatedEntry);

      // Save to appData
      const updatedPresets = (appData.presets || []).map((p) => (p.id === editingPreset.id ? updatedEntry : p));
      updateAppData({ ...appData, presets: updatedPresets });
      showToast('预设数据文件已更新覆盖', 'success');
    } catch (err) {
      showToast('无法解析上传的 JSON 文件', 'error');
    }
  };

  const getPresetRegexScripts = (preset: PresetEntry | null): any[] => {
    if (!preset) return [];
    if (Array.isArray(preset.regexScripts) && preset.regexScripts.length > 0) {
      return preset.regexScripts;
    }
    const data = preset.jsonData;
    if (!data) return [];
    if (Array.isArray(data.regex_scripts)) return data.regex_scripts;
    if (Array.isArray(data.extensions?.regex_scripts)) return data.extensions.regex_scripts;
    if (Array.isArray(data.regexes)) return data.regexes;
    if (Array.isArray(data.user_regexes)) return data.user_regexes;
    if (Array.isArray(data.regex)) return data.regex;
    if (Array.isArray(data)) {
      const isRegexArray = data.some(
        (item) =>
          item &&
          typeof item === 'object' &&
          (item.findRegex || item.find_regex || item.pattern || item.scriptName || item.script_name || item.replaceString || item.replace_string)
      );
      if (isRegexArray) return data;
    }
    if (data.findRegex || data.find_regex || data.pattern || data.scriptName || data.script_name) {
      return [data];
    }
    return [];
  };

  const handleUploadPresetRegexFile = async (file: File) => {
    if (!editingPreset) return;
    try {
      const text = await file.text();
      const json = JSON.parse(text);
      let newRegexScripts: any[] = [];

      if (Array.isArray(json)) {
        newRegexScripts = json;
      } else if (json && typeof json === 'object') {
        if (Array.isArray(json.regex_scripts)) {
          newRegexScripts = json.regex_scripts;
        } else if (Array.isArray(json.regexes)) {
          newRegexScripts = json.regexes;
        } else if (Array.isArray(json.extensions?.regex_scripts)) {
          newRegexScripts = json.extensions.regex_scripts;
        } else if (json.findRegex || json.find_regex || json.pattern || json.scriptName || json.script_name) {
          newRegexScripts = [json];
        }
      }

      if (newRegexScripts.length === 0) {
        showToast('未能识别有效的正则数据，请检查文件格式', 'error');
        return;
      }

      const normalizedScripts = newRegexScripts.map((rx) => {
        const name = rx.scriptName || rx.script_name || rx.name || rx.title || '导入正则';
        const find = rx.findRegex || rx.find_regex || rx.pattern || rx.find || rx.regex || '';
        const replace = rx.replaceString || rx.replace_string || rx.replacement || rx.replace || '';
        return {
          ...rx,
          scriptName: name, script_name: name, name: name,
          findRegex: find, find_regex: find, pattern: find,
          replaceString: replace, replace_string: replace, replacement: replace,
          disabled: Boolean(rx.disabled),
        };
      });

      const currentRegex = getPresetRegexScripts(editingPreset);
      const combined = [...currentRegex, ...normalizedScripts];

      const updatedPreset: PresetEntry = {
        ...editingPreset,
        regexScripts: combined,
        jsonData: {
          ...(typeof editingPreset.jsonData === 'object' ? editingPreset.jsonData : {}),
          extensions: {
            ...(editingPreset.jsonData?.extensions || {}),
            regex_scripts: combined,
          },
          regex_scripts: combined,
        },
      };

      setEditingPreset(updatedPreset);
      const updatedPresets = (appData.presets || []).map((p) => (p.id === editingPreset.id ? updatedPreset : p));
      updateAppData({ ...appData, presets: updatedPresets });
      showToast(`已成功自动识别并添加 ${normalizedScripts.length} 条正则`, 'success');
    } catch (err: any) {
      console.error(err);
      showToast('解析正则文件失败: ' + err.message, 'error');
    }
  };

  // ==================== ST PLUGINS & SCRIPTS STATE & HANDLERS ====================
  const pluginScriptFileInputRef = useRef<HTMLInputElement>(null);

  const [pluginSearchQuery, setPluginSearchQuery] = useState('');
  const [pluginCategoryFilter, setPluginCategoryFilter] = useState('全部分组');
  const [pluginBatchMode, setPluginBatchMode] = useState(false);
  const [selectedPluginIds, setSelectedPluginIds] = useState<string[]>([]);

  // Selection popup state for top bar '+' button when on 'st-plugins' page
  const [showAddPluginTypeModal, setShowAddPluginTypeModal] = useState(false);

  // Add Plugin (Link) Modal State
  const [showAddPluginModal, setShowAddPluginModal] = useState(false);
  const [pluginForm, setPluginForm] = useState({
    name: '',
    url: '',
    contact: '',
    description: '',
  });

  // Detail Modal State for Plugin or Script
  // 插件保持原来的单页详情；脚本独立使用双 Tab 详情。
  const [editingPlugin, setEditingPlugin] = useState<PluginEntry | null>(null);
  const [scriptDetailTab, setScriptDetailTab] = useState<'info' | 'code'>('info');
  const [isScriptCodeExpanded, setIsScriptCodeExpanded] = useState(false);
  const [scriptSearchQuery, setScriptSearchQuery] = useState('');
  const [scriptContentDraft, setScriptContentDraft] = useState('');

  // New Group Modal State for Plugins
  const [showNewPluginGroupModal, setShowNewPluginGroupModal] = useState(false);
  const [newPluginGroupName, setNewPluginGroupName] = useState('');

  // Manage Group Modal State for Plugins
  const [managingPluginCategory, setManagingPluginCategory] = useState<string | null>(null);
  const [renamePluginCategoryInput, setRenamePluginCategoryInput] = useState('');
  const [showDeletePluginCategoryConfirm, setShowDeletePluginCategoryConfirm] = useState(false);
  const [deleteItemsWithPluginCategory, setDeleteItemsWithPluginCategory] = useState(false);

  // Batch Move Modal State for Plugins
  const [showPluginBatchMoveModal, setShowPluginBatchMoveModal] = useState(false);
  const [batchTargetPluginCategory, setBatchTargetPluginCategory] = useState('');

  // List and Filtering
  const pluginsList = appData.plugins || [];

  const filteredPlugins = pluginsList.filter((item) => {
    if (pluginCategoryFilter !== '全部分组' && (item.category || '默认') !== pluginCategoryFilter) {
      return false;
    }
    if (!pluginSearchQuery.trim()) return true;
    const q = pluginSearchQuery.toLowerCase();
    return (
      item.name.toLowerCase().includes(q) ||
      (item.url && item.url.toLowerCase().includes(q)) ||
      (item.contact && item.contact.toLowerCase().includes(q)) ||
      (item.author && item.author.toLowerCase().includes(q)) ||
      (item.category && item.category.toLowerCase().includes(q)) ||
      (item.source && item.source.toLowerCase().includes(q)) ||
      (item.description && item.description.toLowerCase().includes(q)) ||
      (item.fileName && item.fileName.toLowerCase().includes(q))
    );
  });

  // Handlers
  const handleAddPluginSubmit = () => {
    if (!pluginForm.name.trim()) {
      showToast('请填写插件名称', 'error');
      return;
    }
    if (!pluginForm.url.trim()) {
      showToast('请填写插件地址', 'error');
      return;
    }

    const newPlugin: PluginEntry = {
      id: Date.now().toString(36) + Math.random().toString(36).substring(2, 6),
      type: 'plugin',
      name: pluginForm.name.trim(),
      url: pluginForm.url.trim(),
      contact: pluginForm.contact.trim(),
      description: pluginForm.description.trim(),
      category: pluginCategoryFilter !== '全部分组' ? pluginCategoryFilter : '默认',
      createdAt: Date.now(),
      updatedAt: Date.now(),
    };

    const updated = [...pluginsList, newPlugin];
    updateAppData({ ...appData, plugins: updated });
    setPluginForm({ name: '', url: '', contact: '', description: '' });
    setShowAddPluginModal(false);
    showToast('插件添加成功！', 'success');
  };

  const handleScriptFileUpload = async (files: FileList) => {
    const fileArray = Array.from(files);
    if (fileArray.length === 0) return;

    const newEntries: PluginEntry[] = [];

    for (let i = 0; i < fileArray.length; i++) {
      const file = fileArray[i];
      try {
        const text = await file.text();
        let json: any = {};
        try {
          json = JSON.parse(text);
        } catch (e) {
          json = {};
        }

        const cleanFileName = file.name.replace(/\.[^/.]+$/, '');
        const nameClean = json.name || json.script_name || json.scriptName || json.qrName || json.qr_name || json.title || json.label || cleanFileName || '未命名脚本';
        const authorClean = json.author || json.creator || json.display_name || json.authorName || '未知作者';

        const entry: PluginEntry = {
          id: Date.now().toString(36) + Math.random().toString(36).substring(2, 6) + i,
          type: 'script',
          name: nameClean,
          fileName: file.name,
          author: authorClean,
          category: pluginCategoryFilter !== '全部分组' ? pluginCategoryFilter : '默认',
          source: json.source || json.link || '',
          description: json.description || json.summary || '',
          jsonData: json,
          createdAt: Date.now(),
          updatedAt: Date.now(),
        };

        newEntries.push(entry);
      } catch (err: any) {
        showToast(`读取文件 "${file.name}" 失败: ${err.message}`, 'error');
      }
    }

    if (newEntries.length > 0) {
      const updated = [...pluginsList, ...newEntries];
      updateAppData({ ...appData, plugins: updated });
      showToast(`成功导入 ${newEntries.length} 个脚本文件！`, 'success');
    }
  };

  const getPluginScriptContent = (plugin: PluginEntry | null) => {
    if (!plugin) return '';
    if (typeof plugin.jsonData === 'string') return plugin.jsonData;
    try {
      return JSON.stringify(plugin.jsonData ?? {}, null, 2);
    } catch {
      return String(plugin.jsonData ?? '');
    }
  };

  const handleOpenPluginDetail = (item: PluginEntry) => {
    setEditingPlugin({ ...item });
    if (item.type === 'script') {
      setScriptDetailTab('info');
      setScriptSearchQuery('');
      setScriptContentDraft(getPluginScriptContent(item));
      setIsScriptCodeExpanded(false);
    }
  };

  const handleExportPluginScriptJson = (plugin: PluginEntry) => {
    if (plugin.type !== 'script') return;
    const content = getPluginScriptContent(plugin);
    const blob = new Blob([content], { type: 'application/json;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${(plugin.name || plugin.fileName || 'script').replace(/[\\/:*?"<>|]/g, '_')}.json`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
    showToast('脚本 JSON 已导出', 'success');
  };

  const handleSaveScriptContent = () => {
    if (!editingPlugin || editingPlugin.type !== 'script') return;
    try {
      const parsed = JSON.parse(scriptContentDraft);
      const updatedPlugin = {
        ...editingPlugin,
        jsonData: parsed,
        updatedAt: Date.now(),
      };
      setEditingPlugin(updatedPlugin);
      const updated = pluginsList.map((p) => (p.id === updatedPlugin.id ? updatedPlugin : p));
      updateAppData({ ...appData, plugins: updated });
      showToast('脚本内容已保存', 'success');
    } catch {
      showToast('脚本内容不是有效的 JSON，请检查后再保存', 'error');
    }
  };

  const handleCreateNewPluginCategory = () => {
    const trimmed = newPluginGroupName.trim();
    if (!trimmed) return;
    if (trimmed === '全部分组') {
      showToast('"全部分组"为系统保留名称', 'error');
      return;
    }
    const currentCats = appData.pluginCategories || ['默认'];
    if (currentCats.includes(trimmed)) {
      showToast('该分组名称已存在', 'error');
      return;
    }

    updateAppData({
      ...appData,
      pluginCategories: [...currentCats, trimmed],
    });
    setNewPluginGroupName('');
    setShowNewPluginGroupModal(false);
    setPluginCategoryFilter(trimmed);
    showToast(`成功新建分组: ${trimmed}`, 'success');
  };

  const handleRenamePluginCategory = () => {
    if (!managingPluginCategory) return;
    const trimmed = renamePluginCategoryInput.trim();
    if (!trimmed) return;
    if (trimmed === '全部分组') {
      showToast('"全部分组"为系统保留名称', 'error');
      return;
    }

    const currentCats = appData.pluginCategories || ['默认'];
    const updatedCats = currentCats.map((c) => (c === managingPluginCategory ? trimmed : c));
    const updatedPlugins = pluginsList.map((p) =>
      (p.category || '默认') === managingPluginCategory ? { ...p, category: trimmed } : p
    );

    updateAppData({
      ...appData,
      plugins: updatedPlugins,
      pluginCategories: updatedCats,
    });

    if (pluginCategoryFilter === managingPluginCategory) {
      setPluginCategoryFilter(trimmed);
    }
    showToast(`分组已重命名为 "${trimmed}"`, 'success');
    setManagingPluginCategory(null);
  };

  const handleConfirmDeletePluginCategory = () => {
    if (!managingPluginCategory) return;

    let updatedPlugins = [...pluginsList];
    if (deleteItemsWithPluginCategory) {
      updatedPlugins = updatedPlugins.filter((p) => (p.category || '默认') !== managingPluginCategory);
    } else {
      updatedPlugins = updatedPlugins.map((p) =>
        (p.category || '默认') === managingPluginCategory ? { ...p, category: '默认' } : p
      );
    }

    const currentCats = appData.pluginCategories || ['默认'];
    const updatedCats = currentCats.filter((c) => c !== managingPluginCategory);

    updateAppData({
      ...appData,
      plugins: updatedPlugins,
      pluginCategories: updatedCats,
    });

    showToast(`分组 "${managingPluginCategory}" 已删除`, 'info');
    if (pluginCategoryFilter === managingPluginCategory) {
      setPluginCategoryFilter('全部分组');
    }
    setManagingPluginCategory(null);
    setShowDeletePluginCategoryConfirm(false);
    setDeleteItemsWithPluginCategory(false);
  };

  const handleBatchDeletePlugins = () => {
    if (selectedPluginIds.length === 0) return;
    if (!confirm(`确定要删除选中的 ${selectedPluginIds.length} 个插件/脚本吗？`)) return;

    const updated = pluginsList.filter((p) => !selectedPluginIds.includes(p.id));
    updateAppData({ ...appData, plugins: updated });

    showToast(`已成功删除 ${selectedPluginIds.length} 个插件/脚本`, 'info');
    setSelectedPluginIds([]);
    setPluginBatchMode(false);
  };

  const handleConfirmBatchMovePlugins = () => {
    if (!batchTargetPluginCategory) {
      showToast('请选择目标分组', 'error');
      return;
    }
    const target = batchTargetPluginCategory;
    const updated = pluginsList.map((p) =>
      selectedPluginIds.includes(p.id) ? { ...p, category: target } : p
    );

    updateAppData({ ...appData, plugins: updated });

    setSelectedPluginIds([]);
    setPluginBatchMode(false);
    setShowPluginBatchMoveModal(false);
    setBatchTargetPluginCategory('');
    showToast(`已将 ${selectedPluginIds.length} 个插件/脚本移动到 “${target}”`, 'success');
  };

  // ==================== NORMAL CHARACTER CARDS STATE & HANDLERS ====================
  const normalCardFileInputRef = useRef<HTMLInputElement>(null);
  const normalCardCoverInputRef = useRef<HTMLInputElement>(null);
  const normalCardUpdateFileInputRef = useRef<HTMLInputElement>(null);

  const [normalCardSearchQuery, setNormalCardSearchQuery] = useState('');
  const [normalCardCategoryFilter, setNormalCardCategoryFilter] = useState('全部分组');
  const [normalCardBatchMode, setNormalCardBatchMode] = useState(false);
  const [selectedNormalCardIds, setSelectedNormalCardIds] = useState<string[]>([]);

  // Single-Page Detail Modal State
  const [editingNormalCard, setEditingNormalCard] = useState<NormalCardEntry | null>(null);

  // Zoomed/Expanded Document Content Modal State
  const [showExpandedContentModal, setShowExpandedContentModal] = useState(false);

  // New Group Modal State
  const [showNewNormalCardGroupModal, setShowNewNormalCardGroupModal] = useState(false);
  const [newNormalCardGroupName, setNewNormalCardGroupName] = useState('');

  // Manage Group Modal State
  const [managingNormalCardCategory, setManagingNormalCardCategory] = useState<string | null>(null);
  const [renameNormalCardCategoryInput, setRenameNormalCardCategoryInput] = useState('');
  const [showDeleteNormalCardCategoryConfirm, setShowDeleteNormalCardCategoryConfirm] = useState(false);
  const [deleteItemsWithNormalCardCategory, setDeleteItemsWithNormalCardCategory] = useState(false);

  // Batch Move Modal State
  const [showNormalCardBatchMoveModal, setShowNormalCardBatchMoveModal] = useState(false);
  const [batchTargetNormalCardCategory, setBatchTargetNormalCardCategory] = useState('');

  const parseNormalCardFile = async (file: File): Promise<{ fileName: string; content: string }> => {
    const cleanFileName = file.name.replace(/\.[^/.]+$/, '');
    if (file.name.toLowerCase().endsWith('.docx')) {
      try {
        const arrayBuffer = await file.arrayBuffer();
        const result = await mammoth.extractRawText({ arrayBuffer });
        return { fileName: cleanFileName, content: result.value || '' };
      } catch (e: any) {
        console.error('Failed to parse docx', e);
        return { fileName: cleanFileName, content: `[Docx 读取解析失败: ${e.message}]` };
      }
    } else {
      const text = await file.text();
      return { fileName: cleanFileName, content: text };
    }
  };

  const normalCardsList = appData.normalCards || [];

  const filteredNormalCards = normalCardsList.filter((item) => {
    if (normalCardCategoryFilter !== '全部分组' && (item.category || '默认') !== normalCardCategoryFilter) {
      return false;
    }
    if (!normalCardSearchQuery.trim()) return true;
    const q = normalCardSearchQuery.toLowerCase();
    return (
      item.fileName.toLowerCase().includes(q) ||
      (item.charName && item.charName.toLowerCase().includes(q)) ||
      (item.author && item.author.toLowerCase().includes(q)) ||
      (item.category && item.category.toLowerCase().includes(q)) ||
      (item.source && item.source.toLowerCase().includes(q)) ||
      (item.content && item.content.toLowerCase().includes(q))
    );
  });

  const confirmNormalZipImport = async () => {
    if (!normalZipPreview) return;
    const selectedFiles = normalZipPreview.files.filter((item) => item.selected && /\.(docx|txt)$/i.test(item.name));
    if (!selectedFiles.length) {
      showToast('请至少选择一个文档导入', 'error');
      return;
    }
    const newEntries: NormalCardEntry[] = [];
    for (let i = 0; i < selectedFiles.length; i++) {
      const item = selectedFiles[i];
      const cleanName = item.name.split('/').pop()!.replace(/\.[^/.]+$/, '') || '未命名角色卡';
      newEntries.push({
        id: Date.now().toString(36) + Math.random().toString(36).substring(2, 6) + i,
        fileName: cleanName,
        charName: cleanName,
        author: '未知作者',
        category: normalCardCategoryFilter !== '全部分组' ? normalCardCategoryFilter : '默认',
        source: '', coverImage: null, content: item.content || '', createdAt: Date.now(), updatedAt: Date.now(),
      });
    }
    if (newEntries.length) {
      updateAppData((prev) => ({ ...prev, normalCards: [...(prev.normalCards || []), ...newEntries] }));
      showToast(`ZIP 已导入 ${newEntries.length} 个文档`, 'success');
    }
    setNormalZipPreview(null);
  };

  const toggleNormalZipFileSelection = (name: string) => {
    setNormalZipPreview((prev) => prev ? {
      ...prev,
      files: prev.files.map((f) => f.name === name ? { ...f, selected: !f.selected } : f),
    } : prev);
  };


  const handleNormalCardFileUpload = async (files: FileList) => {
    const fileArray = Array.from(files);
    if (fileArray.length === 0) return;
    const zipFile = fileArray.find((f) => f.name.toLowerCase().endsWith('.zip'));
    if (zipFile) {
      try {
        const bytes = new Uint8Array(await zipFile.arrayBuffer());
        const extracted = unzipSync(bytes);
        const previewFiles: { name: string; size: number; content: string; selected: boolean }[] = [];
        for (const [name, data] of Object.entries(extracted)) {
          if (name.endsWith('/') || !/\.(docx|txt)$/i.test(name)) continue;
          let content = '';
          if (/\.txt$/i.test(name)) content = strFromU8(data as Uint8Array);
          else {
            const ab = (data as Uint8Array).buffer.slice((data as Uint8Array).byteOffset, (data as Uint8Array).byteOffset + (data as Uint8Array).byteLength);
            const res = await mammoth.extractRawText({ arrayBuffer: ab });
            content = res.value || '';
          }
          previewFiles.push({ name, size: (data as Uint8Array).byteLength, content, selected: true });
        }
        setNormalZipPreview({ fileName: zipFile.name, files: previewFiles });
        showToast(`ZIP 已解压，识别到 ${previewFiles.length} 个可导入文件，请确认`, 'info');
      } catch (err: any) {
        showToast(`ZIP 解压失败: ${err.message || '格式错误'}`, 'error');
      }
      return;
    }

    const newEntries: NormalCardEntry[] = [];
    const showProgress = fileArray.length > 8;
    if (showProgress) showToast(`开始导入 ${fileArray.length} 个文件…`, 'info');
    for (let i = 0; i < fileArray.length; i++) {
      const file = fileArray[i];
      try {
        const parsed = await parseNormalCardFile(file);
        newEntries.push({
          id: Date.now().toString(36) + Math.random().toString(36).substring(2, 6) + i,
          fileName: parsed.fileName || '未命名角色卡', charName: parsed.fileName || '未知角色', author: '未知作者',
          category: normalCardCategoryFilter !== '全部分组' ? normalCardCategoryFilter : '默认', source: '', coverImage: null,
          content: parsed.content || '', createdAt: Date.now(), updatedAt: Date.now(),
        });
      } catch (err: any) { showToast(`读取文件 "${file.name}" 失败: ${err.message}`, 'error'); }
      // 批量文档解析（尤其是 docx）较耗 CPU，每个文件之间让出主线程避免长时间无响应。
      if (showProgress && (i + 1) % 5 === 0) showToast(`正在导入 ${i + 1}/${fileArray.length} 个文件…`, 'info');
      await new Promise((resolve) => setTimeout(resolve, 0));
    }
    if (newEntries.length) {
      updateAppData((prev) => ({ ...prev, normalCards: [...(prev.normalCards || []), ...newEntries] }));
      showToast(`成功导入 ${newEntries.length} 个普通角色卡文档！`, 'success');
    }
  };

  const handleUpdateNormalCardTextContent = async (file: File) => {
    if (!editingNormalCard) return;
    try {
      const parsed = await parseNormalCardFile(file);
      setEditingNormalCard({
        ...editingNormalCard,
        fileName: parsed.fileName || editingNormalCard.fileName,
        content: parsed.content || '',
        updatedAt: Date.now(),
      });
      showToast('文本内容已成功更新覆盖！', 'success');
    } catch (err: any) {
      showToast(`更新文本失败: ${err.message}`, 'error');
    }
  };

  const handleNormalCardCoverUpload = async (file: File) => {
    if (!editingNormalCard) return;
    try {
      const base64 = await fileToDataURL(file);
      setEditingNormalCard({
        ...editingNormalCard,
        coverImage: base64,
        updatedAt: Date.now(),
      });
      showToast('角色参考图已更新', 'success');
    } catch (err: any) {
      showToast(`上传图片失败: ${err.message}`, 'error');
    }
  };

  const handleCreateNewNormalCardCategory = () => {
    const trimmed = newNormalCardGroupName.trim();
    if (!trimmed) return;
    if (trimmed === '全部分组') {
      showToast('"全部分组"为系统保留名称', 'error');
      return;
    }
    const currentCats = appData.normalCardCategories || ['默认'];
    if (currentCats.includes(trimmed)) {
      showToast('该分组名称已存在', 'error');
      return;
    }

    updateAppData({
      ...appData,
      normalCardCategories: [...currentCats, trimmed],
    });
    setNewNormalCardGroupName('');
    setShowNewNormalCardGroupModal(false);
    setNormalCardCategoryFilter(trimmed);
    showToast(`成功新建分组: ${trimmed}`, 'success');
  };

  const handleRenameNormalCardCategory = () => {
    if (!managingNormalCardCategory) return;
    const trimmed = renameNormalCardCategoryInput.trim();
    if (!trimmed) return;
    if (trimmed === '全部分组') {
      showToast('"全部分组"为系统保留名称', 'error');
      return;
    }

    const currentCats = appData.normalCardCategories || ['默认'];
    const updatedCats = currentCats.map((c) => (c === managingNormalCardCategory ? trimmed : c));
    const updatedCards = normalCardsList.map((p) =>
      (p.category || '默认') === managingNormalCardCategory ? { ...p, category: trimmed } : p
    );

    updateAppData({
      ...appData,
      normalCards: updatedCards,
      normalCardCategories: updatedCats,
    });

    if (normalCardCategoryFilter === managingNormalCardCategory) {
      setNormalCardCategoryFilter(trimmed);
    }
    showToast(`分组已重命名为 "${trimmed}"`, 'success');
    setManagingNormalCardCategory(null);
  };

  const handleConfirmDeleteNormalCardCategory = () => {
    if (!managingNormalCardCategory) return;

    let updatedCards = [...normalCardsList];
    if (deleteItemsWithNormalCardCategory) {
      updatedCards = updatedCards.filter((p) => (p.category || '默认') !== managingNormalCardCategory);
    } else {
      updatedCards = updatedCards.map((p) =>
        (p.category || '默认') === managingNormalCardCategory ? { ...p, category: '默认' } : p
      );
    }

    const currentCats = appData.normalCardCategories || ['默认'];
    const updatedCats = currentCats.filter((c) => c !== managingNormalCardCategory);

    updateAppData({
      ...appData,
      normalCards: updatedCards,
      normalCardCategories: updatedCats,
    });

    showToast(`分组 "${managingNormalCardCategory}" 已删除`, 'info');
    if (normalCardCategoryFilter === managingNormalCardCategory) {
      setNormalCardCategoryFilter('全部分组');
    }
    setManagingNormalCardCategory(null);
    setShowDeleteNormalCardCategoryConfirm(false);
    setDeleteItemsWithNormalCardCategory(false);
  };

  const handleBatchDeleteNormalCards = () => {
    if (selectedNormalCardIds.length === 0) return;
    if (!confirm(`确定要删除选中的 ${selectedNormalCardIds.length} 个角色卡吗？`)) return;

    const updated = normalCardsList.filter((p) => !selectedNormalCardIds.includes(p.id));
    updateAppData({ ...appData, normalCards: updated });

    showToast(`已成功删除 ${selectedNormalCardIds.length} 个角色卡`, 'info');
    setSelectedNormalCardIds([]);
    setNormalCardBatchMode(false);
  };

  const handleConfirmBatchMoveNormalCards = () => {
    if (!batchTargetNormalCardCategory) {
      showToast('请选择目标分组', 'error');
      return;
    }
    const target = batchTargetNormalCardCategory;
    const updated = normalCardsList.map((p) =>
      selectedNormalCardIds.includes(p.id) ? { ...p, category: target } : p
    );

    updateAppData({ ...appData, normalCards: updated });

    setSelectedNormalCardIds([]);
    setNormalCardBatchMode(false);
    setShowNormalCardBatchMoveModal(false);
    setBatchTargetNormalCardCategory('');
    showToast(`已将 ${selectedNormalCardIds.length} 个角色卡移动到 “${target}”`, 'success');
  };

  // ==================== API STORAGE STATE & HANDLERS ====================
  const [apiSearchQuery, setApiSearchQuery] = useState('');
  const [apiCategoryFilter, setApiCategoryFilter] = useState('全部分组');
  const [apiBatchMode, setApiBatchMode] = useState(false);
  const [selectedApiIds, setSelectedApiIds] = useState<string[]>([]);

  // Add API Modal State
  const [showAddApiModal, setShowAddApiModal] = useState(false);
  const [newApiForm, setNewApiForm] = useState<{
    name: string;
    url: string;
    keys: ApiKeyItem[];
    description: string;
    category: string;
  }>({
    name: '',
    url: '',
    keys: [{ id: 'k-1', memo: '', key: '' }],
    description: '',
    category: '默认',
  });

  // Edit/Detail Single-Page Modal State
  const [editingApi, setEditingApi] = useState<ApiEntry | null>(null);

  // New Group Modal State
  const [showNewApiGroupModal, setShowNewApiGroupModal] = useState(false);
  const [newApiGroupName, setNewApiGroupName] = useState('');

  // Manage Group Modal State
  const [managingApiCategory, setManagingApiCategory] = useState<string | null>(null);
  const [renameApiCategoryInput, setRenameApiCategoryInput] = useState('');
  const [showDeleteApiCategoryConfirm, setShowDeleteApiCategoryConfirm] = useState(false);
  const [deleteItemsWithApiCategory, setDeleteItemsWithApiCategory] = useState(false);

  // Batch Move Modal State
  const [showApiBatchMoveModal, setShowApiBatchMoveModal] = useState(false);
  const [batchTargetApiCategory, setBatchTargetApiCategory] = useState('');

  const apisList = appData.apis || [];

  const filteredApis = apisList.filter((item) => {
    if (apiCategoryFilter !== '全部分组' && (item.category || '默认') !== apiCategoryFilter) {
      return false;
    }
    if (!apiSearchQuery.trim()) return true;
    const q = apiSearchQuery.toLowerCase();
    const nameMatch = item.name.toLowerCase().includes(q);
    const urlMatch = item.url.toLowerCase().includes(q);
    const descMatch = (item.description || '').toLowerCase().includes(q);
    const catMatch = (item.category || '').toLowerCase().includes(q);
    const keyMatch = item.keys.some(
      (k) => k.key.toLowerCase().includes(q) || (k.memo || '').toLowerCase().includes(q)
    );
    return nameMatch || urlMatch || descMatch || catMatch || keyMatch;
  });

  const handleOpenAddApiModal = () => {
    setNewApiForm({
      name: '',
      url: '',
      keys: [{ id: Date.now().toString(36) + '1', memo: '', key: '' }],
      description: '',
      category: apiCategoryFilter !== '全部分组' ? apiCategoryFilter : '默认',
    });
    setShowAddApiModal(true);
  };

  const handleAddApi = () => {
    if (!newApiForm.name.trim()) {
      showToast('请填写 API 名称', 'error');
      return;
    }
    if (!newApiForm.url.trim()) {
      showToast('请填写 API 地址', 'error');
      return;
    }

    const newEntry: ApiEntry = {
      id: Date.now().toString(36) + Math.random().toString(36).substring(2, 6),
      name: newApiForm.name.trim(),
      url: newApiForm.url.trim(),
      keys: newApiForm.keys,
      description: newApiForm.description.trim(),
      category: apiCategoryFilter !== '全部分组' ? apiCategoryFilter : (newApiForm.category || '默认'),
      createdAt: Date.now(),
      updatedAt: Date.now(),
    };

    updateAppData({
      ...appData,
      apis: [...apisList, newEntry],
    });

    setShowAddApiModal(false);
    showToast(`成功添加 API “${newEntry.name}”`, 'success');
  };

  const handleCreateNewApiCategory = () => {
    const trimmed = newApiGroupName.trim();
    if (!trimmed) return;
    if (trimmed === '全部分组') {
      showToast('"全部分组"为系统保留名称', 'error');
      return;
    }
    const currentCats = appData.apiCategories || ['默认'];
    if (currentCats.includes(trimmed)) {
      showToast('该分组名称已存在', 'error');
      return;
    }

    updateAppData({
      ...appData,
      apiCategories: [...currentCats, trimmed],
    });
    setNewApiGroupName('');
    setShowNewApiGroupModal(false);
    setApiCategoryFilter(trimmed);
    showToast(`成功新建分组: ${trimmed}`, 'success');
  };

  const handleRenameApiCategory = () => {
    if (!managingApiCategory) return;
    const trimmed = renameApiCategoryInput.trim();
    if (!trimmed) return;
    if (trimmed === '全部分组') {
      showToast('"全部分组"为系统保留名称', 'error');
      return;
    }

    const currentCats = appData.apiCategories || ['默认'];
    const updatedCats = currentCats.map((c) => (c === managingApiCategory ? trimmed : c));
    const updatedApis = apisList.map((p) =>
      (p.category || '默认') === managingApiCategory ? { ...p, category: trimmed } : p
    );

    updateAppData({
      ...appData,
      apis: updatedApis,
      apiCategories: updatedCats,
    });

    if (apiCategoryFilter === managingApiCategory) {
      setApiCategoryFilter(trimmed);
    }
    showToast(`分组已重命名为 "${trimmed}"`, 'success');
    setManagingApiCategory(null);
  };

  const handleConfirmDeleteApiCategory = () => {
    if (!managingApiCategory) return;

    let updatedApis = [...apisList];
    if (deleteItemsWithApiCategory) {
      updatedApis = updatedApis.filter((p) => (p.category || '默认') !== managingApiCategory);
    } else {
      updatedApis = updatedApis.map((p) =>
        (p.category || '默认') === managingApiCategory ? { ...p, category: '默认' } : p
      );
    }

    const currentCats = appData.apiCategories || ['默认'];
    const updatedCats = currentCats.filter((c) => c !== managingApiCategory);

    updateAppData({
      ...appData,
      apis: updatedApis,
      apiCategories: updatedCats,
    });

    showToast(`分组 "${managingApiCategory}" 已删除`, 'info');
    if (apiCategoryFilter === managingApiCategory) {
      setApiCategoryFilter('全部分组');
    }
    setManagingApiCategory(null);
    setShowDeleteApiCategoryConfirm(false);
    setDeleteItemsWithApiCategory(false);
  };

  const handleBatchDeleteApis = () => {
    if (selectedApiIds.length === 0) return;
    if (!confirm(`确定要删除选中的 ${selectedApiIds.length} 个 API 吗？`)) return;

    const updated = apisList.filter((p) => !selectedApiIds.includes(p.id));
    updateAppData({ ...appData, apis: updated });

    showToast(`已成功删除 ${selectedApiIds.length} 个 API`, 'info');
    setSelectedApiIds([]);
    setApiBatchMode(false);
  };

  const handleConfirmBatchMoveApis = () => {
    if (!batchTargetApiCategory) {
      showToast('请选择目标分组', 'error');
      return;
    }
    const target = batchTargetApiCategory;
    const updated = apisList.map((p) =>
      selectedApiIds.includes(p.id) ? { ...p, category: target } : p
    );

    updateAppData({ ...appData, apis: updated });

    setSelectedApiIds([]);
    setApiBatchMode(false);
    setShowApiBatchMoveModal(false);
    setBatchTargetApiCategory('');
    showToast(`已将 ${selectedApiIds.length} 个 API 移动到 “${target}”`, 'success');
  };

  // ==================== FONT STORAGE STATE & HANDLERS ====================
  const [fontSearchQuery, setFontSearchQuery] = useState('');
  const [fontCategoryFilter, setFontCategoryFilter] = useState('全部分组');
  const [fontBatchMode, setFontBatchMode] = useState(false);
  const [selectedFontIds, setSelectedFontIds] = useState<string[]>([]);

  // Selected Font for Preview
  const [activePreviewFontId, setActivePreviewFontId] = useState<string | null>(null);
  const [previewText, setPreviewText] = useState('永和九年，歲在癸丑，暮春之初 | The quick brown fox jumps over the lazy dog 1234567890');

  // Modals for Font
  const [showAddFontChoiceModal, setShowAddFontChoiceModal] = useState(false);
  const [showAddFontUrlModal, setShowAddFontUrlModal] = useState(false);
  const [newFontUrlForm, setNewFontUrlForm] = useState({ name: '', url: '', category: '默认' });

  // File Upload Ref for Font
  const fontFileInputRef = useRef<HTMLInputElement>(null);

  // Edit Font Modal State
  const [editingFont, setEditingFont] = useState<FontEntry | null>(null);

  // Group Modals State for Font
  const [showNewFontGroupModal, setShowNewFontGroupModal] = useState(false);
  const [newFontGroupName, setNewFontGroupName] = useState('');
  const [managingFontCategory, setManagingFontCategory] = useState<string | null>(null);
  const [renameFontCategoryInput, setRenameFontCategoryInput] = useState('');
  const [deleteItemsWithFontCategory, setDeleteItemsWithFontCategory] = useState(false);
  const [showFontBatchMoveModal, setShowFontBatchMoveModal] = useState(false);
  const [batchTargetFontCategory, setBatchTargetFontCategory] = useState('');

  const fontsList = appData.fonts || [];

  // Dynamic @font-face injection for custom fonts
  useEffect(() => {
    const fList = appData.fonts || [];
    let styleEl = document.getElementById('custom-fonts-style') as HTMLStyleElement | null;
    if (!styleEl) {
      styleEl = document.createElement('style');
      styleEl.id = 'custom-fonts-style';
      document.head.appendChild(styleEl);
    }

    const cssRules = fList
      .map((f) => {
        const src = f.fileData || f.url;
        if (!src) return '';
        return `@font-face {
  font-family: '${f.fontFamily}';
  src: url('${src}');
  font-display: swap;
}`;
      })
      .join('\n');

    styleEl.textContent = cssRules;
  }, [appData.fonts]);

  const filteredFonts = fontsList.filter((item) => {
    if (fontCategoryFilter !== '全部分组' && (item.category || '默认') !== fontCategoryFilter) {
      return false;
    }
    if (!fontSearchQuery.trim()) return true;
    const q = fontSearchQuery.toLowerCase();
    const nameMatch = item.name.toLowerCase().includes(q);
    const urlMatch = (item.url || '').toLowerCase().includes(q);
    const catMatch = (item.category || '').toLowerCase().includes(q);
    return nameMatch || urlMatch || catMatch;
  });

  const handleAddFontUrl = () => {
    if (!newFontUrlForm.name.trim()) {
      showToast('请填写字体名称', 'error');
      return;
    }
    if (!newFontUrlForm.url.trim()) {
      showToast('请填写字体 URL', 'error');
      return;
    }

    const fontFamilyName = 'Font_' + Date.now().toString(36);
    const newEntry: FontEntry = {
      id: Date.now().toString(36) + Math.random().toString(36).substring(2, 6),
      name: newFontUrlForm.name.trim(),
      url: newFontUrlForm.url.trim(),
      fontFamily: fontFamilyName,
      category: fontCategoryFilter !== '全部分组' ? fontCategoryFilter : (newFontUrlForm.category || '默认'),
      createdAt: Date.now(),
      updatedAt: Date.now(),
    };

    updateAppData({
      ...appData,
      fonts: [...fontsList, newEntry],
    });

    setShowAddFontUrlModal(false);
    setActivePreviewFontId(newEntry.id);
    setNewFontUrlForm({ name: '', url: '', category: '默认' });
    showToast(`成功添加字体 “${newEntry.name}”`, 'success');
  };

  const processFontFiles = async (files: File[]) => {
    const added: FontEntry[] = [];
    for (const file of files) {
      const rawFileName = file.name;
      const fontName = rawFileName.substring(0, rawFileName.lastIndexOf('.')) || rawFileName;
      try {
        const result = await new Promise<string>((resolve, reject) => {
          const reader = new FileReader();
          reader.onload = () => resolve((reader.result as string) || '');
          reader.onerror = () => reject(new Error('读取失败'));
          reader.readAsDataURL(file);
        });
        if (!result) throw new Error('读取失败');
        added.push({
          id: Date.now().toString(36) + Math.random().toString(36).substring(2, 6),
          name: fontName,
          fileData: result,
          fontFamily: 'Font_' + Date.now().toString(36) + Math.random().toString(36).slice(2, 5),
          category: fontCategoryFilter !== '全部分组' ? fontCategoryFilter : '默认',
          createdAt: Date.now(),
          updatedAt: Date.now(),
        });
      } catch {
        showToast(`读取字体文件 “${file.name}” 失败`, 'error');
      }
      // 大批量文件时主动让出主线程，避免 Android/低内存设备长时间卡死。
      await new Promise((resolve) => setTimeout(resolve, 0));
    }
    if (added.length) {
      updateAppData((prev) => ({ ...prev, fonts: [...(prev.fonts || []), ...added] }));
      setActivePreviewFontId(added[0].id);
      showToast(`已添加 ${added.length} 个字体`, 'success');
    }
    setShowAddFontChoiceModal(false);
  };

  const handleFileUploadFont = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = Array.from(e.target.files || []) as File[];
    if (!files.length) return;
    await processFontFiles(files);
    e.target.value = '';
  };

  const handleCreateNewFontCategory = () => {
    const trimmed = newFontGroupName.trim();
    if (!trimmed) return;
    if (trimmed === '全部分组') {
      showToast('"全部分组"为系统保留名称', 'error');
      return;
    }
    const currentCats = appData.fontCategories || ['默认'];
    if (currentCats.includes(trimmed)) {
      showToast('该分组名称已存在', 'error');
      return;
    }

    updateAppData({
      ...appData,
      fontCategories: [...currentCats, trimmed],
    });
    setNewFontGroupName('');
    setShowNewFontGroupModal(false);
    setFontCategoryFilter(trimmed);
    showToast(`成功新建分组: ${trimmed}`, 'success');
  };

  const handleRenameFontCategory = () => {
    if (!managingFontCategory) return;
    const trimmed = renameFontCategoryInput.trim();
    if (!trimmed) return;
    if (trimmed === '全部分组') {
      showToast('"全部分组"为系统保留名称', 'error');
      return;
    }

    const currentCats = appData.fontCategories || ['默认'];
    const updatedCats = currentCats.map((c) => (c === managingFontCategory ? trimmed : c));
    const updatedFonts = fontsList.map((p) =>
      (p.category || '默认') === managingFontCategory ? { ...p, category: trimmed } : p
    );

    updateAppData({
      ...appData,
      fonts: updatedFonts,
      fontCategories: updatedCats,
    });

    if (fontCategoryFilter === managingFontCategory) {
      setFontCategoryFilter(trimmed);
    }
    showToast(`分组已重命名为 "${trimmed}"`, 'success');
    setManagingFontCategory(null);
  };

  const handleConfirmDeleteFontCategory = () => {
    if (!managingFontCategory) return;

    let updatedFonts = [...fontsList];
    if (deleteItemsWithFontCategory) {
      updatedFonts = updatedFonts.filter((p) => (p.category || '默认') !== managingFontCategory);
    } else {
      updatedFonts = updatedFonts.map((p) =>
        (p.category || '默认') === managingFontCategory ? { ...p, category: '默认' } : p
      );
    }

    const currentCats = appData.fontCategories || ['默认'];
    const updatedCats = currentCats.filter((c) => c !== managingFontCategory);

    updateAppData({
      ...appData,
      fonts: updatedFonts,
      fontCategories: updatedCats,
    });

    showToast(`分组 "${managingFontCategory}" 已删除`, 'info');
    if (fontCategoryFilter === managingFontCategory) {
      setFontCategoryFilter('全部分组');
    }
    setManagingFontCategory(null);
    setDeleteItemsWithFontCategory(false);
  };

  const handleBatchDeleteFonts = () => {
    if (selectedFontIds.length === 0) return;
    if (!confirm(`确定要删除选中的 ${selectedFontIds.length} 个字体吗？`)) return;

    const updated = fontsList.filter((p) => !selectedFontIds.includes(p.id));
    updateAppData({ ...appData, fonts: updated });

    showToast(`已成功删除 ${selectedFontIds.length} 个字体`, 'info');
    setSelectedFontIds([]);
    setFontBatchMode(false);
  };

  const handleConfirmBatchMoveFonts = () => {
    if (!batchTargetFontCategory) {
      showToast('请选择目标分组', 'error');
      return;
    }
    const target = batchTargetFontCategory;
    const updated = fontsList.map((p) =>
      selectedFontIds.includes(p.id) ? { ...p, category: target } : p
    );

    updateAppData({ ...appData, fonts: updated });

    setSelectedFontIds([]);
    setFontBatchMode(false);
    setShowFontBatchMoveModal(false);
    setBatchTargetFontCategory('');
    showToast(`已将 ${selectedFontIds.length} 个字体移动到 “${target}”`, 'success');
  };

  const handleExportFontFile = (font: FontEntry) => {
    if (font.fileData) {
      const link = document.createElement('a');
      link.href = font.fileData;
      link.download = `${font.name}.ttf`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      showToast(`正在导出字体 “${font.name}.ttf”`, 'success');
    } else if (font.url) {
      const link = document.createElement('a');
      link.href = font.url;
      link.download = `${font.name}.ttf`;
      link.target = '_blank';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      showToast(`正在导出字体 “${font.name}.ttf”`, 'success');
    } else {
      showToast('该字体没有可导出的文件或 URL', 'error');
    }
  };

  // ==================== EXTRA STORIES STATE & HANDLERS ====================
  const [extraStorySearchQuery, setExtraStorySearchQuery] = useState('');
  const [extraStoryCategoryFilter, setExtraStoryCategoryFilter] = useState('全部分组');
  const [extraStoryBatchMode, setExtraStoryBatchMode] = useState(false);
  const [selectedExtraStoryIds, setSelectedExtraStoryIds] = useState<string[]>([]);

  // Modals for Extra Story
  const [showAddExtraStoryChoiceModal, setShowAddExtraStoryChoiceModal] = useState(false);
  const [showAddExtraStoryManualModal, setShowAddExtraStoryManualModal] = useState(false);
  const [newExtraStoryManualForm, setNewExtraStoryManualForm] = useState({
    title: '',
    author: '',
    content: '',
    category: '默认',
  });

  // File Upload Ref for Extra Story (.docx, .txt)
  const extraStoryFileInputRef = useRef<HTMLInputElement>(null);

  // Single-Page Popup / Editing Modal (Requirement 4 & 5)
  const [editingExtraStory, setEditingExtraStory] = useState<ExtraStoryEntry | null>(null);
  const [isContentExpanded, setIsContentExpanded] = useState(false);

  // Group Modals State for Extra Story
  const [showNewExtraStoryGroupModal, setShowNewExtraStoryGroupModal] = useState(false);
  const [newExtraStoryGroupName, setNewExtraStoryGroupName] = useState('');
  const [managingExtraStoryCategory, setManagingExtraStoryCategory] = useState<string | null>(null);
  const [renameExtraStoryCategoryInput, setRenameExtraStoryCategoryInput] = useState('');
  const [deleteItemsWithExtraStoryCategory, setDeleteItemsWithExtraStoryCategory] = useState(false);
  const [showExtraStoryBatchMoveModal, setShowExtraStoryBatchMoveModal] = useState(false);
  const [batchTargetExtraStoryCategory, setBatchTargetExtraStoryCategory] = useState('');
  const [extraStoryImportPreview, setExtraStoryImportPreview] = useState<ExtraStoryEntry[] | null>(null);
  const [selectedExtraStoryImportIds, setSelectedExtraStoryImportIds] = useState<string[]>([]);
  const [extraStoryImportVisibleCount, setExtraStoryImportVisibleCount] = useState(60);

  const extraStoriesList = appData.extraStories || [];

  const filteredExtraStories = extraStoriesList.filter((item) => {
    if (extraStoryCategoryFilter !== '全部分组' && (item.category || '默认') !== extraStoryCategoryFilter) {
      return false;
    }
    if (!extraStorySearchQuery.trim()) return true;
    const q = extraStorySearchQuery.toLowerCase();
    const titleMatch = item.title.toLowerCase().includes(q);
    const authorMatch = (item.author || '').toLowerCase().includes(q);
    const contentMatch = item.content.toLowerCase().includes(q);
    const catMatch = (item.category || '').toLowerCase().includes(q);
    return titleMatch || authorMatch || contentMatch || catMatch;
  });

  const handleAddExtraStoryManual = () => {
    if (!newExtraStoryManualForm.title.trim()) {
      showToast('请填写番外名称', 'error');
      return;
    }
    if (!newExtraStoryManualForm.content.trim()) {
      showToast('请填写番外内容', 'error');
      return;
    }

    const newEntry: ExtraStoryEntry = {
      id: Date.now().toString(36) + Math.random().toString(36).substring(2, 6),
      title: newExtraStoryManualForm.title.trim(),
      author: newExtraStoryManualForm.author.trim() || undefined,
      content: newExtraStoryManualForm.content.trim(),
      category: extraStoryCategoryFilter !== '全部分组' ? extraStoryCategoryFilter : (newExtraStoryManualForm.category || '默认'),
      createdAt: Date.now(),
      updatedAt: Date.now(),
    };

    updateAppData({
      ...appData,
      extraStories: [...extraStoriesList, newEntry],
    });

    setShowAddExtraStoryManualModal(false);
    setNewExtraStoryManualForm({ title: '', author: '', content: '', category: '默认' });
    showToast(`成功录入番外 “${newEntry.title}”`, 'success');
  };

  const parseMultiExtraStories = (baseName: string, text: string): { title: string; author?: string; content: string }[] => {
    const normalized = text.replace(/\r\n?/g, '\n').replace(/\u00a0/g, ' ').trim();
    if (!normalized) return [];

    // 优先识别明确的“番外/外传/小剧场/第X章”标题行，避免普通正文中的空行被误切。
    const lines = normalized.split('\n');
    const heading = /^\s*(?:[【\[]?\s*(?:番外|外传|小剧场|特别篇|番外篇)\s*(?:[一二三四五六七八九十百千万0-9]+)?\s*[】\]]?|第\s*[一二三四五六七八九十百千万0-9]+\s*[章节回幕]|chapter\s*\d+)\s*[:：._\-—–]?\s*(.*)$/i;
    const starts: number[] = [];
    lines.forEach((line, index) => {
      if (heading.test(line.trim())) starts.push(index);
    });

    if (starts.length > 1) {
      return starts.map((startLine, idx) => {
        const endLine = idx + 1 < starts.length ? starts[idx + 1] : lines.length;
        const block = lines.slice(startLine, endLine).join('\n').trim();
        const first = lines[startLine].trim();
        const match = first.match(heading);
        const suffix = match?.[1]?.trim();
        const title = suffix || first.replace(/^[【\[]|[】\]]$/g, '') || `${baseName} - 番外${idx + 1}`;
        const content = lines.slice(startLine + 1, endLine).join('\n').trim() || block;
        return { title, content };
      }).filter((item) => item.content.trim().length > 0);
    }

    // 再兼容常见的分隔线：=== / --- / ### / ***。
    const dividerParts = normalized.split(/\n\s*(?:={3,}|-{3,}|_{3,}|\*{3,}|#{2,})\s*\n/g).map((p) => p.trim()).filter(Boolean);
    if (dividerParts.length > 1) {
      return dividerParts.map((part, idx) => {
        const partLines = part.split('\n');
        const first = partLines[0].trim();
        const title = first.length > 0 && first.length <= 50 ? first : `${baseName} - 番外${idx + 1}`;
        const content = (first === title && partLines.length > 1 ? partLines.slice(1).join('\n') : part).trim();
        return { title, content: content || part };
      }).filter((item) => item.content.trim().length > 0);
    }

    return [{ title: baseName, content: normalized }];
  };

  const buildExtraStoryEntries = (fileName: string, rawText: string, category: string): ExtraStoryEntry[] => {
    const baseName = fileName.replace(/\.[^/.]+$/, '') || fileName;
    const parsedStories = parseMultiExtraStories(baseName, rawText);
    const now = Date.now();
    return parsedStories.map((story, index) => ({
      id: `${now.toString(36)}-${index}-${Math.random().toString(36).slice(2, 7)}`,
      title: story.title,
      author: story.author || undefined,
      content: story.content,
      category,
      createdAt: now + index,
      updatedAt: now + index,
    }));
  };

  const commitExtraStoryEntries = (entries: ExtraStoryEntry[]) => {
    if (!entries.length) {
      showToast('没有可导入的番外', 'error');
      return;
    }
    // 一次性提交，避免“大量条目逐条 setState + 持久化”造成移动端内存峰值。
    updateAppData((prev) => ({
      ...prev,
      extraStories: [...(prev.extraStories || []), ...entries],
    }));
    showToast(entries.length === 1 ? `已成功导入番外 “${entries[0].title}”` : `已导入 ${entries.length} 个番外`, 'success');
  };

  const processExtraStoryTextImport = (fileName: string, rawText: string) => {
    const targetCat = extraStoryCategoryFilter !== '全部分组' ? extraStoryCategoryFilter : '默认';
    const entries = buildExtraStoryEntries(fileName, rawText, targetCat);
    if (!entries.length) {
      showToast('文档内容为空', 'error');
      return;
    }
    if (entries.length === 1) {
      commitExtraStoryEntries(entries);
      setShowAddExtraStoryChoiceModal(false);
      return;
    }
    setExtraStoryImportPreview(entries);
    setSelectedExtraStoryImportIds(entries.map((e) => e.id));
    setExtraStoryImportVisibleCount(Math.min(60, entries.length));
    setShowAddExtraStoryChoiceModal(false);
  };

  const processExtraStoryFiles = async (files: File[]) => {
    const targetCat = extraStoryCategoryFilter !== '全部分组' ? extraStoryCategoryFilter : '默认';
    const allEntries: ExtraStoryEntry[] = [];
    for (const file of files) {
      try {
        const ext = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();
        let rawText = '';
        if (ext === '.docx') {
          const arrayBuffer = await file.arrayBuffer();
          const res = await mammoth.extractRawText({ arrayBuffer });
          rawText = res.value || '';
        } else {
          rawText = await file.text();
        }
        allEntries.push(...buildExtraStoryEntries(file.name, rawText, targetCat));
      } catch {
        showToast(`解析番外文件 “${file.name}” 失败`, 'error');
      }
      await new Promise((resolve) => setTimeout(resolve, 0));
    }
    if (!allEntries.length) {
      showToast('没有识别到可导入的番外', 'error');
      return;
    }
    if (allEntries.length === 1) {
      commitExtraStoryEntries(allEntries);
      return;
    }
    setExtraStoryImportPreview(allEntries);
    setSelectedExtraStoryImportIds(allEntries.map((e) => e.id));
    setExtraStoryImportVisibleCount(Math.min(60, allEntries.length));
  };

  const toggleExtraStoryImportSelection = (id: string) => {
    setSelectedExtraStoryImportIds((ids) => ids.includes(id) ? ids.filter((x) => x !== id) : [...ids, id]);
  };

  const updateExtraStoryImportPreviewItem = (id: string, patch: Partial<ExtraStoryEntry>) => {
    setExtraStoryImportPreview((prev) => prev ? prev.map((item) => item.id === id ? { ...item, ...patch, updatedAt: Date.now() } : item) : prev);
  };

  const mergeSelectedExtraStoryImports = () => {
    if (!extraStoryImportPreview || selectedExtraStoryImportIds.length < 2) {
      showToast('请至少选择两条番外再融合', 'error');
      return;
    }
    const selectedSet = new Set(selectedExtraStoryImportIds);
    const selected = extraStoryImportPreview.filter((item) => selectedSet.has(item.id));
    if (selected.length < 2) return;
    const firstIndex = extraStoryImportPreview.findIndex((item) => item.id === selected[0].id);
    const merged: ExtraStoryEntry = {
      ...selected[0],
      title: selected[0].title || `融合番外-${firstIndex + 1}`,
      content: selected.map((item) => item.content.trim()).filter(Boolean).join('\n\n'),
      updatedAt: Date.now(),
    };
    const remaining = extraStoryImportPreview.filter((item) => !selectedSet.has(item.id));
    remaining.splice(Math.min(firstIndex, remaining.length), 0, merged);
    setExtraStoryImportPreview(remaining);
    setSelectedExtraStoryImportIds([merged.id]);
    showToast(`已将 ${selected.length} 条番外融合为 1 条`, 'success');
  };

  const commitSelectedExtraStoryImports = () => {
    const selected = (extraStoryImportPreview || []).filter((item) => selectedExtraStoryImportIds.includes(item.id) && item.content.trim());
    if (!selected.length) {
      showToast('请至少选择一条有内容的番外', 'error');
      return;
    }
    commitExtraStoryEntries(selected.map((item) => ({ ...item, title: item.title.trim() || '未命名番外', content: item.content.trim(), updatedAt: Date.now() })));
    setExtraStoryImportPreview(null);
    setSelectedExtraStoryImportIds([]);
    setExtraStoryImportVisibleCount(60);
  };

  const handleFileUploadExtraStory = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = Array.from(e.target.files || []) as File[];
    e.target.value = '';
    if (!files.length) return;
    showToast(`正在读取 ${files.length} 个番外文件…`, 'info');
    await processExtraStoryFiles(files);
  };

  const handleCreateNewExtraStoryGroup = () => {
    const trimmed = newExtraStoryGroupName.trim();
    if (!trimmed) return;
    if (trimmed === '全部分组') {
      showToast('"全部分组"为系统保留名称', 'error');
      return;
    }
    const currentCats = appData.extraStoryCategories || ['默认'];
    if (currentCats.includes(trimmed)) {
      showToast('该分组名称已存在', 'error');
      return;
    }

    updateAppData({
      ...appData,
      extraStoryCategories: [...currentCats, trimmed],
    });
    setNewExtraStoryGroupName('');
    setShowNewExtraStoryGroupModal(false);
    setExtraStoryCategoryFilter(trimmed);
    showToast(`成功新建分组: ${trimmed}`, 'success');
  };

  const handleRenameExtraStoryCategory = () => {
    if (!managingExtraStoryCategory) return;
    const trimmed = renameExtraStoryCategoryInput.trim();
    if (!trimmed) return;
    if (trimmed === '全部分组') {
      showToast('"全部分组"为系统保留名称', 'error');
      return;
    }

    const currentCats = appData.extraStoryCategories || ['默认'];
    const updatedCats = currentCats.map((c) => (c === managingExtraStoryCategory ? trimmed : c));
    const updatedStories = extraStoriesList.map((p) =>
      (p.category || '默认') === managingExtraStoryCategory ? { ...p, category: trimmed } : p
    );

    updateAppData({
      ...appData,
      extraStories: updatedStories,
      extraStoryCategories: updatedCats,
    });

    if (extraStoryCategoryFilter === managingExtraStoryCategory) {
      setExtraStoryCategoryFilter(trimmed);
    }
    showToast(`分组已重命名为 "${trimmed}"`, 'success');
    setManagingExtraStoryCategory(null);
  };

  const handleConfirmDeleteExtraStoryCategory = () => {
    if (!managingExtraStoryCategory) return;

    let updatedStories = [...extraStoriesList];
    if (deleteItemsWithExtraStoryCategory) {
      updatedStories = updatedStories.filter((p) => (p.category || '默认') !== managingExtraStoryCategory);
    } else {
      updatedStories = updatedStories.map((p) =>
        (p.category || '默认') === managingExtraStoryCategory ? { ...p, category: '默认' } : p
      );
    }

    const currentCats = appData.extraStoryCategories || ['默认'];
    const updatedCats = currentCats.filter((c) => c !== managingExtraStoryCategory);

    updateAppData({
      ...appData,
      extraStories: updatedStories,
      extraStoryCategories: updatedCats,
    });

    showToast(`分组 "${managingExtraStoryCategory}" 已删除`, 'info');
    if (extraStoryCategoryFilter === managingExtraStoryCategory) {
      setExtraStoryCategoryFilter('全部分组');
    }
    setManagingExtraStoryCategory(null);
    setDeleteItemsWithExtraStoryCategory(false);
  };

  const handleBatchDeleteExtraStories = () => {
    if (selectedExtraStoryIds.length === 0) return;
    if (!confirm(`确定要删除选中的 ${selectedExtraStoryIds.length} 个番外吗？`)) return;

    const updated = extraStoriesList.filter((p) => !selectedExtraStoryIds.includes(p.id));
    updateAppData({ ...appData, extraStories: updated });

    showToast(`已成功删除 ${selectedExtraStoryIds.length} 个番外`, 'info');
    setSelectedExtraStoryIds([]);
    setExtraStoryBatchMode(false);
  };

  const handleConfirmBatchMoveExtraStories = () => {
    if (!batchTargetExtraStoryCategory) {
      showToast('请选择目标分组', 'error');
      return;
    }
    const target = batchTargetExtraStoryCategory;
    const updated = extraStoriesList.map((p) =>
      selectedExtraStoryIds.includes(p.id) ? { ...p, category: target } : p
    );

    updateAppData({ ...appData, extraStories: updated });

    setSelectedExtraStoryIds([]);
    setExtraStoryBatchMode(false);
    setShowExtraStoryBatchMoveModal(false);
    setBatchTargetExtraStoryCategory('');
    showToast(`已将 ${selectedExtraStoryIds.length} 个番外移动到 “${target}”`, 'success');
  };

  // ==================== CHAT MEMES (聊天梗) STATE & HANDLERS ====================
  const [chatMemeSearchQuery, setChatMemeSearchQuery] = useState('');
  const [chatMemeSortOrder, setChatMemeSortOrder] = useState<'default' | 'az' | 'za' | 'newest' | 'oldest'>('default');
  const [chatMemeCategoryFilter, setChatMemeCategoryFilter] = useState('全部分组');
  const [chatMemeBatchMode, setChatMemeBatchMode] = useState(false);
  const [selectedChatMemeIds, setSelectedChatMemeIds] = useState<string[]>([]);
  const [selectedChatImportIds, setSelectedChatImportIds] = useState<string[]>([]);
  const [showChatAddChoiceModal, setShowChatAddChoiceModal] = useState(false);
  const [showChatManualModal, setShowChatManualModal] = useState(false);
  const [chatManualInput, setChatManualInput] = useState('');
  const [chatImportPreview, setChatImportPreview] = useState<ChatMemeEntry[] | null>(null);
  const [chatImportSourceName, setChatImportSourceName] = useState('');
  const [editingChatMeme, setEditingChatMeme] = useState<ChatMemeEntry | null>(null);
  const [showNewChatMemeGroupModal, setShowNewChatMemeGroupModal] = useState(false);
  const [newChatMemeGroupName, setNewChatMemeGroupName] = useState('');
  const [managingChatMemeGroup, setManagingChatMemeGroup] = useState<string | null>(null);
  const [renameChatMemeGroupInput, setRenameChatMemeGroupInput] = useState('');
  const [deleteCardsWithChatMemeGroup, setDeleteCardsWithChatMemeGroup] = useState(false);
  const [chatMemeExportFormat, setChatMemeExportFormat] = useState<'docx' | 'txt' | 'json'>('txt');
  const [showChatExportModal, setShowChatExportModal] = useState(false);
  const chatMemeFileInputRef = useRef<HTMLInputElement>(null);
  const chatMemeGroupPressTimer = useRef<ReturnType<typeof setTimeout> | null>(null);

  const chatMemesList = appData.chatMemes || [];
  const filteredChatMemes = chatMemesList.filter((item) => {
    if (chatMemeCategoryFilter !== '全部分组' && (item.category || '默认') !== chatMemeCategoryFilter) return false;
    const q = chatMemeSearchQuery.trim().toLowerCase();
    if (!q) return true;
    return item.content.toLowerCase().includes(q) || (item.category || '').toLowerCase().includes(q);
  }).sort((a, b) => {
    if (chatMemeSortOrder === 'az') return a.content.localeCompare(b.content, undefined, { numeric: true, sensitivity: 'base' });
    if (chatMemeSortOrder === 'za') return b.content.localeCompare(a.content, undefined, { numeric: true, sensitivity: 'base' });
    if (chatMemeSortOrder === 'newest') return (b.createdAt || 0) - (a.createdAt || 0);
    if (chatMemeSortOrder === 'oldest') return (a.createdAt || 0) - (b.createdAt || 0);
    return 0;
  });

  const splitChatMemeText = (text: string): string[] => {
    return text
      .replace(/\r\n/g, '\n')
      .split(/\n\s*\n+/)
      .map((part) => part.trim())
      .filter(Boolean);
  };

  const makeChatMemeEntries = (contents: string[], category = '默认'): ChatMemeEntry[] => {
    const now = Date.now();
    return contents.filter(Boolean).map((content, index) => ({
      id: `${now.toString(36)}-${Math.random().toString(36).slice(2, 9)}-${index}`,
      content: content.trim(),
      category,
      createdAt: Date.now() + index,
      updatedAt: Date.now() + index,
    }));
  };

  const commitChatMemeImport = () => {
    const pending = (chatImportPreview || []).filter((m) => selectedChatImportIds.includes(m.id));
    if (!pending.length) {
      showToast('没有可导入的聊天梗', 'error');
      return;
    }
    updateAppData((prev) => ({ ...prev, chatMemes: [...(prev.chatMemes || []), ...pending] }));
    setChatImportPreview(null);
    setSelectedChatImportIds([]);
    setChatImportSourceName('');
    setShowChatManualModal(false);
    setChatManualInput('');
    showToast(`已导入 ${pending.length} 条聊天梗`, 'success');
  };

  const updateChatImportPreviewItem = (id: string, content: string) => {
    setChatImportPreview((prev) => prev ? prev.map((m) => m.id === id ? { ...m, content, updatedAt: Date.now() } : m) : prev);
  };

  const toggleChatImportSelection = (id: string) => {
    setSelectedChatImportIds((ids) => ids.includes(id) ? ids.filter((x) => x !== id) : [...ids, id]);
  };

  const mergeSelectedChatImports = () => {
    if (!chatImportPreview || selectedChatImportIds.length < 2) {
      showToast('请至少选择两条聊天梗再合并', 'error');
      return;
    }
    const selected = chatImportPreview.filter((m) => selectedChatImportIds.includes(m.id));
    if (selected.length < 2) return;
    const firstIndex = chatImportPreview.findIndex((m) => m.id === selected[0].id);
    const merged: ChatMemeEntry = {
      ...selected[0],
      content: selected.map((m) => m.content.trim()).filter(Boolean).join('\n\n'),
      updatedAt: Date.now(),
    };
    const selectedSet = new Set(selected.map((m) => m.id));
    const next = chatImportPreview.filter((m) => !selectedSet.has(m.id));
    next.splice(Math.min(firstIndex, next.length), 0, merged);
    setChatImportPreview(next);
    setSelectedChatImportIds([]);
    showToast(`已将 ${selected.length} 条聊天梗合并`, 'success');
  };

  const handleChatManualPreview = () => {
    const contents = splitChatMemeText(chatManualInput);
    if (!contents.length) {
      showToast('请先填写聊天梗内容，梗与梗之间用空行分隔', 'error');
      return;
    }
    const category = chatMemeCategoryFilter !== '全部分组' ? chatMemeCategoryFilter : '默认';
    const previewEntries = makeChatMemeEntries(contents, category);
    setChatImportPreview(previewEntries);
    setSelectedChatImportIds(previewEntries.map((m) => m.id));
    setChatImportSourceName('手动录入');
  };

  const handleChatMemeFileSelect = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const input = e.currentTarget;
    const file = input.files?.[0];
    input.value = '';
    if (!file) return;
    setShowChatAddChoiceModal(false);
    showToast(`正在读取：${file.name}`, 'info');
    try {
      const ext = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();
      let rawText = '';
      if (ext === '.docx') {
        const arrayBuffer = await file.arrayBuffer();
        const res = await mammoth.extractRawText({ arrayBuffer });
        rawText = res.value || '';
      } else {
        rawText = await file.text();
      }
      const contents = splitChatMemeText(rawText);
      if (!contents.length && rawText.trim()) contents.push(rawText.trim());
      if (!contents.length) {
        showToast('文件中没有可识别的聊天梗内容', 'error');
        return;
      }
      const category = chatMemeCategoryFilter !== '全部分组' ? chatMemeCategoryFilter : '默认';
      const previewEntries = makeChatMemeEntries(contents, category);
      setChatImportPreview(previewEntries);
      setSelectedChatImportIds(previewEntries.map((m) => m.id));
      setChatImportSourceName(file.name);
    } catch (err) {
      showToast(`读取文件失败：${file.name}`, 'error');
    }
  };

  const handleCreateChatMemeGroup = () => {
    const name = newChatMemeGroupName.trim();
    if (!name || name === '全部分组') return;
    const cats = appData.chatMemeCategories || ['默认'];
    if (cats.includes(name)) {
      showToast('该分组已经存在', 'error');
      return;
    }
    updateAppData((prev) => ({ ...prev, chatMemeCategories: [...(prev.chatMemeCategories || ['默认']), name] }));
    setNewChatMemeGroupName('');
    setShowNewChatMemeGroupModal(false);
    setChatMemeCategoryFilter(name);
  };

  const handleRenameChatMemeGroup = () => {
    if (!managingChatMemeGroup) return;
    const name = renameChatMemeGroupInput.trim();
    if (!name || name === '全部分组') return;
    const old = managingChatMemeGroup;
    updateAppData((prev) => ({
      ...prev,
      chatMemeCategories: (prev.chatMemeCategories || ['默认']).map((g) => (g === old ? name : g)),
      chatMemes: (prev.chatMemes || []).map((m) => (m.category || '默认') === old ? { ...m, category: name, updatedAt: Date.now() } : m),
    }));
    if (chatMemeCategoryFilter === old) setChatMemeCategoryFilter(name);
    setManagingChatMemeGroup(null);
  };

  const handleDeleteChatMemeGroup = () => {
    if (!managingChatMemeGroup) return;
    const group = managingChatMemeGroup;
    updateAppData((prev) => ({
      ...prev,
      chatMemeCategories: (prev.chatMemeCategories || ['默认']).filter((g) => g !== group),
      chatMemes: (prev.chatMemes || []).filter((m) => (m.category || '默认') !== group),
      ...(deleteCardsWithChatMemeGroup ? { cards: prev.cards.filter((c) => c.group !== group) } : {}),
    }));
    if (chatMemeCategoryFilter === group) setChatMemeCategoryFilter('全部分组');
    setManagingChatMemeGroup(null);
    setDeleteCardsWithChatMemeGroup(false);
    showToast(`分组“${group}”已删除`, 'info');
  };

  const handleDeleteSelectedChatMemes = () => {
    if (!selectedChatMemeIds.length) return;
    updateAppData((prev) => ({ ...prev, chatMemes: (prev.chatMemes || []).filter((m) => !selectedChatMemeIds.includes(m.id)) }));
    showToast(`已删除 ${selectedChatMemeIds.length} 条聊天梗`, 'info');
    setSelectedChatMemeIds([]);
    setChatMemeBatchMode(false);
  };

  const handleMoveSelectedChatMemes = (target: string) => {
    if (!target || !selectedChatMemeIds.length) return;
    updateAppData((prev) => ({
      ...prev,
      chatMemes: (prev.chatMemes || []).map((m) => selectedChatMemeIds.includes(m.id) ? { ...m, category: target, updatedAt: Date.now() } : m),
    }));
    setSelectedChatMemeIds([]);
    setChatMemeBatchMode(false);
  };

  const exportSelectedChatMemes = async () => {
    const selected = chatMemesList.filter((m) => selectedChatMemeIds.includes(m.id));
    if (!selected.length) return;
    const baseName = selected.length === 1 ? '聊天梗' : `聊天梗-${selected.length}条`;
    if (chatMemeExportFormat === 'json') {
      downloadJsonFile(`${baseName}.json`, selected);
    } else if (chatMemeExportFormat === 'txt') {
      const blob = new Blob([selected.map((m) => m.content).join('\n\n')], { type: 'text/plain;charset=utf-8' });
      const url = URL.createObjectURL(blob); const a = document.createElement('a'); a.href = url; a.download = `${baseName}.txt`; a.click(); URL.revokeObjectURL(url);
    } else {
      const paragraphs = selected.flatMap((m, i) => [new Paragraph({ children: [new TextRun(m.content)] }), ...(i < selected.length - 1 ? [new Paragraph('')] : [])]);
      const doc = new Document({ sections: [{ properties: {}, children: paragraphs }] });
      const blob = await Packer.toBlob(doc);
      const url = URL.createObjectURL(blob); const a = document.createElement('a'); a.href = url; a.download = `${baseName}.docx`; a.click(); URL.revokeObjectURL(url);
    }
    setShowChatExportModal(false);
    showToast(`已导出 ${selected.length} 条聊天梗`, 'success');
  };

  const handleSaveChatMemeEdit = () => {
    if (!editingChatMeme || !editingChatMeme.content.trim()) return;
    updateAppData((prev) => ({ ...prev, chatMemes: (prev.chatMemes || []).map((m) => m.id === editingChatMeme.id ? { ...editingChatMeme, content: editingChatMeme.content.trim(), updatedAt: Date.now() } : m) }));
    setEditingChatMeme(null);
    showToast('聊天梗已保存', 'success');
  };

  const handleDeleteChatMeme = () => {
    if (!editingChatMeme) return;
    updateAppData((prev) => ({ ...prev, chatMemes: (prev.chatMemes || []).filter((m) => m.id !== editingChatMeme.id) }));
    setEditingChatMeme(null);
    showToast('聊天梗已删除', 'info');
  };

  // ==================== STICKERS (表情包) STATE & HANDLERS ====================
  const [stickerSearchQuery, setStickerSearchQuery] = useState('');
  const [stickerCategoryFilter, setStickerCategoryFilter] = useState('全部分组');
  const [stickerBatchMode, setStickerBatchMode] = useState(false);
  const [selectedStickerPackIds, setSelectedStickerPackIds] = useState<string[]>([]);

  // File Upload Ref for Stickers (.docx, .txt, .json)
  const stickerFileInputRef = useRef<HTMLInputElement>(null);

  // Group Modals State for Stickers
  const [showNewStickerGroupModal, setShowNewStickerGroupModal] = useState(false);
  const [newStickerGroupName, setNewStickerGroupName] = useState('');
  const [managingStickerCategory, setManagingStickerCategory] = useState<string | null>(null);
  const [renameStickerCategoryInput, setRenameStickerCategoryInput] = useState('');
  const [deleteItemsWithStickerCategory, setDeleteItemsWithStickerCategory] = useState(false);
  const [showStickerBatchMoveModal, setShowStickerBatchMoveModal] = useState(false);
  const [batchTargetStickerCategory, setBatchTargetStickerCategory] = useState('');

  // Sticker Pack Detail & Edit Modal
  const [editingStickerPack, setEditingStickerPack] = useState<StickerPackEntry | null>(null);
  const [stickerDetailTab, setStickerDetailTab] = useState<'info' | 'atlas' | 'export'>('info');
  const [stickerExportFormat, setStickerExportFormat] = useState<
    'cnColon' | 'enColon' | 'noSpace' | 'space' | 'urlSpaceName' | 'urlName'
  >('cnColon');

  const stickerPacksList = appData.stickerPacks || [];

  const filteredStickerPacks = stickerPacksList.filter((pack) => {
    if (stickerCategoryFilter !== '全部分组' && (pack.category || '默认') !== stickerCategoryFilter) {
      return false;
    }
    if (!stickerSearchQuery.trim()) return true;
    const q = stickerSearchQuery.toLowerCase();
    const titleMatch = pack.title.toLowerCase().includes(q);
    const authorMatch = (pack.author || '').toLowerCase().includes(q);
    const catMatch = (pack.category || '').toLowerCase().includes(q);
    const itemMatch = pack.items.some(
      (it) => it.name.toLowerCase().includes(q) || it.url.toLowerCase().includes(q)
    );
    return titleMatch || authorMatch || catMatch || itemMatch;
  });

  const parseTextToStickerItems = (text: string): StickerItem[] => {
    const lines = text.split('\n').map((l) => l.trim()).filter(Boolean);
    const items: StickerItem[] = [];

    lines.forEach((line, idx) => {
      const urlMatch = line.match(/(https?:\/\/[^\s]+|data:image\/[^\s]+)/i);
      if (urlMatch) {
        const url = urlMatch[0];
        let name = line.replace(url, '').replace(/[：:\s]/g, '').trim();
        if (!name) name = `表情 ${idx + 1}`;
        items.push({
          id: `stk_${Date.now()}_${idx}_${Math.random().toString(36).substring(2, 6)}`,
          name,
          url,
        });
      }
    });

    return items;
  };

  const processStickerFileImport = async (file: File): Promise<StickerPackEntry | null> => {
    const fileName = file.name;
    const baseName = fileName.replace(/\.[^/.]+$/, '') || fileName;
    let items: StickerItem[] = [];
    try {
      if (fileName.toLowerCase().endsWith('.json')) {
        const text = await file.text();
        const parsed = JSON.parse(text);
        if (Array.isArray(parsed)) {
          items = parsed.map((it: any, idx: number) => ({ id: it.id || `stk_${Date.now()}_${idx}`, name: it.name || it.title || `表情 ${idx + 1}`, url: it.url || it.src || (typeof it === 'string' ? it : '') })).filter((it: StickerItem) => it.url);
        } else if (typeof parsed === 'object' && parsed !== null) {
          if (Array.isArray(parsed.items)) {
            items = parsed.items.map((it: any, idx: number) => ({ id: it.id || `stk_${Date.now()}_${idx}`, name: it.name || `表情 ${idx + 1}`, url: it.url || '' })).filter((it: StickerItem) => it.url);
          } else {
            Object.entries(parsed).forEach(([key, val], idx) => {
              if (typeof val === 'string' && (val.startsWith('http') || val.startsWith('data:image'))) items.push({ id: `stk_${Date.now()}_${idx}`, name: key, url: val });
            });
          }
        }
      } else if (fileName.toLowerCase().endsWith('.docx')) {
        const arrayBuffer = await file.arrayBuffer();
        const res = await mammoth.extractRawText({ arrayBuffer });
        items = parseTextToStickerItems(res.value || '');
      } else {
        items = parseTextToStickerItems(await file.text());
      }
      if (!items.length) {
        showToast(`“${file.name}” 未解析到有效的表情包 URL`, 'error');
        return null;
      }
      const targetCat = stickerCategoryFilter !== '全部分组' ? stickerCategoryFilter : '默认';
      return { id: Date.now().toString(36) + Math.random().toString(36).substring(2, 6), title: baseName, author: '默认', category: targetCat, items, createdAt: Date.now(), updatedAt: Date.now() };
    } catch (err: any) {
      showToast(`解析文件 “${file.name}” 失败: ${err.message || '格式不符'}`, 'error');
      return null;
    }
  };

  const processStickerFiles = async (files: File[]) => {
    const packs: StickerPackEntry[] = [];
    for (const file of files) {
      const pack = await processStickerFileImport(file);
      if (pack) packs.push(pack);
      await new Promise((resolve) => setTimeout(resolve, 0));
    }
    if (packs.length) {
      updateAppData((prev) => ({ ...prev, stickerPacks: [...(prev.stickerPacks || []), ...packs] }));
      showToast(`成功导入 ${packs.length} 个表情包图集`, 'success');
    }
  };

  const handleFileUploadSticker = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = Array.from(e.target.files || []) as File[];
    if (!files.length) return;
    await processStickerFiles(files);
    e.target.value = '';
  };

  const handleCreateNewStickerGroup = () => {
    const trimmed = newStickerGroupName.trim();
    if (!trimmed) return;
    if (trimmed === '全部分组') {
      showToast('"全部分组"为系统保留名称', 'error');
      return;
    }
    const currentCats = appData.stickerCategories || ['默认'];
    if (currentCats.includes(trimmed)) {
      showToast('该分组名称已存在', 'error');
      return;
    }

    updateAppData({
      ...appData,
      stickerCategories: [...currentCats, trimmed],
    });
    setNewStickerGroupName('');
    setShowNewStickerGroupModal(false);
    setStickerCategoryFilter(trimmed);
    showToast(`成功新建分组: ${trimmed}`, 'success');
  };

  const handleRenameStickerCategory = () => {
    if (!managingStickerCategory) return;
    const trimmed = renameStickerCategoryInput.trim();
    if (!trimmed) return;
    if (trimmed === '全部分组') {
      showToast('"全部分组"为系统保留名称', 'error');
      return;
    }

    const currentCats = appData.stickerCategories || ['默认'];
    const updatedCats = currentCats.map((c) => (c === managingStickerCategory ? trimmed : c));
    const updatedPacks = stickerPacksList.map((p) =>
      (p.category || '默认') === managingStickerCategory ? { ...p, category: trimmed } : p
    );

    updateAppData({
      ...appData,
      stickerPacks: updatedPacks,
      stickerCategories: updatedCats,
    });

    if (stickerCategoryFilter === managingStickerCategory) {
      setStickerCategoryFilter(trimmed);
    }
    showToast(`分组已重命名为 "${trimmed}"`, 'success');
    setManagingStickerCategory(null);
  };

  const handleConfirmDeleteStickerCategory = () => {
    if (!managingStickerCategory) return;

    let updatedPacks = [...stickerPacksList];
    if (deleteItemsWithStickerCategory) {
      updatedPacks = updatedPacks.filter((p) => (p.category || '默认') !== managingStickerCategory);
    } else {
      updatedPacks = updatedPacks.map((p) =>
        (p.category || '默认') === managingStickerCategory ? { ...p, category: '默认' } : p
      );
    }

    const currentCats = appData.stickerCategories || ['默认'];
    const updatedCats = currentCats.filter((c) => c !== managingStickerCategory);

    updateAppData({
      ...appData,
      stickerPacks: updatedPacks,
      stickerCategories: updatedCats,
    });

    showToast(`分组 "${managingStickerCategory}" 已删除`, 'info');
    if (stickerCategoryFilter === managingStickerCategory) {
      setStickerCategoryFilter('全部分组');
    }
    setManagingStickerCategory(null);
    setDeleteItemsWithStickerCategory(false);
  };

  const handleBatchDeleteStickerPacks = () => {
    if (selectedStickerPackIds.length === 0) return;
    if (!confirm(`确定要删除选中的 ${selectedStickerPackIds.length} 个表情包图集吗？`)) return;

    const updated = stickerPacksList.filter((p) => !selectedStickerPackIds.includes(p.id));
    updateAppData({ ...appData, stickerPacks: updated });

    showToast(`已成功删除 ${selectedStickerPackIds.length} 个表情包图集`, 'info');
    setSelectedStickerPackIds([]);
    setStickerBatchMode(false);
  };

  const handleConfirmBatchMoveStickerPacks = () => {
    if (!batchTargetStickerCategory) {
      showToast('请选择目标分组', 'error');
      return;
    }
    const target = batchTargetStickerCategory;
    const updated = stickerPacksList.map((p) =>
      selectedStickerPackIds.includes(p.id) ? { ...p, category: target } : p
    );

    updateAppData({ ...appData, stickerPacks: updated });
    setSelectedStickerPackIds([]);
    setStickerBatchMode(false);
    setShowStickerBatchMoveModal(false);
    setBatchTargetStickerCategory('');
    showToast(`已将 ${selectedStickerPackIds.length} 个图集移动到 “${target}”`, 'success');
  };

  const generateStickerExportText = (pack: StickerPackEntry, formatKey: string) => {
    return pack.items
      .map((item) => {
        const name = item.name || '';
        const url = item.url || '';
        switch (formatKey) {
          case 'cnColon': return `${name}：${url}`;
          case 'enColon': return `${name}:${url}`;
          case 'noSpace': return `${name}${url}`;
          case 'space': return `${name} ${url}`;
          case 'urlSpaceName': return `${url} ${name}`;
          case 'urlName': return `${url}${name}`;
          default: return `${name}：${url}`;
        }
      })
      .join('\n');
  };

  const exportStickerPackDocx = (pack: StickerPackEntry, formatKey: string) => {
    const textLines = generateStickerExportText(pack, formatKey).split('\n');
    const htmlContent = `<!DOCTYPE html><html><head><meta charset="utf-8"></head><body>${textLines.map((l) => `<p>${l}</p>`).join('')}</body></html>`;
    const blob = new Blob([htmlContent], { type: 'application/msword;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${pack.title || '表情包图集'}.docx`;
    a.click();
    URL.revokeObjectURL(url);
    showToast('已导出 .docx 文件', 'success');
  };

  const exportStickerPackTxt = (pack: StickerPackEntry, formatKey: string) => {
    const content = generateStickerExportText(pack, formatKey);
    const blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${pack.title || '表情包图集'}.txt`;
    a.click();
    URL.revokeObjectURL(url);
    showToast('已导出 .txt 文件', 'success');
  };

  const exportStickerPackJson = (pack: StickerPackEntry) => {
    const jsonStr = JSON.stringify(pack, null, 2);
    const blob = new Blob([jsonStr], { type: 'application/json;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${pack.title || '表情包图集'}.json`;
    a.click();
    URL.revokeObjectURL(url);
    showToast('已导出 .json 文件', 'success');
  };

  // ==================== WORLD BOOKS (世界书) STATE & HANDLERS ====================
  const [worldBookSearchQuery, setWorldBookSearchQuery] = useState('');
  const [worldBookCategoryFilter, setWorldBookCategoryFilter] = useState('全部分组');
  const [worldBookBatchMode, setWorldBookBatchMode] = useState(false);
  const [selectedWorldBookIds, setSelectedWorldBookIds] = useState<string[]>([]);

  // File Upload Ref for World Books (.docx, .txt, .json)
  const worldBookFileInputRef = useRef<HTMLInputElement>(null);

  // Group Modals State for World Books
  const [showNewWorldBookGroupModal, setShowNewWorldBookGroupModal] = useState(false);
  const [newWorldBookGroupName, setNewWorldBookGroupName] = useState('');
  const [managingWorldBookCategory, setManagingWorldBookCategory] = useState<string | null>(null);
  const [renameWorldBookCategoryInput, setRenameWorldBookCategoryInput] = useState('');
  const [deleteItemsWithWorldBookCategory, setDeleteItemsWithWorldBookCategory] = useState(false);
  const [showWorldBookBatchMoveModal, setShowWorldBookBatchMoveModal] = useState(false);
  const [batchTargetWorldBookCategory, setBatchTargetWorldBookCategory] = useState('');

  // World Book Detail & Edit Modal
  const [editingWorldBook, setEditingWorldBook] = useState<WorldBookEntry | null>(null);
  const [isWorldBookContentExpanded, setIsWorldBookContentExpanded] = useState(false);

  // ST 角色卡世界书：新增条目弹窗
  const [showAddCardWorldBookEntryModal, setShowAddCardWorldBookEntryModal] = useState(false);
  const [newCardWorldBookEntryForm, setNewCardWorldBookEntryForm] = useState({
    comment: '',
    keys: '',
    content: '',
  });

  const worldBooksList = appData.worldBooks || [];

  const filteredWorldBooks = worldBooksList.filter((wb) => {
    if (worldBookCategoryFilter !== '全部分组' && (wb.category || '默认') !== worldBookCategoryFilter) {
      return false;
    }
    if (!worldBookSearchQuery.trim()) return true;
    const q = worldBookSearchQuery.toLowerCase();
    const titleMatch = wb.title.toLowerCase().includes(q);
    const authorMatch = (wb.author || '').toLowerCase().includes(q);
    const catMatch = (wb.category || '').toLowerCase().includes(q);
    const contentMatch = (wb.content || '').toLowerCase().includes(q);
    return titleMatch || authorMatch || catMatch || contentMatch;
  });

  const extractWorldBookJsonDocument = (parsed: any, fileName: string) => {
    const normalized = normalizeWorldBookEntriesFromJson(parsed);
    const book = normalized.book || {};
    const entries = normalized.entries || [];
    return {
      title: book.name || book.title || parsed?.name || parsed?.title || fileName.replace(/\.[^/.]+$/, '') || '未命名世界书',
      author: book.author || parsed?.author || '默认',
      entries,
      raw: parsed,
    };
  };

  const commitJsonWorldBookImport = () => {
    if (!pendingJsonWorldBook) return;
    const targetCat = worldBookCategoryFilter !== '全部分组' ? worldBookCategoryFilter : '默认';
    const now = Date.now();
    const entry: WorldBookEntry = {
      id: `wb_${now}_${Math.random().toString(36).substring(2, 6)}`,
      title: pendingJsonWorldBook.title,
      author: pendingJsonWorldBook.author,
      content: JSON.stringify(pendingJsonWorldBook.raw, null, 2),
      category: targetCat,
      createdAt: now,
      updatedAt: now,
      entries: pendingJsonWorldBook.entries,
      jsonData: pendingJsonWorldBook.raw,
      importFormat: 'json',
    };
    updateAppData((prev) => ({ ...prev, worldBooks: [...(prev.worldBooks || []), entry] }));
    setPendingJsonWorldBook(null);
    setShowJsonWorldBookImportPreview(false);
    showToast(`成功导入世界书 “${entry.title}”，识别到 ${entry.entries?.length || 0} 个条目`, 'success');
  };

  const processWorldBookFileImport = async (file: File) => {
    const fileName = file.name;
    const baseName = fileName.replace(/\.[^/.]+$/, '') || fileName;
    let title = baseName;
    let author = '默认';
    let content = '';

    try {
      if (fileName.toLowerCase().endsWith('.json')) {
        const rawText = await file.text();
        // 兼容 UTF-8 BOM、代码块包裹以及首尾空白。
        const text = rawText.replace(/^\uFEFF/, '').trim().replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/i, '').trim();
        const parsed = JSON.parse(text);
        const doc = extractWorldBookJsonDocument(parsed, fileName);
        setPendingJsonWorldBook(doc);
        setShowJsonWorldBookImportPreview(true);
        return;
      } else if (fileName.toLowerCase().endsWith('.docx')) {
        const arrayBuffer = await file.arrayBuffer();
        const res = await mammoth.extractRawText({ arrayBuffer });
        content = res.value || '';
      } else {
        content = await file.text();
      }

      const targetCat = worldBookCategoryFilter !== '全部分组' ? worldBookCategoryFilter : '默认';
      const newEntry: WorldBookEntry = {
        id: `wb_${Date.now()}_${Math.random().toString(36).substring(2, 6)}`,
        title,
        author,
        content,
        category: targetCat,
        createdAt: Date.now(),
        updatedAt: Date.now(),
        importFormat: fileName.toLowerCase().endsWith('.docx') ? 'docx' : 'txt',
      };

      updateAppData((prev) => ({ ...prev, worldBooks: [...(prev.worldBooks || []), newEntry] }));
      showToast(`成功导入世界书 “${newEntry.title}”`, 'success');
    } catch (err: any) {
      showToast(`解析文件失败: ${err.message || '格式不符'}`, 'error');
    }
  };

  const handleFileUploadWorldBook = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = Array.from(e.target.files || []) as File[];
    if (files.length === 0) return;
    // JSON 需要先进入识别/确认面板；docx/txt 维持原来的直接导入方式。
    for (const file of files) {
      await processWorldBookFileImport(file);
      // JSON 会打开确认面板，一次只处理一个 JSON；其余 docx/txt 可继续顺序导入。
      if (file.name.toLowerCase().endsWith('.json')) break;
      await new Promise((resolve) => setTimeout(resolve, 0));
    }
    e.target.value = '';
  };

  const handleCreateNewWorldBookGroup = () => {
    const trimmed = newWorldBookGroupName.trim();
    if (!trimmed) return;
    if (trimmed === '全部分组') {
      showToast('"全部分组"为系统保留名称', 'error');
      return;
    }
    const currentCats = appData.worldBookCategories || ['默认'];
    if (currentCats.includes(trimmed)) {
      showToast('该分组名称已存在', 'error');
      return;
    }

    updateAppData({
      ...appData,
      worldBookCategories: [...currentCats, trimmed],
    });
    setNewWorldBookGroupName('');
    setShowNewWorldBookGroupModal(false);
    setWorldBookCategoryFilter(trimmed);
    showToast(`成功新建分组: ${trimmed}`, 'success');
  };

  const handleRenameWorldBookCategory = () => {
    if (!managingWorldBookCategory) return;
    const trimmed = renameWorldBookCategoryInput.trim();
    if (!trimmed) return;
    if (trimmed === '全部分组') {
      showToast('"全部分组"为系统保留名称', 'error');
      return;
    }

    const currentCats = appData.worldBookCategories || ['默认'];
    const updatedCats = currentCats.map((c) => (c === managingWorldBookCategory ? trimmed : c));
    const updatedWorldBooks = worldBooksList.map((p) =>
      (p.category || '默认') === managingWorldBookCategory ? { ...p, category: trimmed } : p
    );

    updateAppData({
      ...appData,
      worldBooks: updatedWorldBooks,
      worldBookCategories: updatedCats,
    });

    if (worldBookCategoryFilter === managingWorldBookCategory) {
      setWorldBookCategoryFilter(trimmed);
    }
    showToast(`分组已重命名为 "${trimmed}"`, 'success');
    setManagingWorldBookCategory(null);
  };

  const handleConfirmDeleteWorldBookCategory = () => {
    if (!managingWorldBookCategory) return;

    let updatedWorldBooks = [...worldBooksList];
    if (deleteItemsWithWorldBookCategory) {
      updatedWorldBooks = updatedWorldBooks.filter((p) => (p.category || '默认') !== managingWorldBookCategory);
    } else {
      updatedWorldBooks = updatedWorldBooks.map((p) =>
        (p.category || '默认') === managingWorldBookCategory ? { ...p, category: '默认' } : p
      );
    }

    const currentCats = appData.worldBookCategories || ['默认'];
    const updatedCats = currentCats.filter((c) => c !== managingWorldBookCategory);

    updateAppData({
      ...appData,
      worldBooks: updatedWorldBooks,
      worldBookCategories: updatedCats,
    });

    showToast(`分组 "${managingWorldBookCategory}" 已删除`, 'info');
    if (worldBookCategoryFilter === managingWorldBookCategory) {
      setWorldBookCategoryFilter('全部分组');
    }
    setManagingWorldBookCategory(null);
    setDeleteItemsWithWorldBookCategory(false);
  };

  const handleBatchDeleteWorldBooks = () => {
    if (selectedWorldBookIds.length === 0) return;
    if (!confirm(`确定要删除选中的 ${selectedWorldBookIds.length} 本世界书吗？`)) return;

    const updated = worldBooksList.filter((p) => !selectedWorldBookIds.includes(p.id));
    updateAppData({ ...appData, worldBooks: updated });

    showToast(`已成功删除 ${selectedWorldBookIds.length} 本世界书`, 'info');
    setSelectedWorldBookIds([]);
    setWorldBookBatchMode(false);
  };

  const handleConfirmBatchMoveWorldBooks = () => {
    if (!batchTargetWorldBookCategory) {
      showToast('请选择目标分组', 'error');
      return;
    }
    const target = batchTargetWorldBookCategory;
    const updated = worldBooksList.map((p) =>
      selectedWorldBookIds.includes(p.id) ? { ...p, category: target } : p
    );

    updateAppData({ ...appData, worldBooks: updated });
    setSelectedWorldBookIds([]);
    setWorldBookBatchMode(false);
    setShowWorldBookBatchMoveModal(false);
    setBatchTargetWorldBookCategory('');
    showToast(`已将 ${selectedWorldBookIds.length} 本世界书移动到 “${target}”`, 'success');
  };

  const exportWorldBookDocx = (wb: WorldBookEntry) => {
    const textLines = (wb.content || '').split('\n');
    const htmlContent = `<!DOCTYPE html><html><head><meta charset="utf-8"></head><body><h2>${wb.title}</h2><p>作者：${wb.author || '默认'}</p>${textLines.map((l) => `<p>${l}</p>`).join('')}</body></html>`;
    const blob = new Blob([htmlContent], { type: 'application/msword;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${wb.title || '世界书'}.docx`;
    a.click();
    URL.revokeObjectURL(url);
    showToast('已导出 .docx 文件', 'success');
  };

  const exportWorldBookTxt = (wb: WorldBookEntry) => {
    const content = `世界书：${wb.title}\n作者：${wb.author || '默认'}\n\n${wb.content || ''}`;
    const blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${wb.title || '世界书'}.txt`;
    a.click();
    URL.revokeObjectURL(url);
    showToast('已导出 .txt 文件', 'success');
  };

  const exportWorldBookJson = (wb: WorldBookEntry) => {
    let payload: any = wb;
    if (wb.importFormat === 'json' && wb.jsonData) {
      payload = JSON.parse(JSON.stringify(wb.jsonData));
      if (payload.character_book && typeof payload.character_book === 'object') {
        payload.character_book.name = wb.title;
        payload.character_book.author = wb.author || payload.character_book.author;
      } else {
        if ('name' in payload) payload.name = wb.title;
        else payload.title = wb.title;
        if (wb.author) payload.author = wb.author;
      }
    }
    const jsonStr = JSON.stringify(payload, null, 2);
    const blob = new Blob([jsonStr], { type: 'application/json;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${wb.title || '世界书'}.json`;
    a.click();
    URL.revokeObjectURL(url);
    showToast('已导出 .json 文件', 'success');
  };

  // 每个板块独立的数据字段。导出时只包含当前板块的数据，不会把其他板块一起带走。
  const SECTION_DATA_CONFIG: Record<string, { dataKeys: string[]; label: string }> = {
    'st-cards': { dataKeys: ['cards', 'groups'], label: 'ST 角色卡' },
    'st-themes': { dataKeys: ['themes', 'themeCategories'], label: 'ST 主题' },
    'st-presets': { dataKeys: ['presets', 'presetCategories'], label: 'ST 预设' },
    'st-plugins': { dataKeys: ['plugins', 'pluginCategories'], label: 'st插件和脚本' },
    'st-mobile': { dataKeys: ['phoneLinks'], label: '小手机链接' },
    'normal-cards': { dataKeys: ['normalCards', 'normalCardCategories'], label: '普通角色卡' },
    'stickers': { dataKeys: ['stickerPacks', 'stickerCategories'], label: '表情包' },
    'worldbook': { dataKeys: ['worldBooks', 'worldBookCategories'], label: '世界书' },
    'extras-app': { dataKeys: ['extraStories', 'extraStoryCategories'], label: '番外小剧场' },
    'chat-memes': { dataKeys: ['chatMemes', 'chatMemeCategories'], label: '聊天梗' },
    'themes': { dataKeys: ['beautifications', 'beautificationCategories'], label: '美化' },
    'fonts': { dataKeys: ['fonts', 'fontCategories'], label: '字体' },
    'api-storage': { dataKeys: ['apis', 'apiCategories'], label: 'API 存储' },
  };

  const getCurrentSectionConfig = () => SECTION_DATA_CONFIG[currentPage] || null;

  const handleExportCurrentSection = () => {
    const config = getCurrentSectionConfig();
    if (!config) return;
    try {
      const data: Record<string, any> = {};
      for (const key of config.dataKeys) data[key] = (appData as any)[key] ?? (key.endsWith('Categories') || key === 'groups' ? [] : []);
      const payload = {
        format: 'toolbox-section-backup',
        version: 1,
        section: currentPage,
        sectionName: config.label,
        exportedAt: new Date().toISOString(),
        data,
      };
      downloadJsonFile(`toolbox-${currentPage}-${formatDateForFileName()}.json`, payload);
      showToast(`已导出「${config.label}」全部数据`, 'success');
    } catch (err: any) {
      console.error('Export current section failed', err);
      showToast(`导出失败：${err?.message || String(err)}`, 'error');
    }
  };

  const handleImportCurrentSection = async (file: File) => {
    const config = getCurrentSectionConfig();
    if (!config) return;
    setIsImportingSection(true);
    try {
      const text = await file.text();
      const parsed = JSON.parse(text);
      // 只接受当前板块导出的备份，避免误把完整备份或其他板块数据覆盖进来。
      if (parsed?.format !== 'toolbox-section-backup' || parsed?.section !== currentPage || !parsed?.data || typeof parsed.data !== 'object') {
        throw new Error(`这不是「${config.label}」的板块备份文件`);
      }
      const incoming = parsed.data as Record<string, any>;
      const nextData: any = { ...appData };
      for (const key of config.dataKeys) {
        const value = incoming[key];
        if (value !== undefined && !Array.isArray(value)) throw new Error(`数据字段「${key}」格式不正确`);
        nextData[key] = Array.isArray(value) ? value : [];
      }
      updateAppData(nextData);
      showToast(`已导入「${config.label}」全部数据`, 'success');
    } catch (err: any) {
      console.error('Import current section failed', err);
      showToast(`导入失败：${err?.message || String(err)}`, 'error');
    } finally {
      setIsImportingSection(false);
    }
  };

  const triggerImportCurrentSection = () => {
    if (isImportingSection) return;
    sectionImportFileInputRef.current?.click();
  };

  // Page Names Dictionary for Header Title
  const PAGE_NAMES: Record<string, string> = {
    'st-cards': 'ST 角色卡',
    'st-themes': 'ST 主题',
    'st-presets': 'ST 预设',
    'st-plugins': 'st插件和脚本',
    'st-mobile': '小手机链接',
    'normal-cards': '普通角色卡',
    'stickers': '表情包',
    'worldbook': '世界书',
    'extras-app': '番外小剧场',
    'chat-memes': '聊天梗',
    'themes': '美化',
    'fonts': '字体',
    'api-storage': 'API 存储',
  };

  // Sync Theme
  useEffect(() => {
    document.documentElement.setAttribute('data-theme', theme);
    if (theme === 'dark') {
      document.documentElement.classList.add('dark');
    } else {
      document.documentElement.classList.remove('dark');
    }
    localStorage.setItem('tavern_vault_theme', theme);
  }, [theme]);

  // Persist App Data
  // 只负责更新 React 状态；真正的持久化放到 effect 中，避免在 setState updater
  // 内产生 IndexedDB / Toast 等副作用，也避免 React 快速更新时保存旧快照。
  const updateAppData = (newAppData: AppData | ((prev: AppData) => AppData)) => {
    appDataDirtyRef.current = true;
    setAppData(newAppData);
  };

  useEffect(() => {
    if (!appDataHydrated) return;
    void saveAppData(appData).then((ok) => {
      if (!ok) showToast('保存失败：本地数据库写入失败，请重试', 'error');
    });
  }, [appData, appDataHydrated]);

  const showToast = (msg: string, type: 'success' | 'error' | 'info' = 'info') => {
    const id = Math.random().toString(36).substring(2, 9);
    setToasts((prev) => [...prev, { id, msg, type }]);
    setTimeout(() => {
      setToasts((prev) => prev.filter((t) => t.id !== id));
    }, 3200);
  };

  // 统一可靠复制：优先 Clipboard API，移动端/非安全上下文失败时回退到 textarea + execCommand。
  const copyTextReliable = async (text: string, successMessage = '已复制到剪贴板') => {
    try {
      if (navigator.clipboard && window.isSecureContext) {
        await navigator.clipboard.writeText(text);
        showToast(successMessage, 'success');
        return true;
      }
    } catch (err) {
      console.warn('Clipboard API failed, using fallback.', err);
    }
    try {
      const textarea = document.createElement('textarea');
      textarea.value = text;
      textarea.setAttribute('readonly', '');
      textarea.style.position = 'fixed';
      textarea.style.opacity = '0';
      textarea.style.pointerEvents = 'none';
      document.body.appendChild(textarea);
      textarea.focus();
      textarea.select();
      textarea.setSelectionRange(0, textarea.value.length);
      const copied = document.execCommand('copy');
      document.body.removeChild(textarea);
      if (copied) {
        showToast(successMessage, 'success');
        return true;
      }
    } catch (err) {
      console.warn('Clipboard fallback failed.', err);
    }
    showToast('复制失败，请长按文本手动复制', 'error');
    return false;
  };

  // Upload Handlers
  const handleFileUpload = async (files: FileList | File[]) => {
    const fileArr = Array.from(files);
    if (fileArr.length === 0) return;

    let successCount = 0;
    let failCount = 0;
    let pendingBatch: CardEntry[] = [];
    const total = fileArr.length;
    const showProgress = total > 8;
    const BATCH_SIZE = 6;
    const groupForNewCards = currentGroup !== '全部' ? currentGroup : '默认';

    if (showProgress) showToast(`开始导入 ${total} 个文件…`, 'info');

    for (let i = 0; i < fileArr.length; i++) {
      const file = fileArr[i];
      try {
        const card = await parseCardFile(file);
        // Default group assigned to current group if not '全部', else '默认'
        card.group = groupForNewCards;
        pendingBatch.push(card);
        successCount++;
      } catch (err: any) {
        failCount++;
        showToast(`文件 "${file.name}" 解析失败: ${err.message}`, 'error');
      }

      const isLast = i === fileArr.length - 1;
      // 大批量导入时按小批次落库并主动让出主线程，避免一次性解析/写入上百个文件
      // 造成长时间无响应甚至在低内存设备上崩溃；同时给出进度提示。
      if (pendingBatch.length >= BATCH_SIZE || isLast) {
        if (pendingBatch.length > 0) {
          const batch = pendingBatch;
          pendingBatch = [];
          updateAppData((prev) => ({ ...prev, cards: [...prev.cards, ...batch] }));
        }
        if (showProgress && !isLast) showToast(`正在导入 ${i + 1}/${total} 个文件…`, 'info');
        await new Promise((resolve) => setTimeout(resolve, 0));
      }
    }

    if (successCount > 0) {
      showToast(`成功导入 ${successCount} 张角色卡！${failCount > 0 ? `（${failCount} 个失败）` : ''}`, 'success');
    } else if (failCount > 0) {
      showToast(`导入失败：${failCount} 个文件均无法解析`, 'error');
    }
  };

  // Delete Single Card
  const handleDeleteCard = (id: string) => {
    if (!confirm('确定要删除这张角色卡吗？')) return;
    const updatedCards = appData.cards.filter((c) => c.id !== id);
    updateAppData({ ...appData, cards: updatedCards });
    if (detailCardId === id) setDetailCardId(null);
    showToast('角色卡已删除', 'info');
  };

  // Export Card Single
  const exportAsJson = (card: CardEntry) => {
    const rawCardData = getCurrentCardData(card);
    const pureData = getPureCardDataForExport(rawCardData);
    const jsonStr = JSON.stringify(pureData, null, 2);
    const blob = new Blob([jsonStr], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${getCardDisplayName(card)}.json`;
    a.click();
    URL.revokeObjectURL(url);
    showToast('JSON 导出成功（已移除附加图片与分类数据）', 'success');
  };

  const exportAsPng = async (card: CardEntry) => {
    try {
      const rawCardData = getCurrentCardData(card);
      const pureData = getPureCardDataForExport(rawCardData);
      const jsonStr = JSON.stringify(pureData);
      const b64 = btoa(unescape(encodeURIComponent(jsonStr)));

      let baseImageUrl = card.coverImage || 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=';

      const img = new Image();
      img.crossOrigin = 'anonymous';
      await new Promise((resolve, reject) => {
        img.onload = resolve;
        img.onerror = reject;
        img.src = baseImageUrl;
      });

      const canvas = document.createElement('canvas');
      canvas.width = img.width;
      canvas.height = img.height;
      const ctx = canvas.getContext('2d');
      if (!ctx) throw new Error('Canvas context error');
      ctx.drawImage(img, 0, 0);

      const blob = await new Promise<Blob | null>((res) => canvas.toBlob(res, 'image/png'));
      if (!blob) throw new Error('Blob creation failed');
      const buf = await blob.arrayBuffer();

      const keyword = card.version === 'v3' ? 'ccv3' : 'chara';
      const newPngBuf = injectPngTextChunk(buf, keyword, b64);
      const pngBlob = new Blob([newPngBuf], { type: 'image/png' });

      const url = URL.createObjectURL(pngBlob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `${getCardDisplayName(card)}.png`;
      a.click();
      URL.revokeObjectURL(url);
      showToast('PNG 导出成功（不包含附加板块内容）', 'success');
    } catch (e: any) {
      console.error(e);
      showToast('PNG 导出失败: ' + e.message, 'error');
    }
  };

  // Update Card Overwrite (Requirement 6)
  const handleUpdateCardOverwrite = async (file: File) => {
    if (!detailCardId) return;
    const cardIndex = appData.cards.findIndex((c) => c.id === detailCardId);
    if (cardIndex === -1) return;

    try {
      const updatedCardParsed = await parseCardFile(file);
      const existingCard = appData.cards[cardIndex];

      // Overwrite rawData, version, fileType, fileName, name, author, coverImage while keeping id, group, screenshots, createdAt
      const overwrittenCard: CardEntry = {
        ...existingCard,
        name: updatedCardParsed.name,
        fileName: updatedCardParsed.fileName,
        fileType: updatedCardParsed.fileType,
        version: updatedCardParsed.version,
        author: updatedCardParsed.authorManual ? existingCard.author : updatedCardParsed.author,
        rawData: updatedCardParsed.rawData,
        coverImage: updatedCardParsed.coverImage || existingCard.coverImage,
        editHistory: {}, // Reset edit history as rawData is newly updated
        edited: false,
        updatedAt: Date.now(),
      };

      const updatedCards = [...appData.cards];
      updatedCards[cardIndex] = overwrittenCard;
      updateAppData({ ...appData, cards: updatedCards });

      showToast('角色卡覆盖更新成功！', 'success');
    } catch (err: any) {
      showToast(`覆盖更新失败: ${err.message}`, 'error');
    }
  };

  // Group Management
  const handleCreateNewGroup = () => {
    const trimmed = newGroupName.trim();
    if (!trimmed) return;
    if (trimmed === '全部') {
      showToast('"全部"为系统保留名称', 'error');
      return;
    }
    if (appData.groups.includes(trimmed)) {
      showToast('该分组已存在', 'error');
      return;
    }

    updateAppData({
      ...appData,
      groups: [...appData.groups, trimmed],
    });
    setNewGroupName('');
    setShowNewGroupModal(false);
    setCurrentGroup(trimmed);
    showToast(`成功新建分组: ${trimmed}`, 'success');
  };

  const handleConfirmDeleteGroup = () => {
    if (!managingGroup) return;

    let updatedCards = [...appData.cards];
    if (deleteCardsWithGroup) {
      // Delete cards in this group
      updatedCards = updatedCards.filter((c) => c.group !== managingGroup);
    } else {
      // Reassign cards to '默认'
      updatedCards = updatedCards.map((c) => (c.group === managingGroup ? { ...c, group: '默认' } : c));
    }

    const updatedGroups = appData.groups.filter((g) => g !== managingGroup);
    updateAppData({
      cards: updatedCards,
      groups: updatedGroups,
    });

    showToast(`分组 "${managingGroup}" 已删除`, 'info');
    if (currentGroup === managingGroup) setCurrentGroup('全部');
    setManagingGroup(null);
    setShowDeleteGroupConfirm(false);
    setDeleteCardsWithGroup(false);
  };

  const handleRenameGroup = () => {
    if (!managingGroup) return;
    const trimmed = renameGroupInput.trim();
    if (!trimmed) return;
    if (trimmed === '全部') {
      showToast('"全部"为系统保留名称', 'error');
      return;
    }

    const updatedGroups = appData.groups.map((g) => (g === managingGroup ? trimmed : g));
    const updatedCards = appData.cards.map((c) => (c.group === managingGroup ? { ...c, group: trimmed } : c));

    updateAppData({
      cards: updatedCards,
      groups: updatedGroups,
    });

    if (currentGroup === managingGroup) setCurrentGroup(trimmed);
    showToast(`分组已重命名为 "${trimmed}"`, 'success');
    setManagingGroup(null);
    setRenameGroupInput('');
  };

  // Batch Operations
  const handleBatchDeleteCards = () => {
    if (selectedCardIds.length === 0) return;
    if (!confirm(`确定要删除选中的 ${selectedCardIds.length} 张角色卡吗？`)) return;

    const updatedCards = appData.cards.filter((c) => !selectedCardIds.includes(c.id));
    updateAppData({ ...appData, cards: updatedCards });
    setSelectedCardIds([]);
    showToast(`已删除 ${selectedCardIds.length} 张角色卡`, 'info');
  };

  const handleBatchMoveCards = () => {
    if (selectedCardIds.length === 0 || !batchTargetGroup) return;
    const updatedCards = appData.cards.map((c) => (selectedCardIds.includes(c.id) ? { ...c, group: batchTargetGroup } : c));
    updateAppData({ ...appData, cards: updatedCards });
    setSelectedCardIds([]);
    setShowBatchMoveModal(false);
    showToast(`已移动 ${selectedCardIds.length} 张卡片到分组 "${batchTargetGroup}"`, 'success');
  };

  // Filtered Cards List
  const baseFilteredCards = appData.cards.filter((card) => {
    // Group filter
    if (currentGroup !== '全部') {
      const cardGrp = card.group || '默认';
      if (cardGrp !== currentGroup) return false;
    }

    // Search query
    if (searchQuery.trim()) {
      const q = searchQuery.toLowerCase();
      const name = getCardDisplayName(card).toLowerCase();
      const author = getCardCreator(card).toLowerCase();
      const tags = getCardTags(card).join(' ').toLowerCase();
      const desc = getCardDescription(card).toLowerCase();
      return name.includes(q) || author.includes(q) || tags.includes(q) || desc.includes(q);
    }

    return true;
  });

  const filteredCards = sortItemList(
    cardSortOrder === 'default' ? [...baseFilteredCards].reverse() : baseFilteredCards,
    cardSortOrder,
    getCardDisplayName,
    (c) => c.createdAt || 0
  );

  // 搜索词/分组/排序方式变化时，把可见数量重置回第一页，避免筛选后仍保留上一次滚动到的很大数值。
  useEffect(() => {
    setCardVisibleCount(CARD_PAGE_SIZE);
  }, [searchQuery, currentGroup, cardSortOrder]);

  // 滚动到底部哨兵元素进入视口时自动加载下一页角色卡。
  useEffect(() => {
    const el = cardLoadMoreRef.current;
    if (!el) return;
    const observer = new IntersectionObserver(
      (entries) => {
        if (entries[0]?.isIntersecting) {
          setCardVisibleCount((prev) => Math.min(prev + CARD_PAGE_SIZE, filteredCards.length));
        }
      },
      { rootMargin: '600px' }
    );
    observer.observe(el);
    return () => observer.disconnect();
  }, [filteredCards.length]);

  const visibleFilteredCards = filteredCards.slice(0, cardVisibleCount);

  // Current Active Detail Card
  const activeDetailCard = appData.cards.find((c) => c.id === detailCardId);
  const activeLiveCardData = activeDetailCard ? getCurrentCardData(activeDetailCard) : null;

  const normalizeAssociationName = (value: string) => {
    let name = (value || '').trim().toLowerCase();
    // 允许“陆宴辞（dk）/dk版”“陆宴辞（社畜版）”与“陆宴辞”识别为同一主体；作者仍必须完全相同。
    name = name.split('/')[0].trim();
    name = name.replace(/[（(【\[].*?[）)】\]]/g, '').trim();
    name = name.replace(/\s+/g, '');
    name = name.replace(/(?:版|ver\.?|v\d+)$/i, '');
    return name;
  };

  const getAssociationCandidates = (card: CardEntry | undefined) => {
    if (!card) return [];
    const name = normalizeAssociationName(getCardDisplayName(card));
    const author = getCardCreator(card).trim().toLowerCase();
    if (!name || !author || author === '未知作者') return [];
    const existingIds = new Set((card.associations || []).map((a) => a.cardId));
    return appData.cards.filter((candidate) => {
      if (candidate.id === card.id || existingIds.has(candidate.id)) return false;
      return normalizeAssociationName(getCardDisplayName(candidate)) === name && getCardCreator(candidate).trim().toLowerCase() === author;
    });
  };

  const buildAssociationComponent = (seedId: string, cards: CardEntry[]) => {
    const seen = new Set<string>();
    const queue = [seedId];
    while (queue.length) {
      const id = queue.shift()!;
      if (seen.has(id)) continue;
      seen.add(id);
      const card = cards.find((c) => c.id === id);
      (card?.associations || []).forEach((a) => {
        if (!seen.has(a.cardId)) queue.push(a.cardId);
      });
    }
    return cards.filter((c) => seen.has(c.id));
  };

  const handleAddCardAssociation = () => {
    if (!activeDetailCard || !associationTargetId) {
      showToast('请先选择要关联的角色卡', 'error');
      return;
    }
    const target = appData.cards.find((c) => c.id === associationTargetId);
    if (!target) {
      showToast('关联目标不存在', 'error');
      return;
    }
    const candidates = getAssociationCandidates(activeDetailCard);
    if (!candidates.some((c) => c.id === target.id)) {
      showToast('只能关联主体相同且作者相同的角色卡', 'error');
      return;
    }

    const now = Date.now();
    const note = associationNote.trim() || undefined;
    let updatedCards = appData.cards.map((c) => ({ ...c, associations: [...(c.associations || [])] }));
    const owner = updatedCards.find((c) => c.id === activeDetailCard.id)!;
    const targetCard = updatedCards.find((c) => c.id === target.id)!;

    const upsert = (card: CardEntry, assoc: CardAssociation) => {
      const existing = (card.associations || []).findIndex((a) => a.cardId === assoc.cardId);
      const next = [...(card.associations || [])];
      if (existing >= 0) next[existing] = { ...next[existing], ...assoc };
      else next.push(assoc);
      card.associations = next;
      card.edited = true;
      card.updatedAt = now;
    };

    upsert(owner, { cardId: target.id, note, noteOwnerId: activeDetailCard.id, isPrimary: associationPrimary, createdAt: now });
    upsert(targetCard, { cardId: activeDetailCard.id, isPrimary: !associationPrimary, createdAt: now });

    // 关联彻底化：把同一主体的整个关联网络补齐。备注始终只属于写备注的那一方，不复制给另一方。
    const component = buildAssociationComponent(activeDetailCard.id, updatedCards);
    const componentIds = new Set(component.map((c) => c.id));
    for (const a of component) {
      for (const b of component) {
        if (a.id === b.id) continue;
        const existing = (a.associations || []).find((x) => x.cardId === b.id);
        if (!existing) {
          upsert(a, { cardId: b.id, isPrimary: false, createdAt: now });
        }
      }
    }

    // 由于新增的边可能扩展组件，再跑一次以确保所有成员两两可见。
    const finalComponent = buildAssociationComponent(activeDetailCard.id, updatedCards);
    const finalIds = new Set(finalComponent.map((c) => c.id));
    updatedCards = updatedCards.map((c) => {
      if (!finalIds.has(c.id)) return c;
      const next = [...(c.associations || [])];
      finalIds.forEach((id) => {
        if (id === c.id || next.some((a) => a.cardId === id)) return;
        next.push({ cardId: id, isPrimary: false, createdAt: now });
      });
      return { ...c, associations: next };
    });

    updateAppData({ ...appData, cards: updatedCards });
    setAssociationTargetId('');
    setAssociationNote('');
    setAssociationPrimary(false);
    showToast(`已关联「${getCardDisplayName(target)}」，关联网络已同步`, 'success');
  };

  const handleRemoveCardAssociation = (relatedId: string) => {
    if (!activeDetailCard) return;
    const updatedCards = appData.cards.map((c) => {
      if (c.id === activeDetailCard.id || c.id === relatedId) {
        return { ...c, associations: (c.associations || []).filter((a) => a.cardId !== relatedId && a.cardId !== activeDetailCard.id), edited: true, updatedAt: Date.now() };
      }
      return c;
    });
    updateAppData({ ...appData, cards: updatedCards });
    showToast('关联已解除', 'info');
  };

  const openAssociatedCard = (cardId: string) => {
    setGachaCard(null);
    setDetailCardId(cardId);
    setDetailTab('overview');
  };

  const handleConfirmAddCardWorldBookEntry = () => {
    if (!activeDetailCard) return;

    const wb = JSON.parse(JSON.stringify(getCardWorldBook(activeDetailCard) || { entries: [] }));
    const entry = {
      comment: newCardWorldBookEntryForm.comment.trim() || `新世界书条目 ${((wb.entries || []).length + 1)}`,
      keys: newCardWorldBookEntryForm.keys
        .split(',')
        .map((k) => k.trim())
        .filter(Boolean),
      key: newCardWorldBookEntryForm.keys
        .split(',')
        .map((k) => k.trim())
        .filter(Boolean),
      content: newCardWorldBookEntryForm.content,
      disable: false,
    };
    wb.entries = [...(Array.isArray(wb.entries) ? wb.entries : []), entry];

    const updatedCards = appData.cards.map((c) =>
      c.id === activeDetailCard.id
        ? { ...c, editHistory: { ...c.editHistory, character_book: wb }, edited: true }
        : c
    );
    updateAppData({ ...appData, cards: updatedCards });
    setNewCardWorldBookEntryForm({ comment: '', keys: '', content: '' });
    setShowAddCardWorldBookEntryModal(false);
    showToast('已新增世界书条目', 'success');
  };

  return (
    <div className="min-h-screen bg-zinc-50 dark:bg-zinc-950 text-zinc-900 dark:text-zinc-100 flex flex-col font-sans antialiased selection:bg-zinc-800 selection:text-white">
      {/* App Shell Container */}
      <div className="flex flex-1 h-screen overflow-hidden relative">

        {/* Backdrop overlay for closing sidebar when open (Requirement 5) */}
        {sidebarOpen && (
          <div
            className="fixed inset-0 bg-black/40 dark:bg-black/60 z-40 transition-opacity"
            onClick={() => setSidebarOpen(false)}
          />
        )}

        {/* Sidebar Navigation */}
        <aside
          className={`fixed top-0 bottom-0 left-0 z-50 w-64 bg-white dark:bg-zinc-900 border-r border-zinc-200 dark:border-zinc-800 flex flex-col transition-transform duration-300 ease-in-out ${
            sidebarOpen ? 'translate-x-0 shadow-2xl' : '-translate-x-full'
          }`}
        >
          <div className="p-4 border-b border-zinc-200 dark:border-zinc-800 flex items-center justify-between flex-shrink-0">
            <h2 className="text-xs font-bold uppercase tracking-widest text-zinc-400">工具箱导航</h2>
            <button
              onClick={() => setSidebarOpen(false)}
              className="p-1.5 rounded-lg hover:bg-zinc-100 dark:hover:bg-zinc-800 text-zinc-500 transition-colors"
            >
              <X className="w-4 h-4" />
            </button>
          </div>

          <nav className="flex-1 overflow-y-auto py-3 px-2 space-y-6">
            {/* SillyTavern Section */}
            <div>
              <div className="px-3 mb-2 text-[11px] font-semibold text-zinc-400 dark:text-zinc-500 uppercase tracking-wider">
                SillyTavern
              </div>
              <div className="space-y-1">
                {[
                  { id: 'st-cards', name: 'ST 角色卡' },
                  { id: 'st-themes', name: 'ST 主题' },
                  { id: 'st-presets', name: 'ST 预设' },
                  { id: 'st-plugins', name: 'st插件和脚本' },
                ].map((item) => (
                  <button
                    key={item.id}
                    onClick={() => {
                      setCurrentPage(item.id);
                      setSidebarOpen(false);
                    }}
                    className={`w-full text-left px-3 py-2 rounded-lg text-sm font-medium transition-colors ${
                      currentPage === item.id
                        ? 'bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 shadow-sm'
                        : 'text-zinc-600 dark:text-zinc-400 hover:bg-zinc-100 dark:hover:bg-zinc-800/80 hover:text-zinc-900 dark:hover:text-zinc-100'
                    }`}
                  >
                    {item.name}
                  </button>
                ))}
              </div>
            </div>

            {/* 通用 Section (Requirement 9: Put all items in 通用 after 小手机链接, remove 其他 section) */}
            <div>
              <div className="px-3 mb-2 text-[11px] font-semibold text-zinc-400 dark:text-zinc-500 uppercase tracking-wider">
                通用
              </div>
              <div className="space-y-1">
                {[
                  { id: 'st-mobile', name: '小手机链接' },
                  { id: 'normal-cards', name: '普通角色卡' },
                  { id: 'stickers', name: '表情包' },
                  { id: 'worldbook', name: '世界书' },
                  { id: 'extras-app', name: '番外小剧场' },
                  { id: 'chat-memes', name: '聊天梗' },
                  { id: 'themes', name: '美化' },
                  { id: 'fonts', name: '字体' },
                  { id: 'api-storage', name: 'API 存储' },
                ].map((item) => (
                  <button
                    key={item.id}
                    onClick={() => {
                      setCurrentPage(item.id);
                      setSidebarOpen(false);
                    }}
                    className={`w-full text-left px-3 py-2 rounded-lg text-sm font-medium transition-colors ${
                      currentPage === item.id
                        ? 'bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 shadow-sm'
                        : 'text-zinc-600 dark:text-zinc-400 hover:bg-zinc-100 dark:hover:bg-zinc-800/80 hover:text-zinc-900 dark:hover:text-zinc-100'
                    }`}
                  >
                    {item.name}
                  </button>
                ))}
              </div>
            </div>
          </nav>

          <div className="p-3 border-t border-zinc-200 dark:border-zinc-800 space-y-2 flex-shrink-0">
            <button
              onClick={() => importBackupInputRef.current?.click()}
              disabled={isImportingBackup}
              className="w-full flex items-center justify-center gap-2 px-3 py-2 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-800 hover:bg-zinc-100 dark:hover:bg-zinc-800 text-zinc-700 dark:text-zinc-300 transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
            >
              <Upload className="w-3.5 h-3.5" /> {isImportingBackup ? '正在导入…' : '导入全部数据（不含字体）'}
            </button>
            <button
              onClick={handleExportFullBackup}
              disabled={isExportingBackup}
              className="w-full flex items-center justify-center gap-2 px-3 py-2 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-800 hover:bg-zinc-100 dark:hover:bg-zinc-800 text-zinc-700 dark:text-zinc-300 transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
            >
              <Download className="w-3.5 h-3.5" /> {isExportingBackup ? '正在导出…' : '导出全部数据（不含字体）'}
            </button>
          </div>
        </aside>

        {/* Main Content View */}
        <div className="flex-1 flex flex-col min-w-0 overflow-hidden">
          {/* Header Bar */}
          <header className="h-14 bg-white dark:bg-zinc-900 border-b border-zinc-200 dark:border-zinc-800 px-4 flex items-center justify-between flex-shrink-0 z-10">
            <div className="flex items-center gap-3">
              <button
                onClick={() => setSidebarOpen((prev) => !prev)}
                className="p-2 rounded-lg hover:bg-zinc-100 dark:hover:bg-zinc-800 text-zinc-600 dark:text-zinc-300 transition-colors"
                title="打开侧边栏"
              >
                <Menu className="w-5 h-5" />
              </button>
              <span className="font-bold text-base tracking-tight text-zinc-900 dark:text-zinc-100 cursor-pointer" onClick={() => setSidebarOpen((p) => !p)}>
                {PAGE_NAMES[currentPage] || 'ST 角色卡'}
              </span>
              <div className="flex items-center gap-1.5 ml-0.5">
                <button
                  type="button"
                  onClick={triggerImportCurrentSection}
                  disabled={isImportingSection}
                  className="h-7 px-2.5 rounded-md border border-zinc-200 dark:border-zinc-700 bg-white dark:bg-zinc-900 text-[12px] font-medium text-zinc-700 dark:text-zinc-200 hover:bg-zinc-100 dark:hover:bg-zinc-800 transition-colors disabled:opacity-50"
                  title="导入当前板块全部数据"
                >
                  {isImportingSection ? '导入中…' : '导入'}
                </button>
                <button
                  type="button"
                  onClick={handleExportCurrentSection}
                  className="h-7 px-2.5 rounded-md border border-zinc-200 dark:border-zinc-700 bg-white dark:bg-zinc-900 text-[12px] font-medium text-zinc-700 dark:text-zinc-200 hover:bg-zinc-100 dark:hover:bg-zinc-800 transition-colors"
                  title="导出当前板块全部数据"
                >
                  导出
                </button>
              </div>
            </div>

            <div className="flex items-center gap-2">
              {currentPage === 'st-themes' && (
                <>
                  <button
                    onClick={() => themeDocumentFileInputRef.current?.click()}
                    className="p-2 rounded-lg border border-zinc-200 dark:border-zinc-800 hover:bg-zinc-100 dark:hover:bg-zinc-800 text-zinc-600 dark:text-zinc-300 transition-colors flex items-center justify-center"
                    title="上传 ST 主题文档 (.docx, .txt, .json, .css)"
                  >
                    <Plus className="w-4 h-4" />
                  </button>
                  <input
                    type="file"
                    ref={themeDocumentFileInputRef}
                    multiple
                    accept=".docx,.txt,.json,.css"
                    className="hidden"
                    onChange={(e) => {
                      if (e.target.files) handleThemeDocumentFileUpload(e.target.files);
                      e.target.value = '';
                    }}
                  />
                </>
              )}
              {currentPage === 'themes' && (
                <>
                  <button
                    onClick={() => beautificationDocumentFileInputRef.current?.click()}
                    className="p-2 rounded-lg border border-zinc-200 dark:border-zinc-800 hover:bg-zinc-100 dark:hover:bg-zinc-800 text-zinc-600 dark:text-zinc-300 transition-colors flex items-center justify-center"
                    title="上传美化文档/图片 (.docx, .txt, .json, .css, .png)"
                  >
                    <Plus className="w-4 h-4" />
                  </button>
                  <input
                    type="file"
                    ref={beautificationDocumentFileInputRef}
                    multiple
                    accept=".docx,.txt,.json,.css,.png,.jpg,.jpeg,.webp,image/*"
                    className="hidden"
                    onChange={(e) => {
                      if (e.target.files) handleBeautificationDocumentFileUpload(e.target.files);
                      e.target.value = '';
                    }}
                  />
                </>
              )}
              {currentPage === 'st-presets' && (
                <button
                  onClick={() => presetFileInputRef.current?.click()}
                  className="p-2 rounded-lg border border-zinc-200 dark:border-zinc-800 hover:bg-zinc-100 dark:hover:bg-zinc-800 text-zinc-600 dark:text-zinc-300 transition-colors flex items-center justify-center"
                  title="上传预设 (JSON)"
                >
                  <Plus className="w-4 h-4" />
                </button>
              )}
              {currentPage === 'st-plugins' && (
                <button
                  onClick={() => setShowAddPluginTypeModal(true)}
                  className="p-2 rounded-lg border border-zinc-200 dark:border-zinc-800 hover:bg-zinc-100 dark:hover:bg-zinc-800 text-zinc-600 dark:text-zinc-300 transition-colors flex items-center justify-center"
                  title="添加插件或导入脚本"
                >
                  <Plus className="w-4 h-4" />
                </button>
              )}
              {currentPage === 'st-mobile' && (
                <button
                  onClick={() => {
                    setPhoneForm({ name: '', url: '', contact: '', description: '' });
                    setShowAddPhoneModal(true);
                  }}
                  className="p-2 rounded-lg border border-zinc-200 dark:border-zinc-800 hover:bg-zinc-100 dark:hover:bg-zinc-800 text-zinc-600 dark:text-zinc-300 transition-colors flex items-center justify-center"
                  title="添加小手机链接"
                >
                  <Plus className="w-4 h-4" />
                </button>
              )}
              {currentPage === 'normal-cards' && (
                <button
                  onClick={() => normalCardFileInputRef.current?.click()}
                  className="p-2 rounded-lg border border-zinc-200 dark:border-zinc-800 hover:bg-zinc-100 dark:hover:bg-zinc-800 text-zinc-600 dark:text-zinc-300 transition-colors flex items-center justify-center"
                  title="上传普通角色卡文档 (.docx, .txt, .zip)"
                >
                  <Plus className="w-4 h-4" />
                </button>
              )}
              {currentPage === 'api-storage' && (
                <button
                  onClick={handleOpenAddApiModal}
                  className="p-2 rounded-lg border border-zinc-200 dark:border-zinc-800 hover:bg-zinc-100 dark:hover:bg-zinc-800 text-zinc-600 dark:text-zinc-300 transition-colors flex items-center justify-center"
                  title="添加 API"
                >
                  <Plus className="w-4 h-4" />
                </button>
              )}
              {currentPage === 'fonts' && (
                <button
                  onClick={() => setShowAddFontChoiceModal(true)}
                  className="p-2 rounded-lg border border-zinc-200 dark:border-zinc-800 hover:bg-zinc-100 dark:hover:bg-zinc-800 text-zinc-600 dark:text-zinc-300 transition-colors flex items-center justify-center"
                  title="添加字体"
                >
                  <Plus className="w-4 h-4" />
                </button>
              )}
              {currentPage === 'extras-app' && (
                <button
                  onClick={() => setShowAddExtraStoryChoiceModal(true)}
                  className="p-2 rounded-lg border border-zinc-200 dark:border-zinc-800 hover:bg-zinc-100 dark:hover:bg-zinc-800 text-zinc-600 dark:text-zinc-300 transition-colors flex items-center justify-center"
                  title="添加番外小剧场"
                >
                  <Plus className="w-4 h-4" />
                </button>
              )}
              {currentPage === 'stickers' && (
                <button
                  onClick={() => openBatchUpload('stickers')}
                  className="p-2 rounded-lg border border-zinc-200 dark:border-zinc-800 hover:bg-zinc-100 dark:hover:bg-zinc-800 text-zinc-600 dark:text-zinc-300 transition-colors flex items-center justify-center"
                  title="上传表情包图集"
                >
                  <Plus className="w-4 h-4" />
                </button>
              )}
              {currentPage === 'worldbook' && (
                <button
                  onClick={() => openBatchUpload('worldbook')}
                  className="p-2 rounded-lg border border-zinc-200 dark:border-zinc-800 hover:bg-zinc-100 dark:hover:bg-zinc-800 text-zinc-600 dark:text-zinc-300 transition-colors flex items-center justify-center"
                  title="上传世界书文档"
                >
                  <Plus className="w-4 h-4" />
                </button>
              )}
              {/* 抽卡仅在 ST 角色卡页面显示 */}
              {currentPage === 'st-cards' && (
                <button
                  onClick={handleDrawRandomCard}
                  disabled={isGachaSpinning}
                  className="px-2.5 py-1.5 text-xs font-semibold rounded-lg border border-zinc-200 dark:border-zinc-800 bg-white dark:bg-zinc-900 text-zinc-800 dark:text-zinc-200 hover:bg-zinc-100 dark:hover:bg-zinc-800 transition-colors whitespace-nowrap"
                  title="随机抽取一张角色卡"
                >
                  {isGachaSpinning ? '抽取中...' : '抽卡'}
                </button>
              )}

              {currentPage === 'chat-memes' && (
                <button
                  onClick={() => setShowChatAddChoiceModal(true)}
                  className="p-2 rounded-lg border border-zinc-200 dark:border-zinc-800 hover:bg-zinc-100 dark:hover:bg-zinc-800 text-zinc-600 dark:text-zinc-300 transition-colors"
                  title="新增聊天梗"
                >
                  <Plus className="w-4 h-4" />
                </button>
              )}

              <button
                onClick={() => setTheme(theme === 'dark' ? 'light' : 'dark')}
                className="p-2 rounded-lg border border-zinc-200 dark:border-zinc-800 hover:bg-zinc-100 dark:hover:bg-zinc-800 text-zinc-600 dark:text-zinc-300 transition-colors"
                title="切换黑白主题"
              >
                {theme === 'dark' ? <Sun className="w-4 h-4" /> : <Moon className="w-4 h-4" />}
              </button>
            </div>
          </header>

          {/* Page Body Container */}
          <div className="flex-1 overflow-y-auto p-4 md:p-6">
            {currentPage === 'st-cards' ? (
              <div className="max-w-7xl mx-auto space-y-5">
                {/* Drag and Drop Upload Zone */}
                <div
                  onClick={() => uploadFileInputRef.current?.click()}
                  onDragOver={(e) => e.preventDefault()}
                  onDrop={(e) => {
                    e.preventDefault();
                    if (e.dataTransfer.files) handleFileUpload(e.dataTransfer.files);
                  }}
                  className="border-2 border-dashed border-zinc-200 dark:border-zinc-800 hover:border-zinc-400 dark:hover:border-zinc-600 rounded-xl p-6 text-center cursor-pointer transition-all bg-white/50 dark:bg-zinc-900/50 hover:bg-white dark:hover:bg-zinc-900"
                >
                  <p className="text-sm font-semibold text-zinc-800 dark:text-zinc-200">
                    拖放角色卡文件到此处，或点击上传文件
                  </p>
                  <p className="text-xs text-zinc-400 dark:text-zinc-500 mt-1">
                    支持 .png / .json / .webp 格式 SillyTavern 角色卡，支持多选批量上传
                  </p>
                </div>

                {/* Toolbar & Search Bar */}
                <div className="space-y-3">
                  <div className="flex flex-col sm:flex-row gap-2.5 items-stretch sm:items-center justify-between">
                    <div className="relative flex-1">
                      <Search className="w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 text-zinc-400" />
                      <input
                        type="text"
                        value={searchQuery}
                        onChange={(e) => setSearchQuery(e.target.value)}
                        placeholder="搜索角色名字、作者、标签或性格描述..."
                        className="w-full pl-9 pr-4 py-1.5 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400 dark:focus:ring-zinc-500"
                      />
                    </div>

                    <div className="flex items-center gap-2 flex-shrink-0">
                      {/* Sort Dropdown (Requirement 12) */}
                      <div className="flex items-center gap-1.5 bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-lg px-2.5 h-[30px]">
                        <ArrowUpDown className="w-3.5 h-3.5 text-zinc-400" />
                        <select
                          value={cardSortOrder}
                          onChange={(e) => setCardSortOrder(e.target.value as any)}
                          className="text-xs leading-none bg-transparent text-zinc-700 dark:text-zinc-300 focus:outline-none cursor-pointer"
                        >
                          <option value="default">默认排序</option>
                          <option value="az">名称 A-Z</option>
                          <option value="za">名称 Z-A</option>
                          <option value="newest">最新添加</option>
                          <option value="oldest">最早添加</option>
                        </select>
                      </div>
                    </div>
                  </div>

                  {/* Group Navigation Bar & Actions (Requirement 4) */}
                  <div className="bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-xl p-2 flex flex-col sm:flex-row items-stretch sm:items-center justify-between gap-2">
                    {/* Groups Scroll List */}
                    <div className="flex items-center gap-1.5 overflow-x-auto pb-1 sm:pb-0 scrollbar-none flex-1 pr-2">
                      {/* Default "全部" Group Tab (Requirement 4: "总卡片留下") */}
                      <button
                        onClick={() => setCurrentGroup('全部')}
                        className={`px-3 py-1.5 text-xs font-semibold rounded-lg whitespace-nowrap transition-all ${
                          currentGroup === '全部'
                            ? 'bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 shadow-sm'
                            : 'text-zinc-600 dark:text-zinc-400 hover:bg-zinc-100 dark:hover:bg-zinc-800'
                        }`}
                      >
                        总卡片 ({appData.cards.length})
                      </button>

                      {/* Custom Groups */}
                      {appData.groups.map((groupName) => {
                        const count = appData.cards.filter((c) => (c.group || '默认') === groupName).length;
                        const isSelected = currentGroup === groupName;

                        return (
                          <div
                            key={groupName}
                            className="relative group flex items-center"
                            onContextMenu={(e) => {
                              e.preventDefault();
                              setManagingGroup(groupName);
                              setRenameGroupInput(groupName);
                            }}
                          >
                            <button
                              onClick={() => setCurrentGroup(groupName)}
                              className={`px-3 py-1.5 text-xs font-medium rounded-lg whitespace-nowrap transition-all flex items-center gap-1.5 ${
                                isSelected
                                  ? 'bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 shadow-sm'
                                  : 'bg-zinc-100 dark:bg-zinc-800/60 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-200 dark:hover:bg-zinc-800'
                              }`}
                            >
                              <span>{groupName}</span>
                              <span className="text-[10px] opacity-70">({count})</span>
                            </button>

                            {/* Long press / Manage Icon on Hover */}
                            <button
                              onClick={(e) => {
                                e.stopPropagation();
                                setManagingGroup(groupName);
                                setRenameGroupInput(groupName);
                              }}
                              className="ml-0.5 p-1 text-zinc-400 hover:text-zinc-800 dark:hover:text-zinc-200 transition-opacity"
                              title="右键或长按管理分组"
                            >
                              <MoreHorizontal className="w-3 h-3" />
                            </button>
                          </div>
                        );
                      })}
                    </div>

                    {/* Group Buttons Right Side: 新建分组 & 选择 */}
                    <div className="flex items-center gap-2 flex-shrink-0 pt-2 sm:pt-0 border-t sm:border-t-0 border-zinc-100 dark:border-zinc-800">
                      <button
                        onClick={() => setShowNewGroupModal(true)}
                        className="flex items-center gap-1 px-2.5 py-1.5 text-xs font-medium rounded-lg bg-zinc-100 dark:bg-zinc-800 text-zinc-800 dark:text-zinc-200 hover:bg-zinc-200 dark:hover:bg-zinc-700 transition-colors"
                      >
                        <FolderPlus className="w-3.5 h-3.5" /> 新建分组
                      </button>

                      <button
                        onClick={() => {
                          setBatchMode((prev) => !prev);
                          setSelectedCardIds([]);
                        }}
                        className={`flex items-center gap-1 px-2.5 py-1.5 text-xs font-medium rounded-lg transition-colors ${
                          batchMode
                            ? 'bg-zinc-800 text-white dark:bg-zinc-200 dark:text-zinc-900'
                            : 'border border-zinc-200 dark:border-zinc-800 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800'
                        }`}
                      >
                        <CheckSquare className="w-3.5 h-3.5" /> {batchMode ? '退出选择' : '选择'}
                      </button>
                    </div>
                  </div>

                  {/* Batch Action Bar */}
                  {batchMode && (
                    <div className="bg-zinc-900 dark:bg-zinc-100 text-white dark:text-zinc-900 rounded-xl p-3 flex flex-col sm:flex-row items-center justify-between gap-3 shadow-lg animate-in fade-in slide-in-from-top-2">
                      <div className="text-xs font-medium">
                        已选择 <span className="font-bold underline">{selectedCardIds.length}</span> 张角色卡
                      </div>
                      <div className="flex items-center gap-2">
                        <button
                          onClick={() => {
                            if (selectedCardIds.length === appData.cards.length) setSelectedCardIds([]);
                            else setSelectedCardIds(filteredCards.map((c) => c.id));
                          }}
                          className="px-2.5 py-1 text-xs font-medium rounded bg-zinc-800 dark:bg-zinc-200 text-zinc-200 dark:text-zinc-800 hover:opacity-80"
                        >
                          {selectedCardIds.length === filteredCards.length ? '取消全选' : '全选当前'}
                        </button>
                        <button
                          disabled={selectedCardIds.length === 0}
                          onClick={() => setShowBatchMoveModal(true)}
                          className="px-2.5 py-1 text-xs font-medium rounded bg-zinc-800 dark:bg-zinc-200 text-zinc-200 dark:text-zinc-800 hover:opacity-80 disabled:opacity-40"
                        >
                          批量移动
                        </button>
                        <button
                          disabled={selectedCardIds.length === 0}
                          onClick={handleBatchDeleteCards}
                          className="px-2.5 py-1 text-xs font-medium rounded bg-rose-600 text-white hover:bg-rose-700 disabled:opacity-40"
                        >
                          批量删除
                        </button>
                      </div>
                    </div>
                  )}
                </div>

                {/* Cards Grid Display */}
                {filteredCards.length === 0 ? (
                  <div className="text-center py-20 bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-xl p-8">
                    <p className="text-sm text-zinc-400 dark:text-zinc-500">
                      {searchQuery ? '未查找到匹配的角色卡' : '该分组下暂无角色卡，拖放文件到上方框内进行上传'}
                    </p>
                  </div>
                ) : (
                  <div className="grid grid-cols-3 sm:grid-cols-4 md:grid-cols-5 lg:grid-cols-6 xl:grid-cols-7 gap-2.5">
                    {visibleFilteredCards.map((card) => {
                      const name = getCardDisplayName(card);
                      const author = getCardCreator(card);
                      const tags = getCardTags(card).slice(0, 3);
                      const isSelected = selectedCardIds.includes(card.id);

                      return (
                        <div
                          key={card.id}
                          onClick={() => {
                            if (batchMode) {
                              if (isSelected) setSelectedCardIds((p) => p.filter((id) => id !== card.id));
                              else setSelectedCardIds((p) => [...p, card.id]);
                            } else {
                              setDetailCardId(card.id);
                              setDetailTab('overview');
                            }
                          }}
                          className={`group bg-white dark:bg-zinc-900 border rounded-lg overflow-hidden cursor-pointer transition-all hover:shadow-md flex flex-col relative ${
                            isSelected
                              ? 'border-zinc-900 dark:border-zinc-100 ring-2 ring-zinc-900 dark:ring-zinc-100'
                              : 'border-zinc-200 dark:border-zinc-800 hover:border-zinc-400 dark:hover:border-zinc-600'
                          }`}
                        >
                          {/* Card Cover Image */}
                          <div className="aspect-[2/3] w-full bg-zinc-100 dark:bg-zinc-800 relative overflow-hidden flex items-center justify-center">
                            {card.coverImage ? (
                              <img src={card.coverImage} alt={name} className="w-full h-full object-contain group-hover:scale-105 transition-transform duration-300" />
                            ) : (
                              <div className="text-[10px] text-zinc-400 flex flex-col items-center gap-0.5">
                                <ImageIcon className="w-5 h-5 opacity-40" />
                                <span>无封面</span>
                              </div>
                            )}

                            {/* Format Badge */}
                            <span className="absolute top-1.5 right-1.5 px-1 py-0.5 rounded bg-black/70 text-white text-[9px] font-bold uppercase tracking-wider backdrop-blur-sm">
                              {card.version}
                            </span>

                            {/* Batch Selection Checkbox */}
                            {batchMode && (
                              <div className="absolute top-1.5 left-1.5">
                                {isSelected ? (
                                  <CheckCircle2 className="w-4 h-4 text-zinc-900 dark:text-zinc-100 fill-white dark:fill-zinc-900" />
                                ) : (
                                  <Circle className="w-4 h-4 text-white/80 drop-shadow" />
                                )}
                              </div>
                            )}
                          </div>

                          {/* Card Info Details */}
                          <div className="p-2 flex-1 flex flex-col justify-between space-y-1">
                            <div>
                              <h3 className="text-[11px] font-bold truncate text-zinc-900 dark:text-zinc-100 leading-tight" title={name}>
                                {name}
                              </h3>
                              <p className="text-[10px] text-zinc-400 dark:text-zinc-500 truncate mt-0.5">
                                by {author}
                              </p>
                            </div>

                            {/* Tags / Group badge */}
                            <div className="flex flex-wrap gap-1">
                              {tags.map((t, idx) => (
                                <span key={idx} className="px-1 py-0.5 text-[8px] bg-zinc-100 dark:bg-zinc-800 text-zinc-600 dark:text-zinc-400 rounded leading-none">
                                  {t}
                                </span>
                              ))}
                              <span className="px-1.5 py-0.5 text-[9px] bg-zinc-200/60 dark:bg-zinc-800 text-zinc-500 rounded ml-auto">
                                {card.group || '默认'}
                              </span>
                            </div>
                          </div>

                          {/* Quick Action buttons */}
                          {!batchMode && (
                            <div className="p-2 pt-0 flex gap-1 border-t border-zinc-100 dark:border-zinc-800/60">
                              <button
                                onClick={(e) => {
                                  e.stopPropagation();
                                  setDetailCardId(card.id);
                                  setDetailTab('overview');
                                }}
                                className="flex-1 py-1 text-[11px] font-medium rounded hover:bg-zinc-100 dark:hover:bg-zinc-800 text-zinc-700 dark:text-zinc-300"
                              >
                                查看
                              </button>
                              <button
                                onClick={(e) => {
                                  e.stopPropagation();
                                  exportAsJson(card);
                                }}
                                className="flex-1 py-1 text-[11px] font-medium rounded hover:bg-zinc-100 dark:hover:bg-zinc-800 text-zinc-700 dark:text-zinc-300"
                              >
                                导出
                              </button>
                              <button
                                onClick={(e) => {
                                  e.stopPropagation();
                                  handleDeleteCard(card.id);
                                }}
                                className="px-2 py-1 text-[11px] font-medium rounded hover:bg-rose-50 dark:hover:bg-rose-950/50 text-rose-600 dark:text-rose-400"
                              >
                                <Trash2 className="w-3 h-3" />
                              </button>
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                )}
                {filteredCards.length > cardVisibleCount && (
                  <div ref={cardLoadMoreRef} className="flex flex-col items-center justify-center py-6 gap-2">
                    <p className="text-[11px] text-zinc-400">已显示 {cardVisibleCount} / {filteredCards.length} 张角色卡</p>
                    <button
                      onClick={() => setCardVisibleCount((prev) => Math.min(prev + CARD_PAGE_SIZE, filteredCards.length))}
                      className="px-4 py-1.5 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-800 hover:bg-zinc-100 dark:hover:bg-zinc-800 text-zinc-700 dark:text-zinc-300 transition-colors"
                    >
                      加载更多
                    </button>
                  </div>
                )}
              </div>
            ) : currentPage === 'st-mobile' ? (
              <div className="max-w-7xl mx-auto space-y-5">
                {/* Search Bar & Batch Mode Controls */}
                <div className="flex flex-col sm:flex-row gap-3 items-stretch sm:items-center justify-between">
                  <div className="relative flex-1">
                    <Search className="w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 text-zinc-400" />
                    <input
                      type="text"
                      value={phoneSearchQuery}
                      onChange={(e) => setPhoneSearchQuery(e.target.value)}
                      placeholder="搜索小手机名称、链接、联系方式或描述..."
                      className="w-full pl-9 pr-4 py-1.5 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400 dark:focus:ring-zinc-500"
                    />
                  </div>

                  <div className="flex items-center gap-2">
                    <button
                      onClick={() => {
                        setPhoneBatchMode(!phoneBatchMode);
                        setSelectedPhoneIds([]);
                      }}
                      className={`px-3.5 py-2 text-xs font-medium rounded-lg border transition-colors flex items-center gap-1.5 whitespace-nowrap ${
                        phoneBatchMode
                          ? 'bg-amber-100 text-amber-900 border-amber-300 dark:bg-amber-950 dark:text-amber-200 dark:border-amber-800'
                          : 'border-zinc-200 dark:border-zinc-800 hover:bg-zinc-100 dark:hover:bg-zinc-800 text-zinc-700 dark:text-zinc-300'
                      }`}
                    >
                      {phoneBatchMode ? '退出选择' : '选择'}
                    </button>

                    {phoneBatchMode && (
                      <>
                        <button
                          onClick={() => {
                            const allIds = filteredPhoneLinks.map((item) => item.id);
                            if (selectedPhoneIds.length === allIds.length) {
                              setSelectedPhoneIds([]);
                            } else {
                              setSelectedPhoneIds(allIds);
                            }
                          }}
                          className="px-3 py-2 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-800 hover:bg-zinc-100 dark:hover:bg-zinc-800 text-zinc-700 dark:text-zinc-300 whitespace-nowrap"
                        >
                          {selectedPhoneIds.length === filteredPhoneLinks.length && filteredPhoneLinks.length > 0
                            ? '取消全选'
                            : '全选'}
                        </button>

                        {selectedPhoneIds.length > 0 && (
                          <button
                            onClick={handleBatchDeletePhoneLinks}
                            className="px-3.5 py-2 text-xs font-bold rounded-lg bg-rose-600 text-white hover:bg-rose-700 flex items-center gap-1 whitespace-nowrap"
                          >
                            <Trash2 className="w-3.5 h-3.5" /> 删除 ({selectedPhoneIds.length})
                          </button>
                        )}
                      </>
                    )}
                  </div>
                </div>

                {/* List of Mobile Links */}
                {filteredPhoneLinks.length === 0 ? (
                  <div className="text-center py-16 bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-xl space-y-3">
                    <p className="text-sm font-medium text-zinc-600 dark:text-zinc-400">
                      {phoneSearchQuery ? '未找到匹配的小手机链接' : '暂无小手机链接'}
                    </p>
                    <p className="text-xs text-zinc-400">
                      点击右上角的 “+” 按钮，添加新的小手机链接
                    </p>
                  </div>
                ) : (
                  <div className="space-y-3">
                    {filteredPhoneLinks.map((item) => {
                      const isSelected = selectedPhoneIds.includes(item.id);

                      return (
                        <div
                          key={item.id}
                          onClick={() => {
                            if (phoneBatchMode) {
                              if (isSelected) {
                                setSelectedPhoneIds((p) => p.filter((id) => id !== item.id));
                              } else {
                                setSelectedPhoneIds((p) => [...p, item.id]);
                              }
                            } else {
                              setEditingPhoneLink({ ...item });
                            }
                          }}
                          className={`p-4 bg-white dark:bg-zinc-900 border rounded-xl cursor-pointer transition-all hover:shadow-md flex items-center justify-between gap-4 ${
                            isSelected
                              ? 'border-rose-500 ring-2 ring-rose-500/30 dark:ring-rose-500/30'
                              : 'border-zinc-200 dark:border-zinc-800 hover:border-zinc-300 dark:hover:border-zinc-700'
                          }`}
                        >
                          {/* Name on Line 1, Link on Line 2 */}
                          <div className="flex-1 min-w-0 space-y-1">
                            <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100 truncate" title={item.name}>
                              {item.name}
                            </h3>
                            <p className="text-xs text-zinc-500 dark:text-zinc-400 truncate font-mono" title={item.url}>
                              {item.url}
                            </p>
                          </div>

                          {phoneBatchMode ? (
                            <div className="flex-shrink-0">
                              {isSelected ? (
                                <CheckCircle2 className="w-5 h-5 text-rose-500 fill-white dark:fill-zinc-900" />
                              ) : (
                                <Circle className="w-5 h-5 text-zinc-300 dark:text-zinc-600" />
                              )}
                            </div>
                          ) : (
                            <div className="flex-shrink-0 text-xs text-zinc-400 dark:text-zinc-500 flex items-center gap-1">
                              <span>编辑 / 查看</span>
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>
            ) : currentPage === 'st-themes' ? (
              <div className="max-w-7xl mx-auto space-y-4">
                {/* ST主题搜索栏独立放在页面顶部 */}
                <div className="relative w-full">
                  <Search className="w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 text-zinc-400" />
                  <input
                    type="text"
                    value={themeSearchQuery}
                    onChange={(e) => setThemeSearchQuery(e.target.value)}
                    placeholder="搜索 ST 主题名称、作者、类型、分类..."
                    className="w-full pl-9 pr-4 py-1.5 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400 dark:focus:ring-zinc-500"
                  />
                </div>

                {/* Category Group Bar */}
                <div className="bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-xl p-2.5 flex flex-col items-stretch gap-2">
                  {/* Category Group Scroll Tabs */}
                  <div className="w-full flex items-center gap-1.5 overflow-x-auto pb-1 scrollbar-none min-w-0">
                    {/* "全部分组" Tab */}
                    <button
                      onClick={() => setThemeCategoryFilter('全部分类')}
                      className={`px-3 py-1.5 text-xs font-semibold rounded-lg whitespace-nowrap transition-all ${
                        themeCategoryFilter === '全部分类'
                          ? 'bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 shadow-sm'
                          : 'text-zinc-600 dark:text-zinc-400 hover:bg-zinc-100 dark:hover:bg-zinc-800'
                      }`}
                    >
                      全部分组 ({themesList.length})
                    </button>

                    {/* Custom Categories */}
                    {Array.from(
                      new Set(['默认', ...(appData.themeCategories || []), ...themesList.map((t) => t.category || '默认')])
                    ).map((catName) => {
                      const count = themesList.filter((t) => (t.category || '默认') === catName).length;
                      const isSelected = themeCategoryFilter === catName;

                      return (
                        <div
                          key={catName}
                          className="relative group flex items-center"
                          onContextMenu={(e) => {
                            e.preventDefault();
                            setManagingThemeGroup(catName);
                            setRenameThemeGroupInput(catName);
                          }}
                        >
                          <button
                            onClick={() => setThemeCategoryFilter(catName)}
                            className={`px-3 py-1.5 text-xs font-medium rounded-lg whitespace-nowrap transition-all flex items-center gap-1.5 ${
                              isSelected
                                ? 'bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 shadow-sm'
                                : 'bg-zinc-100 dark:bg-zinc-800/60 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-200 dark:hover:bg-zinc-800'
                            }`}
                          >
                            <span>{catName}</span>
                            <span className="text-[10px] opacity-70">({count})</span>
                          </button>

                          {/* Long Press / Right-click / Manage Icon */}
                          <button
                            onClick={(e) => {
                              e.stopPropagation();
                              setManagingThemeGroup(catName);
                              setRenameThemeGroupInput(catName);
                            }}
                            className="ml-0.5 p-1 text-zinc-400 hover:text-zinc-800 dark:hover:text-zinc-200 transition-opacity"
                            title="右键或长按/点击管理分组"
                          >
                            <MoreHorizontal className="w-3 h-3" />
                          </button>
                        </div>
                      );
                    })}

                    {/* "新建分组" Button on the right of categories */}
                    <button
                      onClick={() => setShowNewThemeGroupModal(true)}
                      className="flex items-center gap-1 px-2.5 py-1.5 text-xs font-medium rounded-lg bg-zinc-100 dark:bg-zinc-800 text-zinc-800 dark:text-zinc-200 hover:bg-zinc-200 dark:hover:bg-zinc-700 transition-colors whitespace-nowrap ml-auto flex-shrink-0"
                    >
                      <FolderPlus className="w-3.5 h-3.5" /> 新建分组
                    </button>
                  </div>

                  {/* Right Action: Sorting & "选择" (Batch Mode) */}
                  <div className="w-full flex items-center justify-start gap-2 flex-shrink-0">
                    <div className="flex items-center gap-1.5 bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-lg px-2.5 h-[30px] flex-shrink-0">
                      <ArrowUpDown className="w-3.5 h-3.5 text-zinc-400" />
                      <select
                        value={themeSortOrder}
                        onChange={(e) => setThemeSortOrder(e.target.value as any)}
                        className="text-xs leading-none bg-transparent text-zinc-700 dark:text-zinc-300 focus:outline-none cursor-pointer"
                      >
                        <option value="default">默认排序</option>
                        <option value="az">名称 A-Z</option>
                        <option value="za">名称 Z-A</option>
                        <option value="newest">最新添加</option>
                        <option value="oldest">最早添加</option>
                      </select>
                    </div>

                    <button
                      onClick={() => {
                        setThemeBatchMode(!themeBatchMode);
                        setSelectedThemeIds([]);
                      }}
                      className={`px-3.5 py-1.5 text-xs font-medium rounded-lg border transition-colors flex items-center gap-1.5 whitespace-nowrap ${
                        themeBatchMode
                          ? 'bg-amber-100 text-amber-900 border-amber-300 dark:bg-amber-950 dark:text-amber-200 dark:border-amber-800'
                          : 'border-zinc-200 dark:border-zinc-800 hover:bg-zinc-100 dark:hover:bg-zinc-800 text-zinc-700 dark:text-zinc-300'
                      }`}
                    >
                      <CheckSquare className="w-3.5 h-3.5" /> {themeBatchMode ? '退出选择' : '选择'}
                    </button>
                  </div>
                </div>

                {/* Batch Action Toolbar when in Batch Mode */}
                {themeBatchMode && (
                  <div className="bg-zinc-900 dark:bg-zinc-100 text-white dark:text-zinc-900 rounded-xl p-3 flex flex-col sm:flex-row items-center justify-between gap-3 shadow-lg animate-in fade-in slide-in-from-top-2">
                    <div className="text-xs font-medium">
                      已选择 <span className="font-bold underline">{selectedThemeIds.length}</span> 个 ST 主题
                    </div>
                    <div className="flex items-center gap-2">
                      <button
                        onClick={() => {
                          if (selectedThemeIds.length === filteredThemes.length && filteredThemes.length > 0) {
                            setSelectedThemeIds([]);
                          } else {
                            setSelectedThemeIds(filteredThemes.map((item) => item.id));
                          }
                        }}
                        className="px-2.5 py-1 text-xs font-medium rounded bg-zinc-800 dark:bg-zinc-200 text-zinc-200 dark:text-zinc-800 hover:opacity-80"
                      >
                        {selectedThemeIds.length === filteredThemes.length && filteredThemes.length > 0
                          ? '取消全选'
                          : '全选当前'}
                      </button>

                      <button
                        disabled={selectedThemeIds.length === 0}
                        onClick={() => setShowThemeBatchMoveModal(true)}
                        className="px-2.5 py-1 text-xs font-medium rounded bg-zinc-800 dark:bg-zinc-200 text-zinc-200 dark:text-zinc-800 hover:opacity-80 disabled:opacity-40 flex items-center gap-1"
                      >
                        <Move className="w-3 h-3" /> 批量移动
                      </button>

                      <button
                        disabled={selectedThemeIds.length === 0}
                        onClick={handleBatchDeleteThemes}
                        className="px-2.5 py-1 text-xs font-medium rounded bg-rose-600 text-white hover:bg-rose-700 disabled:opacity-40 flex items-center gap-1"
                      >
                        <Trash2 className="w-3 h-3" /> 批量删除
                      </button>
                    </div>
                  </div>
                )}

                {/* Gallery List of ST Themes ("图集形式") */}
                {filteredThemes.length === 0 ? (
                  <div className="text-center py-16 bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-xl space-y-3">
                    <p className="text-sm font-medium text-zinc-600 dark:text-zinc-400">
                      {themeSearchQuery || themeCategoryFilter !== '全部分类'
                        ? '未找到匹配的 ST 主题'
                        : '暂无 ST 主题'}
                    </p>
                    <p className="text-xs text-zinc-400">
                      点击右上角的 “+” 按钮，上传并导入 ST 主题文档 (.docx, .txt, .json, .css)
                    </p>
                  </div>
                ) : (
                  <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-5 xl:grid-cols-6 gap-4 sm:gap-6">
                    {sortItemList(
                      filteredThemes,
                      themeSortOrder,
                      (item) => item.name || '',
                      (item) => item.createdAt || 0
                    ).map((item) => {
                      const isSelected = selectedThemeIds.includes(item.id);

                      return (
                        <div
                          key={item.id}
                          onClick={() => {
                            if (themeBatchMode) {
                              if (isSelected) {
                                setSelectedThemeIds((p) => p.filter((id) => id !== item.id));
                              } else {
                                setSelectedThemeIds((p) => [...p, item.id]);
                              }
                            } else {
                              setEditingTheme({ ...item });
                              setThemeDetailTab('info');
                              setCodeSearchQuery('');
                            }
                          }}
                          className={`group bg-white dark:bg-zinc-900 border rounded-xl overflow-hidden cursor-pointer transition-all hover:shadow-lg flex flex-col relative ${
                            isSelected
                              ? 'border-rose-500 ring-2 ring-rose-500/30 dark:ring-rose-500/30'
                              : 'border-zinc-200 dark:border-zinc-800 hover:border-zinc-300 dark:hover:border-zinc-700'
                          }`}
                        >
                          {/* Card Cover (aspect-[2/3] for gallery format) */}
                          <div className="aspect-[2/3] w-full bg-zinc-100 dark:bg-zinc-800/80 relative overflow-hidden flex items-center justify-center p-2">
                            {item.coverImage ? (
                              <img
                                src={item.coverImage}
                                alt={item.name}
                                className="w-full h-full object-cover transition-transform duration-300 group-hover:scale-105"
                              />
                            ) : (
                              <div className="w-full h-full flex flex-col items-center justify-center text-center p-3 border border-dashed border-zinc-300 dark:border-zinc-700 rounded-lg bg-zinc-50/50 dark:bg-zinc-900/40 space-y-2">
                                <FileText className="w-8 h-8 text-zinc-400 dark:text-zinc-600" />
                                <span className="text-[11px] font-bold text-zinc-600 dark:text-zinc-400 line-clamp-2 leading-tight">
                                  {item.name}
                                </span>
                                <span className="text-[9px] uppercase font-mono text-zinc-400 dark:text-zinc-500">
                                  {item.fileType || '主题'}
                                </span>
                              </div>
                            )}

                            {/* Batch Selection Checkbox */}
                            {themeBatchMode && (
                              <div className="absolute top-2 left-2 z-10 bg-black/50 rounded-full p-0.5">
                                {isSelected ? (
                                  <CheckCircle2 className="w-5 h-5 text-rose-500 fill-white dark:fill-zinc-900" />
                                ) : (
                                  <Circle className="w-5 h-5 text-white/80" />
                                )}
                              </div>
                            )}

                            {/* Category or Type Badge */}
                            {(item.type || (item.category && item.category !== '默认')) && (
                              <span className="absolute top-2 right-2 px-1.5 py-0.5 text-[9px] font-bold rounded bg-black/60 text-white backdrop-blur-sm">
                                {item.type || item.category}
                              </span>
                            )}
                          </div>

                          {/* Theme Title & Author */}
                          <div className="p-3 space-y-1 bg-white dark:bg-zinc-900 border-t border-zinc-100 dark:border-zinc-800">
                            <h3 className="text-xs font-bold text-zinc-900 dark:text-zinc-100 truncate" title={item.name}>
                              {item.name}
                            </h3>
                            <div className="flex items-center justify-between text-[10px] text-zinc-400">
                              <span className="truncate">作者：{item.author || '默认'}</span>
                              <span className="uppercase text-[9px]">{item.fileType || '文档'}</span>
                            </div>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>
            ) : currentPage === 'chat-memes' ? (
              <div className="max-w-7xl mx-auto space-y-5">
                {/* 聊天梗：搜索 / 分组 / 选择控制区，与 ST 角色卡保持相同页面边距 */}
                <div className="bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-xl p-2.5 space-y-2">
                  <div className="flex flex-wrap items-center gap-2">
                    <div className="relative w-full sm:w-64 flex-shrink-0">
                      <Search className="w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 text-zinc-400" />
                      <input
                        type="text"
                        value={chatMemeSearchQuery}
                        onChange={(e) => setChatMemeSearchQuery(e.target.value)}
                        placeholder="搜索聊天梗..."
                        className="w-full pl-9 pr-3 py-1.5 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400"
                      />
                    </div>
                    <div className="flex items-center gap-1.5 overflow-x-auto scrollbar-none min-w-0 flex-1">
                      <button
                        onClick={() => setChatMemeCategoryFilter('全部分组')}
                        className={`px-3 py-1.5 text-xs font-semibold rounded-lg whitespace-nowrap ${chatMemeCategoryFilter === '全部分组' ? 'bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900' : 'text-zinc-600 dark:text-zinc-400 hover:bg-zinc-100 dark:hover:bg-zinc-800'}`}
                      >全部分组 ({chatMemesList.length})</button>
                      {(appData.chatMemeCategories || ['默认']).map((cat) => {
                        const count = chatMemesList.filter((m) => (m.category || '默认') === cat).length;
                        const onLongPressStart = () => {
                          if (chatMemeGroupPressTimer.current) clearTimeout(chatMemeGroupPressTimer.current);
                          chatMemeGroupPressTimer.current = setTimeout(() => {
                            setManagingChatMemeGroup(cat);
                            setRenameChatMemeGroupInput(cat);
                          }, 650);
                        };
                        const clearLongPress = () => {
                          if (chatMemeGroupPressTimer.current) clearTimeout(chatMemeGroupPressTimer.current);
                          chatMemeGroupPressTimer.current = null;
                        };
                        return (
                          <button
                            key={cat}
                            onClick={() => setChatMemeCategoryFilter(cat)}
                            onContextMenu={(e) => { e.preventDefault(); setManagingChatMemeGroup(cat); setRenameChatMemeGroupInput(cat); }}
                            onPointerDown={onLongPressStart}
                            onPointerUp={clearLongPress}
                            onPointerCancel={clearLongPress}
                            className={`px-3 py-1.5 text-xs font-medium rounded-lg whitespace-nowrap ${chatMemeCategoryFilter === cat ? 'bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900' : 'bg-zinc-100 dark:bg-zinc-800/60 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-200 dark:hover:bg-zinc-800'}`}
                            title="长按或右键管理分组"
                          >{cat} ({count})</button>
                        );
                      })}
                      <button
                        onClick={() => setShowNewChatMemeGroupModal(true)}
                        className="flex items-center gap-1 px-2.5 py-1.5 text-xs font-medium rounded-lg bg-zinc-100 dark:bg-zinc-800 text-zinc-800 dark:text-zinc-200 hover:bg-zinc-200 dark:hover:bg-zinc-700 whitespace-nowrap ml-auto flex-shrink-0"
                      ><FolderPlus className="w-3.5 h-3.5" /> 新建分组</button>
                    </div>
                    <div className="w-full flex items-center justify-start gap-2 flex-shrink-0">
                      <div className="flex items-center gap-1.5 bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-lg px-2.5 h-[30px] flex-shrink-0">
                        <ArrowUpDown className="w-3.5 h-3.5 text-zinc-400" />
                        <select
                          value={chatMemeSortOrder}
                          onChange={(e) => setChatMemeSortOrder(e.target.value as typeof chatMemeSortOrder)}
                          className="text-xs leading-none bg-transparent text-zinc-700 dark:text-zinc-300 focus:outline-none cursor-pointer"
                        >
                          <option value="default">默认排序</option>
                          <option value="az">名称 A-Z</option>
                          <option value="za">名称 Z-A</option>
                          <option value="newest">最新添加</option>
                          <option value="oldest">最早添加</option>
                        </select>
                      </div>
                      <button
                        onClick={() => { setChatMemeBatchMode((v) => !v); setSelectedChatMemeIds([]); }}
                        className={`px-3.5 py-1.5 text-xs font-medium rounded-lg border whitespace-nowrap flex items-center gap-1.5 ${chatMemeBatchMode ? 'bg-amber-100 text-amber-900 border-amber-300 dark:bg-amber-950 dark:text-amber-200' : 'border-zinc-200 dark:border-zinc-800 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800'}`}
                      ><CheckSquare className="w-3.5 h-3.5" /> {chatMemeBatchMode ? '退出选择' : '选择'}</button>
                    </div>
                  </div>
                </div>

                {chatMemeBatchMode && (
                  <div className="bg-zinc-900 dark:bg-zinc-100 text-white dark:text-zinc-900 rounded-xl p-3 flex flex-wrap items-center justify-between gap-2 shadow-lg">
                    <div className="text-xs font-medium">已选择 <span className="font-bold underline">{selectedChatMemeIds.length}</span> 条聊天梗</div>
                    <div className="flex flex-wrap items-center gap-2">
                      <button onClick={() => setSelectedChatMemeIds(selectedChatMemeIds.length === filteredChatMemes.length ? [] : filteredChatMemes.map((m) => m.id))} className="px-2.5 py-1 text-xs rounded bg-zinc-800 dark:bg-zinc-200 text-zinc-200 dark:text-zinc-800">{selectedChatMemeIds.length === filteredChatMemes.length && filteredChatMemes.length ? '取消全选' : '全选当前'}</button>
                      <select
                        defaultValue=""
                        onChange={(e) => { handleMoveSelectedChatMemes(e.target.value); e.currentTarget.value = ''; }}
                        disabled={!selectedChatMemeIds.length}
                        className="px-2.5 py-1 text-xs rounded bg-zinc-800 dark:bg-zinc-200 text-zinc-200 dark:text-zinc-800 disabled:opacity-40"
                      ><option value="">移动到分组</option>{(appData.chatMemeCategories || ['默认']).map((g) => <option key={g} value={g}>{g}</option>)}</select>
                      <button disabled={!selectedChatMemeIds.length} onClick={() => setShowChatExportModal(true)} className="px-2.5 py-1 text-xs rounded bg-zinc-800 dark:bg-zinc-200 text-zinc-200 dark:text-zinc-800 disabled:opacity-40"><Download className="w-3 h-3 inline mr-1" />导出</button>
                      <button disabled={!selectedChatMemeIds.length} onClick={handleDeleteSelectedChatMemes} className="px-2.5 py-1 text-xs rounded bg-rose-600 text-white disabled:opacity-40"><Trash2 className="w-3 h-3 inline mr-1" />批量删除</button>
                    </div>
                  </div>
                )}

                <div className="space-y-2">
                  {filteredChatMemes.length === 0 ? (
                    <div className="min-h-[280px] flex items-center justify-center rounded-2xl border border-dashed border-zinc-200 dark:border-zinc-800 text-xs text-zinc-400 bg-white/60 dark:bg-zinc-900/60">暂无聊天梗</div>
                  ) : filteredChatMemes.map((item) => (
                    <div key={item.id} onClick={() => chatMemeBatchMode ? setSelectedChatMemeIds((ids) => ids.includes(item.id) ? ids.filter((id) => id !== item.id) : [...ids, item.id]) : setEditingChatMeme(item)} className={`relative w-full text-left p-4 rounded-xl border bg-white dark:bg-zinc-900 cursor-pointer transition-colors ${selectedChatMemeIds.includes(item.id) ? 'border-zinc-900 dark:border-zinc-100 ring-1 ring-zinc-900/10' : 'border-zinc-200 dark:border-zinc-800 hover:bg-zinc-50 dark:hover:bg-zinc-800/60'}`}>
                      {chatMemeBatchMode && <div className={`absolute top-3 right-3 w-5 h-5 rounded-full border-2 flex items-center justify-center ${selectedChatMemeIds.includes(item.id) ? 'bg-zinc-900 dark:bg-zinc-100 border-zinc-900 dark:border-zinc-100' : 'border-zinc-300 dark:border-zinc-600'}`}>{selectedChatMemeIds.includes(item.id) && <Check className="w-3 h-3 text-white dark:text-zinc-900" />}</div>}
                      <div className="pr-8 text-sm text-zinc-800 dark:text-zinc-200 leading-relaxed whitespace-pre-wrap">{item.content}</div>
                    </div>
                  ))}
                </div>
              </div>
            ) : currentPage === 'themes' ? (
              <div className="max-w-7xl mx-auto space-y-4">
                {/* Search Bar & Category Group Bar */}
                <div className="bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-xl p-2.5 flex flex-col md:flex-row items-stretch md:items-center justify-start gap-3">
                  {/* Search Input (Compact, w-full md:w-72) */}
                  <div className="relative w-full md:w-72 flex-shrink-0">
                    <Search className="w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 text-zinc-400" />
                    <input
                      type="text"
                      value={beautificationSearchQuery}
                      onChange={(e) => setBeautificationSearchQuery(e.target.value)}
                      placeholder="搜索美化名称、作者、类型、分类..."
                      className="w-full pl-9 pr-4 py-1.5 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400 dark:focus:ring-zinc-500"
                    />
                  </div>

                  {/* Category Group Scroll Tabs */}
                  <div className="flex items-center gap-1.5 overflow-x-auto pb-1 md:pb-0 scrollbar-none flex-none min-w-0 max-w-full">
                    {/* "全部分组" Tab */}
                    <button
                      onClick={() => setBeautificationCategoryFilter('全部分类')}
                      className={`px-3 py-1.5 text-xs font-semibold rounded-lg whitespace-nowrap transition-all ${
                        beautificationCategoryFilter === '全部分类'
                          ? 'bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 shadow-sm'
                          : 'text-zinc-600 dark:text-zinc-400 hover:bg-zinc-100 dark:hover:bg-zinc-800'
                      }`}
                    >
                      全部分组 ({beautificationsList.length})
                    </button>

                    {/* Custom Categories */}
                    {Array.from(
                      new Set(['默认', ...(appData.beautificationCategories || []), ...beautificationsList.map((b) => b.category || '默认')])
                    ).map((catName) => {
                      const count = beautificationsList.filter((b) => (b.category || '默认') === catName).length;
                      const isSelected = beautificationCategoryFilter === catName;

                      return (
                        <div
                          key={catName}
                          className="relative group flex items-center"
                          onContextMenu={(e) => {
                            e.preventDefault();
                            setManagingBeautificationGroup(catName);
                            setRenameBeautificationGroupInput(catName);
                          }}
                        >
                          <button
                            onClick={() => setBeautificationCategoryFilter(catName)}
                            className={`px-3 py-1.5 text-xs font-medium rounded-lg whitespace-nowrap transition-all flex items-center gap-1.5 ${
                              isSelected
                                ? 'bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 shadow-sm'
                                : 'bg-zinc-100 dark:bg-zinc-800/60 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-200 dark:hover:bg-zinc-800'
                            }`}
                          >
                            <span>{catName}</span>
                            <span className="text-[10px] opacity-70">({count})</span>
                          </button>

                          {/* Long Press / Right-click / Manage Icon */}
                          <button
                            onClick={(e) => {
                              e.stopPropagation();
                              setManagingBeautificationGroup(catName);
                              setRenameBeautificationGroupInput(catName);
                            }}
                            className="ml-0.5 p-1 text-zinc-400 hover:text-zinc-800 dark:hover:text-zinc-200 transition-opacity"
                            title="右键或长按/点击管理分组"
                          >
                            <MoreHorizontal className="w-3 h-3" />
                          </button>
                        </div>
                      );
                    })}

                    {/* "新建分组" Button on the right of categories */}
                    <button
                      onClick={() => setShowNewBeautificationGroupModal(true)}
                      className="flex items-center gap-1 px-2.5 py-1.5 text-xs font-medium rounded-lg bg-zinc-100 dark:bg-zinc-800 text-zinc-800 dark:text-zinc-200 hover:bg-zinc-200 dark:hover:bg-zinc-700 transition-colors whitespace-nowrap ml-auto flex-shrink-0"
                    >
                      <FolderPlus className="w-3.5 h-3.5" /> 新建分组
                    </button>
                  </div>

                  {/* Right Action: Sorting & "选择" (Batch Mode) */}
                  <div className="flex items-center justify-start gap-2 flex-shrink-0">
                    <div className="flex items-center gap-1.5 bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-lg px-2.5 h-[30px] flex-shrink-0">
                      <ArrowUpDown className="w-3.5 h-3.5 text-zinc-400" />
                      <select
                        value={beautificationSortOrder}
                        onChange={(e) => setBeautificationSortOrder(e.target.value as any)}
                        className="text-xs leading-none bg-transparent text-zinc-700 dark:text-zinc-300 focus:outline-none cursor-pointer"
                      >
                        <option value="default">默认排序</option>
                        <option value="az">名称 A-Z</option>
                        <option value="za">名称 Z-A</option>
                        <option value="newest">最新添加</option>
                        <option value="oldest">最早添加</option>
                      </select>
                    </div>

                    <button
                      onClick={() => {
                        setBeautificationBatchMode(!beautificationBatchMode);
                        setSelectedBeautificationIds([]);
                      }}
                      className={`px-3.5 py-1.5 text-xs font-medium rounded-lg border transition-colors flex items-center gap-1.5 whitespace-nowrap ${
                        beautificationBatchMode
                          ? 'bg-amber-100 text-amber-900 border-amber-300 dark:bg-amber-950 dark:text-amber-200 dark:border-amber-800'
                          : 'border-zinc-200 dark:border-zinc-800 hover:bg-zinc-100 dark:hover:bg-zinc-800 text-zinc-700 dark:text-zinc-300'
                      }`}
                    >
                      <CheckSquare className="w-3.5 h-3.5" /> {beautificationBatchMode ? '退出选择' : '选择'}
                    </button>
                  </div>
                </div>

                {/* Batch Action Toolbar when in Batch Mode */}
                {beautificationBatchMode && (
                  <div className="bg-zinc-900 dark:bg-zinc-100 text-white dark:text-zinc-900 rounded-xl p-3 flex flex-col sm:flex-row items-center justify-between gap-3 shadow-lg animate-in fade-in slide-in-from-top-2">
                    <div className="text-xs font-medium">
                      已选择 <span className="font-bold underline">{selectedBeautificationIds.length}</span> 个美化
                    </div>
                    <div className="flex items-center gap-2">
                      <button
                        onClick={() => {
                          if (selectedBeautificationIds.length === filteredBeautifications.length && filteredBeautifications.length > 0) {
                            setSelectedBeautificationIds([]);
                          } else {
                            setSelectedBeautificationIds(filteredBeautifications.map((item) => item.id));
                          }
                        }}
                        className="px-2.5 py-1 text-xs font-medium rounded bg-zinc-800 dark:bg-zinc-200 text-zinc-200 dark:text-zinc-800 hover:opacity-80"
                      >
                        {selectedBeautificationIds.length === filteredBeautifications.length && filteredBeautifications.length > 0
                          ? '取消全选'
                          : '全选当前'}
                      </button>

                      <button
                        disabled={selectedBeautificationIds.length === 0}
                        onClick={() => setShowBeautificationBatchMoveModal(true)}
                        className="px-2.5 py-1 text-xs font-medium rounded bg-zinc-800 dark:bg-zinc-200 text-zinc-200 dark:text-zinc-800 hover:opacity-80 disabled:opacity-40 flex items-center gap-1"
                      >
                        <Move className="w-3 h-3" /> 批量移动
                      </button>

                      <button
                        disabled={selectedBeautificationIds.length === 0}
                        onClick={handleBatchDeleteBeautifications}
                        className="px-2.5 py-1 text-xs font-medium rounded bg-rose-600 text-white hover:bg-rose-700 disabled:opacity-40 flex items-center gap-1"
                      >
                        <Trash2 className="w-3 h-3" /> 批量删除
                      </button>
                    </div>
                  </div>
                )}

                {/* Gallery List of Beautifications ("图集形式") */}
                {filteredBeautifications.length === 0 ? (
                  <div className="text-center py-16 bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-xl space-y-3">
                    <p className="text-sm font-medium text-zinc-600 dark:text-zinc-400">
                      {beautificationSearchQuery || beautificationCategoryFilter !== '全部分类'
                        ? '未找到匹配的美化'
                        : '暂无美化'}
                    </p>
                    <p className="text-xs text-zinc-400">
                      点击右上角的 “+” 按钮，上传并导入美化文档或图片 (.docx, .txt, .json, .css, .png)
                    </p>
                  </div>
                ) : (
                  <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-5 xl:grid-cols-6 gap-4 sm:gap-6">
                    {sortItemList(
                      filteredBeautifications,
                      beautificationSortOrder,
                      (item) => item.name || '',
                      (item) => item.createdAt || 0
                    ).map((item) => {
                      const isSelected = selectedBeautificationIds.includes(item.id);

                      return (
                        <div
                          key={item.id}
                          onClick={() => {
                            if (beautificationBatchMode) {
                              if (isSelected) {
                                setSelectedBeautificationIds((p) => p.filter((id) => id !== item.id));
                              } else {
                                setSelectedBeautificationIds((p) => [...p, item.id]);
                              }
                            } else {
                              setEditingBeautification({ ...item });
                              setBeautificationDetailTab('preview');
                              setCodeSearchQuery('');
                            }
                          }}
                          className={`group bg-white dark:bg-zinc-900 border rounded-xl overflow-hidden cursor-pointer transition-all hover:shadow-lg flex flex-col relative ${
                            isSelected
                              ? 'border-rose-500 ring-2 ring-rose-500/30 dark:ring-rose-500/30'
                              : 'border-zinc-200 dark:border-zinc-800 hover:border-zinc-300 dark:hover:border-zinc-700'
                          }`}
                        >
                          {/* Card Cover (aspect-[2/3] for gallery format) */}
                          <div className="aspect-[2/3] w-full bg-zinc-100 dark:bg-zinc-800/80 relative overflow-hidden flex items-center justify-center p-2">
                            {item.coverImage ? (
                              <img
                                src={item.coverImage}
                                alt={item.name}
                                className="w-full h-full object-cover transition-transform duration-300 group-hover:scale-105"
                              />
                            ) : (
                              <div className="w-full h-full flex flex-col items-center justify-center text-center p-3 border border-dashed border-zinc-300 dark:border-zinc-700 rounded-lg bg-zinc-50/50 dark:bg-zinc-900/40 space-y-2">
                                <FileText className="w-8 h-8 text-zinc-400 dark:text-zinc-600" />
                                <span className="text-[11px] font-bold text-zinc-600 dark:text-zinc-400 line-clamp-2 leading-tight">
                                  {item.name}
                                </span>
                                <span className="text-[9px] uppercase font-mono text-zinc-400 dark:text-zinc-500">
                                  {item.fileType || '美化'}
                                </span>
                              </div>
                            )}

                            {/* Batch Selection Checkbox */}
                            {beautificationBatchMode && (
                              <div className="absolute top-2 left-2 z-10 bg-black/50 rounded-full p-0.5">
                                {isSelected ? (
                                  <CheckCircle2 className="w-5 h-5 text-rose-500 fill-white dark:fill-zinc-900" />
                                ) : (
                                  <Circle className="w-5 h-5 text-white/80" />
                                )}
                              </div>
                            )}

                            {/* Category or Type Badge */}
                            {(item.type || (item.category && item.category !== '默认')) && (
                              <span className="absolute top-2 right-2 px-1.5 py-0.5 text-[9px] font-bold rounded bg-black/60 text-white backdrop-blur-sm">
                                {item.type || item.category}
                              </span>
                            )}
                          </div>

                          {/* Theme Title & Author */}
                          <div className="p-3 space-y-1 bg-white dark:bg-zinc-900 border-t border-zinc-100 dark:border-zinc-800">
                            <h3 className="text-xs font-bold text-zinc-900 dark:text-zinc-100 truncate" title={item.name}>
                              {item.name}
                            </h3>
                            <div className="flex items-center justify-between text-[10px] text-zinc-400">
                              <span className="truncate">作者：{item.author || '默认'}</span>
                              <span className="uppercase text-[9px]">{item.fileType || '文档'}</span>
                            </div>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>
            ) : currentPage === 'st-presets' ? (
              <div className="max-w-7xl mx-auto space-y-4">
                {/* Search Bar & Category Group Bar */}
                <div className="bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-xl p-2.5 flex flex-col md:flex-row items-stretch md:items-center justify-between gap-3">
                  {/* Search Input */}
                  <div className="relative flex-1 min-w-[200px]">
                    <Search className="w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 text-zinc-400" />
                    <input
                      type="text"
                      value={presetSearchQuery}
                      onChange={(e) => setPresetSearchQuery(e.target.value)}
                      placeholder="搜索预设文件名、作者、分类或内容..."
                      className="w-full pl-9 pr-4 py-1.5 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400 dark:focus:ring-zinc-500"
                    />
                  </div>

                  {/* Category Group Scroll Tabs */}
                  <div className="flex items-center gap-1.5 overflow-x-auto pb-1 md:pb-0 scrollbar-none flex-1">
                    {/* "全部分组" Tab */}
                    <button
                      onClick={() => setPresetCategoryFilter('全部分组')}
                      className={`px-3 py-1.5 text-xs font-semibold rounded-lg whitespace-nowrap transition-all ${
                        presetCategoryFilter === '全部分组'
                          ? 'bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 shadow-sm'
                          : 'text-zinc-600 dark:text-zinc-400 hover:bg-zinc-100 dark:hover:bg-zinc-800'
                      }`}
                    >
                      全部分组 ({presetsList.length})
                    </button>

                    {/* Custom Categories */}
                    {Array.from(
                      new Set(['默认', ...(appData.presetCategories || []), ...presetsList.map((p) => p.category || '默认')])
                    ).map((catName) => {
                      const count = presetsList.filter((p) => (p.category || '默认') === catName).length;
                      const isSelected = presetCategoryFilter === catName;

                      return (
                        <div
                          key={catName}
                          className="relative group flex items-center"
                          onContextMenu={(e) => {
                            e.preventDefault();
                            setManagingPresetGroup(catName);
                            setRenamePresetGroupInput(catName);
                          }}
                        >
                          <button
                            onClick={() => setPresetCategoryFilter(catName)}
                            className={`px-3 py-1.5 text-xs font-medium rounded-lg whitespace-nowrap transition-all flex items-center gap-1.5 ${
                              isSelected
                                ? 'bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 shadow-sm'
                                : 'bg-zinc-100 dark:bg-zinc-800/60 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-200 dark:hover:bg-zinc-800'
                            }`}
                          >
                            <span>{catName}</span>
                            <span className="text-[10px] opacity-70">({count})</span>
                          </button>

                          {/* Management Icon */}
                          <button
                            onClick={(e) => {
                              e.stopPropagation();
                              setManagingPresetGroup(catName);
                              setRenamePresetGroupInput(catName);
                            }}
                            className="ml-0.5 p-1 text-zinc-400 hover:text-zinc-800 dark:hover:text-zinc-200 transition-opacity"
                            title="右键或长按/点击管理分组"
                          >
                            <MoreHorizontal className="w-3 h-3" />
                          </button>
                        </div>
                      );
                    })}

                    {/* "新建分组" Button directly on the right of categories */}
                    <button
                      onClick={() => setShowNewPresetGroupModal(true)}
                      className="flex items-center gap-1 px-2.5 py-1.5 text-xs font-medium rounded-lg bg-zinc-100 dark:bg-zinc-800 text-zinc-800 dark:text-zinc-200 hover:bg-zinc-200 dark:hover:bg-zinc-700 transition-colors whitespace-nowrap ml-auto flex-shrink-0"
                    >
                      <FolderPlus className="w-3.5 h-3.5" /> 新建分组
                    </button>
                  </div>

                  {/* Right Action: Sorting & "选择" (Batch Mode) */}
                  <div className="flex items-center justify-start gap-2 flex-shrink-0">
                    <div className="flex items-center gap-1.5 bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-lg px-2.5 h-[30px] flex-shrink-0">
                      <ArrowUpDown className="w-3.5 h-3.5 text-zinc-400" />
                      <select
                        value={presetSortOrder}
                        onChange={(e) => setPresetSortOrder(e.target.value as any)}
                        className="text-xs leading-none bg-transparent text-zinc-700 dark:text-zinc-300 focus:outline-none cursor-pointer"
                      >
                        <option value="default">默认排序</option>
                        <option value="az">名称 A-Z</option>
                        <option value="za">名称 Z-A</option>
                        <option value="newest">最新添加</option>
                        <option value="oldest">最早添加</option>
                      </select>
                    </div>

                    <button
                      onClick={() => {
                        setPresetBatchMode(!presetBatchMode);
                        setSelectedPresetIds([]);
                      }}
                      className={`px-3.5 py-1.5 text-xs font-medium rounded-lg border transition-colors flex items-center gap-1.5 whitespace-nowrap ${
                        presetBatchMode
                          ? 'bg-amber-100 text-amber-900 border-amber-300 dark:bg-amber-950 dark:text-amber-200 dark:border-amber-800'
                          : 'border-zinc-200 dark:border-zinc-800 hover:bg-zinc-100 dark:hover:bg-zinc-800 text-zinc-700 dark:text-zinc-300'
                      }`}
                    >
                      <CheckSquare className="w-3.5 h-3.5" /> {presetBatchMode ? '退出选择' : '选择'}
                    </button>
                  </div>
                </div>

                {/* Batch Action Toolbar when in Batch Mode */}
                {presetBatchMode && (
                  <div className="bg-zinc-900 dark:bg-zinc-100 text-white dark:text-zinc-900 rounded-xl p-3 flex flex-col sm:flex-row items-center justify-between gap-3 shadow-lg animate-in fade-in slide-in-from-top-2">
                    <div className="text-xs font-medium">
                      已选择 <span className="font-bold underline">{selectedPresetIds.length}</span> 个预设
                    </div>
                    <div className="flex items-center gap-2">
                      <button
                        onClick={() => {
                          if (selectedPresetIds.length === filteredPresets.length && filteredPresets.length > 0) {
                            setSelectedPresetIds([]);
                          } else {
                            setSelectedPresetIds(filteredPresets.map((item) => item.id));
                          }
                        }}
                        className="px-2.5 py-1 text-xs font-medium rounded bg-zinc-800 dark:bg-zinc-200 text-zinc-200 dark:text-zinc-800 hover:opacity-80"
                      >
                        {selectedPresetIds.length === filteredPresets.length && filteredPresets.length > 0
                          ? '取消全选'
                          : '全选当前'}
                      </button>

                      <button
                        disabled={selectedPresetIds.length === 0}
                        onClick={() => setShowPresetBatchMoveModal(true)}
                        className="px-2.5 py-1 text-xs font-medium rounded bg-zinc-800 dark:bg-zinc-200 text-zinc-200 dark:text-zinc-800 hover:opacity-80 disabled:opacity-40 flex items-center gap-1"
                      >
                        <Move className="w-3 h-3" /> 批量移动
                      </button>

                      <button
                        disabled={selectedPresetIds.length === 0}
                        onClick={handleBatchDeletePresets}
                        className="px-2.5 py-1 text-xs font-medium rounded bg-rose-600 text-white hover:bg-rose-700 disabled:opacity-40 flex items-center gap-1"
                      >
                        <Trash2 className="w-3 h-3" /> 批量删除
                      </button>
                    </div>
                  </div>
                )}

                {/* Preset List View */}
                {filteredPresets.length === 0 ? (
                  <div className="text-center py-16 bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-xl space-y-3">
                    <p className="text-sm font-medium text-zinc-600 dark:text-zinc-400">
                      {presetSearchQuery ? '未找到匹配的预设文件' : '暂无预设文件'}
                    </p>
                    <p className="text-xs text-zinc-400">
                      点击右上角的 “+” 按钮上传 JSON 格式的预设文件
                    </p>
                  </div>
                ) : (
                  <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3">
                    {sortItemList(
                      filteredPresets,
                      presetSortOrder,
                      (item) => item.fileName || item.name || '',
                      (item) => item.createdAt || 0
                    ).map((item) => {
                      const isSelected = selectedPresetIds.includes(item.id);

                      return (
                        <div
                          key={item.id}
                          onClick={() => {
                            if (presetBatchMode) {
                              if (isSelected) {
                                setSelectedPresetIds((p) => p.filter((id) => id !== item.id));
                              } else {
                                setSelectedPresetIds((p) => [...p, item.id]);
                              }
                            } else {
                              setEditingPreset({ ...item });
                              setEditingPresetTab('details');
                              setPresetEntrySearchQuery('');
                            }
                          }}
                          className={`p-3.5 bg-white dark:bg-zinc-900 border rounded-xl cursor-pointer transition-all hover:shadow-md flex items-center justify-between gap-3 ${
                            isSelected
                              ? 'border-rose-500 ring-2 ring-rose-500/30 dark:ring-rose-500/30'
                              : 'border-zinc-200 dark:border-zinc-800 hover:border-zinc-300 dark:hover:border-zinc-700'
                          }`}
                        >
                          <div className="flex items-center gap-3 min-w-0 flex-1">
                            <div className="w-10 h-10 rounded-lg bg-zinc-100 dark:bg-zinc-800 flex items-center justify-center flex-shrink-0 text-zinc-500 dark:text-zinc-400">
                              <Sliders className="w-5 h-5" />
                            </div>

                            <div className="min-w-0 flex-1 space-y-1">
                              <h3 className="text-xs font-bold text-zinc-900 dark:text-zinc-100 truncate" title={item.name || item.fileName}>
                                {item.name || item.fileName || '未命名预设'}
                              </h3>
                              <div className="flex items-center gap-2 text-[10px] text-zinc-500 dark:text-zinc-400 truncate">
                                <span className="px-1.5 py-0.2 rounded bg-zinc-100 dark:bg-zinc-800 text-zinc-600 dark:text-zinc-300">
                                  {item.category || '默认'}
                                </span>
                                {item.author && <span className="truncate">作者: {item.author}</span>}
                              </div>
                            </div>
                          </div>

                          {presetBatchMode ? (
                            <div className="flex-shrink-0">
                              {isSelected ? (
                                <CheckCircle2 className="w-5 h-5 text-rose-500 fill-white dark:fill-zinc-900" />
                              ) : (
                                <Circle className="w-5 h-5 text-zinc-300 dark:text-zinc-600" />
                              )}
                            </div>
                          ) : (
                            <span className="text-[10px] text-zinc-400 bg-zinc-50 dark:bg-zinc-800/60 px-2 py-1 rounded-md border border-zinc-100 dark:border-zinc-800">
                              查看/编辑
                            </span>
                          )}
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>
            ) : currentPage === 'st-plugins' ? (
              <div className="max-w-7xl mx-auto space-y-4">
                {/* Top Search Bar & Category Group Bar */}
                <div className="bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-xl p-2.5 flex flex-col md:flex-row items-stretch md:items-center justify-between gap-3">
                  {/* Search Input */}
                  <div className="relative flex-1 min-w-[180px] max-w-md">
                    <Search className="w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 text-zinc-400" />
                    <input
                      type="text"
                      value={pluginSearchQuery}
                      onChange={(e) => setPluginSearchQuery(e.target.value)}
                      placeholder="搜索插件/脚本名称、作者、来源或分类..."
                      className="w-full pl-9 pr-4 py-1.5 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400 dark:focus:ring-zinc-500"
                    />
                  </div>

                  {/* Category Group Scroll Tabs */}
                  <div className="flex items-center gap-1.5 overflow-x-auto pb-1 md:pb-0 scrollbar-none flex-1">
                    {/* "全部分组" Tab */}
                    <button
                      onClick={() => setPluginCategoryFilter('全部分组')}
                      className={`px-3 py-1.5 text-xs font-semibold rounded-lg whitespace-nowrap transition-all ${
                        pluginCategoryFilter === '全部分组'
                          ? 'bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 shadow-sm'
                          : 'text-zinc-600 dark:text-zinc-400 hover:bg-zinc-100 dark:hover:bg-zinc-800'
                      }`}
                    >
                      全部分组 ({pluginsList.length})
                    </button>

                    {/* Custom Categories */}
                    {Array.from(
                      new Set(['默认', ...(appData.pluginCategories || []), ...pluginsList.map((p) => p.category || '默认')])
                    ).map((catName) => {
                      const count = pluginsList.filter((p) => (p.category || '默认') === catName).length;
                      const isSelected = pluginCategoryFilter === catName;

                      return (
                        <div
                          key={catName}
                          className="relative group flex items-center"
                          onContextMenu={(e) => {
                            e.preventDefault();
                            setManagingPluginCategory(catName);
                            setRenamePluginCategoryInput(catName);
                          }}
                        >
                          <button
                            onClick={() => setPluginCategoryFilter(catName)}
                            className={`px-3 py-1.5 text-xs font-medium rounded-lg whitespace-nowrap transition-all flex items-center gap-1.5 ${
                              isSelected
                                ? 'bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 shadow-sm'
                                : 'bg-zinc-100 dark:bg-zinc-800/60 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-200 dark:hover:bg-zinc-800'
                            }`}
                          >
                            <span>{catName}</span>
                            <span className="text-[10px] opacity-70">({count})</span>
                          </button>

                          {/* Management Icon */}
                          <button
                            onClick={(e) => {
                              e.stopPropagation();
                              setManagingPluginCategory(catName);
                              setRenamePluginCategoryInput(catName);
                            }}
                            className="ml-0.5 p-1 text-zinc-400 hover:text-zinc-800 dark:hover:text-zinc-200 transition-opacity"
                            title="右键或点击管理分组"
                          >
                            <MoreHorizontal className="w-3 h-3" />
                          </button>
                        </div>
                      );
                    })}

                    {/* "新建分组" Button directly on the right of categories */}
                    <button
                      onClick={() => setShowNewPluginGroupModal(true)}
                      className="flex items-center gap-1 px-2.5 py-1.5 text-xs font-medium rounded-lg bg-zinc-100 dark:bg-zinc-800 text-zinc-800 dark:text-zinc-200 hover:bg-zinc-200 dark:hover:bg-zinc-700 transition-colors whitespace-nowrap ml-auto flex-shrink-0"
                    >
                      <FolderPlus className="w-3.5 h-3.5" /> 新建分组
                    </button>
                  </div>

                  {/* Right Action: Sorting & "选择" (Batch Mode) */}
                  <div className="flex items-center justify-start gap-2 flex-shrink-0">
                    <div className="flex items-center gap-1.5 bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-lg px-2.5 h-[30px] flex-shrink-0">
                      <ArrowUpDown className="w-3.5 h-3.5 text-zinc-400" />
                      <select
                        value={pluginSortOrder}
                        onChange={(e) => setPluginSortOrder(e.target.value as any)}
                        className="text-xs leading-none bg-transparent text-zinc-700 dark:text-zinc-300 focus:outline-none cursor-pointer"
                      >
                        <option value="default">默认排序</option>
                        <option value="az">名称 A-Z</option>
                        <option value="za">名称 Z-A</option>
                        <option value="newest">最新添加</option>
                        <option value="oldest">最早添加</option>
                      </select>
                    </div>

                    <button
                      onClick={() => {
                        setPluginBatchMode(!pluginBatchMode);
                        setSelectedPluginIds([]);
                      }}
                      className={`px-3.5 py-1.5 text-xs font-medium rounded-lg border transition-colors flex items-center gap-1.5 whitespace-nowrap ${
                        pluginBatchMode
                          ? 'bg-amber-100 text-amber-900 border-amber-300 dark:bg-amber-950 dark:text-amber-200 dark:border-amber-800'
                          : 'border-zinc-200 dark:border-zinc-800 hover:bg-zinc-100 dark:hover:bg-zinc-800 text-zinc-700 dark:text-zinc-300'
                      }`}
                    >
                      <CheckSquare className="w-3.5 h-3.5" /> {pluginBatchMode ? '退出选择' : '选择'}
                    </button>
                  </div>
                </div>

                {/* Batch Action Toolbar when in Batch Mode */}
                {pluginBatchMode && (
                  <div className="bg-zinc-900 dark:bg-zinc-100 text-white dark:text-zinc-900 rounded-xl p-3 flex flex-col sm:flex-row items-center justify-between gap-3 shadow-lg animate-in fade-in slide-in-from-top-2">
                    <div className="text-xs font-medium">
                      已选择 <span className="font-bold underline">{selectedPluginIds.length}</span> 个插件/脚本
                    </div>
                    <div className="flex items-center gap-2">
                      <button
                        onClick={() => {
                          if (selectedPluginIds.length === filteredPlugins.length && filteredPlugins.length > 0) {
                            setSelectedPluginIds([]);
                          } else {
                            setSelectedPluginIds(filteredPlugins.map((item) => item.id));
                          }
                        }}
                        className="px-2.5 py-1 text-xs font-medium rounded bg-zinc-800 dark:bg-zinc-200 text-zinc-200 dark:text-zinc-800 hover:opacity-80"
                      >
                        {selectedPluginIds.length === filteredPlugins.length && filteredPlugins.length > 0
                          ? '取消全选'
                          : '全选当前'}
                      </button>

                      <button
                        disabled={selectedPluginIds.length === 0}
                        onClick={() => setShowPluginBatchMoveModal(true)}
                        className="px-2.5 py-1 text-xs font-medium rounded bg-zinc-800 dark:bg-zinc-200 text-zinc-200 dark:text-zinc-800 hover:opacity-80 disabled:opacity-40 flex items-center gap-1"
                      >
                        <Move className="w-3 h-3" /> 批量移动
                      </button>

                      <button
                        disabled={selectedPluginIds.length === 0}
                        onClick={handleBatchDeletePlugins}
                        className="px-2.5 py-1 text-xs font-medium rounded bg-rose-600 text-white hover:bg-rose-700 disabled:opacity-40 flex items-center gap-1"
                      >
                        <Trash2 className="w-3 h-3" /> 批量删除
                      </button>
                    </div>
                  </div>
                )}

                {/* Items List */}
                {filteredPlugins.length === 0 ? (
                  <div className="text-center py-20 bg-white dark:bg-zinc-900 rounded-2xl border border-dashed border-zinc-200 dark:border-zinc-800 p-8 space-y-3">
                    <p className="text-xs text-zinc-400">暂无插件或脚本数据</p>
                    <button
                      onClick={() => setShowAddPluginTypeModal(true)}
                      className="px-4 py-2 bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 rounded-xl text-xs font-bold hover:opacity-90 transition-opacity inline-flex items-center gap-1.5 shadow-sm"
                    >
                      <Plus className="w-4 h-4" /> 点击添加插件或导入脚本
                    </button>
                  </div>
                ) : (
                  <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3">
                    {sortItemList(
                      filteredPlugins,
                      pluginSortOrder,
                      (item) => item.name || '',
                      (item) => item.createdAt || 0
                    ).map((item) => {
                      const isSelected = selectedPluginIds.includes(item.id);

                      return (
                        <div
                          key={item.id}
                          onClick={() => {
                            if (pluginBatchMode) {
                              if (isSelected) {
                                setSelectedPluginIds(selectedPluginIds.filter((id) => id !== item.id));
                              } else {
                                setSelectedPluginIds([...selectedPluginIds, item.id]);
                              }
                            } else {
                              handleOpenPluginDetail(item);
                            }
                          }}
                          className={`p-3.5 bg-white dark:bg-zinc-900 border rounded-xl transition-all cursor-pointer flex items-center justify-between gap-3 shadow-sm hover:shadow-md ${
                            isSelected
                              ? 'border-rose-500 dark:border-rose-500 bg-rose-50/20 dark:bg-rose-950/20'
                              : 'border-zinc-200 dark:border-zinc-800 hover:border-zinc-300 dark:hover:border-zinc-700'
                          }`}
                        >
                          <div className="min-w-0 flex-1 space-y-1.5">
                            <div className="flex items-center gap-2">
                              {item.type === 'script' ? (
                                <span className="px-1.5 py-0.5 rounded text-[10px] font-bold bg-emerald-100 text-emerald-800 dark:bg-emerald-950/80 dark:text-emerald-300">
                                  脚本
                                </span>
                              ) : (
                                <span className="px-1.5 py-0.5 rounded text-[10px] font-bold bg-indigo-100 text-indigo-800 dark:bg-indigo-950/80 dark:text-indigo-300">
                                  插件
                                </span>
                              )}
                              <h3 className="text-xs font-bold text-zinc-900 dark:text-zinc-100 truncate" title={item.name}>
                                {item.name}
                              </h3>
                            </div>

                            <div className="flex items-center gap-2 text-[10px] text-zinc-500 dark:text-zinc-400 truncate">
                              <span className="px-1.5 py-0.2 rounded bg-zinc-100 dark:bg-zinc-800 text-zinc-600 dark:text-zinc-300">
                                {item.category || '默认'}
                              </span>
                              {item.type === 'plugin' ? (
                                <span className="truncate text-zinc-400 font-mono">{item.url}</span>
                              ) : (
                                <span className="truncate">作者: {item.author || '未知作者'}</span>
                              )}
                            </div>
                          </div>

                          {pluginBatchMode ? (
                            <div className="flex-shrink-0">
                              {isSelected ? (
                                <CheckCircle2 className="w-5 h-5 text-rose-500 fill-white dark:fill-zinc-900" />
                              ) : (
                                <Circle className="w-5 h-5 text-zinc-300 dark:text-zinc-600" />
                              )}
                            </div>
                          ) : (
                            <span className="text-[10px] text-zinc-400 bg-zinc-50 dark:bg-zinc-800/60 px-2 py-1 rounded-md border border-zinc-100 dark:border-zinc-800 whitespace-nowrap">
                              查看/编辑
                            </span>
                          )}
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>
            ) : currentPage === 'normal-cards' ? (
              <div className="max-w-7xl mx-auto space-y-4">
                {/* Search & Categories Bar */}
                <div className="bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl p-3 flex flex-col md:flex-row items-stretch md:items-center justify-between gap-3 shadow-sm">
                  {/* Search Box */}
                  <div className="relative w-full md:w-72">
                    <Search className="w-3.5 h-3.5 absolute left-3 top-1/2 -translate-y-1/2 text-zinc-400" />
                    <input
                      type="text"
                      value={normalCardSearchQuery}
                      onChange={(e) => setNormalCardSearchQuery(e.target.value)}
                      placeholder="搜索角色卡、真名、作者..."
                      className="w-full pl-9 pr-4 py-1.5 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400 dark:focus:ring-zinc-500"
                    />
                  </div>

                  {/* Categories & Actions */}
                  <div className="flex flex-wrap items-center gap-2 flex-1 min-w-0 justify-between md:justify-end">
                    <div className="flex items-center gap-1.5 overflow-x-auto scrollbar-none flex-1 min-w-0">
                      {/* 全部分组 Tab */}
                      <button
                        onClick={() => setNormalCardCategoryFilter('全部分组')}
                        className={`px-3 py-1.5 text-xs font-semibold rounded-xl whitespace-nowrap transition-all ${
                          normalCardCategoryFilter === '全部分组'
                            ? 'bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 shadow-sm'
                            : 'text-zinc-600 dark:text-zinc-400 hover:bg-zinc-100 dark:hover:bg-zinc-800'
                        }`}
                      >
                        全部分组 ({normalCardsList.length})
                      </button>

                      {/* Custom Categories */}
                      {Array.from(
                        new Set(['默认', ...(appData.normalCardCategories || []), ...normalCardsList.map((p) => p.category || '默认')])
                      ).map((cat) => {
                        const count = normalCardsList.filter((p) => (p.category || '默认') === cat).length;
                        const isSelected = normalCardCategoryFilter === cat;

                        return (
                          <div
                            key={cat}
                            className="relative group flex items-center"
                            onContextMenu={(e) => {
                              e.preventDefault();
                              setManagingNormalCardCategory(cat);
                              setRenameNormalCardCategoryInput(cat);
                            }}
                          >
                            <button
                              onClick={() => setNormalCardCategoryFilter(cat)}
                              className={`px-3 py-1.5 text-xs font-semibold rounded-xl whitespace-nowrap transition-all flex items-center gap-1.5 ${
                                isSelected
                                  ? 'bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 shadow-sm'
                                  : 'text-zinc-600 dark:text-zinc-400 hover:bg-zinc-100 dark:hover:bg-zinc-800'
                              }`}
                            >
                              <span>{cat}</span>
                              <span
                                className={`text-[10px] px-1.5 py-0.2 rounded-full ${
                                  isSelected
                                    ? 'bg-zinc-700 text-zinc-100 dark:bg-zinc-300 dark:text-zinc-900'
                                    : 'bg-zinc-100 dark:bg-zinc-800 text-zinc-500'
                                }`}
                              >
                                {count}
                              </span>
                            </button>

                            <button
                              onClick={(e) => {
                                e.stopPropagation();
                                setManagingNormalCardCategory(cat);
                                setRenameNormalCardCategoryInput(cat);
                              }}
                              className="p-1 text-zinc-400 hover:text-zinc-700 dark:hover:text-zinc-200"
                              title="管理分组"
                            >
                              <MoreHorizontal className="w-3.5 h-3.5" />
                            </button>
                          </div>
                        );
                      })}

                      {/* 新建分组 Button (Right next to categories) */}
                      <button
                        onClick={() => setShowNewNormalCardGroupModal(true)}
                        className="px-2.5 py-1.5 text-xs font-semibold rounded-xl border border-dashed border-zinc-300 dark:border-zinc-700 text-zinc-600 dark:text-zinc-400 hover:bg-zinc-100 dark:hover:bg-zinc-800 flex items-center gap-1 whitespace-nowrap ml-auto"
                      >
                        <Plus className="w-3.5 h-3.5" /> 新建分组
                      </button>
                    </div>

                    {/* Sorting & 选择 (Batch Mode Toggle) Button */}
                    <div className="flex items-center justify-start gap-2 flex-shrink-0 basis-full">
                      <div className="flex items-center gap-1.5 bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-lg px-2.5 h-[30px] flex-shrink-0">
                        <ArrowUpDown className="w-3.5 h-3.5 text-zinc-400" />
                        <select
                          value={normalCardSortOrder}
                          onChange={(e) => setNormalCardSortOrder(e.target.value as any)}
                          className="text-xs leading-none bg-transparent text-zinc-700 dark:text-zinc-300 focus:outline-none cursor-pointer"
                        >
                          <option value="default">默认排序</option>
                          <option value="az">名称 A-Z</option>
                          <option value="za">名称 Z-A</option>
                          <option value="newest">最新添加</option>
                          <option value="oldest">最早添加</option>
                        </select>
                      </div>

                      <button
                        onClick={() => {
                          setNormalCardBatchMode(!normalCardBatchMode);
                          setSelectedNormalCardIds([]);
                        }}
                        className={`flex items-center gap-1 px-3 py-1.5 text-xs font-semibold rounded-xl transition-all whitespace-nowrap flex-shrink-0 ${
                          normalCardBatchMode
                            ? 'bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 shadow-sm'
                            : 'border border-zinc-200 dark:border-zinc-800 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800'
                        }`}
                      >
                        <CheckSquare className="w-3.5 h-3.5" /> {normalCardBatchMode ? '退出选择' : '选择'}
                      </button>
                    </div>
                  </div>
                </div>

                {/* Batch Mode Toolbar */}
                {normalCardBatchMode && (
                  <div className="bg-zinc-900 dark:bg-zinc-100 text-white dark:text-zinc-900 rounded-2xl p-3 flex flex-col sm:flex-row items-center justify-between gap-3 shadow-lg animate-in fade-in">
                    <div className="text-xs font-medium">
                      已选择 <span className="font-bold underline">{selectedNormalCardIds.length}</span> 个角色卡
                    </div>
                    <div className="flex items-center gap-2">
                      <button
                        onClick={() => {
                          if (selectedNormalCardIds.length === filteredNormalCards.length) setSelectedNormalCardIds([]);
                          else setSelectedNormalCardIds(filteredNormalCards.map((p) => p.id));
                        }}
                        className="px-3 py-1.5 text-xs font-medium rounded-lg bg-zinc-800 dark:bg-zinc-200 text-zinc-200 dark:text-zinc-800 hover:opacity-90"
                      >
                        {selectedNormalCardIds.length === filteredNormalCards.length ? '取消全选' : '全选当前'}
                      </button>
                      <button
                        disabled={selectedNormalCardIds.length === 0}
                        onClick={() => setShowNormalCardBatchMoveModal(true)}
                        className="px-3 py-1.5 text-xs font-medium rounded-lg bg-zinc-800 dark:bg-zinc-200 text-zinc-200 dark:text-zinc-800 hover:opacity-90 disabled:opacity-40"
                      >
                        批量移动
                      </button>
                      <button
                        disabled={selectedNormalCardIds.length === 0}
                        onClick={handleBatchDeleteNormalCards}
                        className="px-3 py-1.5 text-xs font-bold rounded-lg bg-rose-600 text-white hover:bg-rose-700 disabled:opacity-40"
                      >
                        批量删除
                      </button>
                    </div>
                  </div>
                )}

                {/* Cards List Display */}
                {filteredNormalCards.length === 0 ? (
                  <div className="text-center py-20 bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl p-8">
                    <FileText className="w-10 h-10 text-zinc-300 dark:text-zinc-700 mx-auto mb-3" />
                    <p className="text-xs font-medium text-zinc-500 dark:text-zinc-400">
                      {normalCardSearchQuery ? '未找到匹配的角色卡' : '暂无普通角色卡，点击顶部“+”上传 .docx / .txt / .zip 文件'}
                    </p>
                  </div>
                ) : (
                  <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3">
                    {sortItemList(
                      filteredNormalCards,
                      normalCardSortOrder,
                      (item) => item.charName || item.fileName || '',
                      (item) => item.createdAt || 0
                    ).map((item) => {
                      const isSelected = selectedNormalCardIds.includes(item.id);

                      return (
                        <div
                          key={item.id}
                          onClick={() => {
                            if (normalCardBatchMode) {
                              if (isSelected) {
                                setSelectedNormalCardIds(selectedNormalCardIds.filter((id) => id !== item.id));
                              } else {
                                setSelectedNormalCardIds([...selectedNormalCardIds, item.id]);
                              }
                            } else {
                              setEditingNormalCard(item);
                            }
                          }}
                          className={`bg-white dark:bg-zinc-900 border rounded-2xl p-3.5 transition-all cursor-pointer flex items-center justify-between gap-3 shadow-sm hover:shadow-md ${
                            isSelected
                              ? 'border-zinc-900 dark:border-zinc-100 ring-2 ring-zinc-900/10 dark:ring-zinc-100/10'
                              : 'border-zinc-200 dark:border-zinc-800 hover:border-zinc-400 dark:hover:border-zinc-600'
                          }`}
                        >
                          {/* Thumbnail / File Icon */}
                          {item.coverImage ? (
                            <img
                              src={item.coverImage}
                              alt={item.fileName}
                              className="w-12 h-14 rounded-xl object-cover border border-zinc-200 dark:border-zinc-700 flex-shrink-0"
                            />
                          ) : (
                            <div className="w-12 h-14 rounded-xl bg-zinc-100 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 flex flex-col items-center justify-center text-zinc-400 flex-shrink-0">
                              <FileText className="w-6 h-6 text-zinc-400" />
                            </div>
                          )}

                          {/* Info */}
                          <div className="flex-1 min-w-0">
                            <h3 className="text-xs font-bold text-zinc-900 dark:text-zinc-100 truncate" title={item.fileName}>
                              {item.fileName}
                            </h3>
                            <div className="text-[11px] text-zinc-500 dark:text-zinc-400 truncate mt-0.5">
                              真名: {item.charName || '未知'} | 作者: {item.author || '未知作者'}
                            </div>
                            <div className="flex items-center gap-1.5 mt-1.5">
                              <span className="px-1.5 py-0.2 rounded text-[10px] bg-zinc-100 dark:bg-zinc-800 text-zinc-600 dark:text-zinc-300 font-medium">
                                {item.category || '默认'}
                              </span>
                              {item.source && (
                                <span className="px-1.5 py-0.2 rounded text-[10px] bg-blue-50 text-blue-700 dark:bg-blue-950/60 dark:text-blue-300 truncate max-w-[120px]">
                                  {item.source}
                                </span>
                              )}
                            </div>
                          </div>

                          {/* Right Action / Batch Toggle */}
                          {normalCardBatchMode ? (
                            <div className="flex-shrink-0">
                              {isSelected ? (
                                <CheckCircle2 className="w-5 h-5 text-rose-500 fill-white dark:fill-zinc-900" />
                              ) : (
                                <Circle className="w-5 h-5 text-zinc-300 dark:text-zinc-600" />
                              )}
                            </div>
                          ) : (
                            <span className="text-[10px] text-zinc-400 bg-zinc-50 dark:bg-zinc-800/60 px-2 py-1 rounded-md border border-zinc-100 dark:border-zinc-800 whitespace-nowrap">
                              查看/编辑
                            </span>
                          )}
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>
            ) : currentPage === 'api-storage' ? (
              <div className="max-w-7xl mx-auto space-y-4">
                {/* Search Bar & Category Toolbar */}
                <div className="bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-xl p-2.5 flex flex-col md:flex-row items-stretch md:items-center justify-between gap-3">
                  {/* Search Input */}
                  <div className="relative w-full md:w-72">
                    <Search className="w-3.5 h-3.5 absolute left-3 top-1/2 -translate-y-1/2 text-zinc-400" />
                    <input
                      type="text"
                      value={apiSearchQuery}
                      onChange={(e) => setApiSearchQuery(e.target.value)}
                      placeholder="搜索 API 名称、地址、Key、描述..."
                      className="w-full pl-9 pr-4 py-1.5 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400 dark:focus:ring-zinc-500"
                    />
                  </div>

                  {/* Categories & Actions */}
                  <div className="flex flex-wrap items-center gap-2 flex-1 min-w-0 justify-between md:justify-end">
                    <div className="flex items-center gap-1.5 overflow-x-auto scrollbar-none flex-1 min-w-0">
                      {/* 全部分组 Tab */}
                      <button
                        onClick={() => setApiCategoryFilter('全部分组')}
                        className={`px-3 py-1.5 text-xs font-semibold rounded-xl whitespace-nowrap transition-all ${
                          apiCategoryFilter === '全部分组'
                            ? 'bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 shadow-sm'
                            : 'text-zinc-600 dark:text-zinc-400 hover:bg-zinc-100 dark:hover:bg-zinc-800'
                        }`}
                      >
                        全部分组 ({apisList.length})
                      </button>

                      {/* Custom Categories */}
                      {Array.from(
                        new Set(['默认', ...(appData.apiCategories || []), ...apisList.map((p) => p.category || '默认')])
                      ).map((cat) => {
                        const count = apisList.filter((p) => (p.category || '默认') === cat).length;
                        const isSelected = apiCategoryFilter === cat;

                        return (
                          <div
                            key={cat}
                            className="relative group flex items-center"
                            onContextMenu={(e) => {
                              e.preventDefault();
                              setManagingApiCategory(cat);
                              setRenameApiCategoryInput(cat);
                            }}
                          >
                            <button
                              onClick={() => setApiCategoryFilter(cat)}
                              className={`px-3 py-1.5 text-xs font-semibold rounded-xl whitespace-nowrap transition-all flex items-center gap-1.5 ${
                                isSelected
                                  ? 'bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 shadow-sm'
                                  : 'text-zinc-600 dark:text-zinc-400 hover:bg-zinc-100 dark:hover:bg-zinc-800'
                              }`}
                            >
                              <span>{cat}</span>
                              <span
                                className={`text-[10px] px-1.5 py-0.2 rounded-full ${
                                  isSelected
                                    ? 'bg-zinc-700 text-zinc-100 dark:bg-zinc-300 dark:text-zinc-900'
                                    : 'bg-zinc-100 dark:bg-zinc-800 text-zinc-500'
                                }`}
                              >
                                {count}
                              </span>
                            </button>

                            <button
                              onClick={(e) => {
                                e.stopPropagation();
                                setManagingApiCategory(cat);
                                setRenameApiCategoryInput(cat);
                              }}
                              className="p-1 text-zinc-400 hover:text-zinc-700 dark:hover:text-zinc-200"
                              title="右键或点击管理分组"
                            >
                              <MoreHorizontal className="w-3.5 h-3.5" />
                            </button>
                          </div>
                        );
                      })}

                      {/* 新建分组 Button (Right next to categories) */}
                      <button
                        onClick={() => setShowNewApiGroupModal(true)}
                        className="px-2.5 py-1.5 text-xs font-semibold rounded-xl border border-dashed border-zinc-300 dark:border-zinc-700 text-zinc-600 dark:text-zinc-400 hover:bg-zinc-100 dark:hover:bg-zinc-800 flex items-center gap-1 whitespace-nowrap flex-shrink-0 ml-auto"
                      >
                        <Plus className="w-3.5 h-3.5" /> 新建分组
                      </button>
                    </div>

                    {/* Sorting & 选择 (Batch Mode Toggle) Button */}
                    <div className="flex items-center justify-start gap-2 flex-shrink-0 basis-full">
                      <div className="flex items-center gap-1.5 bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-lg px-2.5 h-[30px] flex-shrink-0">
                        <ArrowUpDown className="w-3.5 h-3.5 text-zinc-400" />
                        <select
                          value={apiSortOrder}
                          onChange={(e) => setApiSortOrder(e.target.value as any)}
                          className="text-xs leading-none bg-transparent text-zinc-700 dark:text-zinc-300 focus:outline-none cursor-pointer"
                        >
                          <option value="default">默认排序</option>
                          <option value="az">名称 A-Z</option>
                          <option value="za">名称 Z-A</option>
                          <option value="newest">最新添加</option>
                          <option value="oldest">最早添加</option>
                        </select>
                      </div>

                      <button
                        onClick={() => {
                          setApiBatchMode(!apiBatchMode);
                          setSelectedApiIds([]);
                        }}
                        className={`flex items-center gap-1 px-3 py-1.5 text-xs font-semibold rounded-xl transition-all whitespace-nowrap flex-shrink-0 ${
                          apiBatchMode
                            ? 'bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 shadow-sm'
                            : 'border border-zinc-200 dark:border-zinc-800 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800'
                        }`}
                      >
                        <CheckSquare className="w-3.5 h-3.5" /> {apiBatchMode ? '退出选择' : '选择'}
                      </button>
                    </div>
                  </div>
                </div>

                {/* Batch Mode Toolbar */}
                {apiBatchMode && (
                  <div className="bg-zinc-900 dark:bg-zinc-100 text-white dark:text-zinc-900 rounded-2xl p-3 flex flex-col sm:flex-row items-center justify-between gap-3 shadow-lg animate-in fade-in">
                    <div className="text-xs font-medium">
                      已选择 <span className="font-bold underline">{selectedApiIds.length}</span> 个 API
                    </div>
                    <div className="flex items-center gap-2">
                      <button
                        onClick={() => {
                          if (selectedApiIds.length === filteredApis.length) setSelectedApiIds([]);
                          else setSelectedApiIds(filteredApis.map((p) => p.id));
                        }}
                        className="px-3 py-1.5 text-xs font-medium rounded-lg bg-zinc-800 dark:bg-zinc-200 text-zinc-200 dark:text-zinc-800 hover:opacity-90"
                      >
                        {selectedApiIds.length === filteredApis.length ? '取消全选' : '全选当前'}
                      </button>
                      <button
                        disabled={selectedApiIds.length === 0}
                        onClick={() => setShowApiBatchMoveModal(true)}
                        className="px-3 py-1.5 text-xs font-medium rounded-lg bg-zinc-800 dark:bg-zinc-200 text-zinc-200 dark:text-zinc-800 hover:opacity-90 disabled:opacity-40"
                      >
                        批量移动
                      </button>
                      <button
                        disabled={selectedApiIds.length === 0}
                        onClick={handleBatchDeleteApis}
                        className="px-3 py-1.5 text-xs font-bold rounded-lg bg-rose-600 text-white hover:bg-rose-700 disabled:opacity-40"
                      >
                        批量删除
                      </button>
                    </div>
                  </div>
                )}

                {/* API List */}
                {filteredApis.length === 0 ? (
                  <div className="text-center py-20 bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl p-8 space-y-3">
                    <FileCode className="w-10 h-10 text-zinc-300 dark:text-zinc-700 mx-auto mb-2" />
                    <p className="text-xs font-medium text-zinc-500 dark:text-zinc-400">
                      {apiSearchQuery ? '未找到匹配的 API' : '暂无 API 数据，点击右上角 “+” 按钮添加'}
                    </p>
                    {!apiSearchQuery && (
                      <button
                        onClick={handleOpenAddApiModal}
                        className="px-4 py-2 bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 rounded-xl text-xs font-bold hover:opacity-90 inline-flex items-center gap-1.5 shadow-sm"
                      >
                        <Plus className="w-4 h-4" /> 添加 API
                      </button>
                    )}
                  </div>
                ) : (
                  <div className="space-y-3">
                    {sortItemList(
                      filteredApis,
                      apiSortOrder,
                      (item) => item.name || '',
                      (item) => item.createdAt || 0
                    ).map((item) => {
                      const isSelected = selectedApiIds.includes(item.id);

                      return (
                        <div
                          key={item.id}
                          onClick={() => {
                            if (apiBatchMode) {
                              if (isSelected) {
                                setSelectedApiIds((p) => p.filter((id) => id !== item.id));
                              } else {
                                setSelectedApiIds((p) => [...p, item.id]);
                              }
                            } else {
                              setEditingApi(JSON.parse(JSON.stringify(item)));
                            }
                          }}
                          className={`p-4 bg-white dark:bg-zinc-900 border rounded-xl cursor-pointer transition-all hover:shadow-md flex flex-col md:flex-row md:items-center justify-between gap-3 ${
                            isSelected
                              ? 'border-rose-500 ring-2 ring-rose-500/30 dark:ring-rose-500/30'
                              : 'border-zinc-200 dark:border-zinc-800 hover:border-zinc-300 dark:hover:border-zinc-700'
                          }`}
                        >
                          <div className="flex-1 min-w-0 space-y-1.5">
                            <div className="flex items-center gap-2">
                              <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100 truncate" title={item.name}>
                                {item.name}
                              </h3>
                              <span className="px-2 py-0.5 rounded text-[10px] font-semibold bg-zinc-100 dark:bg-zinc-800 text-zinc-600 dark:text-zinc-300">
                                {item.category || '默认'}
                              </span>
                            </div>

                            <p className="text-xs text-zinc-500 dark:text-zinc-400 truncate font-mono" title={item.url}>
                              {item.url}
                            </p>

                            {/* Display all listed Keys */}
                            {item.keys && item.keys.length > 0 && (
                              <div className="flex flex-wrap gap-1.5 pt-1">
                                {item.keys.map((k, idx) => (
                                  <div
                                    key={k.id || idx}
                                    className="inline-flex items-center gap-1.5 px-2 py-0.5 rounded bg-zinc-50 dark:bg-zinc-800/80 border border-zinc-200/80 dark:border-zinc-700/80 text-[11px]"
                                  >
                                    {k.memo && (
                                      <span className="font-semibold text-zinc-600 dark:text-zinc-300 bg-zinc-200/60 dark:bg-zinc-700/60 px-1 rounded text-[10px]">
                                        {k.memo}
                                      </span>
                                    )}
                                    <span className="font-mono text-zinc-800 dark:text-zinc-200">
                                      {k.key ? (k.key.length > 18 ? `${k.key.substring(0, 8)}...${k.key.substring(k.key.length - 6)}` : k.key) : '（无Key）'}
                                    </span>
                                  </div>
                                ))}
                              </div>
                            )}

                            {item.description && (
                              <p className="text-xs text-zinc-400 dark:text-zinc-500 line-clamp-1">
                                {item.description}
                              </p>
                            )}
                          </div>

                          {apiBatchMode ? (
                            <div className="flex-shrink-0 self-end md:self-center">
                              {isSelected ? (
                                <CheckCircle2 className="w-5 h-5 text-rose-500 fill-white dark:fill-zinc-900" />
                              ) : (
                                <Circle className="w-5 h-5 text-zinc-300 dark:text-zinc-600" />
                              )}
                            </div>
                          ) : (
                            <div className="flex-shrink-0 text-xs text-zinc-400 dark:text-zinc-500 flex items-center gap-1 self-end md:self-center">
                              <span className="px-2 py-1 bg-zinc-50 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700/80 rounded-md">查看/编辑</span>
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>
            ) : currentPage === 'fonts' ? (
              <div className="max-w-7xl mx-auto space-y-4">
                {/* Hidden File Input for Font Upload */}
                <input
                  type="file"
                  multiple
                  ref={fontFileInputRef}
                  onChange={handleFileUploadFont}
                  accept=".ttf,.otf,.woff,.woff2"
                  className="hidden"
                />

                {/* Search Bar & Categories Toolbar */}
                <div className="bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-xl p-2.5 flex flex-col md:flex-row items-stretch md:items-center justify-between gap-3">
                  {/* Search Input */}
                  <div className="relative w-full md:w-72">
                    <Search className="w-3.5 h-3.5 absolute left-3 top-1/2 -translate-y-1/2 text-zinc-400" />
                    <input
                      type="text"
                      value={fontSearchQuery}
                      onChange={(e) => setFontSearchQuery(e.target.value)}
                      placeholder="搜索字体名称..."
                      className="w-full pl-9 pr-4 py-1.5 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400 dark:focus:ring-zinc-500"
                    />
                  </div>

                  {/* Categories & Actions */}
                  <div className="flex flex-wrap items-center gap-2 flex-1 min-w-0 justify-between md:justify-end">
                    <div className="flex items-center gap-1.5 overflow-x-auto scrollbar-none flex-1 min-w-0">
                      {/* 全部分组 Tab */}
                      <button
                        onClick={() => setFontCategoryFilter('全部分组')}
                        className={`px-3 py-1.5 text-xs font-semibold rounded-xl whitespace-nowrap transition-all ${
                          fontCategoryFilter === '全部分组'
                            ? 'bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 shadow-sm'
                            : 'text-zinc-600 dark:text-zinc-400 hover:bg-zinc-100 dark:hover:bg-zinc-800'
                        }`}
                      >
                        全部分组 ({fontsList.length})
                      </button>

                      {/* Custom Categories */}
                      {Array.from(
                        new Set(['默认', ...(appData.fontCategories || []), ...fontsList.map((p) => p.category || '默认')])
                      ).map((cat) => {
                        const count = fontsList.filter((p) => (p.category || '默认') === cat).length;
                        const isSelected = fontCategoryFilter === cat;

                        return (
                          <div
                            key={cat}
                            className="relative group flex items-center"
                            onContextMenu={(e) => {
                              e.preventDefault();
                              setManagingFontCategory(cat);
                              setRenameFontCategoryInput(cat);
                            }}
                          >
                            <button
                              onClick={() => setFontCategoryFilter(cat)}
                              className={`px-3 py-1.5 text-xs font-semibold rounded-xl whitespace-nowrap transition-all flex items-center gap-1.5 ${
                                isSelected
                                  ? 'bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 shadow-sm'
                                  : 'text-zinc-600 dark:text-zinc-400 hover:bg-zinc-100 dark:hover:bg-zinc-800'
                              }`}
                            >
                              <span>{cat}</span>
                              <span
                                className={`text-[10px] px-1.5 py-0.2 rounded-full ${
                                  isSelected
                                    ? 'bg-zinc-700 text-zinc-100 dark:bg-zinc-300 dark:text-zinc-900'
                                    : 'bg-zinc-100 dark:bg-zinc-800 text-zinc-500'
                                }`}
                              >
                                {count}
                              </span>
                            </button>

                            <button
                              onClick={(e) => {
                                e.stopPropagation();
                                setManagingFontCategory(cat);
                                setRenameFontCategoryInput(cat);
                              }}
                              className="p-1 text-zinc-400 hover:text-zinc-700 dark:hover:text-zinc-200"
                              title="右键或点击管理分组"
                            >
                              <MoreHorizontal className="w-3.5 h-3.5" />
                            </button>
                          </div>
                        );
                      })}

                      {/* 新建分组 Button */}
                      <button
                        onClick={() => setShowNewFontGroupModal(true)}
                        className="px-2.5 py-1.5 text-xs font-semibold rounded-xl border border-dashed border-zinc-300 dark:border-zinc-700 text-zinc-600 dark:text-zinc-400 hover:bg-zinc-100 dark:hover:bg-zinc-800 flex items-center gap-1 whitespace-nowrap flex-shrink-0 ml-auto"
                      >
                        <Plus className="w-3.5 h-3.5" /> 新建分组
                      </button>
                    </div>

                    {/* Sorting & 选择 (Batch Mode Toggle) Button */}
                    <div className="flex items-center justify-start gap-2 flex-shrink-0 basis-full">
                      <div className="flex items-center gap-1.5 bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-lg px-2.5 h-[30px] flex-shrink-0">
                        <ArrowUpDown className="w-3.5 h-3.5 text-zinc-400" />
                        <select
                          value={fontSortOrder}
                          onChange={(e) => setFontSortOrder(e.target.value as any)}
                          className="text-xs leading-none bg-transparent text-zinc-700 dark:text-zinc-300 focus:outline-none cursor-pointer"
                        >
                          <option value="default">默认排序</option>
                          <option value="az">名称 A-Z</option>
                          <option value="za">名称 Z-A</option>
                          <option value="newest">最新添加</option>
                          <option value="oldest">最早添加</option>
                        </select>
                      </div>

                      <button
                        onClick={() => {
                          setFontBatchMode(!fontBatchMode);
                          setSelectedFontIds([]);
                        }}
                        className={`flex items-center gap-1 px-3 py-1.5 text-xs font-semibold rounded-xl transition-all whitespace-nowrap flex-shrink-0 ${
                          fontBatchMode
                            ? 'bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 shadow-sm'
                            : 'border border-zinc-200 dark:border-zinc-800 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800'
                        }`}
                      >
                        <CheckSquare className="w-3.5 h-3.5" /> {fontBatchMode ? '退出选择' : '选择'}
                      </button>
                    </div>
                  </div>
                </div>

                {/* Batch Mode Toolbar */}
                {fontBatchMode && (
                  <div className="bg-zinc-900 dark:bg-zinc-100 text-white dark:text-zinc-900 rounded-2xl p-3 flex flex-col sm:flex-row items-center justify-between gap-3 shadow-lg animate-in fade-in">
                    <div className="text-xs font-medium">
                      已选择 <span className="font-bold underline">{selectedFontIds.length}</span> 个字体
                    </div>
                    <div className="flex items-center gap-2">
                      <button
                        onClick={() => {
                          if (selectedFontIds.length === filteredFonts.length) setSelectedFontIds([]);
                          else setSelectedFontIds(filteredFonts.map((p) => p.id));
                        }}
                        className="px-3 py-1.5 text-xs font-medium rounded-lg bg-zinc-800 dark:bg-zinc-200 text-zinc-200 dark:text-zinc-800 hover:opacity-90"
                      >
                        {selectedFontIds.length === filteredFonts.length ? '取消全选' : '全选当前'}
                      </button>
                      <button
                        disabled={selectedFontIds.length === 0}
                        onClick={() => setShowFontBatchMoveModal(true)}
                        className="px-3 py-1.5 text-xs font-medium rounded-lg bg-zinc-800 dark:bg-zinc-200 text-zinc-200 dark:text-zinc-800 hover:opacity-90 disabled:opacity-40"
                      >
                        批量移动
                      </button>
                      <button
                        disabled={selectedFontIds.length === 0}
                        onClick={handleBatchDeleteFonts}
                        className="px-3 py-1.5 text-xs font-bold rounded-lg bg-rose-600 text-white hover:bg-rose-700 disabled:opacity-40"
                      >
                        批量删除
                      </button>
                    </div>
                  </div>
                )}

                {/* Font Cards List */}
                {filteredFonts.length === 0 ? (
                  <div className="text-center py-20 bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl p-8 space-y-3">
                    <p className="text-xs font-medium text-zinc-500 dark:text-zinc-400">
                      {fontSearchQuery ? '未找到匹配的字体' : '暂无字体数据，点击右上角 “+” 按钮添加'}
                    </p>
                    {!fontSearchQuery && (
                      <button
                        onClick={() => setShowAddFontChoiceModal(true)}
                        className="px-4 py-2 bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 rounded-xl text-xs font-bold hover:opacity-90 inline-flex items-center gap-1.5 shadow-sm"
                      >
                        <Plus className="w-4 h-4" /> 添加字体
                      </button>
                    )}
                  </div>
                ) : (
                  <div className="space-y-3">
                    {sortItemList(
                      filteredFonts,
                      fontSortOrder,
                      (item) => item.name || '',
                      (item) => item.createdAt || 0
                    ).map((item) => {
                      const isSelected = selectedFontIds.includes(item.id);
                      const isActivePreview = activePreviewFontId === item.id;

                      return (
                        <div
                          key={item.id}
                          onClick={() => {
                            if (fontBatchMode) {
                              if (isSelected) {
                                setSelectedFontIds((p) => p.filter((id) => id !== item.id));
                              } else {
                                setSelectedFontIds((p) => [...p, item.id]);
                              }
                            } else {
                              setActivePreviewFontId(item.id);
                              showToast(`已切换预览字体为 “${item.name}”`, 'info');
                            }
                          }}
                          className={`p-4 bg-white dark:bg-zinc-900 border rounded-xl cursor-pointer transition-all hover:shadow-md flex flex-col md:flex-row md:items-center justify-between gap-3 ${
                            isActivePreview
                              ? 'border-zinc-900 dark:border-zinc-100 ring-2 ring-zinc-900/10 dark:ring-zinc-100/10'
                              : isSelected
                              ? 'border-rose-500 ring-2 ring-rose-500/30'
                              : 'border-zinc-200 dark:border-zinc-800 hover:border-zinc-300 dark:hover:border-zinc-700'
                          }`}
                        >
                          <div className="flex-1 min-w-0 space-y-1">
                            <div className="flex items-center gap-2">
                              <h3
                                className="text-sm font-bold text-zinc-900 dark:text-zinc-100 truncate"
                                style={{ fontFamily: item.fontFamily }}
                              >
                                {item.name}
                              </h3>
                              <span className="px-2 py-0.5 rounded text-[10px] font-semibold bg-zinc-100 dark:bg-zinc-800 text-zinc-600 dark:text-zinc-300">
                                {item.category || '默认'}
                              </span>
                              {isActivePreview && (
                                <span className="px-2 py-0.5 rounded text-[10px] font-bold bg-emerald-100 text-emerald-700 dark:bg-emerald-950 dark:text-emerald-300">
                                  预览中
                                </span>
                              )}
                            </div>

                            <p
                              className="text-xs text-zinc-600 dark:text-zinc-300 truncate pt-0.5"
                              style={{ fontFamily: item.fontFamily }}
                            >
                              永和九年，歲在癸丑 / AaBbCc 123
                            </p>

                            {item.url ? (
                              <p className="text-[11px] text-zinc-400 font-mono truncate">{item.url}</p>
                            ) : (
                              <p className="text-[11px] text-zinc-400 italic">本地文件上传</p>
                            )}
                          </div>

                          {fontBatchMode ? (
                            <div className="flex-shrink-0 self-end md:self-center">
                              {isSelected ? (
                                <CheckCircle2 className="w-5 h-5 text-rose-500 fill-white dark:fill-zinc-900" />
                              ) : (
                                <Circle className="w-5 h-5 text-zinc-300 dark:text-zinc-600" />
                              )}
                            </div>
                          ) : (
                            <div className="flex-shrink-0 flex items-center gap-2 self-end md:self-center">
                              <button
                                onClick={(e) => {
                                  e.stopPropagation();
                                  setEditingFont(JSON.parse(JSON.stringify(item)));
                                }}
                                className="px-3 py-1.5 text-xs font-semibold rounded-lg bg-zinc-100 dark:bg-zinc-800 hover:bg-zinc-200 dark:hover:bg-zinc-700 text-zinc-800 dark:text-zinc-200 transition-colors flex items-center gap-1"
                              >
                                <Edit3 className="w-3.5 h-3.5" /> 编辑
                              </button>
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>
            ) : currentPage === 'extras-app' ? (
              <div className="max-w-7xl mx-auto space-y-4">
                {/* Hidden File Input for Document Upload */}
                <input
                  type="file"
                  multiple
                  ref={extraStoryFileInputRef}
                  onChange={handleFileUploadExtraStory}
                  accept=".docx,.txt"
                  className="hidden"
                />

                {/* Toolbar (Search, Categories, Group Actions, Batch Mode) */}
                <div className="bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-xl p-2.5 flex flex-col md:flex-row items-stretch md:items-center justify-between gap-3">
                  {/* Search Input */}
                  <div className="relative w-full md:w-72">
                    <Search className="w-3.5 h-3.5 absolute left-3 top-1/2 -translate-y-1/2 text-zinc-400" />
                    <input
                      type="text"
                      value={extraStorySearchQuery}
                      onChange={(e) => setExtraStorySearchQuery(e.target.value)}
                      placeholder="搜索小剧场名称、作者或内容..."
                      className="w-full pl-9 pr-4 py-1.5 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400 dark:focus:ring-zinc-500"
                    />
                  </div>

                  {/* Categories & Actions */}
                  <div className="flex flex-wrap items-center gap-2 flex-1 min-w-0 justify-between md:justify-end">
                    <div className="flex items-center gap-1.5 overflow-x-auto scrollbar-none flex-1 min-w-0">
                      {/* 全部分组 Tab */}
                      <button
                        onClick={() => setExtraStoryCategoryFilter('全部分组')}
                        className={`px-3 py-1.5 text-xs font-semibold rounded-xl whitespace-nowrap transition-all ${
                          extraStoryCategoryFilter === '全部分组'
                            ? 'bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 shadow-sm'
                            : 'text-zinc-600 dark:text-zinc-400 hover:bg-zinc-100 dark:hover:bg-zinc-800'
                        }`}
                      >
                        全部分组 ({extraStoriesList.length})
                      </button>

                      {/* Custom Categories */}
                      {Array.from(
                        new Set(['默认', ...(appData.extraStoryCategories || []), ...extraStoriesList.map((p) => p.category || '默认')])
                      ).map((cat) => {
                        const count = extraStoriesList.filter((p) => (p.category || '默认') === cat).length;
                        const isSelected = extraStoryCategoryFilter === cat;

                        return (
                          <div
                            key={cat}
                            className="relative group flex items-center"
                            onContextMenu={(e) => {
                              e.preventDefault();
                              setManagingExtraStoryCategory(cat);
                              setRenameExtraStoryCategoryInput(cat);
                            }}
                          >
                            <button
                              onClick={() => setExtraStoryCategoryFilter(cat)}
                              className={`px-3 py-1.5 text-xs font-semibold rounded-xl whitespace-nowrap transition-all flex items-center gap-1.5 ${
                                isSelected
                                  ? 'bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 shadow-sm'
                                  : 'text-zinc-600 dark:text-zinc-400 hover:bg-zinc-100 dark:hover:bg-zinc-800'
                              }`}
                            >
                              <span>{cat}</span>
                              <span
                                className={`text-[10px] px-1.5 py-0.2 rounded-full ${
                                  isSelected
                                    ? 'bg-zinc-700 text-zinc-100 dark:bg-zinc-300 dark:text-zinc-900'
                                    : 'bg-zinc-100 dark:bg-zinc-800 text-zinc-500'
                                }`}
                              >
                                {count}
                              </span>
                            </button>

                            <button
                              onClick={(e) => {
                                e.stopPropagation();
                                setManagingExtraStoryCategory(cat);
                                setRenameExtraStoryCategoryInput(cat);
                              }}
                              className="p-1 text-zinc-400 hover:text-zinc-700 dark:hover:text-zinc-200"
                              title="右键或点击管理分组"
                            >
                              <MoreHorizontal className="w-3.5 h-3.5" />
                            </button>
                          </div>
                        );
                      })}

                      {/* 新建分组 Button */}
                      <button
                        onClick={() => setShowNewExtraStoryGroupModal(true)}
                        className="px-2.5 py-1.5 text-xs font-semibold rounded-xl border border-dashed border-zinc-300 dark:border-zinc-700 text-zinc-600 dark:text-zinc-400 hover:bg-zinc-100 dark:hover:bg-zinc-800 flex items-center gap-1 whitespace-nowrap flex-shrink-0 ml-auto"
                      >
                        <Plus className="w-3.5 h-3.5" /> 新建分组
                      </button>
                    </div>

                    {/* Sorting & 选择 (Batch Mode Toggle) Button */}
                    <div className="flex items-center justify-start gap-2 flex-shrink-0 basis-full">
                      <div className="flex items-center gap-1.5 bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-lg px-2.5 h-[30px] flex-shrink-0">
                        <ArrowUpDown className="w-3.5 h-3.5 text-zinc-400" />
                        <select
                          value={extraStorySortOrder}
                          onChange={(e) => setExtraStorySortOrder(e.target.value as any)}
                          className="text-xs leading-none bg-transparent text-zinc-700 dark:text-zinc-300 focus:outline-none cursor-pointer"
                        >
                          <option value="default">默认排序</option>
                          <option value="az">名称 A-Z</option>
                          <option value="za">名称 Z-A</option>
                          <option value="newest">最新添加</option>
                          <option value="oldest">最早添加</option>
                        </select>
                      </div>

                      <button
                        onClick={() => {
                          setExtraStoryBatchMode(!extraStoryBatchMode);
                          setSelectedExtraStoryIds([]);
                        }}
                        className={`flex items-center gap-1 px-3 py-1.5 text-xs font-semibold rounded-xl transition-all whitespace-nowrap flex-shrink-0 ${
                          extraStoryBatchMode
                            ? 'bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 shadow-sm'
                            : 'border border-zinc-200 dark:border-zinc-800 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800'
                        }`}
                      >
                        <CheckSquare className="w-3.5 h-3.5" /> {extraStoryBatchMode ? '退出选择' : '选择'}
                      </button>
                    </div>
                  </div>
                </div>

                {/* Batch Mode Toolbar */}
                {extraStoryBatchMode && (
                  <div className="bg-zinc-900 dark:bg-zinc-100 text-white dark:text-zinc-900 rounded-2xl p-3 flex flex-col sm:flex-row items-center justify-between gap-3 shadow-lg animate-in fade-in">
                    <div className="text-xs font-medium">
                      已选择 <span className="font-bold underline">{selectedExtraStoryIds.length}</span> 个番外
                    </div>
                    <div className="flex items-center gap-2">
                      <button
                        onClick={() => {
                          if (selectedExtraStoryIds.length === filteredExtraStories.length) setSelectedExtraStoryIds([]);
                          else setSelectedExtraStoryIds(filteredExtraStories.map((p) => p.id));
                        }}
                        className="px-3 py-1.5 text-xs font-medium rounded-lg bg-zinc-800 dark:bg-zinc-200 text-zinc-200 dark:text-zinc-800 hover:opacity-90"
                      >
                        {selectedExtraStoryIds.length === filteredExtraStories.length ? '取消全选' : '全选当前'}
                      </button>
                      <button
                        disabled={selectedExtraStoryIds.length === 0}
                        onClick={() => setShowExtraStoryBatchMoveModal(true)}
                        className="px-3 py-1.5 text-xs font-medium rounded-lg bg-zinc-800 dark:bg-zinc-200 text-zinc-200 dark:text-zinc-800 hover:opacity-90 disabled:opacity-40"
                      >
                        批量移动
                      </button>
                      <button
                        disabled={selectedExtraStoryIds.length === 0}
                        onClick={handleBatchDeleteExtraStories}
                        className="px-3 py-1.5 text-xs font-bold rounded-lg bg-rose-600 text-white hover:bg-rose-700 disabled:opacity-40"
                      >
                        批量删除
                      </button>
                    </div>
                  </div>
                )}

                {/* Extra Story List */}
                {filteredExtraStories.length === 0 ? (
                  <div className="text-center py-20 bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl p-8 space-y-3">
                    <p className="text-xs font-medium text-zinc-500 dark:text-zinc-400">
                      {extraStorySearchQuery ? '未找到匹配的番外小剧场' : '暂无番外小剧场，点击右上角 “+” 按钮录入或上传'}
                    </p>
                    {!extraStorySearchQuery && (
                      <button
                        onClick={() => setShowAddExtraStoryChoiceModal(true)}
                        className="px-4 py-2 bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 rounded-xl text-xs font-bold hover:opacity-90 inline-flex items-center gap-1.5 shadow-sm"
                      >
                        <Plus className="w-4 h-4" /> 添加番外小剧场
                      </button>
                    )}
                  </div>
                ) : (
                  <div className="space-y-3">
                    {sortItemList(
                      filteredExtraStories,
                      extraStorySortOrder,
                      (item) => item.title || '',
                      (item) => item.createdAt || 0
                    ).map((item) => {
                      const isSelected = selectedExtraStoryIds.includes(item.id);

                      return (
                        <div
                          key={item.id}
                          onClick={() => {
                            if (extraStoryBatchMode) {
                              if (isSelected) {
                                setSelectedExtraStoryIds((p) => p.filter((id) => id !== item.id));
                              } else {
                                setSelectedExtraStoryIds((p) => [...p, item.id]);
                              }
                            } else {
                              setEditingExtraStory(JSON.parse(JSON.stringify(item)));
                              setIsContentExpanded(false);
                            }
                          }}
                          className={`p-4 bg-white dark:bg-zinc-900 border rounded-xl cursor-pointer transition-all hover:shadow-md flex items-center justify-between gap-3 ${
                            isSelected
                              ? 'border-rose-500 ring-2 ring-rose-500/30'
                              : 'border-zinc-200 dark:border-zinc-800 hover:border-zinc-300 dark:hover:border-zinc-700'
                          }`}
                        >
                          <div className="flex-1 min-w-0 space-y-1">
                            <div className="flex items-center gap-2">
                              <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100 truncate">
                                {item.title}
                              </h3>
                              <span className="px-2 py-0.5 rounded text-[10px] font-semibold bg-zinc-100 dark:bg-zinc-800 text-zinc-600 dark:text-zinc-300 flex-shrink-0">
                                {item.category || '默认'}
                              </span>
                            </div>

                            {item.author && (
                              <p className="text-xs text-zinc-500 dark:text-zinc-400 truncate">
                                作者/来源：{item.author}
                              </p>
                            )}

                            <p className="text-xs text-zinc-400 dark:text-zinc-500 line-clamp-2 leading-relaxed">
                              {item.content}
                            </p>
                          </div>

                          {extraStoryBatchMode ? (
                            <div className="flex-shrink-0">
                              {isSelected ? (
                                <CheckCircle2 className="w-5 h-5 text-rose-500 fill-white dark:fill-zinc-900" />
                              ) : (
                                <Circle className="w-5 h-5 text-zinc-300 dark:text-zinc-600" />
                              )}
                            </div>
                          ) : (
                            <div className="flex-shrink-0 text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-200 p-1.5 rounded-lg hover:bg-zinc-100 dark:hover:bg-zinc-800 transition-colors">
                              <Edit3 className="w-4 h-4" />
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>
            ) : currentPage === 'stickers' ? (
              <div className="max-w-7xl mx-auto space-y-4">
                {/* Hidden File Input for Sticker Document Upload */}
                <input
                  type="file"
                  multiple
                  ref={stickerFileInputRef}
                  onChange={handleFileUploadSticker}
                  accept=".docx,.txt,.json"
                  className="hidden"
                />

                {/* Toolbar (Search, Categories, Group Actions, Batch Mode) */}
                <div className="bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-xl p-2.5 flex flex-col md:flex-row items-stretch md:items-center justify-between gap-3">
                  {/* Search Input */}
                  <div className="relative w-full md:w-72">
                    <Search className="w-3.5 h-3.5 absolute left-3 top-1/2 -translate-y-1/2 text-zinc-400" />
                    <input
                      type="text"
                      value={stickerSearchQuery}
                      onChange={(e) => setStickerSearchQuery(e.target.value)}
                      placeholder="搜索表情包名称、作者或内容..."
                      className="w-full pl-9 pr-4 py-1.5 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400 dark:focus:ring-zinc-500"
                    />
                  </div>

                  {/* Categories & Actions */}
                  <div className="flex flex-wrap items-center gap-2 flex-1 min-w-0">
                    <div className="w-full flex items-center gap-1.5 overflow-x-auto scrollbar-none">
                      {/* 全部分组 Tab */}
                      <button
                        onClick={() => setStickerCategoryFilter('全部分组')}
                        className={`px-3 py-1.5 text-xs font-semibold rounded-xl whitespace-nowrap transition-all ${
                          stickerCategoryFilter === '全部分组'
                            ? 'bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 shadow-sm'
                            : 'text-zinc-600 dark:text-zinc-400 hover:bg-zinc-100 dark:hover:bg-zinc-800'
                        }`}
                      >
                        全部分组 ({stickerPacksList.length})
                      </button>

                      {/* Custom Categories */}
                      {Array.from(
                        new Set(['默认', ...(appData.stickerCategories || []), ...stickerPacksList.map((p) => p.category || '默认')])
                      ).map((cat) => {
                        const count = stickerPacksList.filter((p) => (p.category || '默认') === cat).length;
                        const isSelected = stickerCategoryFilter === cat;

                        return (
                          <div
                            key={cat}
                            className="relative group flex items-center"
                            onContextMenu={(e) => {
                              e.preventDefault();
                              setManagingStickerCategory(cat);
                              setRenameStickerCategoryInput(cat);
                            }}
                          >
                            <button
                              onClick={() => setStickerCategoryFilter(cat)}
                              className={`px-3 py-1.5 text-xs font-semibold rounded-xl whitespace-nowrap transition-all flex items-center gap-1.5 ${
                                isSelected
                                  ? 'bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 shadow-sm'
                                  : 'text-zinc-600 dark:text-zinc-400 hover:bg-zinc-100 dark:hover:bg-zinc-800'
                              }`}
                            >
                              <span>{cat}</span>
                              <span
                                className={`text-[10px] px-1.5 py-0.2 rounded-full ${
                                  isSelected
                                    ? 'bg-zinc-700 text-zinc-100 dark:bg-zinc-300 dark:text-zinc-900'
                                    : 'bg-zinc-100 dark:bg-zinc-800 text-zinc-500'
                                }`}
                              >
                                {count}
                              </span>
                            </button>

                            <button
                              onClick={(e) => {
                                e.stopPropagation();
                                setManagingStickerCategory(cat);
                                setRenameStickerCategoryInput(cat);
                              }}
                              className="p-1 text-zinc-400 hover:text-zinc-700 dark:hover:text-zinc-200"
                              title="右键或点击管理分组"
                            >
                              <MoreHorizontal className="w-3.5 h-3.5" />
                            </button>
                          </div>
                        );
                      })}

                      {/* 新建分组 Button */}
                      <button
                        onClick={() => setShowNewStickerGroupModal(true)}
                        className="px-2.5 py-1.5 text-xs font-semibold rounded-xl border border-dashed border-zinc-300 dark:border-zinc-700 text-zinc-600 dark:text-zinc-400 hover:bg-zinc-100 dark:hover:bg-zinc-800 flex items-center gap-1 whitespace-nowrap flex-shrink-0 ml-auto"
                      >
                        <Plus className="w-3.5 h-3.5" /> 新建分组
                      </button>
                    </div>

                    {/* 选择 (Batch Mode Toggle) Button */}
                    <button
                      onClick={() => {
                        setStickerBatchMode(!stickerBatchMode);
                        setSelectedStickerPackIds([]);
                      }}
                      className={`flex items-center gap-1 px-3 py-1.5 text-xs font-semibold rounded-xl transition-all whitespace-nowrap flex-shrink-0 ${
                        stickerBatchMode
                          ? 'bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 shadow-sm'
                          : 'border border-zinc-200 dark:border-zinc-800 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800'
                      }`}
                    >
                      <CheckSquare className="w-3.5 h-3.5" /> {stickerBatchMode ? '退出选择' : '选择'}
                    </button>
                  </div>
                </div>

                {/* Batch Mode Toolbar */}
                {stickerBatchMode && (
                  <div className="bg-zinc-900 dark:bg-zinc-100 text-white dark:text-zinc-900 rounded-2xl p-3 flex flex-col sm:flex-row items-center justify-between gap-3 shadow-lg animate-in fade-in">
                    <div className="text-xs font-medium">
                      已选择 <span className="font-bold underline">{selectedStickerPackIds.length}</span> 个图集
                    </div>
                    <div className="flex items-center gap-2">
                      <button
                        onClick={() => {
                          if (selectedStickerPackIds.length === filteredStickerPacks.length) setSelectedStickerPackIds([]);
                          else setSelectedStickerPackIds(filteredStickerPacks.map((p) => p.id));
                        }}
                        className="px-3 py-1.5 text-xs font-medium rounded-lg bg-zinc-800 dark:bg-zinc-200 text-zinc-200 dark:text-zinc-800 hover:opacity-90"
                      >
                        {selectedStickerPackIds.length === filteredStickerPacks.length ? '取消全选' : '全选当前'}
                      </button>
                      <button
                        disabled={selectedStickerPackIds.length === 0}
                        onClick={() => setShowStickerBatchMoveModal(true)}
                        className="px-3 py-1.5 text-xs font-medium rounded-lg bg-zinc-800 dark:bg-zinc-200 text-zinc-200 dark:text-zinc-800 hover:opacity-90 disabled:opacity-40"
                      >
                        批量移动
                      </button>
                      <button
                        disabled={selectedStickerPackIds.length === 0}
                        onClick={handleBatchDeleteStickerPacks}
                        className="px-3 py-1.5 text-xs font-bold rounded-lg bg-rose-600 text-white hover:bg-rose-700 disabled:opacity-40"
                      >
                        批量删除
                      </button>
                    </div>
                  </div>
                )}

                {/* Sticker Pack Grid List (Matched with ST Card grid margins) */}
                {filteredStickerPacks.length === 0 ? (
                  <div className="text-center py-20 bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl p-8 space-y-3">
                    <p className="text-xs font-medium text-zinc-500 dark:text-zinc-400">
                      {stickerSearchQuery ? '未找到匹配的表情包图集' : '暂无表情包图集，点击右上角 “+” 按钮上传文档 (.docx / .txt / .json)'}
                    </p>
                    {!stickerSearchQuery && (
                      <button
                        onClick={() => openBatchUpload('stickers')}
                        className="px-4 py-2 bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 rounded-xl text-xs font-bold hover:opacity-90 inline-flex items-center gap-1.5 shadow-sm"
                      >
                        <Plus className="w-4 h-4" /> 上传表情包文档
                      </button>
                    )}
                  </div>
                ) : (
                  <div className="grid grid-cols-3 sm:grid-cols-4 md:grid-cols-5 lg:grid-cols-6 xl:grid-cols-7 gap-2.5">
                    {filteredStickerPacks.map((pack) => {
                      const isSelected = selectedStickerPackIds.includes(pack.id);
                      const firstSticker = pack.items[0];

                      return (
                        <div
                          key={pack.id}
                          onClick={() => {
                            if (stickerBatchMode) {
                              if (isSelected) {
                                setSelectedStickerPackIds((p) => p.filter((id) => id !== pack.id));
                              } else {
                                setSelectedStickerPackIds((p) => [...p, pack.id]);
                              }
                            } else {
                              setEditingStickerPack(JSON.parse(JSON.stringify(pack)));
                              setStickerDetailTab('info');
                            }
                          }}
                          className={`group bg-white dark:bg-zinc-900 border rounded-lg overflow-hidden cursor-pointer transition-all hover:shadow-md flex flex-col relative ${
                            isSelected
                              ? 'border-zinc-900 dark:border-zinc-100 ring-2 ring-zinc-900 dark:ring-zinc-100'
                              : 'border-zinc-200 dark:border-zinc-800 hover:border-zinc-400 dark:hover:border-zinc-600'
                          }`}
                        >
                          {/* Sticker Cover / Preview Image */}
                          <div className="aspect-[2/3] w-full bg-zinc-100 dark:bg-zinc-800 relative overflow-hidden flex items-center justify-center p-2">
                            {firstSticker?.url ? (
                              <img
                                src={firstSticker.url}
                                alt={pack.title}
                                className="w-full h-full object-cover group-hover:scale-105 transition-transform duration-300"
                                onError={(e) => {
                                  (e.target as HTMLElement).style.display = 'none';
                                }}
                              />
                            ) : (
                              <div className="text-[10px] text-zinc-400 flex flex-col items-center gap-1">
                                <ImageIcon className="w-6 h-6 opacity-40" />
                                <span>无表情图</span>
                              </div>
                            )}

                            {/* Sticker Count Badge */}
                            <span className="absolute top-1.5 right-1.5 px-1.5 py-0.5 rounded bg-black/70 text-white text-[9px] font-bold tracking-wider backdrop-blur-sm">
                              {pack.items.length} 张
                            </span>

                            {/* Batch Mode Selection Checkbox */}
                            {stickerBatchMode && (
                              <div className="absolute top-1.5 left-1.5">
                                {isSelected ? (
                                  <CheckCircle2 className="w-4 h-4 text-zinc-900 dark:text-zinc-100 fill-white dark:fill-zinc-900" />
                                ) : (
                                  <Circle className="w-4 h-4 text-white/80 drop-shadow" />
                                )}
                              </div>
                            )}
                          </div>

                          {/* Title & Info */}
                          <div className="p-2 flex-1 flex flex-col justify-between space-y-1 bg-white dark:bg-zinc-900">
                            <h3 className="text-xs font-bold text-zinc-900 dark:text-zinc-100 truncate group-hover:text-rose-500 transition-colors">
                              {pack.title}
                            </h3>
                            <div className="flex items-center justify-between text-[10px] text-zinc-400">
                              <span className="truncate">{pack.author || '默认'}</span>
                              <span className="px-1.5 py-0.2 rounded bg-zinc-100 dark:bg-zinc-800 text-zinc-500 flex-shrink-0">
                                {pack.category || '默认'}
                              </span>
                            </div>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>
            ) : currentPage === 'worldbook' ? (
              <div className="max-w-7xl mx-auto space-y-4">
                {/* Hidden File Input for World Book Upload */}
                <input
                  type="file"
                  multiple
                  ref={worldBookFileInputRef}
                  onChange={handleFileUploadWorldBook}
                  accept=".docx,.txt,.json"
                  className="hidden"
                />

                {/* Toolbar (Search, Categories, Group Actions, Batch Mode) */}
                <div className="bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-xl p-2.5 flex flex-col md:flex-row items-stretch md:items-center justify-between gap-3">
                  {/* Search Input (w-full md:w-72 - 搜索栏左右不用太长) */}
                  <div className="relative w-full md:w-72">
                    <Search className="w-3.5 h-3.5 absolute left-3 top-1/2 -translate-y-1/2 text-zinc-400" />
                    <input
                      type="text"
                      value={worldBookSearchQuery}
                      onChange={(e) => setWorldBookSearchQuery(e.target.value)}
                      placeholder="搜索世界书..."
                      className="w-full pl-9 pr-4 py-1.5 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400 dark:focus:ring-zinc-500"
                    />
                  </div>

                  {/* Categories & Actions */}
                  <div className="flex flex-wrap items-center gap-2 flex-1 min-w-0">
                    <div className="w-full flex items-center gap-1.5 overflow-x-auto scrollbar-none">
                      {/* 全部分组 Tab */}
                      <button
                        onClick={() => setWorldBookCategoryFilter('全部分组')}
                        className={`px-3 py-1.5 text-xs font-semibold rounded-xl whitespace-nowrap transition-all ${
                          worldBookCategoryFilter === '全部分组'
                            ? 'bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 shadow-sm'
                            : 'text-zinc-600 dark:text-zinc-400 hover:bg-zinc-100 dark:hover:bg-zinc-800'
                        }`}
                      >
                        全部分组 ({worldBooksList.length})
                      </button>

                      {/* Custom Categories */}
                      {Array.from(
                        new Set(['默认', ...(appData.worldBookCategories || []), ...worldBooksList.map((p) => p.category || '默认')])
                      ).map((cat) => {
                        const count = worldBooksList.filter((p) => (p.category || '默认') === cat).length;
                        const isSelected = worldBookCategoryFilter === cat;

                        return (
                          <div
                            key={cat}
                            className="relative group flex items-center"
                            onContextMenu={(e) => {
                              e.preventDefault();
                              setManagingWorldBookCategory(cat);
                              setRenameWorldBookCategoryInput(cat);
                            }}
                          >
                            <button
                              onClick={() => setWorldBookCategoryFilter(cat)}
                              className={`px-3 py-1.5 text-xs font-semibold rounded-xl whitespace-nowrap transition-all flex items-center gap-1.5 ${
                                isSelected
                                  ? 'bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 shadow-sm'
                                  : 'text-zinc-600 dark:text-zinc-400 hover:bg-zinc-100 dark:hover:bg-zinc-800'
                              }`}
                            >
                              <span>{cat}</span>
                              <span
                                className={`text-[10px] px-1.5 py-0.2 rounded-full ${
                                  isSelected
                                    ? 'bg-zinc-700 text-zinc-100 dark:bg-zinc-300 dark:text-zinc-900'
                                    : 'bg-zinc-100 dark:bg-zinc-800 text-zinc-500'
                                }`}
                              >
                                {count}
                              </span>
                            </button>

                            <button
                              onClick={(e) => {
                                e.stopPropagation();
                                setManagingWorldBookCategory(cat);
                                setRenameWorldBookCategoryInput(cat);
                              }}
                              className="p-1 text-zinc-400 hover:text-zinc-700 dark:hover:text-zinc-200"
                              title="右键或点击管理分组"
                            >
                              <MoreHorizontal className="w-3.5 h-3.5" />
                            </button>
                          </div>
                        );
                      })}

                      {/* 新建分组 Button (全部分组右侧) */}
                      <button
                        onClick={() => setShowNewWorldBookGroupModal(true)}
                        className="px-2.5 py-1.5 text-xs font-semibold rounded-xl border border-dashed border-zinc-300 dark:border-zinc-700 text-zinc-600 dark:text-zinc-400 hover:bg-zinc-100 dark:hover:bg-zinc-800 flex items-center gap-1 whitespace-nowrap flex-shrink-0 ml-auto"
                      >
                        <Plus className="w-3.5 h-3.5" /> 新建分组
                      </button>
                    </div>

                    {/* Sorting & 选择 (Batch Mode Toggle) Button */}
                    <div className="basis-full flex items-center justify-start gap-2 flex-shrink-0">
                      <div className="flex items-center gap-1.5 bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-lg px-2.5 h-[30px] flex-shrink-0">
                        <ArrowUpDown className="w-3.5 h-3.5 text-zinc-400" />
                        <select
                          value={worldBookSortOrder}
                          onChange={(e) => setWorldBookSortOrder(e.target.value as any)}
                          className="text-xs leading-none bg-transparent text-zinc-700 dark:text-zinc-300 focus:outline-none cursor-pointer"
                        >
                          <option value="default">默认排序</option>
                          <option value="az">名称 A-Z</option>
                          <option value="za">名称 Z-A</option>
                          <option value="newest">最新添加</option>
                          <option value="oldest">最早添加</option>
                        </select>
                      </div>

                      <button
                        onClick={() => {
                          setWorldBookBatchMode(!worldBookBatchMode);
                          setSelectedWorldBookIds([]);
                        }}
                        className={`flex items-center gap-1 px-3 py-1.5 text-xs font-semibold rounded-xl transition-all whitespace-nowrap flex-shrink-0 ${
                          worldBookBatchMode
                            ? 'bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 shadow-sm'
                            : 'border border-zinc-200 dark:border-zinc-800 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800'
                        }`}
                      >
                        <CheckSquare className="w-3.5 h-3.5" /> {worldBookBatchMode ? '退出选择' : '选择'}
                      </button>
                    </div>
                  </div>
                </div>

                {/* Batch Mode Toolbar */}
                {worldBookBatchMode && (
                  <div className="bg-zinc-900 dark:bg-zinc-100 text-white dark:text-zinc-900 rounded-2xl p-3 flex flex-col sm:flex-row items-center justify-between gap-3 shadow-lg animate-in fade-in">
                    <div className="text-xs font-medium">
                      已选择 <span className="font-bold underline">{selectedWorldBookIds.length}</span> 本世界书
                    </div>
                    <div className="flex items-center gap-2">
                      <button
                        onClick={() => {
                          if (selectedWorldBookIds.length === filteredWorldBooks.length) setSelectedWorldBookIds([]);
                          else setSelectedWorldBookIds(filteredWorldBooks.map((p) => p.id));
                        }}
                        className="px-3 py-1.5 text-xs font-medium rounded-lg bg-zinc-800 dark:bg-zinc-200 text-zinc-200 dark:text-zinc-800 hover:opacity-90"
                      >
                        {selectedWorldBookIds.length === filteredWorldBooks.length ? '取消全选' : '全选当前'}
                      </button>
                      <button
                        disabled={selectedWorldBookIds.length === 0}
                        onClick={() => setShowWorldBookBatchMoveModal(true)}
                        className="px-3 py-1.5 text-xs font-medium rounded-lg bg-zinc-800 dark:bg-zinc-200 text-zinc-200 dark:text-zinc-800 hover:opacity-90 disabled:opacity-40"
                      >
                        批量移动
                      </button>
                      <button
                        disabled={selectedWorldBookIds.length === 0}
                        onClick={handleBatchDeleteWorldBooks}
                        className="px-3 py-1.5 text-xs font-bold rounded-lg bg-rose-600 text-white hover:bg-rose-700 disabled:opacity-40"
                      >
                        批量删除
                      </button>
                    </div>
                  </div>
                )}

                {/* World Books List (Aligned with Phone Links panel margins and style) */}
                {filteredWorldBooks.length === 0 ? (
                  <div className="text-center py-20 bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl p-8 space-y-3">
                    <p className="text-xs font-medium text-zinc-500 dark:text-zinc-400">
                      {worldBookSearchQuery ? '未找到匹配的世界书' : '暂无世界书，点击右上角 “+” 按钮上传文档 (.docx / .txt / .json)'}
                    </p>
                    {!worldBookSearchQuery && (
                      <button
                        onClick={() => openBatchUpload('worldbook')}
                        className="px-4 py-2 bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 rounded-xl text-xs font-bold hover:opacity-90 inline-flex items-center gap-1.5 shadow-sm"
                      >
                        <Plus className="w-4 h-4" /> 上传世界书文档
                      </button>
                    )}
                  </div>
                ) : (
                  <div className="space-y-3">
                    {sortItemList(
                      filteredWorldBooks,
                      worldBookSortOrder,
                      (item) => item.title || '',
                      (item) => item.createdAt || 0
                    ).map((item) => {
                      const isSelected = selectedWorldBookIds.includes(item.id);

                      return (
                        <div
                          key={item.id}
                          onClick={() => {
                            if (worldBookBatchMode) {
                              if (isSelected) {
                                setSelectedWorldBookIds((p) => p.filter((id) => id !== item.id));
                              } else {
                                setSelectedWorldBookIds((p) => [...p, item.id]);
                              }
                            } else {
                              setEditingWorldBook(JSON.parse(JSON.stringify(item)));
                              setIsWorldBookContentExpanded(false);
                            }
                          }}
                          className={`p-4 bg-white dark:bg-zinc-900 border rounded-xl cursor-pointer transition-all hover:shadow-md flex items-center justify-between gap-4 ${
                            isSelected
                              ? 'border-zinc-900 dark:border-zinc-100 ring-2 ring-zinc-900/20 dark:ring-zinc-100/20'
                              : 'border-zinc-200 dark:border-zinc-800 hover:border-zinc-300 dark:hover:border-zinc-700'
                          }`}
                        >
                          {/* Title & Author */}
                          <div className="flex-1 min-w-0 space-y-1">
                            <div className="flex items-center gap-2">
                              <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100 truncate" title={item.title}>
                                {item.title}
                              </h3>
                              <span className="px-2 py-0.5 rounded text-[10px] font-semibold bg-zinc-100 dark:bg-zinc-800 text-zinc-600 dark:text-zinc-300 flex-shrink-0">
                                {item.category || '默认'}
                              </span>
                            </div>
                            <p className="text-xs text-zinc-500 dark:text-zinc-400 truncate">
                              作者：{item.author || '默认'}
                            </p>
                          </div>

                          {worldBookBatchMode ? (
                            <div className="flex-shrink-0">
                              {isSelected ? (
                                <CheckCircle2 className="w-5 h-5 text-rose-500 fill-white dark:fill-zinc-900" />
                              ) : (
                                <Circle className="w-5 h-5 text-zinc-300 dark:text-zinc-600" />
                              )}
                            </div>
                          ) : (
                            <div className="flex-shrink-0 text-xs text-zinc-400 dark:text-zinc-500 flex items-center gap-1">
                              <span>查看 / 编辑</span>
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>
            ) : (
              /* Placeholder Page for other non-st-cards pages */
              <div className="text-center py-24 max-w-md mx-auto">
                <h2 className="text-lg font-bold text-zinc-800 dark:text-zinc-200">
                  {currentPage} 板块
                </h2>
                <p className="text-xs text-zinc-400 mt-2">
                  此功能模块已就绪并归属于“通用”分组，请期待功能扩展。
                </p>
                <span className="inline-block mt-4 px-3 py-1 bg-zinc-200 dark:bg-zinc-800 text-zinc-600 dark:text-zinc-400 text-xs rounded-full">
                  Coming Soon
                </span>
              </div>
            )}
          </div>
        </div>
      </div>

      {/* ST 角色卡专用 JSON 导入输入 */}
      <input type="file" ref={cardWorldBookImportFileInputRef} accept=".json,application/json" className="hidden" onChange={(e) => handleCardSectionFileSelected(e, 'worldbook')} />
      <input type="file" ref={cardRegexFileInputRef} accept=".json,application/json" className="hidden" onChange={(e) => handleCardSectionFileSelected(e, 'regex')} />
      <input type="file" ref={cardQrFileInputRef} accept=".json,application/json" className="hidden" onChange={(e) => handleCardSectionFileSelected(e, 'qr')} />

      <input
        type="file"
        ref={sectionImportFileInputRef}
        accept=".json,application/json"
        className="hidden"
        onChange={(e) => {
          const file = e.target.files?.[0];
          if (file) void handleImportCurrentSection(file);
          e.target.value = '';
        }}
      />

      {/* Hidden File Inputs */}
      <input
        type="file"
        ref={themeFileInputRef}
        multiple
        accept=".json,application/json"
        className="hidden"
        onChange={(e) => {
          if (e.target.files) handleThemeFileUpload(e.target.files);
          e.target.value = '';
        }}
      />

      <input
        type="file"
        ref={themeCoverInputRef}
        accept="image/*"
        className="hidden"
        onChange={async (e) => {
          const file = e.target.files?.[0];
          if (file && editingTheme) {
            const scaled = await processImageFile(file, 600, 800);
            setEditingTheme((prev) => (prev ? { ...prev, coverImage: scaled } : null));
            showToast('封面图已设置，点击“确认”保存变更', 'info');
          }
          e.target.value = '';
        }}
      />

      <input
        type="file"
        ref={uploadFileInputRef}
        multiple
        accept=".png,.json,.webp,image/png,application/json"
        className="hidden"
        onChange={(e) => {
          if (e.target.files) handleFileUpload(e.target.files);
          e.target.value = '';
        }}
      />

      <input
        type="file"
        ref={qrFileInputRef}
        accept=".json,application/json,text/plain"
        className="hidden"
        onChange={(e) => {
          const file = e.target.files?.[0];
          if (file) handleQrFileImport(file);
          e.target.value = '';
        }}
      />

      <input
        type="file"
        ref={importBackupInputRef}
        accept=".json"
        className="hidden"
        onChange={(e) => {
          const file = e.target.files?.[0];
          e.target.value = '';
          if (!file) return;
          setIsImportingBackup(true);
          showToast('正在读取备份文件，数据量较大时请稍候…', 'info');
          const reader = new FileReader();
          reader.onload = () => {
            // 大文件 JSON.parse 是同步操作，放进 setTimeout 让上面的提示 toast 先绘制出来，
            // 避免用户以为页面卡死；解析/合并过程仍全程包裹在 try/catch 中，
            // 出错时只弹出错误提示，不会导致页面崩溃。
            setTimeout(() => {
              try {
                const data = JSON.parse(reader.result as string);
                if (data && typeof data === 'object' && Array.isArray(data.cards) && Array.isArray(data.groups)) {
                  // 侧边栏全部数据备份明确不包含字体；即使导入旧版本备份中带有 fonts/fontCategories，
                  // 这里也忽略这两个字段，避免覆盖当前字体数据。
                  const { fonts: _fonts, fontCategories: _fontCategories, ...backupData } = data;
                  updateAppData((prev) => ({ ...prev, ...backupData }));
                  showToast('全部数据导入成功！已恢复除字体外的全部板块数据', 'success');
                } else {
                  showToast('备份文件格式不符合要求', 'error');
                }
              } catch (err: any) {
                console.error('Import full backup failed', err);
                showToast('导入失败：文件可能已损坏或数据量过大（' + (err?.message || String(err)) + '）', 'error');
              } finally {
                setIsImportingBackup(false);
              }
            }, 20);
          };
          reader.onerror = () => {
            showToast('导入失败：文件读取出错', 'error');
            setIsImportingBackup(false);
          };
          reader.readAsText(file);
        }}
      />

      <input
        type="file"
        ref={updateCardInputRef}
        accept=".png,.json,image/png,application/json"
        className="hidden"
        onChange={(e) => {
          const file = e.target.files?.[0];
          if (file) handleUpdateCardOverwrite(file);
          e.target.value = '';
        }}
      />

      <input
        type="file"
        ref={presetRegexFileInputRef}
        accept=".json,.regex,application/json,text/plain"
        className="hidden"
        onChange={(e) => {
          const file = e.target.files?.[0];
          if (file) handleUploadPresetRegexFile(file);
          e.target.value = '';
        }}
      />

      {/* ==================== CARD DETAIL MODAL (Fixed Dimensions: Requirement 1) ==================== */}
      {detailCardId && activeDetailCard && activeLiveCardData && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-3 sm:p-4 pt-[calc(1rem+env(safe-area-inset-top,0px))] pb-[calc(1rem+env(safe-area-inset-bottom,0px))] animate-in fade-in">
          {/* Requirement 1: Fixed Modal Dimensions so no matter which tab is clicked, modal NEVER changes size */}
          <div className="w-full max-w-4xl h-[85vh] min-h-[550px] max-h-[750px] bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl flex flex-col overflow-hidden transition-none">
            {/* Modal Header */}
            <div className="px-6 py-4 border-b border-zinc-200 dark:border-zinc-800 flex items-center justify-between flex-shrink-0 bg-zinc-50/50 dark:bg-zinc-900/50">
              <div className="flex items-center gap-2 min-w-0">
                <h2 className="text-base font-bold truncate text-zinc-900 dark:text-zinc-100">
                  {getCardDisplayName(activeDetailCard)}
                </h2>
                <span className="px-2 py-0.5 text-[10px] font-bold uppercase rounded bg-zinc-200 dark:bg-zinc-800 text-zinc-700 dark:text-zinc-300">
                  {activeDetailCard.version}
                </span>
              </div>
              <button
                onClick={() => setDetailCardId(null)}
                className="p-1.5 rounded-lg hover:bg-zinc-200 dark:hover:bg-zinc-800 text-zinc-500 transition-colors"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Modal Navigation Tabs */}
            <div className="flex border-b border-zinc-200 dark:border-zinc-800 px-6 gap-2 overflow-x-auto flex-shrink-0 bg-white dark:bg-zinc-900 scrollbar-none">
              {[
                { id: 'overview', name: '概览' },
                { id: 'greetings', name: '开场白' },
                { id: 'worldbook', name: '世界书' },
                { id: 'regex', name: '正则' },
                { id: 'qr', name: 'QR' },
                { id: 'raw', name: '原始数据' },
                { id: 'extras', name: '其他' },
              ].map((tab) => (
                <button
                  key={tab.id}
                  onClick={() => setDetailTab(tab.id as any)}
                  className={`py-3 px-3 text-xs font-semibold border-b-2 whitespace-nowrap transition-colors ${
                    detailTab === tab.id
                      ? 'border-zinc-900 text-zinc-900 dark:border-zinc-100 dark:text-zinc-100'
                      : 'border-transparent text-zinc-500 hover:text-zinc-800 dark:hover:text-zinc-200'
                  }`}
                >
                  {tab.name}
                </button>
              ))}
            </div>

            {/* Modal Body: Scrollable Vertical Area within Fixed Height */}
            <div key={detailTab} className="flex-1 overflow-y-auto p-6 space-y-6">
              {/* TAB 1: 概览 (Overview) */}
              {detailTab === 'overview' && (
                <div className="grid grid-cols-1 md:grid-cols-12 gap-6">
                  {/* Left Column: Cover & Export / Update Options */}
                  <div className="md:col-span-4 space-y-4">
                    {/* Cover Main Box */}
                    <div className="aspect-[2/3] w-full rounded-xl overflow-hidden bg-zinc-100 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-800 relative flex items-center justify-center">
                      {activeDetailCard.coverImage ? (
                        <img src={activeDetailCard.coverImage} alt="Cover" className="w-full h-full object-contain" />
                      ) : (
                        <span className="text-xs text-zinc-400">无封面</span>
                      )}
                    </div>

                    {/* Change / Add Cover Button */}
                    <button
                      onClick={() => coverUploadInputRef.current?.click()}
                      className="w-full py-2 text-xs font-medium rounded-lg border border-dashed border-zinc-300 dark:border-zinc-700 hover:border-zinc-500 text-zinc-600 dark:text-zinc-400 transition-colors"
                    >
                      点击替换/更新封面图
                    </button>
                    <input
                      type="file"
                      ref={coverUploadInputRef}
                      accept="image/*"
                      className="hidden"
                      onChange={async (e) => {
                        const file = e.target.files?.[0];
                        if (file) {
                          const original = await fileToDataURL(file);
                          const updatedCards = appData.cards.map((c) =>
                            c.id === activeDetailCard.id ? { ...c, coverImage: original } : c
                          );
                          updateAppData({ ...appData, cards: updatedCards });
                          showToast('封面更新成功', 'success');
                        }
                        e.target.value = '';
                      }}
                    />

                    {/* Export & Update Section (Requirement 6: Add Update Character Card button) */}
                    <div className="space-y-2 pt-2 border-t border-zinc-200 dark:border-zinc-800">
                      <div className="text-xs font-bold text-zinc-500 uppercase tracking-wider">导出与覆盖更新</div>
                      <div className="grid grid-cols-2 gap-2">
                        <button
                          onClick={() => exportAsJson(activeDetailCard)}
                          className="py-1.5 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-800 hover:bg-zinc-100 dark:hover:bg-zinc-800 text-zinc-800 dark:text-zinc-200"
                        >
                          导出 JSON
                        </button>
                        <button
                          onClick={() => exportAsPng(activeDetailCard)}
                          className="py-1.5 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-800 hover:bg-zinc-100 dark:hover:bg-zinc-800 text-zinc-800 dark:text-zinc-200"
                        >
                          导出 PNG
                        </button>
                      </div>

                      {/* Requirement 6: "在概览里导出png的右侧新加一个更新角色卡，点击之后可以任选png/json格式更新，更新后覆盖先前的" */}
                      <button
                        onClick={() => updateCardInputRef.current?.click()}
                        className="w-full py-2 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:opacity-90 flex items-center justify-center gap-1.5 shadow-sm"
                      >
                        <RefreshCw className="w-3.5 h-3.5" /> 更新角色卡 (覆盖原卡)
                      </button>
                    </div>
                  </div>

                  {/* Right Column: Editable Information */}
                  <div className="md:col-span-8 space-y-4">
                    {/* Name Field */}
                    <div>
                      <label className="text-xs font-semibold text-zinc-500 dark:text-zinc-400 block mb-1">
                        角色名称
                      </label>
                      <input
                        type="text"
                        defaultValue={getCardDisplayName(activeDetailCard)}
                        onBlur={(e) => {
                          const val = e.target.value.trim();
                          if (val && val !== getCardDisplayName(activeDetailCard)) {
                            const updatedCards = appData.cards.map((c) =>
                              c.id === activeDetailCard.id
                                ? {
                                    ...c,
                                    name: val,
                                    editHistory: { ...c.editHistory, name: val },
                                    edited: true,
                                  }
                                : c
                            );
                            updateAppData({ ...appData, cards: updatedCards });
                          }
                        }}
                        className="w-full px-3 py-2 text-sm bg-zinc-50 dark:bg-zinc-800/50 border border-zinc-200 dark:border-zinc-700/60 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400"
                      />
                    </div>

                    {/* Author Field */}
                    <div>
                      <label className="text-xs font-semibold text-zinc-500 dark:text-zinc-400 block mb-1">
                        角色作者
                      </label>
                      <input
                        type="text"
                        id="authorInputVal"
                        defaultValue={getCardCreator(activeDetailCard)}
                        className="w-full px-3 py-2 text-sm bg-zinc-50 dark:bg-zinc-800/50 border border-zinc-200 dark:border-zinc-700/60 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400"
                      />
                    </div>

                    {/* Group Selection Field */}
                    <div>
                      <label className="text-xs font-semibold text-zinc-500 dark:text-zinc-400 block mb-1">
                        所属分组
                      </label>
                      <select
                        value={activeDetailCard.group || '默认'}
                        onChange={(e) => {
                          const newGrp = e.target.value;
                          const updatedCards = appData.cards.map((c) =>
                            c.id === activeDetailCard.id ? { ...c, group: newGrp } : c
                          );
                          updateAppData({ ...appData, cards: updatedCards });
                          showToast(`所属分组已更新为 "${newGrp}"`, 'success');
                        }}
                        className="w-full px-3 py-2 text-sm bg-zinc-50 dark:bg-zinc-800/50 border border-zinc-200 dark:border-zinc-700/60 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400"
                      >
                        {appData.groups.map((g) => (
                          <option key={g} value={g}>
                            {g}
                          </option>
                        ))}
                      </select>
                    </div>

                    {/* Description Field */}
                    <div>
                      <div className="flex items-center justify-between mb-1">
                        <label className="text-xs font-semibold text-zinc-500 dark:text-zinc-400 flex items-center gap-1.5">
                          角色描述 (Description)
                          <span className="text-[10px] font-normal text-zinc-400 bg-zinc-100 dark:bg-zinc-800 px-1.5 py-0.5 rounded">
                            估算 ~{estimateTokens(getCardDescription(activeDetailCard))} tokens
                          </span>
                        </label>
                        <button
                          onClick={() =>
                            setFullscreenData({
                              title: '角色描述',
                              content: getCardDescription(activeDetailCard),
                              type: 'text',
                            })
                          }
                          className="text-[11px] text-zinc-500 hover:text-zinc-900 dark:hover:text-zinc-100 flex items-center gap-1"
                        >
                          <Maximize2 className="w-3 h-3" /> 放大查看
                        </button>
                      </div>
                      <textarea
                        rows={6}
                        defaultValue={getCardDescription(activeDetailCard)}
                        onBlur={(e) => {
                          const val = e.target.value;
                          const updatedCards = appData.cards.map((c) =>
                            c.id === activeDetailCard.id
                              ? {
                                  ...c,
                                  editHistory: { ...c.editHistory, description: val },
                                  edited: true,
                                }
                              : c
                          );
                          updateAppData({ ...appData, cards: updatedCards });
                        }}
                        className="w-full px-3 py-2 text-sm bg-zinc-50 dark:bg-zinc-800/50 border border-zinc-200 dark:border-zinc-700/60 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400 leading-relaxed"
                      />
                    </div>
                  </div>
                </div>
              )}

              {/* TAB 2: 开场白 (Greetings - Editable) */}
              {detailTab === 'greetings' && (
                <div className="space-y-6">
                  {/* 自定义用户名：只替换界面中的 {{user}}，导出仍保留 {{user}} */}
                  <div className="p-3 rounded-xl border border-zinc-200 dark:border-zinc-700 bg-zinc-50/80 dark:bg-zinc-800/40">
                    <div className="flex items-center justify-between gap-3">
                      <div className="min-w-0">
                        <div className="text-xs font-bold text-zinc-800 dark:text-zinc-200">自定义用户名</div>
                        <div className="text-[10px] text-zinc-400 mt-0.5">仅用于当前角色卡的开场白预览，不会修改或导出原始 {'{{user}}'} 占位符</div>
                      </div>
                      <input
                        type="text"
                        value={customGreetingUsername}
                        onChange={(e) => setCustomGreetingUsername(e.target.value)}
                        placeholder="例如：小明"
                        className="w-36 shrink-0 px-3 py-2 text-xs bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400"
                      />
                    </div>
                  </div>

                  {/* Main Greeting (主开场白) */}
                  <div className="space-y-2">
                    <div className="space-y-2">
                      <div className="flex items-center justify-between gap-2">
                        <span className="text-xs font-bold text-zinc-800 dark:text-zinc-200 flex items-center gap-1.5">
                          主开场白 (支持编辑)
                          <span className="text-[10px] font-normal text-zinc-400 bg-zinc-100 dark:bg-zinc-800 px-1.5 py-0.5 rounded">
                            估算 ~{estimateTokens(getCardGreeting(activeDetailCard))} tokens
                          </span>
                        </span>
                      </div>
                      <div className="flex items-center justify-start gap-2">
                        <button
                          type="button"
                          onClick={() => downloadJsonFile(`${getCardDisplayName(activeDetailCard)}_st开场白.json`, {
                            first_mes: getCardGreeting(activeDetailCard),
                            alternate_greetings: getCardAlternateGreetings(activeDetailCard),
                          })}
                          className="px-2.5 py-1 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800 flex items-center gap-1"
                        >
                          <Download className="w-3.5 h-3.5" /> 导出 JSON
                        </button>
                        <button
                          onClick={() => setFullscreenData({ title: '主开场白', content: applyCustomGreetingUsername(getCardGreeting(activeDetailCard)), type: 'text' })}
                          className="px-2.5 py-1 text-xs text-zinc-500 hover:text-zinc-900 dark:hover:text-zinc-100 border border-transparent hover:border-zinc-200 dark:hover:border-zinc-700 rounded-lg flex items-center gap-1"
                        >
                          <Maximize2 className="w-3 h-3" /> 放大查看
                        </button>
                      </div>
                    </div>
                    <textarea
                      key={`main-greeting-${activeDetailCard.id}-${customGreetingUsername}`}
                      rows={5}
                      defaultValue={applyCustomGreetingUsername(getCardGreeting(activeDetailCard))}
                      onBlur={(e) => {
                        const val = restoreCustomGreetingUsername(e.target.value);
                        const updatedCards = appData.cards.map((c) =>
                          c.id === activeDetailCard.id
                            ? { ...c, editHistory: { ...c.editHistory, first_mes: val }, edited: true }
                            : c
                        );
                        updateAppData({ ...appData, cards: updatedCards });
                        showToast('主开场白已更新', 'success');
                      }}
                      placeholder="请输入主开场白内容..."
                      className="w-full p-3 text-xs leading-relaxed bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-xl focus:outline-none focus:ring-1 focus:ring-zinc-400 text-zinc-900 dark:text-zinc-100 resize-none"
                    />
                  </div>

                  {/* Alternate Greetings List (备用开场白) */}
                  <div className="space-y-3">
                    <div className="space-y-2">
                      <div className="text-xs font-bold text-zinc-800 dark:text-zinc-200">
                        备用开场白 ({getCardAlternateGreetings(activeDetailCard).length})
                      </div>
                      <div className="flex items-center justify-start gap-2">
                        <span className="text-[10px] font-normal text-zinc-400 bg-zinc-100 dark:bg-zinc-800 px-1.5 py-1 rounded">
                          总计 ~{getCardAlternateGreetings(activeDetailCard).reduce((sum, g) => sum + estimateTokens(g), 0)} tokens
                        </span>
                        <button
                          onClick={() => {
                            setShowAddAltGreetingModal(true);
                            setNewAltGreetingInputText('');
                          }}
                          className="px-2.5 py-1 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:opacity-90 flex items-center gap-1"
                        >
                          <Plus className="w-3.5 h-3.5" /> 添加备用开场白
                        </button>
                      </div>
                    </div>

                    {getCardAlternateGreetings(activeDetailCard).length === 0 ? (
                      <p className="text-xs text-zinc-400 italic py-2">暂无备用开场白</p>
                    ) : (
                      getCardAlternateGreetings(activeDetailCard).map((altGreeting, idx) => (
                        <div key={idx} className="p-3 bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-xl space-y-2">
                          <div className="flex items-center justify-between">
                            <span className="text-xs font-bold text-zinc-500 flex items-center gap-1.5">
                              备用 #{idx + 1}
                              <span className="text-[10px] font-normal text-zinc-400">
                                (~{estimateTokens(altGreeting)} tokens)
                              </span>
                            </span>
                            <div className="flex items-center gap-2">
                              <button
                                onClick={() =>
                                  setFullscreenData({
                                    title: `备用开场白 #${idx + 1}`,
                                    content: applyCustomGreetingUsername(altGreeting),
                                    type: 'text',
                                  })
                                }
                                className="text-[11px] text-zinc-500 hover:text-zinc-900 dark:hover:text-zinc-100 flex items-center gap-1"
                              >
                                <Maximize2 className="w-3 h-3" /> 放大
                              </button>
                              <button
                                onClick={() => {
                                  const currentAlts = getCardAlternateGreetings(activeDetailCard);
                                  const updatedAlts = currentAlts.filter((_, i) => i !== idx);
                                  const updatedCards = appData.cards.map((c) =>
                                    c.id === activeDetailCard.id
                                      ? { ...c, editHistory: { ...c.editHistory, alternate_greetings: updatedAlts }, edited: true }
                                      : c
                                  );
                                  updateAppData({ ...appData, cards: updatedCards });
                                  showToast('已删除备用开场白', 'info');
                                }}
                                className="text-xs text-rose-500 hover:text-rose-700 p-1"
                              >
                                <Trash2 className="w-3.5 h-3.5" />
                              </button>
                            </div>
                          </div>
                          <textarea
                            key={`alt-greeting-${activeDetailCard.id}-${idx}-${customGreetingUsername}`}
                            rows={3}
                            defaultValue={applyCustomGreetingUsername(altGreeting)}
                            onBlur={(e) => {
                              const val = restoreCustomGreetingUsername(e.target.value);
                              const currentAlts = [...getCardAlternateGreetings(activeDetailCard)];
                              currentAlts[idx] = val;
                              const updatedCards = appData.cards.map((c) =>
                                c.id === activeDetailCard.id
                                  ? { ...c, editHistory: { ...c.editHistory, alternate_greetings: currentAlts }, edited: true }
                                  : c
                              );
                              updateAppData({ ...appData, cards: updatedCards });
                              showToast('备用开场白已更新', 'success');
                            }}
                            placeholder="输入备用开场白内容..."
                            className="w-full p-2.5 text-xs leading-relaxed bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400 text-zinc-900 dark:text-zinc-100 resize-none"
                          />
                        </div>
                      ))
                    )}
                  </div>
                </div>
              )}

              {/* TAB 3: 世界书 (WorldBook - Editable) */}
              {detailTab === 'worldbook' && (() => {
                const currentWb = getCardWorldBook(activeDetailCard) || { entries: [] };
                const entries = currentWb.entries || [];
                const totalWbTokens = entries.reduce(
                  (sum: number, entry: any) => sum + estimateTokens(entry.content || '') + estimateTokens(Array.isArray(entry.keys) ? entry.keys.join(', ') : entry.keys || ''),
                  0
                );

                return (
                  <div className="space-y-4">
                    {/* 世界书总名称 (Requirement 6) */}
                    <div className="p-3 bg-zinc-100/70 dark:bg-zinc-800/80 border border-zinc-200 dark:border-zinc-700/60 rounded-xl space-y-1">
                      <label className="text-[11px] font-bold text-zinc-700 dark:text-zinc-300 block">
                        世界书总名称
                      </label>
                      <input
                        type="text"
                        defaultValue={currentWb.name || `${getCardDisplayName(activeDetailCard)} 的世界书`}
                        onBlur={(e) => {
                          const val = e.target.value.trim();
                          const wb = JSON.parse(JSON.stringify(getCardWorldBook(activeDetailCard) || { entries: [] }));
                          wb.name = val;
                          const updatedCards = appData.cards.map((c) =>
                            c.id === activeDetailCard.id
                              ? { ...c, editHistory: { ...c.editHistory, character_book: wb }, edited: true }
                              : c
                          );
                          updateAppData({ ...appData, cards: updatedCards });
                          showToast('世界书总名称已更新', 'success');
                        }}
                        placeholder="请输入世界书总名称..."
                        className="w-full px-3 py-1.5 text-xs bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-700 rounded-lg text-zinc-900 dark:text-zinc-100 focus:outline-none font-medium"
                      />
                    </div>

                    <div className="space-y-2">
                      <div className="flex items-center justify-between gap-2">
                        <span className="text-xs font-bold text-zinc-800 dark:text-zinc-200">世界书条目 ({entries.length})</span>
                        <span className="text-[10px] font-normal text-zinc-400 bg-zinc-100 dark:bg-zinc-800 px-1.5 py-1 rounded">总计 ~{totalWbTokens} tokens</span>
                      </div>
                      <div className="flex items-center justify-start gap-1.5 flex-nowrap">
                        <button
                          type="button"
                          onClick={() => downloadJsonFile(`${getCardDisplayName(activeDetailCard)}_worldbook.json`, currentWb)}
                          className="px-2.5 py-1 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800 flex items-center gap-1"
                        >
                          <Download className="w-3.5 h-3.5" /> 导出 JSON
                        </button>
                        <button
                          type="button"
                          onClick={() => setCardSectionImportModal('worldbook')}
                          className="px-2.5 py-1 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800 flex items-center gap-1"
                        >
                          <Upload className="w-3.5 h-3.5" /> 导入
                        </button>
                        <button
                        onClick={() => {
                          setNewCardWorldBookEntryForm({ comment: '', keys: '', content: '' });
                          setShowAddCardWorldBookEntryModal(true);
                        }}
                        className="px-2.5 py-1 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:opacity-90 flex items-center gap-1"
                      >
                        <Plus className="w-3.5 h-3.5" /> 添加条目
                        </button>
                      </div>
                    </div>

                    {entries.length === 0 ? (
                      <p className="text-xs text-zinc-400 italic py-4 text-center">此角色卡暂无世界书条目</p>
                    ) : (
                      <div className="space-y-3">
                        {entries.map((entry: any, i: number) => {
                          const entryTokens = estimateTokens(entry.content || '') + estimateTokens(Array.isArray(entry.keys) ? entry.keys.join(', ') : entry.keys || '');
                          return (
                            <div key={i} className="p-3.5 bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-xl space-y-2.5">
                              <div className="flex items-center justify-between gap-2">
                                <input
                                  type="text"
                                  defaultValue={entry.comment || entry.name || `条目 #${i + 1}`}
                                  onBlur={(e) => {
                                    const val = e.target.value;
                                    const wb = JSON.parse(JSON.stringify(getCardWorldBook(activeDetailCard)));
                                    wb.entries[i].comment = val;
                                    const updatedCards = appData.cards.map((c) =>
                                      c.id === activeDetailCard.id
                                        ? { ...c, editHistory: { ...c.editHistory, character_book: wb }, edited: true }
                                        : c
                                    );
                                    updateAppData({ ...appData, cards: updatedCards });
                                    showToast('条目标题已保存', 'success');
                                  }}
                                  placeholder="条目备注 / 标题..."
                                  className="px-2.5 py-1 text-xs font-bold bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-700 rounded-lg text-zinc-900 dark:text-zinc-100 flex-1 min-w-0"
                                />
                                <span className="text-[10px] text-zinc-400 font-normal">
                                  ~{entryTokens} tokens
                                </span>
                            <div className="flex items-center gap-1.5 flex-shrink-0">
                              <button
                                onClick={() =>
                                  setFullscreenData({
                                    title: entry.comment || `条目 #${i + 1}`,
                                    content: entry.content || '',
                                    type: 'text',
                                  })
                                }
                                className="text-[11px] text-zinc-500 hover:text-zinc-900 dark:hover:text-zinc-100 p-1 flex-shrink-0"
                              >
                                <Maximize2 className="w-3.5 h-3.5" />
                              </button>
                              <button
                                onClick={() => {
                                  const wb = JSON.parse(JSON.stringify(getCardWorldBook(activeDetailCard)));
                                  wb.entries.splice(i, 1);
                                  const updatedCards = appData.cards.map((c) =>
                                    c.id === activeDetailCard.id
                                      ? { ...c, editHistory: { ...c.editHistory, character_book: wb }, edited: true }
                                      : c
                                  );
                                  updateAppData({ ...appData, cards: updatedCards });
                                  showToast('世界书条目已删除', 'info');
                                }}
                                className="text-xs text-rose-500 hover:text-rose-700 p-1 flex-shrink-0"
                              >
                                <Trash2 className="w-3.5 h-3.5" />
                              </button>
                            </div>
                          </div>

                          {/* Trigger Keywords */}
                          <div>
                            <label className="text-[10px] font-semibold text-zinc-400 block mb-0.5">触发关键词 (英文逗号分隔)</label>
                            <input
                              type="text"
                              defaultValue={Array.isArray(entry.key || entry.keys) ? (entry.key || entry.keys).join(', ') : (entry.key || entry.keys || '')}
                              onBlur={(e) => {
                                const val = e.target.value.split(',').map((k) => k.trim()).filter(Boolean);
                                const wb = JSON.parse(JSON.stringify(getCardWorldBook(activeDetailCard)));
                                wb.entries[i].keys = val;
                                wb.entries[i].key = val;
                                const updatedCards = appData.cards.map((c) =>
                                  c.id === activeDetailCard.id
                                    ? { ...c, editHistory: { ...c.editHistory, character_book: wb }, edited: true }
                                    : c
                                );
                                updateAppData({ ...appData, cards: updatedCards });
                                showToast('关键词已保存', 'success');
                              }}
                              placeholder="如: 关键词1, 关键词2..."
                              className="w-full px-2.5 py-1 text-xs bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-700 rounded-lg text-zinc-800 dark:text-zinc-200"
                            />
                          </div>

                          {/* Content text */}
                          <div>
                            <label className="text-[10px] font-semibold text-zinc-400 block mb-0.5">条目具体正文内容</label>
                            <textarea
                              rows={4}
                              defaultValue={entry.content || ''}
                              onBlur={(e) => {
                                const val = e.target.value;
                                const wb = JSON.parse(JSON.stringify(getCardWorldBook(activeDetailCard)));
                                wb.entries[i].content = val;
                                const updatedCards = appData.cards.map((c) =>
                                  c.id === activeDetailCard.id
                                    ? { ...c, editHistory: { ...c.editHistory, character_book: wb }, edited: true }
                                    : c
                                );
                                updateAppData({ ...appData, cards: updatedCards });
                                showToast('世界书内容已保存', 'success');
                              }}
                              placeholder="请输入世界书条目正文..."
                              className="w-full p-2.5 text-xs leading-relaxed bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none text-zinc-900 dark:text-zinc-100 resize-none"
                            />
                          </div>
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>
              );
            })()}

              {/* TAB 4: 正则 (Regex - Editable) */}
              {detailTab === 'regex' && (
                <div className="space-y-4">
                  <div className="relative w-full">
                    <Search className="w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 text-zinc-400" />
                    <input
                      type="text"
                      value={cardRegexSearchQuery}
                      onChange={(e) => setCardRegexSearchQuery(e.target.value)}
                      placeholder="搜索正则脚本名称、查找或替换内容..."
                      className="w-full pl-9 pr-4 py-1.5 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400 dark:focus:ring-zinc-500"
                    />
                  </div>

                  <div className="space-y-2">
                    <div className="text-xs font-bold text-zinc-800 dark:text-zinc-200">
                      正则脚本 ({getCardRegex(activeDetailCard)?.length || 0})
                    </div>
                    <div className="flex items-center justify-start gap-2 flex-nowrap">
                      <button
                        type="button"
                        onClick={() => downloadJsonFile(`${getCardDisplayName(activeDetailCard)}_regex.json`, getCardRegex(activeDetailCard) || [])}
                        className="px-2.5 py-1 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800 flex items-center gap-1.5"
                      >
                        <Download className="w-3.5 h-3.5" /> 导出 JSON
                      </button>
                      <button
                        type="button"
                        onClick={() => setCardSectionImportModal('regex')}
                        className="px-2.5 py-1 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800 flex items-center gap-1.5"
                      >
                        <Upload className="w-3.5 h-3.5" /> 导入
                      </button>
                      <button
                        type="button"
                        onClick={() => {
                          const currentRegex = [...(getCardRegex(activeDetailCard) || [])];
                          currentRegex.push({
                            scriptName: '新正则脚本', script_name: '新正则脚本', name: '新正则脚本',
                            findRegex: '', find_regex: '', pattern: '',
                            replaceString: '', replace_string: '', replacement: '',
                            disabled: false,
                          });
                          saveCardRegexList(currentRegex);
                          setEditingCardRegex({ index: currentRegex.length - 1, scriptName: '新正则脚本', findRegex: '', replaceString: '' });
                        }}
                        className="px-2.5 py-1 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:opacity-90 flex items-center gap-1"
                      >
                        <Plus className="w-3.5 h-3.5" /> 添加正则脚本
                      </button>
                    </div>
                  </div>

                  {(!getCardRegex(activeDetailCard) || getCardRegex(activeDetailCard).length === 0) ? (
                    <div className="text-center py-8 space-y-3 bg-zinc-50 dark:bg-zinc-800/30 border border-dashed border-zinc-200 dark:border-zinc-800 rounded-2xl">
                      <p className="text-xs text-zinc-400">此角色卡暂无正则脚本</p>
                      <button
                        onClick={() => setCardSectionImportModal('regex')}
                        className="px-3.5 py-1.5 text-xs font-bold rounded-xl bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:opacity-90 inline-flex items-center gap-1.5 shadow-sm"
                      >
                        <Upload className="w-3.5 h-3.5" /> 上传正则脚本文件
                      </button>
                    </div>
                  ) : (
                    <div className="space-y-3">
                      {getCardRegex(activeDetailCard).map((rx: any, i: number) => {
                        const scriptName = rx.scriptName || rx.script_name || rx.name || rx.title || `正则 #${i + 1}`;
                        const findRegex = rx.findRegex || rx.find_regex || rx.pattern || rx.find || rx.regex || '';
                        const replaceString = rx.replaceString || rx.replace_string || rx.replacement || rx.replace || '';
                        const regexQuery = cardRegexSearchQuery.trim().toLowerCase();
                        if (regexQuery && !`${scriptName} ${findRegex} ${replaceString}`.toLowerCase().includes(regexQuery)) return null;
                        return (
                          <div key={i} className="p-3.5 bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-xl">
                            <div className="flex items-center justify-between gap-2">
                              <button
                                type="button"
                                onClick={() => setEditingCardRegex({ index: i, scriptName, findRegex, replaceString })}
                                className="min-w-0 flex-1 text-left"
                              >
                                <div className="text-xs font-bold text-zinc-900 dark:text-zinc-100 truncate">{scriptName}</div>
                                <div className="text-[11px] text-zinc-400 mt-1 truncate">查找：{findRegex || '（空）'}</div>
                                <div className="text-[11px] text-zinc-400 truncate">替换：{replaceString || '（空）'}</div>
                              </button>
                              <div className="flex items-center gap-1.5 flex-shrink-0">
                                <button
                                  type="button"
                                  onClick={() => setEditingCardRegex({ index: i, scriptName, findRegex, replaceString })}
                                  className="px-2.5 py-1 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 hover:bg-zinc-100 dark:hover:bg-zinc-800"
                                >编辑</button>
                                <button
                                  type="button"
                                  onClick={() => {
                                    const rxList = JSON.parse(JSON.stringify(getCardRegex(activeDetailCard)));
                                    rxList.splice(i, 1);
                                    saveCardRegexList(rxList);
                                    showToast('正则脚本已删除', 'info');
                                  }}
                                  className="p-1 text-zinc-400 hover:text-rose-500"
                                  title="删除"
                                ><Trash2 className="w-3.5 h-3.5" /></button>
                              </div>
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  )}
                </div>
              )}

              {/* TAB 6: 原始数据 (Raw Data) */}
              {detailTab === 'raw' && (
                <div className="space-y-2">
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-bold text-zinc-700 dark:text-zinc-300">完整 JSON 原始数据</span>
                    <button
                      onClick={() =>
                        setFullscreenData({
                          title: '原始 JSON 数据',
                          content: JSON.stringify(activeLiveCardData, null, 2),
                          type: 'json',
                        })
                      }
                      className="text-[11px] text-zinc-500 hover:text-zinc-900 dark:hover:text-zinc-100 flex items-center gap-1"
                    >
                      <Maximize2 className="w-3 h-3" /> 放大查看
                    </button>
                  </div>
                  <pre className="p-4 bg-zinc-900 text-zinc-100 text-xs font-mono rounded-lg overflow-x-auto max-h-96">
                    {JSON.stringify(activeLiveCardData, null, 2)}
                  </pre>
                </div>
              )}

              {/* TAB 6: 其他 (Extras) */}
              {detailTab === 'extras' && (
                <div className="space-y-6">
                  {/* Source Field (来源: QQ号/群、社区链接等，选填) */}
                  <div className="p-4 bg-zinc-50 dark:bg-zinc-800/50 border border-zinc-200 dark:border-zinc-700/60 rounded-xl space-y-1.5">
                    <label className="block text-xs font-bold text-zinc-800 dark:text-zinc-200">
                      来源 / 出处（QQ号/群、Discord、社区链接等，选填）
                    </label>
                    <input
                      type="text"
                      defaultValue={activeDetailCard.source || ''}
                      onBlur={(e) => {
                        const val = e.target.value;
                        const updatedCards = appData.cards.map((c) =>
                          c.id === activeDetailCard.id ? { ...c, source: val } : c
                        );
                        updateAppData({ ...appData, cards: updatedCards });
                        showToast('来源信息已保存', 'success');
                      }}
                      placeholder="例如: QQ群: 12345678 / https://t.me/..."
                      className="w-full px-3 py-2 text-xs bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400 text-zinc-900 dark:text-zinc-100"
                    />
                  </div>

                  {/* Section 1: 作者的话 */}
                  <div className="space-y-3">
                    <div className="flex flex-wrap items-center justify-between gap-2">
                      <h3 className="text-xs font-bold text-zinc-800 dark:text-zinc-200">
                        作者的话（截图） ({activeDetailCard.screenshots?.authorsNote?.length || 0})
                      </h3>
                      <div className="flex flex-wrap items-center gap-2">
                        <button
                          onClick={() => authorNoteInputRef.current?.click()}
                          className="px-2.5 py-1 text-xs font-medium rounded bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 flex items-center gap-1"
                        >
                          <Upload className="w-3 h-3" /> 上传截图
                        </button>

                        {(activeDetailCard.screenshots?.authorsNote?.length || 0) > 0 && (
                          <button
                            onClick={() => {
                              setAuthorNoteBatchMode((p) => !p);
                              setSelectedAuthorNoteIndices([]);
                            }}
                            className={`px-2.5 py-1 text-xs font-medium rounded border transition-colors ${
                              authorNoteBatchMode
                                ? 'bg-amber-100 text-amber-900 border-amber-300 dark:bg-amber-950 dark:text-amber-200 dark:border-amber-800'
                                : 'border-zinc-200 dark:border-zinc-800 hover:bg-zinc-100 dark:hover:bg-zinc-800'
                            }`}
                          >
                            {authorNoteBatchMode ? '退出批量' : '批量删除截图'}
                          </button>
                        )}

                        {authorNoteBatchMode && (activeDetailCard.screenshots?.authorsNote?.length || 0) > 0 && (
                          <>
                            <button
                              onClick={() => {
                                const allCount = activeDetailCard.screenshots?.authorsNote?.length || 0;
                                if (selectedAuthorNoteIndices.length === allCount) {
                                  setSelectedAuthorNoteIndices([]);
                                } else {
                                  setSelectedAuthorNoteIndices(Array.from({ length: allCount }, (_, i) => i));
                                }
                              }}
                              className="px-2 py-1 text-xs font-medium rounded border border-zinc-200 dark:border-zinc-800 hover:bg-zinc-100 dark:hover:bg-zinc-800"
                            >
                              {selectedAuthorNoteIndices.length === (activeDetailCard.screenshots?.authorsNote?.length || 0)
                                ? '取消全选'
                                : '全选'}
                            </button>

                            {selectedAuthorNoteIndices.length > 0 && (
                              <button
                                onClick={() => {
                                  if (!confirm(`确定要删除选中的 ${selectedAuthorNoteIndices.length} 张“作者的话”截图吗？`)) return;
                                  const updatedNotes = (activeDetailCard.screenshots?.authorsNote || []).filter(
                                    (_, idx) => !selectedAuthorNoteIndices.includes(idx)
                                  );
                                  const updatedCards = appData.cards.map((c) =>
                                    c.id === activeDetailCard.id
                                      ? {
                                          ...c,
                                          screenshots: {
                                            authorsNote: updatedNotes,
                                            favoriteScenes: c.screenshots?.favoriteScenes || [],
                                          },
                                        }
                                      : c
                                  );
                                  updateAppData({ ...appData, cards: updatedCards });
                                  setSelectedAuthorNoteIndices([]);
                                  setAuthorNoteBatchMode(false);
                                  showToast('已批量删除选中的截图', 'info');
                                }}
                                className="px-2.5 py-1 text-xs font-bold rounded bg-rose-600 text-white hover:bg-rose-700 flex items-center gap-1"
                              >
                                <Trash2 className="w-3 h-3" /> 删除选中 ({selectedAuthorNoteIndices.length})
                              </button>
                            )}
                          </>
                        )}
                      </div>
                    </div>

                    <input
                      type="file"
                      ref={authorNoteInputRef}
                      multiple
                      accept="image/*"
                      className="hidden"
                      onChange={async (e) => {
                        const files = Array.from(e.target.files || []) as File[];
                        if (files.length === 0) return;
                        const newScreenshots: string[] = [];
                        for (const f of files) {
                          const scaled = await processImageFile(f, 800, 800);
                          newScreenshots.push(scaled);
                          await new Promise((resolve) => setTimeout(resolve, 0));
                        }
                        const existing = activeDetailCard.screenshots?.authorsNote || [];
                        const updatedCards = appData.cards.map((c) =>
                          c.id === activeDetailCard.id
                            ? {
                                ...c,
                                screenshots: {
                                  authorsNote: [...existing, ...newScreenshots],
                                  favoriteScenes: c.screenshots?.favoriteScenes || [],
                                },
                              }
                            : c
                        );
                        updateAppData({ ...appData, cards: updatedCards });
                        showToast(`已添加 ${files.length} 张“作者的话”截图`, 'success');
                        e.target.value = '';
                      }}
                    />

                    {/* Screenshots Grid for Author's Notes */}
                    {(!activeDetailCard.screenshots?.authorsNote || activeDetailCard.screenshots.authorsNote.length === 0) ? (
                      <p className="text-xs text-zinc-400 italic">暂无“作者的话”截图</p>
                    ) : (
                      <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 gap-3">
                        {activeDetailCard.screenshots.authorsNote.map((imgSrc, idx) => {
                          const isSelected = selectedAuthorNoteIndices.includes(idx);
                          return (
                            <div
                              key={idx}
                              onClick={() => {
                                if (authorNoteBatchMode) {
                                  if (isSelected) setSelectedAuthorNoteIndices((p) => p.filter((i) => i !== idx));
                                  else setSelectedAuthorNoteIndices((p) => [...p, idx]);
                                } else {
                                  setFullscreenData({ title: `作者的话截图 #${idx + 1}`, content: imgSrc, type: 'text' });
                                }
                              }}
                              className={`aspect-video rounded-lg overflow-hidden bg-zinc-100 dark:bg-zinc-800 border relative group cursor-pointer transition-all ${
                                isSelected ? 'ring-2 ring-rose-500 border-transparent shadow-md' : 'border-zinc-200 dark:border-zinc-800'
                              }`}
                            >
                              <img src={imgSrc} alt="Author Note" className="w-full h-full object-cover" />

                              {/* Single Delete Button */}
                              {!authorNoteBatchMode && (
                                <button
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    const updatedNotes = activeDetailCard.screenshots!.authorsNote.filter((_, i) => i !== idx);
                                    const updatedCards = appData.cards.map((c) =>
                                      c.id === activeDetailCard.id
                                        ? { ...c, screenshots: { ...c.screenshots, authorsNote: updatedNotes, favoriteScenes: c.screenshots?.favoriteScenes || [] } }
                                        : c
                                    );
                                    updateAppData({ ...appData, cards: updatedCards });
                                    showToast('照片已删除', 'info');
                                  }}
                                  className="absolute top-1 right-1 p-1 rounded-full bg-black/60 text-white opacity-0 group-hover:opacity-100 transition-opacity hover:bg-rose-600"
                                  title="删除照片"
                                >
                                  <X className="w-3 h-3" />
                                </button>
                              )}

                              {/* Batch Selection Checkbox */}
                              {authorNoteBatchMode && (
                                <div className="absolute top-1 left-1 bg-black/40 rounded-full p-0.5">
                                  {isSelected ? (
                                    <CheckCircle2 className="w-5 h-5 text-rose-500 fill-white" />
                                  ) : (
                                    <Circle className="w-5 h-5 text-white/80 drop-shadow" />
                                  )}
                                </div>
                              )}
                            </div>
                          );
                        })}
                      </div>
                    )}
                  </div>

                  {/* Section 2: 回忆 */}
                  <div className="space-y-3 pt-4 border-t border-zinc-200 dark:border-zinc-800">
                    <div className="flex flex-wrap items-center justify-between gap-2">
                      <h3 className="text-xs font-bold text-zinc-800 dark:text-zinc-200">
                        回忆（截图） ({activeDetailCard.screenshots?.favoriteScenes?.length || 0})
                      </h3>
                      <div className="flex flex-wrap items-center gap-2">
                        <button
                          onClick={() => memoryInputRef.current?.click()}
                          className="px-2.5 py-1 text-xs font-medium rounded bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 flex items-center gap-1"
                        >
                          <Upload className="w-3 h-3" /> 上传截图
                        </button>

                        {(activeDetailCard.screenshots?.favoriteScenes?.length || 0) > 0 && (
                          <button
                            onClick={() => {
                              setMemoryBatchMode((p) => !p);
                              setSelectedMemoryIndices([]);
                            }}
                            className={`px-2.5 py-1 text-xs font-medium rounded border transition-colors ${
                              memoryBatchMode
                                ? 'bg-amber-100 text-amber-900 border-amber-300 dark:bg-amber-950 dark:text-amber-200 dark:border-amber-800'
                                : 'border-zinc-200 dark:border-zinc-800 hover:bg-zinc-100 dark:hover:bg-zinc-800'
                            }`}
                          >
                            {memoryBatchMode ? '退出批量' : '批量删除截图'}
                          </button>
                        )}

                        {memoryBatchMode && (activeDetailCard.screenshots?.favoriteScenes?.length || 0) > 0 && (
                          <>
                            <button
                              onClick={() => {
                                const allCount = activeDetailCard.screenshots?.favoriteScenes?.length || 0;
                                if (selectedMemoryIndices.length === allCount) {
                                  setSelectedMemoryIndices([]);
                                } else {
                                  setSelectedMemoryIndices(Array.from({ length: allCount }, (_, i) => i));
                                }
                              }}
                              className="px-2 py-1 text-xs font-medium rounded border border-zinc-200 dark:border-zinc-800 hover:bg-zinc-100 dark:hover:bg-zinc-800"
                            >
                              {selectedMemoryIndices.length === (activeDetailCard.screenshots?.favoriteScenes?.length || 0)
                                ? '取消全选'
                                : '全选'}
                            </button>

                            {selectedMemoryIndices.length > 0 && (
                              <button
                                onClick={() => {
                                  if (!confirm(`确定要删除选中的 ${selectedMemoryIndices.length} 张“回忆”截图吗？`)) return;
                                  const updatedScenes = (activeDetailCard.screenshots?.favoriteScenes || []).filter(
                                    (_, idx) => !selectedMemoryIndices.includes(idx)
                                  );
                                  const updatedCards = appData.cards.map((c) =>
                                    c.id === activeDetailCard.id
                                      ? { ...c, screenshots: { authorsNote: c.screenshots?.authorsNote || [], favoriteScenes: updatedScenes } }
                                      : c
                                  );
                                  updateAppData({ ...appData, cards: updatedCards });
                                  setSelectedMemoryIndices([]);
                                  setMemoryBatchMode(false);
                                  showToast('已批量删除选中的截图', 'info');
                                }}
                                className="px-2.5 py-1 text-xs font-bold rounded bg-rose-600 text-white hover:bg-rose-700 flex items-center gap-1"
                              >
                                <Trash2 className="w-3 h-3" /> 删除选中 ({selectedMemoryIndices.length})
                              </button>
                            )}
                          </>
                        )}
                      </div>
                    </div>

                    <input
                      type="file"
                      ref={memoryInputRef}
                      multiple
                      accept="image/*"
                      className="hidden"
                      onChange={async (e) => {
                        const files = Array.from(e.target.files || []) as File[];
                        if (files.length === 0) return;
                        const newScreenshots: string[] = [];
                        for (const f of files) {
                          const scaled = await processImageFile(f, 800, 800);
                          newScreenshots.push(scaled);
                          await new Promise((resolve) => setTimeout(resolve, 0));
                        }
                        const existing = activeDetailCard.screenshots?.favoriteScenes || [];
                        const updatedCards = appData.cards.map((c) =>
                          c.id === activeDetailCard.id
                            ? {
                                ...c,
                                screenshots: {
                                  authorsNote: c.screenshots?.authorsNote || [],
                                  favoriteScenes: [...existing, ...newScreenshots],
                                },
                              }
                            : c
                        );
                        updateAppData({ ...appData, cards: updatedCards });
                        showToast(`已添加 ${files.length} 张“回忆”截图`, 'success');
                        e.target.value = '';
                      }}
                    />

                    {/* Screenshots Grid for Memories */}
                    {(!activeDetailCard.screenshots?.favoriteScenes || activeDetailCard.screenshots.favoriteScenes.length === 0) ? (
                      <p className="text-xs text-zinc-400 italic">暂无“回忆”截图</p>
                    ) : (
                      <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 gap-3">
                        {activeDetailCard.screenshots.favoriteScenes.map((imgSrc, idx) => {
                          const isSelected = selectedMemoryIndices.includes(idx);
                          return (
                            <div
                              key={idx}
                              onClick={() => {
                                if (memoryBatchMode) {
                                  if (isSelected) setSelectedMemoryIndices((p) => p.filter((i) => i !== idx));
                                  else setSelectedMemoryIndices((p) => [...p, idx]);
                                } else {
                                  setFullscreenData({ title: `回忆截图 #${idx + 1}`, content: imgSrc, type: 'text' });
                                }
                              }}
                              className={`aspect-video rounded-lg overflow-hidden bg-zinc-100 dark:bg-zinc-800 border relative group cursor-pointer transition-all ${
                                isSelected ? 'ring-2 ring-rose-500 border-transparent shadow-md' : 'border-zinc-200 dark:border-zinc-800'
                              }`}
                            >
                              <img src={imgSrc} alt="Memory Scene" className="w-full h-full object-cover" />

                              {/* Single Delete Button */}
                              {!memoryBatchMode && (
                                <button
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    const updatedScenes = activeDetailCard.screenshots!.favoriteScenes.filter((_, i) => i !== idx);
                                    const updatedCards = appData.cards.map((c) =>
                                      c.id === activeDetailCard.id
                                        ? { ...c, screenshots: { authorsNote: c.screenshots?.authorsNote || [], favoriteScenes: updatedScenes } }
                                        : c
                                    );
                                    updateAppData({ ...appData, cards: updatedCards });
                                    showToast('照片已删除', 'info');
                                  }}
                                  className="absolute top-1 right-1 p-1 rounded-full bg-black/60 text-white opacity-0 group-hover:opacity-100 transition-opacity hover:bg-rose-600"
                                  title="删除照片"
                                >
                                  <X className="w-3 h-3" />
                                </button>
                              )}

                              {/* Batch Selection Checkbox */}
                              {memoryBatchMode && (
                                <div className="absolute top-1 left-1 bg-black/40 rounded-full p-0.5">
                                  {isSelected ? (
                                    <CheckCircle2 className="w-5 h-5 text-rose-500 fill-white" />
                                  ) : (
                                    <Circle className="w-5 h-5 text-white/80 drop-shadow" />
                                  )}
                                </div>
                              )}
                            </div>
                          );
                        })}
                      </div>
                    )}
                  </div>

                  {/* Section 3: 关联角色卡 */}
                  <div className="space-y-3 pt-4 border-t border-zinc-200 dark:border-zinc-800">
                    <div className="flex items-center justify-between gap-2">
                      <div>
                        <h3 className="text-xs font-bold text-zinc-800 dark:text-zinc-200 flex items-center gap-1.5">
                          <Link2 className="w-3.5 h-3.5" /> 关联角色卡
                        </h3>
                        <p className="text-[10px] text-zinc-400 mt-0.5">仅显示“角色名 + 作者”都完全相同的角色卡</p>
                      </div>
                    </div>

                    {getAssociationCandidates(activeDetailCard).length > 0 ? (
                      <div className="p-3 rounded-xl bg-zinc-50 dark:bg-zinc-800/50 border border-zinc-200 dark:border-zinc-800 space-y-3">
                        <select
                          value={associationTargetId}
                          onChange={(e) => setAssociationTargetId(e.target.value)}
                          className="w-full px-3 py-2 text-xs bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-700 rounded-lg text-zinc-800 dark:text-zinc-200 focus:outline-none"
                        >
                          <option value="">选择要关联的角色卡…</option>
                          {getAssociationCandidates(activeDetailCard).map((card) => (
                            <option key={card.id} value={card.id}>
                              {getCardDisplayName(card)} · 作者：{getCardCreator(card)}
                            </option>
                          ))}
                        </select>

                        <div className="flex flex-col sm:flex-row gap-2">
                          <input
                            type="text"
                            value={associationNote}
                            onChange={(e) => setAssociationNote(e.target.value)}
                            placeholder="描述 / 备注（可选）"
                            className="flex-1 px-3 py-2 text-xs bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-700 rounded-lg text-zinc-800 dark:text-zinc-200 focus:outline-none"
                          />
                          <label className="flex items-center gap-1.5 px-2.5 py-2 text-xs text-zinc-700 dark:text-zinc-300 whitespace-nowrap cursor-pointer">
                            <input
                              type="checkbox"
                              checked={associationPrimary}
                              onChange={(e) => setAssociationPrimary(e.target.checked)}
                              className="w-3.5 h-3.5"
                            />
                            将关联卡设为主卡
                          </label>
                          <button
                            type="button"
                            onClick={handleAddCardAssociation}
                            className="px-3 py-2 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:opacity-90 whitespace-nowrap"
                          >
                            添加关联
                          </button>
                        </div>
                      </div>
                    ) : (
                      <p className="text-xs text-zinc-400 italic">没有找到作者和角色名都相同、且尚未关联的角色卡。</p>
                    )}

                    {(activeDetailCard.associations || []).length > 0 ? (
                      <div className="space-y-2">
                        {(activeDetailCard.associations || []).map((association) => {
                          const relatedCard = appData.cards.find((c) => c.id === association.cardId);
                          if (!relatedCard) return null;
                          return (
                            <div key={association.cardId} className="flex items-start gap-3 p-3 rounded-xl border border-zinc-200 dark:border-zinc-800 bg-white dark:bg-zinc-900">
                              <div className="w-9 h-12 rounded-lg overflow-hidden bg-zinc-100 dark:bg-zinc-800 flex-shrink-0">
                                {relatedCard.coverImage ? (
                                  <img src={relatedCard.coverImage} alt="" className="w-full h-full object-contain" />
                                ) : (
                                  <div className="w-full h-full flex items-center justify-center text-zinc-400"><FileText className="w-4 h-4" /></div>
                                )}
                              </div>
                              <div className="flex-1 min-w-0">
                                <div className="flex items-center gap-1.5 flex-wrap">
                                  <span className="text-xs font-bold text-zinc-900 dark:text-zinc-100 truncate">{getCardDisplayName(relatedCard)}</span>
                                  {association.isPrimary && (
                                    <span className="px-1.5 py-0.5 rounded text-[10px] font-semibold bg-amber-100 text-amber-800 dark:bg-amber-950/60 dark:text-amber-200">主卡</span>
                                  )}
                                </div>
                                <div className="text-[10px] text-zinc-400 mt-0.5 truncate">作者：{getCardCreator(relatedCard)}</div>
                                {association.note && association.noteOwnerId === activeDetailCard.id && (
                                  <div className="text-[11px] text-zinc-600 dark:text-zinc-300 mt-1.5 leading-relaxed">备注：{association.note}</div>
                                )}
                              </div>
                              <div className="flex items-center gap-1 flex-shrink-0">
                                <button
                                  type="button"
                                  onClick={() => openAssociatedCard(relatedCard.id)}
                                  className="px-2 py-1 text-[10px] font-semibold rounded-lg border border-zinc-200 dark:border-zinc-700 hover:bg-zinc-100 dark:hover:bg-zinc-800 text-zinc-600 dark:text-zinc-300"
                                  title="跳转到此角色卡"
                                >跳转</button>
                                <button
                                  type="button"
                                  onClick={() => handleRemoveCardAssociation(relatedCard.id)}
                                  className="p-1 text-zinc-400 hover:text-rose-500"
                                  title="解除关联"
                                >
                                  <X className="w-3.5 h-3.5" />
                                </button>
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    ) : (
                      <p className="text-xs text-zinc-400 italic">暂无已关联角色卡。</p>
                    )}
                  </div>
                </div>
              )}

              {/* TAB 5: QR (JSON) */}
              {detailTab === 'qr' && (() => {
                const qrDoc = activeDetailCard.qrData && typeof activeDetailCard.qrData === 'object'
                  ? normalizeQrDocument(activeDetailCard.qrData)
                  : { version: 2, name: '', qrList: [], idIndex: 0 };
                const visibleQr = (qrDoc.qrList || []).filter((item: any) => {
                  if (!qrSearchQuery.trim()) return true;
                  const q = qrSearchQuery.trim().toLowerCase();
                  return `${qrDoc.name || ''} ${item.label || ''} ${item.title || ''} ${item.message || ''}`.toLowerCase().includes(q);
                });
                return (
                  <div className="space-y-4">
                    <div className="space-y-2">
                      <div className="relative w-full">
                        <Search className="w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 text-zinc-400" />
                        <input
                          type="text"
                          value={qrSearchQuery}
                          onChange={(e) => setQrSearchQuery(e.target.value)}
                          placeholder="搜索 QR 总名称..."
                          className="w-full pl-9 pr-4 py-1.5 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400 dark:focus:ring-zinc-500"
                        />
                      </div>
                      <div className="flex items-center justify-start gap-2 flex-wrap">
                        <button type="button" onClick={() => setCardSectionImportModal('qr')} className="px-3 py-2 text-xs font-medium rounded-xl border border-zinc-200 dark:border-zinc-700 hover:bg-zinc-100 dark:hover:bg-zinc-800 flex items-center gap-1">
                          <Upload className="w-3.5 h-3.5" /> 导入
                        </button>
                        <button type="button" onClick={addManualQr} className="px-3 py-2 text-xs font-bold rounded-xl bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:opacity-90 flex items-center gap-1">
                          <Plus className="w-3.5 h-3.5" /> 增加 QR
                        </button>
                        <button type="button" onClick={() => downloadJsonFile(`${qrDoc.name || getCardDisplayName(activeDetailCard)}_qr.json`, qrDoc)} className="px-3 py-2 text-xs font-medium rounded-xl border border-zinc-200 dark:border-zinc-700 hover:bg-zinc-100 dark:hover:bg-zinc-800 flex items-center gap-1">
                          <Download className="w-3.5 h-3.5" /> 导出 JSON
                        </button>
                      </div>
                    </div>

                    {qrDoc.name && (
                      <div className="px-3 py-2 text-xs text-zinc-500 dark:text-zinc-400 bg-zinc-50 dark:bg-zinc-800/50 rounded-xl border border-zinc-200 dark:border-zinc-800">QR 总名称：<span className="font-semibold text-zinc-800 dark:text-zinc-200">{qrDoc.name}</span></div>
                    )}

                    {visibleQr.length > 0 ? (
                      <div className="space-y-3">
                        {visibleQr.map((item: any, i: number) => (
                          <div key={`${item.id ?? i}-${i}`} className="p-3.5 bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-xl">
                            <div className="flex items-center justify-between gap-2 mb-2">
                              <div className="min-w-0">
                                <div className="text-xs font-bold text-zinc-900 dark:text-zinc-100">{item.label || item.title || `QR #${i + 1}`}</div>
                                <div className="text-[10px] text-zinc-400 mt-0.5">顺序 {i + 1} · ID {item.id ?? '-'}</div>
                              </div>
                              <div className="flex items-center gap-1.5 flex-shrink-0">
                                <span className="text-[10px] px-1.5 py-0.5 rounded bg-zinc-200 dark:bg-zinc-700 text-zinc-500 dark:text-zinc-300">{item.preventAutoExecute ? '手动' : '自动'}</span>
                                <button type="button" onClick={() => setEditingQrItem({ index: qrDoc.qrList.indexOf(item), item: JSON.parse(JSON.stringify(item)) })} className="px-2 py-1 text-[10px] font-semibold rounded-lg border border-zinc-200 dark:border-zinc-700 hover:bg-zinc-100 dark:hover:bg-zinc-800">编辑</button>
                                <button type="button" onClick={() => deleteCardQrItem(qrDoc.qrList.indexOf(item))} className="px-2 py-1 text-[10px] font-semibold rounded-lg border border-red-200 text-red-600 dark:border-red-900/60 dark:text-red-400 hover:bg-red-50 dark:hover:bg-red-950/30">删除</button>
                              </div>
                            </div>
                            <div className="text-[11px] text-zinc-500 dark:text-zinc-400 mb-2">消息：{item.message || '（空）'}</div>
                            <pre className="p-2.5 bg-zinc-900 text-zinc-100 text-[10px] font-mono rounded-lg overflow-x-auto max-h-44 leading-relaxed">{JSON.stringify(item, null, 2)}</pre>
                          </div>
                        ))}
                      </div>
                    ) : (
                      <div className="text-center py-10 text-xs text-zinc-400 border border-dashed border-zinc-200 dark:border-zinc-800 rounded-2xl">暂无 QR 条目</div>
                    )}
                  </div>
                );
              })()}
            </div>

            {/* Modal Bottom Sticky Footer Bar (Requirement 11) */}
            <div className="px-6 py-3 border-t border-zinc-200 dark:border-zinc-800 flex items-center justify-between flex-shrink-0 bg-zinc-50/50 dark:bg-zinc-900/50">
              <button
                type="button"
                onClick={() => {
                  if (confirm(`确定要删除角色卡 “${getCardDisplayName(activeDetailCard)}” 吗？`)) {
                    const updatedCards = appData.cards.filter((c) => c.id !== activeDetailCard.id);
                    updateAppData({ ...appData, cards: updatedCards });
                    setDetailCardId(null);
                    showToast('已彻底删除该角色卡', 'info');
                  }
                }}
                className="px-4 py-2 text-xs font-bold rounded-xl bg-rose-600 text-white hover:bg-rose-700 flex items-center gap-1 shadow-sm"
              >
                <Trash2 className="w-3.5 h-3.5" /> 删除角色卡
              </button>

              <div className="flex items-center gap-2.5">
                <button
                  type="button"
                  onClick={() => setDetailCardId(null)}
                  className="px-4 py-2 text-xs font-medium rounded-xl border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800"
                >
                  取消
                </button>
                <button
                  type="button"
                  onClick={() => {
                    const authorInputEl = document.getElementById('authorInputVal') as HTMLInputElement;
                    const newAuthor = authorInputEl?.value.trim() || '';
                    const updatedCards = appData.cards.map((c) =>
                      c.id === activeDetailCard.id
                        ? {
                            ...c,
                            author: newAuthor,
                            authorManual: true,
                            editHistory: { ...c.editHistory, author: newAuthor },
                            edited: true,
                          }
                        : c
                    );
                    updateAppData({ ...appData, cards: updatedCards });
                    showToast('角色卡详情与修改已成功保存！', 'success');
                  }}
                  className="px-5 py-2 text-xs font-bold rounded-xl bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:opacity-90 shadow-sm flex items-center gap-1.5"
                >
                  <Save className="w-3.5 h-3.5" /> 保存修改
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ==================== NEW GROUP MODAL (Requirement 4) ==================== */}
      {showNewGroupModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 p-4 animate-in fade-in">
          <div className="bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl w-full max-w-sm p-5 space-y-4 shadow-xl">
            <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100">新建分组</h3>
            <div>
              <label className="text-xs font-semibold text-zinc-500 mb-1 block">分组名字</label>
              <input
                type="text"
                autoFocus
                value={newGroupName}
                onChange={(e) => setNewGroupName(e.target.value)}
                placeholder="请输入分组名称..."
                onKeyDown={(e) => {
                  if (e.key === 'Enter') handleCreateNewGroup();
                }}
                className="w-full px-3 py-2 text-sm bg-zinc-50 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400"
              />
            </div>
            <div className="flex justify-end gap-2 pt-2">
              <button
                onClick={() => setShowNewGroupModal(false)}
                className="px-3 py-1.5 text-xs font-medium text-zinc-600 dark:text-zinc-400 hover:bg-zinc-100 dark:hover:bg-zinc-800 rounded-lg"
              >
                取消
              </button>
              <button
                onClick={handleCreateNewGroup}
                className="px-4 py-1.5 text-xs font-bold bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 rounded-lg hover:opacity-90"
              >
                确定
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================== MANAGE / DELETE GROUP MODAL (Requirement 4) ==================== */}
      {managingGroup && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 p-4 animate-in fade-in">
          <div className="bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl w-full max-w-md p-5 space-y-4 shadow-xl">
            <div className="flex items-center justify-between border-b border-zinc-200 dark:border-zinc-800 pb-3">
              <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100">
                管理分组: {managingGroup}
              </h3>
              <button onClick={() => setManagingGroup(null)} className="p-1 rounded text-zinc-400 hover:text-zinc-600">
                <X className="w-4 h-4" />
              </button>
            </div>

            {/* Rename Group Section */}
            <div className="space-y-2">
              <label className="text-xs font-semibold text-zinc-500 block">重命名分组</label>
              <div className="flex gap-2">
                <input
                  type="text"
                  value={renameGroupInput}
                  onChange={(e) => setRenameGroupInput(e.target.value)}
                  className="flex-1 px-3 py-1.5 text-xs bg-zinc-50 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 rounded-lg"
                />
                <button
                  onClick={handleRenameGroup}
                  className="px-3 py-1.5 text-xs font-medium bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 rounded-lg"
                >
                  保存名字
                </button>
              </div>
            </div>

            {/* Delete Group Confirmation Dialog with Requirement 4 Circular Toggle */}
            <div className="pt-3 border-t border-zinc-200 dark:border-zinc-800 space-y-3">
              <div className="text-xs font-bold text-rose-600 dark:text-rose-400">删除分组</div>
              <p className="text-xs text-zinc-500 dark:text-zinc-400 leading-relaxed">
                是否要删除分组“{managingGroup}”？删除前请选择是否一并清理组内角色卡。
              </p>

              {/* Circular Selection Toggle (Requirement 4: "搞成那种圆圈，别人点击圆圈 圆圈内部变黑就是会一起删掉，如果没选就不删") */}
              <div
                onClick={() => setDeleteCardsWithGroup((prev) => !prev)}
                className="flex items-center gap-3 p-3 bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-xl cursor-pointer hover:bg-zinc-100 dark:hover:bg-zinc-800 transition-colors"
              >
                {/* Circular Indicator */}
                <div
                  className={`w-5 h-5 rounded-full border-2 flex items-center justify-center transition-all ${
                    deleteCardsWithGroup
                      ? 'border-zinc-900 dark:border-zinc-100 bg-white dark:bg-zinc-900'
                      : 'border-zinc-400 dark:border-zinc-600 bg-transparent'
                  }`}
                >
                  {deleteCardsWithGroup && (
                    /* Circle center turns dark/black when selected */
                    <div className="w-2.5 h-2.5 rounded-full bg-zinc-900 dark:bg-zinc-100" />
                  )}
                </div>

                <div className="text-xs font-medium text-zinc-800 dark:text-zinc-200">
                  是否要一并把分组里的角色卡给删掉？
                </div>
              </div>

              <div className="flex justify-end gap-2 pt-2">
                <button
                  onClick={() => setManagingGroup(null)}
                  className="px-3 py-1.5 text-xs font-medium text-zinc-600 dark:text-zinc-400 hover:bg-zinc-100 dark:hover:bg-zinc-800 rounded-lg"
                >
                  取消
                </button>
                <button
                  onClick={handleConfirmDeleteGroup}
                  className="px-4 py-1.5 text-xs font-bold bg-rose-600 text-white hover:bg-rose-700 rounded-lg"
                >
                  确认删除分组
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ==================== BATCH MOVE MODAL ==================== */}
      {showBatchMoveModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 p-4 animate-in fade-in">
          <div className="bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl w-full max-w-sm p-5 space-y-4 shadow-xl">
            <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100">
              批量移动到分组 ({selectedCardIds.length} 张)
            </h3>
            <div>
              <label className="text-xs font-semibold text-zinc-500 mb-1 block">选择目标分组</label>
              <select
                value={batchTargetGroup}
                onChange={(e) => setBatchTargetGroup(e.target.value)}
                className="w-full px-3 py-2 text-sm bg-zinc-50 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 rounded-lg"
              >
                <option value="">-- 请选择分组 --</option>
                {appData.groups.map((g) => (
                  <option key={g} value={g}>
                    {g}
                  </option>
                ))}
              </select>
            </div>
            <div className="flex justify-end gap-2 pt-2">
              <button
                onClick={() => setShowBatchMoveModal(false)}
                className="px-3 py-1.5 text-xs font-medium text-zinc-600 dark:text-zinc-400 hover:bg-zinc-100 dark:hover:bg-zinc-800 rounded-lg"
              >
                取消
              </button>
              <button
                disabled={!batchTargetGroup}
                onClick={handleBatchMoveCards}
                className="px-4 py-1.5 text-xs font-bold bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 rounded-lg hover:opacity-90 disabled:opacity-40"
              >
                确认移动
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================== NEW THEME GROUP MODAL ==================== */}
      {showNewThemeGroupModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 p-4 animate-in fade-in">
          <div className="bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl w-full max-w-sm p-5 space-y-4 shadow-xl">
            <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100">新建主题美化分组</h3>
            <div>
              <label className="text-xs font-semibold text-zinc-500 mb-1 block">分组名字</label>
              <input
                type="text"
                autoFocus
                value={newThemeGroupName}
                onChange={(e) => setNewThemeGroupName(e.target.value)}
                placeholder="请输入分组名称..."
                onKeyDown={(e) => {
                  if (e.key === 'Enter') handleCreateNewThemeGroup();
                }}
                className="w-full px-3 py-2 text-sm bg-zinc-50 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400"
              />
            </div>
            <div className="flex justify-end gap-2 pt-2">
              <button
                onClick={() => {
                  setShowNewThemeGroupModal(false);
                  setNewThemeGroupName('');
                }}
                className="px-3 py-1.5 text-xs font-medium text-zinc-600 dark:text-zinc-400 hover:bg-zinc-100 dark:hover:bg-zinc-800 rounded-lg"
              >
                取消
              </button>
              <button
                onClick={handleCreateNewThemeGroup}
                className="px-4 py-1.5 text-xs font-bold bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 rounded-lg hover:opacity-90"
              >
                确定
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================== MANAGE / DELETE THEME GROUP MODAL ==================== */}
      {managingThemeGroup && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 p-4 animate-in fade-in">
          <div className="bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl w-full max-w-md p-5 space-y-4 shadow-xl">
            <div className="flex items-center justify-between border-b border-zinc-200 dark:border-zinc-800 pb-3">
              <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100">
                管理主题分组: {managingThemeGroup}
              </h3>
              <button onClick={() => setManagingThemeGroup(null)} className="p-1 rounded text-zinc-400 hover:text-zinc-600">
                <X className="w-4 h-4" />
              </button>
            </div>

            {/* Rename Group Section */}
            <div className="space-y-2">
              <label className="text-xs font-semibold text-zinc-500 block">重命名分组</label>
              <div className="flex gap-2">
                <input
                  type="text"
                  value={renameThemeGroupInput}
                  onChange={(e) => setRenameThemeGroupInput(e.target.value)}
                  className="flex-1 px-3 py-1.5 text-xs bg-zinc-50 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400"
                />
                <button
                  onClick={handleRenameThemeGroup}
                  className="px-3 py-1.5 text-xs font-medium bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 rounded-lg whitespace-nowrap"
                >
                  保存名字
                </button>
              </div>
            </div>

            {/* Delete Group Confirmation Dialog with Circular Toggle */}
            <div className="pt-3 border-t border-zinc-200 dark:border-zinc-800 space-y-3">
              <div className="text-xs font-bold text-rose-600 dark:text-rose-400">删除分组</div>
              <p className="text-xs text-zinc-500 dark:text-zinc-400 leading-relaxed">
                是否要删除分组“{managingThemeGroup}”？删除前请选择是否一并清理组内主题美化。
              </p>

              {/* Circular Selection Toggle */}
              <div
                onClick={() => setDeleteThemesWithGroup((prev) => !prev)}
                className="flex items-center gap-3 p-3 bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-xl cursor-pointer hover:bg-zinc-100 dark:hover:bg-zinc-800 transition-colors"
              >
                {/* Circular Indicator */}
                <div
                  className={`w-5 h-5 rounded-full border-2 flex items-center justify-center transition-all ${
                    deleteThemesWithGroup
                      ? 'border-zinc-900 dark:border-zinc-100 bg-white dark:bg-zinc-900'
                      : 'border-zinc-400 dark:border-zinc-600 bg-transparent'
                  }`}
                >
                  {deleteThemesWithGroup && (
                    /* Circle center turns dark/black when selected */
                    <div className="w-2.5 h-2.5 rounded-full bg-zinc-900 dark:bg-zinc-100" />
                  )}
                </div>

                <div className="text-xs font-medium text-zinc-800 dark:text-zinc-200">
                  是否要一并把分组里的主题美化给删掉？
                </div>
              </div>

              <div className="flex justify-end gap-2 pt-2">
                <button
                  onClick={() => setManagingThemeGroup(null)}
                  className="px-3 py-1.5 text-xs font-medium text-zinc-600 dark:text-zinc-400 hover:bg-zinc-100 dark:hover:bg-zinc-800 rounded-lg"
                >
                  取消
                </button>
                <button
                  onClick={handleConfirmDeleteThemeGroup}
                  className="px-4 py-1.5 text-xs font-bold bg-rose-600 text-white hover:bg-rose-700 rounded-lg"
                >
                  确认删除分组
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ==================== BATCH MOVE THEME MODAL ==================== */}
      {showThemeBatchMoveModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 p-4 animate-in fade-in">
          <div className="bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl w-full max-w-sm p-5 space-y-4 shadow-xl">
            <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100">
              批量移动主题美化 ({selectedThemeIds.length} 个)
            </h3>
            <div className="space-y-2">
              <label className="text-xs font-semibold text-zinc-500 mb-1 block">选择或输入目标分组</label>
              <select
                value={batchTargetThemeGroup}
                onChange={(e) => setBatchTargetThemeGroup(e.target.value)}
                className="w-full px-3 py-2 text-xs bg-zinc-50 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none"
              >
                <option value="">-- 请选择现有分组 --</option>
                {Array.from(new Set(['默认', ...(appData.themeCategories || [])])).map((cat) => (
                  <option key={cat} value={cat}>
                    {cat}
                  </option>
                ))}
              </select>
              <input
                type="text"
                value={batchTargetThemeGroup}
                onChange={(e) => setBatchTargetThemeGroup(e.target.value)}
                placeholder="或直接输入新的分组名称..."
                className="w-full px-3 py-2 text-xs bg-zinc-50 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400"
              />
            </div>
            <div className="flex justify-end gap-2 pt-2">
              <button
                onClick={() => {
                  setShowThemeBatchMoveModal(false);
                  setBatchTargetThemeGroup('');
                }}
                className="px-3 py-1.5 text-xs font-medium text-zinc-600 dark:text-zinc-400 hover:bg-zinc-100 dark:hover:bg-zinc-800 rounded-lg"
              >
                取消
              </button>
              <button
                disabled={!batchTargetThemeGroup.trim()}
                onClick={handleBatchMoveThemes}
                className="px-4 py-1.5 text-xs font-bold bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 rounded-lg hover:opacity-90 disabled:opacity-40"
              >
                确认移动
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================== ADD PHONE LINK MODAL ==================== */}
      {showAddPhoneModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 animate-in fade-in">
          <div className="w-full max-w-lg bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl p-6 space-y-4">
            <div className="flex items-center justify-between border-b border-zinc-100 dark:border-zinc-800 pb-3">
              <h3 className="text-base font-bold text-zinc-900 dark:text-zinc-100">添加小手机链接</h3>
              <button
                onClick={() => setShowAddPhoneModal(false)}
                className="p-1 rounded-lg text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-200 hover:bg-zinc-100 dark:hover:bg-zinc-800"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="space-y-3.5">
              {/* 小手机名称 (必填) */}
              <div>
                <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                  小手机名称 <span className="text-rose-500">*</span>
                </label>
                <input
                  type="text"
                  value={phoneForm.name}
                  onChange={(e) => setPhoneForm((p) => ({ ...p, name: e.target.value }))}
                  placeholder="请输入小手机名称..."
                  className="w-full px-3 py-2 text-sm bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400 dark:focus:ring-zinc-500 text-zinc-900 dark:text-zinc-100"
                />
              </div>

              {/* 填写链接 (必填) */}
              <div>
                <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                  填写链接 <span className="text-rose-500">*</span>
                </label>
                <input
                  type="text"
                  value={phoneForm.url}
                  onChange={(e) => setPhoneForm((p) => ({ ...p, url: e.target.value }))}
                  placeholder="请输入链接地址..."
                  className="w-full px-3 py-2 text-sm bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400 dark:focus:ring-zinc-500 text-zinc-900 dark:text-zinc-100"
                />
              </div>

              {/* 联系方式 (选填) */}
              <div>
                <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                  联系方式（小红书，qq群等等）
                </label>
                <input
                  type="text"
                  value={phoneForm.contact}
                  onChange={(e) => setPhoneForm((p) => ({ ...p, contact: e.target.value }))}
                  placeholder="选填，例如: 小红书: xxx, QQ群: 123456"
                  className="w-full px-3 py-2 text-sm bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400 dark:focus:ring-zinc-500 text-zinc-900 dark:text-zinc-100"
                />
              </div>

              {/* 描述 (选填) */}
              <div>
                <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                  描述
                </label>
                <textarea
                  rows={3}
                  value={phoneForm.description}
                  onChange={(e) => setPhoneForm((p) => ({ ...p, description: e.target.value }))}
                  placeholder="选填，请输入备注或描述信息..."
                  className="w-full px-3 py-2 text-sm bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400 dark:focus:ring-zinc-500 text-zinc-900 dark:text-zinc-100 resize-none"
                />
              </div>
            </div>

            {/* Bottom Action Buttons */}
            <div className="flex items-center justify-end gap-2.5 pt-2 border-t border-zinc-100 dark:border-zinc-800">
              <button
                onClick={() => setShowAddPhoneModal(false)}
                className="px-4 py-2 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800"
              >
                取消
              </button>
              <button
                onClick={handleAddPhoneLink}
                className="px-4 py-2 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:bg-zinc-800 dark:hover:bg-zinc-200"
              >
                确定
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================== EDIT PHONE LINK MODAL ==================== */}
      {editingPhoneLink && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 animate-in fade-in">
          <div className="w-full max-w-lg bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl p-6 space-y-4">
            <div className="flex items-center justify-between border-b border-zinc-100 dark:border-zinc-800 pb-3">
              <h3 className="text-base font-bold text-zinc-900 dark:text-zinc-100">查看 / 编辑小手机链接</h3>
              <button
                onClick={() => setEditingPhoneLink(null)}
                className="p-1 rounded-lg text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-200 hover:bg-zinc-100 dark:hover:bg-zinc-800"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="space-y-3.5">
              {/* 小手机名称 */}
              <div>
                <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                  小手机名称 <span className="text-rose-500">*</span>
                </label>
                <input
                  type="text"
                  value={editingPhoneLink.name}
                  onChange={(e) => setEditingPhoneLink((p) => p ? { ...p, name: e.target.value } : null)}
                  placeholder="请输入小手机名称..."
                  className="w-full px-3 py-2 text-sm bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400 dark:focus:ring-zinc-500 text-zinc-900 dark:text-zinc-100"
                />
              </div>

              {/* 链接 (带有复制键) */}
              <div>
                <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                  链接 <span className="text-rose-500">*</span>
                </label>
                <div className="flex items-center gap-2">
                  <input
                    type="text"
                    value={editingPhoneLink.url}
                    onChange={(e) => setEditingPhoneLink((p) => p ? { ...p, url: e.target.value } : null)}
                    placeholder="请输入链接地址..."
                    className="flex-1 px-3 py-2 text-sm bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400 dark:focus:ring-zinc-500 text-zinc-900 dark:text-zinc-100"
                  />
                  <button
                    onClick={() => {
                      if (editingPhoneLink.url) {
                        navigator.clipboard.writeText(editingPhoneLink.url);
                        showToast('链接已复制到剪贴板！', 'success');
                      }
                    }}
                    className="p-2 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800 flex items-center gap-1 flex-shrink-0"
                    title="复制完整链接"
                  >
                    <Copy className="w-4 h-4" />
                    <span>复制</span>
                  </button>
                </div>
              </div>

              {/* 联系方式 */}
              <div>
                <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                  联系方式（小红书，qq群等等）
                </label>
                <input
                  type="text"
                  value={editingPhoneLink.contact || ''}
                  onChange={(e) => setEditingPhoneLink((p) => p ? { ...p, contact: e.target.value } : null)}
                  placeholder="选填，例如: 小红书: xxx, QQ群: 123456"
                  className="w-full px-3 py-2 text-sm bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400 dark:focus:ring-zinc-500 text-zinc-900 dark:text-zinc-100"
                />
              </div>

              {/* 描述 */}
              <div>
                <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                  描述
                </label>
                <textarea
                  rows={3}
                  value={editingPhoneLink.description || ''}
                  onChange={(e) => setEditingPhoneLink((p) => p ? { ...p, description: e.target.value } : null)}
                  placeholder="选填，请输入备注或描述信息..."
                  className="w-full px-3 py-2 text-sm bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400 dark:focus:ring-zinc-500 text-zinc-900 dark:text-zinc-100 resize-none"
                />
              </div>
            </div>

            {/* Bottom Action Buttons: 删除, 取消, 确定 */}
            <div className="flex items-center justify-between pt-2 border-t border-zinc-100 dark:border-zinc-800">
              <button
                onClick={() => handleDeleteSinglePhoneLink(editingPhoneLink.id)}
                className="px-3.5 py-2 text-xs font-bold rounded-lg bg-rose-600 text-white hover:bg-rose-700 flex items-center gap-1"
              >
                <Trash2 className="w-3.5 h-3.5" /> 删除
              </button>

              <div className="flex items-center gap-2.5">
                <button
                  onClick={() => setEditingPhoneLink(null)}
                  className="px-4 py-2 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800"
                >
                  取消
                </button>
                <button
                  onClick={handleSaveEditedPhoneLink}
                  className="px-4 py-2 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:bg-zinc-800 dark:hover:bg-zinc-200"
                >
                  确定
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ==================== EDIT THEME DETAIL MODAL (美化弹窗) ==================== */}
      {editingTheme && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-3 sm:p-4 pt-[calc(1rem+env(safe-area-inset-top,0px))] pb-[calc(1rem+env(safe-area-inset-bottom,0px))] animate-in fade-in">
          <div className="w-full max-w-4xl h-[85vh] min-h-[550px] max-h-[750px] bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl flex flex-col overflow-hidden transition-none">
            {/* Hidden Input for Cover Image */}
            <input
              type="file"
              ref={themeCoverInputRef}
              accept="image/*"
              className="hidden"
              onChange={(e) => {
                const file = e.target.files?.[0];
                if (file) {
                  const reader = new FileReader();
                  reader.onload = () => {
                    if (reader.result) {
                      setEditingTheme((p) => p ? { ...p, coverImage: reader.result as string } : null);
                    }
                  };
                  reader.readAsDataURL(file);
                }
                e.target.value = '';
              }}
            />

            {/* Modal Header */}
            <div className="px-6 py-4 border-b border-zinc-200 dark:border-zinc-800 flex items-center justify-between flex-shrink-0 bg-zinc-50/50 dark:bg-zinc-900/50">
              <div className="flex items-center gap-2 min-w-0">
                <h2 className="text-base font-bold truncate text-zinc-900 dark:text-zinc-100">
                  {editingTheme.name || 'ST主题详情'}
                </h2>
                <span className="px-2 py-0.5 text-[10px] font-bold uppercase rounded bg-zinc-200 dark:bg-zinc-800 text-zinc-700 dark:text-zinc-300">
                  {editingTheme.fileType?.toUpperCase() || 'ST主题'}
                </span>
              </div>
              <button
                type="button"
                onClick={() => setEditingTheme(null)}
                className="p-1.5 rounded-lg hover:bg-zinc-200 dark:hover:bg-zinc-800 text-zinc-500 transition-colors"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Modal Navigation Tabs */}
            <div className="flex border-b border-zinc-200 dark:border-zinc-800 px-6 gap-2 flex-shrink-0 bg-white dark:bg-zinc-900">
              <button
                type="button"
                onClick={() => setThemeDetailTab('info')}
                className={`py-3 px-4 text-xs font-bold border-b-2 transition-colors ${
                  themeDetailTab === 'info'
                    ? 'border-zinc-900 text-zinc-900 dark:border-zinc-100 dark:text-zinc-100'
                    : 'border-transparent text-zinc-500 hover:text-zinc-800 dark:hover:text-zinc-200'
                }`}
              >
                基本信息
              </button>
              <button
                type="button"
                onClick={() => setThemeDetailTab('code')}
                className={`py-3 px-4 text-xs font-bold border-b-2 transition-colors ${
                  themeDetailTab === 'code'
                    ? 'border-zinc-900 text-zinc-900 dark:border-zinc-100 dark:text-zinc-100'
                    : 'border-transparent text-zinc-500 hover:text-zinc-800 dark:hover:text-zinc-200'
                }`}
              >
                文档内容
              </button>
            </div>

            {/* Modal Body */}
            <div className="flex-1 overflow-y-auto p-6 space-y-6">
              {/* TAB 1: 美化预览(封面图) 与 基础属性 */}
              {themeDetailTab === 'info' && (
                <div className="bg-zinc-50/50 dark:bg-zinc-800/20 border border-zinc-200 dark:border-zinc-800 rounded-2xl p-4">
                  <h3 className="text-xs font-bold text-zinc-800 dark:text-zinc-200 mb-4 flex items-center gap-1.5">
                    <ImageIcon className="w-4 h-4 text-zinc-500" /> 美化预览与基本信息
                  </h3>
                  <div className="grid grid-cols-1 md:grid-cols-12 gap-6">
                    {/* Left Column: Cover Image Preview & Upload */}
                    <div className="md:col-span-4 space-y-3">
                      <div className="aspect-[2/3] w-full rounded-xl overflow-hidden bg-zinc-100 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-800 relative flex items-center justify-center p-2">
                        {editingTheme.coverImage ? (
                          <img src={editingTheme.coverImage} alt="Theme Cover" className="w-full h-full object-cover" />
                        ) : (
                          <div className="text-center p-4 space-y-2">
                            <ImageIcon className="w-10 h-10 text-zinc-400 mx-auto" />
                            <p className="text-xs text-zinc-400">暂无封面图片</p>
                          </div>
                        )}
                      </div>

                      <button
                        type="button"
                        onClick={() => themeCoverInputRef.current?.click()}
                        className="w-full py-2 text-xs font-medium rounded-lg border border-dashed border-zinc-300 dark:border-zinc-700 hover:border-zinc-500 text-zinc-700 dark:text-zinc-300 transition-colors flex items-center justify-center gap-1.5"
                      >
                        <Upload className="w-3.5 h-3.5" /> 上传图片 (当封面图)
                      </button>
                    </div>

                    {/* Right Column: Fields (名称, 作者, 类型, 来源, 分组, 创建时间单占一行) */}
                    <div className="md:col-span-8 space-y-3">
                      {/* 创建时间 (Requirement 10: 单独占一行) */}
                      <div className="p-2.5 bg-zinc-100 dark:bg-zinc-800/50 border border-zinc-200/80 dark:border-zinc-700/60 rounded-xl text-xs text-zinc-600 dark:text-zinc-300 font-medium">
                        创建时间：{editingTheme.createdAt ? new Date(editingTheme.createdAt).toLocaleString('zh-CN') : '未知时间'}
                      </div>

                      <div>
                        <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                          主题名称 <span className="text-rose-500">*</span>
                        </label>
                        <input
                          type="text"
                          value={editingTheme.name}
                          onChange={(e) => setEditingTheme((p) => p ? { ...p, name: e.target.value } : null)}
                          placeholder="请输入主题名称..."
                          className="w-full px-3 py-2 text-xs bg-white dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400 text-zinc-900 dark:text-zinc-100"
                        />
                      </div>

                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                        <div>
                          <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                            作者
                          </label>
                          <input
                            type="text"
                            value={editingTheme.author || ''}
                            onChange={(e) => setEditingTheme((p) => p ? { ...p, author: e.target.value } : null)}
                            placeholder="作者名字 (默认)..."
                            className="w-full px-3 py-2 text-xs bg-white dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400 text-zinc-900 dark:text-zinc-100"
                          />
                        </div>

                        <div>
                          <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                            来源
                          </label>
                          <input
                            type="text"
                            value={editingTheme.source || ''}
                            onChange={(e) => setEditingTheme((p) => p ? { ...p, source: e.target.value } : null)}
                            placeholder="DC链接 / Q群 / 论坛等..."
                            className="w-full px-3 py-2 text-xs bg-white dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400 text-zinc-900 dark:text-zinc-100"
                          />
                        </div>
                      </div>

                      <div>
                        <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                          所属分组
                        </label>
                        <select
                          value={editingTheme.category || '默认'}
                          onChange={(e) => setEditingTheme((p) => p ? { ...p, category: e.target.value } : null)}
                          className="w-full px-3 py-2 text-xs bg-white dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400 text-zinc-900 dark:text-zinc-100"
                        >
                          {Array.from(new Set(['默认', ...(appData.themeCategories || [])])).map((cat) => (
                            <option key={cat} value={cat}>
                              {cat}
                            </option>
                          ))}
                        </select>
                      </div>

                      <div>
                        <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                          描述 / 备注
                        </label>
                        <textarea
                          rows={3}
                          value={editingTheme.description || ''}
                          onChange={(e) => setEditingTheme((p) => p ? { ...p, description: e.target.value } : null)}
                          placeholder="请输入美化描述或说明..."
                          className="w-full px-3 py-2 text-xs bg-white dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400 text-zinc-900 dark:text-zinc-100 resize-none leading-relaxed"
                        />
                      </div>
                    </div>
                  </div>
                </div>
              )}

              {/* TAB 2: 文档内容 */}
              {themeDetailTab === 'code' && (
                <div className="bg-zinc-50/50 dark:bg-zinc-800/20 border border-zinc-200 dark:border-zinc-800 rounded-2xl p-4 space-y-3">
                  <div className="flex items-center justify-between gap-3">
                    <h3 className="text-xs font-bold text-zinc-800 dark:text-zinc-200 flex items-center gap-1.5 whitespace-nowrap">
                      <FileText className="w-4 h-4 text-zinc-500" /> 文档内容
                    </h3>
                    <button
                      type="button"
                      onClick={() => setIsThemeCodeExpanded(true)}
                      className="px-3 py-1.5 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800 flex items-center gap-1.5 whitespace-nowrap"
                    >
                      <Maximize2 className="w-3.5 h-3.5" /> 放大展开
                    </button>
                  </div>

                  {/* 搜索栏独立放在文档内容界面顶部 */}
                  <div className="relative w-full">
                    <Search className="w-3.5 h-3.5 absolute left-3 top-1/2 -translate-y-1/2 text-zinc-400" />
                    <input
                      type="text"
                      value={codeSearchQuery}
                      onChange={(e) => setCodeSearchQuery(e.target.value)}
                      placeholder="搜索/定位代码内容..."
                      className="w-full pl-9 pr-4 py-1.5 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400 dark:focus:ring-zinc-500"
                    />
                  </div>

                  {/* Document Content Textarea */}
                  <div className="space-y-1">
                    <textarea
                      rows={10}
                      value={editingTheme.content || editingTheme.rawJsonString || ''}
                      onChange={(e) =>
                        setEditingTheme((p) =>
                          p ? { ...p, content: e.target.value, rawJsonString: e.target.value } : null
                        )
                      }
                      placeholder="文档或代码内容..."
                      className="w-full p-3 font-mono text-xs bg-zinc-950 text-zinc-100 rounded-xl border border-zinc-800 focus:outline-none focus:ring-1 focus:ring-zinc-600 leading-relaxed resize-none"
                    />
                    {codeSearchQuery.trim() && (
                      <div className="text-[11px] text-zinc-500 dark:text-zinc-400">
                        代码中匹配关键词 “{codeSearchQuery}” 的频次: {(
                          ((editingTheme.content || editingTheme.rawJsonString || '').match(
                            new RegExp(codeSearchQuery.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi')
                          ) || []).length
                        )} 处
                      </div>
                    )}
                  </div>

                  {/* 导出 JSON 按钮 */}
                  <div className="pt-3 border-t border-zinc-200/60 dark:border-zinc-800/60 space-y-2">
                    <span className="text-xs font-semibold text-zinc-500 block">导出文档格式:</span>
                    <div>
                      <button
                        type="button"
                        onClick={() => exportThemeDocument(editingTheme, 'json')}
                        className="px-4 py-2 text-xs font-semibold rounded-xl border border-zinc-200 dark:border-zinc-700 bg-white dark:bg-zinc-800 text-zinc-800 dark:text-zinc-200 hover:bg-zinc-100 dark:hover:bg-zinc-700 flex items-center gap-1.5 shadow-sm transition-colors"
                      >
                        <Download className="w-3.5 h-3.5" /> 导出 .json
                      </button>
                    </div>
                  </div>
                </div>
              )}
            </div>

            {/* Modal Bottom Buttons (取消, 确定, 删除) */}
            <div className="px-6 py-4 border-t border-zinc-200 dark:border-zinc-800 flex items-center justify-between flex-shrink-0 bg-zinc-50/50 dark:bg-zinc-900/50">
              <button
                type="button"
                onClick={() => handleDeleteSingleTheme(editingTheme.id)}
                className="px-4 py-2 text-xs font-bold rounded-lg bg-rose-600 text-white hover:bg-rose-700 flex items-center gap-1"
              >
                <Trash2 className="w-3.5 h-3.5" /> 删除
              </button>

              <div className="flex items-center gap-2.5">
                <button
                  type="button"
                  onClick={() => setEditingTheme(null)}
                  className="px-4 py-2 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800"
                >
                  取消
                </button>
                <button
                  type="button"
                  onClick={handleSaveEditedTheme}
                  className="px-5 py-2 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:bg-zinc-800 dark:hover:bg-zinc-200"
                >
                  确定
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ==================== EXPANDED THEME CODE MODAL ==================== */}
      {isThemeCodeExpanded && editingTheme && (
        <div className="fixed inset-0 z-50 bg-white dark:bg-black flex flex-col animate-in fade-in pt-[calc(env(safe-area-inset-top,0px)+0.25rem)] pb-[calc(env(safe-area-inset-bottom,0px)+0.25rem)] pl-[env(safe-area-inset-left,0px)] pr-[env(safe-area-inset-right,0px)]">
          <div className="px-6 py-4 border-b border-zinc-200 dark:border-zinc-800 bg-white dark:bg-black flex items-center justify-between flex-shrink-0">
            <div className="flex items-center gap-3 min-w-0">
              <h3 className="font-bold text-sm text-zinc-900 dark:text-zinc-100 truncate">
                {editingTheme.name} - 文档/代码内容 (放大展开编辑)
              </h3>
              <span className="px-2 py-0.5 text-[10px] font-mono bg-zinc-100 dark:bg-zinc-900 text-zinc-600 dark:text-zinc-300 rounded uppercase border border-zinc-200 dark:border-zinc-800">
                {editingTheme.fileType || 'CODE'}
              </span>
            </div>
            <button
              onClick={() => setIsThemeCodeExpanded(false)}
              className="p-1.5 rounded-lg hover:bg-zinc-100 dark:hover:bg-zinc-900 text-zinc-500 dark:text-zinc-400 hover:text-zinc-900 dark:hover:text-zinc-200"
            >
              <X className="w-5 h-5" />
            </button>
          </div>

          <div className="p-4 border-b border-zinc-200 dark:border-zinc-800 bg-zinc-50 dark:bg-zinc-950 flex-shrink-0">
            <div className="relative w-full">
              <Search className="w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 text-zinc-400 dark:text-zinc-500" />
              <input
                type="text"
                value={codeSearchQuery}
                onChange={(e) => setCodeSearchQuery(e.target.value)}
                placeholder="搜索定位代码..."
                className="w-full pl-9 pr-4 py-1.5 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400 dark:focus:ring-zinc-500"
              />
            </div>

            <div className="flex items-center justify-end gap-2 mt-3">
              <button
                type="button"
                onClick={() => {
                  navigator.clipboard.writeText(editingTheme.content || editingTheme.rawJsonString || '');
                  showToast('已复制到剪贴板', 'success');
                }}
                className="px-3 py-1.5 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-800 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-900"
              >
                复制全文
              </button>
              <button
                type="button"
                onClick={() => setIsThemeCodeExpanded(false)}
                className="px-4 py-1.5 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:bg-zinc-800 dark:hover:bg-white"
              >
                完成修改
              </button>
            </div>
          </div>

          <div className="flex-1 p-4 sm:p-6 bg-white dark:bg-black overflow-hidden flex flex-col pb-[env(safe-area-inset-bottom,0px)]">
            <textarea
              value={editingTheme.content || editingTheme.rawJsonString || ''}
              onChange={(e) =>
                setEditingTheme((p) =>
                  p ? { ...p, content: e.target.value, rawJsonString: e.target.value } : null
                )
              }
              placeholder="在此编写或修改代码..."
              className="w-full h-full p-4 font-mono text-xs text-zinc-900 dark:text-zinc-100 bg-white dark:bg-black border border-zinc-200 dark:border-zinc-800 rounded-xl focus:outline-none leading-relaxed resize-none"
            />
          </div>
        </div>
      )}

      {/* ==================== NEW BEAUTIFICATION GROUP MODAL ==================== */}
      {showNewBeautificationGroupModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 p-4 animate-in fade-in">
          <div className="bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl w-full max-w-sm p-5 space-y-4 shadow-xl">
            <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100">新建美化分组</h3>
            <div>
              <label className="text-xs font-semibold text-zinc-500 mb-1 block">分组名字</label>
              <input
                type="text"
                autoFocus
                value={newBeautificationGroupName}
                onChange={(e) => setNewBeautificationGroupName(e.target.value)}
                placeholder="请输入分组名称..."
                onKeyDown={(e) => {
                  if (e.key === 'Enter') handleCreateNewBeautificationGroup();
                }}
                className="w-full px-3 py-2 text-sm bg-zinc-50 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400"
              />
            </div>
            <div className="flex justify-end gap-2 pt-2">
              <button
                onClick={() => {
                  setShowNewBeautificationGroupModal(false);
                  setNewBeautificationGroupName('');
                }}
                className="px-3 py-1.5 text-xs font-medium text-zinc-600 dark:text-zinc-400 hover:bg-zinc-100 dark:hover:bg-zinc-800 rounded-lg"
              >
                取消
              </button>
              <button
                onClick={handleCreateNewBeautificationGroup}
                className="px-4 py-1.5 text-xs font-bold bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 rounded-lg hover:opacity-90"
              >
                确定
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================== MANAGE BEAUTIFICATION GROUP MODAL ==================== */}
      {managingBeautificationGroup && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 p-4 animate-in fade-in">
          <div className="bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl w-full max-w-md p-5 space-y-4 shadow-xl">
            <div className="flex items-center justify-between border-b border-zinc-200 dark:border-zinc-800 pb-3">
              <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100">
                管理美化分组: {managingBeautificationGroup}
              </h3>
              <button onClick={() => setManagingBeautificationGroup(null)} className="p-1 rounded text-zinc-400 hover:text-zinc-600">
                <X className="w-4 h-4" />
              </button>
            </div>

            {/* Rename Group Section */}
            <div className="space-y-2">
              <label className="text-xs font-semibold text-zinc-500 block">重命名分组</label>
              <div className="flex gap-2">
                <input
                  type="text"
                  value={renameBeautificationGroupInput}
                  onChange={(e) => setRenameBeautificationGroupInput(e.target.value)}
                  className="flex-1 px-3 py-1.5 text-xs bg-zinc-50 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400"
                />
                <button
                  onClick={handleRenameBeautificationGroup}
                  className="px-3 py-1.5 text-xs font-medium bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 rounded-lg whitespace-nowrap"
                >
                  保存名字
                </button>
              </div>
            </div>

            {/* Delete Group Confirmation Dialog */}
            <div className="pt-3 border-t border-zinc-200 dark:border-zinc-800 space-y-3">
              <div className="text-xs font-bold text-rose-600 dark:text-rose-400">删除分组</div>
              <p className="text-xs text-zinc-500 dark:text-zinc-400 leading-relaxed">
                是否要删除分组“{managingBeautificationGroup}”？删除前请选择是否一并清理组内美化。
              </p>

              <div
                onClick={() => setDeleteBeautificationsWithGroup((prev) => !prev)}
                className="flex items-center gap-3 p-3 bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-xl cursor-pointer hover:bg-zinc-100 dark:hover:bg-zinc-800 transition-colors"
              >
                <div
                  className={`w-5 h-5 rounded-full border-2 flex items-center justify-center transition-all ${
                    deleteBeautificationsWithGroup
                      ? 'border-zinc-900 dark:border-zinc-100 bg-white dark:bg-zinc-900'
                      : 'border-zinc-400 dark:border-zinc-600 bg-transparent'
                  }`}
                >
                  {deleteBeautificationsWithGroup && (
                    <div className="w-2.5 h-2.5 rounded-full bg-zinc-900 dark:bg-zinc-100" />
                  )}
                </div>

                <div className="text-xs font-medium text-zinc-800 dark:text-zinc-200">
                  是否要一并把分组里的美化给删掉？
                </div>
              </div>

              <div className="flex justify-end gap-2 pt-2">
                <button
                  onClick={() => setManagingBeautificationGroup(null)}
                  className="px-3 py-1.5 text-xs font-medium text-zinc-600 dark:text-zinc-400 hover:bg-zinc-100 dark:hover:bg-zinc-800 rounded-lg"
                >
                  取消
                </button>
                <button
                  onClick={handleConfirmDeleteBeautificationGroup}
                  className="px-4 py-1.5 text-xs font-bold bg-rose-600 text-white hover:bg-rose-700 rounded-lg"
                >
                  确认删除分组
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ==================== BATCH MOVE BEAUTIFICATION MODAL ==================== */}
      {showBeautificationBatchMoveModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 p-4 animate-in fade-in">
          <div className="bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl w-full max-w-sm p-5 space-y-4 shadow-xl">
            <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100">
              批量移动美化 ({selectedBeautificationIds.length} 个)
            </h3>
            <div className="space-y-2">
              <label className="text-xs font-semibold text-zinc-500 mb-1 block">选择或输入目标分组</label>
              <select
                value={batchTargetBeautificationGroup}
                onChange={(e) => setBatchTargetBeautificationGroup(e.target.value)}
                className="w-full px-3 py-2 text-xs bg-zinc-50 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none"
              >
                <option value="">-- 请选择现有分组 --</option>
                {Array.from(new Set(['默认', ...(appData.beautificationCategories || [])])).map((cat) => (
                  <option key={cat} value={cat}>
                    {cat}
                  </option>
                ))}
              </select>
              <input
                type="text"
                value={batchTargetBeautificationGroup}
                onChange={(e) => setBatchTargetBeautificationGroup(e.target.value)}
                placeholder="或直接输入新的分组名称..."
                className="w-full px-3 py-2 text-xs bg-zinc-50 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400"
              />
            </div>
            <div className="flex justify-end gap-2 pt-2">
              <button
                onClick={() => {
                  setShowBeautificationBatchMoveModal(false);
                  setBatchTargetBeautificationGroup('');
                }}
                className="px-3 py-1.5 text-xs font-medium text-zinc-600 dark:text-zinc-400 hover:bg-zinc-100 dark:hover:bg-zinc-800 rounded-lg"
              >
                取消
              </button>
              <button
                disabled={!batchTargetBeautificationGroup.trim()}
                onClick={handleBatchMoveBeautifications}
                className="px-4 py-1.5 text-xs font-bold bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 rounded-lg hover:opacity-90 disabled:opacity-40"
              >
                确认移动
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================== EDIT BEAUTIFICATION DETAIL MODAL ==================== */}
      {editingBeautification && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-3 sm:p-4 pt-[calc(1rem+env(safe-area-inset-top,0px))] pb-[calc(1rem+env(safe-area-inset-bottom,0px))] animate-in fade-in">
          <div className="w-full max-w-4xl h-[85vh] min-h-[550px] max-h-[750px] bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl flex flex-col overflow-hidden transition-none">
            {/* Hidden Input for Cover Image */}
            <input
              type="file"
              ref={beautificationCoverInputRef}
              accept="image/*"
              className="hidden"
              onChange={(e) => {
                const file = e.target.files?.[0];
                if (file) {
                  const reader = new FileReader();
                  reader.onload = () => {
                    if (reader.result) {
                      setEditingBeautification((p) => p ? { ...p, coverImage: reader.result as string } : null);
                    }
                  };
                  reader.readAsDataURL(file);
                }
                e.target.value = '';
              }}
            />

            {/* Modal Header */}
            <div className="px-6 py-4 border-b border-zinc-200 dark:border-zinc-800 flex items-center justify-between flex-shrink-0 bg-zinc-50/50 dark:bg-zinc-900/50">
              <div className="flex items-center gap-2 min-w-0">
                <h2 className="text-base font-bold truncate text-zinc-900 dark:text-zinc-100">
                  {editingBeautification.name || '美化详情'}
                </h2>
                <span className="px-2 py-0.5 text-[10px] font-bold uppercase rounded bg-zinc-200 dark:bg-zinc-800 text-zinc-700 dark:text-zinc-300">
                  {editingBeautification.type || editingBeautification.fileType?.toUpperCase() || '美化'}
                </span>
              </div>
              <button
                type="button"
                onClick={() => setEditingBeautification(null)}
                className="p-1.5 rounded-lg hover:bg-zinc-200 dark:hover:bg-zinc-800 text-zinc-500 transition-colors"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* 横向 Tab 导航 */}
            <div className="flex border-b border-zinc-200 dark:border-zinc-800 px-6 gap-2 flex-shrink-0 bg-white dark:bg-zinc-900">
              <button
                type="button"
                onClick={() => setBeautificationDetailTab('preview')}
                className={`py-3 px-4 text-xs font-bold border-b-2 transition-colors ${
                  beautificationDetailTab === 'preview'
                    ? 'border-zinc-900 text-zinc-900 dark:border-zinc-100 dark:text-zinc-100'
                    : 'border-transparent text-zinc-500 hover:text-zinc-800 dark:hover:text-zinc-200'
                }`}
              >
                美化预览
              </button>
              <button
                type="button"
                onClick={() => setBeautificationDetailTab('code')}
                className={`py-3 px-4 text-xs font-bold border-b-2 transition-colors ${
                  beautificationDetailTab === 'code'
                    ? 'border-zinc-900 text-zinc-900 dark:border-zinc-100 dark:text-zinc-100'
                    : 'border-transparent text-zinc-500 hover:text-zinc-800 dark:hover:text-zinc-200'
                }`}
              >
                文档内容
              </button>
            </div>

            <div className="flex-1 overflow-y-auto p-6 space-y-6">
              {beautificationDetailTab === 'preview' && (
              <div className="bg-zinc-50/50 dark:bg-zinc-800/20 border border-zinc-200 dark:border-zinc-800 rounded-2xl p-4">
                <h3 className="text-xs font-bold text-zinc-800 dark:text-zinc-200 mb-4 flex items-center gap-1.5">
                  <ImageIcon className="w-4 h-4 text-zinc-500" /> 第一栏：美化预览与基本信息
                </h3>
                <div className="grid grid-cols-1 md:grid-cols-12 gap-6">
                  {/* Left Column: Cover Image Preview & Upload */}
                  <div className="md:col-span-4 space-y-3">
                    <div className="aspect-[2/3] w-full rounded-xl overflow-hidden bg-zinc-100 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-800 relative flex items-center justify-center p-2">
                      {editingBeautification.coverImage ? (
                        <img src={editingBeautification.coverImage} alt="Beautification Cover" className="w-full h-full object-cover" />
                      ) : (
                        <div className="text-center p-4 space-y-2">
                          <ImageIcon className="w-10 h-10 text-zinc-400 mx-auto" />
                          <p className="text-xs text-zinc-400">暂无封面图片</p>
                        </div>
                      )}
                    </div>

                    <button
                      type="button"
                      onClick={() => beautificationCoverInputRef.current?.click()}
                      className="w-full py-2 text-xs font-medium rounded-lg border border-dashed border-zinc-300 dark:border-zinc-700 hover:border-zinc-500 text-zinc-700 dark:text-zinc-300 transition-colors flex items-center justify-center gap-1.5"
                    >
                      <Upload className="w-3.5 h-3.5" /> 上传图片 (当封面图)
                    </button>
                  </div>

                  {/* Right Column: Fields (名称, 作者, 类型, 来源, 分组) */}
                  <div className="md:col-span-8 space-y-3">
                    <div>
                      <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                        美化名称 <span className="text-rose-500">*</span>
                      </label>
                      <input
                        type="text"
                        value={editingBeautification.name}
                        onChange={(e) => setEditingBeautification((p) => p ? { ...p, name: e.target.value } : null)}
                        placeholder="请输入美化名称..."
                        className="w-full px-3 py-2 text-xs bg-white dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400 text-zinc-900 dark:text-zinc-100"
                      />
                    </div>

                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                      <div>
                        <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                          作者
                        </label>
                        <input
                          type="text"
                          value={editingBeautification.author || ''}
                          onChange={(e) => setEditingBeautification((p) => p ? { ...p, author: e.target.value } : null)}
                          placeholder="作者名字 (默认)..."
                          className="w-full px-3 py-2 text-xs bg-white dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400 text-zinc-900 dark:text-zinc-100"
                        />
                      </div>

                      <div>
                        <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                          美化类型
                        </label>
                        <input
                          type="text"
                          value={editingBeautification.type || ''}
                          onChange={(e) => setEditingBeautification((p) => p ? { ...p, type: e.target.value } : null)}
                          placeholder="例如: 线上主题, CSS样式..."
                          className="w-full px-3 py-2 text-xs bg-white dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400 text-zinc-900 dark:text-zinc-100"
                        />
                      </div>
                    </div>

                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                      <div>
                        <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                          来源/链接
                        </label>
                        <input
                          type="text"
                          value={editingBeautification.source || ''}
                          onChange={(e) => setEditingBeautification((p) => p ? { ...p, source: e.target.value } : null)}
                          placeholder="来源网址或作者主页..."
                          className="w-full px-3 py-2 text-xs bg-white dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400 text-zinc-900 dark:text-zinc-100"
                        />
                      </div>

                      <div>
                        <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                          所属分组
                        </label>
                        <div className="flex gap-2">
                          <select
                            value={editingBeautification.category || '默认'}
                            onChange={(e) => setEditingBeautification((p) => p ? { ...p, category: e.target.value } : null)}
                            className="flex-1 px-3 py-2 text-xs bg-white dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none text-zinc-900 dark:text-zinc-100"
                          >
                            {Array.from(new Set(['默认', ...(appData.beautificationCategories || [])])).map((cat) => (
                              <option key={cat} value={cat}>
                                {cat}
                              </option>
                            ))}
                          </select>
                          <input
                            type="text"
                            value={editingBeautification.category || ''}
                            onChange={(e) => setEditingBeautification((p) => p ? { ...p, category: e.target.value } : null)}
                            placeholder="或手动新填..."
                            className="w-1/2 px-3 py-2 text-xs bg-white dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none text-zinc-900 dark:text-zinc-100"
                          />
                        </div>
                      </div>
                    </div>

                    <div>
                      <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                        备注/描述信息
                      </label>
                      <textarea
                        rows={2}
                        value={editingBeautification.description || ''}
                        onChange={(e) => setEditingBeautification((p) => p ? { ...p, description: e.target.value } : null)}
                        placeholder="关于此美化的备注说明..."
                        className="w-full px-3 py-2 text-xs bg-white dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400 text-zinc-900 dark:text-zinc-100 resize-none"
                      />
                    </div>
                  </div>
                </div>
              </div>
              )}
              {beautificationDetailTab === 'code' && (
              <div className="bg-zinc-50/50 dark:bg-zinc-800/20 border border-zinc-200 dark:border-zinc-800 rounded-2xl p-4 space-y-3">
                <div className="flex items-center justify-between">
                  <h3 className="text-xs font-bold text-zinc-800 dark:text-zinc-200 flex items-center gap-1.5">
                    <FileCode className="w-4 h-4 text-zinc-500" /> 第二栏：美化核心代码 / 文档文本
                  </h3>
                  <button
                    type="button"
                    onClick={() => setIsBeautificationCodeExpanded(true)}
                    className="px-2.5 py-1 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 hover:bg-zinc-100 dark:hover:bg-zinc-800 text-zinc-700 dark:text-zinc-300 flex items-center gap-1"
                  >
                    <Maximize2 className="w-3 h-3" /> 放大展开编辑
                  </button>
                </div>

                {/* 文档搜索栏 */}
                <div className="relative w-full">
                  <Search className="w-3.5 h-3.5 absolute left-3 top-1/2 -translate-y-1/2 text-zinc-400" />
                  <input
                    type="text"
                    value={codeSearchQuery}
                    onChange={(e) => setCodeSearchQuery(e.target.value)}
                    placeholder="搜索/定位文档内容..."
                    className="w-full pl-9 pr-4 py-2 text-xs bg-white dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400 text-zinc-900 dark:text-zinc-100"
                  />
                </div>

                <div className="relative rounded-xl border border-zinc-200 dark:border-zinc-800 overflow-hidden bg-white dark:bg-zinc-950">
                  <textarea
                    rows={8}
                    value={editingBeautification.content || editingBeautification.rawJsonString || ''}
                    onChange={(e) =>
                      setEditingBeautification((p) =>
                        p ? { ...p, content: e.target.value, rawJsonString: e.target.value } : null
                      )
                    }
                    placeholder="在此编辑美化代码或文档内容..."
                    className="w-full p-4 font-mono text-xs bg-transparent text-zinc-900 dark:text-zinc-100 focus:outline-none leading-relaxed resize-y"
                  />
                </div>

                {/* Bottom Export Action Buttons */}
                <div className="pt-3 border-t border-zinc-200/60 dark:border-zinc-800/60 space-y-2">
                  <div className="text-[11px] text-zinc-500 dark:text-zinc-400">
                    创建时间: {editingBeautification.createdAt ? new Date(editingBeautification.createdAt).toLocaleString('zh-CN') : '未知时间'}
                  </div>

                  <div className="space-y-1.5">
                    <span className="text-xs font-semibold text-zinc-600 dark:text-zinc-300 block">
                      导出格式：
                    </span>
                    <div className="flex flex-wrap items-center gap-2">
                      <button
                        type="button"
                        onClick={() => exportBeautificationDocument(editingBeautification, 'docx')}
                        className="px-3 py-1.5 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 bg-white dark:bg-zinc-800 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-700 flex items-center gap-1 transition-colors"
                      >
                        <Download className="w-3.5 h-3.5" /> 导出 .docx
                      </button>
                      <button
                        type="button"
                        onClick={() => exportBeautificationDocument(editingBeautification, 'txt')}
                        className="px-3 py-1.5 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 bg-white dark:bg-zinc-800 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-700 flex items-center gap-1 transition-colors"
                      >
                        <Download className="w-3.5 h-3.5" /> 导出 .txt
                      </button>
                      <button
                        type="button"
                        onClick={() => exportBeautificationDocument(editingBeautification, 'json')}
                        className="px-3 py-1.5 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 bg-white dark:bg-zinc-800 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-700 flex items-center gap-1 transition-colors"
                      >
                        <Download className="w-3.5 h-3.5" /> 导出 .json
                      </button>
                      <button
                        type="button"
                        onClick={() => exportBeautificationDocument(editingBeautification, 'css')}
                        className="px-3 py-1.5 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 bg-white dark:bg-zinc-800 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-700 flex items-center gap-1 transition-colors"
                      >
                        <Download className="w-3.5 h-3.5" /> 导出 .css
                      </button>
                    </div>
                  </div>
                </div>
              </div>
              )}
            </div>

            {/* Modal Bottom Buttons (取消, 确定, 删除) */}
            <div className="px-6 py-4 border-t border-zinc-200 dark:border-zinc-800 flex items-center justify-between flex-shrink-0 bg-zinc-50/50 dark:bg-zinc-900/50">
              <button
                type="button"
                onClick={() => handleDeleteSingleBeautification(editingBeautification.id)}
                className="px-4 py-2 text-xs font-bold rounded-lg bg-rose-600 text-white hover:bg-rose-700 flex items-center gap-1"
              >
                <Trash2 className="w-3.5 h-3.5" /> 删除
              </button>

              <div className="flex items-center gap-2.5">
                <button
                  type="button"
                  onClick={() => setEditingBeautification(null)}
                  className="px-4 py-2 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800"
                >
                  取消
                </button>
                <button
                  type="button"
                  onClick={handleSaveEditedBeautification}
                  className="px-5 py-2 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:bg-zinc-800 dark:hover:bg-zinc-200"
                >
                  确定
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ==================== EXPANDED BEAUTIFICATION CODE MODAL ==================== */}
      {isBeautificationCodeExpanded && editingBeautification && (
        <div className="fixed inset-0 z-[60] bg-white text-zinc-900 flex flex-col animate-in fade-in pt-[env(safe-area-inset-top,0px)] pb-[env(safe-area-inset-bottom,0px)] pl-[env(safe-area-inset-left,0px)] pr-[env(safe-area-inset-right,0px)]">
          <div className="px-4 sm:px-6 py-3 border-b border-zinc-200 bg-white flex items-center justify-between flex-shrink-0 min-h-[3.5rem]">
            <div className="flex items-center gap-2 min-w-0">
              <h3 className="font-bold text-sm text-zinc-900 truncate">
                {editingBeautification.name} - 文档内容
              </h3>
              <span className="px-2 py-0.5 text-[10px] font-mono bg-zinc-100 text-zinc-600 rounded uppercase border border-zinc-200 flex-shrink-0">
                {editingBeautification.fileType || editingBeautification.type || 'CODE'}
              </span>
            </div>
            <button
              type="button"
              onClick={() => setIsBeautificationCodeExpanded(false)}
              className="p-2 rounded-lg hover:bg-zinc-100 text-zinc-500 hover:text-zinc-900 flex-shrink-0"
              aria-label="关闭全屏文档"
            >
              <X className="w-5 h-5" />
            </button>
          </div>

          <div className="px-4 sm:px-6 py-3 border-b border-zinc-200 bg-white flex-shrink-0">
            <div className="relative w-full">
              <Search className="w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 text-zinc-400" />
              <input
                type="text"
                value={codeSearchQuery}
                onChange={(e) => setCodeSearchQuery(e.target.value)}
                placeholder="搜索/定位文档内容..."
                className="w-full pl-10 pr-4 py-2.5 text-sm bg-white border border-zinc-300 rounded-xl text-zinc-900 placeholder:text-zinc-400 focus:outline-none focus:ring-2 focus:ring-zinc-300"
              />
            </div>
            <div className="flex items-center justify-end gap-2 mt-2">
              <button
                type="button"
                onClick={() => {
                  navigator.clipboard.writeText(editingBeautification.content || editingBeautification.rawJsonString || '');
                  showToast('已复制到剪贴板', 'success');
                }}
                className="px-3 py-1.5 text-xs font-medium rounded-lg border border-zinc-300 text-zinc-700 hover:bg-zinc-100"
              >
                复制全文
              </button>
              <button
                type="button"
                onClick={() => setIsBeautificationCodeExpanded(false)}
                className="px-4 py-1.5 text-xs font-bold rounded-lg bg-zinc-900 text-white hover:bg-zinc-800"
              >
                完成修改
              </button>
            </div>
          </div>

          <div className="flex-1 min-h-0 p-3 sm:p-5 bg-white overflow-hidden pb-[calc(0.75rem+env(safe-area-inset-bottom,0px))]">
            <textarea
              value={editingBeautification.content || editingBeautification.rawJsonString || ''}
              onChange={(e) =>
                setEditingBeautification((p) =>
                  p ? { ...p, content: e.target.value, rawJsonString: e.target.value } : null
                )
              }
              placeholder="在此编写或修改文档内容..."
              className="w-full h-full p-4 font-mono text-xs sm:text-sm text-zinc-900 bg-white border border-zinc-200 rounded-xl focus:outline-none focus:ring-1 focus:ring-zinc-300 leading-relaxed resize-none"
            />
          </div>
        </div>
      )}

      {/* ==================== FULLSCREEN TEXT / VIEWER MODAL ==================== */}
      {fullscreenData && (
        <div className="fixed inset-0 z-50 bg-white dark:bg-zinc-950 flex flex-col animate-in fade-in">
          <div className="px-6 border-b border-zinc-200 dark:border-zinc-800 flex items-center justify-between flex-shrink-0 pt-[env(safe-area-inset-top,0px)] min-h-[calc(3.5rem+env(safe-area-inset-top,0px))] bg-zinc-50/80 dark:bg-zinc-900/80 backdrop-blur-md">
            <h3 className="font-bold text-sm text-zinc-900 dark:text-zinc-100 py-3 truncate max-w-[50vw]">{fullscreenData.title}</h3>
            <div className="flex items-center gap-2 py-3">
              <button
                onClick={() => {
                  navigator.clipboard.writeText(fullscreenData.content);
                  showToast('内容已复制到剪贴板！', 'success');
                }}
                className="px-3 py-1.5 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-800 hover:bg-zinc-100 dark:hover:bg-zinc-800"
              >
                复制全文
              </button>
              <button
                onClick={() => setFullscreenData(null)}
                className="px-3 py-1.5 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900"
              >
                关闭 ✕
              </button>
            </div>
          </div>
          <div className="flex-1 overflow-y-auto p-6 pb-[calc(1.5rem+env(safe-area-inset-bottom,0px))]">
            {fullscreenData.content.startsWith('data:image/') || fullscreenData.content.startsWith('blob:') ? (
              <div className="flex items-center justify-center h-full">
                <img src={fullscreenData.content} alt="Fullscreen" className="max-w-full max-h-[85vh] object-contain rounded-xl shadow-lg" />
              </div>
            ) : (
              <textarea
                readOnly
                value={fullscreenData.content}
                className="w-full h-full min-h-[75vh] p-4 text-sm font-sans leading-relaxed bg-zinc-50 dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-xl focus:outline-none"
              />
            )}
          </div>
        </div>
      )}

      {/* ==================== NEW PRESET GROUP MODAL ==================== */}
      {showNewPresetGroupModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 p-4 animate-in fade-in">
          <div className="bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl w-full max-w-sm p-5 space-y-4 shadow-xl">
            <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100">新建预设分组</h3>
            <div>
              <label className="text-xs font-semibold text-zinc-500 mb-1 block">分组名字</label>
              <input
                type="text"
                autoFocus
                value={newPresetGroupName}
                onChange={(e) => setNewPresetGroupName(e.target.value)}
                placeholder="请输入分组名称..."
                onKeyDown={(e) => {
                  if (e.key === 'Enter') handleCreateNewPresetGroup();
                }}
                className="w-full px-3 py-2 text-sm bg-zinc-50 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400"
              />
            </div>
            <div className="flex justify-end gap-2 pt-2">
              <button
                onClick={() => {
                  setShowNewPresetGroupModal(false);
                  setNewPresetGroupName('');
                }}
                className="px-3 py-1.5 text-xs font-medium text-zinc-600 dark:text-zinc-400 hover:bg-zinc-100 dark:hover:bg-zinc-800 rounded-lg"
              >
                取消
              </button>
              <button
                onClick={handleCreateNewPresetGroup}
                className="px-4 py-1.5 text-xs font-bold bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 rounded-lg hover:opacity-90"
              >
                确定
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================== MANAGE / DELETE PRESET GROUP MODAL ==================== */}
      {managingPresetGroup && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 p-4 animate-in fade-in">
          <div className="bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl w-full max-w-md p-5 space-y-4 shadow-xl">
            <div className="flex items-center justify-between border-b border-zinc-200 dark:border-zinc-800 pb-3">
              <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100">
                管理预设分组: {managingPresetGroup}
              </h3>
              <button onClick={() => setManagingPresetGroup(null)} className="p-1 rounded text-zinc-400 hover:text-zinc-600">
                <X className="w-4 h-4" />
              </button>
            </div>

            {/* Rename Group Section */}
            <div className="space-y-2">
              <label className="text-xs font-semibold text-zinc-500 block">重命名分组</label>
              <div className="flex gap-2">
                <input
                  type="text"
                  value={renamePresetGroupInput}
                  onChange={(e) => setRenamePresetGroupInput(e.target.value)}
                  className="flex-1 px-3 py-1.5 text-xs bg-zinc-50 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400"
                />
                <button
                  onClick={handleRenamePresetGroup}
                  className="px-3 py-1.5 text-xs font-medium bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 rounded-lg whitespace-nowrap"
                >
                  保存名字
                </button>
              </div>
            </div>

            {/* Delete Group Confirmation Dialog with Circular Toggle */}
            <div className="pt-3 border-t border-zinc-200 dark:border-zinc-800 space-y-3">
              <div className="text-xs font-bold text-rose-600 dark:text-rose-400">删除分组</div>
              <p className="text-xs text-zinc-500 dark:text-zinc-400 leading-relaxed">
                是否要删除分组“{managingPresetGroup}”？删除前请选择是否一并清理组内预设。
              </p>

              {/* Circular Selection Toggle */}
              <div
                onClick={() => setDeletePresetsWithGroup((prev) => !prev)}
                className="flex items-center gap-3 p-3 bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-xl cursor-pointer hover:bg-zinc-100 dark:hover:bg-zinc-800 transition-colors"
              >
                {/* Circular Indicator */}
                <div
                  className={`w-5 h-5 rounded-full border-2 flex items-center justify-center transition-all ${
                    deletePresetsWithGroup
                      ? 'border-zinc-900 dark:border-zinc-100 bg-white dark:bg-zinc-900'
                      : 'border-zinc-400 dark:border-zinc-600 bg-transparent'
                  }`}
                >
                  {deletePresetsWithGroup && (
                    <div className="w-2.5 h-2.5 rounded-full bg-zinc-900 dark:bg-zinc-100" />
                  )}
                </div>

                <div className="text-xs font-medium text-zinc-800 dark:text-zinc-200">
                  是否要一并把分组里的预设给删掉？
                </div>
              </div>

              <div className="flex justify-end gap-2 pt-2">
                <button
                  onClick={() => setManagingPresetGroup(null)}
                  className="px-3 py-1.5 text-xs font-medium text-zinc-600 dark:text-zinc-400 hover:bg-zinc-100 dark:hover:bg-zinc-800 rounded-lg"
                >
                  取消
                </button>
                <button
                  onClick={handleConfirmDeletePresetGroup}
                  className="px-4 py-1.5 text-xs font-bold bg-rose-600 text-white hover:bg-rose-700 rounded-lg"
                >
                  确认删除分组
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ==================== BATCH MOVE PRESET MODAL ==================== */}
      {showPresetBatchMoveModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 p-4 animate-in fade-in">
          <div className="bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl w-full max-w-sm p-5 space-y-4 shadow-xl">
            <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100">
              批量移动预设 ({selectedPresetIds.length} 个)
            </h3>
            <div className="space-y-2">
              <label className="text-xs font-semibold text-zinc-500 mb-1 block">选择或输入目标分组</label>
              <select
                value={batchTargetPresetGroup}
                onChange={(e) => setBatchTargetPresetGroup(e.target.value)}
                className="w-full px-3 py-2 text-xs bg-zinc-50 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none"
              >
                <option value="">-- 请选择现有分组 --</option>
                {Array.from(new Set(['默认', ...(appData.presetCategories || [])])).map((cat) => (
                  <option key={cat} value={cat}>
                    {cat}
                  </option>
                ))}
              </select>
              <input
                type="text"
                value={batchTargetPresetGroup}
                onChange={(e) => setBatchTargetPresetGroup(e.target.value)}
                placeholder="或直接输入新的分组名称..."
                className="w-full px-3 py-2 text-xs bg-zinc-50 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400"
              />
            </div>
            <div className="flex justify-end gap-2 pt-2">
              <button
                onClick={() => {
                  setShowPresetBatchMoveModal(false);
                  setBatchTargetPresetGroup('');
                }}
                className="px-3 py-1.5 text-xs font-medium text-zinc-600 dark:text-zinc-400 hover:bg-zinc-100 dark:hover:bg-zinc-800 rounded-lg"
              >
                取消
              </button>
              <button
                disabled={!batchTargetPresetGroup.trim()}
                onClick={handleBatchMovePresets}
                className="px-4 py-1.5 text-xs font-bold bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 rounded-lg hover:opacity-90 disabled:opacity-40"
              >
                确认移动
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================== EDIT PRESET DETAIL MODAL ==================== */}
      {editingPreset && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-3 sm:p-4 pt-[calc(1rem+env(safe-area-inset-top,0px))] pb-[calc(1rem+env(safe-area-inset-bottom,0px))] animate-in fade-in">
          <div className="w-full max-w-3xl h-[85vh] min-h-[550px] max-h-[750px] bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl flex flex-col overflow-hidden">
            {/* Modal Header */}
            <div className="px-6 py-4 border-b border-zinc-200 dark:border-zinc-800 flex items-center justify-between flex-shrink-0 bg-zinc-50/50 dark:bg-zinc-900/50">
              <div className="flex items-center gap-2 min-w-0">
                <Sliders className="w-5 h-5 text-zinc-600 dark:text-zinc-300 flex-shrink-0" />
                <h2 className="text-base font-bold truncate text-zinc-900 dark:text-zinc-100">
                  {editingPreset.fileName || editingPreset.name || '预设详情'}
                </h2>
                <span className="px-2 py-0.5 text-[10px] font-bold uppercase rounded bg-zinc-200 dark:bg-zinc-800 text-zinc-700 dark:text-zinc-300 whitespace-nowrap">
                  JSON 预设
                </span>
              </div>
              <button
                onClick={() => setEditingPreset(null)}
                className="p-1.5 rounded-lg hover:bg-zinc-200 dark:hover:bg-zinc-800 text-zinc-500 transition-colors"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Modal Functional Tabs */}
            <div className="flex border-b border-zinc-200 dark:border-zinc-800 px-6 gap-2 overflow-x-auto flex-shrink-0 bg-white dark:bg-zinc-900 scrollbar-none">
              <button
                onClick={() => setEditingPresetTab('details')}
                className={`py-3 px-4 text-xs font-semibold border-b-2 whitespace-nowrap transition-colors ${
                  editingPresetTab === 'details'
                    ? 'border-zinc-900 text-zinc-900 dark:border-zinc-100 dark:text-zinc-100'
                    : 'border-transparent text-zinc-500 hover:text-zinc-800 dark:hover:text-zinc-200'
                }`}
              >
                预设详情
              </button>
              <button
                onClick={() => setEditingPresetTab('content')}
                className={`py-3 px-4 text-xs font-semibold border-b-2 whitespace-nowrap transition-colors ${
                  editingPresetTab === 'content'
                    ? 'border-zinc-900 text-zinc-900 dark:border-zinc-100 dark:text-zinc-100'
                    : 'border-transparent text-zinc-500 hover:text-zinc-800 dark:hover:text-zinc-200'
                }`}
              >
                文件内容数据
              </button>
              <button
                onClick={() => setEditingPresetTab('regex')}
                className={`py-3 px-4 text-xs font-semibold border-b-2 whitespace-nowrap transition-colors ${
                  editingPresetTab === 'regex'
                    ? 'border-zinc-900 text-zinc-900 dark:border-zinc-100 dark:text-zinc-100'
                    : 'border-transparent text-zinc-500 hover:text-zinc-800 dark:hover:text-zinc-200'
                }`}
              >
                正则内容 ({getPresetRegexScripts(editingPreset).length})
              </button>
            </div>

            {/* Modal Body */}
            <div className="flex-1 overflow-y-auto p-6 space-y-4">
              {editingPresetTab === 'details' && (
                /* Tab 1: 预设详情 (预设名，作者栏，分类，来源 (dc链接等等), 描述) */
                <div className="space-y-4 max-w-2xl mx-auto">
                  {/* 预设名 */}
                  <div>
                    <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                      预设名 <span className="text-rose-500">*</span>
                    </label>
                    <input
                      type="text"
                      value={editingPreset.name}
                      onChange={(e) => setEditingPreset((p) => p ? { ...p, name: e.target.value } : null)}
                      placeholder="请输入预设名称..."
                      className="w-full px-3 py-2 text-sm bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400 text-zinc-900 dark:text-zinc-100"
                    />
                  </div>

                  {/* 作者栏 */}
                  <div>
                    <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                      作者栏
                    </label>
                    <input
                      type="text"
                      value={editingPreset.author || ''}
                      onChange={(e) => setEditingPreset((p) => p ? { ...p, author: e.target.value } : null)}
                      placeholder="请输入作者名称..."
                      className="w-full px-3 py-2 text-sm bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400 text-zinc-900 dark:text-zinc-100"
                    />
                  </div>

                  {/* 分类 */}
                  <div>
                    <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                      分类
                    </label>
                    <input
                      type="text"
                      value={editingPreset.category || ''}
                      onChange={(e) => setEditingPreset((p) => p ? { ...p, category: e.target.value } : null)}
                      placeholder="请输入分类 (如: 默认, 提示词, 模型参数等)..."
                      className="w-full px-3 py-2 text-sm bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400 text-zinc-900 dark:text-zinc-100"
                    />
                  </div>

                  {/* 来源 (dc链接等等) */}
                  <div>
                    <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                      来源（DC链接 / Q群 / 网址等）
                    </label>
                    <input
                      type="text"
                      value={editingPreset.source || ''}
                      onChange={(e) => setEditingPreset((p) => p ? { ...p, source: e.target.value } : null)}
                      placeholder="例如: Discord 链接 / QQ群: 123456 / https://..."
                      className="w-full px-3 py-2 text-sm bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400 text-zinc-900 dark:text-zinc-100"
                    />
                  </div>

                  {/* 描述/备注 */}
                  <div>
                    <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                      描述 / 备注
                    </label>
                    <textarea
                      rows={4}
                      value={editingPreset.description || ''}
                      onChange={(e) => setEditingPreset((p) => p ? { ...p, description: e.target.value } : null)}
                      placeholder="请输入预设的说明、适用场景或备注信息..."
                      className="w-full px-3 py-2 text-sm bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400 text-zinc-900 dark:text-zinc-100 resize-none leading-relaxed"
                    />
                  </div>
                </div>
              )}

              {editingPresetTab === 'content' && (
                /* Tab 2: 文件内容数据 (含条目总数、多选、增删改、上下移动、跨预设移动) */
                <div className="space-y-3.5 flex flex-col h-full">
                  {(() => {
                    const allEntries = getPresetEntriesList(editingPreset);
                    const filteredEntries = allEntries.filter((e) => {
                      if (!presetEntrySearchQuery.trim()) return true;
                      const q = presetEntrySearchQuery.toLowerCase();
                      return e.name.toLowerCase().includes(q) || e.content.toLowerCase().includes(q);
                    });

                    return (
                      <>
                        {/* Header Bar */}
                        <div className="flex flex-wrap items-center justify-between gap-2 pb-2 border-b border-zinc-200 dark:border-zinc-800">
                          <div className="flex items-center gap-2">
                            <span className="text-xs font-bold text-zinc-900 dark:text-zinc-100 flex items-center gap-2">
                              内容列表 <span className="px-2 py-0.5 text-[11px] font-bold rounded-full bg-zinc-100 dark:bg-zinc-800 text-zinc-700 dark:text-zinc-300">共 {allEntries.length} 条</span>
                            </span>
                          </div>

                          <div className="flex flex-wrap items-center gap-2">
                            <button
                              type="button"
                              onClick={() => {
                                setNewPresetEntryForm({ name: '', content: '' });
                                setShowAddPresetEntryModal(true);
                              }}
                              className="px-2.5 py-1 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:opacity-90 flex items-center gap-1 shadow-sm"
                            >
                              <Plus className="w-3.5 h-3.5" /> 添加条目
                            </button>
                            <button
                              type="button"
                              onClick={() => {
                                setPresetEntryBatchMode(!presetEntryBatchMode);
                                setSelectedPresetEntryIndices([]);
                              }}
                              className={`px-2.5 py-1 text-xs font-bold rounded-lg border transition-colors flex items-center gap-1 ${
                                presetEntryBatchMode
                                  ? 'border-zinc-900 bg-zinc-900 text-white dark:border-zinc-100 dark:bg-zinc-100 dark:text-zinc-900'
                                  : 'border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800'
                              }`}
                            >
                              <CheckSquare className="w-3.5 h-3.5" /> {presetEntryBatchMode ? '退出多选' : '多选'}
                            </button>
                          </div>
                        </div>

                        {/* Batch Toolbar if Batch Mode */}
                        {presetEntryBatchMode && (
                          <div className="p-2.5 bg-zinc-100 dark:bg-zinc-800/80 rounded-xl flex flex-wrap items-center justify-between gap-2 border border-zinc-200 dark:border-zinc-700/60 animate-in fade-in">
                            <div className="flex items-center gap-2">
                              <button
                                type="button"
                                onClick={() => {
                                  if (selectedPresetEntryIndices.length === allEntries.length) {
                                    setSelectedPresetEntryIndices([]);
                                  } else {
                                    setSelectedPresetEntryIndices(allEntries.map((item) => item.index));
                                  }
                                }}
                                className="px-2.5 py-1 text-xs font-bold rounded-lg bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-50 dark:hover:bg-zinc-800 flex items-center gap-1 shadow-sm"
                              >
                                {selectedPresetEntryIndices.length === allEntries.length ? (
                                  <>
                                    <Square className="w-3.5 h-3.5" /> 取消全选
                                  </>
                                ) : (
                                  <>
                                    <CheckSquare className="w-3.5 h-3.5" /> 全选 ({allEntries.length})
                                  </>
                                )}
                              </button>
                              <span className="text-xs text-zinc-500 font-medium">
                                已选中 <span className="font-bold text-zinc-900 dark:text-zinc-100">{selectedPresetEntryIndices.length}</span> 条
                              </span>
                            </div>

                            <div className="flex items-center gap-2">
                              <button
                                type="button"
                                disabled={selectedPresetEntryIndices.length === 0}
                                onClick={() => {
                                  if (selectedPresetEntryIndices.length === 0) return;
                                  setMovingEntryIndices(selectedPresetEntryIndices);
                                  setTargetMovePresetId('');
                                  setShowMovePresetEntryModal(true);
                                }}
                                className="px-2.5 py-1 text-xs font-bold rounded-lg border border-zinc-200 dark:border-zinc-700 bg-white dark:bg-zinc-900 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-50 dark:hover:bg-zinc-800 disabled:opacity-40 flex items-center gap-1 shadow-sm"
                              >
                                <ArrowRightLeft className="w-3.5 h-3.5" /> 移动到其他预设
                              </button>
                              <button
                                type="button"
                                disabled={selectedPresetEntryIndices.length === 0}
                                onClick={() => {
                                  if (selectedPresetEntryIndices.length === 0) return;
                                  if (!confirm(`确定要删除选中的 ${selectedPresetEntryIndices.length} 个条目吗？`)) return;
                                  const newList = allEntries.filter((item) => !selectedPresetEntryIndices.includes(item.index));
                                  savePresetEntriesList(editingPreset, newList);
                                  setSelectedPresetEntryIndices([]);
                                  showToast(`已删除 ${selectedPresetEntryIndices.length} 个条目`, 'info');
                                }}
                                className="px-2.5 py-1 text-xs font-bold rounded-lg bg-rose-600 text-white hover:bg-rose-700 disabled:opacity-40 flex items-center gap-1 shadow-sm"
                              >
                                <Trash2 className="w-3.5 h-3.5" /> 批量删除
                              </button>
                            </div>
                          </div>
                        )}

                        {/* Search Input */}
                        <div className="relative">
                          <Search className="w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 text-zinc-400" />
                          <input
                            type="text"
                            value={presetEntrySearchQuery}
                            onChange={(e) => setPresetEntrySearchQuery(e.target.value)}
                            placeholder="搜索定位预设里的各个条目名称或具体内容..."
                            className="w-full pl-9 pr-4 py-1.5 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400 dark:focus:ring-zinc-500"
                          />
                        </div>

                        {/* Entries List */}
                        <div className="flex-1 overflow-y-auto space-y-3 pr-1">
                          {filteredEntries.length === 0 ? (
                            <div className="text-center py-12 text-xs text-zinc-400">
                              {allEntries.length === 0
                                ? '当前预设无任何条目，点击“添加条目”新建'
                                : `未找到匹配 “${presetEntrySearchQuery}” 的预设条目`}
                            </div>
                          ) : (
                            filteredEntries.map((item) => {
                              const isSelected = selectedPresetEntryIndices.includes(item.index);

                              return (
                                <div
                                  key={item.index}
                                  className={`p-3.5 bg-zinc-50 dark:bg-zinc-800/50 border rounded-xl space-y-2 transition-all ${
                                    isSelected
                                      ? 'border-zinc-900 dark:border-zinc-100 bg-zinc-100/80 dark:bg-zinc-800/80 shadow-sm'
                                      : 'border-zinc-200 dark:border-zinc-700/60 hover:border-zinc-300 dark:hover:border-zinc-600'
                                  }`}
                                >
                                  {/* Item Header */}
                                  <div className="flex items-center justify-between gap-2">
                                    <div className="flex items-center gap-2 min-w-0 flex-1">
                                      {presetEntryBatchMode && (
                                        <button
                                          type="button"
                                          onClick={() => {
                                            if (isSelected) {
                                              setSelectedPresetEntryIndices((prev) => prev.filter((i) => i !== item.index));
                                            } else {
                                              setSelectedPresetEntryIndices((prev) => [...prev, item.index]);
                                            }
                                          }}
                                          className="text-zinc-600 dark:text-zinc-300 hover:opacity-80 flex-shrink-0"
                                        >
                                          {isSelected ? (
                                            <CheckSquare className="w-4 h-4 text-zinc-900 dark:text-zinc-100" />
                                          ) : (
                                            <Square className="w-4 h-4 text-zinc-400" />
                                          )}
                                        </button>
                                      )}
                                      <span className="px-1.5 py-0.5 text-[10px] font-bold rounded bg-zinc-200 dark:bg-zinc-700/80 text-zinc-600 dark:text-zinc-300 flex-shrink-0 font-mono">
                                        #{item.index + 1}
                                      </span>
                                      <div className="text-xs font-bold text-zinc-900 dark:text-zinc-100 truncate">
                                        {item.name}
                                      </div>
                                    </div>

                                    {/* Item Operations Bar */}
                                    <div className="flex items-center gap-1 flex-shrink-0">
                                      {/* Move Up */}
                                      <button
                                        type="button"
                                        disabled={item.index === 0}
                                        onClick={() => {
                                          if (item.index === 0) return;
                                          const newList = [...allEntries];
                                          const temp = newList[item.index];
                                          newList[item.index] = newList[item.index - 1];
                                          newList[item.index - 1] = temp;
                                          savePresetEntriesList(editingPreset, newList);
                                          showToast('已向上移动条目', 'success');
                                        }}
                                        className="p-1.5 text-zinc-500 hover:text-zinc-900 dark:hover:text-zinc-100 rounded-lg hover:bg-white dark:hover:bg-zinc-800 disabled:opacity-30 transition-colors"
                                        title="向上移动"
                                      >
                                        <ChevronUp className="w-3.5 h-3.5" />
                                      </button>

                                      {/* Move Down */}
                                      <button
                                        type="button"
                                        disabled={item.index === allEntries.length - 1}
                                        onClick={() => {
                                          if (item.index === allEntries.length - 1) return;
                                          const newList = [...allEntries];
                                          const temp = newList[item.index];
                                          newList[item.index] = newList[item.index + 1];
                                          newList[item.index + 1] = temp;
                                          savePresetEntriesList(editingPreset, newList);
                                          showToast('已向下移动条目', 'success');
                                        }}
                                        className="p-1.5 text-zinc-500 hover:text-zinc-900 dark:hover:text-zinc-100 rounded-lg hover:bg-white dark:hover:bg-zinc-800 disabled:opacity-30 transition-colors"
                                        title="向下移动"
                                      >
                                        <ChevronDown className="w-3.5 h-3.5" />
                                      </button>

                                      {/* Edit */}
                                      <button
                                        type="button"
                                        onClick={() => setEditingPresetEntryModal({ index: item.index, name: item.name, content: item.content })}
                                        className="p-1.5 text-zinc-500 hover:text-zinc-900 dark:hover:text-zinc-100 rounded-lg hover:bg-white dark:hover:bg-zinc-800 transition-colors"
                                        title="编辑此条目"
                                      >
                                        <Edit3 className="w-3.5 h-3.5" />
                                      </button>

                                      {/* Transfer / Move to other Preset */}
                                      <button
                                        type="button"
                                        onClick={() => {
                                          setMovingEntryIndices([item.index]);
                                          setTargetMovePresetId('');
                                          setShowMovePresetEntryModal(true);
                                        }}
                                        className="p-1.5 text-zinc-500 hover:text-zinc-900 dark:hover:text-zinc-100 rounded-lg hover:bg-white dark:hover:bg-zinc-800 transition-colors"
                                        title="移动到其他已上传的预设"
                                      >
                                        <ArrowRightLeft className="w-3.5 h-3.5" />
                                      </button>

                                      {/* Delete */}
                                      <button
                                        type="button"
                                        onClick={() => {
                                          if (!confirm(`确定要删除条目 “${item.name}” 吗？`)) return;
                                          const newList = allEntries.filter((_, idx) => idx !== item.index);
                                          savePresetEntriesList(editingPreset, newList);
                                          showToast('已删除条目', 'info');
                                        }}
                                        className="p-1.5 text-zinc-400 hover:text-rose-500 rounded-lg hover:bg-white dark:hover:bg-zinc-800 transition-colors"
                                        title="删除此条目"
                                      >
                                        <Trash2 className="w-3.5 h-3.5" />
                                      </button>
                                    </div>
                                  </div>

                                  {/* Item Content Preview */}
                                  <div className="whitespace-pre-wrap break-words font-mono text-xs leading-relaxed max-h-36 overflow-y-auto p-2.5 bg-white dark:bg-zinc-900/80 rounded-lg text-zinc-700 dark:text-zinc-300 border border-zinc-200/60 dark:border-zinc-800/80">
                                    {item.content || <span className="text-zinc-400 italic">（空内容）</span>}
                                  </div>
                                </div>
                              );
                            })
                          )}
                        </div>
                      </>
                    );
                  })()}
                </div>
              )}

              {editingPresetTab === 'regex' && (
                /* Tab 3: 正则内容 (识别与编辑，支持上传正则) */
                <div className="space-y-4">
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-bold text-zinc-800 dark:text-zinc-200">
                      正则脚本 ({getPresetRegexScripts(editingPreset).length})
                    </span>
                    <div className="flex items-center gap-2">
                      <button
                        type="button"
                        onClick={() => presetRegexFileInputRef.current?.click()}
                        className="px-3 py-1.5 text-xs font-bold rounded-lg border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800 flex items-center gap-1.5"
                      >
                        <Upload className="w-3.5 h-3.5" /> 上传正则 (.json)
                      </button>
                      <button
                        type="button"
                        onClick={() => {
                          const currentRegex = getPresetRegexScripts(editingPreset);
                          const updated = [...currentRegex, {
                            scriptName: '新正则', script_name: '新正则', name: '新正则',
                            findRegex: '', find_regex: '', pattern: '',
                            replaceString: '', replace_string: '', replacement: '',
                            disabled: false,
                          }];
                          setEditingPreset((p) => p ? { ...p, regexScripts: updated } : null);
                        }}
                        className="px-3 py-1.5 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:opacity-90 flex items-center gap-1"
                      >
                        <Plus className="w-3.5 h-3.5" /> 添加正则
                      </button>
                    </div>
                  </div>

                  {getPresetRegexScripts(editingPreset).length === 0 ? (
                    <div className="text-center py-12 space-y-3 bg-zinc-50 dark:bg-zinc-800/30 border border-dashed border-zinc-200 dark:border-zinc-800 rounded-2xl">
                      <p className="text-xs text-zinc-400">当前预设未自带正则内容</p>
                      <button
                        type="button"
                        onClick={() => presetRegexFileInputRef.current?.click()}
                        className="px-4 py-2 text-xs font-bold rounded-xl bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:opacity-90 inline-flex items-center gap-2 shadow-sm"
                      >
                        <Upload className="w-3.5 h-3.5" /> 上传并自动识别正则
                      </button>
                    </div>
                  ) : (
                    <div className="space-y-3">
                      {getPresetRegexScripts(editingPreset).map((rx: any, idx: number) => {
                        const scriptName = rx.scriptName || rx.script_name || rx.name || rx.title || `正则 #${idx + 1}`;
                        const findRegex = rx.findRegex || rx.find_regex || rx.pattern || rx.find || rx.regex || '';
                        const replaceString = rx.replaceString || rx.replace_string || rx.replacement || rx.replace || '';

                        return (
                          <div key={idx} className="p-3.5 bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-xl space-y-2.5">
                            <div className="flex items-center justify-between gap-2">
                              <input
                                type="text"
                                value={scriptName}
                                onChange={(e) => {
                                  const val = e.target.value;
                                  const currentRegex = [...getPresetRegexScripts(editingPreset)];
                                  currentRegex[idx] = { ...currentRegex[idx], scriptName: val, script_name: val, name: val };
                                  setEditingPreset((p) => p ? { ...p, regexScripts: currentRegex } : null);
                                }}
                                placeholder="脚本名称..."
                                className="px-2.5 py-1 text-xs font-bold bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-700 rounded-lg text-zinc-900 dark:text-zinc-100 flex-1"
                              />
                              <button
                                type="button"
                                onClick={() => {
                                  const currentRegex = [...getPresetRegexScripts(editingPreset)];
                                  currentRegex.splice(idx, 1);
                                  setEditingPreset((p) => p ? { ...p, regexScripts: currentRegex } : null);
                                }}
                                className="text-xs text-rose-500 hover:text-rose-700 p-1"
                              >
                                <Trash2 className="w-3.5 h-3.5" />
                              </button>
                            </div>

                            <div>
                              <label className="text-[10px] font-semibold text-amber-600 dark:text-amber-400 block mb-0.5">查找 (Find Pattern)</label>
                              <textarea
                                rows={2}
                                value={findRegex}
                                onChange={(e) => {
                                  const val = e.target.value;
                                  const currentRegex = [...getPresetRegexScripts(editingPreset)];
                                  currentRegex[idx] = { ...currentRegex[idx], findRegex: val, find_regex: val, pattern: val };
                                  setEditingPreset((p) => p ? { ...p, regexScripts: currentRegex } : null);
                                }}
                                placeholder="查找正则表达式..."
                                className="w-full px-2.5 py-1.5 text-xs font-mono bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-700 rounded-lg text-amber-600 dark:text-amber-400 focus:outline-none resize-y"
                              />
                            </div>

                            <div>
                              <label className="text-[10px] font-semibold text-emerald-600 dark:text-emerald-400 block mb-0.5">替换 (Replacement String)</label>
                              <textarea
                                rows={3}
                                value={replaceString}
                                onChange={(e) => {
                                  const val = e.target.value;
                                  const currentRegex = [...getPresetRegexScripts(editingPreset)];
                                  currentRegex[idx] = { ...currentRegex[idx], replaceString: val, replace_string: val, replacement: val };
                                  setEditingPreset((p) => p ? { ...p, regexScripts: currentRegex } : null);
                                }}
                                placeholder="替换内容..."
                                className="w-full px-2.5 py-1.5 text-xs font-mono bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-700 rounded-lg text-emerald-600 dark:text-emerald-400 focus:outline-none resize-y"
                              />
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  )}
                </div>
              )}
            </div>

            {/* Modal Bottom Action Bar (Req 3: 取消, 确认, 导出 json格式, 删除, 更新) */}
            <div className="px-6 py-4 border-t border-zinc-200 dark:border-zinc-800 flex flex-wrap items-center justify-between gap-2 flex-shrink-0 bg-zinc-50/50 dark:bg-zinc-900/50">
              <div className="flex items-center gap-2">
                <button
                  onClick={handleDeleteEditingPreset}
                  className="px-3.5 py-2 text-xs font-bold rounded-lg bg-rose-600 text-white hover:bg-rose-700 flex items-center gap-1"
                >
                  <Trash2 className="w-3.5 h-3.5" /> 删除
                </button>
                <button
                  onClick={() => presetUpdateInputRef.current?.click()}
                  className="px-3.5 py-2 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800 flex items-center gap-1.5"
                  title="上传更新后的预设，可以覆盖之前的数据"
                >
                  <RefreshCw className="w-3.5 h-3.5" /> 更新
                </button>
              </div>

              <div className="flex items-center gap-2">
                <button
                  onClick={handleExportEditingPreset}
                  className="px-3.5 py-2 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800 flex items-center gap-1.5"
                >
                  <Download className="w-3.5 h-3.5" /> 导出 (JSON)
                </button>
                <button
                  onClick={() => setEditingPreset(null)}
                  className="px-3.5 py-2 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800"
                >
                  取消
                </button>
                <button
                  onClick={handleSaveEditingPreset}
                  className="px-4 py-2 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:bg-zinc-800 dark:hover:bg-zinc-200"
                >
                  确认
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ==================== EXPANDED ENTRY ZOOM/EDIT MODAL ==================== */}
      {expandedPresetEntry && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 animate-in fade-in">
          <div className="w-full max-w-2xl bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl p-6 space-y-4">
            <div className="flex items-center justify-between border-b border-zinc-100 dark:border-zinc-800 pb-3">
              <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100 truncate">
                条目详情: {expandedPresetEntry.key}
              </h3>
              <button
                onClick={() => setExpandedPresetEntry(null)}
                className="p-1 rounded-lg text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-200"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div>
              <label className="block text-xs font-semibold text-zinc-500 mb-1">条目内容数据 (支持查看与修改)</label>
              <textarea
                rows={12}
                value={
                  typeof expandedPresetEntry.value === 'object'
                    ? JSON.stringify(expandedPresetEntry.value, null, 2)
                    : String(expandedPresetEntry.value)
                }
                onChange={(e) => {
                  const val = e.target.value;
                  let parsedVal: any = val;
                  try {
                    parsedVal = JSON.parse(val);
                  } catch {
                    // plain text string
                  }
                  setExpandedPresetEntry((prev) => (prev ? { ...prev, value: parsedVal } : null));
                }}
                className="w-full p-4 font-mono text-xs bg-zinc-950 text-zinc-100 border border-zinc-800 rounded-xl focus:outline-none resize-none leading-relaxed"
              />
            </div>

            <div className="flex justify-end gap-2 pt-2 border-t border-zinc-100 dark:border-zinc-800">
              <button
                onClick={() => setExpandedPresetEntry(null)}
                className="px-4 py-2 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800"
              >
                关闭
              </button>
              <button
                onClick={() => {
                  if (!editingPreset) return;
                  const key = expandedPresetEntry.key;
                  const newValue = expandedPresetEntry.value;

                  let updatedJsonData = { ...editingPreset.jsonData };

                  if (Array.isArray(editingPreset.jsonData)) {
                    updatedJsonData = editingPreset.jsonData.map((item, idx) => {
                      const itemKey = item?.name || item?.identifier || item?.role || `条目 #${idx + 1}`;
                      if (itemKey === key) return typeof newValue === 'object' ? newValue : { ...item, content: newValue };
                      return item;
                    });
                  } else if (editingPreset.jsonData && typeof editingPreset.jsonData === 'object') {
                    updatedJsonData[key] = newValue;
                  }

                  setEditingPreset({
                    ...editingPreset,
                    jsonData: updatedJsonData,
                    rawJsonString: JSON.stringify(updatedJsonData, null, 2),
                  });
                  setExpandedPresetEntry(null);
                  showToast('条目内容修改已更新', 'success');
                }}
                className="px-4 py-2 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:bg-zinc-800 dark:hover:bg-zinc-200"
              >
                保存此条目修改
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================== ADD PRESET ENTRY MODAL ==================== */}
      {showAddPresetEntryModal && editingPreset && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 animate-in fade-in">
          <div className="w-full max-w-lg bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl p-6 space-y-4">
            <div className="flex items-center justify-between border-b border-zinc-100 dark:border-zinc-800 pb-3">
              <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100">
                新建预设条目
              </h3>
              <button
                type="button"
                onClick={() => setShowAddPresetEntryModal(false)}
                className="p-1 rounded-lg text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-200"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="space-y-3">
              <div>
                <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                  条目名称 <span className="text-rose-500">*</span>
                </label>
                <input
                  type="text"
                  value={newPresetEntryForm.name}
                  onChange={(e) => setNewPresetEntryForm((prev) => ({ ...prev, name: e.target.value }))}
                  placeholder="请输入条目名称 (如: 主系统提示词, Role, Style)..."
                  className="w-full px-3 py-2 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400 text-zinc-900 dark:text-zinc-100"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                  条目内容
                </label>
                <textarea
                  rows={8}
                  value={newPresetEntryForm.content}
                  onChange={(e) => setNewPresetEntryForm((prev) => ({ ...prev, content: e.target.value }))}
                  placeholder="请输入条目的文本内容或段落..."
                  className="w-full p-3 font-mono text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-xl focus:outline-none focus:ring-1 focus:ring-zinc-400 text-zinc-900 dark:text-zinc-100 resize-none leading-relaxed"
                />
              </div>
            </div>

            <div className="flex justify-end gap-2 pt-2 border-t border-zinc-100 dark:border-zinc-800">
              <button
                type="button"
                onClick={() => setShowAddPresetEntryModal(false)}
                className="px-4 py-2 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800"
              >
                取消
              </button>
              <button
                type="button"
                disabled={!newPresetEntryForm.name.trim()}
                onClick={() => {
                  if (!newPresetEntryForm.name.trim()) return;
                  const allEntries = getPresetEntriesList(editingPreset);
                  const newList = [...allEntries, { name: newPresetEntryForm.name.trim(), content: newPresetEntryForm.content }];
                  savePresetEntriesList(editingPreset, newList);
                  setShowAddPresetEntryModal(false);
                  setNewPresetEntryForm({ name: '', content: '' });
                  showToast('成功添加新条目', 'success');
                }}
                className="px-4 py-2 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:opacity-90 disabled:opacity-40"
              >
                保存条目
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================== EDIT PRESET ENTRY MODAL ==================== */}
      {editingPresetEntryModal && editingPreset && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 animate-in fade-in">
          <div className="w-full max-w-lg bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl p-6 space-y-4">
            <div className="flex items-center justify-between border-b border-zinc-100 dark:border-zinc-800 pb-3">
              <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100">
                编辑条目 #{editingPresetEntryModal.index + 1}
              </h3>
              <button
                type="button"
                onClick={() => setEditingPresetEntryModal(null)}
                className="p-1 rounded-lg text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-200"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="space-y-3">
              <div>
                <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                  条目名称
                </label>
                <input
                  type="text"
                  value={editingPresetEntryModal.name}
                  onChange={(e) => setEditingPresetEntryModal((prev) => prev ? { ...prev, name: e.target.value } : null)}
                  placeholder="条目名称..."
                  className="w-full px-3 py-2 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400 text-zinc-900 dark:text-zinc-100"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                  条目内容
                </label>
                <textarea
                  rows={8}
                  value={editingPresetEntryModal.content}
                  onChange={(e) => setEditingPresetEntryModal((prev) => prev ? { ...prev, content: e.target.value } : null)}
                  placeholder="条目具体文本内容..."
                  className="w-full p-3 font-mono text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-xl focus:outline-none focus:ring-1 focus:ring-zinc-400 text-zinc-900 dark:text-zinc-100 resize-none leading-relaxed"
                />
              </div>
            </div>

            <div className="flex justify-end gap-2 pt-2 border-t border-zinc-100 dark:border-zinc-800">
              <button
                type="button"
                onClick={() => setEditingPresetEntryModal(null)}
                className="px-4 py-2 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800"
              >
                取消
              </button>
              <button
                type="button"
                onClick={() => {
                  if (!editingPresetEntryModal) return;
                  const allEntries = getPresetEntriesList(editingPreset);
                  const newList = [...allEntries];
                  newList[editingPresetEntryModal.index] = {
                    ...newList[editingPresetEntryModal.index],
                    name: editingPresetEntryModal.name,
                    content: editingPresetEntryModal.content,
                  };
                  savePresetEntriesList(editingPreset, newList);
                  setEditingPresetEntryModal(null);
                  showToast('条目更新已保存', 'success');
                }}
                className="px-4 py-2 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:opacity-90"
              >
                保存修改
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================== MOVE ENTRY TO OTHER PRESET MODAL ==================== */}
      {showMovePresetEntryModal && editingPreset && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 animate-in fade-in">
          <div className="w-full max-w-md bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl p-6 space-y-4">
            <div className="flex items-center justify-between border-b border-zinc-100 dark:border-zinc-800 pb-3">
              <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100">
                跨预设移动条目 ({movingEntryIndices.length} 个)
              </h3>
              <button
                type="button"
                onClick={() => setShowMovePresetEntryModal(false)}
                className="p-1 rounded-lg text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-200"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="space-y-3">
              <p className="text-xs text-zinc-500">
                选择要将选中的条目移动/复制到的目标预设：
              </p>
              <div>
                <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                  目标预设 <span className="text-rose-500">*</span>
                </label>
                <select
                  value={targetMovePresetId}
                  onChange={(e) => setTargetMovePresetId(e.target.value)}
                  className="w-full px-3 py-2 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400 text-zinc-900 dark:text-zinc-100"
                >
                  <option value="">-- 请选择目标预设 --</option>
                  {(appData.presets || [])
                    .filter((p) => p.id !== editingPreset.id)
                    .map((p) => (
                      <option key={p.id} value={p.id}>
                        {p.fileName || p.name} ({p.category || '默认'})
                      </option>
                    ))}
                </select>
              </div>
            </div>

            <div className="flex justify-end gap-2 pt-2 border-t border-zinc-100 dark:border-zinc-800">
              <button
                type="button"
                onClick={() => setShowMovePresetEntryModal(false)}
                className="px-4 py-2 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800"
              >
                取消
              </button>
              <button
                type="button"
                disabled={!targetMovePresetId}
                onClick={() => {
                  if (!targetMovePresetId) return;
                  const targetPreset = (appData.presets || []).find((p) => p.id === targetMovePresetId);
                  if (!targetPreset) return;

                  const currentEntries = getPresetEntriesList(editingPreset);
                  const itemsToMove = currentEntries.filter((e) => movingEntryIndices.includes(e.index));
                  const remainingEntries = currentEntries.filter((e) => !movingEntryIndices.includes(e.index));

                  // Append to target preset
                  const targetEntries = getPresetEntriesList(targetPreset);
                  const updatedTargetEntries = [...targetEntries, ...itemsToMove];

                  // Save remaining in current preset
                  savePresetEntriesList(editingPreset, remainingEntries);

                  // Save target preset
                  savePresetEntriesList(targetPreset, updatedTargetEntries);

                  setShowMovePresetEntryModal(false);
                  setSelectedPresetEntryIndices([]);
                  showToast(`已成功将 ${itemsToMove.length} 个条目移动到 “${targetPreset.name}”`, 'success');
                }}
                className="px-4 py-2 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:opacity-90 disabled:opacity-40"
              >
                确认移动
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Hidden File Inputs for ST Presets */}
      <input
        type="file"
        ref={presetFileInputRef}
        multiple
        accept=".json,application/json"
        className="hidden"
        onChange={(e) => {
          if (e.target.files) handlePresetFileUpload(e.target.files);
          e.target.value = '';
        }}
      />

      <input
        type="file"
        ref={presetUpdateInputRef}
        accept=".json,application/json"
        className="hidden"
        onChange={(e) => {
          const file = e.target.files?.[0];
          if (file) handleUpdateEditingPresetFile(file);
          e.target.value = '';
        }}
      />

      {/* Hidden File Input for ST Plugins & Scripts */}
      <input
        type="file"
        ref={pluginScriptFileInputRef}
        multiple
        accept=".json,application/json"
        className="hidden"
        onChange={(e) => {
          if (e.target.files) handleScriptFileUpload(e.target.files);
          e.target.value = '';
        }}
      />

      {/* ==================== ST PLUGINS: ADD TYPE SELECTION MODAL ==================== */}
      {showAddPluginTypeModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 animate-in fade-in pt-[calc(1rem+env(safe-area-inset-top,0px))]">
          <div className="w-full max-w-sm bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl p-5 space-y-4">
            <div className="flex items-center justify-between border-b border-zinc-100 dark:border-zinc-800 pb-3">
              <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100">选择添加模式</h3>
              <button
                onClick={() => setShowAddPluginTypeModal(false)}
                className="p-1 rounded-lg text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-200"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="space-y-2.5">
              <button
                onClick={() => {
                  setShowAddPluginTypeModal(false);
                  setShowAddPluginModal(true);
                }}
                className="w-full p-3 bg-zinc-50 dark:bg-zinc-800/80 hover:bg-zinc-100 dark:hover:bg-zinc-800 border border-zinc-200 dark:border-zinc-700/60 rounded-xl text-left transition-all flex items-center justify-between group"
              >
                <div>
                  <div className="text-xs font-bold text-zinc-900 dark:text-zinc-100">添加插件</div>
                  <div className="text-[10px] text-zinc-400">填写插件名称、插件链接、联系方式及描述</div>
                </div>
                <Plus className="w-4 h-4 text-zinc-400 group-hover:text-zinc-900 dark:group-hover:text-zinc-100" />
              </button>

              <button
                onClick={() => {
                  setShowAddPluginTypeModal(false);
                  pluginScriptFileInputRef.current?.click();
                }}
                className="w-full p-3 bg-zinc-50 dark:bg-zinc-800/80 hover:bg-zinc-100 dark:hover:bg-zinc-800 border border-zinc-200 dark:border-zinc-700/60 rounded-xl text-left transition-all flex items-center justify-between group"
              >
                <div>
                  <div className="text-xs font-bold text-zinc-900 dark:text-zinc-100">导入脚本</div>
                  <div className="text-[10px] text-zinc-400">选择并上传 .json 格式的脚本文件</div>
                </div>
                <Upload className="w-4 h-4 text-zinc-400 group-hover:text-zinc-900 dark:group-hover:text-zinc-100" />
              </button>
            </div>

            <div className="pt-2 flex justify-end">
              <button
                onClick={() => setShowAddPluginTypeModal(false)}
                className="px-4 py-1.5 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800"
              >
                取消
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================== ST PLUGINS: ADD PLUGIN LINK MODAL ==================== */}
      {showAddPluginModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 animate-in fade-in pt-[calc(1rem+env(safe-area-inset-top,0px))]">
          <div className="w-full max-w-sm bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl p-5 space-y-4">
            <div className="flex items-center justify-between border-b border-zinc-100 dark:border-zinc-800 pb-3">
              <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100">添加插件</h3>
              <button
                onClick={() => setShowAddPluginModal(false)}
                className="p-1 rounded-lg text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-200"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="space-y-3">
              <div>
                <label className="text-xs font-semibold text-zinc-700 dark:text-zinc-300 block mb-1">
                  插件名称 <span className="text-rose-500">*</span>
                </label>
                <input
                  type="text"
                  value={pluginForm.name}
                  onChange={(e) => setPluginForm({ ...pluginForm, name: e.target.value })}
                  placeholder="输入插件名称"
                  className="w-full px-3 py-2 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-xl text-zinc-900 dark:text-zinc-100 focus:outline-none"
                />
              </div>

              <div>
                <label className="text-xs font-semibold text-zinc-700 dark:text-zinc-300 block mb-1">
                  插件地址 <span className="text-rose-500">*</span>
                </label>
                <input
                  type="text"
                  value={pluginForm.url}
                  onChange={(e) => setPluginForm({ ...pluginForm, url: e.target.value })}
                  placeholder="输入插件 URL / 下载链接"
                  className="w-full px-3 py-2 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-xl text-zinc-900 dark:text-zinc-100 focus:outline-none"
                />
              </div>

              <div>
                <label className="text-xs font-semibold text-zinc-700 dark:text-zinc-300 block mb-1">
                  联系方式 <span className="text-zinc-400 font-normal">(DC链接/QQ群等，选填)</span>
                </label>
                <input
                  type="text"
                  value={pluginForm.contact}
                  onChange={(e) => setPluginForm({ ...pluginForm, contact: e.target.value })}
                  placeholder="例如: Discord链接 / QQ群"
                  className="w-full px-3 py-2 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-xl text-zinc-900 dark:text-zinc-100 focus:outline-none"
                />
              </div>

              <div>
                <label className="text-xs font-semibold text-zinc-700 dark:text-zinc-300 block mb-1">
                  描述 <span className="text-zinc-400 font-normal">(选填)</span>
                </label>
                <textarea
                  rows={3}
                  value={pluginForm.description}
                  onChange={(e) => setPluginForm({ ...pluginForm, description: e.target.value })}
                  placeholder="填写插件功能或用途说明..."
                  className="w-full px-3 py-2 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-xl text-zinc-900 dark:text-zinc-100 focus:outline-none resize-none"
                />
              </div>
            </div>

            <div className="pt-2 flex justify-end gap-2 border-t border-zinc-100 dark:border-zinc-800">
              <button
                onClick={() => setShowAddPluginModal(false)}
                className="px-4 py-2 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800"
              >
                取消
              </button>
              <button
                onClick={handleAddPluginSubmit}
                className="px-4 py-2 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:bg-zinc-800 dark:hover:bg-zinc-200"
              >
                确定
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================== ST PLUGINS: SINGLE-PAGE DETAIL / EDIT MODAL ==================== */}
      {editingPlugin && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 animate-in fade-in pt-[calc(1rem+env(safe-area-inset-top,0px))] pb-[calc(1rem+env(safe-area-inset-bottom,0px))]">
          <div className="w-full max-w-3xl h-[85vh] min-h-[550px] max-h-[750px] bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl flex flex-col overflow-hidden">
            {/* Header */}
            <div className="px-6 py-4 border-b border-zinc-200 dark:border-zinc-800 flex items-center justify-between flex-shrink-0 bg-zinc-50/50 dark:bg-zinc-900/50">
              <div className="flex items-center gap-2">
                {editingPlugin.type === 'script' ? (
                  <span className="px-2 py-0.5 rounded text-xs font-bold bg-emerald-100 text-emerald-800 dark:bg-emerald-950/80 dark:text-emerald-300">
                    脚本详情
                  </span>
                ) : (
                  <span className="px-2 py-0.5 rounded text-xs font-bold bg-indigo-100 text-indigo-800 dark:bg-indigo-950/80 dark:text-indigo-300">
                    插件详情
                  </span>
                )}
                <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100 truncate max-w-[240px]">
                  {editingPlugin.name}
                </h3>
              </div>
              <button
                onClick={() => setEditingPlugin(null)}
                className="p-1 rounded-lg text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-200"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Modal Body: 插件保持原单页；脚本独立使用横向 Tab */}
            {editingPlugin.type === 'script' ? (
              <div className="flex-1 min-h-0 flex flex-col overflow-hidden">
                {/* Script Tabs */}
                <div className="px-6 pt-3 border-b border-zinc-200 dark:border-zinc-800 flex items-center gap-1 flex-shrink-0 bg-white dark:bg-zinc-900">
                  <button
                    type="button"
                    onClick={() => setScriptDetailTab('info')}
                    className={`px-4 py-2.5 text-xs font-semibold border-b-2 transition-colors ${
                      scriptDetailTab === 'info'
                        ? 'border-zinc-900 text-zinc-900 dark:border-zinc-100 dark:text-zinc-100'
                        : 'border-transparent text-zinc-400 hover:text-zinc-700 dark:hover:text-zinc-200'
                    }`}
                  >
                    脚本信息
                  </button>
                  <button
                    type="button"
                    onClick={() => setScriptDetailTab('code')}
                    className={`px-4 py-2.5 text-xs font-semibold border-b-2 transition-colors ${
                      scriptDetailTab === 'code'
                        ? 'border-zinc-900 text-zinc-900 dark:border-zinc-100 dark:text-zinc-100'
                        : 'border-transparent text-zinc-400 hover:text-zinc-700 dark:hover:text-zinc-200'
                    }`}
                  >
                    脚本内容
                  </button>
                </div>

                {scriptDetailTab === 'info' ? (
                  <div className="p-6 overflow-y-auto space-y-4 flex-1">
                    <div>
                      <label className="text-xs font-semibold text-zinc-700 dark:text-zinc-300 block mb-1">脚本名称 <span className="text-rose-500">*</span></label>
                      <input
                        type="text"
                        value={editingPlugin.name}
                        onChange={(e) => setEditingPlugin({ ...editingPlugin, name: e.target.value })}
                        className="w-full px-3 py-2 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-xl text-zinc-900 dark:text-zinc-100 focus:outline-none"
                      />
                    </div>

                    <div>
                      <label className="text-xs font-semibold text-zinc-700 dark:text-zinc-300 block mb-1">作者栏</label>
                      <input
                        type="text"
                        value={editingPlugin.author || editingPlugin.contact || ''}
                        onChange={(e) => setEditingPlugin({ ...editingPlugin, author: e.target.value, contact: e.target.value })}
                        placeholder="填写作者名称或联系信息"
                        className="w-full px-3 py-2 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-xl text-zinc-900 dark:text-zinc-100 focus:outline-none"
                      />
                    </div>

                    <div>
                      <label className="text-xs font-semibold text-zinc-700 dark:text-zinc-300 block mb-1">分类</label>
                      <select
                        value={editingPlugin.category || '默认'}
                        onChange={(e) => setEditingPlugin({ ...editingPlugin, category: e.target.value })}
                        className="w-full px-3 py-2 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-xl text-zinc-900 dark:text-zinc-100 focus:outline-none"
                      >
                        {Array.from(new Set(['默认', ...(appData.pluginCategories || []), ...pluginsList.map((p) => p.category || '默认')])).map((cat) => (
                          <option key={cat} value={cat}>{cat}</option>
                        ))}
                      </select>
                    </div>

                    <div>
                      <label className="text-xs font-semibold text-zinc-700 dark:text-zinc-300 block mb-1">来源 <span className="text-zinc-400 font-normal">(DC链接/QQ群等)</span></label>
                      <input
                        type="text"
                        value={editingPlugin.source || ''}
                        onChange={(e) => setEditingPlugin({ ...editingPlugin, source: e.target.value })}
                        placeholder="例如: Discord链接 / QQ群"
                        className="w-full px-3 py-2 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-xl text-zinc-900 dark:text-zinc-100 focus:outline-none"
                      />
                    </div>

                    <div>
                      <label className="text-xs font-semibold text-zinc-700 dark:text-zinc-300 block mb-1">描述</label>
                      <textarea
                        rows={4}
                        value={editingPlugin.description || ''}
                        onChange={(e) => setEditingPlugin({ ...editingPlugin, description: e.target.value })}
                        placeholder="填写脚本的详细说明..."
                        className="w-full px-3 py-2 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-xl text-zinc-900 dark:text-zinc-100 focus:outline-none resize-none leading-relaxed"
                      />
                    </div>

                    {/* 只有脚本拥有此按钮；插件不会出现 */}
                    <div className="pt-3 border-t border-zinc-200 dark:border-zinc-800">
                      <button
                        type="button"
                        onClick={() => handleExportPluginScriptJson(editingPlugin)}
                        className="px-3.5 py-2 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800 flex items-center gap-1.5"
                      >
                        <Download className="w-3.5 h-3.5" /> 导出 JSON
                      </button>
                    </div>
                  </div>
                ) : (
                  <div className="flex-1 min-h-0 p-4 flex flex-col gap-3 overflow-hidden">
                    <div className="flex items-center justify-between gap-3 flex-shrink-0">
                      <h3 className="text-xs font-bold text-zinc-800 dark:text-zinc-200 flex items-center gap-1.5">
                        <FileCode className="w-4 h-4 text-zinc-500" /> 脚本内容
                      </h3>
                      <button
                        type="button"
                        onClick={() => setIsScriptCodeExpanded(true)}
                        className="px-3 py-1.5 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800 flex items-center gap-1.5 whitespace-nowrap"
                      >
                        <Maximize2 className="w-3.5 h-3.5" /> 放大展开
                      </button>
                    </div>

                    <div className="relative w-full flex-shrink-0">
                      <Search className="w-3.5 h-3.5 absolute left-3 top-1/2 -translate-y-1/2 text-zinc-400" />
                      <input
                        type="text"
                        value={scriptSearchQuery}
                        onChange={(e) => setScriptSearchQuery(e.target.value)}
                        placeholder="搜索/定位脚本内容..."
                        className="w-full pl-9 pr-4 py-1.5 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400 dark:focus:ring-zinc-500"
                      />
                    </div>

                    <div className="flex-1 min-h-0 relative rounded-xl border border-zinc-200 dark:border-zinc-800 overflow-hidden bg-zinc-950">
                      <textarea
                        value={scriptContentDraft}
                        onChange={(e) => setScriptContentDraft(e.target.value)}
                        spellCheck={false}
                        className="w-full h-full p-4 font-mono text-xs bg-zinc-950 text-zinc-100 focus:outline-none resize-none leading-relaxed"
                        placeholder="暂无脚本 JSON 内容"
                      />
                    </div>

                    {scriptSearchQuery.trim() && (
                      <div className="text-[11px] text-zinc-500 dark:text-zinc-400 flex-shrink-0">
                        内容中匹配关键词“{scriptSearchQuery}”的频次：{(
                          (scriptContentDraft.match(new RegExp(scriptSearchQuery.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi')) || []).length
                        )} 处
                      </div>
                    )}
                  </div>
                )}
              </div>
            ) : (
              <div className="p-6 overflow-y-auto space-y-4 flex-1">
                {/* 名称 */}
                <div>
                  <label className="text-xs font-semibold text-zinc-700 dark:text-zinc-300 block mb-1">
                    插件名称 <span className="text-rose-500">*</span>
                  </label>
                  <input
                    type="text"
                    value={editingPlugin.name}
                    onChange={(e) => setEditingPlugin({ ...editingPlugin, name: e.target.value })}
                    className="w-full px-3 py-2 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-xl text-zinc-900 dark:text-zinc-100 focus:outline-none"
                  />
                </div>

                {/* 插件地址 */}
                <div>
                  <label className="text-xs font-semibold text-zinc-700 dark:text-zinc-300 block mb-1">插件地址 <span className="text-rose-500">*</span></label>
                  <div className="flex items-center gap-2">
                    <input
                      type="text"
                      value={editingPlugin.url || ''}
                      onChange={(e) => setEditingPlugin({ ...editingPlugin, url: e.target.value })}
                      placeholder="https://..."
                      className="w-full px-3 py-2 text-xs font-mono bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-xl text-zinc-900 dark:text-zinc-100 focus:outline-none flex-1"
                    />
                    <button
                      onClick={() => {
                        if (editingPlugin.url) {
                          navigator.clipboard.writeText(editingPlugin.url);
                          showToast('插件地址已复制到剪贴板！', 'success');
                        }
                      }}
                      className="px-3 py-2 text-xs font-medium rounded-xl bg-zinc-100 dark:bg-zinc-800 hover:bg-zinc-200 dark:hover:bg-zinc-700 text-zinc-700 dark:text-zinc-300 flex items-center gap-1 flex-shrink-0"
                    >
                      <Copy className="w-3.5 h-3.5" /> 复制
                    </button>
                  </div>
                </div>

                {/* 作者栏 */}
                <div>
                  <label className="text-xs font-semibold text-zinc-700 dark:text-zinc-300 block mb-1">作者栏</label>
                  <input
                    type="text"
                    value={editingPlugin.author || editingPlugin.contact || ''}
                    onChange={(e) => setEditingPlugin({ ...editingPlugin, author: e.target.value, contact: e.target.value })}
                    placeholder="填写作者名称或联系信息"
                    className="w-full px-3 py-2 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-xl text-zinc-900 dark:text-zinc-100 focus:outline-none"
                  />
                </div>

                {/* 分类 */}
                <div>
                  <label className="text-xs font-semibold text-zinc-700 dark:text-zinc-300 block mb-1">分类</label>
                  <select
                    value={editingPlugin.category || '默认'}
                    onChange={(e) => setEditingPlugin({ ...editingPlugin, category: e.target.value })}
                    className="w-full px-3 py-2 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-xl text-zinc-900 dark:text-zinc-100 focus:outline-none"
                  >
                    {Array.from(new Set(['默认', ...(appData.pluginCategories || []), ...pluginsList.map((p) => p.category || '默认')])).map((cat) => (
                      <option key={cat} value={cat}>{cat}</option>
                    ))}
                  </select>
                </div>

                {/* 来源 */}
                <div>
                  <label className="text-xs font-semibold text-zinc-700 dark:text-zinc-300 block mb-1">来源 <span className="text-zinc-400 font-normal">(DC链接/QQ群等)</span></label>
                  <input
                    type="text"
                    value={editingPlugin.source || editingPlugin.contact || ''}
                    onChange={(e) => setEditingPlugin({ ...editingPlugin, source: e.target.value })}
                    placeholder="例如: Discord链接 / QQ群"
                    className="w-full px-3 py-2 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-xl text-zinc-900 dark:text-zinc-100 focus:outline-none"
                  />
                </div>

                {/* 描述 */}
                <div>
                  <label className="text-xs font-semibold text-zinc-700 dark:text-zinc-300 block mb-1">描述</label>
                  <textarea
                    rows={4}
                    value={editingPlugin.description || ''}
                    onChange={(e) => setEditingPlugin({ ...editingPlugin, description: e.target.value })}
                    placeholder="填写插件功能或用途说明..."
                    className="w-full px-3 py-2 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-xl text-zinc-900 dark:text-zinc-100 focus:outline-none resize-none leading-relaxed"
                  />
                </div>
              </div>
            )}

            {/* Bottom Buttons: 插件保持原样；脚本独立保存 */}
            <div className="px-6 py-4 border-t border-zinc-200 dark:border-zinc-800 flex flex-wrap items-center justify-between gap-2 flex-shrink-0 bg-zinc-50/50 dark:bg-zinc-900/50">
              <button
                onClick={() => {
                  if (!confirm(`确定要删除“${editingPlugin.name}”吗？`)) return;
                  const updated = pluginsList.filter((p) => p.id !== editingPlugin.id);
                  updateAppData({ ...appData, plugins: updated });
                  setEditingPlugin(null);
                  showToast('项目已成功删除', 'info');
                }}
                className="px-3.5 py-2 text-xs font-bold rounded-lg bg-rose-600 text-white hover:bg-rose-700 flex items-center gap-1"
              >
                <Trash2 className="w-3.5 h-3.5" /> 删除
              </button>

              <div className="flex items-center gap-2">
                <button
                  onClick={() => setEditingPlugin(null)}
                  className="px-3 py-2 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800"
                >
                  取消
                </button>

                <button
                  onClick={() => {
                    if (!editingPlugin.name.trim()) {
                      showToast('请输入名称', 'error');
                      return;
                    }
                    if (editingPlugin.type === 'plugin' && !editingPlugin.url?.trim()) {
                      showToast('请输入插件地址', 'error');
                      return;
                    }
                    if (editingPlugin.type === 'script') {
                      try {
                        JSON.parse(scriptContentDraft);
                      } catch {
                        showToast('脚本内容不是有效的 JSON，请检查后再保存', 'error');
                        return;
                      }
                    }
                    const finalPlugin = editingPlugin.type === 'script'
                      ? { ...editingPlugin, jsonData: (() => { try { return JSON.parse(scriptContentDraft); } catch { return editingPlugin.jsonData; } })(), updatedAt: Date.now() }
                      : editingPlugin;
                    const updated = pluginsList.map((p) => (p.id === finalPlugin.id ? finalPlugin : p));
                    updateAppData({ ...appData, plugins: updated });
                    setEditingPlugin(null);
                    showToast('信息已保存更新！', 'success');
                  }}
                  className="px-4 py-2 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:bg-zinc-800 dark:hover:bg-zinc-200"
                >
                  确认
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ==================== ST PLUGINS: SCRIPT CONTENT FULLSCREEN ==================== */}
      {isScriptCodeExpanded && editingPlugin?.type === 'script' && (
        <div className="fixed inset-0 z-[60] bg-white text-zinc-900 flex flex-col animate-in fade-in pt-[env(safe-area-inset-top,0px)] pb-[env(safe-area-inset-bottom,0px)] pl-[env(safe-area-inset-left,0px)] pr-[env(safe-area-inset-right,0px)]">
          <div className="px-4 sm:px-6 py-3 border-b border-zinc-200 bg-white flex items-center justify-between flex-shrink-0 min-h-[3.5rem]">
            <div className="flex items-center gap-2 min-w-0">
              <h3 className="font-bold text-sm text-zinc-900 truncate">{editingPlugin.name} - 脚本内容</h3>
              <span className="px-2 py-0.5 text-[10px] font-mono bg-zinc-100 text-zinc-600 rounded uppercase border border-zinc-200 flex-shrink-0">JSON</span>
            </div>
            <button
              type="button"
              onClick={() => setIsScriptCodeExpanded(false)}
              className="p-2 rounded-lg hover:bg-zinc-100 text-zinc-500 hover:text-zinc-900 flex-shrink-0"
              aria-label="关闭全屏脚本内容"
            >
              <X className="w-5 h-5" />
            </button>
          </div>

          <div className="px-4 sm:px-6 py-3 border-b border-zinc-200 bg-white flex-shrink-0">
            <div className="relative w-full">
              <Search className="w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 text-zinc-400" />
              <input
                type="text"
                value={scriptSearchQuery}
                onChange={(e) => setScriptSearchQuery(e.target.value)}
                placeholder="搜索/定位脚本内容..."
                className="w-full pl-10 pr-4 py-2.5 text-sm bg-white border border-zinc-300 rounded-xl text-zinc-900 placeholder:text-zinc-400 focus:outline-none focus:ring-2 focus:ring-zinc-300"
              />
            </div>
            <div className="flex items-center justify-end gap-2 mt-2">
              <button
                type="button"
                onClick={() => {
                  navigator.clipboard.writeText(scriptContentDraft);
                  showToast('脚本内容已复制到剪贴板', 'success');
                }}
                className="px-3 py-1.5 text-xs font-medium rounded-lg border border-zinc-300 text-zinc-700 hover:bg-zinc-100"
              >
                复制全文
              </button>
              <button
                type="button"
                onClick={handleSaveScriptContent}
                className="px-4 py-1.5 text-xs font-bold rounded-lg bg-zinc-900 text-white hover:bg-zinc-800"
              >
                保存修改
              </button>
              <button
                type="button"
                onClick={() => handleExportPluginScriptJson(editingPlugin)}
                className="px-3 py-1.5 text-xs font-medium rounded-lg border border-zinc-300 text-zinc-700 hover:bg-zinc-100"
              >
                导出 JSON
              </button>
            </div>
          </div>

          <div className="flex-1 min-h-0 p-3 sm:p-5 bg-white overflow-hidden pb-[calc(0.75rem+env(safe-area-inset-bottom,0px))]">
            <textarea
              value={scriptContentDraft}
              onChange={(e) => setScriptContentDraft(e.target.value)}
              spellCheck={false}
              placeholder="暂无脚本 JSON 内容"
              className="w-full h-full p-4 font-mono text-xs sm:text-sm text-zinc-900 bg-white border border-zinc-200 rounded-xl focus:outline-none focus:ring-1 focus:ring-zinc-300 leading-relaxed resize-none"
            />
          </div>
        </div>
      )}

      {/* ==================== ST PLUGINS: NEW GROUP MODAL ==================== */}
      {showNewPluginGroupModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 animate-in fade-in pt-[calc(1rem+env(safe-area-inset-top,0px))]">
          <div className="w-full max-w-sm bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl p-5 space-y-4">
            <div className="flex items-center justify-between border-b border-zinc-100 dark:border-zinc-800 pb-3">
              <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100">新建插件分组</h3>
              <button
                onClick={() => setShowNewPluginGroupModal(false)}
                className="p-1 rounded-lg text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-200"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div>
              <label className="text-xs font-semibold text-zinc-700 dark:text-zinc-300 block mb-1">
                分组名称
              </label>
              <input
                type="text"
                value={newPluginGroupName}
                onChange={(e) => setNewPluginGroupName(e.target.value)}
                placeholder="例如: 正则过滤 / 美化UI / 核心脚本..."
                className="w-full px-3 py-2 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-xl text-zinc-900 dark:text-zinc-100 focus:outline-none"
              />
            </div>

            <div className="flex justify-end gap-2 pt-2 border-t border-zinc-100 dark:border-zinc-800">
              <button
                onClick={() => setShowNewPluginGroupModal(false)}
                className="px-4 py-2 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800"
              >
                取消
              </button>
              <button
                onClick={handleCreateNewPluginCategory}
                className="px-4 py-2 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:bg-zinc-800 dark:hover:bg-zinc-200"
              >
                创建
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================== ST PLUGINS: MANAGE GROUP MODAL ==================== */}
      {managingPluginCategory && !showDeletePluginCategoryConfirm && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 animate-in fade-in pt-[calc(1rem+env(safe-area-inset-top,0px))]">
          <div className="w-full max-w-sm bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl p-5 space-y-4">
            <div className="flex items-center justify-between border-b border-zinc-100 dark:border-zinc-800 pb-3">
              <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100">
                管理分组: {managingPluginCategory}
              </h3>
              <button
                onClick={() => setManagingPluginCategory(null)}
                className="p-1 rounded-lg text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-200"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="space-y-3">
              <div>
                <label className="text-xs font-semibold text-zinc-700 dark:text-zinc-300 block mb-1">
                  重命名分组
                </label>
                <div className="flex gap-2">
                  <input
                    type="text"
                    value={renamePluginCategoryInput}
                    onChange={(e) => setRenamePluginCategoryInput(e.target.value)}
                    className="w-full px-3 py-1.5 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-xl text-zinc-900 dark:text-zinc-100 focus:outline-none flex-1"
                  />
                  <button
                    onClick={handleRenamePluginCategory}
                    className="px-3 py-1.5 text-xs font-bold rounded-xl bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:opacity-90"
                  >
                    重命名
                  </button>
                </div>
              </div>

              {managingPluginCategory !== '默认' && (
                <div className="pt-2 border-t border-zinc-100 dark:border-zinc-800">
                  <button
                    onClick={() => setShowDeletePluginCategoryConfirm(true)}
                    className="w-full py-2 text-xs font-bold text-rose-600 hover:bg-rose-50 dark:hover:bg-rose-950/40 rounded-xl border border-rose-200 dark:border-rose-900 transition-colors flex items-center justify-center gap-1.5"
                  >
                    <Trash2 className="w-3.5 h-3.5" /> 删除此分组
                  </button>
                </div>
              )}
            </div>

            <div className="flex justify-end pt-2 border-t border-zinc-100 dark:border-zinc-800">
              <button
                onClick={() => setManagingPluginCategory(null)}
                className="px-4 py-1.5 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800"
              >
                关闭
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================== ST PLUGINS: DELETE GROUP CONFIRMATION MODAL ==================== */}
      {showDeletePluginCategoryConfirm && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 animate-in fade-in pt-[calc(1rem+env(safe-area-inset-top,0px))]">
          <div className="w-full max-w-sm bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl p-5 space-y-4">
            <div className="flex items-center justify-between border-b border-zinc-100 dark:border-zinc-800 pb-3">
              <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100">确认删除分组</h3>
              <button
                onClick={() => setShowDeletePluginCategoryConfirm(false)}
                className="p-1 rounded-lg text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-200"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <p className="text-xs text-zinc-600 dark:text-zinc-300">
              确定要删除分组 “<span className="font-bold underline">{managingPluginCategory}</span>” 吗？
            </p>

            {/* Circular toggle option */}
            <div
              onClick={() => setDeleteItemsWithPluginCategory(!deleteItemsWithPluginCategory)}
              className="p-3 bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-xl cursor-pointer flex items-center gap-3 select-none hover:border-zinc-300 dark:hover:border-zinc-600 transition-colors"
            >
              {/* Circular Icon */}
              <div
                className={`w-5 h-5 rounded-full border-2 flex items-center justify-center transition-all ${
                  deleteItemsWithPluginCategory
                    ? 'border-zinc-900 dark:border-zinc-100 bg-zinc-900 dark:bg-zinc-100'
                    : 'border-zinc-300 dark:border-zinc-600 bg-transparent'
                }`}
              >
                {deleteItemsWithPluginCategory && (
                  <div className="w-2 h-2 rounded-full bg-white dark:bg-zinc-900" />
                )}
              </div>
              <span className="text-xs text-zinc-800 dark:text-zinc-200 font-medium">
                一并把分组里的插件和脚本删除
              </span>
            </div>

            <p className="text-[10px] text-zinc-400">
              {deleteItemsWithPluginCategory
                ? '注意：该分组内的所有插件和脚本将被永久删除！'
                : '提示：未勾选时，该分组内的插件和脚本将保留并自动归类到 “默认” 分组。'}
            </p>

            <div className="flex justify-end gap-2 pt-2 border-t border-zinc-100 dark:border-zinc-800">
              <button
                onClick={() => setShowDeletePluginCategoryConfirm(false)}
                className="px-4 py-2 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800"
              >
                取消
              </button>
              <button
                onClick={handleConfirmDeletePluginCategory}
                className="px-4 py-2 text-xs font-bold rounded-lg bg-rose-600 text-white hover:bg-rose-700"
              >
                确定删除
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================== ST PLUGINS: BATCH MOVE MODAL ==================== */}
      {showPluginBatchMoveModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 animate-in fade-in pt-[calc(1rem+env(safe-area-inset-top,0px))]">
          <div className="w-full max-w-sm bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl p-5 space-y-4">
            <div className="flex items-center justify-between border-b border-zinc-100 dark:border-zinc-800 pb-3">
              <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100">批量移动分组</h3>
              <button
                onClick={() => setShowPluginBatchMoveModal(false)}
                className="p-1 rounded-lg text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-200"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div>
              <label className="text-xs font-semibold text-zinc-700 dark:text-zinc-300 block mb-1">
                选择目标分组
              </label>
              <select
                value={batchTargetPluginCategory}
                onChange={(e) => setBatchTargetPluginCategory(e.target.value)}
                className="w-full px-3 py-2 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-xl text-zinc-900 dark:text-zinc-100 focus:outline-none"
              >
                <option value="">-- 请选择分组 --</option>
                {Array.from(
                  new Set(['默认', ...(appData.pluginCategories || []), ...pluginsList.map((p) => p.category || '默认')])
                ).map((cat) => (
                  <option key={cat} value={cat}>
                    {cat}
                  </option>
                ))}
              </select>
            </div>

            <div className="flex justify-end gap-2 pt-2 border-t border-zinc-100 dark:border-zinc-800">
              <button
                onClick={() => setShowPluginBatchMoveModal(false)}
                className="px-4 py-2 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800"
              >
                取消
              </button>
              <button
                onClick={handleConfirmBatchMovePlugins}
                className="px-4 py-2 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:bg-zinc-800 dark:hover:bg-zinc-200"
              >
                确定移动
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Hidden File Inputs for Normal Character Cards */}
      <input
        type="file"
        ref={normalCardFileInputRef}
        multiple
        accept=".docx,.txt,.zip,text/plain,application/vnd.openxmlformats-officedocument.wordprocessingml.document,application/zip"
        className="hidden"
        onChange={(e) => {
          if (e.target.files) handleNormalCardFileUpload(e.target.files);
          e.target.value = '';
        }}
      />
      <input
        type="file"
        ref={normalCardCoverInputRef}
        accept="image/*"
        className="hidden"
        onChange={(e) => {
          if (e.target.files?.[0]) handleNormalCardCoverUpload(e.target.files[0]);
          e.target.value = '';
        }}
      />
      <input
        type="file"
        ref={normalCardUpdateFileInputRef}
        accept=".docx,.txt,text/plain,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        className="hidden"
        onChange={(e) => {
          if (e.target.files?.[0]) handleUpdateNormalCardTextContent(e.target.files[0]);
          e.target.value = '';
        }}
      />

      {normalZipPreview && (
        <div className="fixed inset-0 z-[100] flex items-center justify-center bg-black/60 p-4 pt-[calc(1rem+env(safe-area-inset-top,0px))] pb-[calc(1rem+env(safe-area-inset-bottom,0px))]">
          <div className="w-full max-w-xl max-h-[85vh] bg-white dark:bg-zinc-900 rounded-2xl shadow-2xl overflow-hidden flex flex-col">
            <div className="p-4 border-b border-zinc-200 dark:border-zinc-800">
              <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100">确认导入 ZIP 内的文档</h3>
              <p className="text-[11px] text-zinc-400 mt-1">ZIP：{normalZipPreview.fileName} · 识别到 {normalZipPreview.files.length} 个可导入文档</p>
              <div className="flex items-center justify-between mt-3">
                <span className="text-[11px] text-zinc-500 dark:text-zinc-400">已选择 {normalZipPreview.files.filter((f) => f.selected).length} / {normalZipPreview.files.length}</span>
                <button type="button" onClick={() => setNormalZipPreview((prev) => prev ? { ...prev, files: prev.files.map((f) => ({ ...f, selected: true })) } : prev)} className="text-[11px] font-semibold text-zinc-700 dark:text-zinc-300 hover:underline">全选</button>
              </div>
            </div>
            <div className="p-4 overflow-y-auto space-y-2 flex-1">
              {normalZipPreview.files.map((f) => {
                const displayName = f.name.split('/').pop() || f.name;
                const displayBaseName = displayName.replace(/\.[^/.]+$/, '') || displayName;
                return (
                  <label key={f.name} className="flex items-center gap-3 p-3 rounded-xl bg-zinc-50 dark:bg-zinc-800/50 border border-zinc-200 dark:border-zinc-800 cursor-pointer hover:bg-zinc-100 dark:hover:bg-zinc-800">
                    <input type="checkbox" checked={f.selected} onChange={() => toggleNormalZipFileSelection(f.name)} className="w-4 h-4 accent-zinc-900 dark:accent-zinc-100 flex-shrink-0" />
                    <div className="min-w-0 flex-1">
                      <div className="text-xs font-semibold truncate text-zinc-900 dark:text-zinc-100" title={displayName}>{displayBaseName}</div>
                      <div className="text-[10px] text-zinc-400 truncate" title={displayName}>{displayName} · {f.size.toLocaleString()} bytes</div>
                    </div>
                    {f.selected ? <CheckCircle2 className="w-4 h-4 text-emerald-500 flex-shrink-0" /> : <Circle className="w-4 h-4 text-zinc-300 dark:text-zinc-600 flex-shrink-0" />}
                  </label>
                );
              })}
            </div>
            <div className="p-4 border-t border-zinc-200 dark:border-zinc-800 flex justify-end gap-2 bg-zinc-50/70 dark:bg-zinc-900/70">
              <button type="button" onClick={() => setNormalZipPreview(null)} className="px-4 py-2 text-xs rounded-lg border border-zinc-200 dark:border-zinc-700">取消</button>
              <button type="button" disabled={!normalZipPreview.files.some((f) => f.selected)} onClick={confirmNormalZipImport} className="px-4 py-2 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 disabled:opacity-40 disabled:cursor-not-allowed">确定导入</button>
            </div>
          </div>
        </div>
      )}

      {/* ==================== NORMAL CARDS: SINGLE-PAGE DETAIL MODAL ==================== */}
      {editingNormalCard && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 animate-in fade-in pt-[calc(1rem+env(safe-area-inset-top,0px))]">
          <div className="w-full max-w-lg bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl overflow-hidden flex flex-col max-h-[85vh]">
            {/* Header */}
            <div className="flex items-center justify-between px-5 py-3.5 border-b border-zinc-100 dark:border-zinc-800 bg-zinc-50/50 dark:bg-zinc-800/30">
              <div className="flex items-center gap-2">
                <FileText className="w-4 h-4 text-zinc-500" />
                <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100 truncate max-w-[260px]">
                  普通角色卡详情: {editingNormalCard.fileName}
                </h3>
              </div>
              <button
                onClick={() => setEditingNormalCard(null)}
                className="p-1 rounded-lg text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-200"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            {/* Single-Page Scrollable Body */}
            <div className="p-5 overflow-y-auto space-y-4 flex-1">
              {/* Cover Image Reference */}
              <div className="flex items-center gap-4 bg-zinc-50 dark:bg-zinc-800/50 p-3 rounded-xl border border-zinc-200/60 dark:border-zinc-700/60">
                {editingNormalCard.coverImage ? (
                  <img
                    src={editingNormalCard.coverImage}
                    alt="角色图"
                    className="w-16 h-20 rounded-lg object-contain border border-zinc-200 dark:border-zinc-700 flex-shrink-0"
                  />
                ) : (
                  <div className="w-16 h-20 rounded-lg bg-zinc-200/60 dark:bg-zinc-700/60 flex items-center justify-center text-zinc-400 flex-shrink-0">
                    <ImageIcon className="w-6 h-6" />
                  </div>
                )}
                <div className="space-y-1.5 flex-1">
                  <div className="text-xs font-semibold text-zinc-800 dark:text-zinc-200">角色图参考</div>
                  <div className="flex items-center gap-2">
                    <button
                      onClick={() => normalCardCoverInputRef.current?.click()}
                      className="px-2.5 py-1 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800 flex items-center gap-1"
                    >
                      <Upload className="w-3 h-3" />
                      {editingNormalCard.coverImage ? '更换图片' : '上传图片'}
                    </button>
                    {editingNormalCard.coverImage && (
                      <button
                        onClick={() => setEditingNormalCard({ ...editingNormalCard, coverImage: null })}
                        className="px-2 py-1 text-xs text-rose-500 hover:text-rose-600 font-medium"
                      >
                        删除图片
                      </button>
                    )}
                  </div>
                </div>
              </div>

              {/* Grid Form Fields */}
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                {/* 文件名 */}
                <div>
                  <label className="text-[11px] font-semibold text-zinc-500 dark:text-zinc-400 block mb-1">
                    文件名 <span className="text-rose-500">*</span>
                  </label>
                  <input
                    type="text"
                    value={editingNormalCard.fileName}
                    onChange={(e) => setEditingNormalCard({ ...editingNormalCard, fileName: e.target.value })}
                    className="w-full px-3 py-1.5 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-xl text-zinc-900 dark:text-zinc-100 focus:outline-none"
                  />
                </div>

                {/* 角色真名 */}
                <div>
                  <label className="text-[11px] font-semibold text-zinc-500 dark:text-zinc-400 block mb-1">
                    角色真名
                  </label>
                  <input
                    type="text"
                    value={editingNormalCard.charName || ''}
                    onChange={(e) => setEditingNormalCard({ ...editingNormalCard, charName: e.target.value })}
                    placeholder="角色实际名字"
                    className="w-full px-3 py-1.5 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-xl text-zinc-900 dark:text-zinc-100 focus:outline-none"
                  />
                </div>

                {/* 作者 */}
                <div>
                  <label className="text-[11px] font-semibold text-zinc-500 dark:text-zinc-400 block mb-1">
                    作者
                  </label>
                  <input
                    type="text"
                    value={editingNormalCard.author || ''}
                    onChange={(e) => setEditingNormalCard({ ...editingNormalCard, author: e.target.value })}
                    placeholder="作者/创作者"
                    className="w-full px-3 py-1.5 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-xl text-zinc-900 dark:text-zinc-100 focus:outline-none"
                  />
                </div>

                {/* 分组 */}
                <div>
                  <label className="text-[11px] font-semibold text-zinc-500 dark:text-zinc-400 block mb-1">
                    分组
                  </label>
                  <select
                    value={editingNormalCard.category || '默认'}
                    onChange={(e) => setEditingNormalCard({ ...editingNormalCard, category: e.target.value })}
                    className="w-full px-3 py-1.5 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-xl text-zinc-900 dark:text-zinc-100 focus:outline-none"
                  >
                    {Array.from(
                      new Set(['默认', ...(appData.normalCardCategories || []), ...normalCardsList.map((p) => p.category || '默认')])
                    ).map((cat) => (
                      <option key={cat} value={cat}>
                        {cat}
                      </option>
                    ))}
                  </select>
                </div>
              </div>

              {/* 来源 */}
              <div>
                <label className="text-[11px] font-semibold text-zinc-500 dark:text-zinc-400 block mb-1">
                  联系方式 / 来源 (DC链接/QQ群等)
                </label>
                <input
                  type="text"
                  value={editingNormalCard.source || ''}
                  onChange={(e) => setEditingNormalCard({ ...editingNormalCard, source: e.target.value })}
                  placeholder="例如: https://discord.gg/... 或 QQ群: 123456"
                  className="w-full px-3 py-1.5 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-xl text-zinc-900 dark:text-zinc-100 focus:outline-none"
                />
              </div>

              {/* 文档内容 Preview & Token Estimation (Requirement 7) */}
              <div>
                <div className="flex items-center justify-between mb-1">
                  <label className="text-[11px] font-semibold text-zinc-500 dark:text-zinc-400">
                    文档内容 ( Token估算: <span className="font-mono text-amber-600 dark:text-amber-400 font-bold">{estimateTokens(editingNormalCard.content || '')}</span> )
                  </label>
                  <div className="flex items-center gap-2">
                    <button
                      onClick={() => {
                        navigator.clipboard.writeText(editingNormalCard.content || '');
                        showToast('文档内容已复制到剪贴板！', 'success');
                      }}
                      className="text-xs text-blue-600 dark:text-blue-400 hover:underline flex items-center gap-1 font-medium"
                    >
                      <Copy className="w-3 h-3" /> 复制内容
                    </button>
                    <button
                      onClick={() => setShowExpandedContentModal(true)}
                      className="text-xs text-emerald-600 dark:text-emerald-400 hover:underline flex items-center gap-1 font-medium"
                    >
                      <Maximize2 className="w-3 h-3" /> 放大展开
                    </button>
                  </div>
                </div>
                <div className="p-3 bg-zinc-50 dark:bg-zinc-800/50 border border-zinc-200 dark:border-zinc-700/60 rounded-xl text-xs text-zinc-700 dark:text-zinc-300 font-mono line-clamp-4 leading-relaxed whitespace-pre-wrap select-text">
                  {editingNormalCard.content || '（暂无文档内容）'}
                </div>
              </div>

              {/* Export Buttons: .txt, .docx, .json (Requirement 3) */}
              <div className="space-y-1.5 pt-1">
                <span className="text-[11px] font-semibold text-zinc-500 block">导出文档格式:</span>
                <div className="grid grid-cols-3 gap-2">
                  <button
                    onClick={() => {
                      const blob = new Blob([editingNormalCard.content || ''], { type: 'text/plain;charset=utf-8' });
                      const url = URL.createObjectURL(blob);
                      const a = document.createElement('a');
                      a.href = url;
                      a.download = `${editingNormalCard.fileName}.txt`;
                      a.click();
                      URL.revokeObjectURL(url);
                      showToast('已导出为 TXT 文本文件', 'success');
                    }}
                    className="px-2.5 py-1.5 text-xs font-medium rounded-xl border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800 flex items-center justify-center gap-1"
                  >
                    <Download className="w-3.5 h-3.5" /> .txt
                  </button>

                  <button
                    onClick={() => {
                      exportThemeDocument({ name: editingNormalCard.fileName, content: editingNormalCard.content }, 'docx');
                    }}
                    className="px-2.5 py-1.5 text-xs font-medium rounded-xl border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800 flex items-center justify-center gap-1"
                  >
                    <Download className="w-3.5 h-3.5" /> .docx
                  </button>

                  <button
                    onClick={() => {
                      const dataStr = JSON.stringify(editingNormalCard, null, 2);
                      const blob = new Blob([dataStr], { type: 'application/json' });
                      const url = URL.createObjectURL(blob);
                      const a = document.createElement('a');
                      a.href = url;
                      a.download = `${editingNormalCard.fileName}.json`;
                      a.click();
                      URL.revokeObjectURL(url);
                      showToast('已导出为 JSON 文件', 'success');
                    }}
                    className="px-2.5 py-1.5 text-xs font-medium rounded-xl border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800 flex items-center justify-center gap-1"
                  >
                    <Download className="w-3.5 h-3.5" /> .json
                  </button>
                </div>
              </div>
            </div>

            {/* Bottom Action Bar: 取消, 确定, 删除, 更新 (Requirement 5 Strengthened) */}
            <div className="px-5 py-3 border-t border-zinc-100 dark:border-zinc-800 bg-zinc-50/50 dark:bg-zinc-800/30 flex items-center justify-between gap-2">
              <button
                onClick={() => {
                  if (confirm(`确定要永久删除普通角色卡 “${editingNormalCard.fileName}” 吗？`)) {
                    const currentCards = appData.normalCards || [];
                    const updated = currentCards.filter((p) => p.id !== editingNormalCard.id);
                    updateAppData({ ...appData, normalCards: updated });
                    setEditingNormalCard(null);
                    showToast('角色卡已彻底删除', 'info');
                  }
                }}
                className="px-3 py-1.5 text-xs font-bold rounded-xl bg-rose-600 text-white hover:bg-rose-700 flex items-center gap-1 shadow-sm"
              >
                <Trash2 className="w-3.5 h-3.5" /> 彻底删除
              </button>

              <div className="flex items-center gap-2">
                <button
                  onClick={() => setEditingNormalCard(null)}
                  className="px-3 py-1.5 text-xs font-medium rounded-xl border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800"
                >
                  取消
                </button>
                <button
                  onClick={() => normalCardUpdateFileInputRef.current?.click()}
                  className="px-3 py-1.5 text-xs font-medium rounded-xl border border-zinc-300 dark:border-zinc-600 text-zinc-800 dark:text-zinc-200 hover:bg-zinc-100 dark:hover:bg-zinc-800 flex items-center gap-1"
                  title="重新上传文档覆写替换当前内容"
                >
                  <RefreshCw className="w-3.5 h-3.5" /> 更新
                </button>
                <button
                  onClick={() => {
                    const updated = normalCardsList.map((p) => (p.id === editingNormalCard.id ? editingNormalCard : p));
                    updateAppData({ ...appData, normalCards: updated });
                    setEditingNormalCard(null);
                    showToast('保存成功！', 'success');
                  }}
                  className="px-4 py-1.5 text-xs font-bold rounded-xl bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:bg-zinc-800 dark:hover:bg-zinc-200"
                >
                  确定
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ==================== NORMAL CARDS: EXPANDED CONTENT MODAL ==================== */}
      {showExpandedContentModal && editingNormalCard && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 animate-in fade-in pt-[calc(1rem+env(safe-area-inset-top,0px))]">
          <div className="w-full max-w-3xl bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl overflow-hidden flex flex-col h-[80vh]">
            <div className="flex items-center justify-between px-5 py-3.5 border-b border-zinc-100 dark:border-zinc-800 bg-zinc-50/50 dark:bg-zinc-800/30">
              <div className="flex items-center gap-2">
                <Maximize2 className="w-4 h-4 text-emerald-500" />
                <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100 truncate">
                  文档展开内容 - {editingNormalCard.fileName}
                </h3>
              </div>
              <button
                onClick={() => setShowExpandedContentModal(false)}
                className="p-1 rounded-lg text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-200"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="p-5 flex-1 overflow-y-auto">
              <textarea
                value={editingNormalCard.content || ''}
                onChange={(e) => setEditingNormalCard({ ...editingNormalCard, content: e.target.value })}
                className="w-full h-full p-4 text-xs font-mono bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-xl text-zinc-900 dark:text-zinc-100 focus:outline-none resize-none leading-relaxed"
                placeholder="此处显示文档全部文本..."
              />
            </div>

            <div className="px-5 py-3 border-t border-zinc-100 dark:border-zinc-800 flex items-center justify-between">
              <button
                onClick={() => {
                  navigator.clipboard.writeText(editingNormalCard.content || '');
                  showToast('文档全景内容已复制到剪贴板！', 'success');
                }}
                className="px-3 py-1.5 text-xs font-medium rounded-xl bg-blue-50 text-blue-600 dark:bg-blue-950/60 dark:text-blue-400 hover:bg-blue-100 flex items-center gap-1"
              >
                <Copy className="w-3.5 h-3.5" /> 复制全文
              </button>
              <button
                onClick={() => setShowExpandedContentModal(false)}
                className="px-4 py-1.5 text-xs font-bold rounded-xl bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:bg-zinc-800"
              >
                关闭
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================== NORMAL CARDS: NEW GROUP MODAL ==================== */}
      {showNewNormalCardGroupModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 animate-in fade-in pt-[calc(1rem+env(safe-area-inset-top,0px))]">
          <div className="w-full max-w-sm bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl p-5 space-y-4">
            <div className="flex items-center justify-between border-b border-zinc-100 dark:border-zinc-800 pb-3">
              <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100">新建角色卡分组</h3>
              <button
                onClick={() => setShowNewNormalCardGroupModal(false)}
                className="p-1 rounded-lg text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-200"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div>
              <label className="text-xs font-semibold text-zinc-700 dark:text-zinc-300 block mb-1">
                分组名称
              </label>
              <input
                type="text"
                value={newNormalCardGroupName}
                onChange={(e) => setNewNormalCardGroupName(e.target.value)}
                placeholder="请输入分组名称"
                className="w-full px-3 py-2 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-xl text-zinc-900 dark:text-zinc-100 focus:outline-none"
                onKeyDown={(e) => e.key === 'Enter' && handleCreateNewNormalCardCategory()}
              />
            </div>

            <div className="flex justify-end gap-2 pt-2 border-t border-zinc-100 dark:border-zinc-800">
              <button
                onClick={() => setShowNewNormalCardGroupModal(false)}
                className="px-4 py-2 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800"
              >
                取消
              </button>
              <button
                onClick={handleCreateNewNormalCardCategory}
                className="px-4 py-2 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:bg-zinc-800 dark:hover:bg-zinc-200"
              >
                创建
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================== NORMAL CARDS: MANAGE GROUP MODAL ==================== */}
      {managingNormalCardCategory && !showDeleteNormalCardCategoryConfirm && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 animate-in fade-in pt-[calc(1rem+env(safe-area-inset-top,0px))]">
          <div className="w-full max-w-sm bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl p-5 space-y-4">
            <div className="flex items-center justify-between border-b border-zinc-100 dark:border-zinc-800 pb-3">
              <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100">
                管理分组: {managingNormalCardCategory}
              </h3>
              <button
                onClick={() => setManagingNormalCardCategory(null)}
                className="p-1 rounded-lg text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-200"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div>
              <label className="text-xs font-semibold text-zinc-700 dark:text-zinc-300 block mb-1">
                重命名分组
              </label>
              <div className="flex gap-2">
                <input
                  type="text"
                  value={renameNormalCardCategoryInput}
                  onChange={(e) => setRenameNormalCardCategoryInput(e.target.value)}
                  className="flex-1 px-3 py-1.5 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-xl text-zinc-900 dark:text-zinc-100 focus:outline-none"
                />
                <button
                  onClick={handleRenameNormalCardCategory}
                  className="px-3 py-1.5 text-xs font-semibold rounded-xl bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:bg-zinc-800"
                >
                  确定
                </button>
              </div>
            </div>

            <div className="pt-2 border-t border-zinc-100 dark:border-zinc-800 flex justify-between items-center">
              <button
                onClick={() => setShowDeleteNormalCardCategoryConfirm(true)}
                className="px-3 py-1.5 text-xs font-bold text-rose-600 hover:bg-rose-50 dark:hover:bg-rose-950/60 rounded-xl transition-colors flex items-center gap-1"
              >
                <Trash2 className="w-3.5 h-3.5" /> 删除该分组
              </button>
              <button
                onClick={() => setManagingNormalCardCategory(null)}
                className="px-3 py-1.5 text-xs font-medium rounded-xl border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100"
              >
                关闭
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================== NORMAL CARDS: DELETE GROUP CONFIRM MODAL ==================== */}
      {showDeleteNormalCardCategoryConfirm && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 animate-in fade-in pt-[calc(1rem+env(safe-area-inset-top,0px))]">
          <div className="w-full max-w-sm bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl p-5 space-y-4">
            <div className="flex items-center justify-between border-b border-zinc-100 dark:border-zinc-800 pb-3">
              <h3 className="text-sm font-bold text-rose-600 dark:text-rose-400">确认删除分组</h3>
              <button
                onClick={() => setShowDeleteNormalCardCategoryConfirm(false)}
                className="p-1 rounded-lg text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-200"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <p className="text-xs text-zinc-600 dark:text-zinc-300">
              确定要删除分组 “<span className="font-bold underline">{managingNormalCardCategory}</span>” 吗？
            </p>

            {/* Circular toggle option */}
            <div
              onClick={() => setDeleteItemsWithNormalCardCategory(!deleteItemsWithNormalCardCategory)}
              className="p-3 bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-xl cursor-pointer flex items-center gap-3 select-none hover:border-zinc-300 dark:hover:border-zinc-600 transition-colors"
            >
              {/* Circular Icon */}
              <div
                className={`w-5 h-5 rounded-full border-2 flex items-center justify-center transition-all ${
                  deleteItemsWithNormalCardCategory
                    ? 'border-zinc-900 dark:border-zinc-100 bg-zinc-900 dark:bg-zinc-100'
                    : 'border-zinc-300 dark:border-zinc-600 bg-transparent'
                }`}
              >
                {deleteItemsWithNormalCardCategory && (
                  <div className="w-2 h-2 rounded-full bg-white dark:bg-zinc-900" />
                )}
              </div>
              <span className="text-xs text-zinc-800 dark:text-zinc-200 font-medium">
                一并把分组里的角色卡给删掉
              </span>
            </div>

            <p className="text-[10px] text-zinc-400">
              {deleteItemsWithNormalCardCategory
                ? '注意：该分组内的所有角色卡将被永久删除！'
                : '提示：未勾选时，该分组内的角色卡将保留并自动归类到 “默认” 分组。'}
            </p>

            <div className="flex justify-end gap-2 pt-2 border-t border-zinc-100 dark:border-zinc-800">
              <button
                onClick={() => setShowDeleteNormalCardCategoryConfirm(false)}
                className="px-4 py-2 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800"
              >
                取消
              </button>
              <button
                onClick={handleConfirmDeleteNormalCardCategory}
                className="px-4 py-2 text-xs font-bold rounded-lg bg-rose-600 text-white hover:bg-rose-700"
              >
                确定删除
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================== NORMAL CARDS: BATCH MOVE MODAL ==================== */}
      {showNormalCardBatchMoveModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 animate-in fade-in pt-[calc(1rem+env(safe-area-inset-top,0px))]">
          <div className="w-full max-w-sm bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl p-5 space-y-4">
            <div className="flex items-center justify-between border-b border-zinc-100 dark:border-zinc-800 pb-3">
              <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100">批量移动分组</h3>
              <button
                onClick={() => setShowNormalCardBatchMoveModal(false)}
                className="p-1 rounded-lg text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-200"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div>
              <label className="text-xs font-semibold text-zinc-700 dark:text-zinc-300 block mb-1">
                选择目标分组
              </label>
              <select
                value={batchTargetNormalCardCategory}
                onChange={(e) => setBatchTargetNormalCardCategory(e.target.value)}
                className="w-full px-3 py-2 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-xl text-zinc-900 dark:text-zinc-100 focus:outline-none"
              >
                <option value="">-- 请选择分组 --</option>
                {Array.from(
                  new Set(['默认', ...(appData.normalCardCategories || []), ...normalCardsList.map((p) => p.category || '默认')])
                ).map((cat) => (
                  <option key={cat} value={cat}>
                    {cat}
                  </option>
                ))}
              </select>
            </div>

            <div className="flex justify-end gap-2 pt-2 border-t border-zinc-100 dark:border-zinc-800">
              <button
                onClick={() => setShowNormalCardBatchMoveModal(false)}
                className="px-4 py-2 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800"
              >
                取消
              </button>
              <button
                onClick={handleConfirmBatchMoveNormalCards}
                className="px-4 py-2 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:bg-zinc-800 dark:hover:bg-zinc-200"
              >
                确定移动
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================== ADD API MODAL ==================== */}
      {showAddApiModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 animate-in fade-in pt-[calc(1rem+env(safe-area-inset-top,0px))]">
          <div className="w-full max-w-lg bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl p-6 space-y-4 max-h-[90vh] overflow-y-auto">
            <div className="flex items-center justify-between border-b border-zinc-100 dark:border-zinc-800 pb-3">
              <h3 className="text-base font-bold text-zinc-900 dark:text-zinc-100">添加 API</h3>
              <button
                onClick={() => setShowAddApiModal(false)}
                className="p-1 rounded-lg text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-200 hover:bg-zinc-100 dark:hover:bg-zinc-800"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="space-y-3.5">
              {/* API 名称 */}
              <div>
                <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                  API 名称 <span className="text-rose-500">*</span>
                </label>
                <input
                  type="text"
                  value={newApiForm.name}
                  onChange={(e) => setNewApiForm((p) => ({ ...p, name: e.target.value }))}
                  placeholder="请输入 API 名称..."
                  className="w-full px-3 py-2 text-sm bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400 text-zinc-900 dark:text-zinc-100"
                />
              </div>

              {/* API URL */}
              <div>
                <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                  API 地址 (URL) <span className="text-rose-500">*</span>
                </label>
                <input
                  type="text"
                  value={newApiForm.url}
                  onChange={(e) => setNewApiForm((p) => ({ ...p, url: e.target.value }))}
                  placeholder="请输入 API 地址, 例如 https://api.openai.com/v1..."
                  className="w-full px-3 py-2 text-sm bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400 font-mono text-zinc-900 dark:text-zinc-100"
                />
              </div>

              {/* Key 列表 with memo column and '+' button on right side */}
              <div>
                <div className="flex items-center justify-between mb-1.5">
                  <label className="text-xs font-semibold text-zinc-700 dark:text-zinc-300">
                    API Key 列表
                  </label>
                  <button
                    type="button"
                    onClick={() =>
                      setNewApiForm((p) => ({
                        ...p,
                        keys: [
                          ...p.keys,
                          { id: Date.now().toString(36) + Math.random().toString(36).substring(2, 5), memo: '', key: '' },
                        ],
                      }))
                    }
                    className="flex items-center gap-1 px-2 py-1 text-xs font-semibold rounded bg-zinc-100 dark:bg-zinc-800 hover:bg-zinc-200 dark:hover:bg-zinc-700 text-zinc-800 dark:text-zinc-200 transition-colors"
                  >
                    <Plus className="w-3.5 h-3.5" /> 添加 Key
                  </button>
                </div>

                <div className="space-y-2">
                  {newApiForm.keys.map((kItem, idx) => (
                    <div key={kItem.id || idx} className="flex items-center gap-1.5">
                      <input
                        type="text"
                        value={kItem.memo || ''}
                        onChange={(e) => {
                          const val = e.target.value;
                          setNewApiForm((p) => ({
                            ...p,
                            keys: p.keys.map((k, i) => (i === idx ? { ...k, memo: val } : k)),
                          }));
                        }}
                        placeholder="备注"
                        className="w-12 sm:w-16 flex-shrink-0 px-1.5 py-1.5 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none text-zinc-900 dark:text-zinc-100"
                      />
                      <input
                        type="text"
                        value={kItem.key}
                        onChange={(e) => {
                          const val = e.target.value;
                          setNewApiForm((p) => ({
                            ...p,
                            keys: p.keys.map((k, i) => (i === idx ? { ...k, key: val } : k)),
                          }));
                        }}
                        placeholder="sk-..."
                        className="flex-1 min-w-0 px-2.5 py-1.5 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none font-mono text-zinc-900 dark:text-zinc-100"
                      />
                      {newApiForm.keys.length > 1 && (
                        <button
                          type="button"
                          onClick={() =>
                            setNewApiForm((p) => ({
                              ...p,
                              keys: p.keys.filter((_, i) => i !== idx),
                            }))
                          }
                          className="p-1.5 text-zinc-400 hover:text-rose-500 rounded-lg flex-shrink-0"
                          title="删除此Key"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                        </button>
                      )}
                    </div>
                  ))}
                </div>
              </div>

              {/* 描述 */}
              <div>
                <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                  描述
                </label>
                <textarea
                  rows={2}
                  value={newApiForm.description}
                  onChange={(e) => setNewApiForm((p) => ({ ...p, description: e.target.value }))}
                  placeholder="请输入描述信息（选填）..."
                  className="w-full px-3 py-2 text-sm bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400 text-zinc-900 dark:text-zinc-100 resize-none"
                />
              </div>

              {/* 分组 */}
              <div>
                <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                  所属分组
                </label>
                <select
                  value={newApiForm.category}
                  onChange={(e) => setNewApiForm((p) => ({ ...p, category: e.target.value }))}
                  className="w-full px-3 py-2 text-sm bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none text-zinc-900 dark:text-zinc-100"
                >
                  {Array.from(
                    new Set(['默认', ...(appData.apiCategories || []), ...apisList.map((p) => p.category || '默认')])
                  ).map((cat) => (
                    <option key={cat} value={cat}>
                      {cat}
                    </option>
                  ))}
                </select>
              </div>
            </div>

            {/* Bottom Action Buttons: 取消, 确定 */}
            <div className="flex items-center justify-end gap-2.5 pt-2 border-t border-zinc-100 dark:border-zinc-800">
              <button
                onClick={() => setShowAddApiModal(false)}
                className="px-4 py-2 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800"
              >
                取消
              </button>
              <button
                onClick={handleAddApi}
                className="px-4 py-2 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:bg-zinc-800 dark:hover:bg-zinc-200"
              >
                确定
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================== EDIT/DETAIL SINGLE-PAGE API MODAL ==================== */}
      {editingApi && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 animate-in fade-in pt-[calc(1rem+env(safe-area-inset-top,0px))]">
          <div className="w-full max-w-lg bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl p-6 space-y-4 max-h-[90vh] overflow-y-auto">
            <div className="flex items-center justify-between border-b border-zinc-100 dark:border-zinc-800 pb-3">
              <h3 className="text-base font-bold text-zinc-900 dark:text-zinc-100">API 详情</h3>
              <button
                onClick={() => setEditingApi(null)}
                className="p-1 rounded-lg text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-200 hover:bg-zinc-100 dark:hover:bg-zinc-800"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="space-y-3.5">
              {/* API 名称 */}
              <div>
                <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                  API 名称 <span className="text-rose-500">*</span>
                </label>
                <input
                  type="text"
                  value={editingApi.name}
                  onChange={(e) => setEditingApi((p) => (p ? { ...p, name: e.target.value } : null))}
                  className="w-full px-3 py-2 text-sm bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none text-zinc-900 dark:text-zinc-100"
                />
              </div>

              {/* API 地址 + Copy Button on right */}
              <div>
                <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                  API 地址 (URL) <span className="text-rose-500">*</span>
                </label>
                <div className="flex items-center gap-2">
                  <input
                    type="text"
                    value={editingApi.url}
                    onChange={(e) => setEditingApi((p) => (p ? { ...p, url: e.target.value } : null))}
                    className="flex-1 px-3 py-2 text-sm bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none font-mono text-zinc-900 dark:text-zinc-100"
                  />
                  <button
                    type="button"
                    onClick={() => {
                      if (editingApi.url) {
                        navigator.clipboard.writeText(editingApi.url);
                        showToast('API 地址已复制到剪贴板！', 'success');
                      }
                    }}
                    className="px-3 py-2 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 hover:bg-zinc-100 dark:hover:bg-zinc-800 text-zinc-700 dark:text-zinc-300 flex items-center gap-1 flex-shrink-0"
                    title="复制 API 地址"
                  >
                    <Copy className="w-3.5 h-3.5" /> 复制
                  </button>
                </div>
              </div>

              {/* Key 列表 + Copy Key button on right of each key */}
              <div>
                <div className="flex items-center justify-between mb-1.5">
                  <label className="text-xs font-semibold text-zinc-700 dark:text-zinc-300">
                    API Key 列表
                  </label>
                  <button
                    type="button"
                    onClick={() =>
                      setEditingApi((p) =>
                        p
                          ? {
                              ...p,
                              keys: [
                                ...p.keys,
                                { id: Date.now().toString(36) + Math.random().toString(36).substring(2, 5), memo: '', key: '' },
                              ],
                            }
                          : null
                      )
                    }
                    className="flex items-center gap-1 px-2 py-1 text-xs font-semibold rounded bg-zinc-100 dark:bg-zinc-800 hover:bg-zinc-200 dark:hover:bg-zinc-700 text-zinc-800 dark:text-zinc-200 transition-colors"
                  >
                    <Plus className="w-3.5 h-3.5" /> 添加 Key
                  </button>
                </div>

                <div className="space-y-2">
                  {editingApi.keys.map((kItem, idx) => (
                    <div key={kItem.id || idx} className="flex items-center gap-1.5">
                      <input
                        type="text"
                        value={kItem.memo || ''}
                        onChange={(e) => {
                          const val = e.target.value;
                          setEditingApi((p) =>
                            p
                              ? {
                                  ...p,
                                  keys: p.keys.map((k, i) => (i === idx ? { ...k, memo: val } : k)),
                                }
                              : null
                          );
                        }}
                        placeholder="备注"
                        className="w-12 sm:w-16 flex-shrink-0 px-1.5 py-1.5 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none text-zinc-900 dark:text-zinc-100"
                      />
                      <input
                        type="text"
                        value={kItem.key}
                        onChange={(e) => {
                          const val = e.target.value;
                          setEditingApi((p) =>
                            p
                              ? {
                                  ...p,
                                  keys: p.keys.map((k, i) => (i === idx ? { ...k, key: val } : k)),
                                }
                              : null
                          );
                        }}
                        placeholder="Key 值"
                        className="flex-1 min-w-0 px-2.5 py-1.5 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none font-mono text-zinc-900 dark:text-zinc-100"
                      />
                      <button
                        type="button"
                        onClick={() => {
                          if (kItem.key) {
                            navigator.clipboard.writeText(kItem.key);
                            showToast('Key 已复制（不包含备注）！', 'success');
                          } else {
                            showToast('Key 为空', 'error');
                          }
                        }}
                        className="p-1.5 text-zinc-500 hover:text-zinc-900 dark:hover:text-zinc-100 border border-zinc-200 dark:border-zinc-700 rounded-lg hover:bg-zinc-100 dark:hover:bg-zinc-800 flex-shrink-0"
                        title="单独复制 Key"
                      >
                        <Copy className="w-3.5 h-3.5" />
                      </button>

                      {editingApi.keys.length > 1 && (
                        <button
                          type="button"
                          onClick={() =>
                            setEditingApi((p) =>
                              p
                                ? {
                                    ...p,
                                    keys: p.keys.filter((_, i) => i !== idx),
                                  }
                                : null
                            )
                          }
                          className="p-1.5 text-zinc-400 hover:text-rose-500 rounded-lg flex-shrink-0"
                          title="删除此Key"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                        </button>
                      )}
                    </div>
                  ))}
                </div>
              </div>

              {/* 描述 */}
              <div>
                <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                  描述
                </label>
                <textarea
                  rows={2}
                  value={editingApi.description || ''}
                  onChange={(e) => setEditingApi((p) => (p ? { ...p, description: e.target.value } : null))}
                  placeholder="请输入描述信息..."
                  className="w-full px-3 py-2 text-sm bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none text-zinc-900 dark:text-zinc-100 resize-none"
                />
              </div>

              {/* 分组 */}
              <div>
                <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                  所属分组
                </label>
                <select
                  value={editingApi.category || '默认'}
                  onChange={(e) => setEditingApi((p) => (p ? { ...p, category: e.target.value } : null))}
                  className="w-full px-3 py-2 text-sm bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none text-zinc-900 dark:text-zinc-100"
                >
                  {Array.from(
                    new Set(['默认', ...(appData.apiCategories || []), ...apisList.map((p) => p.category || '默认')])
                  ).map((cat) => (
                    <option key={cat} value={cat}>
                      {cat}
                    </option>
                  ))}
                </select>
              </div>
            </div>

            {/* Bottom Buttons: 取消, 确定, 删除 (Requirement 5) */}
            <div className="flex items-center justify-between pt-3 border-t border-zinc-100 dark:border-zinc-800">
              <button
                type="button"
                onClick={() => {
                  if (confirm(`确定要删除 API “${editingApi.name}” 吗？`)) {
                    const updated = apisList.filter((p) => p.id !== editingApi.id);
                    updateAppData({ ...appData, apis: updated });
                    setEditingApi(null);
                    showToast('API 已删除', 'info');
                  }
                }}
                className="px-3.5 py-2 text-xs font-bold rounded-lg bg-rose-50 text-rose-600 dark:bg-rose-950/60 dark:text-rose-400 hover:bg-rose-100 dark:hover:bg-rose-900/60 flex items-center gap-1"
              >
                <Trash2 className="w-3.5 h-3.5" /> 删除
              </button>

              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={() => setEditingApi(null)}
                  className="px-4 py-2 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800"
                >
                  取消
                </button>
                <button
                  type="button"
                  onClick={() => {
                    if (!editingApi.name.trim() || !editingApi.url.trim()) {
                      showToast('API 名称和地址不能为空', 'error');
                      return;
                    }
                    const updated = apisList.map((p) => (p.id === editingApi.id ? { ...editingApi, updatedAt: Date.now() } : p));
                    updateAppData({ ...appData, apis: updated });
                    setEditingApi(null);
                    showToast('保存成功！', 'success');
                  }}
                  className="px-4 py-2 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:bg-zinc-800 dark:hover:bg-zinc-200"
                >
                  确定
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ==================== NEW API GROUP MODAL ==================== */}
      {showNewApiGroupModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 animate-in fade-in pt-[calc(1rem+env(safe-area-inset-top,0px))]">
          <div className="w-full max-w-sm bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl p-5 space-y-4">
            <div className="flex items-center justify-between border-b border-zinc-100 dark:border-zinc-800 pb-3">
              <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100">新建 API 分组</h3>
              <button
                onClick={() => setShowNewApiGroupModal(false)}
                className="p-1 rounded-lg text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-200"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div>
              <label className="text-xs font-semibold text-zinc-700 dark:text-zinc-300 block mb-1">
                分组名称
              </label>
              <input
                type="text"
                value={newApiGroupName}
                onChange={(e) => setNewApiGroupName(e.target.value)}
                placeholder="请输入分组名字..."
                className="w-full px-3 py-2 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-xl text-zinc-900 dark:text-zinc-100 focus:outline-none"
              />
            </div>

            <div className="flex justify-end gap-2 pt-2 border-t border-zinc-100 dark:border-zinc-800">
              <button
                onClick={() => setShowNewApiGroupModal(false)}
                className="px-4 py-2 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800"
              >
                取消
              </button>
              <button
                onClick={handleCreateNewApiCategory}
                className="px-4 py-2 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:bg-zinc-800 dark:hover:bg-zinc-200"
              >
                确定新建
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================== MANAGE API GROUP MODAL ==================== */}
      {managingApiCategory && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 animate-in fade-in pt-[calc(1rem+env(safe-area-inset-top,0px))]">
          <div className="w-full max-w-sm bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl p-5 space-y-4">
            <div className="flex items-center justify-between border-b border-zinc-100 dark:border-zinc-800 pb-3">
              <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100">
                管理分组 - “{managingApiCategory}”
              </h3>
              <button
                onClick={() => setManagingApiCategory(null)}
                className="p-1 rounded-lg text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-200"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            {/* Rename Section */}
            <div className="space-y-2">
              <label className="text-xs font-semibold text-zinc-500 block">重命名分组</label>
              <div className="flex gap-2">
                <input
                  type="text"
                  value={renameApiCategoryInput}
                  onChange={(e) => setRenameApiCategoryInput(e.target.value)}
                  className="flex-1 px-3 py-1.5 text-xs bg-zinc-50 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 rounded-lg text-zinc-900 dark:text-zinc-100 focus:outline-none"
                />
                <button
                  onClick={handleRenameApiCategory}
                  className="px-3 py-1.5 text-xs font-medium bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 rounded-lg hover:opacity-90"
                >
                  保存名字
                </button>
              </div>
            </div>

            {/* Delete Group Dialog with Circular Toggle (Requirement 1 & 4) */}
            <div className="pt-3 border-t border-zinc-200 dark:border-zinc-800 space-y-3">
              <div className="text-xs font-bold text-rose-600 dark:text-rose-400">删除分组</div>
              <p className="text-xs text-zinc-500 dark:text-zinc-400 leading-relaxed">
                是否要删除分组“{managingApiCategory}”？删除前请选择是否一并清理组内 API。
              </p>

              {/* Circular Selection Toggle */}
              <div
                onClick={() => setDeleteItemsWithApiCategory((prev) => !prev)}
                className="flex items-center gap-3 p-3 bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-xl cursor-pointer hover:bg-zinc-100 dark:hover:bg-zinc-800 transition-colors"
              >
                <div
                  className={`w-5 h-5 rounded-full border-2 flex items-center justify-center transition-all ${
                    deleteItemsWithApiCategory
                      ? 'border-zinc-900 dark:border-zinc-100 bg-white dark:bg-zinc-900'
                      : 'border-zinc-400 dark:border-zinc-600 bg-transparent'
                  }`}
                >
                  {deleteItemsWithApiCategory && (
                    <div className="w-2.5 h-2.5 rounded-full bg-zinc-900 dark:bg-zinc-100" />
                  )}
                </div>

                <div className="text-xs font-medium text-zinc-800 dark:text-zinc-200">
                  是否要一并把分组里的 API 给删掉？
                </div>
              </div>

              <div className="flex justify-end gap-2 pt-2">
                <button
                  onClick={() => setManagingApiCategory(null)}
                  className="px-3 py-1.5 text-xs font-medium text-zinc-600 dark:text-zinc-400 hover:bg-zinc-100 dark:hover:bg-zinc-800 rounded-lg"
                >
                  取消
                </button>
                <button
                  onClick={handleConfirmDeleteApiCategory}
                  className="px-4 py-1.5 text-xs font-bold bg-rose-600 text-white hover:bg-rose-700 rounded-lg"
                >
                  确认删除分组
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ==================== BATCH MOVE API MODAL ==================== */}
      {showApiBatchMoveModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 animate-in fade-in pt-[calc(1rem+env(safe-area-inset-top,0px))]">
          <div className="w-full max-w-sm bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl p-5 space-y-4">
            <div className="flex items-center justify-between border-b border-zinc-100 dark:border-zinc-800 pb-3">
              <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100">批量移动分组</h3>
              <button
                onClick={() => setShowApiBatchMoveModal(false)}
                className="p-1 rounded-lg text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-200"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div>
              <label className="text-xs font-semibold text-zinc-700 dark:text-zinc-300 block mb-1">
                选择目标分组
              </label>
              <select
                value={batchTargetApiCategory}
                onChange={(e) => setBatchTargetApiCategory(e.target.value)}
                className="w-full px-3 py-2 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-xl text-zinc-900 dark:text-zinc-100 focus:outline-none"
              >
                <option value="">-- 请选择分组 --</option>
                {Array.from(
                  new Set(['默认', ...(appData.apiCategories || []), ...apisList.map((p) => p.category || '默认')])
                ).map((cat) => (
                  <option key={cat} value={cat}>
                    {cat}
                  </option>
                ))}
              </select>
            </div>

            <div className="flex justify-end gap-2 pt-2 border-t border-zinc-100 dark:border-zinc-800">
              <button
                onClick={() => setShowApiBatchMoveModal(false)}
                className="px-4 py-2 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800"
              >
                取消
              </button>
              <button
                onClick={handleConfirmBatchMoveApis}
                className="px-4 py-2 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:bg-zinc-800 dark:hover:bg-zinc-200"
              >
                确定移动
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================== ADD FONT CHOICE MODAL (Requirement 2) ==================== */}
      {showAddFontChoiceModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 animate-in fade-in pt-[calc(1rem+env(safe-area-inset-top,0px))]">
          <div className="w-full max-w-xs bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl p-5 space-y-4">
            <div className="flex items-center justify-between border-b border-zinc-100 dark:border-zinc-800 pb-3">
              <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100">添加字体</h3>
              <button
                onClick={() => setShowAddFontChoiceModal(false)}
                className="p-1 rounded-lg text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-200"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            {/* List options aligned with phone link modal style */}
            <div className="space-y-2">
              <button
                onClick={() => {
                  setShowAddFontChoiceModal(false);
                  setShowAddFontUrlModal(true);
                }}
                className="w-full p-3 text-left bg-zinc-50 dark:bg-zinc-800/80 hover:bg-zinc-100 dark:hover:bg-zinc-800 border border-zinc-200 dark:border-zinc-700/80 rounded-xl transition-all group flex items-center justify-between"
              >
                <div>
                  <div className="text-xs font-bold text-zinc-900 dark:text-zinc-100 group-hover:text-rose-500 transition-colors">
                    上传字体 URL
                  </div>
                  <div className="text-[11px] text-zinc-400 mt-0.5">手动填写字体名称与网络网络 URL 地址</div>
                </div>
                <ArrowRightLeft className="w-4 h-4 text-zinc-400 group-hover:text-rose-500 flex-shrink-0 ml-2" />
              </button>

              <button
                onClick={() => {
                  setShowAddFontChoiceModal(false);
                  openBatchUpload('fonts');
                }}
                className="w-full p-3 text-left bg-zinc-50 dark:bg-zinc-800/80 hover:bg-zinc-100 dark:hover:bg-zinc-800 border border-zinc-200 dark:border-zinc-700/80 rounded-xl transition-all group flex items-center justify-between"
              >
                <div>
                  <div className="text-xs font-bold text-zinc-900 dark:text-zinc-100 group-hover:text-rose-500 transition-colors">
                    上传文件
                  </div>
                  <div className="text-[11px] text-zinc-400 mt-0.5">选择本地 .ttf / .otf 文件，自动识别名称</div>
                </div>
                <Upload className="w-4 h-4 text-zinc-400 group-hover:text-rose-500 flex-shrink-0 ml-2" />
              </button>
            </div>

            <div className="pt-2 border-t border-zinc-100 dark:border-zinc-800 text-right">
              <button
                onClick={() => setShowAddFontChoiceModal(false)}
                className="px-3.5 py-1.5 text-xs font-medium text-zinc-500 hover:text-zinc-700 dark:hover:text-zinc-300"
              >
                取消
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================== ADD FONT URL MODAL (Requirement 3) ==================== */}
      {showAddFontUrlModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 animate-in fade-in pt-[calc(1rem+env(safe-area-inset-top,0px))]">
          <div className="w-full max-w-md bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl p-5 space-y-4">
            <div className="flex items-center justify-between border-b border-zinc-100 dark:border-zinc-800 pb-3">
              <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100">上传字体 URL</h3>
              <button
                onClick={() => setShowAddFontUrlModal(false)}
                className="p-1 rounded-lg text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-200"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="space-y-3">
              <div>
                <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                  字体名称 <span className="text-rose-500">*</span>
                </label>
                <input
                  type="text"
                  value={newFontUrlForm.name}
                  onChange={(e) => setNewFontUrlForm((p) => ({ ...p, name: e.target.value }))}
                  placeholder="例如：苹方黑体、思源宋体..."
                  className="w-full px-3 py-2 text-sm bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none text-zinc-900 dark:text-zinc-100"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                  字体 URL 地址 <span className="text-rose-500">*</span>
                </label>
                <input
                  type="text"
                  value={newFontUrlForm.url}
                  onChange={(e) => setNewFontUrlForm((p) => ({ ...p, url: e.target.value }))}
                  placeholder="https://example.com/font.ttf 或 .woff2"
                  className="w-full px-3 py-2 text-sm bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none font-mono text-zinc-900 dark:text-zinc-100 text-xs"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                  所属分组
                </label>
                <select
                  value={newFontUrlForm.category}
                  onChange={(e) => setNewFontUrlForm((p) => ({ ...p, category: e.target.value }))}
                  className="w-full px-3 py-2 text-sm bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none text-zinc-900 dark:text-zinc-100"
                >
                  {Array.from(
                    new Set(['默认', ...(appData.fontCategories || []), ...fontsList.map((p) => p.category || '默认')])
                  ).map((cat) => (
                    <option key={cat} value={cat}>
                      {cat}
                    </option>
                  ))}
                </select>
              </div>
            </div>

            {/* Bottom Buttons: 取消, 确认 (Requirement 3) */}
            <div className="flex items-center justify-end gap-2 pt-3 border-t border-zinc-100 dark:border-zinc-800">
              <button
                type="button"
                onClick={() => setShowAddFontUrlModal(false)}
                className="px-4 py-2 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800"
              >
                取消
              </button>
              <button
                type="button"
                onClick={handleAddFontUrl}
                className="px-4 py-2 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:bg-zinc-800 dark:hover:bg-zinc-200"
              >
                确认
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================== EDIT FONT MODAL (Requirement 5) ==================== */}
      {editingFont && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 animate-in fade-in pt-[calc(1rem+env(safe-area-inset-top,0px))]">
          <div className="w-full max-w-md bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl p-5 space-y-4">
            <div className="flex items-center justify-between border-b border-zinc-100 dark:border-zinc-800 pb-3">
              <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100">编辑字体面板</h3>
              <button
                onClick={() => setEditingFont(null)}
                className="p-1 rounded-lg text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-200"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="space-y-3">
              <div>
                <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                  字体名称
                </label>
                <input
                  type="text"
                  value={editingFont.name}
                  onChange={(e) => setEditingFont((p) => (p ? { ...p, name: e.target.value } : null))}
                  className="w-full px-3 py-2 text-sm bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none text-zinc-900 dark:text-zinc-100"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                  字体 URL（若是文件上传默认空，可手动添加）
                </label>
                <input
                  type="text"
                  value={editingFont.url || ''}
                  onChange={(e) => setEditingFont((p) => (p ? { ...p, url: e.target.value } : null))}
                  placeholder="https://..."
                  className="w-full px-3 py-2 text-sm bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none font-mono text-zinc-900 dark:text-zinc-100 text-xs"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                  所属分组
                </label>
                <select
                  value={editingFont.category || '默认'}
                  onChange={(e) => setEditingFont((p) => (p ? { ...p, category: e.target.value } : null))}
                  className="w-full px-3 py-2 text-sm bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none text-zinc-900 dark:text-zinc-100"
                >
                  {Array.from(
                    new Set(['默认', ...(appData.fontCategories || []), ...fontsList.map((p) => p.category || '默认')])
                  ).map((cat) => (
                    <option key={cat} value={cat}>
                      {cat}
                    </option>
                  ))}
                </select>
              </div>
            </div>

            {/* Bottom Buttons: 取消, 确定, 删除, 导出（ttf格式） (Requirement 5) */}
            <div className="flex items-center justify-between pt-3 border-t border-zinc-100 dark:border-zinc-800">
              <button
                type="button"
                onClick={() => {
                  if (confirm(`确定要删除字体 “${editingFont.name}” 吗？`)) {
                    const updated = fontsList.filter((p) => p.id !== editingFont.id);
                    updateAppData({ ...appData, fonts: updated });
                    if (activePreviewFontId === editingFont.id) setActivePreviewFontId(null);
                    setEditingFont(null);
                    showToast('字体已删除', 'info');
                  }
                }}
                className="px-3 py-1.5 text-xs font-bold rounded-lg bg-rose-50 text-rose-600 dark:bg-rose-950/60 dark:text-rose-400 hover:bg-rose-100 dark:hover:bg-rose-900/60 flex items-center gap-1"
              >
                <Trash2 className="w-3.5 h-3.5" /> 删除
              </button>

              <div className="flex items-center gap-1.5">
                <button
                  type="button"
                  onClick={() => handleExportFontFile(editingFont)}
                  className="px-3 py-1.5 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 hover:bg-zinc-100 dark:hover:bg-zinc-800 text-zinc-700 dark:text-zinc-300 flex items-center gap-1"
                  title="导出为 ttf 格式文件"
                >
                  <Download className="w-3.5 h-3.5" /> 导出(ttf)
                </button>
                <button
                  type="button"
                  onClick={() => setEditingFont(null)}
                  className="px-3 py-1.5 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800"
                >
                  取消
                </button>
                <button
                  type="button"
                  onClick={() => {
                    if (!editingFont.name.trim()) {
                      showToast('字体名称不能为空', 'error');
                      return;
                    }
                    const updated = fontsList.map((p) =>
                      p.id === editingFont.id ? { ...editingFont, updatedAt: Date.now() } : p
                    );
                    updateAppData({ ...appData, fonts: updated });
                    setEditingFont(null);
                    showToast('保存字体成功！', 'success');
                  }}
                  className="px-3.5 py-1.5 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:bg-zinc-800 dark:hover:bg-zinc-200"
                >
                  确定
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ==================== NEW FONT GROUP MODAL ==================== */}
      {showNewFontGroupModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 animate-in fade-in pt-[calc(1rem+env(safe-area-inset-top,0px))]">
          <div className="w-full max-w-sm bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl p-5 space-y-4">
            <div className="flex items-center justify-between border-b border-zinc-100 dark:border-zinc-800 pb-3">
              <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100">新建字体分组</h3>
              <button
                onClick={() => setShowNewFontGroupModal(false)}
                className="p-1 rounded-lg text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-200"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div>
              <label className="text-xs font-semibold text-zinc-700 dark:text-zinc-300 block mb-1">
                分组名称
              </label>
              <input
                type="text"
                value={newFontGroupName}
                onChange={(e) => setNewFontGroupName(e.target.value)}
                placeholder="请输入分组名字..."
                className="w-full px-3 py-2 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-xl text-zinc-900 dark:text-zinc-100 focus:outline-none"
              />
            </div>

            <div className="flex justify-end gap-2 pt-2 border-t border-zinc-100 dark:border-zinc-800">
              <button
                onClick={() => setShowNewFontGroupModal(false)}
                className="px-4 py-2 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800"
              >
                取消
              </button>
              <button
                onClick={handleCreateNewFontCategory}
                className="px-4 py-2 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:bg-zinc-800 dark:hover:bg-zinc-200"
              >
                确定新建
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================== MANAGE FONT GROUP MODAL ==================== */}
      {managingFontCategory && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 animate-in fade-in pt-[calc(1rem+env(safe-area-inset-top,0px))]">
          <div className="w-full max-w-sm bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl p-5 space-y-4">
            <div className="flex items-center justify-between border-b border-zinc-100 dark:border-zinc-800 pb-3">
              <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100">
                管理分组 - “{managingFontCategory}”
              </h3>
              <button
                onClick={() => setManagingFontCategory(null)}
                className="p-1 rounded-lg text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-200"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            {/* Rename Section */}
            <div className="space-y-2">
              <label className="text-xs font-semibold text-zinc-500 block">重命名分组</label>
              <div className="flex gap-2">
                <input
                  type="text"
                  value={renameFontCategoryInput}
                  onChange={(e) => setRenameFontCategoryInput(e.target.value)}
                  className="flex-1 px-3 py-1.5 text-xs bg-zinc-50 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 rounded-lg text-zinc-900 dark:text-zinc-100 focus:outline-none"
                />
                <button
                  onClick={handleRenameFontCategory}
                  className="px-3 py-1.5 text-xs font-medium bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 rounded-lg hover:opacity-90"
                >
                  保存名字
                </button>
              </div>
            </div>

            {/* Delete Group Dialog with Circular Toggle (Requirement 1) */}
            <div className="pt-3 border-t border-zinc-200 dark:border-zinc-800 space-y-3">
              <div className="text-xs font-bold text-rose-600 dark:text-rose-400">删除分组</div>
              <p className="text-xs text-zinc-500 dark:text-zinc-400 leading-relaxed">
                是否要删除分组“{managingFontCategory}”？删除前请选择是否一并清理组内字体。
              </p>

              {/* Circular Selection Toggle */}
              <div
                onClick={() => setDeleteItemsWithFontCategory((prev) => !prev)}
                className="flex items-center gap-3 p-3 bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-xl cursor-pointer hover:bg-zinc-100 dark:hover:bg-zinc-800 transition-colors"
              >
                <div
                  className={`w-5 h-5 rounded-full border-2 flex items-center justify-center transition-all ${
                    deleteItemsWithFontCategory
                      ? 'border-zinc-900 dark:border-zinc-100 bg-white dark:bg-zinc-900'
                      : 'border-zinc-400 dark:border-zinc-600 bg-transparent'
                  }`}
                >
                  {deleteItemsWithFontCategory && (
                    <div className="w-2.5 h-2.5 rounded-full bg-zinc-900 dark:bg-zinc-100" />
                  )}
                </div>

                <div className="text-xs font-medium text-zinc-800 dark:text-zinc-200">
                  是否要一并把分组里的字体给删掉？
                </div>
              </div>

              <div className="flex justify-end gap-2 pt-2">
                <button
                  onClick={() => setManagingFontCategory(null)}
                  className="px-3 py-1.5 text-xs font-medium text-zinc-600 dark:text-zinc-400 hover:bg-zinc-100 dark:hover:bg-zinc-800 rounded-lg"
                >
                  取消
                </button>
                <button
                  onClick={handleConfirmDeleteFontCategory}
                  className="px-4 py-1.5 text-xs font-bold bg-rose-600 text-white hover:bg-rose-700 rounded-lg"
                >
                  确认删除分组
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ==================== BATCH MOVE FONT MODAL ==================== */}
      {showFontBatchMoveModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 animate-in fade-in pt-[calc(1rem+env(safe-area-inset-top,0px))]">
          <div className="w-full max-w-sm bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl p-5 space-y-4">
            <div className="flex items-center justify-between border-b border-zinc-100 dark:border-zinc-800 pb-3">
              <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100">批量移动分组</h3>
              <button
                onClick={() => setShowFontBatchMoveModal(false)}
                className="p-1 rounded-lg text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-200"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div>
              <label className="text-xs font-semibold text-zinc-700 dark:text-zinc-300 block mb-1">
                选择目标分组
              </label>
              <select
                value={batchTargetFontCategory}
                onChange={(e) => setBatchTargetFontCategory(e.target.value)}
                className="w-full px-3 py-2 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-xl text-zinc-900 dark:text-zinc-100 focus:outline-none"
              >
                <option value="">-- 请选择分组 --</option>
                {Array.from(
                  new Set(['默认', ...(appData.fontCategories || []), ...fontsList.map((p) => p.category || '默认')])
                ).map((cat) => (
                  <option key={cat} value={cat}>
                    {cat}
                  </option>
                ))}
              </select>
            </div>

            <div className="flex justify-end gap-2 pt-2 border-t border-zinc-100 dark:border-zinc-800">
              <button
                onClick={() => setShowFontBatchMoveModal(false)}
                className="px-4 py-2 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800"
              >
                取消
              </button>
              <button
                onClick={handleConfirmBatchMoveFonts}
                className="px-4 py-2 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:bg-zinc-800 dark:hover:bg-zinc-200"
              >
                确定移动
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================== ADD EXTRA STORY CHOICE MODAL ==================== */}
      {showAddExtraStoryChoiceModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 animate-in fade-in pt-[calc(1rem+env(safe-area-inset-top,0px))]">
          <div className="w-full max-w-xs bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl p-5 space-y-4">
            <div className="flex items-center justify-between border-b border-zinc-100 dark:border-zinc-800 pb-3">
              <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100">添加番外小剧场</h3>
              <button
                onClick={() => setShowAddExtraStoryChoiceModal(false)}
                className="p-1 rounded-lg text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-200"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="space-y-2">
              <button
                onClick={() => {
                  setShowAddExtraStoryChoiceModal(false);
                  setShowAddExtraStoryManualModal(true);
                }}
                className="w-full p-3 text-left bg-zinc-50 dark:bg-zinc-800/80 hover:bg-zinc-100 dark:hover:bg-zinc-800 border border-zinc-200 dark:border-zinc-700/80 rounded-xl transition-all group flex items-center justify-between"
              >
                <div>
                  <div className="text-xs font-bold text-zinc-900 dark:text-zinc-100 group-hover:text-rose-500 transition-colors">
                    手动录入
                  </div>
                  <div className="text-[11px] text-zinc-400 mt-0.5">直接填写番外名称、作者及长文本内容</div>
                </div>
                <Edit3 className="w-4 h-4 text-zinc-400 group-hover:text-rose-500 flex-shrink-0 ml-2" />
              </button>

              <button
                onClick={() => {
                  setShowAddExtraStoryChoiceModal(false);
                  openBatchUpload('extras');
                }}
                className="w-full p-3 text-left bg-zinc-50 dark:bg-zinc-800/80 hover:bg-zinc-100 dark:hover:bg-zinc-800 border border-zinc-200 dark:border-zinc-700/80 rounded-xl transition-all group flex items-center justify-between"
              >
                <div>
                  <div className="text-xs font-bold text-zinc-900 dark:text-zinc-100 group-hover:text-rose-500 transition-colors">
                    上传文档 (.docx / .txt)
                  </div>
                  <div className="text-[11px] text-zinc-400 mt-0.5">选择本地文档，支持自动拆分多个番外</div>
                </div>
                <Upload className="w-4 h-4 text-zinc-400 group-hover:text-rose-500 flex-shrink-0 ml-2" />
              </button>
            </div>

            <div className="pt-2 border-t border-zinc-100 dark:border-zinc-800 text-right">
              <button
                onClick={() => setShowAddExtraStoryChoiceModal(false)}
                className="px-3.5 py-1.5 text-xs font-medium text-zinc-500 hover:text-zinc-700 dark:hover:text-zinc-300"
              >
                取消
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================== ADD EXTRA STORY MANUAL MODAL ==================== */}
      {showAddExtraStoryManualModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 animate-in fade-in pt-[calc(1rem+env(safe-area-inset-top,0px))]">
          <div className="w-full max-w-lg bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl p-5 space-y-4">
            <div className="flex items-center justify-between border-b border-zinc-100 dark:border-zinc-800 pb-3">
              <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100">手动录入番外</h3>
              <button
                onClick={() => setShowAddExtraStoryManualModal(false)}
                className="p-1 rounded-lg text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-200"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="space-y-3">
              <div>
                <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                  番外名称 <span className="text-rose-500">*</span>
                </label>
                <input
                  type="text"
                  value={newExtraStoryManualForm.title}
                  onChange={(e) => setNewExtraStoryManualForm((p) => ({ ...p, title: e.target.value }))}
                  placeholder="请输入番外小剧场名称..."
                  className="w-full px-3 py-2 text-sm bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none text-zinc-900 dark:text-zinc-100"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                  作者 / 来源
                </label>
                <input
                  type="text"
                  value={newExtraStoryManualForm.author}
                  onChange={(e) => setNewExtraStoryManualForm((p) => ({ ...p, author: e.target.value }))}
                  placeholder="例如：原作者名、QQ群/社区等..."
                  className="w-full px-3 py-2 text-sm bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none text-zinc-900 dark:text-zinc-100"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                  所属分组
                </label>
                <select
                  value={newExtraStoryManualForm.category}
                  onChange={(e) => setNewExtraStoryManualForm((p) => ({ ...p, category: e.target.value }))}
                  className="w-full px-3 py-2 text-sm bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none text-zinc-900 dark:text-zinc-100"
                >
                  {Array.from(
                    new Set(['默认', ...(appData.extraStoryCategories || []), ...extraStoriesList.map((p) => p.category || '默认')])
                  ).map((cat) => (
                    <option key={cat} value={cat}>
                      {cat}
                    </option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                  番外内容 (长文本) <span className="text-rose-500">*</span>
                </label>
                <textarea
                  rows={6}
                  value={newExtraStoryManualForm.content}
                  onChange={(e) => setNewExtraStoryManualForm((p) => ({ ...p, content: e.target.value }))}
                  placeholder="在此黏贴或输入番外正文内容..."
                  className="w-full px-3 py-2 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none text-zinc-900 dark:text-zinc-100 leading-relaxed font-sans"
                />
              </div>
            </div>

            <div className="flex items-center justify-end gap-2 pt-3 border-t border-zinc-100 dark:border-zinc-800">
              <button
                type="button"
                onClick={() => setShowAddExtraStoryManualModal(false)}
                className="px-4 py-2 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800"
              >
                取消
              </button>
              <button
                type="button"
                onClick={handleAddExtraStoryManual}
                className="px-4 py-2 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:bg-zinc-800 dark:hover:bg-zinc-200"
              >
                确认录入
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================== EDIT EXTRA STORY MODAL (Requirement 4 & 5) ==================== */}
      {editingExtraStory && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 animate-in fade-in pt-[calc(1rem+env(safe-area-inset-top,0px))]">
          <div className="w-full max-w-2xl bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl p-5 space-y-4 max-h-[90vh] flex flex-col">
            <div className="flex items-center justify-between border-b border-zinc-100 dark:border-zinc-800 pb-3 flex-shrink-0">
              <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100">番外详情与编辑</h3>
              <button
                onClick={() => setEditingExtraStory(null)}
                className="p-1 rounded-lg text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-200"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="space-y-3 overflow-y-auto flex-1 pr-1">
              <div>
                <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                  番外名称
                </label>
                <input
                  type="text"
                  value={editingExtraStory.title}
                  onChange={(e) => setEditingExtraStory((p) => (p ? { ...p, title: e.target.value } : null))}
                  className="w-full px-3 py-2 text-sm bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none text-zinc-900 dark:text-zinc-100 font-bold"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                  作者 / 来源
                </label>
                <input
                  type="text"
                  value={editingExtraStory.author || ''}
                  onChange={(e) => setEditingExtraStory((p) => (p ? { ...p, author: e.target.value } : null))}
                  placeholder="可在此填写作者或来源..."
                  className="w-full px-3 py-2 text-sm bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none text-zinc-900 dark:text-zinc-100"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                  所属分组
                </label>
                <select
                  value={editingExtraStory.category || '默认'}
                  onChange={(e) => setEditingExtraStory((p) => (p ? { ...p, category: e.target.value } : null))}
                  className="w-full px-3 py-2 text-sm bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none text-zinc-900 dark:text-zinc-100"
                >
                  {Array.from(
                    new Set(['默认', ...(appData.extraStoryCategories || []), ...extraStoriesList.map((p) => p.category || '默认')])
                  ).map((cat) => (
                    <option key={cat} value={cat}>
                      {cat}
                    </option>
                  ))}
                </select>
              </div>

              <div>
                <div className="flex items-center justify-between mb-1">
                  <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300">
                    番外内容
                  </label>
                  <div className="flex items-center gap-1.5">
                    {/* Copy Button */}
                    <button
                      type="button"
                      onClick={() => {
                        void copyTextReliable(editingExtraStory.content, '已复制番外内容到剪贴板！');
                      }}
                      className="px-2.5 py-1 text-[11px] font-medium rounded-md bg-zinc-100 dark:bg-zinc-800 hover:bg-zinc-200 dark:hover:bg-zinc-700 text-zinc-700 dark:text-zinc-300 transition-colors flex items-center gap-1"
                    >
                      <Copy className="w-3 h-3" /> 复制
                    </button>

                    {/* Expand Button */}
                    <button
                      type="button"
                      onClick={() => setIsContentExpanded(true)}
                      className="px-2.5 py-1 text-[11px] font-medium rounded-md bg-zinc-100 dark:bg-zinc-800 hover:bg-zinc-200 dark:hover:bg-zinc-700 text-zinc-700 dark:text-zinc-300 transition-colors flex items-center gap-1"
                    >
                      <Maximize2 className="w-3 h-3" /> 放大展开
                    </button>
                  </div>
                </div>

                <textarea
                  rows={10}
                  value={editingExtraStory.content}
                  onChange={(e) => setEditingExtraStory((p) => (p ? { ...p, content: e.target.value } : null))}
                  className="w-full px-3 py-2 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none text-zinc-900 dark:text-zinc-100 leading-relaxed font-sans"
                />
              </div>
            </div>

            {/* Bottom Buttons: 取消, 确定, 删除 (Requirement 5) */}
            <div className="flex items-center justify-between pt-3 border-t border-zinc-100 dark:border-zinc-800 flex-shrink-0">
              <button
                type="button"
                onClick={() => {
                  if (confirm(`确定要删除番外 “${editingExtraStory.title}” 吗？`)) {
                    const updated = extraStoriesList.filter((p) => p.id !== editingExtraStory.id);
                    updateAppData({ ...appData, extraStories: updated });
                    setEditingExtraStory(null);
                    showToast('番外已删除', 'info');
                  }
                }}
                className="px-3.5 py-1.5 text-xs font-bold rounded-lg bg-rose-50 text-rose-600 dark:bg-rose-950/60 dark:text-rose-400 hover:bg-rose-100 dark:hover:bg-rose-900/60 flex items-center gap-1"
              >
                <Trash2 className="w-3.5 h-3.5" /> 删除
              </button>

              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={() => setEditingExtraStory(null)}
                  className="px-4 py-1.5 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800"
                >
                  取消
                </button>
                <button
                  type="button"
                  onClick={() => {
                    if (!editingExtraStory.title.trim()) {
                      showToast('番外名称不能为空', 'error');
                      return;
                    }
                    const updated = extraStoriesList.map((p) =>
                      p.id === editingExtraStory.id ? { ...editingExtraStory, updatedAt: Date.now() } : p
                    );
                    updateAppData({ ...appData, extraStories: updated });
                    setEditingExtraStory(null);
                    showToast('保存成功！', 'success');
                  }}
                  className="px-4 py-1.5 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:bg-zinc-800 dark:hover:bg-zinc-200"
                >
                  确定
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ==================== EXPANDED CONTENT MODAL ==================== */}
      {isContentExpanded && editingExtraStory && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/80 p-4 sm:p-6 animate-in fade-in">
          <div className="w-full max-w-4xl h-[85vh] bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl p-6 flex flex-col space-y-4">
            <div className="flex items-center justify-between border-b border-zinc-100 dark:border-zinc-800 pb-3 flex-shrink-0">
              <div className="flex items-center gap-2">
                <h3 className="text-base font-bold text-zinc-900 dark:text-zinc-100">
                  {editingExtraStory.title} (放大正文模式)
                </h3>
              </div>
              <div className="flex items-center gap-2">
                <button
                  onClick={() => {
                    void copyTextReliable(editingExtraStory.content, '已复制番外内容到剪贴板！');
                  }}
                  className="px-3 py-1.5 text-xs font-medium rounded-lg bg-zinc-100 dark:bg-zinc-800 hover:bg-zinc-200 dark:hover:bg-zinc-700 text-zinc-700 dark:text-zinc-300 flex items-center gap-1.5"
                >
                  <Copy className="w-3.5 h-3.5" /> 复制正文
                </button>
                <button
                  onClick={() => setIsContentExpanded(false)}
                  className="p-1.5 rounded-lg text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-200"
                >
                  <X className="w-5 h-5" />
                </button>
              </div>
            </div>

            <textarea
              value={editingExtraStory.content}
              onChange={(e) => setEditingExtraStory((p) => (p ? { ...p, content: e.target.value } : null))}
              className="w-full flex-1 p-4 bg-zinc-50 dark:bg-zinc-800/50 border border-zinc-200 dark:border-zinc-700 rounded-xl text-sm leading-relaxed focus:outline-none text-zinc-900 dark:text-zinc-100 font-sans resize-none"
            />

            <div className="flex justify-end flex-shrink-0 pt-2 border-t border-zinc-100 dark:border-zinc-800">
              <button
                onClick={() => setIsContentExpanded(false)}
                className="px-5 py-2 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:opacity-90"
              >
                关闭放大视图
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================== NEW EXTRA STORY GROUP MODAL ==================== */}
      {showNewExtraStoryGroupModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 animate-in fade-in pt-[calc(1rem+env(safe-area-inset-top,0px))]">
          <div className="w-full max-w-sm bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl p-5 space-y-4">
            <div className="flex items-center justify-between border-b border-zinc-100 dark:border-zinc-800 pb-3">
              <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100">新建番外分组</h3>
              <button
                onClick={() => setShowNewExtraStoryGroupModal(false)}
                className="p-1 rounded-lg text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-200"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div>
              <label className="text-xs font-semibold text-zinc-700 dark:text-zinc-300 block mb-1">
                分组名称
              </label>
              <input
                type="text"
                value={newExtraStoryGroupName}
                onChange={(e) => setNewExtraStoryGroupName(e.target.value)}
                placeholder="请输入分组名字..."
                className="w-full px-3 py-2 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-xl text-zinc-900 dark:text-zinc-100 focus:outline-none"
              />
            </div>

            <div className="flex justify-end gap-2 pt-2 border-t border-zinc-100 dark:border-zinc-800">
              <button
                onClick={() => setShowNewExtraStoryGroupModal(false)}
                className="px-4 py-2 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800"
              >
                取消
              </button>
              <button
                onClick={handleCreateNewExtraStoryGroup}
                className="px-4 py-2 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:bg-zinc-800 dark:hover:bg-zinc-200"
              >
                确定新建
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================== MANAGE EXTRA STORY GROUP MODAL ==================== */}
      {managingExtraStoryCategory && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 animate-in fade-in pt-[calc(1rem+env(safe-area-inset-top,0px))]">
          <div className="w-full max-w-sm bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl p-5 space-y-4">
            <div className="flex items-center justify-between border-b border-zinc-100 dark:border-zinc-800 pb-3">
              <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100">
                管理分组 - “{managingExtraStoryCategory}”
              </h3>
              <button
                onClick={() => setManagingExtraStoryCategory(null)}
                className="p-1 rounded-lg text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-200"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            {/* Rename Section */}
            <div className="space-y-2">
              <label className="text-xs font-semibold text-zinc-500 block">重命名分组</label>
              <div className="flex gap-2">
                <input
                  type="text"
                  value={renameExtraStoryCategoryInput}
                  onChange={(e) => setRenameExtraStoryCategoryInput(e.target.value)}
                  className="flex-1 px-3 py-1.5 text-xs bg-zinc-50 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 rounded-lg text-zinc-900 dark:text-zinc-100 focus:outline-none"
                />
                <button
                  onClick={handleRenameExtraStoryCategory}
                  className="px-3 py-1.5 text-xs font-medium bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 rounded-lg hover:opacity-90"
                >
                  保存名字
                </button>
              </div>
            </div>

            {/* Delete Group Dialog with Circular Toggle (Requirement 1 & ST cards logic) */}
            <div className="pt-3 border-t border-zinc-200 dark:border-zinc-800 space-y-3">
              <div className="text-xs font-bold text-rose-600 dark:text-rose-400">删除分组</div>
              <p className="text-xs text-zinc-500 dark:text-zinc-400 leading-relaxed">
                是否要删除分组“{managingExtraStoryCategory}”？删除前请选择是否一并清理组内番外。
              </p>

              {/* Circular Selection Toggle */}
              <div
                onClick={() => setDeleteItemsWithExtraStoryCategory((prev) => !prev)}
                className="flex items-center gap-3 p-3 bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-xl cursor-pointer hover:bg-zinc-100 dark:hover:bg-zinc-800 transition-colors"
              >
                <div
                  className={`w-5 h-5 rounded-full border-2 flex items-center justify-center transition-all ${
                    deleteItemsWithExtraStoryCategory
                      ? 'border-zinc-900 dark:border-zinc-100 bg-white dark:bg-zinc-900'
                      : 'border-zinc-400 dark:border-zinc-600 bg-transparent'
                  }`}
                >
                  {deleteItemsWithExtraStoryCategory && (
                    <div className="w-2.5 h-2.5 rounded-full bg-zinc-900 dark:bg-zinc-100" />
                  )}
                </div>

                <div className="text-xs font-medium text-zinc-800 dark:text-zinc-200">
                  是否要一并把分组里的番外给删掉？
                </div>
              </div>

              <div className="flex justify-end gap-2 pt-2">
                <button
                  onClick={() => setManagingExtraStoryCategory(null)}
                  className="px-3 py-1.5 text-xs font-medium text-zinc-600 dark:text-zinc-400 hover:bg-zinc-100 dark:hover:bg-zinc-800 rounded-lg"
                >
                  取消
                </button>
                <button
                  onClick={handleConfirmDeleteExtraStoryCategory}
                  className="px-4 py-1.5 text-xs font-bold bg-rose-600 text-white hover:bg-rose-700 rounded-lg"
                >
                  确认删除分组
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ==================== BATCH MOVE EXTRA STORY MODAL ==================== */}
      {showExtraStoryBatchMoveModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 animate-in fade-in pt-[calc(1rem+env(safe-area-inset-top,0px))]">
          <div className="w-full max-w-sm bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl p-5 space-y-4">
            <div className="flex items-center justify-between border-b border-zinc-100 dark:border-zinc-800 pb-3">
              <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100">批量移动分组</h3>
              <button
                onClick={() => setShowExtraStoryBatchMoveModal(false)}
                className="p-1 rounded-lg text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-200"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div>
              <label className="text-xs font-semibold text-zinc-700 dark:text-zinc-300 block mb-1">
                选择目标分组
              </label>
              <select
                value={batchTargetExtraStoryCategory}
                onChange={(e) => setBatchTargetExtraStoryCategory(e.target.value)}
                className="w-full px-3 py-2 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-xl text-zinc-900 dark:text-zinc-100 focus:outline-none"
              >
                <option value="">-- 请选择分组 --</option>
                {Array.from(
                  new Set(['默认', ...(appData.extraStoryCategories || []), ...extraStoriesList.map((p) => p.category || '默认')])
                ).map((cat) => (
                  <option key={cat} value={cat}>
                    {cat}
                  </option>
                ))}
              </select>
            </div>

            <div className="flex justify-end gap-2 pt-2 border-t border-zinc-100 dark:border-zinc-800">
              <button
                onClick={() => setShowExtraStoryBatchMoveModal(false)}
                className="px-4 py-2 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800"
              >
                取消
              </button>
              <button
                onClick={handleConfirmBatchMoveExtraStories}
                className="px-4 py-2 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:bg-zinc-800 dark:hover:bg-zinc-200"
              >
                确定移动
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================== EDIT STICKER PACK DETAIL MODAL ==================== */}
      {editingStickerPack && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 animate-in fade-in pt-[calc(1rem+env(safe-area-inset-top,0px))]">
          <div className="w-full max-w-3xl h-[85vh] bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl flex flex-col overflow-hidden">
            {/* Modal Header & Tabs */}
            <div className="p-4 border-b border-zinc-200 dark:border-zinc-800 flex items-center justify-between flex-shrink-0">
              <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100 truncate max-w-xs">
                表情包图集 - {editingStickerPack.title}
              </h3>
              <button
                onClick={() => setEditingStickerPack(null)}
                className="p-1 rounded-lg text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-200"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            {/* Tabs Navigation Bar */}
            <div className="border-b border-zinc-200 dark:border-zinc-800 flex gap-2 px-6 overflow-x-auto scrollbar-none flex-shrink-0 bg-zinc-50/50 dark:bg-zinc-900/50">
              {[
                { id: 'info', name: '基础信息' },
                { id: 'atlas', name: '表情包图集详情' },
                { id: 'export', name: '格式转换页' },
              ].map((tab) => (
                <button
                  key={tab.id}
                  onClick={() => setStickerDetailTab(tab.id as any)}
                  className={`py-3 px-4 text-xs font-semibold border-b-2 whitespace-nowrap transition-colors ${
                    stickerDetailTab === tab.id
                      ? 'border-zinc-900 text-zinc-900 dark:border-zinc-100 dark:text-zinc-100 font-bold'
                      : 'border-transparent text-zinc-500 hover:text-zinc-800 dark:hover:text-zinc-200'
                  }`}
                >
                  {tab.name}
                </button>
              ))}
            </div>

            {/* Tab Body Content */}
            <div className="flex-1 overflow-y-auto p-6 space-y-4">
              {/* TAB 1: 基础信息 */}
              {stickerDetailTab === 'info' && (
                <div className="space-y-4 max-w-md mx-auto py-4">
                  <div>
                    <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                      表情包图集名称 <span className="text-rose-500">*</span>
                    </label>
                    <input
                      type="text"
                      value={editingStickerPack.title}
                      onChange={(e) => setEditingStickerPack((p) => (p ? { ...p, title: e.target.value } : null))}
                      className="w-full px-3 py-2 text-sm bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none text-zinc-900 dark:text-zinc-100 font-bold"
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                      作者 / 来源
                    </label>
                    <input
                      type="text"
                      value={editingStickerPack.author || ''}
                      onChange={(e) => setEditingStickerPack((p) => (p ? { ...p, author: e.target.value } : null))}
                      placeholder="例如：画师名、原出处或社区..."
                      className="w-full px-3 py-2 text-sm bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none text-zinc-900 dark:text-zinc-100"
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                      所属分组
                    </label>
                    <select
                      value={editingStickerPack.category || '默认'}
                      onChange={(e) => setEditingStickerPack((p) => (p ? { ...p, category: e.target.value } : null))}
                      className="w-full px-3 py-2 text-sm bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none text-zinc-900 dark:text-zinc-100"
                    >
                      {Array.from(
                        new Set(['默认', ...(appData.stickerCategories || []), ...stickerPacksList.map((p) => p.category || '默认')])
                      ).map((cat) => (
                        <option key={cat} value={cat}>
                          {cat}
                        </option>
                      ))}
                    </select>
                  </div>
                </div>
              )}

              {/* TAB 2: 表情包图集详情 (Preview & Edit List) */}
              {stickerDetailTab === 'atlas' && (
                <div className="space-y-3">
                  <div className="flex items-center justify-between pb-2 border-b border-zinc-200 dark:border-zinc-800">
                    <span className="text-xs font-bold text-zinc-700 dark:text-zinc-300">
                      包含表情 ({editingStickerPack.items.length} 个)
                    </span>
                    <button
                      type="button"
                      onClick={() => {
                        setEditingStickerPack((p) =>
                          p
                            ? {
                                ...p,
                                items: [
                                  ...p.items,
                                  {
                                    id: `stk_${Date.now()}_${Math.random().toString(36).substring(2, 6)}`,
                                    name: `表情 ${p.items.length + 1}`,
                                    url: '',
                                  },
                                ],
                              }
                            : null
                        );
                      }}
                      className="px-3 py-1.5 text-xs font-semibold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:opacity-90 inline-flex items-center gap-1"
                    >
                      <Plus className="w-3.5 h-3.5" /> 添加新表情
                    </button>
                  </div>

                  <div className="space-y-2.5 max-h-[55vh] overflow-y-auto pr-1">
                    {editingStickerPack.items.map((item, idx) => (
                      <div
                        key={item.id}
                        className="flex items-center gap-3 p-3 bg-zinc-50 dark:bg-zinc-800/50 border border-zinc-200 dark:border-zinc-700/60 rounded-xl"
                      >
                        {/* Left: Fixed uniform size image preview */}
                        <div className="w-16 h-16 sm:w-20 sm:h-20 flex-shrink-0 bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-700 rounded-lg overflow-hidden flex items-center justify-center p-1 relative">
                          {item.url ? (
                            <img
                              src={item.url}
                              alt={item.name}
                              className="w-full h-full object-cover"
                              onError={(e) => {
                                (e.target as HTMLElement).style.display = 'none';
                              }}
                            />
                          ) : (
                            <ImageIcon className="w-6 h-6 text-zinc-300 dark:text-zinc-600" />
                          )}
                        </div>

                        {/* Right: Top Name & Bottom URL */}
                        <div className="flex-1 min-w-0 space-y-2">
                          <div>
                            <label className="text-[10px] font-semibold text-zinc-400 block mb-0.5">表情包名称</label>
                            <input
                              type="text"
                              value={item.name}
                              onChange={(e) => {
                                const val = e.target.value;
                                setEditingStickerPack((p) =>
                                  p
                                    ? {
                                        ...p,
                                        items: p.items.map((it) => (it.id === item.id ? { ...it, name: val } : it)),
                                      }
                                    : null
                                );
                              }}
                              placeholder="如：开心 / 震惊..."
                              className="w-full px-2.5 py-1 text-xs bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none text-zinc-900 dark:text-zinc-100 font-semibold"
                            />
                          </div>

                          <div>
                            <label className="text-[10px] font-semibold text-zinc-400 block mb-0.5">表情包 URL</label>
                            <input
                              type="text"
                              value={item.url}
                              onChange={(e) => {
                                const val = e.target.value;
                                setEditingStickerPack((p) =>
                                  p
                                    ? {
                                        ...p,
                                        items: p.items.map((it) => (it.id === item.id ? { ...it, url: val } : it)),
                                      }
                                    : null
                                );
                              }}
                              placeholder="http://... 或 https://..."
                              className="w-full px-2.5 py-1 text-xs bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-700 rounded-lg focus:outline-none text-zinc-900 dark:text-zinc-100 font-mono text-[11px]"
                            />
                          </div>
                        </div>

                        {/* Delete Sticker Item */}
                        <button
                          type="button"
                          onClick={() => {
                            setEditingStickerPack((p) =>
                              p ? { ...p, items: p.items.filter((it) => it.id !== item.id) } : null
                            );
                          }}
                          className="p-1.5 text-zinc-400 hover:text-rose-500 rounded-lg hover:bg-zinc-100 dark:hover:bg-zinc-800 transition-colors flex-shrink-0"
                          title="删除此表情"
                        >
                          <Trash2 className="w-4 h-4" />
                        </button>
                      </div>
                    ))}
                  </div>
                </div>
              )}

              {/* TAB 3: 格式转换页 */}
              {stickerDetailTab === 'export' && (
                <div className="space-y-4">
                  <div className="space-y-2">
                    <label className="block text-xs font-bold text-zinc-700 dark:text-zinc-300">
                      选择导出格式规范
                    </label>
                    <div className="grid grid-cols-2 sm:grid-cols-3 gap-2">
                      {[
                        { key: 'cnColon', label: '名称：url' },
                        { key: 'enColon', label: '名称:url' },
                        { key: 'noSpace', label: '名称url' },
                        { key: 'space', label: '名称 url' },
                        { key: 'urlSpaceName', label: 'url 名称' },
                        { key: 'urlName', label: 'url名称' },
                      ].map((fmt) => (
                        <button
                          key={fmt.key}
                          type="button"
                          onClick={() => setStickerExportFormat(fmt.key as any)}
                          className={`p-2.5 text-xs font-medium rounded-xl border text-center transition-all ${
                            stickerExportFormat === fmt.key
                              ? 'border-zinc-900 dark:border-zinc-100 bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 font-bold shadow-sm'
                              : 'border-zinc-200 dark:border-zinc-700 bg-zinc-50 dark:bg-zinc-800/60 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800'
                          }`}
                        >
                          {fmt.label}
                        </button>
                      ))}
                    </div>
                  </div>

                  {/* Formatted Output Preview Textarea */}
                  <div>
                    <div className="flex items-center justify-between mb-1">
                      <label className="block text-xs font-semibold text-zinc-500">格式转换预览</label>
                      <button
                        type="button"
                        onClick={() => {
                          const txt = generateStickerExportText(editingStickerPack, stickerExportFormat);
                          navigator.clipboard.writeText(txt);
                          showToast('已复制转换后的内容到剪贴板！', 'success');
                        }}
                        className="text-[11px] text-zinc-500 hover:text-zinc-800 dark:hover:text-zinc-200 flex items-center gap-1"
                      >
                        <Copy className="w-3 h-3" /> 复制预览文本
                      </button>
                    </div>
                    <textarea
                      readOnly
                      rows={6}
                      value={generateStickerExportText(editingStickerPack, stickerExportFormat)}
                      className="w-full px-3 py-2 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-xl font-mono leading-relaxed text-zinc-800 dark:text-zinc-200 focus:outline-none"
                    />
                  </div>

                  {/* Export Document Buttons */}
                  <div className="pt-2 border-t border-zinc-200 dark:border-zinc-800 space-y-2">
                    <label className="block text-xs font-bold text-zinc-700 dark:text-zinc-300">
                      按当前格式导出文档
                    </label>
                    <div className="grid grid-cols-3 gap-2">
                      <button
                        type="button"
                        onClick={() => exportStickerPackDocx(editingStickerPack, stickerExportFormat)}
                        className="py-2 text-xs font-bold rounded-xl bg-zinc-100 dark:bg-zinc-800 hover:bg-zinc-200 dark:hover:bg-zinc-700 text-zinc-800 dark:text-zinc-200 transition-colors flex items-center justify-center gap-1.5"
                      >
                        <Download className="w-3.5 h-3.5" /> 导出 .docx
                      </button>
                      <button
                        type="button"
                        onClick={() => exportStickerPackTxt(editingStickerPack, stickerExportFormat)}
                        className="py-2 text-xs font-bold rounded-xl bg-zinc-100 dark:bg-zinc-800 hover:bg-zinc-200 dark:hover:bg-zinc-700 text-zinc-800 dark:text-zinc-200 transition-colors flex items-center justify-center gap-1.5"
                      >
                        <Download className="w-3.5 h-3.5" /> 导出 .txt
                      </button>
                      <button
                        type="button"
                        onClick={() => exportStickerPackJson(editingStickerPack)}
                        className="py-2 text-xs font-bold rounded-xl bg-zinc-100 dark:bg-zinc-800 hover:bg-zinc-200 dark:hover:bg-zinc-700 text-zinc-800 dark:text-zinc-200 transition-colors flex items-center justify-center gap-1.5"
                      >
                        <Download className="w-3.5 h-3.5" /> 导出 .json
                      </button>
                    </div>
                  </div>
                </div>
              )}
            </div>

            {/* Modal Bottom Buttons: 取消, 确定, 删除 (Requirement 4) */}
            <div className="flex items-center justify-between p-4 border-t border-zinc-200 dark:border-zinc-800 bg-zinc-50 dark:bg-zinc-900/50 flex-shrink-0">
              <button
                type="button"
                onClick={() => {
                  if (confirm(`确定要删除表情包图集 “${editingStickerPack.title}” 吗？`)) {
                    const updated = stickerPacksList.filter((p) => p.id !== editingStickerPack.id);
                    updateAppData({ ...appData, stickerPacks: updated });
                    setEditingStickerPack(null);
                    showToast('表情包图集已删除', 'info');
                  }
                }}
                className="px-4 py-2 text-xs font-bold rounded-lg bg-rose-50 text-rose-600 dark:bg-rose-950/60 dark:text-rose-400 hover:bg-rose-100 dark:hover:bg-rose-900/60 transition-colors flex items-center gap-1.5"
              >
                <Trash2 className="w-3.5 h-3.5" /> 删除
              </button>

              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={() => setEditingStickerPack(null)}
                  className="px-4 py-2 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800 transition-colors"
                >
                  取消
                </button>
                <button
                  type="button"
                  onClick={() => {
                    if (!editingStickerPack.title.trim()) {
                      showToast('表情包图集名称不能为空', 'error');
                      return;
                    }
                    const updated = stickerPacksList.map((p) =>
                      p.id === editingStickerPack.id ? { ...editingStickerPack, updatedAt: Date.now() } : p
                    );
                    updateAppData({ ...appData, stickerPacks: updated });
                    setEditingStickerPack(null);
                    showToast('保存成功！', 'success');
                  }}
                  className="px-4 py-2 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:opacity-90 transition-opacity"
                >
                  确定
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ==================== NEW STICKER GROUP MODAL ==================== */}
      {showNewStickerGroupModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 animate-in fade-in pt-[calc(1rem+env(safe-area-inset-top,0px))]">
          <div className="w-full max-w-sm bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl p-5 space-y-4">
            <div className="flex items-center justify-between border-b border-zinc-100 dark:border-zinc-800 pb-3">
              <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100">新建表情包分组</h3>
              <button
                onClick={() => setShowNewStickerGroupModal(false)}
                className="p-1 rounded-lg text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-200"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div>
              <label className="text-xs font-semibold text-zinc-700 dark:text-zinc-300 block mb-1">
                分组名称
              </label>
              <input
                type="text"
                value={newStickerGroupName}
                onChange={(e) => setNewStickerGroupName(e.target.value)}
                placeholder="请输入分组名字..."
                className="w-full px-3 py-2 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-xl text-zinc-900 dark:text-zinc-100 focus:outline-none"
              />
            </div>

            <div className="flex justify-end gap-2 pt-2 border-t border-zinc-100 dark:border-zinc-800">
              <button
                onClick={() => setShowNewStickerGroupModal(false)}
                className="px-4 py-2 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800"
              >
                取消
              </button>
              <button
                onClick={handleCreateNewStickerGroup}
                className="px-4 py-2 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:bg-zinc-800 dark:hover:bg-zinc-200"
              >
                确定新建
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================== MANAGE STICKER GROUP MODAL ==================== */}
      {managingStickerCategory && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 animate-in fade-in pt-[calc(1rem+env(safe-area-inset-top,0px))]">
          <div className="w-full max-w-sm bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl p-5 space-y-4">
            <div className="flex items-center justify-between border-b border-zinc-100 dark:border-zinc-800 pb-3">
              <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100">
                管理分组 - “{managingStickerCategory}”
              </h3>
              <button
                onClick={() => setManagingStickerCategory(null)}
                className="p-1 rounded-lg text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-200"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            {/* Rename Section */}
            <div className="space-y-2">
              <label className="text-xs font-semibold text-zinc-500 block">重命名分组</label>
              <div className="flex gap-2">
                <input
                  type="text"
                  value={renameStickerCategoryInput}
                  onChange={(e) => setRenameStickerCategoryInput(e.target.value)}
                  className="flex-1 px-3 py-1.5 text-xs bg-zinc-50 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 rounded-lg text-zinc-900 dark:text-zinc-100 focus:outline-none"
                />
                <button
                  onClick={handleRenameStickerCategory}
                  className="px-3 py-1.5 text-xs font-medium bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 rounded-lg hover:opacity-90"
                >
                  保存名字
                </button>
              </div>
            </div>

            {/* Delete Group Dialog with Circular Toggle (Requirement 1 & ST cards logic) */}
            <div className="pt-3 border-t border-zinc-200 dark:border-zinc-800 space-y-3">
              <div className="text-xs font-bold text-rose-600 dark:text-rose-400">删除分组</div>
              <p className="text-xs text-zinc-500 dark:text-zinc-400 leading-relaxed">
                是否要删除分组“{managingStickerCategory}”？删除前请选择是否一并清理组内表情包。
              </p>

              {/* Circular Selection Toggle */}
              <div
                onClick={() => setDeleteItemsWithStickerCategory((prev) => !prev)}
                className="flex items-center gap-3 p-3 bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-xl cursor-pointer hover:bg-zinc-100 dark:hover:bg-zinc-800 transition-colors"
              >
                <div
                  className={`w-5 h-5 rounded-full border-2 flex items-center justify-center transition-all ${
                    deleteItemsWithStickerCategory
                      ? 'border-zinc-900 dark:border-zinc-100 bg-white dark:bg-zinc-900'
                      : 'border-zinc-400 dark:border-zinc-600 bg-transparent'
                  }`}
                >
                  {deleteItemsWithStickerCategory && (
                    <div className="w-2.5 h-2.5 rounded-full bg-zinc-900 dark:bg-zinc-100" />
                  )}
                </div>

                <div className="text-xs font-medium text-zinc-800 dark:text-zinc-200">
                  是否要一并把分组里的表情包给删掉？
                </div>
              </div>

              <div className="flex justify-end gap-2 pt-2">
                <button
                  onClick={() => setManagingStickerCategory(null)}
                  className="px-3 py-1.5 text-xs font-medium text-zinc-600 dark:text-zinc-400 hover:bg-zinc-100 dark:hover:bg-zinc-800 rounded-lg"
                >
                  取消
                </button>
                <button
                  onClick={handleConfirmDeleteStickerCategory}
                  className="px-4 py-1.5 text-xs font-bold bg-rose-600 text-white hover:bg-rose-700 rounded-lg"
                >
                  确认删除分组
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ==================== BATCH MOVE STICKER PACKS MODAL ==================== */}
      {showStickerBatchMoveModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 animate-in fade-in pt-[calc(1rem+env(safe-area-inset-top,0px))]">
          <div className="w-full max-w-sm bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl p-5 space-y-4">
            <div className="flex items-center justify-between border-b border-zinc-100 dark:border-zinc-800 pb-3">
              <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100">批量移动分组</h3>
              <button
                onClick={() => setShowStickerBatchMoveModal(false)}
                className="p-1 rounded-lg text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-200"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div>
              <label className="text-xs font-semibold text-zinc-700 dark:text-zinc-300 block mb-1">
                选择目标分组
              </label>
              <select
                value={batchTargetStickerCategory}
                onChange={(e) => setBatchTargetStickerCategory(e.target.value)}
                className="w-full px-3 py-2 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-xl text-zinc-900 dark:text-zinc-100 focus:outline-none"
              >
                <option value="">-- 请选择分组 --</option>
                {Array.from(
                  new Set(['默认', ...(appData.stickerCategories || []), ...stickerPacksList.map((p) => p.category || '默认')])
                ).map((cat) => (
                  <option key={cat} value={cat}>
                    {cat}
                  </option>
                ))}
              </select>
            </div>

            <div className="flex justify-end gap-2 pt-2 border-t border-zinc-100 dark:border-zinc-800">
              <button
                onClick={() => setShowStickerBatchMoveModal(false)}
                className="px-4 py-2 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800"
              >
                取消
              </button>
              <button
                onClick={handleConfirmBatchMoveStickerPacks}
                className="px-4 py-2 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:bg-zinc-800 dark:hover:bg-zinc-200"
              >
                确定移动
              </button>
            </div>
          </div>
        </div>
      )}

      {extraStoryImportPreview && (
        <div className="fixed inset-0 z-[89] flex items-center justify-center bg-black/60 p-3 pt-[calc(0.75rem+env(safe-area-inset-top,0px))]">
          <div className="w-full max-w-3xl max-h-[88vh] bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl overflow-hidden flex flex-col">
            <div className="p-4 border-b border-zinc-200 dark:border-zinc-800 flex items-center justify-between gap-3">
              <div className="min-w-0"><h3 className="text-sm font-bold truncate">预览导入番外</h3><p className="text-[11px] text-zinc-400 mt-0.5">已识别 {extraStoryImportPreview.length} 条 · 当前选择 {selectedExtraStoryImportIds.length} 条</p></div>
              <button onClick={() => { setExtraStoryImportPreview(null); setSelectedExtraStoryImportIds([]); }} className="p-1 text-zinc-400"><X className="w-4 h-4" /></button>
            </div>
            <div className="px-4 py-2 border-b border-zinc-100 dark:border-zinc-800 flex flex-wrap items-center gap-2">
              <button onClick={() => setSelectedExtraStoryImportIds(extraStoryImportPreview.map((e) => e.id))} className="px-2.5 py-1.5 text-xs rounded-lg border">全选</button>
              <button onClick={() => setSelectedExtraStoryImportIds([])} className="px-2.5 py-1.5 text-xs rounded-lg border">取消全选</button>
              <button onClick={mergeSelectedExtraStoryImports} disabled={selectedExtraStoryImportIds.length < 2} className="px-2.5 py-1.5 text-xs rounded-lg border disabled:opacity-40">融合所选</button>
              <span className="text-[11px] text-zinc-400 ml-auto">可直接编辑；选择后只导入选中的条目</span>
            </div>
            <div className="p-4 overflow-y-auto space-y-3">
              {extraStoryImportPreview.slice(0, extraStoryImportVisibleCount).map((item, index) => {
                const selected = selectedExtraStoryImportIds.includes(item.id);
                return (
                  <div key={item.id} className={`rounded-xl border p-3 ${selected ? 'border-zinc-300 dark:border-zinc-700' : 'border-zinc-200 dark:border-zinc-800 opacity-60'}`}>
                    <div className="flex items-center gap-2 mb-2">
                      <button onClick={() => toggleExtraStoryImportSelection(item.id)} className="flex-shrink-0" title={selected ? '取消导入' : '选择导入'}>
                        {selected ? <CheckCircle2 className="w-5 h-5 text-zinc-900 dark:text-zinc-100" /> : <Circle className="w-5 h-5 text-zinc-400" />}
                      </button>
                      <span className="text-[11px] text-zinc-400">#{index + 1}</span>
                      <input value={item.title} onChange={(e) => updateExtraStoryImportPreviewItem(item.id, { title: e.target.value })} className="flex-1 min-w-0 px-2.5 py-1.5 text-xs font-semibold rounded-lg border bg-zinc-50 dark:bg-zinc-800/60 border-zinc-200 dark:border-zinc-700" placeholder="番外名称" />
                    </div>
                    <textarea value={item.content} onChange={(e) => updateExtraStoryImportPreviewItem(item.id, { content: e.target.value })} className="w-full min-h-[120px] px-3 py-2 text-xs leading-relaxed rounded-xl border bg-white dark:bg-zinc-900 border-zinc-200 dark:border-zinc-700 resize-y focus:outline-none" />
                  </div>
                );
              })}
              {extraStoryImportVisibleCount < extraStoryImportPreview.length && (
                <button onClick={() => setExtraStoryImportVisibleCount((n) => Math.min(n + 60, extraStoryImportPreview.length))} className="w-full py-2.5 text-xs rounded-xl border border-dashed text-zinc-500 hover:bg-zinc-50 dark:hover:bg-zinc-800">继续显示后面的 {Math.min(60, extraStoryImportPreview.length - extraStoryImportVisibleCount)} 条</button>
              )}
            </div>
            <div className="p-4 border-t border-zinc-200 dark:border-zinc-800 flex justify-end gap-2">
              <button onClick={() => { setExtraStoryImportPreview(null); setSelectedExtraStoryImportIds([]); }} className="px-4 py-2 text-xs rounded-lg border">取消</button>
              <button onClick={commitSelectedExtraStoryImports} disabled={!selectedExtraStoryImportIds.length} className="px-4 py-2 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 disabled:opacity-40">确定导入（{selectedExtraStoryImportIds.length}）</button>
            </div>
          </div>
        </div>
      )}

      {showJsonWorldBookImportPreview && pendingJsonWorldBook && (
        <div className="fixed inset-0 z-[88] flex items-center justify-center bg-black/60 p-4 pt-[calc(1rem+env(safe-area-inset-top,0px))]">
          <div className="w-full max-w-4xl max-h-[88vh] bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl overflow-hidden flex flex-col">
            <div className="p-4 border-b border-zinc-200 dark:border-zinc-800 flex items-center justify-between">
              <div><h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100">确认导入 JSON 世界书</h3><p className="text-[11px] text-zinc-400 mt-0.5">识别到 {pendingJsonWorldBook.entries?.length || 0} 个条目</p></div>
              <button onClick={() => { setPendingJsonWorldBook(null); setShowJsonWorldBookImportPreview(false); }} className="p-1 text-zinc-400"><X className="w-4 h-4" /></button>
            </div>
            <div className="p-5 overflow-y-auto grid grid-cols-1 md:grid-cols-[1fr_1.8fr] gap-4">
              <div className="space-y-3">
                <div className="text-xs font-bold">世界书名称</div><input value={pendingJsonWorldBook.title} onChange={(e) => setPendingJsonWorldBook((p: any) => ({ ...p, title: e.target.value }))} className="w-full px-3 py-2 text-xs border rounded-xl bg-zinc-50 dark:bg-zinc-800/60 border-zinc-200 dark:border-zinc-700" />
                <div className="text-xs font-bold">作者名</div><input value={pendingJsonWorldBook.author} onChange={(e) => setPendingJsonWorldBook((p: any) => ({ ...p, author: e.target.value }))} className="w-full px-3 py-2 text-xs border rounded-xl bg-zinc-50 dark:bg-zinc-800/60 border-zinc-200 dark:border-zinc-700" />
              </div>
              <div className="space-y-3"><div className="relative"><Search className="w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 text-zinc-400" /><input value={jsonWorldBookSearchQuery} onChange={(e) => setJsonWorldBookSearchQuery(e.target.value)} placeholder="搜索条目..." className="w-full pl-9 pr-4 py-1.5 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400 dark:focus:ring-zinc-500" /></div><div className="space-y-2">{(pendingJsonWorldBook.entries || []).map((entry: any, i: number) => { const label=entry.comment||entry.name||entry.title||`条目 #${i+1}`; const content=typeof entry.content==='string'?entry.content:JSON.stringify(entry.content||entry,null,2); const q=jsonWorldBookSearchQuery.trim().toLowerCase(); if(q&&!`${label} ${content}`.toLowerCase().includes(q)) return null; return <div key={i} className="p-3 border rounded-xl border-zinc-200 dark:border-zinc-800 bg-zinc-50 dark:bg-zinc-800/50"><div className="text-xs font-bold">{label}</div><div className="text-[11px] text-zinc-500 mt-1 whitespace-pre-wrap line-clamp-5">{content}</div></div>; })}</div></div>
            </div>
            <div className="p-4 border-t border-zinc-200 dark:border-zinc-800 flex justify-end gap-2"><button onClick={() => { setPendingJsonWorldBook(null); setShowJsonWorldBookImportPreview(false); }} className="px-4 py-2 text-xs rounded-lg border border-zinc-200 dark:border-zinc-700">取消</button><button onClick={commitJsonWorldBookImport} className="px-4 py-2 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900">确定导入</button></div>
          </div>
        </div>
      )}

      {editingQrItem && activeDetailCard && (
        <div className="fixed inset-0 z-[90] flex items-center justify-center bg-black/60 p-4 pt-[calc(1rem+env(safe-area-inset-top,0px))]">
          <div className="w-full max-w-lg bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl overflow-hidden">
            <div className="p-4 border-b border-zinc-200 dark:border-zinc-800 flex items-center justify-between">
              <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100">编辑 QR</h3>
              <button type="button" onClick={() => setEditingQrItem(null)} className="p-1 text-zinc-400 hover:text-zinc-700 dark:hover:text-zinc-200"><X className="w-4 h-4" /></button>
            </div>
            <div className="p-5 space-y-3">
              {[
                ['label', 'QR 名称 / Label'],
                ['title', '标题'],
                ['message', '消息'],
                ['automationId', 'Automation ID'],
              ].map(([key, label]) => (
                <div key={key}>
                  <label className="block text-xs font-semibold text-zinc-600 dark:text-zinc-300 mb-1">{label}</label>
                  <input
                    value={editingQrItem.item?.[key] || ''}
                    onChange={(e) => setEditingQrItem((prev) => prev ? { ...prev, item: { ...prev.item, [key]: e.target.value } } : prev)}
                    className="w-full px-3 py-2 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-xl text-zinc-900 dark:text-zinc-100 focus:outline-none"
                  />
                </div>
              ))}
              <div className="grid grid-cols-2 gap-3">
                {[
                  ['preventAutoExecute', '手动执行'],
                  ['isHidden', '隐藏'],
                  ['executeOnStartup', '启动时执行'],
                  ['executeOnUser', '用户输入时执行'],
                  ['executeOnAi', 'AI 时执行'],
                  ['executeOnChatChange', '聊天切换时执行'],
                ].map(([key, label]) => (
                  <label key={key} className="flex items-center gap-2 text-xs text-zinc-700 dark:text-zinc-300">
                    <input type="checkbox" checked={!!editingQrItem.item?.[key]} onChange={(e) => setEditingQrItem((prev) => prev ? { ...prev, item: { ...prev.item, [key]: e.target.checked } } : prev)} />
                    {label}
                  </label>
                ))}
              </div>
            </div>
            <div className="p-4 border-t border-zinc-200 dark:border-zinc-800 flex justify-end gap-2">
              <button type="button" onClick={() => setEditingQrItem(null)} className="px-4 py-2 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700">取消</button>
              <button type="button" onClick={saveEditedQrItem} className="px-4 py-2 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900">保存</button>
            </div>
          </div>
        </div>
      )}

      {/* ==================== WORLD BOOK DETAIL & EDIT MODAL (Single Page) ==================== */}
      {editingWorldBook && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 animate-in fade-in pt-[calc(1rem+env(safe-area-inset-top,0px))]">
          <div className="w-full max-w-2xl bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl flex flex-col overflow-hidden max-h-[90vh]">
            {/* Modal Header */}
            <div className="p-4 border-b border-zinc-200 dark:border-zinc-800 flex items-center justify-between flex-shrink-0">
              <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100 truncate">
                世界书详情 - {editingWorldBook.title}
              </h3>
              <button
                onClick={() => setEditingWorldBook(null)}
                className="p-1 rounded-lg text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-200"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            {/* Modal Body (Single Page) */}
            <div className="p-5 overflow-y-auto space-y-4 flex-1">
              {editingWorldBook.importFormat === 'json' ? (
                <div className="space-y-4">
                  <div className="grid grid-cols-1 md:grid-cols-[1fr_1.7fr] gap-4">
                    <div className="space-y-3">
                      <div>
                        <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">世界书名称</label>
                        <input value={editingWorldBook.title} onChange={(e) => setEditingWorldBook((p) => p ? { ...p, title: e.target.value } : null)} className="w-full px-3 py-2 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-xl" />
                      </div>
                      <div>
                        <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">作者名</label>
                        <input value={editingWorldBook.author || ''} onChange={(e) => setEditingWorldBook((p) => p ? { ...p, author: e.target.value } : null)} className="w-full px-3 py-2 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-xl" />
                      </div>
                      <div>
                        <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">分组</label>
                        <select value={editingWorldBook.category || '默认'} onChange={(e) => setEditingWorldBook((p) => p ? { ...p, category: e.target.value } : null)} className="w-full px-3 py-2 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-xl">
                          {Array.from(new Set(['默认', ...(appData.worldBookCategories || []), ...worldBooksList.map((p) => p.category || '默认')])).map((cat) => <option key={cat} value={cat}>{cat}</option>)}
                        </select>
                      </div>
                      <div className="rounded-xl border border-zinc-200 dark:border-zinc-800 p-3 bg-zinc-50 dark:bg-zinc-800/40">
                        <div className="text-xs font-bold text-zinc-700 dark:text-zinc-300 mb-1">导出格式</div>
                        <div className="text-[11px] text-zinc-500 dark:text-zinc-400">JSON</div>
                        <button type="button" onClick={() => exportWorldBookJson(editingWorldBook)} className="mt-2 w-full px-3 py-2 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 flex items-center justify-center gap-1.5"><Download className="w-3.5 h-3.5" /> 导出 JSON</button>
                      </div>
                    </div>
                    <div className="space-y-3 min-w-0">
                      <div className="text-xs font-bold text-zinc-800 dark:text-zinc-200">条目 ({editingWorldBook.entries?.length || 0})</div>
                      <div className="relative w-full">
                        <Search className="w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 text-zinc-400" />
                        <input value={jsonWorldBookSearchQuery} onChange={(e) => setJsonWorldBookSearchQuery(e.target.value)} placeholder="搜索条目名称、关键词或正文..." className="w-full pl-9 pr-4 py-1.5 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-400 dark:focus:ring-zinc-500" />
                      </div>
                      <div className="space-y-2 max-h-[58vh] overflow-y-auto pr-1">
                        {(editingWorldBook.entries || []).map((entry: any, i: number) => {
                          const label = entry.comment || entry.name || entry.title || `条目 #${i + 1}`;
                          const keys = Array.isArray(entry.keys || entry.key) ? (entry.keys || entry.key).join(', ') : (entry.keys || entry.key || '');
                          const content = typeof entry.content === 'string' ? entry.content : JSON.stringify(entry.content || entry, null, 2);
                          const q = jsonWorldBookSearchQuery.trim().toLowerCase();
                          if (q && !`${label} ${keys} ${content}`.toLowerCase().includes(q)) return null;
                          return (
                            <div key={i} className="p-3 rounded-xl border border-zinc-200 dark:border-zinc-800 bg-zinc-50 dark:bg-zinc-800/50">
                              <div className="flex items-center justify-between gap-2"><div className="text-xs font-bold text-zinc-900 dark:text-zinc-100 truncate">{label}</div><span className="text-[10px] text-zinc-400 flex-shrink-0">#{i + 1}</span></div>
                              {keys && <div className="text-[10px] text-zinc-400 mt-1 truncate">关键词：{keys}</div>}
                              <div className="text-[11px] text-zinc-600 dark:text-zinc-300 mt-2 whitespace-pre-wrap break-words line-clamp-6">{content || '（空）'}</div>
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  </div>
                </div>
              ) : (
              <>
              {/* 世界书名称 */}
              <div>
                <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                  世界书名称 <span className="text-rose-500">*</span>
                </label>
                <input
                  type="text"
                  value={editingWorldBook.title}
                  onChange={(e) => setEditingWorldBook((prev) => (prev ? { ...prev, title: e.target.value } : null))}
                  className="w-full px-3 py-2 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-xl text-zinc-900 dark:text-zinc-100 font-bold focus:outline-none"
                />
              </div>

              {/* 作者名 */}
              <div>
                <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                  作者名
                </label>
                <input
                  type="text"
                  value={editingWorldBook.author || ''}
                  onChange={(e) => setEditingWorldBook((prev) => (prev ? { ...prev, author: e.target.value } : null))}
                  placeholder="例如：原作者名或来源..."
                  className="w-full px-3 py-2 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-xl text-zinc-900 dark:text-zinc-100 focus:outline-none"
                />
              </div>

              {/* 所属分组 */}
              <div>
                <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1">
                  所属分组
                </label>
                <select
                  value={editingWorldBook.category || '默认'}
                  onChange={(e) => setEditingWorldBook((prev) => (prev ? { ...prev, category: e.target.value } : null))}
                  className="w-full px-3 py-2 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-xl text-zinc-900 dark:text-zinc-100 focus:outline-none"
                >
                  {Array.from(
                    new Set(['默认', ...(appData.worldBookCategories || []), ...worldBooksList.map((p) => p.category || '默认')])
                  ).map((cat) => (
                    <option key={cat} value={cat}>
                      {cat}
                    </option>
                  ))}
                </select>
              </div>

              {/* 世界书内容 (复制键与放大展开) */}
              <div>
                <div className="flex items-center justify-between mb-1.5">
                  <label className="block text-xs font-bold text-zinc-700 dark:text-zinc-300">
                    世界书内容
                  </label>
                  <div className="flex items-center gap-2">
                    <button
                      type="button"
                      onClick={() => {
                        navigator.clipboard.writeText(editingWorldBook.content || '');
                        showToast('世界书内容已复制到剪贴板！', 'success');
                      }}
                      className="px-2.5 py-1 text-[11px] font-semibold text-zinc-600 dark:text-zinc-300 bg-zinc-100 dark:bg-zinc-800 hover:bg-zinc-200 dark:hover:bg-zinc-700 rounded-lg flex items-center gap-1 transition-colors"
                    >
                      <Copy className="w-3 h-3" /> 复制
                    </button>
                    <button
                      type="button"
                      onClick={() => setIsWorldBookContentExpanded(!isWorldBookContentExpanded)}
                      className="px-2.5 py-1 text-[11px] font-semibold text-zinc-600 dark:text-zinc-300 bg-zinc-100 dark:bg-zinc-800 hover:bg-zinc-200 dark:hover:bg-zinc-700 rounded-lg flex items-center gap-1 transition-colors"
                    >
                      <Maximize2 className="w-3 h-3" /> {isWorldBookContentExpanded ? '折叠收起' : '放大展开'}
                    </button>
                  </div>
                </div>

                <textarea
                  value={editingWorldBook.content}
                  onChange={(e) => setEditingWorldBook((prev) => (prev ? { ...prev, content: e.target.value } : null))}
                  rows={isWorldBookContentExpanded ? 16 : 5}
                  placeholder="输入或修改世界书内容..."
                  className="w-full px-3 py-2 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-xl leading-relaxed text-zinc-900 dark:text-zinc-100 focus:outline-none font-mono"
                />
              </div>

              {/* 在世界书内容的下方加一个导出（docx，txt，json） */}
              <div className="pt-2 border-t border-zinc-200 dark:border-zinc-800 space-y-2">
                <label className="block text-xs font-bold text-zinc-700 dark:text-zinc-300">
                  导出世界书文档
                </label>
                <div className="grid grid-cols-3 gap-2">
                  <button
                    type="button"
                    onClick={() => exportWorldBookDocx(editingWorldBook)}
                    className="py-2 text-xs font-bold rounded-xl bg-zinc-100 dark:bg-zinc-800 hover:bg-zinc-200 dark:hover:bg-zinc-700 text-zinc-800 dark:text-zinc-200 transition-colors flex items-center justify-center gap-1.5"
                  >
                    <Download className="w-3.5 h-3.5" /> 导出 .docx
                  </button>
                  <button
                    type="button"
                    onClick={() => exportWorldBookTxt(editingWorldBook)}
                    className="py-2 text-xs font-bold rounded-xl bg-zinc-100 dark:bg-zinc-800 hover:bg-zinc-200 dark:hover:bg-zinc-700 text-zinc-800 dark:text-zinc-200 transition-colors flex items-center justify-center gap-1.5"
                  >
                    <Download className="w-3.5 h-3.5" /> 导出 .txt
                  </button>
                  <button
                    type="button"
                    onClick={() => exportWorldBookJson(editingWorldBook)}
                    className="py-2 text-xs font-bold rounded-xl bg-zinc-100 dark:bg-zinc-800 hover:bg-zinc-200 dark:hover:bg-zinc-700 text-zinc-800 dark:text-zinc-200 transition-colors flex items-center justify-center gap-1.5"
                  >
                    <Download className="w-3.5 h-3.5" /> 导出 .json
                  </button>
                </div>
              </div>
              </>
              )}
            </div>

            {/* Modal Bottom Buttons: 取消, 确定, 删除 (Requirement 4) */}
            <div className="flex items-center justify-between p-4 border-t border-zinc-200 dark:border-zinc-800 bg-zinc-50 dark:bg-zinc-900/50 flex-shrink-0">
              <button
                type="button"
                onClick={() => {
                  if (confirm(`确定要删除世界书 “${editingWorldBook.title}” 吗？`)) {
                    const updated = worldBooksList.filter((p) => p.id !== editingWorldBook.id);
                    updateAppData({ ...appData, worldBooks: updated });
                    setEditingWorldBook(null);
                    showToast('世界书已删除', 'info');
                  }
                }}
                className="px-4 py-2 text-xs font-bold rounded-lg bg-rose-50 text-rose-600 dark:bg-rose-950/60 dark:text-rose-400 hover:bg-rose-100 dark:hover:bg-rose-900/60 transition-colors flex items-center gap-1.5"
              >
                <Trash2 className="w-3.5 h-3.5" /> 删除
              </button>

              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={() => setEditingWorldBook(null)}
                  className="px-4 py-2 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800 transition-colors"
                >
                  取消
                </button>
                <button
                  type="button"
                  onClick={() => {
                    if (!editingWorldBook.title.trim()) {
                      showToast('世界书名称不能为空', 'error');
                      return;
                    }
                    const updated = worldBooksList.map((p) =>
                      p.id === editingWorldBook.id ? { ...editingWorldBook, updatedAt: Date.now() } : p
                    );
                    updateAppData({ ...appData, worldBooks: updated });
                    setEditingWorldBook(null);
                    showToast('保存成功！', 'success');
                  }}
                  className="px-4 py-2 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:opacity-90 transition-opacity"
                >
                  确定
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ==================== NEW WORLD BOOK GROUP MODAL ==================== */}
      {showNewWorldBookGroupModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 animate-in fade-in pt-[calc(1rem+env(safe-area-inset-top,0px))]">
          <div className="w-full max-w-sm bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl p-5 space-y-4">
            <div className="flex items-center justify-between border-b border-zinc-100 dark:border-zinc-800 pb-3">
              <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100">新建世界书分组</h3>
              <button
                onClick={() => setShowNewWorldBookGroupModal(false)}
                className="p-1 rounded-lg text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-200"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div>
              <label className="text-xs font-semibold text-zinc-700 dark:text-zinc-300 block mb-1">
                分组名称
              </label>
              <input
                type="text"
                value={newWorldBookGroupName}
                onChange={(e) => setNewWorldBookGroupName(e.target.value)}
                placeholder="请输入分组名字..."
                className="w-full px-3 py-2 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-xl text-zinc-900 dark:text-zinc-100 focus:outline-none"
              />
            </div>

            <div className="flex justify-end gap-2 pt-2 border-t border-zinc-100 dark:border-zinc-800">
              <button
                onClick={() => setShowNewWorldBookGroupModal(false)}
                className="px-4 py-2 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800"
              >
                取消
              </button>
              <button
                onClick={handleCreateNewWorldBookGroup}
                className="px-4 py-2 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:bg-zinc-800 dark:hover:bg-zinc-200"
              >
                确定新建
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================== MANAGE WORLD BOOK GROUP MODAL ==================== */}
      {managingWorldBookCategory && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 animate-in fade-in pt-[calc(1rem+env(safe-area-inset-top,0px))]">
          <div className="w-full max-w-sm bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl p-5 space-y-4">
            <div className="flex items-center justify-between border-b border-zinc-100 dark:border-zinc-800 pb-3">
              <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100">
                管理分组 - “{managingWorldBookCategory}”
              </h3>
              <button
                onClick={() => setManagingWorldBookCategory(null)}
                className="p-1 rounded-lg text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-200"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            {/* Rename Section */}
            <div className="space-y-2">
              <label className="text-xs font-semibold text-zinc-500 block">重命名分组</label>
              <div className="flex gap-2">
                <input
                  type="text"
                  value={renameWorldBookCategoryInput}
                  onChange={(e) => setRenameWorldBookCategoryInput(e.target.value)}
                  className="flex-1 px-3 py-1.5 text-xs bg-zinc-50 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 rounded-lg text-zinc-900 dark:text-zinc-100 focus:outline-none"
                />
                <button
                  onClick={handleRenameWorldBookCategory}
                  className="px-3 py-1.5 text-xs font-medium bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 rounded-lg hover:opacity-90"
                >
                  保存名字
                </button>
              </div>
            </div>

            {/* Delete Group Dialog with Circular Toggle (Requirement 1 & ST cards logic) */}
            <div className="pt-3 border-t border-zinc-200 dark:border-zinc-800 space-y-3">
              <div className="text-xs font-bold text-rose-600 dark:text-rose-400">删除分组</div>
              <p className="text-xs text-zinc-500 dark:text-zinc-400 leading-relaxed">
                是否要删除分组“{managingWorldBookCategory}”？删除前请选择是否一并清理组内世界书。
              </p>

              {/* Circular Selection Toggle */}
              <div
                onClick={() => setDeleteItemsWithWorldBookCategory((prev) => !prev)}
                className="flex items-center gap-3 p-3 bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-xl cursor-pointer hover:bg-zinc-100 dark:hover:bg-zinc-800 transition-colors"
              >
                <div
                  className={`w-5 h-5 rounded-full border-2 flex items-center justify-center transition-all ${
                    deleteItemsWithWorldBookCategory
                      ? 'border-zinc-900 dark:border-zinc-100 bg-white dark:bg-zinc-900'
                      : 'border-zinc-400 dark:border-zinc-600 bg-transparent'
                  }`}
                >
                  {deleteItemsWithWorldBookCategory && (
                    <div className="w-2.5 h-2.5 rounded-full bg-zinc-900 dark:bg-zinc-100" />
                  )}
                </div>

                <div className="text-xs font-medium text-zinc-800 dark:text-zinc-200">
                  是否要一并把分组里的角色卡给删掉？
                </div>
              </div>

              <div className="flex justify-end gap-2 pt-2">
                <button
                  onClick={() => setManagingWorldBookCategory(null)}
                  className="px-3 py-1.5 text-xs font-medium text-zinc-600 dark:text-zinc-400 hover:bg-zinc-100 dark:hover:bg-zinc-800 rounded-lg"
                >
                  取消
                </button>
                <button
                  onClick={handleConfirmDeleteWorldBookCategory}
                  className="px-4 py-1.5 text-xs font-bold bg-rose-600 text-white hover:bg-rose-700 rounded-lg"
                >
                  确认删除分组
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ==================== BATCH MOVE WORLD BOOKS MODAL ==================== */}
      {showWorldBookBatchMoveModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4 animate-in fade-in pt-[calc(1rem+env(safe-area-inset-top,0px))]">
          <div className="w-full max-w-sm bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl p-5 space-y-4">
            <div className="flex items-center justify-between border-b border-zinc-100 dark:border-zinc-800 pb-3">
              <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100">批量移动分组</h3>
              <button
                onClick={() => setShowWorldBookBatchMoveModal(false)}
                className="p-1 rounded-lg text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-200"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div>
              <label className="text-xs font-semibold text-zinc-700 dark:text-zinc-300 block mb-1">
                选择目标分组
              </label>
              <select
                value={batchTargetWorldBookCategory}
                onChange={(e) => setBatchTargetWorldBookCategory(e.target.value)}
                className="w-full px-3 py-2 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700/60 rounded-xl text-zinc-900 dark:text-zinc-100 focus:outline-none"
              >
                <option value="">-- 请选择分组 --</option>
                {Array.from(
                  new Set(['默认', ...(appData.worldBookCategories || []), ...worldBooksList.map((p) => p.category || '默认')])
                ).map((cat) => (
                  <option key={cat} value={cat}>
                    {cat}
                  </option>
                ))}
              </select>
            </div>

            <div className="flex justify-end gap-2 pt-2 border-t border-zinc-100 dark:border-zinc-800">
              <button
                onClick={() => setShowWorldBookBatchMoveModal(false)}
                className="px-4 py-2 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800"
              >
                取消
              </button>
              <button
                onClick={handleConfirmBatchMoveWorldBooks}
                className="px-4 py-2 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:bg-zinc-800 dark:hover:bg-zinc-200"
              >
                确定移动
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================== ST 角色卡：导入单条 / 全部选择 ==================== */}
      {cardSectionImportModal && activeDetailCard && (
        <div className="fixed inset-0 z-[88] flex items-center justify-center bg-black/60 p-4 pt-[calc(1rem+env(safe-area-inset-top,0px))] pb-[calc(1rem+env(safe-area-inset-bottom,0px))] animate-in fade-in">
          <div className="w-full max-w-sm bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl p-5 space-y-4">
            <div className="flex items-center justify-between">
              <div>
                <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100">导入{cardSectionImportModal === 'worldbook' ? '世界书' : cardSectionImportModal === 'regex' ? '正则' : 'QR'}</h3>
                <p className="text-[11px] text-zinc-400 mt-1">请选择导入方式</p>
              </div>
              <button type="button" onClick={() => setCardSectionImportModal(null)} className="p-1 text-zinc-400 hover:text-zinc-700 dark:hover:text-zinc-200"><X className="w-4 h-4" /></button>
            </div>
            <div className="grid grid-cols-2 gap-2">
              <button type="button" onClick={() => {
                setCardSectionImportMode('single');
                setCardSectionImportModal(null);
                if (cardSectionImportModal === 'worldbook') cardWorldBookImportFileInputRef.current?.click();
                else if (cardSectionImportModal === 'regex') cardRegexFileInputRef.current?.click();
                else cardQrFileInputRef.current?.click();
              }} className="p-3 text-xs font-bold rounded-xl border border-zinc-200 dark:border-zinc-700 hover:bg-zinc-100 dark:hover:bg-zinc-800">导入单条</button>
              <button type="button" onClick={() => {
                setCardSectionImportMode('all');
                setCardSectionImportModal(null);
                if (cardSectionImportModal === 'worldbook') cardWorldBookImportFileInputRef.current?.click();
                else if (cardSectionImportModal === 'regex') cardRegexFileInputRef.current?.click();
                else cardQrFileInputRef.current?.click();
              }} className="p-3 text-xs font-bold rounded-xl bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:opacity-90">导入全部</button>
            </div>
            <button type="button" onClick={() => setCardSectionImportModal(null)} className="w-full px-3 py-2 text-xs text-zinc-500 hover:text-zinc-800 dark:hover:text-zinc-200">取消</button>
          </div>
        </div>
      )}

      {/* ==================== ST 角色卡：编辑正则弹窗 ==================== */}
      {editingCardRegex && activeDetailCard && (
        <div className="fixed inset-0 z-[89] flex items-center justify-center bg-black/60 p-4 pt-[calc(1rem+env(safe-area-inset-top,0px))] pb-[calc(1rem+env(safe-area-inset-bottom,0px))] animate-in fade-in">
          <div className="w-full max-w-2xl bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl overflow-hidden">
            <div className="p-4 border-b border-zinc-200 dark:border-zinc-800 flex items-center justify-between">
              <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100">编辑正则</h3>
              <button type="button" onClick={() => setEditingCardRegex(null)} className="p-1 text-zinc-400 hover:text-zinc-700 dark:hover:text-zinc-200"><X className="w-4 h-4" /></button>
            </div>
            <div className="p-5 space-y-4">
              <div><label className="text-xs font-semibold text-zinc-700 dark:text-zinc-300 block mb-1.5">正则名</label><input autoFocus value={editingCardRegex.scriptName} onChange={(e) => setEditingCardRegex((p) => p ? { ...p, scriptName: e.target.value } : p)} className="w-full px-3 py-2 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-xl text-zinc-900 dark:text-zinc-100 focus:outline-none" /></div>
              <div><label className="text-xs font-semibold text-amber-600 dark:text-amber-400 block mb-1.5">查找</label><textarea rows={5} value={editingCardRegex.findRegex} onChange={(e) => setEditingCardRegex((p) => p ? { ...p, findRegex: e.target.value } : p)} className="w-full p-3 text-xs font-mono bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-xl text-zinc-900 dark:text-zinc-100 focus:outline-none resize-y" /></div>
              <div><label className="text-xs font-semibold text-emerald-600 dark:text-emerald-400 block mb-1.5">替换</label><textarea rows={6} value={editingCardRegex.replaceString} onChange={(e) => setEditingCardRegex((p) => p ? { ...p, replaceString: e.target.value } : p)} className="w-full p-3 text-xs font-mono bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-xl text-zinc-900 dark:text-zinc-100 focus:outline-none resize-y" /></div>
            </div>
            <div className="px-5 py-3 border-t border-zinc-200 dark:border-zinc-800 flex justify-end gap-2 bg-zinc-50/70 dark:bg-zinc-900/70">
              <button type="button" onClick={() => setEditingCardRegex(null)} className="px-4 py-2 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700">取消</button>
              <button type="button" onClick={() => {
                const list = JSON.parse(JSON.stringify(getCardRegex(activeDetailCard) || []));
                const idx = editingCardRegex.index;
                if (!list[idx]) list[idx] = {};
                list[idx] = { ...list[idx], scriptName: editingCardRegex.scriptName, script_name: editingCardRegex.scriptName, name: editingCardRegex.scriptName, findRegex: editingCardRegex.findRegex, find_regex: editingCardRegex.findRegex, pattern: editingCardRegex.findRegex, replaceString: editingCardRegex.replaceString, replace_string: editingCardRegex.replaceString, replacement: editingCardRegex.replaceString };
                saveCardRegexList(list);
                setEditingCardRegex(null);
                showToast('正则已保存', 'success');
              }} className="px-5 py-2 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900">确认</button>
            </div>
          </div>
        </div>
      )}

      {/* ==================== ST 角色卡：添加备用开场白弹窗 ==================== */}
      {showAddAltGreetingModal && activeDetailCard && (
        <div className="fixed inset-0 z-[85] flex items-center justify-center bg-black/60 p-4 animate-in fade-in">
          <div className="w-full max-w-2xl bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl flex flex-col overflow-hidden">
            <div className="p-4 border-b border-zinc-200 dark:border-zinc-800 flex items-center justify-between">
              <div>
                <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100">添加备用开场白</h3>
                <p className="text-[11px] text-zinc-400 mt-0.5">确认后会按当前顺序追加到最后，不会改变已有开场白顺序。</p>
              </div>
              <button
                type="button"
                onClick={() => setShowAddAltGreetingModal(false)}
                className="p-1 rounded-lg text-zinc-400 hover:text-zinc-700 dark:hover:text-zinc-200"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="p-5">
              <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1.5">
                开场白内容
              </label>
              <textarea
                autoFocus
                value={newAltGreetingInputText}
                onChange={(e) => setNewAltGreetingInputText(e.target.value)}
                placeholder="请输入备用开场白内容，可填写长文本..."
                className="w-full min-h-[280px] p-3 text-xs leading-relaxed bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-xl focus:outline-none focus:ring-1 focus:ring-zinc-400 text-zinc-900 dark:text-zinc-100 resize-y"
              />
            </div>

            <div className="px-5 py-3 border-t border-zinc-200 dark:border-zinc-800 flex items-center justify-end gap-2 bg-zinc-50/70 dark:bg-zinc-900/70">
              <button
                type="button"
                onClick={() => {
                  setShowAddAltGreetingModal(false);
                  setNewAltGreetingInputText('');
                }}
                className="px-4 py-2 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800"
              >
                取消
              </button>
              <button
                type="button"
                onClick={() => {
                  const value = newAltGreetingInputText;
                  if (!value.trim()) {
                    showToast('请填写开场白内容', 'error');
                    return;
                  }
                  const currentAlts = getCardAlternateGreetings(activeDetailCard);
                  const updatedCards = appData.cards.map((c) =>
                    c.id === activeDetailCard.id
                      ? {
                          ...c,
                          editHistory: {
                            ...c.editHistory,
                            alternate_greetings: [...currentAlts, value],
                          },
                          edited: true,
                        }
                      : c
                  );
                  updateAppData({ ...appData, cards: updatedCards });
                  setShowAddAltGreetingModal(false);
                  setNewAltGreetingInputText('');
                  showToast('备用开场白已添加', 'success');
                }}
                className="px-5 py-2 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:opacity-90"
              >
                确认
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================== ST 角色卡：添加世界书条目弹窗 ==================== */}
      {showAddCardWorldBookEntryModal && activeDetailCard && (
        <div className="fixed inset-0 z-[85] flex items-center justify-center bg-black/60 p-4 animate-in fade-in">
          <div className="w-full max-w-2xl bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl flex flex-col overflow-hidden">
            <div className="p-4 border-b border-zinc-200 dark:border-zinc-800 flex items-center justify-between">
              <div>
                <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100">添加世界书条目</h3>
                <p className="text-[11px] text-zinc-400 mt-0.5">确认后按当前顺序追加到最后，方便世界书较多时集中填写。</p>
              </div>
              <button
                type="button"
                onClick={() => setShowAddCardWorldBookEntryModal(false)}
                className="p-1 rounded-lg text-zinc-400 hover:text-zinc-700 dark:hover:text-zinc-200"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="p-5 space-y-4">
              <div>
                <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1.5">条目标题 / 备注</label>
                <input
                  autoFocus
                  type="text"
                  value={newCardWorldBookEntryForm.comment}
                  onChange={(e) => setNewCardWorldBookEntryForm((p) => ({ ...p, comment: e.target.value }))}
                  placeholder="例如：人物背景、世界规则..."
                  className="w-full px-3 py-2 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-xl text-zinc-900 dark:text-zinc-100 focus:outline-none"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1.5">触发关键词</label>
                <input
                  type="text"
                  value={newCardWorldBookEntryForm.keys}
                  onChange={(e) => setNewCardWorldBookEntryForm((p) => ({ ...p, keys: e.target.value }))}
                  placeholder="关键词1, 关键词2..."
                  className="w-full px-3 py-2 text-xs bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-xl text-zinc-900 dark:text-zinc-100 focus:outline-none"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-zinc-700 dark:text-zinc-300 mb-1.5">条目正文</label>
                <textarea
                  rows={12}
                  value={newCardWorldBookEntryForm.content}
                  onChange={(e) => setNewCardWorldBookEntryForm((p) => ({ ...p, content: e.target.value }))}
                  placeholder="请输入世界书条目正文，可填写长文本..."
                  className="w-full p-3 text-xs leading-relaxed bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-xl focus:outline-none focus:ring-1 focus:ring-zinc-400 text-zinc-900 dark:text-zinc-100 resize-y"
                />
              </div>
            </div>

            <div className="px-5 py-3 border-t border-zinc-200 dark:border-zinc-800 flex items-center justify-end gap-2 bg-zinc-50/70 dark:bg-zinc-900/70">
              <button
                type="button"
                onClick={() => {
                  setShowAddCardWorldBookEntryModal(false);
                  setNewCardWorldBookEntryForm({ comment: '', keys: '', content: '' });
                }}
                className="px-4 py-2 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800"
              >
                取消
              </button>
              <button
                type="button"
                onClick={handleConfirmAddCardWorldBookEntry}
                className="px-5 py-2 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:opacity-90"
              >
                确认
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================== ST 角色卡：QR JSON 导入弹窗 ==================== */}
      {showQrImportModal && activeDetailCard && (
        <div className="fixed inset-0 z-[90] flex items-center justify-center bg-black/60 p-4 animate-in fade-in">
          <div className="w-full max-w-2xl bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl flex flex-col overflow-hidden">
            <div className="p-4 border-b border-zinc-200 dark:border-zinc-800 flex items-center justify-between">
              <div>
                <h3 className="text-sm font-bold text-zinc-900 dark:text-zinc-100">导入 QR JSON</h3>
                <p className="text-[11px] text-zinc-400 mt-0.5">支持 version 2、包含 qrList 数组的 JSON 文件，也可以直接粘贴 JSON。</p>
              </div>
              <button type="button" onClick={() => setShowQrImportModal(false)} className="p-1 rounded-lg text-zinc-400 hover:text-zinc-700 dark:hover:text-zinc-200">
                <X className="w-4 h-4" />
              </button>
            </div>
            <div className="p-5">
              <textarea
                autoFocus
                value={qrInputText}
                onChange={(e) => setQrInputText(e.target.value)}
                placeholder={'在这里粘贴 QR JSON，例如：{\n  "version": 2,\n  "name": "新同层手机喵的qr",\n  "qrList": [...]\n}'}
                className="w-full min-h-[360px] p-3 text-xs leading-relaxed font-mono bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 rounded-xl focus:outline-none focus:ring-1 focus:ring-zinc-400 text-zinc-900 dark:text-zinc-100 resize-y"
              />
            </div>
            <div className="px-5 py-3 border-t border-zinc-200 dark:border-zinc-800 flex items-center justify-between gap-2 bg-zinc-50/70 dark:bg-zinc-900/70">
              <button
                type="button"
                onClick={() => qrFileInputRef.current?.click()}
                className="px-3 py-2 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800 flex items-center gap-1"
              >
                <Upload className="w-3.5 h-3.5" /> 选择 JSON 文件
              </button>
              <div className="flex items-center gap-2">
                <button type="button" onClick={() => { setShowQrImportModal(false); setQrInputText(''); }} className="px-4 py-2 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 text-zinc-700 dark:text-zinc-300 hover:bg-zinc-100 dark:hover:bg-zinc-800">取消</button>
                <button type="button" disabled={!qrInputText.trim()} onClick={() => importQrJsonText(qrInputText)} className="px-5 py-2 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:opacity-90 disabled:opacity-40">确认导入</button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ==================== GACHA RANDOM CARD MODAL (Requirement 9) ==================== */}
      {currentPage === 'st-cards' && gachaCard && (
        <div className="fixed inset-0 z-[80] flex items-center justify-center bg-black/75 p-4 animate-in fade-in backdrop-blur-sm pt-[calc(1rem+env(safe-area-inset-top,0px))]">
          <div className="w-full max-w-sm bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl overflow-hidden animate-in zoom-in-95 duration-200">
            <div className="p-4 bg-white dark:bg-zinc-900 text-zinc-900 dark:text-zinc-100 border-b border-zinc-100 dark:border-zinc-800 flex items-center justify-between">
              <div className="flex items-center gap-2 font-bold text-sm">
                命运抽卡 - 随机选中!
              </div>
              <button
                onClick={() => setGachaCard(null)}
                className="p-1 rounded-lg text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-200 transition-colors"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="p-5 text-center space-y-4">
              <div className="aspect-[2/3] w-36 mx-auto rounded-xl overflow-hidden bg-zinc-100 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 shadow-md">
                {gachaCard.coverImage ? (
                  <img
                    src={gachaCard.coverImage}
                    alt={getCardDisplayName(gachaCard)}
                    className="w-full h-full object-cover"
                  />
                ) : (
                  <div className="w-full h-full flex items-center justify-center text-zinc-400">
                    <ImageIcon className="w-8 h-8" />
                  </div>
                )}
              </div>

              <div>
                <h3 className="font-bold text-base text-zinc-900 dark:text-zinc-100">
                  {getCardDisplayName(gachaCard)}
                </h3>
                <p className="text-xs text-zinc-500 dark:text-zinc-400 mt-0.5">
                  作者: {getCardCreator(gachaCard) || '未知作者'}
                </p>
                {getCardDescription(gachaCard) && (
                  <p className="text-xs text-zinc-600 dark:text-zinc-300 mt-2 line-clamp-3 bg-zinc-50 dark:bg-zinc-800/60 p-2.5 rounded-lg border border-zinc-200/60 dark:border-zinc-800/60 leading-relaxed text-left">
                    {getCardDescription(gachaCard)}
                  </p>
                )}
              </div>

              {(gachaCard.associations || []).length > 0 && (
                <div className="text-left border-t border-zinc-200 dark:border-zinc-800 pt-3 space-y-2">
                  <div className="text-xs font-bold text-zinc-700 dark:text-zinc-300">已关联角色卡</div>
                  <div className="flex flex-wrap gap-1.5">
                    {(gachaCard.associations || []).map((a) => {
                      const related = appData.cards.find((c) => c.id === a.cardId);
                      if (!related) return null;
                      return (
                        <button key={a.cardId} type="button" onClick={() => openAssociatedCard(related.id)} className="px-2.5 py-1.5 text-[11px] rounded-lg border border-zinc-200 dark:border-zinc-700 hover:bg-zinc-100 dark:hover:bg-zinc-800 text-zinc-700 dark:text-zinc-300">
                          {getCardDisplayName(related)} · 跳转
                        </button>
                      );
                    })}
                  </div>
                </div>
              )}

              <div className="flex gap-2 pt-2">
                <button
                  onClick={handleDrawRandomCard}
                  className="flex-1 py-2 text-xs font-bold rounded-xl border border-zinc-300 dark:border-zinc-700 hover:bg-zinc-100 dark:hover:bg-zinc-800 text-zinc-800 dark:text-zinc-200 flex items-center justify-center gap-1.5"
                >
                  再抽一次
                </button>
                <button
                  onClick={() => {
                    setDetailCardId(gachaCard.id);
                    setDetailTab('overview');
                    setGachaCard(null);
                  }}
                  className="flex-1 py-2 text-xs font-bold rounded-xl bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 hover:opacity-90 flex items-center justify-center gap-1.5"
                >
                  查看详情
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ==================== CHAT MEME MODALS ==================== */}
      {showChatAddChoiceModal && (
        <div className="fixed inset-0 z-[90] flex items-center justify-center bg-black/50 p-4 pb-[calc(1rem+env(safe-area-inset-bottom,0px))]">
          <div className="w-full max-w-sm bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl p-4 space-y-3">
            <div className="flex items-center justify-between"><h3 className="text-sm font-bold">新增聊天梗</h3><button onClick={() => setShowChatAddChoiceModal(false)} className="p-1 rounded-lg text-zinc-400 hover:bg-zinc-100 dark:hover:bg-zinc-800"><X className="w-4 h-4" /></button></div>
            <button onClick={() => { setShowChatAddChoiceModal(false); setShowChatManualModal(true); setChatManualInput(''); setChatImportPreview(null); }} className="w-full p-3 text-left rounded-xl border border-zinc-200 dark:border-zinc-800 hover:bg-zinc-50 dark:hover:bg-zinc-800/60"><div className="text-xs font-bold">手动录入</div><div className="text-[11px] text-zinc-400 mt-1">一个或多个聊天梗，梗与梗之间用空行分隔</div></button>
            <label className="w-full p-3 text-left rounded-xl border border-zinc-200 dark:border-zinc-800 hover:bg-zinc-50 dark:hover:bg-zinc-800/60 cursor-pointer block">
              <div className="text-xs font-bold">上传文档</div>
              <div className="text-[11px] text-zinc-400 mt-1">支持 DOCX / TXT，并自动按空行拆分多个聊天梗</div>
              <input ref={chatMemeFileInputRef} type="file" accept=".docx,.txt,text/plain,application/vnd.openxmlformats-officedocument.wordprocessingml.document" className="sr-only" onChange={handleChatMemeFileSelect} />
            </label>
          </div>
        </div>
      )}

      {showChatManualModal && (
        <div className="fixed inset-0 z-[95] flex items-center justify-center bg-black/50 p-4 pt-[calc(1rem+env(safe-area-inset-top,0px))] pb-[calc(1rem+env(safe-area-inset-bottom,0px))]">
          <div className="w-full max-w-sm max-h-[82vh] bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl flex flex-col overflow-hidden">
            <div className="p-4 border-b border-zinc-200 dark:border-zinc-800 flex items-center justify-between"><h3 className="text-sm font-bold">手动录入聊天梗</h3><button onClick={() => setShowChatManualModal(false)} className="p-1 rounded-lg text-zinc-400"><X className="w-4 h-4" /></button></div>
            <div className="p-4 overflow-y-auto space-y-3">
              <label className="text-xs font-semibold text-zinc-500">聊天梗内容</label>
              <textarea value={chatManualInput} onChange={(e) => { const value = e.target.value; setChatManualInput(value); const parts = splitChatMemeText(value); { const previewEntries = parts.length ? makeChatMemeEntries(parts, chatMemeCategoryFilter === '全部分组' ? '默认' : chatMemeCategoryFilter) : []; setChatImportPreview(previewEntries.length ? previewEntries : null); setSelectedChatImportIds(previewEntries.map((m) => m.id)); setChatImportSourceName('手动录入'); } }} rows={10} placeholder="梗与梗之间留一个空行……" className="w-full px-3 py-2 text-xs leading-relaxed bg-zinc-50 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 rounded-xl resize-none focus:outline-none focus:ring-1 focus:ring-zinc-400" />
              {chatImportPreview && <div className="space-y-2"><div className="text-[11px] font-semibold text-zinc-500">预览全部内容（{chatImportPreview.length} 条）</div>{chatImportPreview.map((m, i) => <div key={m.id} className="p-3 rounded-xl bg-zinc-50 dark:bg-zinc-800/60 border border-zinc-200 dark:border-zinc-700 text-xs leading-relaxed whitespace-pre-wrap"><span className="text-[10px] text-zinc-400 block mb-1">聊天梗 {i + 1}</span>{m.content}</div>)}</div>}
            </div>
            <div className="p-4 border-t border-zinc-200 dark:border-zinc-800 flex justify-end gap-2"><button onClick={() => { setShowChatManualModal(false); setChatImportPreview(null); }} className="px-3.5 py-1.5 text-xs rounded-lg border border-zinc-200 dark:border-zinc-700">取消</button><button onClick={() => { if (!chatImportPreview) handleChatManualPreview(); else commitChatMemeImport(); }} disabled={!chatManualInput.trim()} className="px-4 py-1.5 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 disabled:opacity-40">确定</button></div>
          </div>
        </div>
      )}

      {chatImportPreview && !showChatManualModal && (
        <div className="fixed inset-0 z-[96] flex items-center justify-center bg-black/50 p-4 pt-[calc(1rem+env(safe-area-inset-top,0px))] pb-[calc(1rem+env(safe-area-inset-bottom,0px))]">
          <div className="w-full max-w-sm max-h-[82vh] bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl flex flex-col overflow-hidden">
            <div className="p-4 border-b border-zinc-200 dark:border-zinc-800"><h3 className="text-sm font-bold">预览聊天梗</h3><p className="text-[11px] text-zinc-400 mt-1 truncate">{chatImportSourceName} · 共 {chatImportPreview.length} 条</p></div>
            <div className="p-4 border-b border-zinc-200 dark:border-zinc-800 space-y-2">
              <div className="flex flex-wrap items-center gap-2 text-[11px] text-zinc-500">
                <button onClick={() => setSelectedChatImportIds(selectedChatImportIds.length === chatImportPreview.length ? [] : chatImportPreview.map((m) => m.id))} className="px-2.5 py-1 rounded-lg border border-zinc-200 dark:border-zinc-700">{selectedChatImportIds.length === chatImportPreview.length ? '取消全选' : '全选'}</button>
                <button onClick={mergeSelectedChatImports} disabled={selectedChatImportIds.length < 2} className="px-2.5 py-1 rounded-lg border border-zinc-200 dark:border-zinc-700 disabled:opacity-40">融合所选 ({selectedChatImportIds.length})</button>
                <span>可自由选择导入条目，也可以编辑内容后再导入</span>
              </div>
            </div>
            <div className="p-4 overflow-y-auto space-y-2">
              {chatImportPreview.map((m, i) => {
                const checked = selectedChatImportIds.includes(m.id);
                return <div key={m.id} className={`p-3 rounded-xl border text-xs leading-relaxed ${checked ? 'border-zinc-900 dark:border-zinc-100 ring-1 ring-zinc-900/10' : 'border-zinc-200 dark:border-zinc-700'} bg-zinc-50 dark:bg-zinc-800/60`}>
                  <div className="flex items-center gap-2 mb-2">
                    <button onClick={() => toggleChatImportSelection(m.id)} className={`w-5 h-5 rounded-full border-2 flex items-center justify-center flex-shrink-0 ${checked ? 'bg-zinc-900 dark:bg-zinc-100 border-zinc-900 dark:border-zinc-100' : 'border-zinc-300 dark:border-zinc-600'}`}>{checked && <Check className="w-3 h-3 text-white dark:text-zinc-900" />}</button>
                    <span className="text-[10px] text-zinc-400">聊天梗 {i + 1}</span>
                  </div>
                  <textarea value={m.content} onChange={(e) => updateChatImportPreviewItem(m.id, e.target.value)} rows={Math.min(10, Math.max(3, m.content.split('\n').length + 1))} className="w-full px-3 py-2 bg-white/70 dark:bg-zinc-900/60 border border-zinc-200 dark:border-zinc-700 rounded-lg resize-y focus:outline-none focus:ring-1 focus:ring-zinc-400 leading-relaxed" />
                </div>;
              })}
            </div>
            <div className="p-4 border-t border-zinc-200 dark:border-zinc-800 flex justify-end gap-2"><button onClick={() => { setChatImportPreview(null); setSelectedChatImportIds([]); setChatImportSourceName(''); }} className="px-3.5 py-1.5 text-xs rounded-lg border border-zinc-200 dark:border-zinc-700">取消</button><button onClick={commitChatMemeImport} disabled={!chatImportPreview.some((m) => selectedChatImportIds.includes(m.id))} className="px-4 py-1.5 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 disabled:opacity-40">确定导入 ({selectedChatImportIds.length})</button></div>
          </div>
        </div>
      )}

      {editingChatMeme && (
        <div className="fixed inset-0 z-[95] flex items-center justify-center bg-black/50 p-4 pt-[calc(1rem+env(safe-area-inset-top,0px))] pb-[calc(1rem+env(safe-area-inset-bottom,0px))]">
          <div className="w-full max-w-4xl h-[85vh] min-h-[550px] max-h-[750px] bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl flex flex-col overflow-hidden">
            <div className="p-4 border-b border-zinc-200 dark:border-zinc-800 flex items-center justify-between"><h3 className="text-sm font-bold">聊天梗</h3><button onClick={() => setEditingChatMeme(null)} className="p-1.5 rounded-lg text-zinc-400 hover:bg-zinc-100 dark:hover:bg-zinc-800"><X className="w-4 h-4" /></button></div>
            <div className="p-5 flex-1 overflow-y-auto"><label className="text-xs font-semibold text-zinc-500 block mb-2">聊天梗内容</label><textarea value={editingChatMeme.content} onChange={(e) => setEditingChatMeme({ ...editingChatMeme, content: e.target.value })} className="w-full h-full min-h-[380px] px-4 py-3 text-sm leading-relaxed bg-zinc-50 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 rounded-xl resize-none focus:outline-none focus:ring-1 focus:ring-zinc-400" /></div>
            <div className="p-4 border-t border-zinc-200 dark:border-zinc-800 flex items-center justify-between gap-2"><button onClick={() => navigator.clipboard?.writeText(editingChatMeme.content).then(() => showToast('聊天梗已复制', 'success')).catch(() => showToast('复制失败，请手动复制', 'error'))} className="px-3.5 py-1.5 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 flex items-center gap-1.5"><Copy className="w-3.5 h-3.5" />复制</button><div className="flex gap-2"><button onClick={() => setEditingChatMeme(null)} className="px-3.5 py-1.5 text-xs rounded-lg border border-zinc-200 dark:border-zinc-700">取消</button><button onClick={handleSaveChatMemeEdit} className="px-4 py-1.5 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900">确定</button><button onClick={handleDeleteChatMeme} className="px-4 py-1.5 text-xs font-bold rounded-lg bg-rose-600 text-white">删除</button></div></div>
          </div>
        </div>
      )}

      {showNewChatMemeGroupModal && (
        <div className="fixed inset-0 z-[95] flex items-center justify-center bg-black/50 p-4"><div className="w-full max-w-sm bg-white dark:bg-zinc-900 rounded-2xl shadow-2xl p-5 space-y-4"><h3 className="text-sm font-bold">新建聊天梗分组</h3><input autoFocus value={newChatMemeGroupName} onChange={(e) => setNewChatMemeGroupName(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && handleCreateChatMemeGroup()} placeholder="请输入分组名字" className="w-full px-3 py-2 text-sm bg-zinc-50 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 rounded-lg" /><div className="flex justify-end gap-2"><button onClick={() => setShowNewChatMemeGroupModal(false)} className="px-3 py-1.5 text-xs rounded-lg border border-zinc-200 dark:border-zinc-700">取消</button><button onClick={handleCreateChatMemeGroup} className="px-4 py-1.5 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900">确定</button></div></div></div>
      )}

      {managingChatMemeGroup && (
        <div className="fixed inset-0 z-[95] flex items-center justify-center bg-black/50 p-4"><div className="w-full max-w-sm bg-white dark:bg-zinc-900 rounded-2xl shadow-2xl p-5 space-y-4"><h3 className="text-sm font-bold">管理分组：{managingChatMemeGroup}</h3><div className="flex gap-2"><input value={renameChatMemeGroupInput} onChange={(e) => setRenameChatMemeGroupInput(e.target.value)} className="flex-1 px-3 py-1.5 text-xs bg-zinc-50 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 rounded-lg" /><button onClick={handleRenameChatMemeGroup} className="px-3 py-1.5 text-xs rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900">重命名</button></div><div className="pt-3 border-t border-zinc-200 dark:border-zinc-800"><p className="text-xs font-bold text-rose-600 mb-2">删除分组</p><div onClick={() => setDeleteCardsWithChatMemeGroup((v) => !v)} className="flex items-center gap-3 p-3 rounded-xl bg-zinc-50 dark:bg-zinc-800/60 cursor-pointer"><div className={`w-5 h-5 rounded-full border-2 flex items-center justify-center ${deleteCardsWithChatMemeGroup ? 'border-zinc-900 dark:border-zinc-100' : 'border-zinc-400'}`}>{deleteCardsWithChatMemeGroup && <div className="w-2.5 h-2.5 rounded-full bg-zinc-900 dark:bg-zinc-100" />}</div><span className="text-xs">是否一并删除该分组里的角色卡？</span></div></div><div className="flex justify-end gap-2"><button onClick={() => { setManagingChatMemeGroup(null); setDeleteCardsWithChatMemeGroup(false); }} className="px-3 py-1.5 text-xs rounded-lg border border-zinc-200 dark:border-zinc-700">取消</button><button onClick={handleDeleteChatMemeGroup} className="px-4 py-1.5 text-xs font-bold rounded-lg bg-rose-600 text-white">确认删除分组</button></div></div></div>
      )}

      {showChatExportModal && (
        <div className="fixed inset-0 z-[95] flex items-center justify-center bg-black/50 p-4"><div className="w-full max-w-sm bg-white dark:bg-zinc-900 rounded-2xl shadow-2xl p-5 space-y-4"><h3 className="text-sm font-bold">导出聊天梗（{selectedChatMemeIds.length} 条）</h3><div className="grid grid-cols-3 gap-2">{(['docx','txt','json'] as const).map((f) => <button key={f} onClick={() => setChatMemeExportFormat(f)} className={`py-2 rounded-lg text-xs border ${chatMemeExportFormat === f ? 'bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 border-zinc-900' : 'border-zinc-200 dark:border-zinc-700'}`}>{f.toUpperCase()}</button>)}</div><div className="flex justify-end gap-2"><button onClick={() => setShowChatExportModal(false)} className="px-3 py-1.5 text-xs rounded-lg border border-zinc-200 dark:border-zinc-700">取消</button><button onClick={exportSelectedChatMemes} className="px-4 py-1.5 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900">导出</button></div></div></div>
      )}

      {/* ==================== TOAST NOTIFICATIONS ==================== */}
      <div className="fixed bottom-6 left-1/2 -translate-x-1/2 z-[100] flex flex-col-reverse items-center gap-2 pointer-events-none toast-container pb-[env(safe-area-inset-bottom,0px)]">
        {toasts.map((t) => (
          <div
            key={t.id}
            className={`px-4 py-2.5 rounded-xl shadow-2xl text-xs font-medium max-w-sm sm:max-w-md pointer-events-auto border animate-in slide-in-from-bottom duration-200 text-center ${
              t.type === 'success'
                ? 'bg-emerald-950/95 text-emerald-200 border-emerald-800 backdrop-blur-md'
                : t.type === 'error'
                ? 'bg-rose-950/95 text-rose-200 border-rose-800 backdrop-blur-md'
                : 'bg-zinc-900/95 text-zinc-100 border-zinc-800 backdrop-blur-md'
            }`}
          >
            {t.msg}
          </div>
        ))}
      </div>
      {/* ==================== MOBILE-FRIENDLY BATCH UPLOAD QUEUE ==================== */}
      {batchUploadKind && (
        <div className="fixed inset-0 z-[100] flex items-center justify-center bg-black/60 p-4 pt-[calc(1rem+env(safe-area-inset-top,0px))] pb-[calc(1rem+env(safe-area-inset-bottom,0px))]">
          <div className="w-full max-w-lg max-h-[85vh] overflow-hidden bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-2xl shadow-2xl flex flex-col">
            <div className="p-4 border-b border-zinc-100 dark:border-zinc-800 flex items-center justify-between">
              <div>
                <h3 className="text-sm font-bold">批量上传</h3>
                <p className="text-[11px] text-zinc-400 mt-1">Android 如果系统一次只能选择一个文件，可以反复点击“继续添加”，最后一次性导入。</p>
              </div>
              <button onClick={() => { setBatchUploadKind(null); setBatchUploadFiles([]); }} className="p-1.5 rounded-lg text-zinc-400 hover:bg-zinc-100 dark:hover:bg-zinc-800"><X className="w-4 h-4" /></button>
            </div>
            <div className="p-4 flex-1 overflow-y-auto space-y-2">
              {batchUploadFiles.length === 0 ? (
                <div className="py-10 text-center text-xs text-zinc-400 border border-dashed border-zinc-300 dark:border-zinc-700 rounded-xl">
                  还没有选择文件
                </div>
              ) : batchUploadFiles.map((file, index) => (
                <div key={`${file.name}-${file.size}-${file.lastModified}-${index}`} className="flex items-center gap-3 p-2.5 bg-zinc-50 dark:bg-zinc-800/60 rounded-lg border border-zinc-200 dark:border-zinc-700">
                  <FileText className="w-4 h-4 text-zinc-400 flex-shrink-0" />
                  <div className="min-w-0 flex-1"><div className="text-xs font-medium truncate">{file.name}</div><div className="text-[10px] text-zinc-400">{(file.size / 1024 / 1024).toFixed(2)} MB</div></div>
                  <button onClick={() => removeBatchUploadFile(index)} className="p-1.5 text-zinc-400 hover:text-rose-500"><Trash2 className="w-3.5 h-3.5" /></button>
                </div>
              ))}
            </div>
            <div className="p-4 border-t border-zinc-100 dark:border-zinc-800 flex flex-wrap items-center justify-between gap-2">
              <input ref={batchUploadFileInputRef} type="file" multiple accept={batchUploadAccept[batchUploadKind]} onChange={handleBatchUploadFileSelect} className="hidden" />
              <button onClick={() => batchUploadFileInputRef.current?.click()} className="px-3.5 py-2 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700 hover:bg-zinc-100 dark:hover:bg-zinc-800">+ 继续添加</button>
              <div className="flex gap-2 ml-auto">
                <button onClick={() => { setBatchUploadKind(null); setBatchUploadFiles([]); }} className="px-3.5 py-2 text-xs font-medium rounded-lg border border-zinc-200 dark:border-zinc-700">取消</button>
                <button disabled={!batchUploadFiles.length} onClick={confirmBatchUpload} className="px-4 py-2 text-xs font-bold rounded-lg bg-zinc-900 text-white dark:bg-zinc-100 dark:text-zinc-900 disabled:opacity-40">开始导入 ({batchUploadFiles.length})</button>
              </div>
            </div>
          </div>
        </div>
      )}

    </div>
  );
}
