import type { DictResult } from '@/lib/dict-types';
import crypto from 'crypto';

// ─── Free Dictionary API (English) ────────────────────

interface FreeDictDefinition {
  definition: string;
  example?: string;
}

interface FreeDictMeaning {
  partOfSpeech: string;
  definitions: FreeDictDefinition[];
}

interface FreeDictEntry {
  word: string;
  phonetic?: string;
  phonetics?: { text?: string; audio?: string }[];
  meanings: FreeDictMeaning[];
}

const FREE_DICT_URL = 'https://api.dictionaryapi.dev/api/v2/entries/en/';

async function fetchEnglishDefinition(word: string): Promise<{
  phonetic?: string;
  partOfSpeech?: string;
  enDefinitions: string[];
} | null> {
  try {
    const res = await fetch(`${FREE_DICT_URL}${encodeURIComponent(word)}`, {
      signal: AbortSignal.timeout(5000),
    });
    if (!res.ok) return null;

    const data: FreeDictEntry[] = await res.json();
    if (!data?.length) return null;

    const entry = data[0];

    const phonetic =
      entry.phonetic ||
      entry.phonetics?.find((p) => p.text)?.text ||
      undefined;

    const enDefinitions: string[] = [];
    let primaryPos: string | undefined;

    for (const meaning of entry.meanings) {
      if (!primaryPos) primaryPos = meaning.partOfSpeech;
      for (const def of meaning.definitions) {
        enDefinitions.push(def.definition);
      }
    }

    if (enDefinitions.length === 0) return null;

    return {
      phonetic,
      partOfSpeech: primaryPos,
      enDefinitions: enDefinitions.slice(0, 5),
    };
  } catch {
    return null;
  }
}

// ─── Chinese Translation Providers ────────────────────

// -- MyMemory (free, no API key) --
async function translateWithMyMemory(text: string): Promise<string | null> {
  try {
    const url = `https://api.mymemory.translated.net/get?q=${encodeURIComponent(text)}&langpair=en|zh`;
    const res = await fetch(url, { signal: AbortSignal.timeout(5000) });
    if (!res.ok) return null;
    const data = await res.json();
    const translated = data?.responseData?.translatedText;
    if (!translated || translated.toLowerCase() === text.toLowerCase()) return null;
    return translated;
  } catch {
    return null;
  }
}

// -- Youdao (set YOUDAA_APP_KEY and YOUDAA_APP_SECRET in .env.local) --
async function translateWithYoudao(text: string): Promise<string | null> {
  const appKey = process.env.YOUDAA_APP_KEY;
  const secret = process.env.YOUDAA_APP_SECRET;
  if (!appKey || !secret) return null;

  try {
    const salt = Date.now().toString();
    const encoder = new TextEncoder();
    const data = encoder.encode(appKey + text + salt + secret);
    const hashBuffer = await crypto.subtle.digest('SHA-256', data);
    const hashArray = Array.from(new Uint8Array(hashBuffer));
    const sign = hashArray.map((b) => b.toString(16).padStart(2, '0')).join('');

    const params = new URLSearchParams({
      q: text,
      from: 'en',
      to: 'zh-CHS',
      appKey,
      salt,
      sign,
    });

    const res = await fetch('https://openapi.youdao.com/api', {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: params.toString(),
      signal: AbortSignal.timeout(5000),
    });

    if (!res.ok) return null;
    const result = await res.json();
    return result?.translation?.[0] || null;
  } catch {
    return null;
  }
}

// -- Baidu (set BAIDU_TRANS_APP_ID and BAIDU_TRANS_KEY in .env.local) --
async function translateWithBaidu(text: string): Promise<string | null> {
  const appId = process.env.BAIDU_TRANS_APP_ID;
  const key = process.env.BAIDU_TRANS_KEY;
  if (!appId || !key) return null;

  try {
    const salt = Date.now().toString();
    const sign = crypto
      .createHash('md5')
      .update(appId + text + salt + key)
      .digest('hex');

    const params = new URLSearchParams({
      q: text,
      from: 'en',
      to: 'zh',
      appid: appId,
      salt,
      sign,
    });

    const res = await fetch('https://fanyi-api.baidu.com/api/trans/vip/translate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: params.toString(),
      signal: AbortSignal.timeout(5000),
    });

    if (!res.ok) return null;
    const result = await res.json();
    return result?.trans_result?.[0]?.dst || null;
  } catch {
    return null;
  }
}

async function fetchChineseTranslation(word: string): Promise<string | null> {
  // Try Youdao first if configured
  const youdaoResult = await translateWithYoudao(word);
  if (youdaoResult) return youdaoResult;

  // Try Baidu next if configured
  const baiduResult = await translateWithBaidu(word);
  if (baiduResult) return baiduResult;

  // Fall back to MyMemory (free, no key required)
  return translateWithMyMemory(word);
}

// ─── Main lookup ──────────────────────────────────────

export async function lookupWord(word: string): Promise<DictResult | null> {
  const clean = word.trim().toLowerCase();
  if (!clean) return null;

  const [enResult, zhResult] = await Promise.all([
    fetchEnglishDefinition(clean),
    fetchChineseTranslation(clean),
  ]);

  // Discard translation if it just echoes the input
  const zhDef = zhResult && zhResult.toLowerCase() !== clean ? zhResult : null;

  if (!enResult && !zhDef) return null;

  const providers: string[] = [];
  if (enResult) providers.push('Free Dictionary');
  if (zhDef) {
    if (process.env.YOUDAA_APP_KEY) providers.push('Youdao');
    else if (process.env.BAIDU_TRANS_APP_ID) providers.push('Baidu');
    else providers.push('MyMemory');
  }

  return {
    word: clean,
    phonetic: enResult?.phonetic,
    partOfSpeech: enResult?.partOfSpeech,
    enDefinitions: enResult?.enDefinitions ?? [],
    zhDefinition: zhDef ?? undefined,
    source: providers.join(' + ') || undefined,
  };
}
