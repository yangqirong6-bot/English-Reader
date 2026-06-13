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

async function fetchFreeDictDefinition(word: string): Promise<{
  phonetic?: string;
  partOfSpeech?: string;
  enDefinitions: string[];
  source: string;
} | null> {
  try {
    const res = await fetch(
      `https://api.dictionaryapi.dev/api/v2/entries/en/${encodeURIComponent(word)}`,
      { signal: AbortSignal.timeout(5000) },
    );
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
      source: 'Free Dictionary',
    };
  } catch {
    return null;
  }
}

// ─── Wiktionary API (English, free, no key) ──────────

interface WiktionaryDefinition {
  definition: string;
  partOfSpeech?: string;
}

async function fetchWiktionaryDefinition(word: string): Promise<{
  phonetic?: string;
  partOfSpeech?: string;
  enDefinitions: string[];
  source: string;
} | null> {
  try {
    const url = `https://en.wiktionary.org/api/rest_v1/page/definition/${encodeURIComponent(word)}`;
    const res = await fetch(url, { signal: AbortSignal.timeout(5000) });
    if (!res.ok) return null;

    const data: Record<string, WiktionaryDefinition[]> = await res.json();
    if (!data) return null;

    const enDefinitions: string[] = [];
    let primaryPos: string | undefined;

    const langKey = Object.keys(data).find((k) => k === 'en') || Object.keys(data)[0];
    if (!langKey || !data[langKey]) return null;

    for (const item of data[langKey]) {
      if (!primaryPos && item.partOfSpeech) primaryPos = item.partOfSpeech;
      if (item.definition) {
        const clean = item.definition.replace(/<[^>]+>/g, '');
        enDefinitions.push(clean);
      }
    }

    if (enDefinitions.length === 0) return null;

    return {
      partOfSpeech: primaryPos,
      enDefinitions: enDefinitions.slice(0, 5),
      source: 'Wiktionary',
    };
  } catch {
    return null;
  }
}

// ─── Combined English lookup (parallel) ───────────────

async function fetchEnglishDefinition(word: string): Promise<{
  result: {
    phonetic?: string;
    partOfSpeech?: string;
    enDefinitions: string[];
  } | null;
  source: string;
}> {
  // Fire both in parallel, take first success
  const [freeDict, wiktionary] = await Promise.allSettled([
    fetchFreeDictDefinition(word),
    fetchWiktionaryDefinition(word),
  ]);

  if (freeDict.status === 'fulfilled' && freeDict.value) {
    return { result: freeDict.value, source: freeDict.value.source };
  }
  if (wiktionary.status === 'fulfilled' && wiktionary.value) {
    return { result: wiktionary.value, source: wiktionary.value.source };
  }

  return { result: null, source: 'Free Dictionary' };
}

// ─── Chinese Translation Providers ────────────────────

// Google Translate (free, no key)
async function translateWithGoogle(text: string): Promise<string | null> {
  try {
    const url = `https://translate.googleapis.com/translate_a/single?client=gtx&sl=en&tl=zh-CN&dt=t&q=${encodeURIComponent(text)}`;
    const res = await fetch(url, { signal: AbortSignal.timeout(5000) });
    if (!res.ok) return null;
    const data = await res.json();
    // Response format: [[["translated text","original",...]],...]
    const translated = data?.[0]?.[0]?.[0];
    if (!translated || translated.toLowerCase() === text.toLowerCase()) return null;
    return translated;
  } catch {
    return null;
  }
}

// DeepL (requires API key)
async function translateWithDeepL(text: string): Promise<string | null> {
  const apiKey = process.env.DEEPL_API_KEY;
  if (!apiKey) return null;

  try {
    const isFree = apiKey.endsWith(':fx');
    const baseUrl = isFree
      ? 'https://api-free.deepl.com/v2/translate'
      : 'https://api.deepl.com/v2/translate';

    const params = new URLSearchParams({
      text,
      source_lang: 'EN',
      target_lang: 'ZH',
    });

    const res = await fetch(baseUrl, {
      method: 'POST',
      headers: {
        'Authorization': `DeepL-Auth-Key ${apiKey}`,
        'Content-Type': 'application/x-www-form-urlencoded',
      },
      body: params.toString(),
      signal: AbortSignal.timeout(5000),
    });

    if (!res.ok) return null;
    const result = await res.json();
    const translated = result?.translations?.[0]?.text;
    if (!translated || translated.toLowerCase() === text.toLowerCase()) return null;
    return translated;
  } catch {
    return null;
  }
}

// MyMemory (free, no key)
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

// Youdao (requires API key) — uses Node.js crypto, NOT Web crypto.subtle
async function translateWithYoudao(text: string): Promise<string | null> {
  const appKey = process.env.YOUDAA_APP_KEY;
  const secret = process.env.YOUDAA_APP_SECRET;
  if (!appKey || !secret) return null;

  try {
    const salt = Date.now().toString();
    const sign = crypto
      .createHash('sha256')
      .update(appKey + text + salt + secret)
      .digest('hex');

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
    const translated = result?.translation?.[0];
    if (!translated || translated.toLowerCase() === text.toLowerCase()) return null;
    return translated;
  } catch {
    return null;
  }
}

// Baidu (requires API key)
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

    const res = await fetch(
      'https://fanyi-api.baidu.com/api/trans/vip/translate',
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: params.toString(),
        signal: AbortSignal.timeout(5000),
      },
    );

    if (!res.ok) return null;
    const result = await res.json();
    const translated = result?.trans_result?.[0]?.dst;
    if (!translated || translated.toLowerCase() === text.toLowerCase()) return null;
    return translated;
  } catch {
    return null;
  }
}

// Run all translators in parallel, pick best result by priority
export async function fetchChineseTranslation(word: string): Promise<{
  text: string;
  source: string;
} | null> {
  const results = await Promise.allSettled([
    translateWithDeepL(word),
    translateWithYoudao(word),
    translateWithBaidu(word),
    translateWithGoogle(word),
    translateWithMyMemory(word),
  ]);

  const sources = ['DeepL', 'Youdao', 'Baidu', 'Google', 'MyMemory'];

  for (let i = 0; i < results.length; i++) {
    const r = results[i];
    if (r.status === 'fulfilled' && r.value) {
      return { text: r.value, source: sources[i] };
    }
  }
  return null;
}

// ─── Main lookup ──────────────────────────────────────

export async function lookupWord(word: string): Promise<DictResult | null> {
  const clean = word.trim().toLowerCase();
  if (!clean) return null;

  const [{ result: enResult, source: enSource }, zhResult] = await Promise.all([
    fetchEnglishDefinition(clean),
    fetchChineseTranslation(clean),
  ]);

  if (!enResult && !zhResult) return null;

  const providers: string[] = [];
  if (enResult) providers.push(enSource);
  if (zhResult) providers.push(zhResult.source);

  return {
    word: clean,
    phonetic: enResult?.phonetic,
    partOfSpeech: enResult?.partOfSpeech,
    enDefinitions: enResult?.enDefinitions ?? [],
    zhDefinition: zhResult?.text ?? undefined,
    source: providers.join(' + ') || undefined,
  };
}
