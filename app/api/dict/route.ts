import { NextResponse } from 'next/server';
import { MDX } from 'js-mdict';
import path from 'path';

const dictPath = path.join(process.cwd(), 'dicts', 'oald10.mdx');

if (!(globalThis as any).__oxfordDict) {
  console.log('Loading Oxford dictionary into persistent memory (one-time)...');
  (globalThis as any).__oxfordDict = new MDX(dictPath);
}

const FALLBACK_URL = 'https://api.dictionaryapi.dev/api/v2/entries/en/';
const WIKIPEDIA_URL = 'https://en.wikipedia.org/api/rest_v1/page/summary/';

// Simple stem-based lemmatization (no external package needed)
const STEM_SUFFIXES = [
  'ies', 'es', 's',
  'ied', 'ed', 'd',
  'ing',
  'ly',
  'er', 'est',
  'tion', 'ment', 'ness', 'able', 'ible',
];

function tryLemmatization(word: string): string[] {
  const candidates: string[] = [];
  for (const suffix of STEM_SUFFIXES) {
    if (word.length > suffix.length + 2 && word.endsWith(suffix)) {
      const root = word.slice(0, -suffix.length);
      candidates.push(root);
      // try adding back 'e' (e.g. baking → bake, moving → move)
      candidates.push(root + 'e');
    }
  }
  return [...new Set(candidates)];
}

function lookupWithFallback(dict: MDX, word: string): { html: string; fromLemmatization?: boolean } | null {
  // 1. Exact match
  const exact = dict.lookup(word);
  if (exact?.definition) return { html: exact.definition };

  // 2. Try lemmatization
  const candidates = tryLemmatization(word);
  for (const candidate of candidates) {
    const result = dict.lookup(candidate);
    if (result?.definition) return { html: result.definition, fromLemmatization: true };
  }

  return null;
}

interface FreeDictPhonetic {
  text?: string;
  audio?: string;
}

interface FreeDictDefinition {
  definition: string;
  example?: string;
  synonyms?: string[];
  antonyms?: string[];
}

interface FreeDictMeaning {
  partOfSpeech: string;
  definitions: FreeDictDefinition[];
  synonyms?: string[];
  antonyms?: string[];
}

interface FreeDictEntry {
  word: string;
  phonetic?: string;
  phonetics?: FreeDictPhonetic[];
  meanings: FreeDictMeaning[];
  origin?: string;
  license?: object;
  sourceUrls?: string[];
}

interface WikipediaPage {
  title: string;
  extract: string;
  extract_html?: string;
  description?: string;
  thumbnail?: { source: string; width: number; height: number };
  content_urls?: { desktop: { page: string } };
}

function buildFallbackHtml(data: FreeDictEntry[]): string {
  const entry = data[0];

  const phonetic =
    entry.phonetic ||
    entry.phonetics?.find((p) => p.text)?.text ||
    '';

  const meaningsHtml = entry.meanings
    .map((m) => {
      const defsHtml = m.definitions
        .slice(0, 5)
        .map((d) => {
          const exampleHtml = d.example
            ? `<div class="border-l-2 border-gray-300 pl-2 mt-1 text-gray-500 text-sm">${escapeHtml(d.example)}</div>`
            : '';
          return `<li class="text-gray-800 text-sm leading-relaxed">${escapeHtml(d.definition)}${exampleHtml}</li>`;
        })
        .join('');

      return `<div class="mb-3">
        <span class="text-blue-600 font-semibold italic text-sm">${escapeHtml(m.partOfSpeech)}</span>
        <ul class="list-disc pl-5 mt-1 space-y-1.5">${defsHtml}</ul>
      </div>`;
    })
    .join('');

  return `<div class="online-dict-fallback p-2">
    <div class="border-b border-gray-200 pb-2 mb-3">
      <span class="text-lg font-bold text-gray-900">${escapeHtml(entry.word)}</span>
      ${phonetic ? `<span class="text-gray-400 text-sm ml-2">${escapeHtml(phonetic)}</span>` : ''}
    </div>
    ${meaningsHtml}
  </div>`;
}

function buildWikipediaHtml(page: WikipediaPage): string {
  const extract = page.extract_html || page.extract;
  if (!extract) throw new Error('No extract');

  return `<div class="online-dict-fallback p-2">
    <div class="border-b border-gray-200 pb-2 mb-3">
      <span class="text-lg font-bold text-gray-900">${escapeHtml(page.title)}</span>
      ${page.description ? `<span class="text-gray-400 text-sm ml-2">${escapeHtml(page.description)}</span>` : ''}
    </div>
    <div class="text-sm text-gray-700 leading-relaxed [&>p]:mb-2">${extract}</div>
    <div class="mt-3 text-xs text-gray-400 text-right">via Wikipedia</div>
  </div>`;
}

function escapeHtml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

export async function GET(request: Request) {
  const { searchParams } = new URL(request.url);
  const word = searchParams.get('word');

  if (!word) {
    return NextResponse.json({ error: 'Word is required' }, { status: 400 });
  }

  const cleanWord = word.trim().toLowerCase();

  // ── Phase 1: Local MDX lookup ──
  try {
    const dict = (globalThis as any).__oxfordDict as MDX;
    const localResult = lookupWithFallback(dict, cleanWord);

    if (localResult) {
      return NextResponse.json({ html: localResult.html });
    }
  } catch (error) {
    console.error('MDX Lookup Error:', error);
    // fall through to network fallback instead of failing immediately
  }

  // ── Phase 2: Free Dictionary API ──
  try {
    const res = await fetch(`${FALLBACK_URL}${encodeURIComponent(cleanWord)}`, {
      signal: AbortSignal.timeout(5000),
    });

    if (res.ok) {
      try {
        const data: FreeDictEntry[] = await res.json();
        const html = buildFallbackHtml(data);
        return NextResponse.json({ html });
      } catch (err) {
        console.error('Free Dictionary: parse/build failed, trying Wikipedia...', err);
      }
    }
  } catch (error) {
    console.error('Free Dictionary: network error, trying Wikipedia...', error);
  }

  // ── Phase 3: Wikipedia fallback ──
  try {
    const res = await fetch(`${WIKIPEDIA_URL}${encodeURIComponent(cleanWord)}`, {
      signal: AbortSignal.timeout(5000),
      headers: { 'Accept': 'application/json' },
    });

    if (res.ok) {
      const data: WikipediaPage = await res.json();
      if (data.extract) {
        const html = buildWikipediaHtml(data);
        return NextResponse.json({ html });
      }
    }
  } catch (error) {
    console.error('Wikipedia fallback error:', error);
  }

  return NextResponse.json({ error: 'Not found' }, { status: 404 });
}
