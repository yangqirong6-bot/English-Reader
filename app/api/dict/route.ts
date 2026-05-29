import { NextResponse } from 'next/server';
import { MDX } from 'js-mdict';
import path from 'path';
import { lookupWord } from '@/lib/dictionary';

const dictPath = path.join(process.cwd(), 'dicts', 'oald10.mdx');

if (!(globalThis as any).__oxfordDict) {
  console.log('Loading Oxford dictionary into persistent memory (one-time)...');
  (globalThis as any).__oxfordDict = new MDX(dictPath);
}

// ── MDX redirect resolution ──────────────────────────

function resolveLink(dict: MDX, word: string, visited: Set<string> = new Set()): { html: string } | null {
  if (visited.has(word)) return null;
  visited.add(word);

  const result = dict.lookup(word);
  if (!result?.definition) return null;

  const linkPrefix = '@@@LINK=';
  if (result.definition.startsWith(linkPrefix)) {
    const target = result.definition.slice(linkPrefix.length).trim();
    console.log(`  ↳ @@@LINK: "${word}" → "${target}"`);
    return resolveLink(dict, target, visited);
  }

  return { html: result.definition };
}

// ── Stem-based lemmatization ─────────────────────────

const STEM_SUFFIXES = [
  { suffix: 'ies', restore: 'y' },
  { suffix: 'es', restore: 'e' },
  { suffix: 's', restore: '' },
  { suffix: 'ied', restore: 'y' },
  { suffix: 'ed', restore: 'e' },
  { suffix: 'd', restore: '' },
  { suffix: 'ing', restore: 'e' },
  { suffix: 'ly', restore: '' },
  { suffix: 'er', restore: 'e' },
  { suffix: 'est', restore: 'e' },
  { suffix: 'tion', restore: 'e' },
  { suffix: 'ment', restore: '' },
  { suffix: 'ness', restore: '' },
  { suffix: 'able', restore: 'e' },
  { suffix: 'ible', restore: 'e' },
];

function tryLemmatization(word: string): string[] {
  const candidates: string[] = [];
  for (const { suffix, restore } of STEM_SUFFIXES) {
    if (word.length > suffix.length + 2 && word.endsWith(suffix)) {
      const root = word.slice(0, -suffix.length);
      candidates.push(root + restore);
      candidates.push(root + 'e');
    }
  }
  return [...new Set(candidates)];
}

function mdxLookup(dict: MDX, word: string): { html: string } | null {
  // 1. Exact match with @@@LINK= resolution
  const exact = resolveLink(dict, word);
  if (exact) return exact;

  // 2. Try lemmatization with @@@LINK= resolution
  const candidates = tryLemmatization(word);
  for (const candidate of candidates) {
    const result = resolveLink(dict, candidate);
    if (result) return result;
  }

  return null;
}

// ── API Route ────────────────────────────────────────

export async function GET(request: Request) {
  const { searchParams } = new URL(request.url);
  const word = searchParams.get('word');

  if (!word) {
    return NextResponse.json({ error: 'Word is required' }, { status: 400 });
  }

  const cleanWord = word.trim().toLowerCase();

  // ── Phase 1: Local Oxford MDX lookup ──
  try {
    const dict = (globalThis as any).__oxfordDict as MDX;
    const localResult = mdxLookup(dict, cleanWord);

    if (localResult) {
      return NextResponse.json(localResult);
    }
  } catch (error) {
    console.error('MDX Lookup Error:', error);
    // fall through to online fallback
  }

  // ── Phase 2: Online dictionary fallback ──
  try {
    const result = await lookupWord(cleanWord);
    if (result) {
      return NextResponse.json(result);
    }
  } catch (error) {
    console.error('Online lookup error:', error);
  }

  return NextResponse.json({ error: 'Not found' }, { status: 404 });
}
