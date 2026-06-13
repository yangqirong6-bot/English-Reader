'use client';

import { useState, useMemo, useCallback, useRef, useEffect } from 'react';
import { speakSentence, stopPlayback, speakWord } from '@/lib/azure-tts';
import { Document, Packer, Paragraph, TextRun, AlignmentType, HeadingLevel } from 'docx';
import Draggable from 'react-draggable';
import type { DictResult } from '@/lib/dict-types';

type DictApiResponse = DictResult | { html: string };

// ─── Types ───────────────────────────────────────────
interface WordToken {
  display: string;
  clean: string;
  id: string;
}

interface SentenceData {
  words: WordToken[];
  id: string;
}


// ─── Sentence splitting ─────────────────────────────
function splitSentences(text: string): string[] {
  const t = text.replace(/\s+/g, ' ').trim();
  if (!t) return [];

  // Protect abbreviations
  const abbrMarkers: { marker: string; original: string }[] = [];
  let processed = t;

  const patterns: [RegExp, string, number][] = [
    [/\b(Mr|Mrs|Ms|Dr|Prof|Sr|Jr|St|Sgt|Capt|Lt|Gen|Col|Sra|Srta)\./gi, '$1\x01', 0],
    [/\b(etc|vs|dept|est|approx|govt|Co|Inc|Ltd|Corp|Plc|LLC)\./gi, '$1\x02', 1],
    [/\b([A-Z])\.([A-Z])\./g, '$1\x03A', 2],
    [/\b(a\.m|p\.m)\./gi, '$1\x04', 3],
    [/\b(Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\./gi, '$1\x05', 4],
  ];

  patterns.forEach(([regex, replacement, idx]) => {
    const marker = String.fromCharCode(0x01 + idx);
    processed = processed.replace(regex, (...args) => {
      const match = args[0] as string;
      abbrMarkers.push({ marker, original: match });
      return match.replace(/\./g, marker);
    });
  });

  // Split at sentence boundaries
  const raw = processed.split(/(?<=[.!?])\s+(?=["'‘’“"(（\d]*[A-Z0-9])/);

  return raw
    .map((s: string) => {
      let restored = s;
      for (const { marker, original } of abbrMarkers) {
        restored = restored.replace(new RegExp(marker.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g'), '.');
      }
      restored = restored.replace(/[\x01-\x05]/g, '.');
      return restored.trim();
    })
    .filter(Boolean);
}

// ─── Word tokenization ───────────────────────────────
function tokenizeWords(sentence: string, sentenceIndex: number): WordToken[] {
  const parts = sentence.match(/\S+/g) || [];
  return parts.map((token, wordIdx) => {
    const id = `s${sentenceIndex}_w${wordIdx}`;
    const clean = token.replace(/^[^a-zA-Z0-9]+|[^a-zA-Z0-9]+$/g, '');
    return { display: token, clean, id };
  });
}

function parseArticle(text: string): SentenceData[] {
  const sentences = splitSentences(text);
  return sentences.map((s, i) => ({
    words: tokenizeWords(s, i),
    id: `sent_${i}`,
  }));
}

// ─── Component ───────────────────────────────────────
interface ArticleParserProps {
  onWordClick?: (word: string) => void;
  placeholder?: string;
  vocab: Set<string>;
  onToggleVocab: (word: string) => void;
}

export default function ArticleParser({ onWordClick, placeholder, vocab, onToggleVocab }: ArticleParserProps) {
  const [text, setText] = useState('');
  const [parsed, setParsed] = useState(false);
  const textareaRef = useRef<HTMLTextAreaElement>(null);
  const [activeWord, setActiveWord] = useState<WordToken | null>(null);
  const [dictData, setDictData] = useState<DictApiResponse | null>(null);
  const [isDictLoading, setIsDictLoading] = useState(false);
  const [playingSentenceId, setPlayingSentenceId] = useState<string | null>(null);
  const [testMode, setTestMode] = useState(false);
  const [correctTestIds, setCorrectTestIds] = useState<Set<string>>(new Set());
  const [flashIds, setFlashIds] = useState<Set<string>>(new Set());
  const [tooltipPos, setTooltipPos] = useState<{ top: number; left: number }>({ top: 0, left: 0 });
  const tooltipRef = useRef<HTMLDivElement>(null);
  const containerRef = useRef<HTMLDivElement>(null);
  const ttsAbortRef = useRef(false);
  const testInputRefs = useRef<Map<string, HTMLInputElement>>(new Map());
  const dictCacheRef = useRef<Map<string, DictApiResponse | null>>(new Map());
  const [showTranslation, setShowTranslation] = useState(false);
  const [sentenceTranslations, setSentenceTranslations] = useState<Map<string, string>>(new Map());
  const [translationsLoading, setTranslationsLoading] = useState<Set<string>>(new Set());
  const translationCacheRef = useRef<Map<string, string>>(new Map());

  const sentences = useMemo(() => {
    if (!parsed || !text.trim()) return [];
    return parseArticle(text);
  }, [text, parsed]);

  // ── Word click handler ──
  const handleWordClick = useCallback(
    async (e: React.MouseEvent, word: WordToken) => {
      e.stopPropagation();
      if (!word.clean) return;

      onWordClick?.(word.clean);

      // Stop sentence playback when user clicks a word
      if (playingSentenceId) {
        ttsAbortRef.current = true;
        stopPlayback();
        setPlayingSentenceId(null);
      }

      // Position tooltip near clicked word
      const rect = (e.target as HTMLElement).getBoundingClientRect();
      const tooltipWidth = 340;
      const tooltipHeight = 300;
      const top = Math.min(rect.bottom + 8, window.innerHeight - tooltipHeight - 16);
      const left = Math.min(
        Math.max(rect.left + rect.width / 2 - tooltipWidth / 2, 16),
        window.innerWidth - tooltipWidth - 16,
      );
      setTooltipPos({ top, left });

      // Show tooltip immediately with loading state
      setActiveWord(word);
      setDictData(null);
      setIsDictLoading(true);

      // Check cache first
      const cacheKey = word.clean.toLowerCase();
      const cached = dictCacheRef.current.get(cacheKey);
      if (cached !== undefined) {
        setDictData(cached);
        setIsDictLoading(false);
        return;
      }

      // Fetch dictionary definition
      try {
        const res = await fetch(`/api/dict?word=${encodeURIComponent(word.clean)}`);
        let result: DictApiResponse | null = null;
        if (res.ok) {
          const data = await res.json();
          if (!data.error) {
            result = data as DictApiResponse;
            setDictData(result);
          } else {
            setDictData(null);
          }
        } else {
          setDictData(null);
        }
        dictCacheRef.current.set(cacheKey, result);
      } catch {
        setDictData(null);
        dictCacheRef.current.set(cacheKey, null);
      } finally {
        setIsDictLoading(false);
      }
    },
    [onWordClick, playingSentenceId],
  );

  // ── Close tooltip ──
  useEffect(() => {
    if (!activeWord) return;

    const handleEscape = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        setActiveWord(null);
        setDictData(null);
      }
    };

    const id = setTimeout(() => {
      document.addEventListener('keydown', handleEscape);
    }, 0);

    return () => {
      clearTimeout(id);
      document.removeEventListener('keydown', handleEscape);
    };
  }, [activeWord]);

  // ── Parse handler ──
  const handleParse = useCallback(() => {
    // Read directly from DOM to avoid stale React state during hydration
    const currentText = textareaRef.current?.value ?? text;
    if (!currentText.trim()) return;
    if (!parsed || currentText !== text) {
      setText(currentText);
    }
    ttsAbortRef.current = true;
    stopPlayback();
    setParsed(true);
    setActiveWord(null);
    setDictData(null);
    setSentenceTranslations(new Map());
    setTranslationsLoading(new Set());
    translationCacheRef.current.clear();

    setPlayingSentenceId(null);
    if (testMode) {
      setTestMode(false);
      setCorrectTestIds(new Set());
      setFlashIds(new Set());
    }
  }, [text, testMode, parsed]);

  // ── Text change handler ──
  const handleTextChange = useCallback(
    (e: React.ChangeEvent<HTMLTextAreaElement>) => {
      setText(e.target.value);
      if (parsed) setParsed(false);
      setActiveWord(null);
      setDictData(null);
      if (playingSentenceId) {
        ttsAbortRef.current = true;
        stopPlayback();
        setPlayingSentenceId(null);
      }
    },
    [parsed, playingSentenceId],
  );

  // ── Keyboard shortcut ──
  const handleKeyDown = useCallback(
    (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
      if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
        e.preventDefault();
        handleParse();
      }
    },
    [handleParse],
  );

  // ── Close dictionary ──
  const closeDict = useCallback(() => {
    setActiveWord(null);
    setDictData(null);

  }, []);

  // ── Sentence TTS playback ──
  const handleSentencePlay = useCallback(
    async (sentenceId: string, text: string) => {
      // If already playing this sentence, stop it
      if (playingSentenceId === sentenceId) {
        ttsAbortRef.current = true;
        stopPlayback();
        setPlayingSentenceId(null);
        return;
      }

      // Stop any current playback and start new one
      ttsAbortRef.current = false;
      stopPlayback();
      setPlayingSentenceId(sentenceId);
      setActiveWord(null);
      setDictData(null);
  

      try {
        await speakSentence(text);
        if (!ttsAbortRef.current) {
          setPlayingSentenceId(null);
        }
      } catch (err) {
        if (!ttsAbortRef.current) {
          console.error('TTS playback error:', err);
          setPlayingSentenceId(null);
        }
      }
    },
    [playingSentenceId],
  );

  // Clean up TTS on unmount
  useEffect(() => {
    return () => {
      ttsAbortRef.current = true;
      stopPlayback();
    };
  }, []);

  // ── Translation fetching ──
  useEffect(() => {
    if (!showTranslation || sentences.length === 0) return;

    let cancelled = false;

    const fetchTranslations = async () => {
      for (const sentence of sentences) {
        if (cancelled) break;
        const sentenceText = sentence.words
          .map((w) => w.display)
          .join(' ')
          .trim();

        // Skip if already loaded
        if (sentenceTranslations.has(sentence.id)) continue;

        // Check cache
        const cacheKey = sentenceText;
        if (translationCacheRef.current.has(cacheKey)) {
          setSentenceTranslations((prev) => {
            const next = new Map(prev);
            next.set(sentence.id, translationCacheRef.current.get(cacheKey)!);
            return next;
          });
          continue;
        }

        setTranslationsLoading((prev) => new Set(prev).add(sentence.id));

        try {
          const res = await fetch(
            `/api/translate?text=${encodeURIComponent(sentenceText)}`,
          );
          if (res.ok && !cancelled) {
            const data = await res.json();
            if (data.translation) {
              translationCacheRef.current.set(cacheKey, data.translation);
              setSentenceTranslations((prev) => {
                const next = new Map(prev);
                next.set(sentence.id, data.translation);
                return next;
              });
            }
          }
        } catch {
          // ignore
        } finally {
          if (!cancelled) {
            setTranslationsLoading((prev) => {
              const next = new Set(prev);
              next.delete(sentence.id);
              return next;
            });
          }
        }
      }
    };

    fetchTranslations();

    return () => {
      cancelled = true;
    };
  }, [showTranslation, sentences]);

  // ── Test mode ──
  const toggleTestMode = useCallback(() => {
    setTestMode((prev) => {
      const next = !prev;
      if (prev) {
        // Exiting test mode — reset answers
        setCorrectTestIds(new Set());
        setFlashIds(new Set());
      } else {
        // Entering test mode — stop playback, close dictionary
        ttsAbortRef.current = true;
        stopPlayback();
        setPlayingSentenceId(null);
        setActiveWord(null);
        setDictData(null);
    
      }
      return next;
    });
  }, []);

  const handleTestKeyDown = useCallback(
    (e: React.KeyboardEvent<HTMLInputElement>, wordId: string, correctAnswer: string) => {
      if (e.key !== 'Enter') return;
      const input = e.currentTarget;
      const typed = input.value.trim();

      if (typed.toLowerCase() === correctAnswer.toLowerCase()) {
        setCorrectTestIds((prev) => {
          const next = new Set(prev);
          next.add(wordId);
          return next;
        });
        // Auto-advance to the next unfilled input by DOM order
        requestAnimationFrame(() => {
          const inputs = document.querySelectorAll<HTMLInputElement>('input[data-word-id]');
          const idx = Array.from(inputs).findIndex((el) => el.dataset.wordId === wordId);
          const next = Array.from(inputs).slice(idx + 1).find((el) => !el.value);
          if (next) next.focus();
        });
      } else {
        // Flash red
        setFlashIds((prev) => new Set(prev).add(wordId));
        setTimeout(() => {
          setFlashIds((prev) => {
            const next = new Set(prev);
            next.delete(wordId);
            return next;
          });
          // Re-focus
          const inp = testInputRefs.current.get(wordId);
          if (inp) { inp.value = ''; inp.focus(); }
        }, 400);
      }
    },
    [],
  );

  // ── Export to Word ──
  const [exporting, setExporting] = useState(false);

  const exportToWord = useCallback(async () => {
    if (vocab.size === 0 || exporting) return;
    setExporting(true);
    try {
      const vocabArray = Array.from(vocab);

      // One paragraph per word
      const wordParagraphs = vocabArray.map(
        (word) =>
          new Paragraph({
            spacing: { before: 120, after: 120 },
            children: [new TextRun({ text: word, size: 24, bold: true })],
          }),
      );

      // Build the document
      const doc = new Document({
        title: '雅思外刊阅读精读生词本',
        creator: 'English Learning App',
        styles: {
          default: {
            document: {
              run: { font: 'Arial', size: 22 },
              paragraph: { spacing: { after: 100 } },
            },
          },
        },
        sections: [
          {
            children: [
              new Paragraph({
                text: '雅思外刊阅读精读生词本',
                heading: HeadingLevel.HEADING_1,
                alignment: AlignmentType.CENTER,
                spacing: { after: 600, before: 200 },
              }),
              ...wordParagraphs,
            ],
          },
        ],
      });

      const blob = await Packer.toBlob(doc);
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'IELTS_My_Vocabulary.docx';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    } catch (err) {
      console.error('Error exporting Word document:', err);
    } finally {
      setExporting(false);
    }
  }, [vocab, exporting]);

  // ── Derived state ──
  const hasText = text.trim().length > 0;
  const parsedCount = sentences.reduce((sum, s) => sum + s.words.length, 0);

  // ── Render ──
  return (
    <div ref={containerRef} className="mx-auto max-w-3xl">
      {/* ── Input area ── */}
      <div className="mb-6 space-y-3">
        <textarea
          value={text}
          onChange={handleTextChange}
          onKeyDown={handleKeyDown}
          placeholder={placeholder ?? 'Paste an English article here...'}
          rows={6}
          className="w-full resize-y rounded-lg border border-gray-300 bg-white p-4 text-base leading-relaxed text-gray-900 shadow-sm transition-colors placeholder:text-gray-400 focus:border-blue-500 focus:outline-none focus:ring-2 focus:ring-blue-500/20"
        />

        <div className="flex items-center justify-between">
          <span className="text-sm text-gray-500">
            {text.length > 0 && `${text.length} characters`}
          </span>

          <button
            onClick={handleParse}
            className="inline-flex items-center gap-2 rounded-lg bg-blue-600 px-6 py-2.5 text-sm font-semibold text-white shadow-sm transition-all hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-blue-500/40 touch-manipulation min-h-[44px]"
          >
            <svg className="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
              <path strokeLinecap="round" strokeLinejoin="round" d="M10.325 4.317c.426-1.756 2.924-1.756 3.35 0a1.724 1.724 0 002.573 1.066c1.543-.94 3.31.826 2.37 2.37a1.724 1.724 0 001.066 2.573c1.756.426 1.756 2.924 0 3.35a1.724 1.724 0 00-1.066 2.573c.94 1.543-.826 3.31-2.37 2.37a1.724 1.724 0 00-2.573 1.066c-.426 1.756-2.924 1.756-3.35 0a1.724 1.724 0 00-2.573-1.066c-1.543.94-3.31-.826-2.37-2.37a1.724 1.724 0 00-1.066-2.573c-1.756-.426-1.756-2.924 0-3.35a1.724 1.724 0 001.066-2.573c-.94-1.543.826-3.31 2.37-2.37.996.608 2.296.07 2.572-1.065z" />
              <path strokeLinecap="round" strokeLinejoin="round" d="M15 12a3 3 0 11-6 0 3 3 0 016 0z" />
            </svg>
            Parse
          </button>
        </div>

        {parsed && (
          <div className="flex items-center justify-between">
            <p className="text-xs text-gray-400">
              Ctrl / Cmd + Enter &middot; {parsedCount} words across {sentences.length} sentences
              {vocab.size > 0 && ` · ${vocab.size} saved`}
            </p>
            <div className="flex items-center gap-2">
              <button
                onClick={() => setShowTranslation((prev) => !prev)}
                className={`
                  inline-flex items-center gap-1.5 rounded-lg px-3 py-1.5 text-xs font-semibold transition-all
                  ${
                    showTranslation
                      ? 'bg-emerald-100 text-emerald-700 ring-1 ring-emerald-300 hover:bg-emerald-200'
                      : 'bg-gray-100 text-gray-600 hover:bg-gray-200'
                  }
                `}
              >
                <svg className="h-3.5 w-3.5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                  <path strokeLinecap="round" strokeLinejoin="round" d="M3 5h12M9 3v2m1.048 9.5A18.022 18.022 0 016.412 9m6.088 9h7M11 21l5-10 5 10M12.751 5C11.783 10.77 8.07 15.61 3 18.129" />
                </svg>
                {showTranslation ? 'Hide Translation' : 'Translate'}
              </button>
              {vocab.size > 0 && (
                <>
                  <button
                    onClick={exportToWord}
                    disabled={exporting}
                    className="inline-flex items-center gap-1.5 rounded-lg bg-gradient-to-r from-violet-600 to-indigo-600 px-3 py-1.5 text-xs font-semibold text-white shadow-sm shadow-violet-300 transition-all hover:from-violet-700 hover:to-indigo-700 hover:shadow-md hover:shadow-violet-400 disabled:opacity-50 disabled:cursor-not-allowed"
                  >
                    <svg className="h-3.5 w-3.5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                      <path strokeLinecap="round" strokeLinejoin="round" d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                    </svg>
                    {exporting ? 'Generating...' : 'Export Word'}
                  </button>
                  <button
                    onClick={toggleTestMode}
                    className={`
                      inline-flex items-center gap-1.5 rounded-lg px-3 py-1.5 text-xs font-semibold transition-all
                      ${
                        testMode
                          ? 'bg-indigo-100 text-indigo-700 ring-1 ring-indigo-300 hover:bg-indigo-200'
                          : 'bg-gray-100 text-gray-600 hover:bg-gray-200'
                      }
                    `}
                  >
                    <svg className="h-3.5 w-3.5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                      {testMode
                        ? <path strokeLinecap="round" strokeLinejoin="round" d="M6 18L18 6M6 6l12 12" />
                        : <path strokeLinecap="round" strokeLinejoin="round" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                      }
                    </svg>
                    {testMode ? 'Exit Test Mode' : 'Dictation Test'}
                  </button>
                </>
              )}
            </div>
          </div>
        )}
      </div>

      {/* ── Empty state ── */}
      {!parsed && (
        <div className="flex flex-col items-center justify-center rounded-xl border-2 border-dashed border-gray-200 py-16 text-center text-gray-400">
          <svg className="mb-3 h-10 w-10" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={1.5}>
            <path strokeLinecap="round" strokeLinejoin="round" d="M19.5 14.25v-2.625a3.375 3.375 0 00-3.375-3.375h-1.5A1.125 1.125 0 0113.5 7.125v-1.5a3.375 3.375 0 00-3.375-3.375H8.25m0 12.75h7.5m-7.5 3H12M10.5 2.25H5.625c-.621 0-1.125.504-1.125 1.125v17.25c0 .621.504 1.125 1.125 1.125h12.75c.621 0 1.125-.504 1.125-1.125V11.25a9 9 0 00-9-9z" />
          </svg>
          <p className="text-sm">Paste an article above and click Parse to get started</p>
        </div>
      )}

      {/* ── Parse failure ── */}
      {parsed && sentences.length === 0 && hasText && (
        <div className="rounded-xl border border-yellow-200 bg-yellow-50 px-4 py-6 text-center text-sm text-yellow-700">
          Could not split the text into sentences. Try pasting a different article.
        </div>
      )}

      {/* ── Test mode banner ── */}
      {testMode && (
        <div className="mb-4 flex items-center gap-2 rounded-lg border border-indigo-200 bg-indigo-50 px-4 py-2.5 text-sm text-indigo-700">
          <svg className="h-4 w-4 flex-shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
            <path strokeLinecap="round" strokeLinejoin="round" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
          </svg>
          <span>
            <strong>Dictation Test Mode</strong> — Type the correct spelling of each word and press <kbd className="rounded bg-indigo-200 px-1 py-0.5 font-mono text-xs">Enter</kbd> to check.
          </span>
        </div>
      )}

      {/* ── Parse result ── */}
      {parsed && sentences.length > 0 && (
        <div className={`space-y-3 ${testMode ? 'rounded-xl border border-indigo-100 bg-white/80 p-4' : ''}`}>
          {sentences.map((sentence) => {
            const isPlaying = playingSentenceId === sentence.id;
            const sentenceText = sentence.words
              .map((w) => w.display)
              .join(' ')
              .trim();

            return (
              <div
                key={sentence.id}
                className={`
                  group relative flex items-start gap-2 rounded-xl px-3 py-2
                  transition-all duration-300
                  ${
                    isPlaying
                      ? 'bg-blue-50 shadow-sm ring-1 ring-blue-200'
                      : 'hover:bg-gray-50/60'
                  }
                `}
              >
                {/* ── Sentence play button ── */}
                {typeof window !== 'undefined' && window.speechSynthesis && (
                  <button
                    onClick={(e) => {
                      e.stopPropagation();
                      handleSentencePlay(sentence.id, sentenceText);
                    }}
                    aria-label={isPlaying ? 'Stop' : 'Play sentence'}
                    title={isPlaying ? 'Stop' : 'Play sentence'}
                    className={`
                      mt-1 flex-shrink-0 rounded-lg p-1.5 transition-colors
                      focus:outline-none focus:ring-2 focus:ring-blue-300
                      ${
                        isPlaying
                          ? 'bg-blue-100 text-blue-600 hover:bg-blue-200'
                          : 'text-gray-400 opacity-0 group-hover:opacity-100 hover:text-blue-500 hover:opacity-100'
                      }
                    `}
                  >
                    {isPlaying ? (
                      /* Stop icon */
                      <svg className="h-4 w-4" viewBox="0 0 24 24" fill="currentColor">
                        <rect x="6" y="6" width="12" height="12" rx="1" />
                      </svg>
                    ) : (
                      /* Play icon */
                      <svg className="h-4 w-4" viewBox="0 0 24 24" fill="currentColor">
                        <path d="M8 5.14v13.72a1 1 0 001.5.86l11-6.86a1 1 0 000-1.72l-11-6.86A1 1 0 008 5.14z" />
                      </svg>
                    )}
                  </button>
                )}

                {/* ── Words + Translation ── */}
                <div className="flex-1 min-w-0">
                <p className="leading-[2.2]">
                  {sentence.words.map((word) => {
                    const isSaved = vocab.has(word.clean);

                    // ── Test mode: render input for vocab words ──
                    if (testMode && isSaved) {
                      const isCorrect = correctTestIds.has(word.id);
                      const isFlashing = flashIds.has(word.id);

                      if (isCorrect) {
                        return (
                          <span
                            key={word.id}
                            className="inline-block rounded px-0.5 text-base font-medium leading-[2.2] text-green-600"
                          >
                            {word.clean}{' '}
                          </span>
                        );
                      }

                      return (
                        <input
                          key={word.id}
                          ref={(el) => {
                            if (el) testInputRefs.current.set(word.id, el);
                            else testInputRefs.current.delete(word.id);
                          }}
                          type="text"
                          defaultValue=""
                          spellCheck={false}
                          autoComplete="off"
                          data-word-id={word.id}
                          placeholder={' '.repeat(Math.max(word.clean.length - 1, 1))}
                          data-word={word.clean}
                          style={{
                            width: `calc(${word.clean.length}ch + 1.2rem)`,
                            minWidth: '4rem',
                          }}
                          onKeyDown={(e) => handleTestKeyDown(e, word.id, word.clean)}
                          className={`
                            mx-0.5 inline-block rounded border-b-2 bg-transparent
                            px-1.5 py-0 text-center text-base
                            outline-none transition-all duration-200
                            ${
                              isFlashing
                                ? 'border-red-400 bg-red-50'
                                : 'border-gray-300 focus:border-indigo-400'
                            }
                          `}
                        />
                      );
                    }

                    // ── Normal mode: clickable word ──
                    return (
                      <span
                        key={word.id}
                        onClick={(e) => handleWordClick(e, word)}
                        onKeyDown={(e) => {
                          if (e.key === 'Enter' || e.key === ' ') {
                            e.preventDefault();
                            const clickEvent = { ...e, target: e.currentTarget } as unknown as React.MouseEvent;
                            handleWordClick(clickEvent, word);
                          }
                        }}
                        tabIndex={testMode ? -1 : 0}
                        role="button"
                        aria-label={`Word: ${word.clean}`}
                        className={`
                          inline-block cursor-pointer rounded px-0.5 text-base leading-[2.2] transition-colors
                          focus:outline-none focus:ring-2 focus:ring-yellow-300/60
                          ${isSaved ? 'text-red-600 hover:bg-red-100' : 'text-gray-800 hover:bg-yellow-200'}
                        `}
                      >
                        {word.display}{' '}
                      </span>
                    );
                  })}
                </p>

                {/* ── Translation ── */}
                {showTranslation && (
                  translationsLoading.has(sentence.id) ? (
                    <p className="mt-1 pl-2 border-l-2 border-gray-200 text-sm text-gray-400 animate-pulse">
                      Translating...
                    </p>
                  ) : sentenceTranslations.has(sentence.id) ? (
                    <p className="mt-1 pl-2 border-l-2 border-emerald-300 text-sm text-gray-600 leading-relaxed">
                      {sentenceTranslations.get(sentence.id)}
                    </p>
                  ) : null
                )}
                </div>
              </div>
            );
          })}
        </div>
      )}

      {/* ── Dictionary Popover ── */}
      {activeWord && (
        <div className="fixed z-[9999]" style={{ top: tooltipPos.top, left: tooltipPos.left }}>
          <Draggable nodeRef={tooltipRef} handle=".drag-handle">
            <div
              ref={tooltipRef}
              role="dialog"
              aria-label={`Dictionary: ${activeWord.clean}`}
              className="flex flex-col resize overflow-hidden min-w-[300px] min-h-[250px] max-w-[90vw] max-h-[90vh] pb-8 animate-in fade-in slide-in-from-top-2 rounded-xl border border-gray-200 bg-white shadow-xl"
            >
            {/* Drag handle header */}
            <div className="drag-handle cursor-move flex-shrink-0 flex items-center justify-between border-b border-gray-100 bg-gray-50 px-4 py-2.5 rounded-t-xl">
              <div className="flex items-center gap-2 min-w-0">
                <svg className="h-4 w-4 flex-shrink-0 text-gray-400" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                  <path strokeLinecap="round" strokeLinejoin="round" d="M4 8h16M4 16h16" />
                </svg>
                <span className="text-sm font-bold text-gray-800 truncate">{activeWord.clean}</span>
              </div>
              <div className="flex items-center gap-1">
                {/* Speak button */}
                <button
                  onClick={(e) => { e.stopPropagation(); speakWord(activeWord.clean); }}
                  aria-label="Pronounce"
                  className="rounded-lg p-1.5 text-gray-500 transition-colors hover:bg-gray-100 hover:text-gray-700"
                >
                  <svg className="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                    <path strokeLinecap="round" strokeLinejoin="round" d="M15.536 8.464a5 5 0 010 7.072M17.95 6.05a8 8 0 010 11.9M6.5 8.8l4.3-4.3a1 1 0 011.7.7v13.6a1 1 0 01-1.7.7l-4.3-4.3H4a1 1 0 01-1-1v-4.4a1 1 0 011-1h2.5z" />
                  </svg>
                </button>
                {/* Close button */}
                <button
                  onClick={(e) => { e.stopPropagation(); closeDict(); }}
                  aria-label="Close"
                  className="rounded-lg p-1.5 text-gray-400 transition-colors hover:bg-gray-100 hover:text-gray-600"
                >
                  <svg className="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                    <path strokeLinecap="round" strokeLinejoin="round" d="M6 18L18 6M6 6l12 12" />
                  </svg>
                </button>
              </div>
            </div>

            {/* Body */}
            <div className="flex min-h-0 flex-1 flex-col px-4 py-3 overflow-y-auto overscroll-contain">
              {isDictLoading ? (
                <div className="flex flex-col items-center justify-center py-8 gap-2">
                  <svg className="h-5 w-5 animate-spin text-indigo-400" fill="none" viewBox="0 0 24 24">
                    <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
                    <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z" />
                  </svg>
                  <span className="text-xs text-gray-400">Looking up &ldquo;{activeWord.clean}&rdquo;...</span>
                </div>
              ) : dictData && 'html' in dictData ? (
                <div
                  className="oxford-dict oxford-dict-content text-sm p-4 h-full w-full overflow-y-auto"
                  dangerouslySetInnerHTML={{ __html: dictData.html }}
                />
              ) : dictData ? (
                <div className="flex min-h-0 flex-1 flex-col overflow-y-auto space-y-3">
                  {/* Phonetic + Part of Speech */}
                  {(dictData.phonetic || dictData.partOfSpeech) && (
                    <div className="flex items-center gap-2 text-sm">
                      {dictData.phonetic && (
                        <span className="text-gray-500 font-mono">{dictData.phonetic}</span>
                      )}
                      {dictData.partOfSpeech && (
                        <span className="inline-block rounded bg-blue-50 px-2 py-0.5 text-xs font-medium text-blue-600 italic">
                          {dictData.partOfSpeech}
                        </span>
                      )}
                    </div>
                  )}

                  {/* English definitions */}
                  {dictData.enDefinitions.length > 0 && (
                    <div>
                      <span className="text-xs font-semibold text-gray-400 uppercase tracking-wide">English</span>
                      <ol className="mt-1 space-y-1.5 list-decimal list-inside">
                        {dictData.enDefinitions.map((def, i) => (
                          <li key={i} className="text-sm text-gray-700 leading-relaxed">{def}</li>
                        ))}
                      </ol>
                    </div>
                  )}

                  {/* Chinese definition */}
                  {dictData.zhDefinition && (
                    <div>
                      <span className="text-xs font-semibold text-gray-400 uppercase tracking-wide">中文</span>
                      <p className="mt-1 text-sm text-gray-700">{dictData.zhDefinition}</p>
                    </div>
                  )}

                  {/* Source */}
                  {dictData.source && (
                    <p className="text-[10px] text-gray-300 text-right mt-2">{dictData.source}</p>
                  )}
                </div>
              ) : (
                <p className="py-4 text-center text-sm text-gray-400">
                  No definition found for &ldquo;{activeWord.clean}&rdquo;
                </p>
              )}
            </div>

            {/* Footer */}
            {!isDictLoading && (
              <div className="flex-shrink-0 border-t border-gray-100 px-4 py-2.5">
                <button
                  onClick={(e) => { e.stopPropagation(); onToggleVocab(activeWord.clean); }}
                  className={`flex w-full items-center justify-center gap-1.5 rounded-lg px-3 py-2 text-sm font-medium transition-colors ${
                    vocab.has(activeWord.clean)
                      ? 'bg-red-50 text-red-600 hover:bg-red-100'
                      : 'bg-gray-50 text-gray-600 hover:bg-gray-100 hover:text-gray-800'
                  }`}
                >
                  <svg className="h-4 w-4" fill={vocab.has(activeWord.clean) ? 'currentColor' : 'none'} viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                    <path strokeLinecap="round" strokeLinejoin="round" d="M5 5a2 2 0 012-2h10a2 2 0 012 2v16l-7-3.5L5 21V5z" />
                  </svg>
                  {vocab.has(activeWord.clean) ? 'Remove from vocabulary' : 'Add to vocabulary'}
                </button>
              </div>
            )}

            {/* Resize handle indicator */}
            <div className="pointer-events-none absolute bottom-0 right-0 flex items-end justify-end p-1 text-gray-300">
              <svg width="16" height="16" viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.5" className="drop-shadow-sm">
                <line x1="11" y1="16" x2="16" y2="11" />
                <line x1="6" y1="16" x2="16" y2="6" />
                <line x1="1" y1="16" x2="16" y2="1" />
              </svg>
            </div>
          </div>
          </Draggable>
        </div>
      )}
    </div>
  );
}
