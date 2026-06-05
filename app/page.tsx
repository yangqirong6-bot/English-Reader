'use client';

import { useState, useCallback } from 'react';
import ArticleParser from '@/components/ArticleParser';
import VocabularySidebar from '@/components/VocabularySidebar';

export default function Home() {
  const [vocab, setVocab] = useState<Set<string>>(new Set());

  const toggleVocab = useCallback((word: string) => {
    setVocab((prev) => {
      const next = new Set(prev);
      if (next.has(word)) next.delete(word);
      else next.add(word);
      return next;
    });
  }, []);

  return (
    <main className="mx-auto min-h-screen max-w-6xl px-4 py-10">
      {/* Header */}
      <div className="mb-8 text-center">
        <h1 className="text-3xl font-bold tracking-tight text-gray-900">
          English Reader
        </h1>
        <p className="mt-1.5 text-sm text-gray-500">
          Paste an article, learn words, test your spelling
        </p>
      </div>

      <div className="flex gap-6 justify-center">
        <div className="flex-1 min-w-0 max-w-3xl">
          <ArticleParser
            placeholder="Paste an English article here…"
            onWordClick={(word) => console.log('Selected:', word)}
            vocab={vocab}
            onToggleVocab={toggleVocab}
          />
        </div>
        <VocabularySidebar
          words={Array.from(vocab)}
          onRemove={toggleVocab}
        />
      </div>
    </main>
  );
}
