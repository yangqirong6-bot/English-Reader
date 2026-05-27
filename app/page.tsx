'use client';

import ArticleParser from '@/components/ArticleParser';

export default function Home() {
  return (
    <main className="mx-auto min-h-screen max-w-4xl px-4 py-10">
      {/* Header */}
      <div className="mb-8 text-center">
        <h1 className="text-3xl font-bold tracking-tight text-gray-900">
          English Reader
        </h1>
        <p className="mt-1.5 text-sm text-gray-500">
          Paste an article, learn words, test your spelling
        </p>
      </div>

      <ArticleParser
        placeholder="Paste an English article here…"
        onWordClick={(word) => console.log('Selected:', word)}
      />
    </main>
  );
}
