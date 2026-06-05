'use client';

function speakWord(word: string) {
  if (typeof window === 'undefined') return;
  window.speechSynthesis.cancel();
  const u = new SpeechSynthesisUtterance(word);
  u.rate = 0.85;
  u.lang = 'en-US';
  const voices = window.speechSynthesis.getVoices();
  const voice = voices.find((v) => v.lang.startsWith('en'));
  if (voice) u.voice = voice;
  window.speechSynthesis.speak(u);
}

interface VocabularySidebarProps {
  words: string[];
  onRemove: (word: string) => void;
}

export default function VocabularySidebar({ words, onRemove }: VocabularySidebarProps) {
  if (words.length === 0) return null;

  return (
    <div className="w-full lg:w-56 flex-shrink-0">
      <div className="sticky top-4 rounded-xl border border-gray-200 bg-white shadow-sm">
        <div className="border-b border-gray-100 px-4 py-3">
          <h2 className="text-sm font-semibold text-gray-800">生词本</h2>
          <p className="text-xs text-gray-400 mt-0.5">{words.length} word{words.length !== 1 ? 's' : ''}</p>
        </div>
        <ul className="max-h-[70vh] overflow-y-auto px-2 py-2 space-y-0.5">
          {words.map((word) => (
            <li
              key={word}
              className="flex items-center justify-between rounded-lg px-2 py-1.5 hover:bg-gray-50 group"
            >
              <button
                onClick={() => speakWord(word)}
                className="text-sm text-gray-700 hover:text-blue-600 text-left truncate"
                title="Click to pronounce"
              >
                {word}
              </button>
              <button
                onClick={() => onRemove(word)}
                className="opacity-0 group-hover:opacity-100 text-gray-400 hover:text-red-500 transition-opacity p-0.5 flex-shrink-0"
                title="Remove from vocabulary"
              >
                <svg className="h-3.5 w-3.5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                  <path strokeLinecap="round" strokeLinejoin="round" d="M6 18L18 6M6 6l12 12" />
                </svg>
              </button>
            </li>
          ))}
        </ul>
      </div>
    </div>
  );
}
