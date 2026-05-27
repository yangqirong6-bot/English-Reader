/**
 * Sentence TTS playback using the browser's Web Speech API.
 *
 * No API keys or external services needed — works offline for supported voices.
 * Falls back gracefully if the browser has no en-GB voice available.
 *
 * Usage:
 *   import { speakSentence, stopPlayback } from '@/lib/azure-tts';
 *
 *   await speakSentence('Hello world.');
 *   speakSentence('Hello world.', { subscriptionKey: '', region: '' }); // legacy compat
 */

export interface TtsConfig {
  /** Kept for backwards compatibility — Web Speech API ignores these. */
  subscriptionKey?: string;
  region?: string;
}

// ── State ───────────────────────────────────────────

let currentUtterance: SpeechSynthesisUtterance | null = null;
let fallbackResolve: (() => void) | null = null;
let voicePromise: Promise<SpeechSynthesisVoice[]> | null = null;

// ── Voice loading ───────────────────────────────────

function loadVoices(): Promise<SpeechSynthesisVoice[]> {
  if (voicePromise) return voicePromise;

  voicePromise = new Promise((resolve) => {
    const voices = window.speechSynthesis.getVoices();
    if (voices.length > 0) {
      resolve(voices);
      return;
    }

    // Chrome loads voices asynchronously
    const timeout = setTimeout(() => {
      resolve(window.speechSynthesis.getVoices());
    }, 3000);

    window.speechSynthesis.onvoiceschanged = () => {
      clearTimeout(timeout);
      resolve(window.speechSynthesis.getVoices());
    };
  });

  return voicePromise;
}

function pickVoice(voices: SpeechSynthesisVoice[]): SpeechSynthesisVoice | null {
  // 1st choice — Edge online neural: en-GB Sonia Natural
  const soniaNatural = voices.find(
    (v) => v.lang.startsWith('en-GB') && v.name.includes('Sonia') && v.name.includes('Natural'),
  );
  if (soniaNatural) {
    console.log('🔊 TTS voice:', soniaNatural.name);
    return soniaNatural;
  }

  // 2nd choice — any Edge "Natural" online voice (en-GB)
  const natural = voices.find(
    (v) => v.lang.startsWith('en-GB') && v.name.includes('Natural'),
  );
  if (natural) {
    console.log('🔊 TTS voice:', natural.name);
    return natural;
  }

  // 3rd choice — any British female (e.g. Google UK English Female)
  const britishFemale = voices.find(
    (v) =>
      v.lang.startsWith('en-GB') &&
      (/female|hazel|sonia|susan/i.test(v.name) || v.name.includes('(UK)')),
  );
  if (britishFemale) {
    console.log('🔊 TTS voice:', britishFemale.name);
    return britishFemale;
  }

  // 4th choice — any British voice
  const british = voices.find((v) => v.lang.startsWith('en-GB'));
  if (british) {
    console.log('🔊 TTS voice:', british.name);
    return british;
  }

  // 5th choice — any English voice
  const english = voices.find((v) => v.lang.startsWith('en'));
  if (english) {
    console.log('🔊 TTS voice:', english.name);
    return english;
  }

  console.log('🔊 No English TTS voice found, using system default');
  return null;
}

// ── Stop ────────────────────────────────────────────

/**
 * Stop any currently playing TTS audio and settle the pending promise.
 */
export function stopPlayback(): void {
  window.speechSynthesis.cancel();
  currentUtterance = null;
  fallbackResolve?.();
  fallbackResolve = null;
}

// ── Speak ───────────────────────────────────────────

/**
 * Speak a sentence using the browser's native speech synthesis.
 * Automatically selects a British English voice if available.
 *
 * @returns A promise that resolves when playback completes or is stopped.
 */
export async function speakSentence(
  text: string,
  _config?: TtsConfig,
): Promise<void> {
  // Kill any previous playback
  stopPlayback();

  if (!text.trim()) return;

  // Load voices (cached after first call)
  const voices = await loadVoices();
  const voice = pickVoice(voices);

  return new Promise<void>((resolve) => {
    const utterance = new SpeechSynthesisUtterance(text);

    utterance.rate = 0.9;           // slightly slower for dictation
    utterance.pitch = 1;
    utterance.volume = 1;
    utterance.lang = 'en-GB';

    if (voice) utterance.voice = voice;

    currentUtterance = utterance;
    fallbackResolve = resolve;

    utterance.onend = () => {
      currentUtterance = null;
      fallbackResolve = null;
      resolve();
    };

    utterance.onerror = () => {
      // Resolve anyway — don't let a speech glitch break the UX
      currentUtterance = null;
      fallbackResolve = null;
      resolve();
    };

    window.speechSynthesis.speak(utterance);
  });
}
