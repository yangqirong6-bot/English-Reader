import { NextResponse } from 'next/server';
import { fetchChineseTranslation } from '@/lib/dictionary';

export async function GET(request: Request) {
  const { searchParams } = new URL(request.url);
  const text = searchParams.get('text');

  if (!text) {
    return NextResponse.json({ error: 'Text is required' }, { status: 400 });
  }

  const trimmed = text.trim();
  if (trimmed.length === 0) {
    return NextResponse.json({ error: 'Text is empty' }, { status: 400 });
  }

  try {
    const result = await fetchChineseTranslation(trimmed);
    if (result) {
      return NextResponse.json({ translation: result.text, source: result.source });
    }
    return NextResponse.json({ error: 'Translation failed' }, { status: 404 });
  } catch (error) {
    console.error('Translation error:', error);
    return NextResponse.json({ error: 'Translation error' }, { status: 500 });
  }
}
