import { NextResponse } from 'next/server';
import { lookupWord } from '@/lib/dictionary';

export async function GET(request: Request) {
  const { searchParams } = new URL(request.url);
  const word = searchParams.get('word');

  if (!word) {
    return NextResponse.json({ error: 'Word is required' }, { status: 400 });
  }

  try {
    const result = await lookupWord(word.trim().toLowerCase());
    if (result) {
      return NextResponse.json(result);
    }
    return NextResponse.json({ error: 'Not found' }, { status: 404 });
  } catch (error) {
    console.error('Dictionary lookup error:', error);
    return NextResponse.json({ error: 'Lookup failed' }, { status: 500 });
  }
}
