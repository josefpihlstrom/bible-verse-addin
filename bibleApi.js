const API_KEY = "9143bcf8e5fc5b5d169146ce13ef478b";
const BASE_URL = "https://api.scripture.api.bible/v1";
const FREE_BIBLE_ID = "bba9f40183526463-01"; // Berean standard Bible

async function fetchBibles() {
  const res = await fetch(`${BASE_URL}/bibles?language=swe`, {
    headers: { accept: "application/json", "api-key": API_KEY },
  });
  return res.json();
}

async function fetchBooks(bibleId) {
  const res = await fetch(`${BASE_URL}/bibles/${bibleId}/books`, {
    headers: { accept: "application/json", "api-key": API_KEY },
  });
  return res.json();
}

async function fetchChapters(bibleId, bookId) {
  const res = await fetch(`${BASE_URL}/bibles/${bibleId}/books/${bookId}/chapters`, {
    headers: { accept: "application/json", "api-key": API_KEY },
  });
  return res.json();
}

async function fetchVerses(bibleId, chapterId) {
  const res = await fetch(`${BASE_URL}/bibles/${bibleId}/chapters/${chapterId}/verses`, {
    headers: { accept: "application/json", "api-key": API_KEY },
  });
  return res.json();
}

async function fetchVerseById(bibleId, verseId) {
  const url = `${BASE_URL}/bibles/${bibleId}/verses/${verseId}` +
              `?content-type=html` +
              `&include-notes=false` +
              `&include-titles=true` +
              `&include-chapter-numbers=false` +
              `&include-verse-numbers=true` +
              `&include-verse-spans=false` +
              `&parallels=${FREE_BIBLE_ID}` +
              `&use-org-id=false`;

  const res = await fetch(url, {
    headers: { accept: "application/json", "api-key": API_KEY },
  });

  if (!res.ok) {
    throw new Error(`Failed to fetch verse: ${verseId}`);
  }

  const data = await res.json();

  return {
    reference: data.data.reference,
    baseText: data.data.content, // Main Bible selected
    parallelText: data.data.parallels[0].content, // Parallel verse
    parallelBookID : data.data.parallels[0].bookId, // Parallel book id
  };
}

export {
  fetchBibles,
  fetchBooks,
  fetchChapters,
  fetchVerses,
  fetchVerseById
};
