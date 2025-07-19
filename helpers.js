// Strips HTML tags and decodes entities using a temporary DOM element
function stripHtml(html) {
  const tmp = document.createElement("div");
  tmp.innerHTML = html;
  return tmp.textContent || tmp.innerText || "";
}

// Formats the verse text by removing extra content and applying [verseNumber]
function formatVerseText(verseReference, verseText) {
  const plainText = stripHtml(verseText);

  const match = verseReference.match(/\d+$/);
  const verseNumber = match ? match[0] : "";

  if (!verseNumber) return plainText.trim();

  const index = plainText.indexOf(verseNumber);

  let cleanedText = plainText;

  if (index > 0) {
    cleanedText = plainText.substring(index + 1);
  } else if (index === 0) {
    cleanedText = plainText.substring(verseNumber.length);
  }

  cleanedText = cleanedText.trim();

  return `[${verseNumber}] ${cleanedText}`;
}

export { formatVerseText, stripHtml };
