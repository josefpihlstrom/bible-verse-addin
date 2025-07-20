export async function insertVersesToSlide(verses) {
  const firstVerse = verses[0];
  const lastVerse = verses[verses.length - 1];

  const firstRefParts = firstVerse.reference.split(" ");
  const bookName = firstRefParts.slice(0, -1).join(" ");
  const [firstChapter, firstVerseNumber] = firstRefParts.at(-1).split(":");

  const lastRefParts = lastVerse.reference.split(" ");
  const [lastChapter, lastVerseNumber] = lastRefParts.at(-1).split(":");

  const sameChapter = firstChapter === lastChapter;
  const baseRef = `${bookName} ${firstChapter}:${firstVerseNumber === lastVerseNumber
    ? firstVerseNumber
    : `${firstVerseNumber}${sameChapter ? `–${lastVerseNumber}` : `–${lastChapter}:${lastVerseNumber}`}`
  }`;

  const sfbTitle = `${baseRef} [SFB]`;
  const chapterAndVerse = baseRef.substring(baseRef.lastIndexOf(" ") + 1);
  const fbvTitle = `${verses[0].parallelBookID} ${chapterAndVerse} [BSB]`;

  // Format verse text with superscript markers and new lines
  const sfbFormatted = verses.map(v =>
    `${getSuperscriptMarker(v.reference)} ${v.baseText}`
  ).join("\n");

  const fbvFormatted = verses.map(v =>
    `${getSuperscriptMarker(v.reference)} ${v.parallelText || "[Not available]"}`
  ).join("\n");

  await PowerPoint.run(async (context) => {
    const selectedSlides = context.presentation.getSelectedSlides();
    selectedSlides.load("items");
    await context.sync();

    if (selectedSlides.items.length === 0) {
      console.log("No slide selected.");
      return;
    }

    const slide = selectedSlides.items[0];

    const boxWidth = 400;
    const boxHeight = 500;
    const topMargin = 100;
    const leftMargin = 60;
    const spacing = 40;

    const sfbTextbox = slide.shapes.addTextBox(`${sfbTitle}\n${sfbFormatted}`, {
      left: leftMargin,
      top: topMargin,
      width: boxWidth,
      height: boxHeight,
    });

    const fbvTextbox = slide.shapes.addTextBox(`${fbvTitle}\n${fbvFormatted}`, {
      left: leftMargin + boxWidth + spacing,
      top: topMargin,
      width: boxWidth,
      height: boxHeight,
    });

    // Style: bold title, font size 36, superscript verse numbers
    [sfbTextbox, fbvTextbox].forEach(textbox => {
      const paragraphs = textbox.textFrame.textRange.paragraphs;
      const lines = textbox.textFrame.textRange.text.split("\n");

      textbox.textFrame.textRange.font.size = 36;

      if (lines.length > 0) {
        paragraphs.getItemAt(0).font.bold = true;
      }

      // Apply superscript to verse markers like [1], [2], etc.
      for (let i = 1; i < paragraphs.items.length; i++) {
        const para = paragraphs.getItemAt(i);
        const matches = [...para.text.matchAll(/\[(\d+)\]/g)];
        for (const match of matches) {
          const index = match.index;
          const length = match[0].length;
          const range = para.getSubstring(index, length);
          range.font.superscript = true;
        }
      }
    });

    await context.sync();
  });
}

function getSuperscriptMarker(reference) {
  const versePart = reference.split(":")[1];
  const verseNumber = versePart.split("–")[0];
  return `[${verseNumber}]`;
}
