import { formatVerseText } from "./helpers.js";

// Helper to build PowerPoint-formatted rich text parts
function buildFormattedText(context, title, verses) {
  const lines = [`${title}`, ...verses.map(v => formatVerseText(v.reference, v.baseText))];
  const paragraph = context.presentation.text.createTextRange("");
  let textRange = paragraph;

  lines.forEach((line, i) => {
    const isTitle = i === 0;
    const range = context.presentation.text.createTextRange(line + "\n");

    if (isTitle) {
      range.font.bold = true;
    } else {
      // Superscript verse markers like [1]
      const match = line.match(/^\[(\d+)\]/);
      if (match) {
        const verseNum = match[0];
        const rest = line.substring(verseNum.length);

        const supRange = context.presentation.text.createTextRange(verseNum);
        supRange.font.superscript = true;

        const restRange = context.presentation.text.createTextRange(rest + "\n");

        range.text = ""; // Clear and rebuild
        range.insertTextRange("Start", supRange);
        range.insertTextRange("End", restRange);
      }
    }

    textRange.insertTextRange("End", range);
  });

  return textRange;
}

export async function insertVersesToSlide(verses) {
  await PowerPoint.run(async context => {
    const presentation = context.presentation;
    const slide = presentation.slides.getActiveSlide();

    const boxWidth = 400;
    const boxHeight = 500;
    const topMargin = 100;
    const leftMargin = 60;
    const spacing = 40;

    const shapes = slide.shapes;
    shapes.load("items");
    await context.sync();

    shapes.items.forEach(shape => shape.load("textFrame/textRange/text"));
    await context.sync();

    for (const shape of shapes.items) {
      if (shape.textFrame?.textRange?.text?.includes("[SFB]") || shape.textFrame?.textRange?.text?.includes("[BSB]")) {
        shape.delete();
      }
    }

    // Filter verses
    const sfbVerses = verses.map(v => ({
      reference: v.reference,
      baseText: v.baseText,
    }));

    const fbvVerses = verses.map(v => ({
      reference: v.reference,
      baseText: v.parallelText || "[Not available]",
    }));

    const sfbTitle = `${verses[0].reference.split(":")[0]} [SFB]`;
    const fbvTitle = `${verses[0].reference.split(":")[0]} [BSB]`;

    const sfbShape = shapes.addTextBox("", {
      left: leftMargin,
      top: topMargin,
      width: boxWidth,
      height: boxHeight,
    });

    const fbvShape = shapes.addTextBox("", {
      left: leftMargin + boxWidth + spacing,
      top: topMargin,
      width: boxWidth,
      height: boxHeight,
    });

    // Load text ranges
    sfbShape.textFrame.textRange.load("text");
    fbvShape.textFrame.textRange.load("text");
    await context.sync();

    // Apply rich formatting
    const sfbFormatted = buildFormattedText(context, sfbTitle, sfbVerses);
    const fbvFormatted = buildFormattedText(context, fbvTitle, fbvVerses);

    sfbShape.textFrame.textRange.text = "";
    fbvShape.textFrame.textRange.text = "";

    sfbShape.textFrame.textRange.insertTextRange("Start", sfbFormatted);
    fbvShape.textFrame.textRange.insertTextRange("Start", fbvFormatted);

    await context.sync();
  });
}
