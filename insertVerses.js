import { formatVerseText } from "./helpers.js";

// Helper to build PowerPoint-formatted rich text parts
function buildFormattedText(context, title, verseText) {
  const lines = [`${title}`, ...verseText.split(/(?=\[\d+\])/)];
  const paragraph = context.presentation.text.createTextRange("");
  let textRange = paragraph;

  lines.forEach((line, i) => {
    const isTitle = i === 0;
    const range = context.presentation.text.createTextRange(line + "\n");
    if (isTitle) {
      range.font.bold = true;
    } else {
      // Superscript verse markers
      const match = line.match(/^\[(\d+)\]/);
      if (match) {
        const verseNum = match[0];
        const rest = line.substring(verseNum.length);

        const supRange = context.presentation.text.createTextRange(verseNum);
        supRange.font.superscript = true;

        const textRangeRest = context.presentation.text.createTextRange(rest + "\n");

        range.text = ""; // reset to append parts
        range.insertTextRange("Start", supRange);
        range.insertTextRange("End", textRangeRest);
      }
    }

    textRange.insertTextRange("End", range);
  });

  return textRange;
}

export async function insertVersesToSlide(context, slide, sfbTitle, sfbFormatted, fbvTitle, fbvFormatted) {
  const boxWidth = 400;
  const boxHeight = 500;
  const topMargin = 100;
  const leftMargin = 60;
  const spacing = 40;

  const shapes = slide.shapes;

  // Remove previous verse boxes if needed
  shapes.load("items");
  await context.sync();

  shapes.items.forEach(shape => shape.load("text"));
  await context.sync();

  for (const shape of shapes.items) {
    if (shape.text && (shape.text.includes("[SFB]") || shape.text.includes("[BSB]"))) {
      shape.delete();
    }
  }

  // Insert new textboxes
  const sfbShape = shapes.addTextBox(sfbTitle + "\n" + sfbFormatted, {
    left: leftMargin,
    top: topMargin,
    width: boxWidth,
    height: boxHeight,
  });

  const fbvShape = shapes.addTextBox(fbvTitle + "\n" + fbvFormatted, {
    left: leftMargin + boxWidth + spacing,
    top: topMargin,
    width: boxWidth,
    height: boxHeight,
  });

  // Load their text ranges
  sfbShape.textFrame.textRange.load("text");
  fbvShape.textFrame.textRange.load("text");
  await context.sync();

  // Set font styles
  [sfbShape, fbvShape].forEach(shape => {
    const textRange = shape.textFrame.textRange;
    textRange.font.size = 36;

    // Bold the first line (title)
    const titleEndIndex = textRange.text.indexOf("\n");
    if (titleEndIndex !== -1) {
      const titleRange = textRange.getSubstring(0, titleEndIndex);
      titleRange.font.bold = true;
    }

    // Superscript verse markers like [1], [2], etc.
    const verseMatches = [...textRange.text.matchAll(/\[\d+\]/g)];
    verseMatches.forEach(match => {
      const range = textRange.getSubstring(match.index, match.index + match[0].length);
      range.font.superscript = true;
    });
  });

  await context.sync();
}
