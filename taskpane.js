import {
  fetchBibles,
  fetchBooks,
  fetchChapters,
  fetchVerses,
  fetchVerseById,
} from "./bibleApi.js";
import {
  clearDropdown,
  populateDropdown,
  clearCheckboxes,
  createVerseCheckbox,
} from "./ui.js";
import { formatVerseText } from "./helpers.js";

let bibleList = [];
let selectedBibleId = null;
let selectedBookId = null;

Office.onReady(() => {
  if (Office.context.host !== Office.HostType.PowerPoint) return;

  const bibleDropdown = document.getElementById("bible-dropdown");
  const bookDropdown = document.getElementById("book-dropdown");
  const chapterDropdown = document.getElementById("chapter-dropdown");
  const verseCheckboxesContainer = document.getElementById("verse-checkboxes");
  const insertVersesButton = document.getElementById("insert-verses");

  loadBibles();

  bibleDropdown.addEventListener("change", async () => {
    selectedBibleId = bibleDropdown.value;
    clearDropdown(bookDropdown);
    clearDropdown(chapterDropdown);
    clearCheckboxes(verseCheckboxesContainer);
    bookDropdown.disabled = true;
    chapterDropdown.disabled = true;
    await loadBooks(selectedBibleId);
  });

  bookDropdown.addEventListener("change", async () => {
    selectedBookId = bookDropdown.value;
    clearDropdown(chapterDropdown);
    clearCheckboxes(verseCheckboxesContainer);
    chapterDropdown.disabled = true;
    await loadChapters(selectedBibleId, selectedBookId);
  });

  chapterDropdown.addEventListener("change", async () => {
    clearCheckboxes(verseCheckboxesContainer);
    const chapterId = chapterDropdown.value;
    if (!chapterId) return;
    await loadVerses(selectedBibleId, chapterId, verseCheckboxesContainer);
  });

  insertVersesButton.addEventListener("click", async () => {
    const selectedVerseIds = Array.from(
      verseCheckboxesContainer.querySelectorAll("input[type='checkbox']:checked")
    ).map(cb => cb.value);

    if (selectedVerseIds.length === 0) {
      console.log("No verses selected.");
      return;
    }

    try {
      const versePromises = selectedVerseIds.map(id =>
        fetchVerseById(selectedBibleId, id)
      );
      const verses = await Promise.all(versePromises);

      // Format main and parallel versions
      const sfbFormatted = verses
        .map(v => `${formatVerseText(v.reference, v.baseText)}`)
        .join(" ");
      const fbvFormatted = verses
        .map(v => `${formatVerseText(v.reference, v.parallelText || "[Not available]")}`)
        .join(" ");

      // Generate clean combined reference title like: "Genesis 1:1–2"
      const firstVerse = verses[0];
      const lastVerse = verses[verses.length - 1];

      // Use reference string like "Genesis 1:1"
      const firstRefParts = firstVerse.reference.split(" ");
      const bookName = firstRefParts.slice(0, -1).join(" "); // handles multi-word books like "1 Thessalonians"
      const [firstChapter, firstVerseNumber] = firstRefParts.at(-1).split(":");

      const lastRefParts = lastVerse.reference.split(" ");
      const [lastChapter, lastVerseNumber] = lastRefParts.at(-1).split(":");

      const sameChapter = firstChapter === lastChapter;
      const baseRef = `${bookName} ${firstChapter}:${firstVerseNumber === lastVerseNumber
        ? firstVerseNumber
        : `${firstVerseNumber}${sameChapter ? `–${lastVerseNumber}` : `–${lastChapter}:${lastVerseNumber}`}`
      }`;

      const sfbTitle = `${baseRef} [SFB]`;
      const lastSpaceIndex = baseRef.lastIndexOf(" ");
      const chapterAndVerse = baseRef.substring(lastSpaceIndex + 1);
      const fbvTitle = `${verses[0].parallelBookID} ${chapterAndVerse} [BSB]`;

      const sfbText = `${sfbTitle}\n${sfbFormatted}`;
      const fbvText = `${fbvTitle}\n${fbvFormatted}`;

      // Insert textboxes on selected slide
      await PowerPoint.run(async (context) => {
        const selectedSlides = context.presentation.getSelectedSlides();
        selectedSlides.load("items");
        await context.sync();

        if (selectedSlides.items.length === 0) {
          console.log("No slide selected.");
          return;
        }

        const slide = selectedSlides.items[0];

      // Insert side-by-side textboxes
      const boxWidth = 400;
      const boxHeight = 500;
      const topMargin = 100;
      const leftMargin = 60;
      const spacing = 40;

      const mainShape = slide.shapes.addTextBox(sfbText, {
        left: leftMargin,
        top: topMargin,
        width: boxWidth,
        height: boxHeight,
      });

      const parallelShape = slide.shapes.addTextBox(fbvText, {
        left: leftMargin + boxWidth + spacing,
        top: topMargin,
        width: boxWidth,
        height: boxHeight,
      });

      mainShape.textFrame.font.size = 36;
      parallelShape.textFrame.font.size = 36;

        await context.sync();
        console.log("Inserted selected verses.");
      });
    } catch (error) {
      console.error("Error inserting verses:", error);
    }
  });

  // Load list of Bibles in Swedish (or selected language)
  async function loadBibles() {
    try {
      const data = await fetchBibles();
      bibleList = data.data;
      populateDropdown(bibleDropdown, bibleList, "Välj en Bibel");
    } catch (error) {
      console.error("Failed to load Bible list:", error);
      bibleDropdown.innerHTML = "<option>Error loading Bibles</option>";
    }
  }

  // Load books for selected Bible
  async function loadBooks(bibleId) {
    try {
      const data = await fetchBooks(bibleId);
      populateDropdown(bookDropdown, data.data, "Välj en bok");
    } catch (error) {
      console.error("Failed to load books:", error);
      bookDropdown.innerHTML = "<option>Error loading books</option>";
    }
  }

  // Load chapters for selected book
  async function loadChapters(bibleId, bookId) {
    try {
      const data = await fetchChapters(bibleId, bookId);
      populateDropdown(chapterDropdown, data.data, "Välj ett kapitel");
    } catch (error) {
      console.error("Failed to load chapters:", error);
      chapterDropdown.innerHTML = "<option>Error loading chapters</option>";
    }
  }

  // Load checkboxes for verses in selected chapter
  async function loadVerses(bibleId, chapterId, container) {
    try {
      const data = await fetchVerses(bibleId, chapterId);
      data.data.forEach((verse, index) => {
        const checkboxLabel = createVerseCheckbox(verse, index);
        container.appendChild(checkboxLabel);
        container.appendChild(document.createElement("br"));
      });
    } catch (error) {
      console.error("Failed to load verses:", error);
      container.textContent = "Error loading verses.";
    }
  }
});
