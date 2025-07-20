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
import { insertVersesToSlide } from "./insertVerses.js";

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

      await insertVersesToSlide(verses);
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
