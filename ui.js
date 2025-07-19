function clearDropdown(dropdown) {
  dropdown.innerHTML = "";
}

function populateDropdown(dropdown, items, placeholder) {
  dropdown.innerHTML = "";
  const defaultOption = document.createElement("option");
  defaultOption.textContent = placeholder;
  defaultOption.disabled = false;
  defaultOption.selected = true;
  dropdown.appendChild(defaultOption);

  items.forEach((item) => {
    const option = document.createElement("option");
    option.value = item.id;
    option.textContent = item.name || item.reference;
    dropdown.appendChild(option);
  });

  dropdown.disabled = false;
}

function clearCheckboxes(container) {
  container.innerHTML = "";
}

function createVerseCheckbox(verse, index) {
  const label = document.createElement("label");
  const checkbox = document.createElement("input");
  checkbox.type = "checkbox";
  checkbox.value = verse.id;
  label.appendChild(checkbox);
  label.append(`Verse ${index + 1}`);
  return label;
}

export { clearDropdown, populateDropdown, clearCheckboxes, createVerseCheckbox };
