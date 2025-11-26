// Email Upload Dialog Logic for Paperless-ngx PDF Uploader

let currentMessage = null;
let allAttachments = [];
let selectedTags = [];
let availableTags = [];
let fuse = null;
let selectedSuggestionIndex = -1;

document.addEventListener('DOMContentLoaded', async function () {
  await loadEmailData();
  setupEventListeners();
  await loadPaperlessData();
});

async function loadEmailData() {
  try {
    const result = await browser.storage.local.get('emailUploadData');
    const uploadData = result.emailUploadData;

    if (!uploadData) {
      showError('Keine E-Mail-Daten gefunden. Bitte versuchen Sie es erneut.');
      return;
    }

    currentMessage = uploadData.message;
    allAttachments = uploadData.attachments || [];

    // Populate email preview
    document.getElementById('emailFrom').textContent = currentMessage.author || '';
    document.getElementById('emailTo').textContent = currentMessage.recipients ? currentMessage.recipients.join(', ') : '';
    document.getElementById('emailSubject').textContent = currentMessage.subject || '(Kein Betreff)';
    document.getElementById('emailDate').textContent = new Date(currentMessage.date).toLocaleDateString('de-DE', {
      weekday: 'long',
      year: 'numeric',
      month: 'long',
      day: 'numeric',
      hour: '2-digit',
      minute: '2-digit'
    });

    // Set default title to email subject
    document.getElementById('documentTitle').value = currentMessage.subject || 'E-Mail';

    // Populate attachment list
    await populateAttachmentList();

    // Show main content
    document.getElementById('loadingSection').style.display = 'none';
    document.getElementById('mainContent').style.display = 'block';

  } catch (error) {
    console.error('Error loading email data:', error);
    showError('Fehler beim Laden der E-Mail-Daten: ' + error.message);
  }
}

async function populateAttachmentList() {
  const listContainer = document.getElementById('attachmentList');
  const noAttachmentsMsg = document.getElementById('noAttachments');
  const selectionControls = document.getElementById('selectionControls');

  // Clear existing items
  while (listContainer.firstChild) {
    listContainer.removeChild(listContainer.firstChild);
  }

  if (allAttachments.length === 0) {
    noAttachmentsMsg.style.display = 'block';
    selectionControls.style.display = 'none';
    return;
  }

  noAttachmentsMsg.style.display = 'none';
  selectionControls.style.display = 'flex';

  for (const [index, attachment] of allAttachments.entries()) {
    const item = document.createElement('div');
    item.className = 'attachment-item';

    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.className = 'attachment-checkbox';
    checkbox.id = `attachment-${index}`;
    checkbox.dataset.index = index;
    checkbox.checked = true; // Default: all selected

    const infoDiv = document.createElement('div');
    infoDiv.className = 'attachment-info';

    const icon = getAttachmentIcon(attachment.name, attachment.contentType);

    const nameDiv = document.createElement('div');
    nameDiv.className = 'attachment-name';
    nameDiv.innerHTML = `<span class="attachment-icon">${icon}</span>${escapeHtmlSafe(attachment.name)}`;

    const sizeDiv = document.createElement('div');
    sizeDiv.className = 'attachment-size';
    // Format file size
    const sizeText = await browser.messengerUtilities.formatFileSize(attachment.size);
    sizeDiv.textContent = sizeText;

    infoDiv.appendChild(nameDiv);
    infoDiv.appendChild(sizeDiv);
    item.appendChild(checkbox);
    item.appendChild(infoDiv);
    listContainer.appendChild(item);
  }
}

async function loadPaperlessData() {
  try {
    const settings = await getPaperlessSettings();
    if (!settings.paperlessUrl || !settings.paperlessToken) {
      return;
    }

    // Load correspondents
    try {
      const response = await makePaperlessRequest('/api/correspondents/', {}, settings);
      if (response.ok) {
        const data = await response.json();
        const correspondents = data.results.map(c => ({ id: c.id, name: c.name }));
        const select = document.getElementById('correspondent');
        correspondents.forEach(correspondent => {
          const option = document.createElement('option');
          option.value = correspondent.id;
          option.textContent = correspondent.name;
          select.appendChild(option);
        });
      }
    } catch (err) {
      console.error('Failed to fetch correspondents:', err);
    }

    // Load document types
    try {
      const response = await makePaperlessRequest('/api/document_types/', {}, settings);
      if (response.ok) {
        const data = await response.json();
        const documentTypes = data.results.map(d => ({ id: d.id, name: d.name }));
        const select = document.getElementById('documentType');
        documentTypes.forEach(docType => {
          const option = document.createElement('option');
          option.value = docType.id;
          option.textContent = docType.name;
          select.appendChild(option);
        });
      }
    } catch (err) {
      console.error('Failed to fetch document types:', err);
    }

    // Load tags
    try {
      const response = await makePaperlessRequest('/api/tags/', {}, settings);
      if (response.ok) {
        const data = await response.json();
        availableTags = data.results.map(t => ({ id: t.id, name: t.name }));
      }
    } catch (err) {
      console.error('Failed to fetch tags:', err);
    }

  } catch (error) {
    console.error('Error loading Paperless data:', error);
  }
}

function setupEventListeners() {
  // Form submission
  document.getElementById('uploadForm').addEventListener('submit', handleUpload);

  // Cancel button
  document.getElementById('cancelBtn').addEventListener('click', () => {
    window.close();
  });

  // Select all/none buttons
  document.getElementById('selectAllBtn').addEventListener('click', () => {
    const checkboxes = document.querySelectorAll('.attachment-checkbox');
    checkboxes.forEach(cb => cb.checked = true);
  });

  document.getElementById('selectNoneBtn').addEventListener('click', () => {
    const checkboxes = document.querySelectorAll('.attachment-checkbox');
    checkboxes.forEach(cb => cb.checked = false);
  });

  // Tags input
  const tagInput = document.querySelector('.tag-input');
  tagInput.addEventListener('keydown', handleTagInput);
  tagInput.addEventListener('input', handleTagAutocomplete);

  // Hide suggestions when clicking outside
  document.addEventListener('click', function (event) {
    const tagsContainer = document.getElementById('tagsInput');
    if (!tagsContainer.contains(event.target)) {
      hideSuggestions();
    }
  });
}

function handleTagInput(event) {
  if (event.key === 'Enter') {
    event.preventDefault();

    const suggestions = document.querySelectorAll('.suggestion-item');
    if (selectedSuggestionIndex >= 0 && suggestions[selectedSuggestionIndex]) {
      const selectedTag = suggestions[selectedSuggestionIndex].textContent;
      addTag(selectedTag);
      event.target.value = '';
      hideSuggestions();
      return;
    }

    const tagValue = event.target.value.trim();
    if (tagValue && !selectedTags.includes(tagValue)) {
      addTag(tagValue);
      event.target.value = '';
      hideSuggestions();
    }
  } else if (event.key === 'Backspace' && event.target.value === '') {
    if (selectedTags.length > 0) {
      removeTag(selectedTags[selectedTags.length - 1]);
    }
  } else if (event.key === 'ArrowDown') {
    event.preventDefault();
    navigateSuggestions(1);
  } else if (event.key === 'ArrowUp') {
    event.preventDefault();
    navigateSuggestions(-1);
  } else if (event.key === 'Escape') {
    hideSuggestions();
  }
}

function handleTagAutocomplete(event) {
  const query = event.target.value.trim();

  if (query.length === 0) {
    hideSuggestions();
    return;
  }

  if (!fuse && availableTags.length > 0) {
    const options = {
      includeScore: true,
      threshold: 0.4,
      keys: ['name']
    };
    fuse = new Fuse(availableTags, options);
  }

  if (fuse) {
    const results = fuse.search(query);
    showSuggestions(results.map(result => result.item), query);
  }
}

function addTag(tagName) {
  if (!selectedTags.includes(tagName)) {
    selectedTags.push(tagName);
    renderTags();
  }
}

function removeTag(tagName) {
  selectedTags = selectedTags.filter(tag => tag !== tagName);
  renderTags();
}

function renderTags() {
  const tagsContainer = document.getElementById('tagsInput');
  const tagInput = tagsContainer.querySelector('.tag-input');

  // Remove existing tag elements
  tagsContainer.querySelectorAll('.tag-item').forEach(el => el.remove());

  // Add tag elements
  selectedTags.forEach(tag => {
    const tagElement = document.createElement('div');
    tagElement.className = 'tag-item';

    const tagText = document.createTextNode(tag);
    tagElement.appendChild(tagText);

    const removeButton = document.createElement('span');
    removeButton.className = 'tag-remove';
    removeButton.textContent = '×';
    removeButton.addEventListener('click', () => removeTag(tag));

    tagElement.appendChild(removeButton);
    tagsContainer.insertBefore(tagElement, tagInput);
  });
}

function showSuggestions(tags, query) {
  hideSuggestions();

  if (tags.length === 0) return;

  const suggestionsContainer = document.getElementById('tagSuggestions');
  selectedSuggestionIndex = -1;

  const filteredTags = tags.filter(tag => !selectedTags.includes(tag.name));

  if (filteredTags.length === 0) return;

  const tagsToShow = filteredTags.slice(0, 5);

  tagsToShow.forEach((tag, index) => {
    const suggestionItem = document.createElement('div');
    suggestionItem.className = 'suggestion-item';

    const regex = new RegExp(`(${escapeRegexStr(query)})`, 'gi');
    const parts = tag.name.split(regex);

    parts.forEach(part => {
      if (part.toLowerCase() === query.toLowerCase()) {
        const mark = document.createElement('mark');
        mark.textContent = part;
        suggestionItem.appendChild(mark);
      } else {
        suggestionItem.appendChild(document.createTextNode(part));
      }
    });

    suggestionItem.addEventListener('click', () => {
      addTag(tag.name);
      const tagInput = document.querySelector('.tag-input');
      tagInput.value = '';
      hideSuggestions();
      tagInput.focus();
    });

    suggestionsContainer.appendChild(suggestionItem);
  });

  suggestionsContainer.style.display = 'block';
}

function hideSuggestions() {
  const suggestionsContainer = document.getElementById('tagSuggestions');
  while (suggestionsContainer.firstChild) {
    suggestionsContainer.removeChild(suggestionsContainer.firstChild);
  }
  suggestionsContainer.style.display = 'none';
  selectedSuggestionIndex = -1;
}

function navigateSuggestions(direction) {
  const suggestions = document.querySelectorAll('.suggestion-item');
  if (suggestions.length === 0) return;

  if (selectedSuggestionIndex >= 0) {
    suggestions[selectedSuggestionIndex].classList.remove('selected');
  }

  selectedSuggestionIndex += direction;

  if (selectedSuggestionIndex < 0) {
    selectedSuggestionIndex = suggestions.length - 1;
  } else if (selectedSuggestionIndex >= suggestions.length) {
    selectedSuggestionIndex = 0;
  }

  suggestions[selectedSuggestionIndex].classList.add('selected');
}

function escapeRegexStr(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function escapeHtmlSafe(str) {
  if (!str) return '';
  const div = document.createElement('div');
  div.textContent = str;
  return div.innerHTML;
}

async function handleUpload(event) {
  event.preventDefault();

  const uploadBtn = document.getElementById('uploadBtn');
  const originalText = uploadBtn.textContent;
  uploadBtn.disabled = true;
  uploadBtn.textContent = '⏳ Wird hochgeladen...';

  try {
    clearMessages();

    // Collect selected attachments
    const selectedAttachments = [];
    const checkboxes = document.querySelectorAll('.attachment-checkbox:checked');
    checkboxes.forEach(cb => {
      const index = parseInt(cb.dataset.index);
      selectedAttachments.push(allAttachments[index]);
    });

    // Collect form data
    const formData = new FormData(event.target);
    
    // Convert correspondent and document_type to integer IDs if present
    const correspondentValue = formData.get('correspondent');
    const correspondentId = correspondentValue ? parseInt(correspondentValue, 10) : undefined;
    const documentTypeValue = formData.get('document_type');
    const documentTypeId = documentTypeValue ? parseInt(documentTypeValue, 10) : undefined;

    // Convert selectedTags (names) to IDs using availableTags
    const tagIds = selectedTags
      .map(tagName => {
        const found = availableTags.find(t => t.name === tagName);
        return found ? found.id : undefined;
      })
      .filter(id => id !== undefined);

    const uploadOptions = {
      title: formData.get('title') || currentMessage.subject || 'E-Mail',
      correspondent: correspondentId,
      document_type: documentTypeId,
      tags: tagIds,
      addPaperlessTag: document.getElementById('addPaperlessTag').checked
    };

    // Send upload request to background script
    const response = await browser.runtime.sendMessage({
      action: 'uploadEmailToPaperless',
      messageData: currentMessage,
      selectedAttachments: selectedAttachments,
      uploadOptions: uploadOptions
    });

    if (response && response.success) {
      showSuccess('E-Mail wurde erfolgreich an Paperless-ngx gesendet!');
      setTimeout(() => window.close(), 2000);
    } else {
      showError('Fehler beim Hochladen: ' + (response?.error || 'Unbekannter Fehler'));
    }

  } catch (error) {
    console.error('Upload error:', error);
    showError('Fehler beim Hochladen: ' + error.message);
  } finally {
    uploadBtn.disabled = false;
    uploadBtn.textContent = originalText;
  }
}

// Make removeTag available globally for the tag elements
window.removeTag = removeTag;
