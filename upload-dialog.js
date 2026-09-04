let currentAttachments = [];
let currentMessage = null;
let availableTags = [];
let documentTypes = [];

document.addEventListener('DOMContentLoaded', async function () {
  await loadUploadData();
  setupEventListeners();
  await loadPaperlessData();
  setupPlusButtons();
});
// Add event listeners for plus buttons to create new correspondent/document type
function setupPlusButtons() {
  const addCorrespondentBtn = document.getElementById('addCorrespondentBtn');
  if (addCorrespondentBtn) {
    addCorrespondentBtn.addEventListener('click', async () => {
      await createNewCorrespondent();
    });
  }
  const addDocumentTypeBtn = document.getElementById('addDocumentTypeBtn');
  if (addDocumentTypeBtn) {
    addDocumentTypeBtn.addEventListener('click', async () => {
      await createNewDocumentType();
    });
  }

  // Listen for messages from popup windows
  window.addEventListener('message', handlePopupMessage);
}

async function createNewCorrespondent() {
  try {
    await createCenteredWindow(browser.runtime.getURL('create-correspondent.html'), 600, 600);
  } catch (error) {
    console.error('Error opening correspondent creation window:', error);
  }
}

async function createNewDocumentType() {
  try {
    await createCenteredWindow(browser.runtime.getURL('create-document-type.html'), 600, 600);
  } catch (error) {
    console.error('Error opening document type creation window:', error);
  }
}

function handlePopupMessage(event) {
  if (event.data.action === 'correspondentCreated' && event.data.success) {
    // Repopulate correspondents and select the new one
    repopulateCorrespondents().then(() => {
      if (event.data.correspondent) {
        const select = document.getElementById('correspondent');
        select.value = event.data.correspondent.id;
        showSuccess(`Correspondent "${event.data.correspondent.name}" created successfully!`);
      }
    });
  } else if (event.data.action === 'documentTypeCreated' && event.data.success) {
    // Repopulate document types and select the new one
    repopulateDocumentTypes().then(() => {
      if (event.data.documentType) {
        const select = document.getElementById('documentType');
        select.value = event.data.documentType.id;
        showSuccess(`Document type "${event.data.documentType.name}" created successfully!`);
      }
    });
  }
}

async function repopulateCorrespondents() {
  try {
    const settings = await getPaperlessSettings();
    if (!settings.paperlessUrl || !settings.paperlessToken) return;

    const response = await makePaperlessRequest('/api/correspondents/?page_size=1000', {}, settings);

    if (response.ok) {
      const data = await response.json();
      const correspondents = data.results.map(c => ({ id: c.id, name: c.name }));

      const select = document.getElementById('correspondent');
      const currentValue = select.value;

      // Clear existing options except the first one
      while (select.children.length > 1) {
        select.removeChild(select.lastChild);
      }

      // Add all correspondents
      correspondents.forEach(correspondent => {
        const option = document.createElement('option');
        option.value = correspondent.id;
        option.textContent = correspondent.name;
        select.appendChild(option);
      });

      // Restore selection if it still exists
      if (currentValue && select.querySelector(`option[value="${currentValue}"]`)) {
        select.value = currentValue;
      }
    }
  } catch (error) {
    console.error('Error repopulating correspondents:', error);
  }
}

async function repopulateDocumentTypes() {
  try {
    const settings = await getPaperlessSettings();
    if (!settings.paperlessUrl || !settings.paperlessToken) return;

    const response = await makePaperlessRequest('/api/document_types/?page_size=1000', {}, settings);

    if (response.ok) {
      const data = await response.json();
      const documentTypes = data.results.map(d => ({ id: d.id, name: d.name }));

      const select = document.getElementById('documentType');
      const currentValue = select.value;

      // Clear existing options except the first one
      while (select.children.length > 1) {
        select.removeChild(select.lastChild);
      }

      // Add all document types
      documentTypes.forEach(docType => {
        const option = document.createElement('option');
        option.value = docType.id;
        option.textContent = docType.name;
        select.appendChild(option);
      });

      // Restore selection if it still exists
      if (currentValue && select.querySelector(`option[value="${currentValue}"]`)) {
        select.value = currentValue;
      }
    }
  } catch (error) {
    console.error('Error repopulating document types:', error);
  }
}

async function loadUploadData() {
  try {
    const result = await browser.storage.local.get('currentUploadData');
    const uploadData = result.currentUploadData;

    if (!uploadData) {
      showError("No upload data found. Please try again.");
      return;
    }

    currentMessage = uploadData.message;
    currentAttachments = uploadData.attachments;

    // Populate email info
    document.getElementById('emailFrom').textContent = currentMessage.author;
    document.getElementById('emailSubject').textContent = currentMessage.subject;
    document.getElementById('emailDate').textContent = new Date(currentMessage.date).toLocaleDateString();

    // Populate file list
    const fileList = document.getElementById('fileList');
    currentAttachments.forEach(attachment => {
      const li = document.createElement('li');
      li.className = 'file-item';
      li.textContent = `📄 ${attachment.name} (${browser.messengerUtilities.formatFileSize(attachment.size)})`;
      fileList.appendChild(li);
    });

    // Set default title (first attachment name without extension)
    if (currentAttachments.length > 0) {
      const defaultTitle = currentAttachments[0].name.replace(/\.pdf$/i, '');
      document.getElementById('documentTitle').value = defaultTitle;
    }

    // Set default date to email date
    const emailDate = new Date(currentMessage.date);
    document.getElementById('documentDate').value = emailDate.toISOString().split('T')[0];

    // Show main content
    document.getElementById('loadingSection').style.display = 'none';
    document.getElementById('mainContent').style.display = 'block';

  } catch (error) {
    console.error('Error loading upload data:', error);
    showError('Error loading data: ' + error.message);
  }
}

async function loadPaperlessData() {
  try {
    // Load settings
    const settings = await getPaperlessSettings();

    // Fetch correspondents from Paperless-ngx API if settings are available
    let correspondents = [];
    if (settings.paperlessUrl && settings.paperlessToken) {
      try {
        const response = await makePaperlessRequest('/api/correspondents/?page_size=1000', {}, settings);
        if (response.ok) {
          const data = await response.json();
          // Store both name and id for each correspondent
          correspondents = data.results.map(c => ({ id: c.id, name: c.name }));
          // You can use 'correspondents' as needed here
          // Example: console.log(correspondents);
        }
      } catch (err) {
        console.error('Failed to fetch correspondents from Paperless-ngx:', err);
      }
    }

    if (correspondents.length > 0) {
      const correspondentSelect = document.getElementById('correspondent');
      correspondents.forEach(correspondent => {
        const option = document.createElement('option');
        option.value = correspondent.id;
        option.textContent = correspondent.name;
        correspondentSelect.appendChild(option);
      });

    }


    documentTypes = [];
    // Fetch document types from Paperless-ngx API if settings are available
    if (settings.paperlessUrl && settings.paperlessToken) {
      try {
        const response = await makePaperlessRequest('/api/document_types/?page_size=1000', {}, settings);
        if (response.ok) {
          const data = await response.json();
          // Store document types
          documentTypes = data.results.map(d => ({ id: d.id, name: d.name }));
        }
      } catch (err) {
        console.error('Failed to fetch document types from Paperless-ngx:', err);
      }
    }

    if (documentTypes.length > 0) {
      const docTypeSelect = document.getElementById('documentType');
      documentTypes.forEach(docType => {
        const option = document.createElement('option');
        option.value = docType.id;
        option.textContent = docType.name;
        docTypeSelect.appendChild(option);
      });
    }

availableTags = [];

if (settings.paperlessUrl && settings.paperlessToken) {
  try {
    const response = await makePaperlessRequest(
      '/api/tags/?page_size=1000',
      {},
      settings
    );

    if (response.ok) {
      const data = await response.json();

      availableTags = data.results.map(tag => ({
        id: tag.id,
        name: tag.name
      }));

      const tagsSelect = document.getElementById('tags');

      availableTags.forEach(tag => {
        const option = document.createElement('option');
        option.value = tag.id;
        option.textContent = tag.name;
        tagsSelect.appendChild(option);
      });
    }
  } catch (err) {
    console.error(
      'Failed to fetch tags from Paperless-ngx:',
      err
    );
  }
}

  } catch (error) {
    console.error('Error loading Paperless data:', error);
    // Continue without the data - it's not critical for basic upload
  }
}

function setupEventListeners() {
  // Form submission
  document.getElementById('uploadForm').addEventListener('submit', handleUpload);

  // Cancel button
  document.getElementById('cancelBtn').addEventListener('click', () => {
    window.close();
  });
}

async function handleUpload(event) {
  event.preventDefault();

  const uploadBtn = document.getElementById('uploadBtn');
  const originalText = setButtonLoading(uploadBtn, '⏳ Uploading...');

  try {
    clearMessages();

    // Collect form data
    const formData = new FormData(event.target);

    // Convert correspondent and document_type to integer IDs if present
    const correspondentValue = formData.get('correspondent');
    const correspondentId = correspondentValue ? parseInt(correspondentValue, 10) : undefined;
    const documentTypeValue = formData.get('document_type');
    const documentTypeId = documentTypeValue ? parseInt(documentTypeValue, 10) : undefined;

    // Convert selectedTags (names) to IDs using availableTags
     const tagsSelect = document.getElementById('tags');

     const tagIds = Array.from(tagsSelect.selectedOptions)
       .map(option => parseInt(option.value, 10))
       .filter(Number.isInteger);

    const uploadOptions = {};

    const title = formData.get('title');
    if (title) uploadOptions.title = title;

    if (correspondentId) uploadOptions.correspondent = correspondentId;

    if (documentTypeId) uploadOptions.document_type = documentTypeId;

    const created = formData.get('created');
    if (created) uploadOptions.created = created;

    if (tagIds.length > 0) uploadOptions.tags = tagIds;

    // Upload each attachment using background.js - let background handle all notifications
    for (const attachment of currentAttachments) {
      try {
        await browser.runtime.sendMessage({
          action: 'uploadWithOptions',
          messageData: currentMessage,
          attachmentData: attachment,
          uploadOptions: uploadOptions
        });
        // Background script handles all success/error notifications
      } catch (error) {
        console.error(`Error sending upload message for ${attachment.name}:`, error);
        // Even message sending errors will be rare, let background handle notifications
      }
    }

    // Show completion message and close dialog
    showSuccess(`Upload requests sent for ${currentAttachments.length} document(s). Check notifications for results.`);
    closeWindowWithDelay(2000);

  } catch (error) {
    console.error('Upload form error:', error);
    showError('Error processing upload form: ' + error.message);
  } finally {
    resetButtonLoading(uploadBtn, originalText);
  }
}
