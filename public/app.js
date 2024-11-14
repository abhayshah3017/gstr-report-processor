document.addEventListener('DOMContentLoaded', () => {
  const form = document.getElementById('uploadForm');
  const fileInput = document.getElementById('file');
  const dropZone = document.querySelector('.drop-zone');
  const fileInfo = document.getElementById('fileInfo');
  const fileName = document.getElementById('fileName');
  const removeFile = document.getElementById('removeFile');
  const submitButton = document.getElementById('submitButton');
  const progress = document.getElementById('progress');
  const progressBar = document.querySelector('.progress-bar');

  // Drag and drop handlers
  dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('drag-over');
  });

  dropZone.addEventListener('dragleave', () => {
    dropZone.classList.remove('drag-over');
  });

  dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
    
    const files = e.dataTransfer.files;
    if (files.length > 0) {
      const file = files[0];
      if (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
        fileInput.files = files;
        updateFileInfo(file);
      } else {
        showError('Please upload an Excel file (.xlsx or .xls)');
      }
    }
  });

  // Click to upload
  dropZone.addEventListener('click', () => {
    fileInput.click();
  });

  fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
      updateFileInfo(e.target.files[0]);
    }
  });

  // Remove file
  removeFile.addEventListener('click', (e) => {
    e.stopPropagation();
    resetForm();
  });

  function updateFileInfo(file) {
    fileName.textContent = file.name;
    fileInfo.classList.remove('hidden');
    submitButton.disabled = false;
  }

  function resetForm() {
    fileInput.value = '';
    fileInfo.classList.add('hidden');
    submitButton.disabled = true;
    progress.classList.add('hidden');
    progressBar.style.width = '0%';
  }

  function showError(message) {
    const error = document.createElement('div');
    error.className = 'error';
    error.textContent = message;
    form.appendChild(error);
    setTimeout(() => error.remove(), 5000);
  }

  // Form submission
  form.addEventListener('submit', async (e) => {
    e.preventDefault();
    
    const formData = new FormData();
    formData.append('file', fileInput.files[0]);

    submitButton.disabled = true;
    progress.classList.remove('hidden');
    progressBar.style.width = '50%';

    try {
      const response = await fetch('/process', {
        method: 'POST',
        body: formData
      });

      progressBar.style.width = '100%';

      if (response.ok) {
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'processed.pdf';
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        a.remove();
        resetForm();
      } else {
        const error = await response.json();
        showError(error.error || 'Error processing file');
      }
    } catch (error) {
      showError('Error uploading file');
    } finally {
      submitButton.disabled = false;
      setTimeout(() => {
        progress.classList.add('hidden');
        progressBar.style.width = '0%';
      }, 1000);
    }
  });
});