document.addEventListener('DOMContentLoaded', function() {
    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('file-input');
    const fileNameDisplay = document.getElementById('file-name');
    const printModeCheckbox = document.getElementById('print_mode');
    const pdfOption = document.getElementById('pdf-option');

    if (dropZone && fileInput) {
        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
            dropZone.addEventListener(eventName, preventDefaults, false);
            document.body.addEventListener(eventName, preventDefaults, false);
        });

        function preventDefaults(e) {
            e.preventDefault();
            e.stopPropagation();
        }

        ['dragenter', 'dragover'].forEach(eventName => {
            dropZone.addEventListener(eventName, highlight, false);
        });

        ['dragleave', 'drop'].forEach(eventName => {
            dropZone.addEventListener(eventName, unhighlight, false);
        });

        function highlight(e) {
            dropZone.classList.add('dragover');
        }

        function unhighlight(e) {
            dropZone.classList.remove('dragover');
        }

        dropZone.addEventListener('drop', handleDrop, false);

        function handleDrop(e) {
            const dt = e.dataTransfer;
            const files = dt.files;

            if (files.length > 0) {
                const file = files[0];
                
                if (file.name.toLowerCase().endsWith('.docx')) {
                    fileInput.files = files;
                    updateFileName(file.name);
                } else {
                    showError('Please upload a .docx file');
                }
            }
        }

        fileInput.addEventListener('change', function(e) {
            if (this.files.length > 0) {
                updateFileName(this.files[0].name);
            }
        });

        function updateFileName(name) {
            if (fileNameDisplay) {
                fileNameDisplay.textContent = 'Selected: ' + name;
                fileNameDisplay.style.display = 'block';
            }
        }

        function showError(message) {
            const existingAlert = document.querySelector('.alert-error');
            if (existingAlert) {
                existingAlert.remove();
            }

            const alertDiv = document.createElement('div');
            alertDiv.className = 'alert alert-error';
            alertDiv.textContent = message;
            
            const container = document.querySelector('.container');
            const header = document.querySelector('header');
            container.insertBefore(alertDiv, header.nextSibling);

            setTimeout(() => {
                alertDiv.remove();
            }, 5000);
        }
    }

    if (printModeCheckbox && pdfOption) {
        printModeCheckbox.addEventListener('change', function() {
            if (this.checked) {
                pdfOption.style.display = 'block';
            } else {
                pdfOption.style.display = 'none';
                document.getElementById('generate_pdf').checked = false;
            }
        });
    }

    const form = document.getElementById('upload-form');
    if (form) {
        form.addEventListener('submit', function(e) {
            const fileInput = document.getElementById('file-input');
            
            if (!fileInput.files || fileInput.files.length === 0) {
                e.preventDefault();
                showError('Please select a .docx file to upload');
                return false;
            }

            const submitBtn = document.querySelector('.submit-btn');
            if (submitBtn) {
                submitBtn.innerHTML = `
                    <svg class="spinner" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <circle cx="12" cy="12" r="10"></circle>
                    </svg>
                    Processing...
                `;
                submitBtn.disabled = true;
                submitBtn.style.opacity = '0.7';

                const style = document.createElement('style');
                style.textContent = `
                    @keyframes spin {
                        0% { transform: rotate(0deg); }
                        100% { transform: rotate(360deg); }
                    }
                    .spinner {
                        animation: spin 1s linear infinite;
                        width: 24px;
                        height: 24px;
                    }
                `;
                document.head.appendChild(style);
            }

            return true;
        });

        function showError(message) {
            const existingAlert = document.querySelector('.alert-error');
            if (existingAlert) {
                existingAlert.remove();
            }

            const alertDiv = document.createElement('div');
            alertDiv.className = 'alert alert-error';
            alertDiv.textContent = message;
            
            const container = document.querySelector('.container');
            const header = document.querySelector('header');
            container.insertBefore(alertDiv, header.nextSibling);

            setTimeout(() => {
                alertDiv.remove();
            }, 5000);
        }
    }
});
