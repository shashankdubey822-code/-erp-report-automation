document.addEventListener('DOMContentLoaded', function() {
    const uploadZone = document.getElementById('uploadZone');
    const fileInput = document.getElementById('erp_file');
    const fileInfo = document.getElementById('fileInfo');
    const fileName = document.getElementById('fileName');
    const removeFile = document.getElementById('removeFile');
    const submitBtn = document.getElementById('submitBtn');
    const browseText = document.querySelector('.text-primary.fw-bold');
    // const spinner = submitBtn.querySelector('.spinner-border'); // Spinner handling might be better in a separate utility or integrated more cleanly

    if(uploadZone) {
        uploadZone.addEventListener('click', () => fileInput.click());
        browseText.addEventListener('click', (e) => {
            e.stopPropagation();
            fileInput.click();
        });

        uploadZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            uploadZone.classList.add('drag-over');
        });

        uploadZone.addEventListener('dragleave', () => {
            uploadZone.classList.remove('drag-over');
        });

        uploadZone.addEventListener('drop', (e) => {
            e.preventDefault();
            uploadZone.classList.remove('drag-over');
            const files = e.dataTransfer.files;
            if (files.length > 0) {
                handleFileSelection(files[0]);
            }
        });
    }

    if(fileInput) {
        fileInput.addEventListener('change', (e) => {
            if (e.target.files.length > 0) {
                handleFileSelection(e.target.files[0]);
            }
        });
    }


    function handleFileSelection(file) {
        const allowedExtensions = ['.csv', '.xls', '.xlsx'];
        const fileExtension = '.' + file.name.split('.').pop().toLowerCase();
        if (!allowedExtensions.includes(fileExtension)) {
            alert('Invalid file type. Please upload a CSV or Excel file.');
            return;
        }
        
        if (file.size > 16 * 1024 * 1024) {
            alert('File size must be less than 16MB.');
            return;
        }

        fileName.textContent = file.name;
        uploadZone.classList.add('is-hidden'); 
        fileInfo.classList.remove('is-hidden'); 
        submitBtn.disabled = false;
    }

    if(removeFile) {
        removeFile.addEventListener('click', () => {
            fileInput.value = '';
            uploadZone.classList.remove('is-hidden');
            fileInfo.classList.add('is-hidden');
            submitBtn.disabled = true;
        });
    }

    // This section animation for sections fading in - keeping it here for now,
    // could be moved to a separate general animation.js if needed.
    const sections = document.querySelectorAll('.text-center.mb-5, .my-5, .my-5.p-4.border.rounded.shadow-sm, .text-center.mt-5.py-4.border-top.text-muted');
    sections.forEach((section, index) => {
        setTimeout(() => {
            section.classList.add('fade-in-up', 'loaded');
        }, (index * 150) + 100);
    });

    // const uploadForm = document.querySelector('form'); // Not used directly here
});
