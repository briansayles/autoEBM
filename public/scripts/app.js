// filepath: /web-front-end/web-front-end/public/scripts/app.js

document.addEventListener('DOMContentLoaded', function() {
    const equipmentFileInput = document.getElementById('equipmentFileInput');
    const customerConfigFileInput = document.getElementById('customerConfigFileInput');
    const customerTemplateFileInput = document.getElementById('customerTemplateFileInput');
    const uploadButton = document.getElementById('uploadButton');
    const messagesDiv = document.getElementById('messages');
    // const customerNameSelect = document.getElementById('customerNameSelect');

    // Disable uploadButton initially
    uploadButton.disabled = true;

    function updateUploadButtonState() {
        const filesChosen = equipmentFileInput.files.length > 0 &&
                            customerConfigFileInput.files.length > 0 &&
                            customerTemplateFileInput.files.length > 0;
        // const customerChosen = customerNameSelect.value !== '' && customerNameSelect.value !== 'Select Customer';
        uploadButton.disabled = !(filesChosen);// && customerChosen);
    }

    equipmentFileInput.addEventListener('change', updateUploadButtonState);
    customerConfigFileInput.addEventListener('change', updateUploadButtonState);
    customerTemplateFileInput.addEventListener('change', updateUploadButtonState);
    // customerNameSelect.addEventListener('change', updateUploadButtonState);

    const selectedFileNameSpan = document.getElementById('selectedFileName');
    equipmentFileInput.addEventListener('change', function() {
        if (equipmentFileInput.files.length > 0) {
            selectedFileNameSpan.textContent = `Selected file: ${equipmentFileInput.files[0].name}`;
        } else {
            selectedFileNameSpan.textContent = '';
        }
    });

    const selectedConfigFileNameSpan = document.getElementById('selectedConfigFileName');
    customerConfigFileInput.addEventListener('change', function() {
        if (customerConfigFileInput.files.length > 0) {
            selectedConfigFileNameSpan.textContent = `Selected file: ${customerConfigFileInput.files[0].name}`;
        } else {
            selectedConfigFileNameSpan.textContent = '';
        }
    });

    const selectedTemplateFileNameSpan = document.getElementById('selectedTemplateFileName');
    customerTemplateFileInput.addEventListener('change', function() {
        if (customerTemplateFileInput.files.length > 0) {
            selectedTemplateFileNameSpan.textContent = `Selected file: ${customerTemplateFileInput.files[0].name}`;
        } else {
            selectedTemplateFileNameSpan.textContent = '';
        }
    });

    uploadButton.addEventListener('click', async function() {
        uploadButton.disabled = true;
        equipmentFileInput.disabled = true;
        customerConfigFileInput.disabled = true;
        customerTemplateFileInput.disabled = true;
        messagesDiv.textContent = 'Processing... Please wait.';
        const file = equipmentFileInput.files[0];
        if (!file) {
            messagesDiv.textContent = 'Please select a file to upload.';
            return;
        }
        const formData = new FormData();
        formData.append('equipmentDataFile', equipmentFileInput.files[0]);
        formData.append('customerConfigFile', document.getElementById('customerConfigFileInput').files[0]);
        formData.append('customerTemplateFile', document.getElementById('customerTemplateFileInput').files[0]);
        formData.append('jobNumber', document.getElementById('jobNumberInput').value);
        formData.append('noExcel', document.getElementById('noExcelCheckbox').checked);
        formData.append('noLabels', document.getElementById('noLabelsCheckbox').checked);
        formData.append('noMerge', document.getElementById('noMergeCheckbox').checked);
        try {
            const response = await fetch('/upload', {
                method: 'POST',
                body: formData,
                headers: {
                    'jobNumber': document.getElementById('jobNumberInput').value,
                    'noExcel': document.getElementById('noExcelCheckbox').checked,
                    'noLabels': document.getElementById('noLabelsCheckbox').checked,
                    'noMerge': document.getElementById('noMergeCheckbox').checked,
                    // 'customerName': customerNameSelect.value
                }
            });
            const result = await response.json();
            console.log(result);
            equipmentFileInput.disabled = false;
            customerConfigFileInput.disabled = false;
            customerTemplateFileInput.disabled = false;
            if (result.error) {
                updateUploadButtonState();
            }
            messagesDiv.textContent = `${result.message}`;
            messagesDiv.appendChild(document.createElement('br'));
            if (result.zipOutput) {
                const downloadLink = document.createElement('a');
                downloadLink.href = `/download?filePath=${encodeURIComponent(result.zipOutput)}`;
                downloadLink.textContent = 'AutoEBM Results File';
                messagesDiv.appendChild(document.createElement('br'));
                messagesDiv.appendChild(downloadLink);
            }
        } catch (error) {
            console.log('There was an error:'  , error);
            messagesDiv.textContent = `Error: ${error}`;
        }
        updateUploadButtonState();
    });

    const noLabelsCheckbox = document.getElementById('noLabelsCheckbox');
    const noMergeCheckbox = document.getElementById('noMergeCheckbox');

    noLabelsCheckbox.addEventListener('change', function () {
        if (noLabelsCheckbox.checked) {
            noMergeCheckbox.checked = true;
            noMergeCheckbox.disabled = true;
        } else {
            noMergeCheckbox.checked = false;
            noMergeCheckbox.disabled = false;
        }
    });
});