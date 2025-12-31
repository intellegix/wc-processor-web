/**
 * Workers' Comp Processor - Frontend Application
 */

document.addEventListener('DOMContentLoaded', function() {
    // State
    const state = {
        asrFile: null,
        armorproFile: null,
        payPeriod: null
    };

    // Elements
    const asrUploadZone = document.getElementById('asr-upload-zone');
    const asrFileInput = document.getElementById('asr-file');
    const armorproUploadZone = document.getElementById('armorpro-upload-zone');
    const armorproFileInput = document.getElementById('armorpro-file');
    const payPeriodInput = document.getElementById('pay-period');
    const processBtn = document.getElementById('process-btn');
    const resetBtn = document.getElementById('reset-btn');
    const retryBtn = document.getElementById('retry-btn');

    // Panels
    const initialState = document.getElementById('initial-state');
    const processingSteps = document.getElementById('processing-steps');
    const resultsPanel = document.getElementById('results-panel');
    const errorState = document.getElementById('error-state');

    // Set default pay period to today
    const today = new Date();
    payPeriodInput.value = today.toISOString().split('T')[0];

    // Setup upload zones
    setupUploadZone(asrUploadZone, asrFileInput, 'asr');
    setupUploadZone(armorproUploadZone, armorproFileInput, 'armorpro');

    // Event listeners
    payPeriodInput.addEventListener('change', updateProcessButton);
    processBtn.addEventListener('click', processReports);
    resetBtn?.addEventListener('click', resetForm);
    retryBtn?.addEventListener('click', resetForm);

    function setupUploadZone(zone, input, type) {
        // Click to upload
        zone.addEventListener('click', (e) => {
            if (!e.target.classList.contains('remove-file')) {
                input.click();
            }
        });

        // File input change
        input.addEventListener('change', (e) => {
            if (e.target.files.length > 0) {
                handleFileSelect(e.target.files[0], type, zone);
            }
        });

        // Drag and drop
        zone.addEventListener('dragover', (e) => {
            e.preventDefault();
            zone.classList.add('drag-over');
        });

        zone.addEventListener('dragleave', () => {
            zone.classList.remove('drag-over');
        });

        zone.addEventListener('drop', (e) => {
            e.preventDefault();
            zone.classList.remove('drag-over');
            if (e.dataTransfer.files.length > 0) {
                handleFileSelect(e.dataTransfer.files[0], type, zone);
            }
        });

        // Remove file button
        const removeBtn = zone.querySelector('.remove-file');
        removeBtn?.addEventListener('click', (e) => {
            e.stopPropagation();
            removeFile(type, zone, input);
        });
    }

    async function handleFileSelect(file, type, zone) {
        // Validate file type
        const validTypes = ['text/csv', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'application/vnd.ms-excel'];
        const validExtensions = ['.csv', '.xlsx', '.xls'];
        const ext = '.' + file.name.split('.').pop().toLowerCase();

        if (!validExtensions.includes(ext)) {
            alert('Please upload a CSV, XLSX, or XLS file.');
            return;
        }

        // Upload file
        const formData = new FormData();
        formData.append('file', file);
        formData.append('type', type);

        try {
            const response = await fetch('/api/upload', {
                method: 'POST',
                body: formData
            });

            const result = await response.json();

            if (result.success) {
                // Update state
                if (type === 'asr') {
                    state.asrFile = result.filename;
                } else {
                    state.armorproFile = result.filename;
                }

                // Update UI
                zone.classList.add('has-file');
                zone.querySelector('.upload-content').classList.add('d-none');
                zone.querySelector('.upload-success').classList.remove('d-none');
                zone.querySelector('.file-name').textContent = file.name;

                updateProcessButton();
            } else {
                alert('Error uploading file: ' + result.error);
            }
        } catch (error) {
            console.error('Upload error:', error);
            alert('Error uploading file. Please try again.');
        }
    }

    function removeFile(type, zone, input) {
        if (type === 'asr') {
            state.asrFile = null;
        } else {
            state.armorproFile = null;
        }

        zone.classList.remove('has-file');
        zone.querySelector('.upload-content').classList.remove('d-none');
        zone.querySelector('.upload-success').classList.add('d-none');
        input.value = '';

        updateProcessButton();
    }

    function updateProcessButton() {
        const hasAsrFile = state.asrFile !== null;
        const hasPayPeriod = payPeriodInput.value !== '';
        processBtn.disabled = !(hasAsrFile && hasPayPeriod);
    }

    async function processReports() {
        // Hide initial state, show processing
        initialState.classList.add('d-none');
        resultsPanel.classList.add('d-none');
        errorState.classList.add('d-none');
        processingSteps.classList.remove('d-none');

        // Disable process button
        processBtn.disabled = true;
        processBtn.innerHTML = '<span class="spinner-border spinner-border-sm me-2"></span>Processing...';

        // Reset all steps
        document.querySelectorAll('.step').forEach(step => {
            step.classList.remove('processing', 'complete', 'error', 'skipped');
            step.querySelector('.step-status').textContent = 'Waiting...';
        });

        try {
            // Format pay period as YYYYMMDD
            const payPeriodDate = new Date(payPeriodInput.value);
            const payPeriod = payPeriodDate.toISOString().split('T')[0].replace(/-/g, '');

            const response = await fetch('/api/process', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({
                    asr_file: state.asrFile,
                    armorpro_file: state.armorproFile,
                    pay_period: payPeriod
                })
            });

            const result = await response.json();

            if (result.success) {
                // Update steps based on results
                result.results.steps.forEach(step => {
                    const stepEl = document.querySelector(`.step[data-step="${step.step}"]`);
                    if (stepEl) {
                        stepEl.classList.remove('processing');
                        stepEl.classList.add(step.status);
                        stepEl.querySelector('.step-status').textContent =
                            step.status === 'complete' ? 'Complete' :
                            step.status === 'skipped' ? 'Skipped' :
                            step.status === 'warning' ? (step.message || 'Warning') :
                            'Error';
                    }
                });

                // Show results
                showResults(result.results);
            } else {
                showError(result.error);
            }
        } catch (error) {
            console.error('Processing error:', error);
            showError('An unexpected error occurred. Please try again.');
        }

        // Reset button
        processBtn.disabled = false;
        processBtn.innerHTML = '<i class="bi bi-gear-fill me-2"></i>Process Reports';
    }

    function showResults(results) {
        processingSteps.classList.add('d-none');
        resultsPanel.classList.remove('d-none');

        // Update summary cards
        if (results.summary) {
            document.getElementById('sum-records').textContent = results.summary.record_count || '--';
            document.getElementById('sum-regular').textContent = formatCurrency(results.summary.regular_wages);
            document.getElementById('sum-overtime').textContent = formatCurrency(results.summary.overtime_wages + results.summary.doubletime_wages);
            document.getElementById('sum-total').textContent = formatCurrency(results.summary.grand_total);
        }

        // Populate download list
        const downloadList = document.getElementById('download-list');
        downloadList.innerHTML = '';

        results.files.forEach(file => {
            const ext = file.type || file.name.split('.').pop();
            const iconClass = ext === 'xlsx' ? 'bi-file-earmark-excel xlsx' : 'bi-file-earmark-text csv';

            const item = document.createElement('a');
            item.href = `/api/download/${file.name}`;
            item.className = 'list-group-item list-group-item-action';
            item.innerHTML = `
                <div class="file-info">
                    <i class="bi ${iconClass} file-icon"></i>
                    <span>${file.name}</span>
                </div>
                <span class="badge bg-primary">Download</span>
            `;
            downloadList.appendChild(item);
        });
    }

    function showError(message) {
        processingSteps.classList.add('d-none');
        errorState.classList.remove('d-none');
        document.getElementById('error-message').textContent = message;
    }

    function resetForm() {
        // Reset state
        state.asrFile = null;
        state.armorproFile = null;

        // Reset upload zones
        [asrUploadZone, armorproUploadZone].forEach(zone => {
            zone.classList.remove('has-file');
            zone.querySelector('.upload-content').classList.remove('d-none');
            zone.querySelector('.upload-success').classList.add('d-none');
        });

        // Reset file inputs
        asrFileInput.value = '';
        armorproFileInput.value = '';

        // Reset pay period to today
        payPeriodInput.value = new Date().toISOString().split('T')[0];

        // Show initial state
        initialState.classList.remove('d-none');
        processingSteps.classList.add('d-none');
        resultsPanel.classList.add('d-none');
        errorState.classList.add('d-none');

        // Update button
        updateProcessButton();

        // Cleanup session files
        fetch('/api/cleanup', { method: 'POST' });
    }

    function formatCurrency(value) {
        if (value === undefined || value === null) return '$--';
        return '$' + value.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    }
});
