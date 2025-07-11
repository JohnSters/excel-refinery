/**
 * ExcelRefinery View Reports JavaScript
 * File: view-reports.js
 * Description: Interactive functionality for the reports dashboard and data analysis interface
 * Author: ExcelRefinery Development Team
 */

(function() {
    'use strict';
    
    // Cache DOM elements
    let cachedElements = {};
    
    // Application state
    let appState = {
        currentFileId: null,
        currentFileName: null,
        currentWorksheet: null,
        currentKeyColumn: null,
        isLoading: false,
        analysisData: null,
        uploadedFiles: new Map()
    };
    
    // Private methods
    const initializeElements = function() {
        try {
            cachedElements = {
                // File and worksheet selectors
                fileSelector: document.getElementById('fileSelector'),
                worksheetSelector: document.getElementById('worksheetSelector'),
                keyColumnSelectorContainer: document.getElementById('keyColumnSelectorContainer'),
                keyColumnSelector: document.getElementById('keyColumnSelector'),
                
                // Action buttons
                refreshDataBtn: document.getElementById('refreshDataBtn'),
                loadFileBtn: document.getElementById('loadFileBtn'),
                
                // Summary cards
                totalAssetsCount: document.getElementById('totalAssetsCount'),
                integrityIssuesCount: document.getElementById('integrityIssuesCount'),
                missingTasksCount: document.getElementById('missingTasksCount'),
                completionRate: document.getElementById('completionRate'),
                
                // Navigation tabs
                tabItems: document.querySelectorAll('.nav-excel-item[data-tab]'),
                tabPanes: document.querySelectorAll('.tab-pane'),
                
                // Tables
                integrityTable: document.getElementById('integrityTable'),
                assetsTable: document.getElementById('assetsTable'),
                tasksTable: document.getElementById('tasksTable'),
                
                // Loading overlay
                loadingOverlay: document.getElementById('loadingOverlay'),
                
                // Upload modal elements
                uploadModal: document.getElementById('uploadModal'),
                uploadModalClose: document.getElementById('uploadModalClose'),
                uploadZoneModal: document.getElementById('uploadZoneModal'),
                uploadContentModal: document.getElementById('uploadContentModal'),
                fileInputModal: document.getElementById('fileInputModal'),
                uploadProgress: document.getElementById('uploadProgress'),
                uploadProgressBar: document.getElementById('uploadProgressBar'),
                uploadProgressText: document.getElementById('uploadProgressText'),
                cancelUpload: document.getElementById('cancelUpload'),
                startUpload: document.getElementById('startUpload')
            };
        } catch (error) {
            console.error('Error initializing DOM elements:', error);
        }
    };
    
    const initializeEventListeners = function() {
        try {
            if (cachedElements.fileSelector) cachedElements.fileSelector.addEventListener('change', handleFileSelection);
            if (cachedElements.worksheetSelector) cachedElements.worksheetSelector.addEventListener('change', handleWorksheetSelection);
            if (cachedElements.keyColumnSelector) cachedElements.keyColumnSelector.addEventListener('change', handleKeyColumnSelection);
            if (cachedElements.refreshDataBtn) cachedElements.refreshDataBtn.addEventListener('click', handleRefreshData);
            if (cachedElements.loadFileBtn) cachedElements.loadFileBtn.addEventListener('click', handleLoadNewFile);
            
            cachedElements.tabItems.forEach(tabItem => {
                tabItem.addEventListener('click', e => {
                    e.preventDefault();
                    switchTab(e.currentTarget.getAttribute('data-tab'));
                });
            });
            
            initializeUploadModal();
            
        } catch (error) {
            console.error('Error initializing event listeners:', error);
        }
    };
    
    const handleFileSelection = async function(e) {
        const selectedFileId = e.target.value;
        if (!selectedFileId) {
            resetWorksheetSelector();
            resetKeyColumnSelector();
            resetAnalysisData();
            return;
        }
        appState.currentFileId = selectedFileId;
        appState.currentFileName = e.target.options[e.target.selectedIndex].text;

        const fileData = appState.uploadedFiles.get(selectedFileId);
        if (fileData && fileData.worksheets) {
            populateWorksheetSelector(fileData.worksheets.map(w => w.name));
        } else {
            await loadWorksheets(selectedFileId);
        }
    };
    
    const handleWorksheetSelection = async function(e) {
        const selectedWorksheet = e.target.value;
        if (!selectedWorksheet) {
            resetKeyColumnSelector();
            resetAnalysisData();
            return;
        }
        appState.currentWorksheet = selectedWorksheet;
        await loadAnalysisData(appState.currentFileId, selectedWorksheet);
    };

    const handleKeyColumnSelection = function(e) {
        appState.currentKeyColumn = e.target.value;
        console.log(`Key column selected: ${appState.currentKeyColumn}`);
        // Future: Trigger re-analysis based on new key
    };
    
    const handleRefreshData = async function(e) {
        e.preventDefault();
        if (!appState.currentFileId || !appState.currentWorksheet) {
            showErrorMessage('Please select a file and worksheet first');
            return;
        }
        await loadAnalysisData(appState.currentFileId, appState.currentWorksheet, true);
    };
    
    const handleLoadNewFile = function(e) {
        e.preventDefault();
        showUploadModal();
    };
    
    const loadWorksheets = async function(fileId) {
        showLoading(true);
        try {
            const response = await fetch(`/api/reports/worksheets/${fileId}`);
            if (!response.ok) throw new Error(`Failed to load worksheets. Status: ${response.status}`);
            const worksheets = await response.json();
            populateWorksheetSelector(worksheets);
        } catch (error) {
            console.error('Error loading worksheets:', error);
            showErrorMessage('Error loading worksheets for the selected file.');
            resetWorksheetSelector();
        } finally {
            showLoading(false);
        }
    };
    
    const loadAnalysisData = async function(fileId, worksheetName) {
        showLoading(true);
        resetAnalysisData();
        try {
            const response = await fetch(`/api/reports/preview/${fileId}/${worksheetName}`);
            if (!response.ok) throw new Error(`Failed to load analysis data. Status: ${response.status}`);
            
            const previewData = await response.json();
            appState.analysisData = previewData;
            
            // For now, display preview in the 'Assets' table
            updateTableWithPreviewData(cachedElements.assetsTable, previewData, "Assets");

            // Populate other tables with a placeholder
            updateTableWithPreviewData(cachedElements.integrityTable, null, "Integrity Issues");
            updateTableWithPreviewData(cachedElements.tasksTable, null, "Tasks");
            
            populateKeyColumnSelector(previewData.headers);

        } catch (error) {
            console.error('Error loading analysis data:', error);
            showErrorMessage('Error loading analysis data for the selected worksheet.');
            resetAnalysisData();
        } finally {
            showLoading(false);
        }
    };
    
    const populateWorksheetSelector = function(worksheets) {
        const selector = cachedElements.worksheetSelector;
        selector.innerHTML = '<option value="">Select a worksheet...</option>';
        worksheets.forEach(worksheet => {
            const option = document.createElement('option');
            option.value = worksheet;
            option.textContent = worksheet;
            selector.appendChild(option);
        });
        selector.disabled = false;
        resetKeyColumnSelector();
        resetAnalysisData();
    };

    const populateKeyColumnSelector = function(headers) {
        const selector = cachedElements.keyColumnSelector;
        const container = cachedElements.keyColumnSelectorContainer;
        selector.innerHTML = '<option value="">Select a key column...</option>';

        if (!headers || headers.length === 0) {
            container.style.display = 'none';
            selector.disabled = true;
            return;
        }

        headers.forEach(header => {
            if (header) { // Ensure header is not empty
                const option = document.createElement('option');
                option.value = header;
                option.textContent = header;
                selector.appendChild(option);
            }
        });

        selector.disabled = false;
        container.style.display = 'block';
    };
    
    const updateTableWithPreviewData = function(table, previewData, typeName) {
        if (!table) return;
        const thead = table.querySelector('thead');
        const tbody = table.querySelector('tbody');
        thead.innerHTML = '';
        tbody.innerHTML = '';

        if (!previewData || !previewData.headers || previewData.headers.length === 0) {
            const placeholderColspan = thead.parentElement.querySelector('tr')?.children.length || 6;
            tbody.innerHTML = `<tr class="no-data-row"><td colspan="${placeholderColspan}" class="text-center"><i class="material-icons">description</i><p>Select a file and worksheet to view ${typeName}</p></td></tr>`;
            return;
        }

        const headerRow = document.createElement('tr');
        previewData.headers.forEach(headerText => {
            const th = document.createElement('th');
            th.textContent = headerText;
            headerRow.appendChild(th);
        });
        thead.appendChild(headerRow);

        if (!previewData.rows || previewData.rows.length === 0) {
             tbody.innerHTML = `<tr class="no-data-row"><td colspan="${previewData.headers.length}" class="text-center"><i class="material-icons">info</i><p>No data rows found in this worksheet.</p></td></tr>`;
             return;
        }

        previewData.rows.forEach(rowData => {
            const row = document.createElement('tr');
            rowData.forEach(cellData => {
                const td = document.createElement('td');
                td.textContent = cellData;
                row.appendChild(td);
            });
            tbody.appendChild(row);
        });
    }
    
    const switchTab = function(tabId) {
        appState.currentTab = tabId;
        cachedElements.tabItems.forEach(item => {
            item.classList.toggle('active', item.getAttribute('data-tab') === tabId);
        });
        cachedElements.tabPanes.forEach(pane => {
            pane.classList.toggle('active', pane.id === tabId + 'Tab');
        });
    };
    
    const resetWorksheetSelector = function() {
        cachedElements.worksheetSelector.innerHTML = '<option value="">Select a file first...</option>';
        cachedElements.worksheetSelector.disabled = true;
    };

    const resetKeyColumnSelector = function() {
        cachedElements.keyColumnSelector.innerHTML = '<option value="">Select worksheet first...</option>';
        cachedElements.keyColumnSelector.disabled = true;
        cachedElements.keyColumnSelectorContainer.style.display = 'none';
    };
    
    const resetAnalysisData = function() {
        // Reset summary cards
        cachedElements.totalAssetsCount.textContent = '-';
        cachedElements.integrityIssuesCount.textContent = '-';
        cachedElements.missingTasksCount.textContent = '-';
        cachedElements.completionRate.textContent = '-';
        
        // Reset tables
        const tables = [cachedElements.integrityTable, cachedElements.assetsTable, cachedElements.tasksTable];
        tables.forEach(table => {
            if (table) {
                const tableBody = table.querySelector('tbody');
                const placeholderColspan = table.querySelector('thead tr')?.children.length || 6;
                if (tableBody) {
                    tableBody.innerHTML = `<tr class="no-data-row"><td colspan="${placeholderColspan}" class="text-center"><i class="material-icons">description</i><p>Select a file and worksheet to view analysis</p></td></tr>`;
                }
            }
        });
        
        resetKeyColumnSelector();
        appState.analysisData = null;
    };
    
    const showLoading = function(show) {
        cachedElements.loadingOverlay.classList.toggle('active', show);
        appState.isLoading = show;
    };
    
    const showErrorMessage = function(message) {
        // Simple alert for now - can be enhanced with custom notification system
        alert(`Error: ${message}`);
    };
    
    const initializeUploadModal = function() {
        if (cachedElements.uploadModalClose) cachedElements.uploadModalClose.addEventListener('click', hideUploadModal);
        if (cachedElements.cancelUpload) cachedElements.cancelUpload.addEventListener('click', hideUploadModal);
        if (cachedElements.uploadModal) {
            cachedElements.uploadModal.addEventListener('click', e => {
                if (e.target === cachedElements.uploadModal) hideUploadModal();
            });
        }
        if (cachedElements.fileInputModal) cachedElements.fileInputModal.addEventListener('change', handleModalFileSelection);
        if (cachedElements.uploadZoneModal) {
            cachedElements.uploadZoneModal.addEventListener('dragover', handleDragOver);
            cachedElements.uploadZoneModal.addEventListener('dragleave', handleDragLeave);
            cachedElements.uploadZoneModal.addEventListener('drop', handleFileDrop);
        }
        if (cachedElements.startUpload) cachedElements.startUpload.addEventListener('click', handleFileUpload);
    };
    
    const showUploadModal = function() {
        cachedElements.uploadModal.classList.add('active');
        resetUploadModal();
    };
    
    const hideUploadModal = function() {
        cachedElements.uploadModal.classList.remove('active');
        resetUploadModal();
    };
    
    const resetUploadModal = function() {
        cachedElements.fileInputModal.value = '';
        cachedElements.startUpload.disabled = true;
        cachedElements.uploadProgress.style.display = 'none';
        cachedElements.uploadContentModal.style.display = 'block';
        cachedElements.uploadProgressBar.style.width = '0%';
        cachedElements.uploadProgressText.textContent = 'Uploading...';
        resetUploadZoneDefaultText();
    };
    
    const handleModalFileSelection = function(e) {
        const files = e.target.files;
        cachedElements.startUpload.disabled = !(files && files.length > 0);
        if (files && files.length > 0) {
            updateUploadZoneWithFileInfo(files[0]);
        } else {
            resetUploadZoneDefaultText();
        }
    };
    
    const handleDragOver = function(e) {
        e.preventDefault();
        e.stopPropagation();
        cachedElements.uploadZoneModal.classList.add('dragover');
    };
    
    const handleDragLeave = function(e) {
        e.preventDefault();
        e.stopPropagation();
        cachedElements.uploadZoneModal.classList.remove('dragover');
    };
    
    const handleFileDrop = function(e) {
        e.preventDefault();
        e.stopPropagation();
        cachedElements.uploadZoneModal.classList.remove('dragover');
        const files = e.dataTransfer.files;
        if (files && files.length > 0) {
            cachedElements.fileInputModal.files = files;
            cachedElements.startUpload.disabled = false;
            updateUploadZoneWithFileInfo(files[0]);
        }
    };
    
    const handleFileUpload = function(e) {
        e.preventDefault();
        const files = cachedElements.fileInputModal.files;
        if (!files || files.length === 0) {
            showErrorMessage('Please select files to upload');
            return;
        }
        cachedElements.uploadContentModal.style.display = 'none';
        cachedElements.uploadProgress.style.display = 'block';
        cachedElements.startUpload.disabled = true;
        cachedElements.cancelUpload.disabled = true;
        uploadFiles(files);
    };
    
    const uploadFiles = async function(files) {
        const formData = new FormData();
        // Currently supports single file upload for analysis clarity
        const file = files[0];
        formData.append('file', file);

        try {
            const xhr = new XMLHttpRequest();
            xhr.open('POST', '/api/reports/upload', true);

            xhr.upload.onprogress = e => {
                if (e.lengthComputable) {
                    const progress = (e.loaded / e.total) * 100;
                    cachedElements.uploadProgressBar.style.width = `${progress}%`;
                    cachedElements.uploadProgressText.textContent = `Uploading ${file.name}... ${Math.round(progress)}%`;
                }
            };

            xhr.onload = () => {
                if (xhr.status === 200) {
                    cachedElements.uploadProgressText.textContent = 'Upload complete! Processing...';
                    const result = JSON.parse(xhr.responseText);
                    addFileToSelector(result);
                    setTimeout(hideUploadModal, 1000);
                } else {
                    const errorResponse = JSON.parse(xhr.responseText);
                    showErrorMessage(`Upload failed: ${errorResponse.message || 'Server error'}`);
                    resetUploadModal();
                }
                cachedElements.cancelUpload.disabled = false;
            };

            xhr.onerror = () => {
                showErrorMessage('An error occurred during the upload. Please check your network connection.');
                resetUploadModal();
                cachedElements.cancelUpload.disabled = false;
            };

            xhr.send(formData);
        } catch (error) {
            console.error('Error uploading files:', error);
            showErrorMessage('Error setting up file upload.');
            resetUploadModal();
            cachedElements.cancelUpload.disabled = false;
        }
    };
    
    const updateUploadZoneWithFileInfo = function(file) {
        if (!cachedElements.uploadContentModal) return;
        const icon = cachedElements.uploadContentModal.querySelector('.upload-icon-modal');
        const h4 = cachedElements.uploadContentModal.querySelector('h4');
        const p = cachedElements.uploadContentModal.querySelector('p');
        const button = cachedElements.uploadContentModal.querySelector('.btn-excel');

        if (icon) icon.textContent = 'description';
        if (h4) h4.textContent = 'File Ready for Upload';
        if (p) p.textContent = file.name;
        if (button) button.style.display = 'none';
    };

    const resetUploadZoneDefaultText = function() {
        if (!cachedElements.uploadContentModal) return;
        const icon = cachedElements.uploadContentModal.querySelector('.upload-icon-modal');
        const h4 = cachedElements.uploadContentModal.querySelector('h4');
        const p = cachedElements.uploadContentModal.querySelector('p');
        const button = cachedElements.uploadContentModal.querySelector('.btn-excel');

        if (icon) icon.textContent = 'cloud_upload';
        if (h4) h4.textContent = 'Drop Excel files here or click to browse';
        if (p) p.textContent = 'Supports .xlsx, .xls, and .csv files • Max 50MB per file';
        if (button) button.style.display = '';
    };

    const addFileToSelector = function(fileAnalysisResult) {
        // Store result for later use to avoid refetching worksheets
        appState.uploadedFiles.set(fileAnalysisResult.fileId, fileAnalysisResult);

        const fileSelector = cachedElements.fileSelector;
        const option = document.createElement('option');
        option.value = fileAnalysisResult.fileId;
        option.textContent = fileAnalysisResult.fileName;

        // Prepend new file to the top
        fileSelector.insertBefore(option, fileSelector.options[1]);
        
        fileSelector.value = fileAnalysisResult.fileId;
        fileSelector.dispatchEvent(new Event('change'));
    };
    
    const loadInitialData = async function() {
        showLoading(true);
        try {
            const response = await fetch('/api/reports/files');
            if (!response.ok) throw new Error('Failed to fetch initial file list.');
            const files = await response.json();
            
            const fileSelector = cachedElements.fileSelector;
            files.forEach(file => {
                const option = document.createElement('option');
                option.value = file.fileId;
                option.textContent = file.fileName;
                fileSelector.appendChild(option);
            });
        } catch (error) {
            console.error('Error loading initial data:', error);
            showErrorMessage('Could not load existing files from the server.');
        } finally {
            showLoading(false);
        }
    };
    
    const init = function() {
        try {
            initializeElements();
            initializeEventListeners();
            loadInitialData();
            console.log('View Reports module initialized successfully');
        } catch (error) {
            console.error('Error initializing View Reports module:', error);
        }
    };
    
    // Initialize module when DOM is ready
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
    
})(); 