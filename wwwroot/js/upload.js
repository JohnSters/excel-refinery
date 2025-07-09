/**
 * ExcelRefinery Upload Page JavaScript
 * File: upload.js
 * Description: Handles Excel file upload, preview, and worksheet selection functionality
 * Author: ExcelRefinery Development Team
 */

(function() {
    'use strict';
    
    // Global variables
    let uploadedFiles = [];
    let selectedWorksheets = new Set();
    let selectedHeaders = new Set();
    
    // Cache DOM elements
    let cachedElements = {};
    
    // Private methods
    const initializeElements = function() {
        try {
            cachedElements = {
                uploadZone: document.getElementById('uploadZone'),
                fileInput: document.getElementById('fileInput'),
                fileListSection: document.getElementById('fileListSection'),
                fileList: document.getElementById('fileList'),
                processSection: document.getElementById('processSection'),
                progressSection: document.getElementById('progressSection'),
                processFiles: document.getElementById('processFiles'),
                previewData: document.getElementById('previewData'),
                clearFiles: document.getElementById('clearFiles'),
                progressBar: document.getElementById('progressBar'),
                progressText: document.getElementById('progressText')
            };
        } catch (error) {
            console.error('Error initializing DOM elements:', error);
        }
    };
    
    const initializeEventListeners = function() {
        try {
            if (!cachedElements.uploadZone) return;
            
            // Click to upload
            cachedElements.uploadZone.addEventListener('click', () => cachedElements.fileInput.click());
            
            // File input change
            cachedElements.fileInput.addEventListener('change', handleFileSelect);
            
            // Drag and drop events
            cachedElements.uploadZone.addEventListener('dragover', handleDragOver);
            cachedElements.uploadZone.addEventListener('dragleave', handleDragLeave);
            cachedElements.uploadZone.addEventListener('drop', handleDrop);
            
            // Process buttons
            if (cachedElements.processFiles) {
                cachedElements.processFiles.addEventListener('click', startProcessing);
            }
            if (cachedElements.previewData) {
                cachedElements.previewData.addEventListener('click', previewData);
            }
            if (cachedElements.clearFiles) {
                cachedElements.clearFiles.addEventListener('click', clearAllFiles);
            }
        } catch (error) {
            console.error('Error initializing event listeners:', error);
        }
    };
    
    const handleDragOver = function(e) {
        try {
            e.preventDefault();
            cachedElements.uploadZone.classList.add('dragover');
        } catch (error) {
            console.error('Error handling drag over:', error);
        }
    };
    
    const handleDragLeave = function(e) {
        try {
            e.preventDefault();
            cachedElements.uploadZone.classList.remove('dragover');
        } catch (error) {
            console.error('Error handling drag leave:', error);
        }
    };
    
    const handleDrop = function(e) {
        try {
            e.preventDefault();
            cachedElements.uploadZone.classList.remove('dragover');
            const files = Array.from(e.dataTransfer.files);
            processFiles(files);
        } catch (error) {
            console.error('Error handling drop:', error);
        }
    };
    
    const handleFileSelect = function(e) {
        try {
            const files = Array.from(e.target.files);
            processFiles(files);
        } catch (error) {
            console.error('Error handling file select:', error);
        }
    };
    
    const processFiles = function(files) {
        try {
            const formData = new FormData();
            files.forEach(file => {
                if (validateFile(file)) {
                    formData.append('files', file);
                }
            });

            if (formData.has('files')) {
                uploadFiles(formData);
            }
        } catch (error) {
            console.error('Error processing files:', error);
        }
    };
    
    const validateFile = function(file) {
        try {
            const validTypes = [
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                'application/vnd.ms-excel',
                'text/csv'
            ];
            
            const validExtensions = ['.xlsx', '.xls', '.csv'];
            const fileExtension = '.' + file.name.split('.').pop().toLowerCase();
            
            if (!validTypes.includes(file.type) && !validExtensions.includes(fileExtension)) {
                alert('Invalid file type. Please upload Excel (.xlsx, .xls) or CSV (.csv) files.');
                return false;
            }
            
            if (file.size > 50 * 1024 * 1024) { // 50MB limit
                alert('File size too large. Maximum size is 50MB.');
                return false;
            }
            
            return true;
        } catch (error) {
            console.error('Error validating file:', error);
            return false;
        }
    };
    
    const uploadFiles = function(formData) {
        try {
            // Show loading spinner
            showLoadingSpinner();
            
            // Show upload progress
            if (window.StylingTemplate && window.StylingTemplate.showDemoAlert) {
                window.StylingTemplate.showDemoAlert('info', 'Uploading and analyzing files...');
            }

            fetch('/Home/UploadFiles', {
                method: 'POST',
                body: formData,
                headers: {
                    'RequestVerificationToken': document.querySelector('input[name="__RequestVerificationToken"]')?.value || ''
                }
            })
            .then(response => response.json())
            .then(data => {
                hideLoadingSpinner();
                
                if (data.success) {
                    // Add uploaded files to our collection
                    data.files.forEach(fileResult => {
                        // Convert server response to client format
                        const fileData = {
                            id: fileResult.fileId,
                            name: fileResult.fileName,
                            size: fileResult.fileSize,
                            type: fileResult.fileType,
                            lastModified: new Date(fileResult.lastModified),
                            worksheets: (fileResult.worksheets || []).map(ws => ({
                                id: ws.id,
                                name: ws.name || 'Unknown Sheet',
                                rowCount: ws.rowCount || 0,
                                columnCount: ws.columnCount || 0,
                                hasHeaders: ws.hasHeaders !== undefined ? ws.hasHeaders : true,
                                selected: ws.selected || false,
                                detectedHeaders: ws.detectedHeaders || [],
                                firstDataRowPreview: ws.firstDataRowPreview || ''
                            })),
                            headers: (fileResult.headers || []).map(header => ({
                                id: header.id,
                                name: header.detectedName || header.name || 'Unknown',
                                standardName: header.standardName || header.detectedName || header.name || 'Unknown',
                                type: header.dataType || header.type || 'Text',
                                selected: header.selected !== undefined ? header.selected : true,
                                isRequired: header.isRequired || false,
                                matchConfidence: header.matchConfidence || 1.0,
                                column: header.column || '',
                                sampleData: header.sampleData || ''
                            })),
                            status: fileResult.status,
                            validationErrors: fileResult.validationErrors || [],
                            validationWarnings: fileResult.validationWarnings || [],
                            qualityScore: fileResult.qualityScore || 0
                        };
                        
                        uploadedFiles.push(fileData);
                    });

                    renderFileList();
                    showProcessSection();

                    if (window.StylingTemplate && window.StylingTemplate.showDemoAlert) {
                        window.StylingTemplate.showDemoAlert('success', `Successfully processed ${data.files.length} file(s)`);
                    }
                } else {
                    console.error('Upload failed:', data.message);
                    if (window.StylingTemplate && window.StylingTemplate.showDemoAlert) {
                        window.StylingTemplate.showDemoAlert('danger', data.message || 'Error uploading files');
                    }
                }
            })
            .catch(error => {
                hideLoadingSpinner();
                console.error('Error uploading files:', error);
                if (window.StylingTemplate && window.StylingTemplate.showDemoAlert) {
                    window.StylingTemplate.showDemoAlert('danger', 'Error uploading files. Please try again.');
                }
            });
        } catch (error) {
            console.error('Error in uploadFiles:', error);
        }
    };
    
    const loadWorksheetHeaders = function(fileId, worksheetName) {
        try {
            fetch(`/Home/GetWorksheetHeaders?fileId=${encodeURIComponent(fileId)}&worksheetName=${encodeURIComponent(worksheetName)}`)
            .then(response => response.json())
            .then(data => {
                if (data.success) {
                    // Update the file's headers
                    const file = uploadedFiles.find(f => f.id === fileId);
                    if (file) {
                        file.headers = data.headers.map(header => ({
                            id: header.id,
                            name: header.detectedName || header.name || 'Unknown',
                            standardName: header.standardName || header.detectedName || header.name || 'Unknown',
                            type: header.dataType || header.type || 'Text',
                            selected: header.selected !== undefined ? header.selected : true,
                            isRequired: header.isRequired || false,
                            matchConfidence: header.matchConfidence || 1.0,
                            column: header.column || '',
                            sampleData: header.sampleData || ''
                        }));
                        
                        renderFileList();
                    }
                } else {
                    console.error('Failed to load headers:', data.message);
                }
            })
            .catch(error => {
                console.error('Error loading worksheet headers:', error);
            });
        } catch (error) {
            console.error('Error in loadWorksheetHeaders:', error);
        }
    };
    
    const renderFileList = function() {
        try {
            if (!cachedElements.fileList) return;
            
            cachedElements.fileList.innerHTML = '';
            
            uploadedFiles.forEach(fileData => {
                const fileElement = createFileElement(fileData);
                cachedElements.fileList.appendChild(fileElement);
            });
            
            if (cachedElements.fileListSection) {
                cachedElements.fileListSection.style.display = 'block';
            }
        } catch (error) {
            console.error('Error rendering file list:', error);
        }
    };
    
    const createFileElement = function(fileData) {
        try {
            const fileDiv = document.createElement('div');
            fileDiv.className = 'file-item';
            fileDiv.innerHTML = `
                <div class="file-header">
                    <div class="file-info">
                        <i class="material-icons file-icon">description</i>
                        <div>
                            <h4 style="color: var(--federal-blue); margin: 0;">${fileData.name}</h4>
                            <span class="status-badge status-${fileData.status}">
                                <i class="material-icons" style="font-size: 1rem;">check_circle</i>
                                Ready for processing
                            </span>
                        </div>
                    </div>
                    <button class="btn-excel btn-excel-danger btn-excel-sm" onclick="UploadHandler.removeFile('${fileData.id}')">
                        <i class="material-icons">delete</i>
                        Remove
                    </button>
                </div>
                
                <div class="file-details">
                    <div class="detail-item">
                        <div class="detail-label">File Size</div>
                        <div class="detail-value">${formatFileSize(fileData.size)}</div>
                    </div>
                    <div class="detail-item">
                        <div class="detail-label">File Type</div>
                        <div class="detail-value">${getFileTypeDisplay(fileData.type, fileData.name)}</div>
                    </div>
                    <div class="detail-item">
                        <div class="detail-label">Last Modified</div>
                        <div class="detail-value">${fileData.lastModified.toLocaleDateString()}</div>
                    </div>
                                            <div class="detail-item">
                            <div class="detail-label">Worksheets</div>
                            <div class="detail-value">${fileData.worksheets.length} sheet(s)</div>
                        </div>
                        <div class="detail-item">
                            <div class="detail-label">Quality Score</div>
                            <div class="detail-value">
                                <span class="quality-score ${getQualityScoreClass(fileData.qualityScore)}">${fileData.qualityScore || 0}%</span>
                            </div>
                        </div>
                    </div>
                    
                    ${fileData.validationErrors && fileData.validationErrors.length > 0 ? `
                    <div class="alert alert-danger mt-2">
                        <strong>Validation Errors:</strong>
                        <ul class="mb-0 mt-1">
                            ${fileData.validationErrors.map(error => `<li>${error}</li>`).join('')}
                        </ul>
                    </div>
                    ` : ''}
                    
                    ${fileData.validationWarnings && fileData.validationWarnings.length > 0 ? `
                    <div class="alert alert-warning mt-2">
                        <strong>Warnings:</strong>
                        <ul class="mb-0 mt-1">
                            ${fileData.validationWarnings.map(warning => `<li>${warning}</li>`).join('')}
                        </ul>
                    </div>
                    ` : ''}
                
                <div class="worksheets-section">
                    <h5 style="color: var(--federal-blue); margin-bottom: 0.5rem;">
                        <i class="material-icons" style="font-size: 1.2rem; vertical-align: middle; margin-right: 0.5rem;">tab</i>
                        Worksheets
                    </h5>
                    <p style="color: var(--text-secondary); font-size: 0.875rem; margin-bottom: 1rem;">
                        Select which worksheets to include in processing
                    </p>
                    <div class="worksheet-grid">
                        ${fileData.worksheets.map(ws => createWorksheetCard(ws, fileData.id)).join('')}
                    </div>
                </div>
                
                <div class="headers-section">
                    <h5 style="color: var(--federal-blue); margin-bottom: 0.5rem;">
                        <i class="material-icons" style="font-size: 1.2rem; vertical-align: middle; margin-right: 0.5rem;">view_column</i>
                        Column Headers
                    </h5>
                    <p style="color: var(--text-secondary); font-size: 0.875rem; margin-bottom: 1rem;">
                        Select which columns to include in processing
                    </p>
                    <div class="headers-grid">
                        ${fileData.headers.map(header => createHeaderItem(header, fileData.id)).join('')}
                    </div>
                </div>
            `;
            
            return fileDiv;
        } catch (error) {
            console.error('Error creating file element:', error);
            return document.createElement('div');
        }
    };
    
    const createWorksheetCard = function(worksheet, fileId) {
        try {
            return `
                <div class="worksheet-card ${worksheet.selected ? 'selected' : ''}" 
                     onclick="UploadHandler.toggleWorksheet('${fileId}', '${worksheet.id}')">
                    <div class="worksheet-name">${worksheet.name}</div>
                    <div class="worksheet-stats">
                        ${(worksheet.rowCount || 0).toLocaleString()} rows • ${worksheet.columnCount || 0} columns
                    </div>
                </div>
            `;
        } catch (error) {
            console.error('Error creating worksheet card:', error);
            return '';
        }
    };
    
    const createHeaderItem = function(header, fileId) {
        try {
            const dataTypeClass = getDataTypeClass(header.type);
            
            return `
                <div class="header-item ${header.selected ? 'selected' : ''}" 
                     onclick="UploadHandler.toggleHeader('${fileId}', '${header.id}')">
                    <input type="checkbox" class="header-checkbox" ${header.selected ? 'checked' : ''} 
                           onchange="event.stopPropagation()">
                    <div class="header-content">
                        <div class="header-name">
                            <strong>${header.name}</strong>
                            <span class="badge ${dataTypeClass} ms-1">${header.type}</span>
                        </div>
                        ${header.sampleData ? `<small class="text-muted d-block">Sample: ${header.sampleData}</small>` : ''}
                    </div>
                </div>
            `;
        } catch (error) {
            console.error('Error creating header item:', error);
            return '';
        }
    };
    
    const getDataTypeClass = function(dataType) {
        switch (dataType) {
            case 'Date': return 'badge-info';
            case 'Numeric': return 'badge-success';
            case 'Boolean': return 'badge-warning';
            case 'Text':
            default: return 'badge-secondary';
        }
    };
    
    const formatFileSize = function(bytes) {
        try {
            if (bytes === 0) return '0 Bytes';
            const k = 1024;
            const sizes = ['Bytes', 'KB', 'MB', 'GB'];
            const i = Math.floor(Math.log(bytes) / Math.log(k));
            return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
        } catch (error) {
            console.error('Error formatting file size:', error);
            return 'Unknown';
        }
    };
    
    const getFileTypeDisplay = function(type, name) {
        try {
            const extension = name.split('.').pop().toLowerCase();
            switch (extension) {
                case 'xlsx': return 'Excel Workbook (.xlsx)';
                case 'xls': return 'Excel 97-2003 (.xls)';
                case 'csv': return 'CSV File (.csv)';
                default: return type || 'Unknown';
            }
        } catch (error) {
            console.error('Error getting file type display:', error);
            return 'Unknown';
        }
    };
    
    const getQualityScoreClass = function(score) {
        try {
            if (score >= 90) return 'quality-excellent';
            if (score >= 75) return 'quality-good';
            if (score >= 50) return 'quality-fair';
            return 'quality-poor';
        } catch (error) {
            console.error('Error getting quality score class:', error);
            return 'quality-poor';
        }
    };
    
    const showProcessSection = function() {
        try {
            if (uploadedFiles.length > 0 && cachedElements.processSection) {
                cachedElements.processSection.style.display = 'block';
            }
        } catch (error) {
            console.error('Error showing process section:', error);
        }
    };
    
    const startProcessing = function() {
        try {
            if (cachedElements.progressSection) {
                cachedElements.progressSection.style.display = 'block';
                simulateProcessing();
            }
        } catch (error) {
            console.error('Error starting processing:', error);
        }
    };
    
    const simulateProcessing = function() {
        try {
            let progress = 0;
            
            const interval = setInterval(() => {
                progress += Math.random() * 15;
                if (progress >= 100) {
                    progress = 100;
                    clearInterval(interval);
                    if (cachedElements.progressText) {
                        cachedElements.progressText.textContent = 'Processing completed successfully!';
                    }
                    
                    setTimeout(() => {
                        if (window.StylingTemplate && window.StylingTemplate.showDemoAlert) {
                            window.StylingTemplate.showDemoAlert('success', 'Files processed successfully! Results are ready for review.');
                        }
                    }, 500);
                } else {
                    if (cachedElements.progressText) {
                        cachedElements.progressText.textContent = `Processing files... ${Math.round(progress)}%`;
                    }
                }
                if (cachedElements.progressBar) {
                    cachedElements.progressBar.style.width = progress + '%';
                }
            }, 500);
        } catch (error) {
            console.error('Error simulating processing:', error);
        }
    };
    
    const previewData = function() {
        try {
            // Find the first file with a selected worksheet
            const fileWithWorksheet = uploadedFiles.find(file => 
                file.worksheets.some(ws => ws.selected)
            );
            
            if (!fileWithWorksheet) {
                if (window.StylingTemplate && window.StylingTemplate.showDemoAlert) {
                    window.StylingTemplate.showDemoAlert('warning', 'Please select a worksheet to preview data.');
                }
                return;
            }
            
            const selectedWorksheet = fileWithWorksheet.worksheets.find(ws => ws.selected);
            
            fetch(`/Home/GetDataPreview?fileId=${encodeURIComponent(fileWithWorksheet.id)}&worksheetName=${encodeURIComponent(selectedWorksheet.name)}&maxRows=10`)
            .then(response => response.json())
            .then(data => {
                if (data.success) {
                    showDataPreviewModal(data.data);
                } else {
                    if (window.StylingTemplate && window.StylingTemplate.showDemoAlert) {
                        window.StylingTemplate.showDemoAlert('danger', data.message || 'Error loading data preview.');
                    }
                }
            })
            .catch(error => {
                console.error('Error getting data preview:', error);
                if (window.StylingTemplate && window.StylingTemplate.showDemoAlert) {
                    window.StylingTemplate.showDemoAlert('danger', 'Error loading data preview.');
                }
            });
        } catch (error) {
            console.error('Error previewing data:', error);
        }
    };
    
    const showDataPreviewModal = function(previewData) {
        try {
            // Create a simple modal to show the data preview
            const modalHtml = `
                <div class="modal fade" id="dataPreviewModal" tabindex="-1">
                    <div class="modal-dialog modal-xl">
                        <div class="modal-content">
                            <div class="modal-header">
                                <h5 class="modal-title">Data Preview - ${previewData.worksheetId}</h5>
                                <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
                            </div>
                            <div class="modal-body">
                                <div class="table-responsive">
                                    <table class="table table-striped table-sm">
                                        <thead>
                                            <tr>
                                                ${previewData.headers.map(header => `<th>${header}</th>`).join('')}
                                            </tr>
                                        </thead>
                                        <tbody>
                                            ${previewData.rows.map(row => 
                                                `<tr>${row.map(cell => `<td>${cell || ''}</td>`).join('')}</tr>`
                                            ).join('')}
                                        </tbody>
                                    </table>
                                </div>
                                <small class="text-muted">
                                    Showing ${previewData.rows.length} of ${previewData.totalRows} rows
                                    ${previewData.hasMoreData ? ' (partial preview)' : ''}
                                </small>
                            </div>
                        </div>
                    </div>
                </div>
            `;
            
            // Remove existing modal if any
            const existingModal = document.getElementById('dataPreviewModal');
            if (existingModal) {
                existingModal.remove();
            }
            
            // Add new modal to page
            document.body.insertAdjacentHTML('beforeend', modalHtml);
            
            // Show the modal
            const modal = new bootstrap.Modal(document.getElementById('dataPreviewModal'));
            modal.show();
            
            // Clean up when modal is hidden
            document.getElementById('dataPreviewModal').addEventListener('hidden.bs.modal', function() {
                this.remove();
            });
        } catch (error) {
            console.error('Error showing data preview modal:', error);
        }
    };
    
    const clearAllFiles = function() {
        try {
            if (confirm('Are you sure you want to clear all uploaded files?')) {
                uploadedFiles = [];
                selectedWorksheets.clear();
                selectedHeaders.clear();
                
                if (cachedElements.fileListSection) {
                    cachedElements.fileListSection.style.display = 'none';
                }
                if (cachedElements.processSection) {
                    cachedElements.processSection.style.display = 'none';
                }
                if (cachedElements.progressSection) {
                    cachedElements.progressSection.style.display = 'none';
                }
                if (cachedElements.fileInput) {
                    cachedElements.fileInput.value = '';
                }
            }
        } catch (error) {
            console.error('Error clearing files:', error);
        }
    };
    
    const showLoadingSpinner = function() {
        try {
            // Create spinner overlay if it doesn't exist
            let spinner = document.getElementById('uploadSpinner');
            if (!spinner) {
                const spinnerHtml = `
                    <div id="uploadSpinner" class="upload-spinner">
                        <div class="spinner-content">
                            <div class="spinner-border text-primary" style="width: 3rem; height: 3rem;"></div>
                            <div class="mt-3">Processing files...</div>
                        </div>
                    </div>
                `;
                document.body.insertAdjacentHTML('beforeend', spinnerHtml);
                spinner = document.getElementById('uploadSpinner');
            }
            spinner.style.display = 'flex';
        } catch (error) {
            console.error('Error showing loading spinner:', error);
        }
    };
    
    const hideLoadingSpinner = function() {
        try {
            const spinner = document.getElementById('uploadSpinner');
            if (spinner) {
                spinner.style.display = 'none';
            }
        } catch (error) {
            console.error('Error hiding loading spinner:', error);
        }
    };
    
    // Public methods
    const UploadHandler = {
        init: function() {
            try {
                initializeElements();
                initializeEventListeners();
                console.log('Upload Handler initialized successfully');
            } catch (error) {
                console.error('Error initializing Upload Handler:', error);
            }
        },
        
        removeFile: function(fileId) {
            try {
                uploadedFiles = uploadedFiles.filter(f => f.id != fileId);
                if (uploadedFiles.length === 0) {
                    if (cachedElements.fileListSection) {
                        cachedElements.fileListSection.style.display = 'none';
                    }
                    if (cachedElements.processSection) {
                        cachedElements.processSection.style.display = 'none';
                    }
                } else {
                    renderFileList();
                }
            } catch (error) {
                console.error('Error removing file:', error);
            }
        },
        
        toggleWorksheet: function(fileId, worksheetId) {
            try {
                const file = uploadedFiles.find(f => f.id == fileId);
                if (file) {
                    // First, unselect all worksheets for this file
                    file.worksheets.forEach(ws => ws.selected = false);
                    
                    // Then select the clicked worksheet
                    const worksheet = file.worksheets.find(ws => ws.id === worksheetId);
                    if (worksheet) {
                        worksheet.selected = true;
                        
                        // Load headers for the selected worksheet
                        loadWorksheetHeaders(fileId, worksheet.name);
                    }
                }
            } catch (error) {
                console.error('Error toggling worksheet:', error);
            }
        },
        
        toggleHeader: function(fileId, headerId) {
            try {
                const file = uploadedFiles.find(f => f.id == fileId);
                if (file) {
                    const header = file.headers.find(h => h.id === headerId);
                    if (header) {
                        header.selected = !header.selected;
                        renderFileList();
                    }
                }
            } catch (error) {
                console.error('Error toggling header:', error);
            }
        },
        
        refresh: function() {
            try {
                initializeElements();
                console.log('Upload Handler refreshed');
            } catch (error) {
                console.error('Error refreshing Upload Handler:', error);
            }
        }
    };
    
    // Auto-initialize when DOM is ready
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', UploadHandler.init);
    } else {
        UploadHandler.init();
    }
    
    // Make UploadHandler available globally for onclick handlers
    window.UploadHandler = UploadHandler;
    
})(); 