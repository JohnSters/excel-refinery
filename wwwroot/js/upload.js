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
            files.forEach(file => {
                if (validateFile(file)) {
                    analyzeFile(file);
                }
            });
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
    
    const analyzeFile = function(file) {
        try {
            // Create mock file analysis (in real app, this would use a library like SheetJS)
            const mockWorksheets = generateMockWorksheets(file.name);
            const mockHeaders = generateMockHeaders();
            
            const fileData = {
                id: Date.now() + Math.random(),
                file: file,
                name: file.name,
                size: file.size,
                type: file.type,
                lastModified: new Date(file.lastModified),
                worksheets: mockWorksheets,
                headers: mockHeaders,
                status: 'ready'
            };
            
            uploadedFiles.push(fileData);
            renderFileList();
            showProcessSection();
        } catch (error) {
            console.error('Error analyzing file:', error);
        }
    };
    
    const generateMockWorksheets = function(filename) {
        try {
            const worksheetNames = ['Equipment_Data', 'Maintenance_Schedule', 'Task_List', 'Summary'];
            const selectedSheets = worksheetNames.slice(0, Math.floor(Math.random() * 3) + 1);
            
            return selectedSheets.map((name, index) => ({
                id: `ws_${index}`,
                name: name,
                rows: Math.floor(Math.random() * 1000) + 100,
                columns: Math.floor(Math.random() * 20) + 10,
                selected: index === 0 // First sheet selected by default
            }));
        } catch (error) {
            console.error('Error generating mock worksheets:', error);
            return [];
        }
    };
    
    const generateMockHeaders = function() {
        try {
            const standardHeaders = [
                'Equipment ID', 'CMMS System', 'Equipment Technical Number',
                'Task ID', 'Task Type', 'Task Description', 'Task Details',
                'Last Date', 'Override Interval', 'Desired Interval',
                'Reoccurring', 'Next Date', 'Next Date Basis',
                'Task Assigned To', 'Reason', 'Related Entity ID'
            ];
            
            return standardHeaders.map((header, index) => ({
                id: `header_${index}`,
                name: header,
                type: getHeaderType(header),
                selected: ['Equipment ID', 'Task ID', 'Task Type'].includes(header)
            }));
        } catch (error) {
            console.error('Error generating mock headers:', error);
            return [];
        }
    };
    
    const getHeaderType = function(header) {
        try {
            if (header.includes('Date')) return 'Date';
            if (header.includes('ID') || header.includes('Interval')) return 'Numeric';
            if (header.includes('Reoccurring')) return 'Boolean';
            return 'Text';
        } catch (error) {
            console.error('Error getting header type:', error);
            return 'Text';
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
                </div>
                
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
                        ${worksheet.rows.toLocaleString()} rows • ${worksheet.columns} columns
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
            return `
                <div class="header-item ${header.selected ? 'selected' : ''}" 
                     onclick="UploadHandler.toggleHeader('${fileId}', '${header.id}')">
                    <input type="checkbox" class="header-checkbox" ${header.selected ? 'checked' : ''} 
                           onchange="event.stopPropagation()">
                    <div class="header-name">${header.name}</div>
                </div>
            `;
        } catch (error) {
            console.error('Error creating header item:', error);
            return '';
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
            if (window.StylingTemplate && window.StylingTemplate.showDemoAlert) {
                window.StylingTemplate.showDemoAlert('info', 'Data preview functionality coming soon!');
            } else {
                alert('Data preview functionality coming soon!');
            }
        } catch (error) {
            console.error('Error previewing data:', error);
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
                    const worksheet = file.worksheets.find(ws => ws.id === worksheetId);
                    if (worksheet) {
                        worksheet.selected = !worksheet.selected;
                        renderFileList();
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