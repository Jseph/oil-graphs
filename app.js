document.addEventListener('DOMContentLoaded', () => {
    // Global Application State
    const AppState = {
        wells: {},           // Map of Well ID -> { id, points, outlierSensitivity, bdfFitReal, outliersReal, bdfFitMbt, outliersMbt }
        wellIds: [],         // Array of string Well IDs (filtered & sorted)
        currentWellIndex: 0,
        rawHeaders: [],
        rawRows: [],
        rawFileRows: [],     // 2D array of all parsed rows from the file
        activeView: 'view-upload',
        charts: {
            analysis: null,
            cdfReal: null,
            cdfMbt: null
        }
    };

    // DOM Elements
    const navButtons = document.querySelectorAll('.nav-btn');
    const viewPanels = document.querySelectorAll('.view-panel');
    const navAnalysisBtn = document.getElementById('nav-analysis-btn');
    const navSummaryBtn = document.getElementById('nav-summary-btn');

    // Upload Elements
    const uploadArea = document.getElementById('upload-area');
    const csvFileInput = document.getElementById('csv-file-input');
    const browseLink = document.getElementById('browse-link');
    const fileInfoLabel = document.getElementById('file-info-label');
    const mappingControls = document.getElementById('mapping-controls');
    const mapWellIdSelect = document.getElementById('map-well-id');
    const mapDateSelect = document.getElementById('map-date');
    const mapDateFormatSelect = document.getElementById('map-date-format');
    const mapVolumeSelect = document.getElementById('map-volume');
    const processCsvBtn = document.getElementById('process-csv-btn');
    const csvPreviewContainer = document.getElementById('csv-preview-container');
    const csvPreviewTable = document.getElementById('csv-preview-table');
    const defaultSensitivitySlider = document.getElementById('default-sensitivity-slider');
    const defaultSensitivityValue = document.getElementById('default-sensitivity-value');
    const hasHeadersCheckbox = document.getElementById('has-headers-checkbox');

    // Analysis Elements
    const wellSearchInput = document.getElementById('well-search');
    const wellsListUl = document.getElementById('wells-list');
    const prevWellBtn = document.getElementById('prev-well-btn');
    const nextWellBtn = document.getElementById('next-well-btn');
    const currentWellInfoSpan = document.getElementById('current-well-info');
    const activeWellNameH2 = document.getElementById('active-well-name');
    const activeWellBadge = document.getElementById('active-well-badge');
    const timeProdRadio = document.getElementById('time-prod-analysis');
    const timeMbtRadio = document.getElementById('time-mbt-analysis');
    const sensitivitySlider = document.getElementById('sensitivity-slider-analysis');
    const sensitivityValueSpan = document.getElementById('sensitivity-value-analysis');

    const pointCountAnalysisSpan = document.getElementById('point-count-analysis');
    const ipValAnalysisSpan = document.getElementById('ip-val-analysis');
    const ttdRealValSpan = document.getElementById('ttd-real-val');
    const ttdMbtValSpan = document.getElementById('ttd-mbt-val');
    const ttdRealDaysSpan = document.getElementById('ttd-real-days');
    const ttdMbtDaysSpan = document.getElementById('ttd-mbt-days');
    const sseValAnalysisSpan = document.getElementById('sse-val-analysis');

    // Summary Elements
    const summaryTotalWellsSpan = document.getElementById('summary-total-wells');
    const summaryAvgIpSpan = document.getElementById('summary-avg-ip');
    const summaryAvgTtdRealSpan = document.getElementById('summary-avg-ttd-real');
    const summaryAvgTtdMbtSpan = document.getElementById('summary-avg-ttd-mbt');
    const summaryTableBody = document.getElementById('summary-table-body');
    const exportCsvBtn = document.getElementById('export-csv-btn');

    // Helper functions
    const daysToMonths = (days) => days / 30.4;
    const formatNumber = (num, decimals = 0) => {
        if (num === null || num === undefined || isNaN(num)) return '--';
        return num.toLocaleString(undefined, { maximumFractionDigits: decimals });
    };

    // --- NAVIGATION LOGIC ---
    const switchView = (targetViewId) => {
        viewPanels.forEach(panel => {
            if (panel.id === targetViewId) {
                panel.classList.add('active');
            } else {
                panel.classList.remove('active');
            }
        });

        navButtons.forEach(btn => {
            if (btn.getAttribute('data-target') === targetViewId) {
                btn.classList.add('active');
            } else {
                btn.classList.remove('active');
            }
        });

        AppState.activeView = targetViewId;

        // Perform view-specific loads
        if (targetViewId === 'view-analysis') {
            renderWellsList();
            loadActiveWell();
        } else if (targetViewId === 'view-summary') {
            renderSummaryPage();
        }
    };

    navButtons.forEach(btn => {
        btn.addEventListener('click', () => {
            if (!btn.disabled) {
                switchView(btn.getAttribute('data-target'));
            }
        });
    });

    // --- CSV PARSING & UPLOAD LOGIC ---
    // Handle File Browsing
    browseLink.addEventListener('click', (e) => {
        e.stopPropagation();
        csvFileInput.click();
    });

    uploadArea.addEventListener('click', () => {
        csvFileInput.click();
    });

    // Drag and Drop
    ['dragenter', 'dragover'].forEach(eventName => {
        uploadArea.addEventListener(eventName, (e) => {
            e.preventDefault();
            e.stopPropagation();
            uploadArea.classList.add('dragover');
        }, false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        uploadArea.addEventListener(eventName, (e) => {
            e.preventDefault();
            e.stopPropagation();
            uploadArea.classList.remove('dragover');
        }, false);
    });

    uploadArea.addEventListener('drop', (e) => {
        const dt = e.dataTransfer;
        const files = dt.files;
        if (files.length > 0) {
            handleSelectedFile(files[0]);
        }
    });

    csvFileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            handleSelectedFile(e.target.files[0]);
        }
    });

    const handleSelectedFile = (file) => {
        const isExcel = file.name.endsWith('.xlsx') || file.name.endsWith('.xls');
        const isCsv = file.name.endsWith('.csv');

        if (!isExcel && !isCsv) {
            alert('Please select a valid CSV or Excel file.');
            return;
        }
        
        fileInfoLabel.textContent = `${file.name} (${(file.size / 1024).toFixed(1)} KB)`;
        
        const reader = new FileReader();
        
        // Reset checkbox to checked on new file upload
        if (hasHeadersCheckbox) {
            hasHeadersCheckbox.checked = true;
        }

        if (isExcel) {
            reader.onload = (event) => {
                try {
                    const data = new Uint8Array(event.target.result);
                    const workbook = XLSX.read(data, { type: 'array', cellDates: true });
                    
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    
                    const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                    if (rows.length === 0) {
                        alert('The Excel file is empty.');
                        return;
                    }
                    
                    // Normalize cell values to strings, preserving empty strings for null/undefined
                    const maxCols = Math.max(...rows.map(r => r.length));
                    const cleanRows = rows.map(row => {
                        const cleanRow = [];
                        for (let i = 0; i < maxCols; i++) {
                            const val = row[i];
                            cleanRow.push(val === null || val === undefined ? '' : String(val).trim());
                        }
                        return cleanRow;
                    });
                    
                    AppState.rawFileRows = cleanRows;
                    processRawFileRows();
                } catch (error) {
                    console.error(error);
                    alert('Error parsing Excel file. Make sure it is a valid .xlsx or .xls file.');
                }
            };
            reader.readAsArrayBuffer(file);
        } else {
            reader.onload = (event) => {
                parseCSVText(event.target.result);
            };
            reader.readAsText(file);
        }
    };

    const parseCSVLine = (line) => {
        const result = [];
        let current = '';
        let inQuotes = false;
        let delimiter = ',';
        
        if (line.includes('\t')) delimiter = '\t';
        else if (line.includes(';')) delimiter = ';';
        
        for (let i = 0; i < line.length; i++) {
            const char = line[i];
            if (char === '"') {
                inQuotes = !inQuotes;
            } else if (char === delimiter && !inQuotes) {
                result.push(current);
                current = '';
            } else {
                current += char;
            }
        }
        result.push(current);
        return result;
    };

    const parseCSVText = (text) => {
        const lines = text.split(/\r?\n/);
        const rows = [];

        for (let i = 0; i < lines.length; i++) {
            const line = lines[i].trim();
            if (line === '') continue;
            
            const parts = parseCSVLine(line).map(h => h.trim().replace(/^["']|["']$/g, ''));
            rows.push(parts);
        }

        if (rows.length === 0) {
            alert('The file is empty.');
            return;
        }

        AppState.rawFileRows = rows;
        processRawFileRows();
    };

    const processRawFileRows = () => {
        const rows = AppState.rawFileRows;
        if (!rows || rows.length === 0) return;

        const hasHeaders = hasHeadersCheckbox ? hasHeadersCheckbox.checked : true;
        let headers = [];
        let dataRows = [];

        if (hasHeaders) {
            headers = rows[0];
            dataRows = rows.slice(1);
        } else {
            // Generate dummy headers like Column 1, Column 2, etc.
            const colCount = rows[0].length;
            for (let i = 0; i < colCount; i++) {
                headers.push(`Column ${i + 1}`);
            }
            dataRows = rows;
        }

        initializeDataMapping(headers, dataRows);
    };

    const initializeDataMapping = (headers, rows) => {
        if (headers.length < 3) {
            alert('Your file must have at least 3 columns (Well ID, Timestamp, Volume).');
            return;
        }

        AppState.rawHeaders = headers;
        AppState.rawRows = rows;

        populateMappingSelectors(headers);
        refreshPreview();

        // Reset default sensitivity slider to 0% when a new file is mapped
        if (defaultSensitivitySlider) {
            defaultSensitivitySlider.value = 0;
        }
        if (defaultSensitivityValue) {
            defaultSensitivityValue.textContent = '0%';
        }

        mappingControls.style.display = 'block';
        csvPreviewContainer.style.display = 'block';
    };

    if (defaultSensitivitySlider && defaultSensitivityValue) {
        defaultSensitivitySlider.addEventListener('input', () => {
            defaultSensitivityValue.textContent = `${defaultSensitivitySlider.value}%`;
        });
    }

    if (hasHeadersCheckbox) {
        hasHeadersCheckbox.addEventListener('change', () => {
            processRawFileRows();
        });
    }

    // Re-render preview when format or date column selection changes
    const refreshPreview = () => {
        if (!AppState.rawHeaders.length) return;
        const dateColIdx = mapDateSelect ? parseInt(mapDateSelect.value) : -1;
        const timeFormat = mapDateFormatSelect ? mapDateFormatSelect.value : 'seconds';
        renderCSVPreview(AppState.rawHeaders, AppState.rawRows.slice(0, 5), dateColIdx, timeFormat);
    };

    if (mapDateFormatSelect) {
        mapDateFormatSelect.addEventListener('change', refreshPreview);
    }
    if (mapDateSelect) {
        mapDateSelect.addEventListener('change', refreshPreview);
    }

    const populateMappingSelectors = (headers) => {
        [mapWellIdSelect, mapDateSelect, mapVolumeSelect].forEach(select => {
            select.innerHTML = '';
            headers.forEach((h, idx) => {
                const opt = new Option(h, idx);
                select.add(opt);
            });
        });

        // Smart column mapping heuristics
        let dateColIdx = -1;
        headers.forEach((h, idx) => {
            const lower = h.toLowerCase();
            if (lower.includes('well') || lower.includes('api') || lower.includes('id') || lower.includes('name')) {
                mapWellIdSelect.value = idx;
            }
            if (lower.includes('date') || lower.includes('time') || lower.includes('stamp')) {
                mapDateSelect.value = idx;
                dateColIdx = idx;
            }
            if (lower.includes('volume') || lower.includes('prod') || lower.includes('oil') || lower.includes('qty') || lower.includes('rate')) {
                mapVolumeSelect.value = idx;
            }
        });

        // Heuristics for Time Format: Excel serial dates for recent years are ~36000–50000.
        // Unix epoch seconds for the same range are in the billions.
        if (dateColIdx !== -1 && rows && rows.length > 0) {
            const sampleVal = rows[0][dateColIdx];
            if (sampleVal) {
                const num = parseFloat(sampleVal);
                if (!isNaN(num)) {
                    mapDateFormatSelect.value = num > 100000000 ? 'seconds' : 'excel';
                }
            }
        }
    };

    // Convert a raw time cell value to a JS Date based on the chosen format
    const parseTimestampToDate = (raw, timeFormat) => {
        const num = parseFloat(raw);
        if (timeFormat === 'seconds') {
            if (!isNaN(num)) return new Date(num * 1000);
        } else if (timeFormat === 'excel') {
            // Excel serial: Day 1 = Jan 1, 1900. Offset vs Unix epoch is 25569 days
            // (accounts for Excel's erroneous Feb 29, 1900 leap year bug)
            if (!isNaN(num)) return new Date((num - 25569) * 86400 * 1000);
        }
        return null;
    };

    const formatDateMMDDYYYY = (date) => {
        if (!date) return '—';
        const mm = String(date.getMonth() + 1).padStart(2, '0');
        const dd = String(date.getDate()).padStart(2, '0');
        const yyyy = date.getFullYear();
        return `${mm}/${dd}/${yyyy}`;
    };

    const renderCSVPreview = (headers, rows, dateColIdx = -1, timeFormat = 'seconds') => {
        let html = '<thead><tr>';
        headers.forEach(h => {
            html += `<th>${h}</th>`;
        });
        html += '</tr></thead><tbody>';
        
        rows.forEach(row => {
            html += '<tr>';
            row.forEach((val, colIdx) => {
                if (colIdx === dateColIdx) {
                    const d = parseTimestampToDate(val, timeFormat);
                    const formatted = d ? formatDateMMDDYYYY(d) : val;
                    html += `<td title="${val}" style="color: var(--accent-color)">${formatted}</td>`;
                } else {
                    html += `<td>${val}</td>`;
                }
            });
            html += '</tr>';
        });
        html += '</tbody>';
        csvPreviewTable.innerHTML = html;
    };

    // Process and Group CSV by Well ID
    processCsvBtn.addEventListener('click', () => {
        const wellIdx = parseInt(mapWellIdSelect.value);
        const dateIdx = parseInt(mapDateSelect.value);
        const volumeIdx = parseInt(mapVolumeSelect.value);
        const timeFormat = mapDateFormatSelect ? mapDateFormatSelect.value : 'seconds';

        if (wellIdx === dateIdx || wellIdx === volumeIdx || dateIdx === volumeIdx) {
            alert('Each mapping selector must map to a unique column.');
            return;
        }

        const rawWells = {};
        AppState.rawRows.forEach(row => {
            if (row.length <= Math.max(wellIdx, dateIdx, volumeIdx)) return;

            const wellId = row[wellIdx].trim();
            const volume = parseFloat(row[volumeIdx]);
            
            let timestamp;
            const rawTime = row[dateIdx];

            if (timeFormat === 'excel') {
                // Excel serial date stored as raw integer days since Jan 1, 1900
                // Keep as-is; delta will be computed directly in days below
                timestamp = parseFloat(rawTime);
            } else {
                // seconds: raw Unix seconds
                timestamp = parseFloat(rawTime);
            }

            if (wellId === '' || isNaN(timestamp) || isNaN(volume) || volume <= 0) return;

            if (!rawWells[wellId]) {
                rawWells[wellId] = {};
            }

            // Sum volume for duplicate timestamps
            if (rawWells[wellId][timestamp]) {
                rawWells[wellId][timestamp] += volume;
            } else {
                rawWells[wellId][timestamp] = volume;
            }
        });

        const defaultSensitivity = defaultSensitivitySlider ? parseInt(defaultSensitivitySlider.value) : 0;
        const processedWells = {};
        const wellIds = [];

        Object.keys(rawWells).forEach(wellId => {
            const timeVolumeMap = rawWells[wellId];
            const sortedTimestamps = Object.keys(timeVolumeMap).map(Number).sort((a, b) => a - b);

            if (sortedTimestamps.length < 2) return; // need at least 2 points

            // Calculate intervals and rates
            const rawPoints = sortedTimestamps.map(t => ({
                t: t,
                volume: timeVolumeMap[t]
            }));

            const dt = [];
            const n = rawPoints.length;
            for (let i = 0; i < n; i++) {
                if (i > 0) {
                    const diff = rawPoints[i].t - rawPoints[i - 1].t;
                    // 'excel' timestamps are already in days; 'seconds' need dividing
                    const deltaDays = (timeFormat === 'excel') ? diff : (diff / 86400);
                    dt.push(deltaDays); // delta in days
                }
            }

            const firstDt = dt.length > 0 ? dt[0] : 30; // fallback to 30 days
            dt.unshift(firstDt);

            const points = [];
            let cumulativeVolume = 0;
            let cumulativeDays = 0;

            for (let i = 0; i < n; i++) {
                const currentDt = dt[i] <= 0 ? 1 : dt[i]; // prevent 0 delta
                cumulativeDays += currentDt;
                cumulativeVolume += rawPoints[i].volume;
                const rate = rawPoints[i].volume / currentDt;

                points.push({
                    tDays: cumulativeDays,
                    rate: rate,
                    cumulativeVolume: cumulativeVolume,
                    mbtDays: rate > 0 ? cumulativeVolume / rate : 0
                });
            }

            // Ensure mathematical values are positive
            const validPoints = points.filter(p => p.tDays > 0 && p.rate > 0 && p.mbtDays > 0);

            if (validPoints.length < 2) return;

            const wellObj = {
                id: wellId,
                points: validPoints,
                outlierSensitivity: defaultSensitivity,
                timeBasis: 'production',
                bdfFitReal: null,
                outliersReal: [],
                bdfFitMbt: null,
                outliersMbt: []
            };

            // Pre-calculate regression results for both time bases
            calculateWellRegressions(wellObj);

            processedWells[wellId] = wellObj;
            wellIds.push(wellId);
        });

        if (wellIds.length === 0) {
            alert('No valid well datasets found. Ensure timestamps and volumes are positive numbers.');
            return;
        }

        wellIds.sort();

        AppState.wells = processedWells;
        AppState.wellIds = wellIds;
        AppState.currentWellIndex = 0;

        // Enable navigation
        navAnalysisBtn.removeAttribute('disabled');
        navSummaryBtn.removeAttribute('disabled');

        // Go to analysis view
        switchView('view-analysis');
    });

    // --- REGRESSION LOGIC ---
    const performPiecewiseRegression = (points) => {
        if (points.length < 2) return null;

        const logX = points.map(p => Math.log(p.x));
        const logY = points.map(p => Math.log(p.y));
        const n = points.length;

        let bestSSE = Infinity;
        let bestIP = 0;
        let bestTTD = 0;

        const minX = points[0].x;
        const maxX = points[n - 1].x;
        
        const logMin = Math.log(minX);
        const logMax = Math.log(maxX);
        const candidates = [];
        for (let i = 0; i <= 50; i++) {
            candidates.push(Math.exp(logMin + (i / 50) * (logMax - logMin)));
        }

        candidates.forEach(ttd => {
            let sumLogYMinusK = 0;
            const k = points.map(p => {
                if (p.x <= ttd) {
                    return -0.5 * Math.log(p.x);
                } else {
                    return 0.5 * Math.log(ttd) - 1.0 * Math.log(p.x);
                }
            });

            for (let i = 0; i < n; i++) {
                sumLogYMinusK += (logY[i] - k[i]);
            }
            const logIP = sumLogYMinusK / n;
            const ip = Math.exp(logIP);

            let sse = 0;
            for (let i = 0; i < n; i++) {
                const logYCalc = logIP + k[i];
                const diff = logY[i] - logYCalc;
                sse += diff * diff;
            }

            if (sse < bestSSE) {
                bestSSE = sse;
                bestIP = ip;
                bestTTD = ttd;
            }
        });

        return { ip: bestIP, ttd: bestTTD, sse: bestSSE };
    };

    const performIterativeRegression = (points, sensitivityPercent) => {
        if (points.length < 2) return { regression: null, outliers: [] };

        const numOutliersToExclude = Math.floor(points.length * (sensitivityPercent / 100));
        
        let currentActivePoints = [...points];
        let outliers = [];
        let finalRegression = null;

        for (let step = 0; step < 5; step++) {
            const regression = performPiecewiseRegression(currentActivePoints);
            if (!regression) break;
            
            finalRegression = regression;
            if (numOutliersToExclude === 0 || step === 4) break;

            const residuals = points.map(p => {
                const logX = Math.log(p.x);
                const logY = Math.log(p.y);
                const logIP = Math.log(regression.ip);
                
                let logYCalc;
                if (p.x <= regression.ttd) {
                    logYCalc = logIP - 0.5 * logX;
                } else {
                    logYCalc = logIP + 0.5 * Math.log(regression.ttd) - 1.0 * logX;
                }
                
                return { point: p, error: Math.pow(logY - logYCalc, 2) };
            });

            residuals.sort((a, b) => b.error - a.error);
            outliers = residuals.slice(0, numOutliersToExclude).map(r => r.point);
            currentActivePoints = points.filter(p => !outliers.includes(p));
            
            if (currentActivePoints.length < 2) break;
        }

        return { regression: finalRegression, outliers: outliers };
    };

    const calculateWellRegressions = (well) => {
        // 1. Real time points
        const realPoints = well.points.map(p => ({ x: p.tDays, y: p.rate }));
        const realResult = performIterativeRegression(realPoints, well.outlierSensitivity);
        well.bdfFitReal = realResult.regression;
        well.outliersReal = realResult.outliers;

        // 2. MBT points
        const mbtPoints = well.points.map(p => ({ x: p.mbtDays, y: p.rate }));
        const mbtResult = performIterativeRegression(mbtPoints, well.outlierSensitivity);
        well.bdfFitMbt = mbtResult.regression;
        well.outliersMbt = mbtResult.outliers;
    };

    const generateTrendLine = (points, regression) => {
        if (!regression) return [];

        const { ip, ttd } = regression;
        const minX = Math.min(...points.map(p => p.x));
        const maxX = Math.max(...points.map(p => p.x));
        
        const linePoints = [];
        linePoints.push({ x: minX, y: ip * Math.pow(minX, -0.5) });

        if (ttd > minX && ttd < maxX) {
            linePoints.push({ x: ttd, y: ip * Math.pow(ttd, -0.5) });
        }

        if (maxX > ttd) {
            linePoints.push({ x: maxX, y: ip * Math.pow(ttd, 0.5) * Math.pow(maxX, -1.0) });
        } else {
            linePoints.push({ x: maxX, y: ip * Math.pow(maxX, -0.5) });
        }

        return linePoints;
    };

    // --- WELL ANALYSIS VIEW LOGIC ---
    const getFilteredWellIds = () => {
        const query = wellSearchInput.value.toLowerCase().trim();
        if (query === '') return AppState.wellIds;
        return AppState.wellIds.filter(id => id.toLowerCase().includes(query));
    };

    const renderWellsList = () => {
        const filteredIds = getFilteredWellIds();
        wellsListUl.innerHTML = '';
        
        filteredIds.forEach(id => {
            const li = document.createElement('li');
            li.className = 'well-item';
            if (AppState.wellIds[AppState.currentWellIndex] === id) {
                li.classList.add('active');
            }
            li.textContent = id;
            li.title = id;
            li.addEventListener('click', () => {
                const idx = AppState.wellIds.indexOf(id);
                if (idx !== -1) {
                    AppState.currentWellIndex = idx;
                    renderWellsList();
                    loadActiveWell();
                }
            });
            wellsListUl.appendChild(li);
        });

        // Update footer info
        const totalFiltered = filteredIds.length;
        const total = AppState.wellIds.length;
        currentWellInfoSpan.textContent = total > 0 ? `${AppState.currentWellIndex + 1} / ${total}` : '0 / 0';
    };

    wellSearchInput.addEventListener('input', () => {
        renderWellsList();
    });

    const loadActiveWell = () => {
        const activeWellId = AppState.wellIds[AppState.currentWellIndex];
        if (!activeWellId) return;

        const well = AppState.wells[activeWellId];
        activeWellNameH2.textContent = well.id;
        activeWellBadge.textContent = `Well ID: ${well.id}`;

        // Sync local inputs
        sensitivitySlider.value = well.outlierSensitivity;
        sensitivityValueSpan.textContent = `${well.outlierSensitivity}%`;

        if (well.timeBasis === 'production') {
            timeProdRadio.checked = true;
        } else {
            timeMbtRadio.checked = true;
        }

        renderActiveWellStats(well);
        renderActiveWellChart(well);
    };

    const renderActiveWellStats = (well) => {
        pointCountAnalysisSpan.textContent = well.points.length;
        
        if (well.bdfFitReal) {
            ipValAnalysisSpan.textContent = formatNumber(well.bdfFitReal.ip);
            ttdRealValSpan.textContent = formatNumber(daysToMonths(well.bdfFitReal.ttd), 2);
            ttdRealDaysSpan.textContent = formatNumber(well.bdfFitReal.ttd, 1);
        } else {
            ipValAnalysisSpan.textContent = '--';
            ttdRealValSpan.textContent = '--';
            ttdRealDaysSpan.textContent = '--';
        }

        if (well.bdfFitMbt) {
            ttdMbtValSpan.textContent = formatNumber(daysToMonths(well.bdfFitMbt.ttd), 2);
            ttdMbtDaysSpan.textContent = formatNumber(well.bdfFitMbt.ttd, 1);
        } else {
            ttdMbtValSpan.textContent = '--';
            ttdMbtDaysSpan.textContent = '--';
        }

        const activeFit = well.timeBasis === 'production' ? well.bdfFitReal : well.bdfFitMbt;
        sseValAnalysisSpan.textContent = activeFit ? activeFit.sse.toFixed(6) : '--';
    };

    const renderActiveWellChart = (well) => {
        const ctx = document.getElementById('analysis-chart').getContext('2d');
        
        if (AppState.charts.analysis) {
            AppState.charts.analysis.destroy();
        }

        const isRealTime = well.timeBasis === 'production';
        
        // Prepare datasets
        const activeX = isRealTime ? 'tDays' : 'mbtDays';
        const outliers = isRealTime ? well.outliersReal : well.outliersMbt;
        
        const activePoints = well.points.filter(p => {
            return !outliers.some(o => o.x === p[activeX] && o.y === p.rate);
        });

        const activeData = activePoints.map(p => ({ x: p[activeX], y: p.rate }));
        const outlierData = outliers;
        
        const regressionResult = isRealTime ? well.bdfFitReal : well.bdfFitMbt;
        const trendData = generateTrendLine(activeData, regressionResult);

        AppState.charts.analysis = new Chart(ctx, {
            type: 'scatter',
            data: {
                datasets: [
                    {
                        label: 'Production Data',
                        data: activeData,
                        backgroundColor: '#3b82f6',
                        borderColor: '#2563eb',
                        pointRadius: 4,
                        pointHoverRadius: 6,
                    },
                    {
                        label: 'Outliers Excluded',
                        data: outlierData,
                        backgroundColor: 'rgba(255, 255, 255, 0.12)',
                        borderColor: 'rgba(255, 255, 255, 0.25)',
                        pointRadius: 4,
                        pointHoverRadius: 5,
                    },
                    {
                        label: 'BDF Trend Line',
                        data: trendData,
                        type: 'line',
                        borderColor: '#10b981',
                        borderWidth: 2,
                        pointRadius: 0,
                        fill: false,
                        showLine: true
                    }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                scales: {
                    x: {
                        type: 'logarithmic',
                        title: { 
                            display: true, 
                            text: isRealTime ? 'Production Time (Days)' : 'Material Balance Time (Days)', 
                            color: '#94a3b8' 
                        },
                        grid: { color: 'rgba(255, 255, 255, 0.04)' },
                        ticks: { color: '#94a3b8' }
                    },
                    y: {
                        type: 'logarithmic',
                        title: { display: true, text: 'Production Rate (bbl/d)', color: '#94a3b8' },
                        grid: { color: 'rgba(255, 255, 255, 0.04)' },
                        ticks: { color: '#94a3b8' }
                    }
                },
                plugins: {
                    legend: { display: false },
                    tooltip: {
                        backgroundColor: 'rgba(11, 14, 20, 0.95)',
                        titleColor: '#fff',
                        bodyColor: '#94a3b8',
                        borderColor: 'rgba(255, 255, 255, 0.1)',
                        borderWidth: 1
                    }
                }
            }
        });
    };

    // Sensitivity Slider Event Listeners
    sensitivitySlider.addEventListener('input', () => {
        const val = parseInt(sensitivitySlider.value);
        sensitivityValueSpan.textContent = `${val}%`;

        const activeWellId = AppState.wellIds[AppState.currentWellIndex];
        if (activeWellId) {
            const well = AppState.wells[activeWellId];
            well.outlierSensitivity = val;
            calculateWellRegressions(well);
            renderActiveWellStats(well);
            renderActiveWellChart(well);
        }
    });

    // Time Basis Change Event Listeners
    [timeProdRadio, timeMbtRadio].forEach(radio => {
        radio.addEventListener('change', () => {
            const activeWellId = AppState.wellIds[AppState.currentWellIndex];
            if (activeWellId) {
                const well = AppState.wells[activeWellId];
                well.timeBasis = timeProdRadio.checked ? 'production' : 'mbt';
                renderActiveWellStats(well);
                renderActiveWellChart(well);
            }
        });
    });

    // Pagination Click Listeners
    prevWellBtn.addEventListener('click', () => {
        if (AppState.wellIds.length === 0) return;
        AppState.currentWellIndex = (AppState.currentWellIndex - 1 + AppState.wellIds.length) % AppState.wellIds.length;
        renderWellsList();
        loadActiveWell();
    });

    nextWellBtn.addEventListener('click', () => {
        if (AppState.wellIds.length === 0) return;
        AppState.currentWellIndex = (AppState.currentWellIndex + 1) % AppState.wellIds.length;
        renderWellsList();
        loadActiveWell();
    });

    // --- SUMMARY VIEW LOGIC ---
    const generateCDFSeries = (values) => {
        const sorted = values.filter(v => v !== null && !isNaN(v) && v > 0).sort((a, b) => a - b);
        if (sorted.length === 0) return [];
        const n = sorted.length;
        return sorted.map((v, idx) => ({
            x: v,
            y: ((idx + 1) / n) * 100
        }));
    };

    const renderSummaryPage = () => {
        const totalWells = AppState.wellIds.length;
        summaryTotalWellsSpan.textContent = totalWells;

        if (totalWells === 0) return;

        // Calculate global values
        let sumIp = 0;
        let sumTtdReal = 0;
        let sumTtdMbt = 0;
        let countIp = 0;
        let countTtdReal = 0;
        let countTtdMbt = 0;

        const ttdRealValues = [];
        const ttdMbtValues = [];

        AppState.wellIds.forEach(id => {
            const well = AppState.wells[id];
            
            if (well.bdfFitReal) {
                sumIp += well.bdfFitReal.ip;
                countIp++;
                const ttdRealMonths = daysToMonths(well.bdfFitReal.ttd);
                sumTtdReal += ttdRealMonths;
                countTtdReal++;
                ttdRealValues.push(ttdRealMonths);
            }
            if (well.bdfFitMbt) {
                const ttdMbtMonths = daysToMonths(well.bdfFitMbt.ttd);
                sumTtdMbt += ttdMbtMonths;
                countTtdMbt++;
                ttdMbtValues.push(ttdMbtMonths);
            }
        });

        summaryAvgIpSpan.textContent = countIp > 0 ? formatNumber(sumIp / countIp) : '--';
        summaryAvgTtdRealSpan.textContent = countTtdReal > 0 ? formatNumber(sumTtdReal / countTtdReal, 2) + ' mo' : '--';
        summaryAvgTtdMbtSpan.textContent = countTtdMbt > 0 ? formatNumber(sumTtdMbt / countTtdMbt, 2) + ' mo' : '--';

        // Render CDF Charts
        renderCDFChart('cdf-real-chart', 'cdfReal', ttdRealValues, 'CDF: Time to BDF (Real Time)', '#3b82f6');
        renderCDFChart('cdf-mbt-chart', 'cdfMbt', ttdMbtValues, 'CDF: Time to BDF (MBT)', '#10b981');

        // Render Summary Table Rows
        renderSummaryTable();
    };

    const renderCDFChart = (canvasId, chartKey, values, label, color) => {
        const ctx = document.getElementById(canvasId).getContext('2d');
        if (AppState.charts[chartKey]) {
            AppState.charts[chartKey].destroy();
        }

        const cdfData = generateCDFSeries(values);

        AppState.charts[chartKey] = new Chart(ctx, {
            type: 'scatter',
            data: {
                datasets: [{
                    label: label,
                    data: cdfData,
                    showLine: true,
                    borderColor: color,
                    backgroundColor: color,
                    borderWidth: 2,
                    pointRadius: 3,
                    fill: false,
                    tension: 0.1
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                scales: {
                    x: {
                        type: 'linear',
                        title: { display: true, text: 'Time to BDF (Months)', color: '#94a3b8' },
                        grid: { color: 'rgba(255, 255, 255, 0.04)' },
                        ticks: { color: '#94a3b8' }
                    },
                    y: {
                        min: 0,
                        max: 100,
                        title: { display: true, text: 'Cumulative Percentile (%)', color: '#94a3b8' },
                        grid: { color: 'rgba(255, 255, 255, 0.04)' },
                        ticks: { color: '#94a3b8' }
                    }
                },
                plugins: {
                    legend: { display: false }
                }
            }
        });
    };

    const renderSummaryTable = () => {
        let html = '';
        AppState.wellIds.forEach(id => {
            const well = AppState.wells[id];
            const pointsCount = well.points.length;
            const ip = well.bdfFitReal ? formatNumber(well.bdfFitReal.ip) : '--';
            const ttdRealD = well.bdfFitReal ? formatNumber(well.bdfFitReal.ttd, 1) : '--';
            const ttdRealM = well.bdfFitReal ? formatNumber(daysToMonths(well.bdfFitReal.ttd), 2) : '--';
            const ttdMbtD = well.bdfFitMbt ? formatNumber(well.bdfFitMbt.ttd, 1) : '--';
            const ttdMbtM = well.bdfFitMbt ? formatNumber(daysToMonths(well.bdfFitMbt.ttd), 2) : '--';
            
            const activeFit = well.timeBasis === 'production' ? well.bdfFitReal : well.bdfFitMbt;
            const sse = activeFit ? activeFit.sse.toFixed(4) : '--';

            html += `
                <tr>
                    <td style="font-weight:600; color:#f1f5f9;">${well.id}</td>
                    <td>${pointsCount}</td>
                    <td>${ip}</td>
                    <td>${ttdRealD}</td>
                    <td style="color:#3b82f6; font-weight:600;">${ttdRealM}</td>
                    <td>${ttdMbtD}</td>
                    <td style="color:#10b981; font-weight:600;">${ttdMbtM}</td>
                    <td>${sse}</td>
                    <td><span class="badge secondary">Planned</span></td>
                </tr>
            `;
        });
        summaryTableBody.innerHTML = html;
    };

    // --- EXPORT CSV LOGIC ---
    exportCsvBtn.addEventListener('click', () => {
        if (AppState.wellIds.length === 0) return;

        let csv = 'Well ID,Points Count,IP (qi bbl/d),TTD Real (Days),TTD Real (Months),TTD MBT (Days),TTD MBT (Months),Real SSE,MBT SSE,ARPS Decline Status\n';
        
        AppState.wellIds.forEach(id => {
            const well = AppState.wells[id];
            const pointsCount = well.points.length;
            
            const ip = well.bdfFitReal ? well.bdfFitReal.ip.toFixed(2) : 'N/A';
            const ttdRealD = well.bdfFitReal ? well.bdfFitReal.ttd.toFixed(2) : 'N/A';
            const ttdRealM = well.bdfFitReal ? daysToMonths(well.bdfFitReal.ttd).toFixed(2) : 'N/A';
            
            const ttdMbtD = well.bdfFitMbt ? well.bdfFitMbt.ttd.toFixed(2) : 'N/A';
            const ttdMbtM = well.bdfFitMbt ? daysToMonths(well.bdfFitMbt.ttd).toFixed(2) : 'N/A';
            
            const sseReal = well.bdfFitReal ? well.bdfFitReal.sse.toFixed(6) : 'N/A';
            const sseMbt = well.bdfFitMbt ? well.bdfFitMbt.sse.toFixed(6) : 'N/A';

            // Escape well ID in quotes
            const escapedId = `"${well.id.replace(/"/g, '""')}"`;
            
            csv += `${escapedId},${pointsCount},${ip},${ttdRealD},${ttdRealM},${ttdMbtD},${ttdMbtM},${sseReal},${sseMbt},Planned\n`;
        });

        const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.setAttribute('href', url);
        link.setAttribute('download', 'wells_decline_analysis_summary.csv');
        link.style.visibility = 'hidden';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    });
});
