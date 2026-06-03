// App State
let holidays = {
    "2025-01-01": "元日",
    "2025-01-13": "成人の日",
    "2025-02-11": "建国記念の日",
    "2025-02-23": "天皇誕生日",
    "2025-02-24": "天皇誕生日 振替休日",
    "2025-03-20": "春分の日",
    "2025-04-29": "昭和の日",
    "2025-05-03": "憲法記念日",
    "2025-05-04": "みどりの日",
    "2025-05-05": "こどもの日",
    "2025-05-06": "みどりの日 振替休日",
    "2025-07-21": "海の日",
    "2025-08-11": "山の日",
    "2025-09-15": "敬老の日",
    "2025-09-23": "秋分の日",
    "2025-10-13": "スポーツの日",
    "2025-11-03": "文化の日",
    "2025-11-23": "勤労感謝の日",
    "2025-11-24": "勤労感謝の日 振替休日",
    "2026-01-01": "元日",
    "2026-01-12": "成人の日",
    "2026-02-11": "建国記念の日",
    "2026-02-23": "天皇誕生日",
    "2026-03-20": "春分の日",
    "2026-04-29": "昭和の日",
    "2026-05-03": "憲法記念日",
    "2026-05-04": "みどりの日",
    "2026-05-05": "こどもの日",
    "2026-05-06": "憲法記念日 振替休日",
    "2026-07-20": "海の日",
    "2026-08-11": "山の日",
    "2026-09-21": "敬老の日",
    "2026-09-22": "国民の休日",
    "2026-09-23": "秋分の日",
    "2026-10-12": "スポーツの日",
    "2026-11-03": "文化の日",
    "2026-11-23": "勤労感謝の日",
    "2027-01-01": "元日",
    "2027-01-11": "成人の日",
    "2027-02-11": "建国記念の日",
    "2027-02-23": "天皇誕生日",
    "2027-03-21": "春分の日",
    "2027-03-22": "春分の日 振替休日",
    "2027-04-29": "昭和の日",
    "2027-05-03": "憲法記念日",
    "2027-05-04": "みどりの日",
    "2027-05-05": "こどもの日",
    "2027-07-19": "海の日",
    "2027-08-11": "山の日",
    "2027-09-20": "敬老の日",
    "2027-09-23": "秋分の日",
    "2027-10-11": "スポーツの日",
    "2027-11-03": "文化の日",
    "2027-11-23": "勤労感謝の日"
};

let currentFileData = null;
let currentIssues = [];
let allRowsData = [];

// Initialize
document.addEventListener('DOMContentLoaded', () => {
    setupEventListeners();
});

// Setup Drag and Drop / Click Upload
function setupEventListeners() {
    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('file-input');
    
    // Drag events
    ['dragenter', 'dragover'].forEach(eventName => {
        dropZone.addEventListener(eventName, (e) => {
            e.preventDefault();
            dropZone.classList.add('dragover');
        }, false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, (e) => {
            e.preventDefault();
            dropZone.classList.remove('dragover');
        }, false);
    });

    dropZone.addEventListener('drop', (e) => {
        const dt = e.dataTransfer;
        const files = dt.files;
        if (files.length > 0) {
            handleFile(files[0]);
        }
    });

    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            handleFile(e.target.files[0]);
        }
    });

    // Tab switcher
    const tabButtons = document.querySelectorAll('.tab-btn');
    tabButtons.forEach(btn => {
        btn.addEventListener('click', () => {
            tabButtons.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            filterIssues(btn.dataset.tab);
        });
    });
}

// Handle uploaded file
function handleFile(file) {
    if (!file.name.endsWith('.xls') && !file.name.endsWith('.xlsx')) {
        alert('Excelファイル (.xls または .xlsx) をアップロードしてください。');
        return;
    }

    const loadStateEl = document.getElementById('loading-state');
    const fileNameEl = document.getElementById('loading-file-name');
    const stepTextEl = document.getElementById('loading-step-text');

    // Show loading safely
    if (fileNameEl) fileNameEl.textContent = `対象ファイル: ${file.name}`;
    if (stepTextEl) stepTextEl.textContent = 'エクセルファイルを読み込んでいます...';
    if (loadStateEl) loadStateEl.classList.remove('hidden');
    
    const emptyState = document.getElementById('empty-state');
    if (emptyState) emptyState.classList.add('hidden');
    
    const dashboardResult = document.getElementById('dashboard-result');
    if (dashboardResult) dashboardResult.classList.add('hidden');

    const reader = new FileReader();
    
    reader.onerror = function(err) {
        console.error('File reader error:', err);
        alert('ファイルの読み込みに失敗しました。');
        if (loadStateEl) loadStateEl.classList.add('hidden');
        if (emptyState) emptyState.classList.remove('hidden');
    };

    reader.onload = function(e) {
        try {
            const arrayBuffer = e.target.result; // Store immediately!
            if (!arrayBuffer) {
                throw new Error('読み込まれたデータが空です。');
            }
            
            // Step 1: Wait for repaint and update message
            setTimeout(() => {
                if (stepTextEl) stepTextEl.textContent = '経費精算ルールと明細データの照合中...';
                
                // Step 2: Perform processing in the next tick to ensure the DOM updates
                setTimeout(() => {
                    try {
                        const data = new Uint8Array(arrayBuffer);
                        const workbook = XLSX.read(data, {type: 'array', cellDates: true, cellNF: false, cellText: false});
                        
                        // Assume the first sheet is the target
                        const firstSheetName = workbook.SheetNames[0];
                        const worksheet = workbook.Sheets[firstSheetName];
                        
                        // Convert to 2D Array to handle freely
                        const sheetData = XLSX.utils.sheet_to_json(worksheet, {header: 1, defval: ''});
                        
                        processExcelData(sheetData, file.name);
                    } catch (error) {
                        console.error('File parsing error:', error);
                        alert('ファイルの読み込み中にエラーが発生しました。ファイル形式を確認してください。');
                        if (loadStateEl) loadStateEl.classList.add('hidden');
                        if (emptyState) emptyState.classList.remove('hidden');
                    }
                }, 300); // 300ms pause to show Rule Verification step clearly
            }, 300); // 300ms pause to show Reading file step clearly
        } catch (outerError) {
            console.error('Outer reader error:', outerError);
            alert('ファイルの処理開始時にエラーが発生しました。');
            if (loadStateEl) loadStateEl.classList.add('hidden');
            if (emptyState) emptyState.classList.remove('hidden');
        }
    };
    reader.readAsArrayBuffer(file);
}

// Parse route string: extracts start/end station and whether it is a round trip
function parseRoute(text) {
    let isRound = text.includes('⇔');
    let separator = isRound ? '⇔' : '→';
    if (!text.includes(separator)) return null;
    
    let parts = text.split(separator);
    let leftPart = parts[0].trim();
    let rightPart = parts[1].trim();
    
    let startStation = extractLastParentheses(leftPart);
    let endStation = extractFirstParentheses(rightPart);
    
    return {
        start: startStation,
        end: endStation,
        isRound: isRound
    };
}

function extractLastParentheses(str) {
    str = str.trim();
    if (!str.endsWith(')')) {
        let lastIdx = str.lastIndexOf(')');
        if (lastIdx === -1) return str;
        str = str.substring(0, lastIdx + 1);
    }
    
    let depth = 0;
    let startIdx = -1;
    for (let i = str.length - 1; i >= 0; i--) {
        if (str[i] === ')') depth++;
        else if (str[i] === '(') {
            depth--;
            if (depth === 0) {
                startIdx = i;
                break;
            }
        }
    }
    if (startIdx !== -1) {
        return str.substring(startIdx + 1, str.length - 1).trim();
    }
    return str;
}

function extractFirstParentheses(str) {
    str = str.trim();
    let startIdx = str.indexOf('(');
    if (startIdx === -1) return str;
    
    let depth = 0;
    let endIdx = -1;
    for (let i = startIdx; i < str.length; i++) {
        if (str[i] === '(') depth++;
        else if (str[i] === ')') {
            depth--;
            if (depth === 0) {
                endIdx = i;
                break;
            }
        }
    }
    if (endIdx !== -1) {
        return str.substring(startIdx + 1, endIdx).trim();
    }
    return str;
}

// Format date to YYYY-MM-DD
function normalizeDate(cellValue) {
    if (!cellValue) return null;
    
    if (cellValue instanceof Date) {
        // Formatted JS Date
        let y = cellValue.getFullYear();
        let m = String(cellValue.getMonth() + 1).padStart(2, '0');
        let d = String(cellValue.getDate()).padStart(2, '0');
        return `${y}-${m}-${d}`;
    }
    
    let str = String(cellValue).trim();
    // Check if it is standard date string like YYYY/MM/DD(木)
    // Strip day of week like (木)
    str = str.replace(/\([^)]+\)/g, '');
    
    // Replace slashes with dashes
    str = str.replace(/\//g, '-');
    
    // Check if format is YYYY-MM-DD
    let match = str.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (match) {
        let y = match[1];
        let m = match[2].padStart(2, '0');
        let d = match[3].padStart(2, '0');
        return `${y}-${m}-${d}`;
    }
    return null;
}

// Get Day of Week Name
function getDayOfWeekStr(dateStr) {
    const days = ['日', '月', '火', '水', '木', '金', '土'];
    const d = new Date(dateStr);
    return days[d.getDay()];
}

// Main Process
function processExcelData(sheetData, filename) {
    currentIssues = [];
    allRowsData = [];

    // 1. Detect Metadata
    let applicantName = '不明';
    let employeeId = '不明';
    let applicationDate = '不明';
    let deptName = '不明';
    let totalAmountClaimed = 0;

    // Standard positioning checks, but search dynamically just in case
    for (let r = 0; r < Math.min(sheetData.length, 15); r++) {
        const row = sheetData[r];
        for (let c = 0; c < row.length; c++) {
            const val = String(row[c]).trim();
            if (val === '氏名') applicantName = String(row[c+1] || '').trim();
            if (val === '社員番号') employeeId = String(row[c+1] || '').trim();
            if (val === '申請日') {
                let rawDate = row[c+1];
                if (rawDate instanceof Date) {
                    applicationDate = normalizeDate(rawDate);
                } else {
                    applicationDate = String(rawDate || '').trim();
                }
            }
            if (val === '所属') deptName = String(row[c+1] || '').trim();
            if (val === '合計金額') totalAmountClaimed = parseFloat(row[c+1]) || 0;
        }
    }

    // 2. Find 明細 Table Headers
    let headerRowIdx = -1;
    let colIdx = {
        no: -1,
        date: -1,
        payee: -1,
        category: -1,
        projId: -1,
        amount: -1,
        remarks: -1
    };

    for (let r = 0; r < sheetData.length; r++) {
        const row = sheetData[r];
        if (row.includes('番号') && row.includes('日付') && row.includes('経費科目')) {
            headerRowIdx = r;
            colIdx.no = row.indexOf('番号');
            colIdx.date = row.indexOf('日付');
            colIdx.payee = row.indexOf('支払先内容');
            colIdx.category = row.indexOf('経費科目');
            colIdx.projId = row.indexOf('プロジェクトID');
            colIdx.amount = row.indexOf('金額');
            colIdx.remarks = row.indexOf('備考・メモ');
            break;
        }
    }

    if (headerRowIdx === -1) {
        alert('経費明細テーブル（「番号」「日付」「経費科目」を含む行）が見つかりませんでした。');
        document.getElementById('loading-state').classList.add('hidden');
        document.getElementById('empty-state').classList.remove('hidden');
        return;
    }

    // 3. Extract Data Rows
    let calculatedTotal = 0;
    for (let r = headerRowIdx + 1; r < sheetData.length; r++) {
        const row = sheetData[r];
        // Must have at least a valid Number or Date to be considered a row
        const noVal = row[colIdx.no];
        const dateVal = row[colIdx.date];
        
        if (noVal === '' && dateVal === '') continue; // Skip empty rows

        const parsedNo = parseFloat(noVal);
        if (isNaN(parsedNo) && dateVal === '') continue; // Skip totals or non-data rows

        const amount = parseFloat(row[colIdx.amount]) || 0;
        calculatedTotal += amount;

        const rowObj = {
            excelRowIdx: r, // 0-indexed row index in Excel
            no: noVal,
            rawDate: dateVal,
            dateStr: normalizeDate(dateVal),
            payee: String(row[colIdx.payee] || '').trim(),
            category: String(row[colIdx.category] || '').trim(),
            projId: String(row[colIdx.projId] || '').trim(),
            amount: amount,
            remarks: String(row[colIdx.remarks] || '').trim(),
            issues: []
        };
        allRowsData.push(rowObj);
    }

    // 4. Run Checks
    runValidationChecks();

    // 5. Update UI
    document.getElementById('loading-state').classList.add('hidden');
    document.getElementById('dashboard-result').classList.remove('hidden');

    // Fill Summary
    document.getElementById('val-applicant').textContent = applicantName;
    document.getElementById('val-emp-id').textContent = employeeId;
    document.getElementById('val-date').textContent = applicationDate;
    document.getElementById('val-dept').textContent = deptName;
    document.getElementById('val-amount').textContent = `¥${totalAmountClaimed.toLocaleString()}`;
    
    // Verification check for totals
    if (Math.abs(calculatedTotal - totalAmountClaimed) > 1) {
        addGeneralIssue('error', '金額不一致', `ファイルの申告合計金額（¥${totalAmountClaimed.toLocaleString()}）と明細行の積算合計金額（¥${calculatedTotal.toLocaleString()}）が一致しません。`);
    }

    renderIssues();
    renderTable(colIdx);
}

// Validation Engine
function runValidationChecks() {
    // Helper to add row issue
    function addRowIssue(row, type, title, desc, colName) {
        const issue = {
            type: type, // 'error' or 'warning'
            rowIdx: row.excelRowIdx,
            excelNo: row.no,
            date: row.rawDate,
            category: row.category,
            title: title,
            desc: desc,
            colName: colName,
            cellName: `${getColLetter(colName)}${row.excelRowIdx + 1}`
        };
        row.issues.push(issue);
        currentIssues.push(issue);
    }

    // First, let's identify HOME and COMPANY stations from commuting expenses
    let homeStations = {};
    let companyStations = {};

    allRowsData.forEach(row => {
        if (row.category === '通勤費') {
            const route = parseRoute(row.payee);
            if (route) {
                if (row.payee.includes('出社')) {
                    homeStations[route.start] = (homeStations[route.start] || 0) + 1;
                    companyStations[route.end] = (companyStations[route.end] || 0) + 1;
                } else if (row.payee.includes('帰宅')) {
                    companyStations[route.start] = (companyStations[route.start] || 0) + 1;
                    homeStations[route.end] = (homeStations[route.end] || 0) + 1;
                } else {
                    // Just add both based on ⇔
                    if (route.isRound) {
                        // For round trip commuting, standard is usually Home ⇔ Company.
                        // Let's assume start is Home, end is Company for sorting later.
                        homeStations[route.start] = (homeStations[route.start] || 0) + 1;
                        companyStations[route.end] = (companyStations[route.end] || 0) + 1;
                    }
                }
            }
        }
    });

    // Detect primary Home/Company station (highest frequency)
    let homeStation = Object.keys(homeStations).reduce((a, b) => homeStations[a] > homeStations[b] ? a : b, null);
    let companyStation = Object.keys(companyStations).reduce((a, b) => companyStations[a] > companyStations[b] ? a : b, null);

    console.log(`Detected Station Configuration - Home: ${homeStation}, Company: ${companyStation}`);

    // Group rows by date for commuting completeness check
    let rowsByDate = {};
    allRowsData.forEach(row => {
        if (row.dateStr) {
            if (!rowsByDate[row.dateStr]) {
                rowsByDate[row.dateStr] = [];
            }
            rowsByDate[row.dateStr].push(row);
        }
    });

    // Tracking for transport project ID consistency
    let transportProjIds = {};

    // Validate Row by Row
    allRowsData.forEach(row => {
        // --- 1. 全経費共通ルール ---
        
        // 1.1 プロジェクトIDチェック
        if (!row.projId) {
            addRowIssue(row, 'error', 'プロジェクトID未入力', 'プロジェクトIDが入力されていません。', 'プロジェクトID');
        } else if (!row.projId.includes('電子雑誌') && !row.projId.includes('飲食')) {
            addRowIssue(row, 'error', 'プロジェクトID不正', 'プロジェクトIDが「電子雑誌」または「飲食」のいずれでもありません。', 'プロジェクトID');
        }

        // 1.2 土日祝日チェック
        if (row.dateStr) {
            const dateObj = new Date(row.dateStr);
            const dayOfWeek = dateObj.getDay(); // 0: Sun, 6: Sat
            const isWeekend = (dayOfWeek === 0 || dayOfWeek === 6);
            const holidayName = holidays[row.dateStr];
            
            if (isWeekend || holidayName) {
                const dayName = isWeekend ? getDayOfWeekStr(row.dateStr) + '曜日' : `祝日(${holidayName})`;
                addRowIssue(row, 'warning', '休日経費申請', `土日・祝日（${dayName}）に経費が申請されています。休日出勤の精算かどうか確認してください。`, '日付');
            }
        } else {
            addRowIssue(row, 'error', '日付フォーマットエラー', '日付を正しく解析できませんでした（例: 2026/05/07(木)）。', '日付');
        }

        // --- 2. 交通費ルール ---
        // Excludes parking/gasoline
        const isNormalTransport = row.category.startsWith('交通費') && 
                                  !row.category.includes('駐車場') && 
                                  !row.category.includes('ガソリン');
                                  
        if (isNormalTransport) {
            // 2.1 利用目的チェック
            if (!row.remarks.includes('【旅色営業のため】')) {
                addRowIssue(row, 'warning', '利用目的不備', '交通費の備考欄に利用目的「【旅色営業のため】」が含まれていません。', '備考・メモ');
            }
            
            // 2.2 プロジェクトID一貫性のための収集
            if (row.projId) {
                transportProjIds[row.projId] = (transportProjIds[row.projId] || 0) + 1;
            }

            // 2.3 新幹線チェック
            if (row.payee.includes('新幹線') || row.remarks.includes('新幹線')) {
                addRowIssue(row, 'error', '新幹線申請不可', '支払先または備考に「新幹線」の表記があります。別の精算経路、または事前承認が必要な項目です。', '支払先内容');
            }
        }

        // --- 4. 交通費（駐車場）ルール ---
        if (row.category.includes('交通費(駐車場)')) {
            // 4.1 支払先内容チェック (No routes)
            if (row.payee.includes('→') || row.payee.includes('⇔') || row.payee.includes('→') || row.payee.includes('⇔')) {
                addRowIssue(row, 'error', '駐車場支払先エラー', '支払先内容がルート表記（→や⇔）になっています。管理会社名を記載してください。', '支払先内容');
            }
            // 4.2 利用目的チェック
            if (!row.remarks.includes('【旅色営業のため】')) {
                addRowIssue(row, 'warning', '利用目的不備', '駐車場の備考欄に利用目的「【旅色営業のため】」が含まれていません。', '備考・メモ');
            }
        }

        // --- 5. 内部飲食代ルール ---
        if (row.category.startsWith('内部飲食代')) {
            // 5.1 必須項目チェック (利用目的、経費枠負担者、金額)
            const hasPurpose = row.remarks.match(/利用目的【[^】]+】/);
            const hasPayer = row.remarks.match(/経費枠負担者【[^】]+】/);
            const hasAmountTag = row.remarks.match(/金額【[^】]+】/);

            let missingTags = [];
            if (!hasPurpose) missingTags.push('利用目的');
            if (!hasPayer) missingTags.push('経費枠負担者');
            if (!hasAmountTag) missingTags.push('金額');

            if (missingTags.length > 0) {
                addRowIssue(row, 'error', '内部飲食記載不足', `内部飲食代の備考欄に必要な記載タグが不足しています: ${missingTags.join(', ')}`, '備考・メモ');
            }

            // 5.2 金額整合性チェック
            if (hasAmountTag) {
                // Extract amount from tag, e.g. "金額【11080円】" -> 11080
                const tagAmountMatch = hasAmountTag[0].match(/金額【(\d+)円】/);
                if (tagAmountMatch) {
                    const tagAmountVal = parseInt(tagAmountMatch[1], 10);
                    if (tagAmountVal !== row.amount) {
                        addRowIssue(row, 'error', '飲食金額不一致', `備考欄の金額（¥${tagAmountVal.toLocaleString()}）と金額列（¥${row.amount.toLocaleString()}）が一致しません。`, '備考・メモ');
                    }
                } else {
                    addRowIssue(row, 'error', '飲食金額タグ形式不正', '備考欄の金額タグの形式が不正です（例：金額【11080円】）。', '備考・メモ');
                }
            }
        }
    });

    // 2.2 プロジェクトID一貫性チェック（全体評価）
    const projIdKeys = Object.keys(transportProjIds);
    if (projIdKeys.length > 1) {
        // Multi project IDs used in transportation. Let's add warnings/errors to all transportation rows
        allRowsData.forEach(row => {
            const isNormalTransport = row.category.startsWith('交通費') && 
                                      !row.category.includes('駐車場') && 
                                      !row.category.includes('ガソリン');
            if (isNormalTransport) {
                addRowIssue(row, 'error', '交通費プロジェクトID不一致', `交通費のプロジェクトIDが一貫していません。明細内で複数のプロジェクトID（${projIdKeys.join(', ')}）が混在しています。`, 'プロジェクトID');
            }
        });
    }

    // --- 3. 通勤費ルール (日別統合判定) ---
    if (homeStation) {
        Object.keys(rowsByDate).forEach(dateStr => {
            const dayRows = rowsByDate[dateStr];
            
            // Check if there is any commuting/travel action on this day
            const hasAnyExpense = dayRows.length > 0;
            if (!hasAnyExpense) return;

            // Commuting details
            let hasRoundCommute = false;
            let hasMorningCommute = false;
            let hasEveningCommute = false;

            // Travel details
            let trainTravels = [];

            dayRows.forEach(row => {
                if (row.category === '通勤費') {
                    if (row.payee.includes('⇔')) {
                        hasRoundCommute = true;
                    } else if (row.payee.includes('出社')) {
                        hasMorningCommute = true;
                    } else if (row.payee.includes('帰宅')) {
                        hasEveningCommute = true;
                    }
                } else if (row.category.startsWith('交通費') && !row.category.includes('駐車場') && !row.category.includes('ガソリン')) {
                    const route = parseRoute(row.payee);
                    if (route) {
                        trainTravels.push(route);
                    }
                }
            });

            // 1. Morning Commute completeness
            let morningOk = hasRoundCommute || hasMorningCommute;
            if (!morningOk) {
                // Check if replaced by a train travel starting from homeStation
                // e.g. (HOME) -> (X) or (HOME) ⇔ (X)
                const replaced = trainTravels.some(r => r.start.includes(homeStation));
                if (replaced) {
                    morningOk = true;
                }
            }

            // 2. Evening Commute completeness
            let eveningOk = hasRoundCommute || hasEveningCommute;
            if (!eveningOk) {
                // Check if replaced by a train travel ending at homeStation
                // e.g. (X) -> (HOME) or (X) ⇔ (HOME)
                const replaced = trainTravels.some(r => r.end.includes(homeStation));
                if (replaced) {
                    eveningOk = true;
                }
            }

            // If morning or evening is missing, we flag it.
            // But wait, if they only registered something completely unrelated (like 内部飲食代 or gasoline only), they might have worked.
            // Let's issue a warning.
            if (!morningOk || !eveningOk) {
                // Generate a combined warning for this date.
                // We will attach the issue to the first row of this date.
                const targetRow = dayRows[0];
                let missingParts = [];
                if (!morningOk) missingParts.push('朝（出社）');
                if (!eveningOk) missingParts.push('帰り（帰宅）');

                addRowIssue(
                    targetRow, 
                    'warning', 
                    '通勤費・交通費漏れ疑い', 
                    `${dateStr} の通勤申請に不足があります: ${missingParts.join(' および ')} の通勤ルート（または自宅発着の交通費）が申請されていません。`, 
                    '経費科目'
                );
            }
        });
    } else {
        console.warn('Could not determine home station. Commuter completeness check skipped.');
    }
}

// Add a General Issue (not tied to a specific row, or representing overall status)
function addGeneralIssue(type, title, desc) {
    const issue = {
        type: type,
        rowIdx: -1,
        excelNo: '-',
        date: '-',
        category: '全体',
        title: title,
        desc: desc,
        cellName: '-'
    };
    currentIssues.push(issue);
}

// Render issues summary and list
function renderIssues() {
    const issuesContainer = document.getElementById('issues-list');
    issuesContainer.innerHTML = '';

    const errors = currentIssues.filter(i => i.type === 'error');
    const warnings = currentIssues.filter(i => i.type === 'warning');

    // Update summary labels
    const errCountBadge = document.getElementById('error-count');
    const warnCountBadge = document.getElementById('warning-count');
    
    errCountBadge.textContent = errors.length;
    warnCountBadge.textContent = warnings.length;

    // Summary Card styling based on alerts
    const statusVal = document.getElementById('status-value');
    const statusCard = document.getElementById('status-card');
    
    if (errors.length > 0) {
        statusVal.textContent = '不備あり (要修正)';
        statusVal.className = 'summary-val status-danger';
    } else if (warnings.length > 0) {
        statusVal.textContent = '確認事項あり';
        statusVal.className = 'summary-val status-warning';
    } else {
        statusVal.textContent = '不備なし (良好)';
        statusVal.className = 'summary-val status-ok';
    }

    if (currentIssues.length === 0) {
        issuesContainer.innerHTML = `
            <div class="empty-state">
                <span class="material-icons" style="font-size: 3rem; color: var(--success); margin-bottom: 1rem;">check_circle</span>
                <p>チェック完了。不備や確認事項は見つかりませんでした！</p>
            </div>
        `;
        return;
    }

    filterIssues('all');
}

// Filter and Display Issues in List
function filterIssues(tab) {
    const container = document.getElementById('issues-list');
    container.innerHTML = '';

    let filtered = currentIssues;
    if (tab === 'errors') {
        filtered = currentIssues.filter(i => i.type === 'error');
    } else if (tab === 'warnings') {
        filtered = currentIssues.filter(i => i.type === 'warning');
    }

    if (filtered.length === 0) {
        container.innerHTML = `
            <div class="empty-state">
                <p>表示する項目がありません。</p>
            </div>
        `;
        return;
    }

    filtered.forEach(issue => {
        const item = document.createElement('div');
        item.className = `issue-item ${issue.type}`;
        
        const icon = issue.type === 'error' ? 'cancel' : 'warning';
        
        let cellText = issue.cellName !== '-' ? `${issue.cellName}セル` : '';
        let rowText = issue.rowIdx !== -1 ? `[No.${issue.excelNo}]` : '';
        let locText = cellText || rowText ? `<div class="issue-loc" onclick="focusRow(${issue.rowIdx})">${rowText} ${cellText}</div>` : '';

        item.innerHTML = `
            <span class="material-icons issue-icon">${icon}</span>
            <div class="issue-content">
                <div class="issue-title">${issue.title}</div>
                <div class="issue-desc">${issue.desc}</div>
                ${locText}
            </div>
        `;
        container.appendChild(item);
    });
}

// Render spreadsheet-like preview table
function renderTable(colIdx) {
    const tableHeader = document.getElementById('table-header');
    const tableBody = document.getElementById('table-body');
    
    tableHeader.innerHTML = '';
    tableBody.innerHTML = '';

    // Create table header
    const headers = ['ステータス', '番号', '日付', '支払先内容', '経費科目', 'プロジェクトID', '金額', '備考・メモ'];
    const headerRow = document.createElement('tr');
    headers.forEach(h => {
        const th = document.createElement('th');
        th.textContent = h;
        if (h === 'ステータス') th.style.textAlign = 'center';
        headerRow.appendChild(th);
    });
    tableHeader.appendChild(headerRow);

    // Create rows
    allRowsData.forEach(row => {
        const tr = document.createElement('tr');
        tr.id = `row-${row.excelRowIdx}`;

        // Row severity
        const hasError = row.issues.some(i => i.type === 'error');
        const hasWarning = row.issues.some(i => i.type === 'warning');

        if (hasError) {
            tr.classList.add('row-has-error');
        } else if (hasWarning) {
            tr.classList.add('row-has-warning');
        }

        // Status Badge cell
        const tdStatus = document.createElement('td');
        tdStatus.className = 'row-status-cell';
        let badgeClass = 'ok';
        let badgeIcon = 'check';
        
        if (hasError) {
            badgeClass = 'error';
            badgeIcon = 'close';
        } else if (hasWarning) {
            badgeClass = 'warning';
            badgeIcon = '!';
        }

        tdStatus.innerHTML = `<span class="row-status-badge ${badgeClass}">${badgeIcon}</span>`;
        tr.appendChild(tdStatus);

        // Standard cells
        const fields = [
            row.no,
            row.rawDate,
            row.payee,
            row.category,
            row.projId,
            typeof row.amount === 'number' ? `¥${row.amount.toLocaleString()}` : row.amount,
            row.remarks
        ];

        fields.forEach(val => {
            const td = document.createElement('td');
            td.textContent = val;
            tr.appendChild(td);
        });

        // Set row click behavior to show hover or alert if has issues
        if (row.issues.length > 0) {
            tr.style.cursor = 'pointer';
            tr.addEventListener('click', () => {
                let alertMsg = row.issues.map(i => `・[${i.title}] ${i.desc}`).join('\n');
                alert(`行 No.${row.no} (${row.rawDate}) の確認事項:\n\n${alertMsg}`);
            });
        }

        tableBody.appendChild(tr);
    });
}

// Highlight and scroll to a specific row
window.focusRow = function(rowIdx) {
    if (rowIdx === -1) return;
    const element = document.getElementById(`row-${rowIdx}`);
    if (element) {
        element.scrollIntoView({ behavior: 'smooth', block: 'center' });
        // Visual flash effect
        element.style.outline = '2px solid var(--primary)';
        setTimeout(() => {
            element.style.outline = 'none';
        }, 2000);
    }
};

// Map column name to Excel letter for reference
function getColLetter(colName) {
    const mapping = {
        '番号': 'A',
        '日付': 'B',
        '支払先内容': 'C',
        '経費科目': 'D',
        'プロジェクトID': 'E',
        '金額': 'F',
        '備考・メモ': 'G'
    };
    return mapping[colName] || '';
}
