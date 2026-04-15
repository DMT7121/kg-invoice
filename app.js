/**
 * KG Invoice Converter - CukCuk to Sapo Format
 * 
 * This application reads CukCuk invoice files (Excel/CSV) and converts them
 * to the Sapo e-invoice format for import.
 * 
 * Sapo format: 35 columns (A-AI), 3 header rows + data rows
 */

// ============================================
// State Management
// ============================================
var AppState = {
    files: [],           // Uploaded file objects
    invoices: [],        // Parsed CukCuk invoices
    sapoData: [],        // Converted Sapo data rows
    currentTab: 'cukcuk-preview'
};

// ============================================
// Sapo Column Definitions (35 columns, A-AI)
// Mapped exactly from Sapo template analysis
// ============================================
var SAPO_HEADERS = {
    // Row 1 (header group names)
    row1: [
        'Ký hiệu*',                     // A (col 0)
        'Mã chứng từ gốc',              // B (col 1)
        'Thông tin người mua',           // C (col 2) - spans to I
        '', '', '', '', '', '',          // D-I (cols 3-8) merged with C
        'Thông tin người nhận',          // J (col 9) - spans to K
        '',                              // K (col 10) merged with J
        'Thông tin giao dịch',           // L (col 11) - spans to P
        '', '', '', '',                  // M-P (cols 12-15) merged with L
        'Chiết khấu cả hóa đơn',        // Q (col 16) - spans to R
        '',                              // R (col 17) merged with Q
        'Thuế GTGT cả hóa đơn (%)',     // S (col 18) standalone
        'Thông tin hàng hóa, dịch vụ',  // T (col 19) - spans to AI
        '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '' // U-AI (cols 20-34)
    ],
    // Row 2 (sub-column names)
    row2: [
        '',                              // A (col 0) - merged from row 1
        '',                              // B (col 1) - merged from row 1
        'Mã khách',                      // C (col 2)
        'Mã số thuế người mua',          // D (col 3)
        'Tên đơn vị',                    // E (col 4)
        'Địa chỉ',                       // F (col 5)
        'Người mua hàng',                // G (col 6)
        'Số điện thoại',                 // H (col 7)
        'Email',                         // I (col 8)
        'Tên người nhận hóa đơn',        // J (col 9)
        'Email nhận hóa đơn',            // K (col 10)
        'Hình thức thanh toán*',         // L (col 11)
        'Tên ngân hàng',                 // M (col 12)
        'Số tài khoản ngân hàng',        // N (col 13)
        'Loại tiền',                     // O (col 14)
        'Tỷ giá',                        // P (col 15)
        '',                              // Q (col 16) - merged from row 1
        '',                              // R (col 17) - merged from row 1
        '',                              // S (col 18) - merged from row 1
        'Tính chất',                     // T (col 19)
        'Mã hàng',                       // U (col 20)
        'Tên hàng hóa/dịch vụ',         // V (col 21)
        'ĐVT',                          // W (col 22)
        'Số lượng',                      // X (col 23)
        'Đơn giá',                       // Y (col 24)
        'Thành tiền',                    // Z (col 25)
        'Chiết khấu/SP',                // AA (col 26) - spans to AB
        '',                              // AB (col 27) merged with AA
        'Tiền chiết khấu',              // AC (col 28)
        '%Tính thuế',                    // AD (col 29)
        'Thành tiền đã trừ chiết khấu', // AE (col 30)
        'Tiền thuế GTGT được giảm',     // AF (col 31)
        'Thuế GTGT (%)',                 // AG (col 32)
        'Tiền thuế GTGT',               // AH (col 33)
        'Tổng tiền'                      // AI (col 34)
    ],
    // Row 3 (sub-sub headers for discount columns)
    row3: [
        '', '', '', '', '', '', '', '', '', // A-I (cols 0-8)
        '', '',                             // J-K (cols 9-10)
        '', '', '', '', '',                 // L-P (cols 11-15)
        'Giá trị', '%',                    // Q-R (cols 16-17)
        '',                                 // S (col 18)
        '', '', '', '', '', '', '',         // T-Z (cols 19-25)
        'Giá trị', '%',                    // AA-AB (cols 26-27)
        '', '', '', '', '', '', ''          // AC-AI (cols 28-34)
    ]
};

// Column indices for Sapo format (35 columns, 0-indexed)
var SAPO_COL = {
    KY_HIEU: 0,           // A - Ký hiệu*
    MA_CHUNG_TU: 1,        // B - Mã chứng từ gốc
    MA_KHACH: 2,           // C - Mã khách
    MA_SO_THUE: 3,         // D - Mã số thuế người mua
    TEN_DON_VI: 4,         // E - Tên đơn vị
    DIA_CHI: 5,            // F - Địa chỉ
    NGUOI_MUA: 6,          // G - Người mua hàng
    SDT: 7,                // H - Số điện thoại
    EMAIL: 8,              // I - Email
    NGUOI_NHAN: 9,         // J - Tên người nhận hóa đơn
    EMAIL_NHAN: 10,        // K - Email nhận hóa đơn
    HINH_THUC_TT: 11,     // L - Hình thức thanh toán*
    NGAN_HANG: 12,         // M - Tên ngân hàng
    SO_TK: 13,             // N - Số tài khoản ngân hàng
    LOAI_TIEN: 14,         // O - Loại tiền
    TY_GIA: 15,            // P - Tỷ giá
    CK_GIA_TRI: 16,        // Q - Chiết khấu cả HĐ - Giá trị
    CK_PHAN_TRAM: 17,      // R - Chiết khấu cả HĐ - %
    THUE_GTGT_HD: 18,      // S - Thuế GTGT cả hóa đơn (%)
    TINH_CHAT: 19,         // T - Tính chất
    MA_HANG: 20,           // U - Mã hàng
    TEN_HANG: 21,          // V - Tên hàng hóa/dịch vụ
    DVT: 22,               // W - ĐVT
    SO_LUONG: 23,          // X - Số lượng
    DON_GIA: 24,           // Y - Đơn giá
    THANH_TIEN: 25,        // Z - Thành tiền
    CK_SP_GIA_TRI: 26,    // AA - Chiết khấu/SP - Giá trị
    CK_SP_PHAN_TRAM: 27,  // AB - Chiết khấu/SP - %
    TIEN_CHIET_KHAU: 28,  // AC - Tiền chiết khấu
    PHAN_TRAM_THUE: 29,   // AD - %Tính thuế
    THANH_TIEN_DA_TRU: 30, // AE - Thành tiền đã trừ chiết khấu
    THUE_GTGT_GIAM: 31,   // AF - Tiền thuế GTGT được giảm
    THUE_GTGT_PCT: 32,    // AG - Thuế GTGT (%)
    TIEN_THUE_GTGT: 33,   // AH - Tiền thuế GTGT
    TONG_TIEN: 34         // AI - Tổng tiền
};

// Total columns
var TOTAL_COLS = 35;

// ============================================
// Initialization
// ============================================
document.addEventListener('DOMContentLoaded', function () {
    initEventListeners();
});

function initEventListeners() {
    var dropZone = document.getElementById('dropZone');
    var fileInput = document.getElementById('fileInput');
    var fileInputMore = document.getElementById('fileInputMore');
    var btnAddMore = document.getElementById('btnAddMore');
    var btnHelp = document.getElementById('btnHelp');
    var btnCloseHelp = document.getElementById('btnCloseHelp');
    var btnExportXlsx = document.getElementById('btnExportXlsx');
    var btnExportCsv = document.getElementById('btnExportCsv');
    var btnReset = document.getElementById('btnReset');

    // Drag & Drop
    dropZone.addEventListener('dragover', function (e) {
        e.preventDefault();
        dropZone.classList.add('drag-over');
    });

    dropZone.addEventListener('dragleave', function (e) {
        e.preventDefault();
        dropZone.classList.remove('drag-over');
    });

    dropZone.addEventListener('drop', function (e) {
        e.preventDefault();
        dropZone.classList.remove('drag-over');
        if (e.dataTransfer.files.length > 0) {
            handleFiles(e.dataTransfer.files);
        }
    });

    // Click to upload
    dropZone.addEventListener('click', function (e) {
        if (e.target.tagName !== 'LABEL' && e.target.tagName !== 'INPUT') {
            fileInput.click();
        }
    });

    fileInput.addEventListener('change', function (e) {
        if (e.target.files.length > 0) {
            handleFiles(e.target.files);
        }
    });

    // Add more files
    btnAddMore.addEventListener('click', function () {
        fileInputMore.click();
    });

    fileInputMore.addEventListener('change', function (e) {
        if (e.target.files.length > 0) {
            handleFiles(e.target.files);
        }
    });

    // Tab switching
    document.querySelectorAll('.tab').forEach(function (tab) {
        tab.addEventListener('click', function () {
            switchTab(tab.getAttribute('data-tab'));
        });
    });

    // Help modal
    btnHelp.addEventListener('click', function () {
        document.getElementById('helpModal').style.display = 'flex';
    });
    btnCloseHelp.addEventListener('click', function () {
        document.getElementById('helpModal').style.display = 'none';
    });
    document.querySelector('.modal-overlay').addEventListener('click', function () {
        document.getElementById('helpModal').style.display = 'none';
    });

    // Export
    btnExportXlsx.addEventListener('click', function () {
        exportToXlsx();
    });
    btnExportCsv.addEventListener('click', function () {
        exportToCsv();
    });

    // Reset
    btnReset.addEventListener('click', resetApp);
}

// ============================================
// File Handling
// ============================================
function handleFiles(fileList) {
    var validFiles = [];
    for (var i = 0; i < fileList.length; i++) {
        var file = fileList[i];
        var ext = file.name.split('.').pop().toLowerCase();
        if (['xlsx', 'xls', 'csv'].indexOf(ext) !== -1) {
            validFiles.push(file);
        }
    }

    if (validFiles.length === 0) {
        showToast('Vui lòng chọn file .xlsx, .xls hoặc .csv', 'error');
        return;
    }

    var processedCount = 0;
    validFiles.forEach(function (file) {
        var reader = new FileReader();
        reader.onload = function (e) {
            try {
                processFile(file.name, e.target.result);
                processedCount++;
                if (processedCount === validFiles.length) {
                    onAllFilesProcessed();
                }
            } catch (err) {
                console.error('Error processing file:', file.name, err);
                showToast('Lỗi đọc file: ' + file.name + ' - ' + err.message, 'error');
                processedCount++;
                if (processedCount === validFiles.length) {
                    onAllFilesProcessed();
                }
            }
        };
        reader.readAsArrayBuffer(file);
    });
}

function processFile(fileName, arrayBuffer) {
    var workbook = XLSX.read(arrayBuffer, { type: 'array', raw: true });

    workbook.SheetNames.forEach(function (sheetName) {
        var sheet = workbook.Sheets[sheetName];
        var rawData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: true });

        if (rawData.length < 3) return;

        // Parse invoices from this sheet
        var invoices = parseCukCukSheet(rawData, fileName, sheetName);
        if (invoices.length > 0) {
            AppState.files.push({ name: fileName, sheet: sheetName, invoiceCount: invoices.length });
            AppState.invoices = AppState.invoices.concat(invoices);
        }
    });
}

function onAllFilesProcessed() {
    if (AppState.invoices.length === 0) {
        showToast('Không tìm thấy hóa đơn CukCuk trong file', 'error');
        return;
    }

    // Convert to Sapo format
    AppState.sapoData = convertToSapo(AppState.invoices);

    // Update UI
    updateFileList();
    renderCukCukPreview();
    renderSapoPreview();
    updateExportDetails();

    // Show sections
    document.getElementById('fileList-section').style.display = 'block';
    document.getElementById('step2-section').style.display = 'block';
    document.getElementById('step3-section').style.display = 'block';

    // Update steps
    setStep(3);

    showToast('Đã chuyển đổi thành công ' + AppState.invoices.length + ' hóa đơn!', 'success');
}

// ============================================
// CukCuk Invoice Parser
// ============================================
function parseCukCukSheet(data, fileName, sheetName) {
    var invoices = [];

    var i = 0;
    while (i < data.length) {
        var row = data[i];
        var firstCell = String(row[0] || '').trim();

        // Check if this is the start of an invoice block
        if (firstCell === 'Hóa đơn' || firstCell === 'Hoá đơn') {
            var invoice = parseInvoiceBlock(data, i);
            if (invoice) {
                invoice.fileName = fileName;
                invoice.sheetName = sheetName;
                invoices.push(invoice);
                i = invoice._endRow + 1;
                continue;
            }
        }

        // Also check for invoice blocks that start with "Số:" directly
        if (row.length > 0) {
            var cellStr = String(row[0] || '');
            if (cellStr.indexOf('Số:') !== -1 || cellStr.indexOf('Số :') !== -1) {
                var invoice = parseInvoiceBlock(data, Math.max(0, i - 1));
                if (invoice) {
                    invoice.fileName = fileName;
                    invoice.sheetName = sheetName;
                    invoices.push(invoice);
                    i = invoice._endRow + 1;
                    continue;
                }
            }
        }

        i++;
    }

    // If no invoices found, try parsing the whole sheet as one
    if (invoices.length === 0 && data.length > 5) {
        var invoice = parseInvoiceBlock(data, 0);
        if (invoice) {
            invoice.fileName = fileName;
            invoice.sheetName = sheetName;
            invoices.push(invoice);
        }
    }

    return invoices;
}

function parseInvoiceBlock(data, startRow) {
    var invoice = {
        number: '',
        date: '',
        timeRange: '',
        cashier: '',
        server: '',
        table: '',
        guestCount: '',
        customerName: '',
        customerPhone: '',
        customerCard: '',
        delivery: '',
        address: '',
        items: [],
        totalAmount: 0,
        paymentMethods: [],
        _endRow: startRow
    };

    var i = startRow;
    var foundItems = false;

    while (i < data.length) {
        var row = data[i];
        var rowStr = row.map(function (c) { return String(c || '').trim(); }).join(' ');

        // Parse header info before items table
        if (!foundItems) {
            parseInfoRow(row, invoice);
        }

        // Check for item table header row
        if (isItemHeaderRow(row)) {
            foundItems = true;
            i++;
            continue;
        }

        // Parse item rows
        if (foundItems) {
            var firstCell = String(row[0] || '').trim();
            var isNumber = /^\d+$/.test(firstCell);

            if (isNumber && row.length >= 5) {
                var item = parseItemRow(row);
                if (item) {
                    invoice.items.push(item);
                }
            }

            // Parse total
            if (rowStr.indexOf('Tổng tiền thanh toán') !== -1 || 
                (rowStr.indexOf('Tổng tiền') !== -1 && rowStr.indexOf('thanh toán') !== -1)) {
                invoice.totalAmount = findNumberInRow(row);
            }

            // Parse payment methods
            parsePaymentRow(row, rowStr, invoice);

            // Check for end of invoice
            if (isEndOfInvoice(data, i, startRow)) {
                invoice._endRow = i;
                break;
            }
        }

        invoice._endRow = i;
        i++;
    }

    if (invoice.items.length === 0) return null;

    // Default payment method
    if (invoice.paymentMethods.length === 0 && invoice.totalAmount > 0) {
        invoice.paymentMethods.push({ method: 'Tiền mặt', amount: invoice.totalAmount });
    }

    return invoice;
}

function parseInfoRow(row, invoice) {
    // Process each cell individually for more reliable parsing
    for (var i = 0; i < row.length; i++) {
        var cell = String(row[i] || '').trim();
        if (!cell) continue;

        // Parse invoice number: "Số: 2602000218" or "Số : 2602000218"
        // Must NOT match "Số người:"
        if (cell.indexOf('Số:') !== -1 || cell.indexOf('Số :') !== -1) {
            if (cell.indexOf('Số người') === -1) {
                var colonIdx = cell.indexOf(':');
                if (colonIdx !== -1) {
                    var afterColon = cell.substring(colonIdx + 1).trim();
                    if (afterColon && /[A-Za-z0-9]/.test(afterColon)) {
                        invoice.number = afterColon;
                    }
                }
            }
        }

        // Parse date: "Ngày: 01/04/2026 (18:23 - 19:21)"
        if (cell.indexOf('Ngày') !== -1 || cell.indexOf('Ngày') !== -1) {
            var dateMatch = cell.match(/(\d{1,2}\/\d{1,2}\/\d{2,4})/);
            if (dateMatch) invoice.date = dateMatch[1];

            var timeMatch = cell.match(/\((\d{1,2}:\d{2}\s*-\s*\d{1,2}:\d{2})\)/);
            if (timeMatch) invoice.timeRange = timeMatch[1];
        }

        // Parse cashier: "Thu ngân: THU NGÂN"
        if (cell.indexOf('Thu ngân') !== -1 || cell.indexOf('Thu Ngân') !== -1) {
            var colonIdx = cell.indexOf(':');
            if (colonIdx !== -1) {
                var val = cell.substring(colonIdx + 1).trim();
                if (val) invoice.cashier = val;
            }
        }

        // Parse server: "Phục vụ: Nguyễn Văn Hoà"
        if (cell.indexOf('Phục vụ') !== -1 || cell.indexOf('phục vụ') !== -1) {
            var colonIdx = cell.indexOf(':');
            if (colonIdx !== -1) {
                var val = cell.substring(colonIdx + 1).trim();
                if (val) invoice.server = val;
            }
        }

        // Parse table: "Bàn: A.18"
        if (cell.indexOf('Bàn') !== -1 && cell.indexOf(':') !== -1) {
            var colonIdx = cell.indexOf(':');
            if (colonIdx !== -1) {
                var val = cell.substring(colonIdx + 1).trim();
                if (val) invoice.table = val;
            }
        }

        // Parse guest count: "Số người: 3"
        if (cell.indexOf('Số người') !== -1) {
            var guestMatch = cell.match(/(\d+)/);
            if (guestMatch) invoice.guestCount = guestMatch[1];
        }

        // Parse customer: "KH: tên khách"
        if (cell.indexOf('KH:') !== -1 || cell.indexOf('KH :') !== -1) {
            var colonIdx = cell.indexOf(':');
            if (colonIdx !== -1) {
                var val = cell.substring(colonIdx + 1).trim();
                if (val) invoice.customerName = val;
            }
        }

        // Parse phone: "ĐT: 0123456789"
        if (cell.indexOf('ĐT') !== -1 || cell.indexOf('ĐT') !== -1) {
            var colonIdx = cell.indexOf(':');
            if (colonIdx !== -1) {
                var val = cell.substring(colonIdx + 1).trim();
                if (val && /\d/.test(val)) invoice.customerPhone = val;
            }
        }
    }
}

function isItemHeaderRow(row) {
    var cells = row.map(function (c) { return String(c || '').trim().toLowerCase(); });
    var hasSTT = cells.indexOf('stt') !== -1;
    var hasTenMon = cells.some(function (c) {
        return c.indexOf('tên món') !== -1 || c.indexOf('ten mon') !== -1 || c === 'tên hàng';
    });
    var hasDVT = cells.some(function (c) { return c === 'đvt' || c === 'dvt'; });
    return hasSTT && (hasTenMon || hasDVT);
}

function parseItemRow(row) {
    try {
        var item = {
            stt: parseInt(String(row[0] || '0')),
            name: String(row[1] || '').trim(),
            unit: String(row[2] || '').trim(),
            quantity: parseCukCukNumber(row[3]),
            unitPrice: parseCukCukNumber(row[4]),
            amount: parseCukCukNumber(row[5]),
            discount: parseCukCukNumber(row[6]),
            vatPercent: parseCukCukNumber(row[7]),
            total: parseCukCukNumber(row[8])
        };

        if (!item.name) return null;

        // Calculate amount if missing
        if (item.amount === 0 && item.quantity > 0 && item.unitPrice > 0) {
            item.amount = item.quantity * item.unitPrice;
        }

        // Calculate total if missing
        if (item.total === 0 && item.amount > 0) {
            var afterDiscount = item.amount - item.discount;
            item.total = afterDiscount;
            if (item.vatPercent > 0) {
                item.total += Math.round(afterDiscount * item.vatPercent / 100);
            }
        }

        return item;
    } catch (e) {
        console.warn('Failed to parse item row:', row, e);
        return null;
    }
}

function parseCukCukNumber(val) {
    if (val === null || val === undefined || val === '') return 0;
    
    // If already a number, return it
    if (typeof val === 'number') return val;
    
    var str = String(val).trim().replace(/"/g, '');

    // Vietnamese number format handling
    if (str.indexOf(',') !== -1 && str.indexOf('.') === -1) {
        // Comma as decimal: "3,00" -> "3.00"
        str = str.replace(',', '.');
    } else if (str.indexOf('.') !== -1 && str.indexOf(',') !== -1) {
        // Both: "3.000,00" -> "3000.00"
        str = str.replace(/\./g, '').replace(',', '.');
    } else if (str.indexOf('.') !== -1) {
        // Dot only: check if thousands separator
        var parts = str.split('.');
        if (parts.length >= 2 && parts[parts.length - 1].length === 3) {
            str = str.replace(/\./g, '');
        }
    }

    var num = parseFloat(str);
    return isNaN(num) ? 0 : num;
}

function findNumberInRow(row) {
    for (var i = row.length - 1; i >= 0; i--) {
        var val = parseCukCukNumber(row[i]);
        if (val > 0) return val;
    }
    return 0;
}

function parsePaymentRow(row, rowStr, invoice) {
    var methods = [
        { name: 'Tiền mặt', label: 'Tiền mặt' },
        { name: 'Chuyển khoản', label: 'Chuyển khoản' },
        { name: 'Thẻ', label: 'Thẻ' }
    ];

    methods.forEach(function (m) {
        if (rowStr.indexOf(m.name + ':') !== -1 || rowStr.indexOf(m.name + ' :') !== -1) {
            var amount = extractPaymentAmount(row, m.name);
            if (amount > 0) {
                // Avoid duplicates
                var exists = invoice.paymentMethods.some(function (p) { return p.method === m.label; });
                if (!exists) {
                    invoice.paymentMethods.push({ method: m.label, amount: amount });
                }
            }
        }
    });
}

function extractPaymentAmount(row, methodName) {
    for (var i = 0; i < row.length; i++) {
        var cell = String(row[i] || '').trim();
        if (cell.indexOf(methodName + ':') !== -1 || cell.indexOf(methodName + ' :') !== -1) {
            // Check this cell for embedded number after ':'
            var colonIdx = cell.indexOf(':');
            if (colonIdx !== -1) {
                var afterColon = cell.substring(colonIdx + 1).trim();
                var val = parseCukCukNumber(afterColon);
                if (val > 0) return val;
            }
            // Check next cells
            for (var j = i + 1; j < Math.min(i + 3, row.length); j++) {
                var v = parseCukCukNumber(row[j]);
                if (v > 0) return v;
            }
        }
    }
    return 0;
}

function isEndOfInvoice(data, currentRow, startRow) {
    if (currentRow >= data.length - 1) return true;

    if (currentRow - startRow > 8) {
        var currentEmpty = isRowEmpty(data[currentRow]);
        var nextEmpty = currentRow + 1 < data.length ? isRowEmpty(data[currentRow + 1]) : true;
        if (currentEmpty && nextEmpty) return true;
    }

    if (currentRow + 1 < data.length) {
        var nextFirst = String(data[currentRow + 1][0] || '').trim();
        if (nextFirst === 'Hóa đơn' || nextFirst === 'Hoá đơn') return true;
    }

    return false;
}

function isRowEmpty(row) {
    if (!row || row.length === 0) return true;
    return row.every(function (c) { return String(c || '').trim() === ''; });
}

// ============================================
// CukCuk to Sapo Converter
// ============================================
function convertToSapo(invoices) {
    var sapoRows = [];

    invoices.forEach(function (invoice) {
        var paymentMethod = determinePaymentMethod(invoice);

        invoice.items.forEach(function (item, itemIndex) {
            var row = new Array(TOTAL_COLS);
            for (var i = 0; i < TOTAL_COLS; i++) row[i] = '';

            // First item row gets invoice header info
            if (itemIndex === 0) {
                row[SAPO_COL.MA_CHUNG_TU] = invoice.number;
                row[SAPO_COL.NGUOI_MUA] = invoice.customerName || '';
                row[SAPO_COL.SDT] = invoice.customerPhone || '';
                row[SAPO_COL.HINH_THUC_TT] = paymentMethod;
                row[SAPO_COL.LOAI_TIEN] = 'VND';
                row[SAPO_COL.TY_GIA] = 1;
            }

            // Product info (always filled)
            row[SAPO_COL.TINH_CHAT] = 'Hàng hóa dịch vụ';
            row[SAPO_COL.TEN_HANG] = item.name;
            row[SAPO_COL.DVT] = item.unit;
            row[SAPO_COL.SO_LUONG] = item.quantity;
            row[SAPO_COL.DON_GIA] = item.unitPrice;
            row[SAPO_COL.THANH_TIEN] = item.amount;

            // Discount per item
            if (item.discount > 0) {
                row[SAPO_COL.CK_SP_GIA_TRI] = item.discount;
                row[SAPO_COL.TIEN_CHIET_KHAU] = item.discount;
            }

            // Amount after discount
            var amountAfterDiscount = item.amount - item.discount;
            row[SAPO_COL.THANH_TIEN_DA_TRU] = amountAfterDiscount;

            // VAT
            if (item.vatPercent > 0) {
                row[SAPO_COL.THUE_GTGT_PCT] = item.vatPercent;
                var vatAmount = Math.round(amountAfterDiscount * item.vatPercent / 100);
                row[SAPO_COL.TIEN_THUE_GTGT] = vatAmount;
                row[SAPO_COL.TONG_TIEN] = amountAfterDiscount + vatAmount;
            } else {
                row[SAPO_COL.TONG_TIEN] = amountAfterDiscount;
            }

            sapoRows.push(row);
        });
    });

    return sapoRows;
}

function determinePaymentMethod(invoice) {
    if (invoice.paymentMethods.length === 0) return 'Tiền mặt';
    if (invoice.paymentMethods.length === 1) {
        var method = invoice.paymentMethods[0].method;
        if (method === 'Thẻ') return 'Thẻ';
        if (method === 'Chuyển khoản') return 'Chuyển khoản';
        return 'Tiền mặt';
    }
    return 'TM/CK';
}

// ============================================
// Export Functions
// ============================================
function exportToXlsx() {
    try {
        var wb = XLSX.utils.book_new();
        var wsData = buildSapoSheetData();
        var ws = XLSX.utils.aoa_to_sheet(wsData);

        // Set column widths
        ws['!cols'] = [
            { wch: 12 },  // A - Ký hiệu
            { wch: 16 },  // B - Mã chứng từ
            { wch: 10 },  // C - Mã khách
            { wch: 16 },  // D - MST
            { wch: 25 },  // E - Tên đơn vị
            { wch: 30 },  // F - Địa chỉ
            { wch: 20 },  // G - Người mua
            { wch: 14 },  // H - SĐT
            { wch: 20 },  // I - Email
            { wch: 20 },  // J - Người nhận
            { wch: 20 },  // K - Email nhận
            { wch: 15 },  // L - HTTT
            { wch: 15 },  // M - Ngân hàng
            { wch: 16 },  // N - STK
            { wch: 8 },   // O - Loại tiền
            { wch: 8 },   // P - Tỷ giá
            { wch: 10 },  // Q - CK giá trị
            { wch: 6 },   // R - CK %
            { wch: 8 },   // S - VAT HĐ
            { wch: 16 },  // T - Tính chất
            { wch: 12 },  // U - Mã hàng
            { wch: 30 },  // V - Tên HH
            { wch: 8 },   // W - ĐVT
            { wch: 8 },   // X - SL
            { wch: 12 },  // Y - Đơn giá
            { wch: 12 },  // Z - Thành tiền
            { wch: 10 },  // AA - CK/SP GiaTri
            { wch: 6 },   // AB - CK/SP %
            { wch: 12 },  // AC - Tiền CK
            { wch: 8 },   // AD - %Thuế
            { wch: 14 },  // AE - TT đã trừ CK
            { wch: 12 },  // AF - Thuế giảm
            { wch: 8 },   // AG - VAT%
            { wch: 12 },  // AH - Tiền VAT
            { wch: 14 }   // AI - Tổng tiền
        ];

        // Set merge cells - match exactly with Sapo template
        ws['!merges'] = [
            // Row 1 group headers
            { s: { r: 0, c: 0 }, e: { r: 2, c: 0 } },   // A1:A3 - Ký hiệu
            { s: { r: 0, c: 1 }, e: { r: 2, c: 1 } },   // B1:B3 - Mã chứng từ
            { s: { r: 0, c: 2 }, e: { r: 0, c: 8 } },   // C1:I1 - Thông tin người mua
            { s: { r: 0, c: 9 }, e: { r: 0, c: 10 } },  // J1:K1 - Thông tin người nhận
            { s: { r: 0, c: 11 }, e: { r: 0, c: 15 } }, // L1:P1 - Thông tin giao dịch
            { s: { r: 0, c: 16 }, e: { r: 1, c: 17 } }, // Q1:R2 - Chiết khấu cả HĐ
            { s: { r: 0, c: 18 }, e: { r: 2, c: 18 } }, // S1:S3 - Thuế GTGT cả HĐ
            { s: { r: 0, c: 19 }, e: { r: 0, c: 34 } }, // T1:AI1 - Thông tin HH, DV
            // Row 2 sub-headers that merge with row 3
            { s: { r: 1, c: 2 }, e: { r: 2, c: 2 } },   // C2:C3 - Mã khách
            { s: { r: 1, c: 3 }, e: { r: 2, c: 3 } },   // D2:D3
            { s: { r: 1, c: 4 }, e: { r: 2, c: 4 } },   // E2:E3
            { s: { r: 1, c: 5 }, e: { r: 2, c: 5 } },   // F2:F3
            { s: { r: 1, c: 6 }, e: { r: 2, c: 6 } },   // G2:G3
            { s: { r: 1, c: 7 }, e: { r: 2, c: 7 } },   // H2:H3
            { s: { r: 1, c: 8 }, e: { r: 2, c: 8 } },   // I2:I3
            { s: { r: 1, c: 9 }, e: { r: 2, c: 9 } },   // J2:J3
            { s: { r: 1, c: 10 }, e: { r: 2, c: 10 } }, // K2:K3
            { s: { r: 1, c: 11 }, e: { r: 2, c: 11 } }, // L2:L3
            { s: { r: 1, c: 12 }, e: { r: 2, c: 12 } }, // M2:M3
            { s: { r: 1, c: 13 }, e: { r: 2, c: 13 } }, // N2:N3
            { s: { r: 1, c: 14 }, e: { r: 2, c: 14 } }, // O2:O3
            { s: { r: 1, c: 15 }, e: { r: 2, c: 15 } }, // P2:P3
            { s: { r: 1, c: 19 }, e: { r: 2, c: 19 } }, // T2:T3
            { s: { r: 1, c: 20 }, e: { r: 2, c: 20 } }, // U2:U3
            { s: { r: 1, c: 21 }, e: { r: 2, c: 21 } }, // V2:V3
            { s: { r: 1, c: 22 }, e: { r: 2, c: 22 } }, // W2:W3
            { s: { r: 1, c: 23 }, e: { r: 2, c: 23 } }, // X2:X3
            { s: { r: 1, c: 24 }, e: { r: 2, c: 24 } }, // Y2:Y3
            { s: { r: 1, c: 25 }, e: { r: 2, c: 25 } }, // Z2:Z3
            { s: { r: 1, c: 26 }, e: { r: 1, c: 27 } }, // AA2:AB2 - Chiết khấu/SP
            { s: { r: 1, c: 28 }, e: { r: 2, c: 28 } }, // AC2:AC3
            { s: { r: 1, c: 29 }, e: { r: 2, c: 29 } }, // AD2:AD3
            { s: { r: 1, c: 30 }, e: { r: 2, c: 30 } }, // AE2:AE3
            { s: { r: 1, c: 31 }, e: { r: 2, c: 31 } }, // AF2:AF3
            { s: { r: 1, c: 32 }, e: { r: 2, c: 32 } }, // AG2:AG3
            { s: { r: 1, c: 33 }, e: { r: 2, c: 33 } }, // AH2:AH3
            { s: { r: 1, c: 34 }, e: { r: 2, c: 34 } }  // AI2:AI3
        ];

        XLSX.utils.book_append_sheet(wb, ws, 'Sapo_Import');

        var now = new Date();
        var dateStr = now.getFullYear().toString() +
            ('0' + (now.getMonth() + 1)).slice(-2) +
            ('0' + now.getDate()).slice(-2) +
            '_' + ('0' + now.getHours()).slice(-2) +
            ('0' + now.getMinutes()).slice(-2);
        var fileName = 'Sapo_Import_' + dateStr + '.xlsx';

        // Use Blob-based download for maximum browser compatibility
        var wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        var blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        var url = URL.createObjectURL(blob);
        var a = document.createElement('a');
        a.href = url;
        a.download = fileName;
        document.body.appendChild(a);
        a.click();
        setTimeout(function () {
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        }, 100);
        showToast('Đã xuất file ' + fileName + ' thành công!', 'success');
    } catch (err) {
        console.error('Export error:', err);
        showToast('Lỗi xuất file: ' + err.message, 'error');
    }
}

function exportToCsv() {
    try {
        var wsData = buildSapoSheetData();
        var csvContent = wsData.map(function (row) {
            return row.map(function (cell) {
                var str = String(cell === null || cell === undefined ? '' : cell);
                if (str.indexOf(',') !== -1 || str.indexOf('"') !== -1 || str.indexOf('\n') !== -1) {
                    str = '"' + str.replace(/"/g, '""') + '"';
                }
                return str;
            }).join(',');
        }).join('\r\n');

        var blob = new Blob(['\ufeff' + csvContent], { type: 'text/csv;charset=utf-8' });
        var url = URL.createObjectURL(blob);
        var a = document.createElement('a');

        var now = new Date();
        var dateStr = now.getFullYear().toString() +
            ('0' + (now.getMonth() + 1)).slice(-2) +
            ('0' + now.getDate()).slice(-2) +
            '_' + ('0' + now.getHours()).slice(-2) +
            ('0' + now.getMinutes()).slice(-2);

        a.href = url;
        a.download = 'Sapo_Import_' + dateStr + '.csv';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);

        showToast('Đã xuất file CSV thành công!', 'success');
    } catch (err) {
        console.error('CSV export error:', err);
        showToast('Lỗi xuất file CSV: ' + err.message, 'error');
    }
}

function buildSapoSheetData() {
    var data = [];

    // 3 header rows
    data.push(SAPO_HEADERS.row1.slice());
    data.push(SAPO_HEADERS.row2.slice());
    data.push(SAPO_HEADERS.row3.slice());

    // Data rows
    AppState.sapoData.forEach(function (row) {
        data.push(row.slice());
    });

    return data;
}

// ============================================
// UI Rendering
// ============================================
function updateFileList() {
    var container = document.getElementById('fileListContainer');
    container.innerHTML = '';

    AppState.files.forEach(function (file, index) {
        var div = document.createElement('div');
        div.className = 'file-item';
        div.innerHTML =
            '<div class="file-item-icon">' +
            '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">' +
            '<path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/>' +
            '<polyline points="14,2 14,8 20,8"/></svg></div>' +
            '<div class="file-item-info">' +
            '<div class="file-item-name">' + escapeHtml(file.name) + '</div>' +
            '<div class="file-item-meta">' +
            '<span>Sheet: ' + escapeHtml(file.sheet) + '</span>' +
            '<span class="file-item-badge">' + file.invoiceCount + ' hóa đơn</span>' +
            '</div></div>' +
            '<div class="file-item-actions">' +
            '<button class="btn-icon btn-danger" onclick="removeFile(' + index + ')" title="Xóa file">' +
            '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">' +
            '<polyline points="3,6 5,6 21,6"/>' +
            '<path d="M19 6v14a2 2 0 01-2 2H7a2 2 0 01-2-2V6m3 0V4a2 2 0 012-2h4a2 2 0 012 2v2"/>' +
            '</svg></button></div>';
        container.appendChild(div);
    });
}

function renderCukCukPreview() {
    var thead = document.getElementById('cukcukTableHead');
    var tbody = document.getElementById('cukcukTableBody');

    thead.innerHTML =
        '<tr>' +
        '<th>Hóa đơn</th><th>Ngày</th><th>HTTT</th>' +
        '<th>STT</th><th>Tên món</th><th>ĐVT</th><th>SL</th>' +
        '<th class="amount-cell">Đơn giá</th><th class="amount-cell">Tiền hàng</th>' +
        '<th class="amount-cell">Tiền KM</th><th>%VAT</th>' +
        '<th class="amount-cell">Thành tiền</th></tr>';

    tbody.innerHTML = '';
    var totalInvoices = AppState.invoices.length;
    var totalItems = 0;
    var grandTotal = 0;

    AppState.invoices.forEach(function (inv) {
        inv.items.forEach(function (item, idx) {
            var tr = document.createElement('tr');
            if (idx === 0) tr.className = 'invoice-header-row';

            var payBadgeClass = getPaymentBadgeClass(inv.paymentMethods);
            var payLabel = determinePaymentMethod(inv);

            tr.innerHTML =
                '<td>' + (idx === 0 ? '<span class="invoice-num">' + escapeHtml(inv.number) + '</span>' : '') + '</td>' +
                '<td>' + (idx === 0 ? escapeHtml(inv.date) : '') + '</td>' +
                '<td>' + (idx === 0 ? '<span class="payment-badge ' + payBadgeClass + '">' + escapeHtml(payLabel) + '</span>' : '') + '</td>' +
                '<td>' + item.stt + '</td>' +
                '<td>' + escapeHtml(item.name) + '</td>' +
                '<td>' + escapeHtml(item.unit) + '</td>' +
                '<td>' + formatQuantity(item.quantity) + '</td>' +
                '<td class="amount-cell">' + formatNumber(item.unitPrice) + '</td>' +
                '<td class="amount-cell">' + formatNumber(item.amount) + '</td>' +
                '<td class="amount-cell">' + (item.discount > 0 ? formatNumber(item.discount) : '-') + '</td>' +
                '<td>' + (item.vatPercent > 0 ? item.vatPercent + '%' : '-') + '</td>' +
                '<td class="amount-cell"><strong>' + formatNumber(item.total) + '</strong></td>';

            tbody.appendChild(tr);
            totalItems++;
            grandTotal += item.total;
        });
    });

    document.getElementById('previewStats').innerHTML =
        '<span class="stat-badge"><strong>' + totalInvoices + '</strong> hóa đơn</span>' +
        '<span class="stat-badge"><strong>' + totalItems + '</strong> dòng</span>' +
        '<span class="stat-badge"><strong>' + formatNumber(grandTotal) + '</strong> VNĐ</span>';
}

function renderSapoPreview() {
    var thead = document.getElementById('sapoTableHead');
    var tbody = document.getElementById('sapoTableBody');

    var headerHtml = '';

    // Row 1 - group headers
    headerHtml += '<tr class="header-row-1">';
    var row1Spans = [
        { text: 'Ký hiệu*', cols: 1 },
        { text: 'Mã chứng từ gốc', cols: 1 },
        { text: 'Thông tin người mua', cols: 7 },
        { text: 'Thông tin người nhận', cols: 2 },
        { text: 'Thông tin giao dịch', cols: 5 },
        { text: 'CK cả HĐ', cols: 2 },
        { text: 'VAT HĐ', cols: 1 },
        { text: 'Thông tin hàng hóa, dịch vụ', cols: 16 }
    ];
    row1Spans.forEach(function (h) {
        headerHtml += '<th colspan="' + h.cols + '">' + h.text + '</th>';
    });
    headerHtml += '</tr>';

    // Row 2 - column names
    headerHtml += '<tr class="header-row-2">';
    SAPO_HEADERS.row2.forEach(function (h) {
        headerHtml += '<th>' + h + '</th>';
    });
    headerHtml += '</tr>';

    thead.innerHTML = headerHtml;

    // Data rows
    tbody.innerHTML = '';
    AppState.sapoData.forEach(function (row) {
        var tr = document.createElement('tr');
        if (row[SAPO_COL.MA_CHUNG_TU]) tr.className = 'invoice-header-row';

        var html = '';
        var amountCols = [24, 25, 26, 28, 30, 33, 34];
        row.forEach(function (cell, colIdx) {
            var cls = amountCols.indexOf(colIdx) !== -1 ? ' class="amount-cell"' : '';
            var displayVal = cell;
            if (typeof cell === 'number' && cell > 0) {
                if (colIdx === 23) {
                    displayVal = formatQuantity(cell);
                } else if (amountCols.indexOf(colIdx) !== -1) {
                    displayVal = formatNumber(cell);
                }
            }
            html += '<td' + cls + '>' + (displayVal !== '' && displayVal !== 0 ? escapeHtml(String(displayVal)) : '') + '</td>';
        });
        tr.innerHTML = html;
        tbody.appendChild(tr);
    });
}

function updateExportDetails() {
    var totalItems = AppState.sapoData.length;
    var totalInvoices = AppState.invoices.length;
    var grandTotal = 0;
    AppState.sapoData.forEach(function (row) {
        grandTotal += (typeof row[SAPO_COL.TONG_TIEN] === 'number' ? row[SAPO_COL.TONG_TIEN] : 0);
    });

    document.getElementById('exportDetails').innerHTML =
        '<div class="export-detail">' +
        '<div class="export-detail-value">' + totalInvoices + '</div>' +
        '<div class="export-detail-label">Hóa đơn</div></div>' +
        '<div class="export-detail">' +
        '<div class="export-detail-value">' + totalItems + '</div>' +
        '<div class="export-detail-label">Dòng sản phẩm</div></div>' +
        '<div class="export-detail">' +
        '<div class="export-detail-value">' + formatNumber(grandTotal) + '</div>' +
        '<div class="export-detail-label">Tổng tiền (VNĐ)</div></div>';
}

// ============================================
// Tab Management
// ============================================
function switchTab(tabId) {
    AppState.currentTab = tabId;
    document.querySelectorAll('.tab').forEach(function (t) { t.classList.remove('active'); });
    document.querySelectorAll('.tab-content').forEach(function (c) { c.style.display = 'none'; });
    document.querySelector('[data-tab="' + tabId + '"]').classList.add('active');
    document.getElementById(tabId).style.display = 'block';
}

// ============================================
// Step Management
// ============================================
function setStep(step) {
    for (var i = 1; i <= 3; i++) {
        var el = document.getElementById('step' + i + '-indicator');
        el.classList.remove('active', 'completed');
        if (i < step) el.classList.add('completed');
        else if (i === step) el.classList.add('active');
    }
    var lines = document.querySelectorAll('.step-line');
    lines.forEach(function (line, idx) {
        if (idx < step - 1) line.classList.add('completed');
        else line.classList.remove('completed');
    });
}

// ============================================
// Utility Functions
// ============================================
function formatNumber(num) {
    if (num === 0 || num === '' || num === null || num === undefined) return '0';
    return Math.round(num).toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.');
}

function formatQuantity(num) {
    if (num === 0) return '0';
    if (num % 1 === 0) return num.toString();
    return num.toFixed(2).replace('.', ',');
}

function escapeHtml(str) {
    if (!str) return '';
    var div = document.createElement('div');
    div.appendChild(document.createTextNode(str));
    return div.innerHTML;
}

function getPaymentBadgeClass(methods) {
    if (methods.length === 0) return 'cash';
    if (methods.length > 1) return 'mixed';
    var m = methods[0].method;
    if (m === 'Chuyển khoản') return 'transfer';
    if (m === 'Thẻ') return 'card';
    return 'cash';
}

function showToast(message, type) {
    var toast = document.getElementById('toast');
    var icon = '';
    if (type === 'success') icon = '✅ ';
    else if (type === 'error') icon = '❌ ';
    else if (type === 'info') icon = 'ℹ️ ';

    toast.textContent = icon + message;
    toast.className = 'toast ' + (type || 'info');
    toast.style.display = 'flex';
    toast.classList.remove('toast-out');

    clearTimeout(toast._timeout);
    toast._timeout = setTimeout(function () {
        toast.classList.add('toast-out');
        setTimeout(function () { toast.style.display = 'none'; }, 300);
    }, 3000);
}

function removeFile(index) {
    var removedFile = AppState.files[index];
    AppState.invoices = AppState.invoices.filter(function (inv) {
        return !(inv.fileName === removedFile.name && inv.sheetName === removedFile.sheet);
    });
    AppState.files.splice(index, 1);

    if (AppState.files.length === 0) { resetApp(); return; }

    AppState.sapoData = convertToSapo(AppState.invoices);
    updateFileList();
    renderCukCukPreview();
    renderSapoPreview();
    updateExportDetails();
    showToast('Đã xóa file ' + removedFile.name, 'info');
}

function resetApp() {
    AppState.files = [];
    AppState.invoices = [];
    AppState.sapoData = [];
    document.getElementById('fileList-section').style.display = 'none';
    document.getElementById('step2-section').style.display = 'none';
    document.getElementById('step3-section').style.display = 'none';
    document.getElementById('fileInput').value = '';
    document.getElementById('fileInputMore').value = '';
    setStep(1);
    switchTab('cukcuk-preview');
}
