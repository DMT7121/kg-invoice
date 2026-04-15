/**
 * KG Invoice Converter - CukCuk to Sapo Format
 * 
 * This application reads CukCuk invoice files (Excel/CSV) and converts them
 * to the Sapo e-invoice format for import.
 */

// ============================================
// State Management
// ============================================
const AppState = {
    files: [],           // Uploaded file objects
    invoices: [],        // Parsed CukCuk invoices
    sapoData: [],        // Converted Sapo data rows
    currentTab: 'cukcuk-preview'
};

// ============================================
// Sapo Column Definitions (36 columns)
// ============================================
const SAPO_HEADERS = {
    row1: [
        'Ký hiệu*', 'Mã chứng từ gốc', 'Thông tin người mua', '', '', '', '', '', '',
        'Thông tin người nhận', '',
        'Thông tin giao dịch', '', '', '', '',
        'Chiết khấu cả hóa đơn', '',
        'Thuế GTGT cả hóa đơn (%)',
        'Thông tin hàng hóa, dịch vụ', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''
    ],
    row2: [
        '', '', 'Mã khách', 'Mã số thuế người mua', 'Tên đơn vị', 'Địa chỉ',
        'Người mua hàng', 'Số điện thoại', 'Email',
        'Tên người nhận hóa đơn', 'Email nhận hóa đơn',
        'Hình thức thanh toán*', 'Tên ngân hàng', 'Số tài khoản ngân hàng', 'Loại tiền', 'Tỷ giá',
        '', '', '',
        '', 'Tính chất', 'Mã hàng', 'Tên hàng hóa/dịch vụ', 'ĐVT', 'Số lượng', 'Đơn giá',
        'Thành tiền', 'Chiết khấu/SP', '', 'Tiền chiết khấu', '%Tính thuế',
        'Thành tiền đã trừ chiết khấu', 'Tiền thuế GTGT được giảm', 'Thuế GTGT (%)',
        'Tiền thuế GTGT', 'Tổng tiền'
    ],
    row3: [
        '', '', '', '', '', '', '', '', '',
        '', '',
        '', '', '', '', '',
        'Giá trị', '%', '',
        '', '', '', '', '', '', '',
        '', 'Giá trị', '%', '', '',
        '', '', '', '', ''
    ]
};

// Column indices for Sapo format
const SAPO_COL = {
    KY_HIEU: 0,
    MA_CHUNG_TU: 1,
    MA_KHACH: 2,
    MA_SO_THUE: 3,
    TEN_DON_VI: 4,
    DIA_CHI: 5,
    NGUOI_MUA: 6,
    SDT: 7,
    EMAIL: 8,
    NGUOI_NHAN: 9,
    EMAIL_NHAN: 10,
    HINH_THUC_TT: 11,
    NGAN_HANG: 12,
    SO_TK: 13,
    LOAI_TIEN: 14,
    TY_GIA: 15,
    CK_GIA_TRI: 16,
    CK_PHAN_TRAM: 17,
    THUE_GTGT_HD: 18,
    SPACER: 19,
    TINH_CHAT: 20,
    MA_HANG: 21,
    TEN_HANG: 22,
    DVT: 23,
    SO_LUONG: 24,
    DON_GIA: 25,
    THANH_TIEN: 26,
    CK_SP_GIA_TRI: 27,
    CK_SP_PHAN_TRAM: 28,
    TIEN_CHIET_KHAU: 29,
    PHAN_TRAM_THUE: 30,
    THANH_TIEN_DA_TRU: 31,
    THUE_GTGT_GIAM: 32,
    THUE_GTGT_PCT: 33,
    TIEN_THUE_GTGT: 34,
    TONG_TIEN: 35
};

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

    // Find all invoice blocks in the sheet
    // Each invoice starts with a row containing "Hóa đơn" or a row with "Số:" pattern
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
        if (i < data.length && row.length > 1) {
            var secondCell = String(row[0] || '') + ' ' + String(row[1] || '');
            if (secondCell.indexOf('Số:') !== -1) {
                // Check if the previous row might be the "Hóa đơn" header that we missed
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

    // If no invoices found with headers, try parsing the whole sheet as a single invoice
    if (invoices.length === 0 && data.length > 5) {
        var invoice = tryParseWholeSheet(data, fileName, sheetName);
        if (invoice) {
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

        // Parse header info
        if (!foundItems) {
            // Parse "Số: xxx, Ngày: xxx, Thu ngân: xxx"
            parseInfoRow(row, invoice);
        }

        // Check for item table header row (STT, Tên món, ĐVT, ...)
        if (isItemHeaderRow(row)) {
            foundItems = true;
            i++;
            continue;
        }

        // Parse item rows (numbered rows after header)
        if (foundItems) {
            var firstCell = String(row[0] || '').trim();
            var isNumber = /^\d+$/.test(firstCell);

            if (isNumber && row.length >= 5) {
                var item = parseItemRow(row);
                if (item) {
                    invoice.items.push(item);
                }
            } else if (rowStr.indexOf('Tổng tiền thanh toán') !== -1 || rowStr.indexOf('Tổng tiền') !== -1) {
                // Parse total
                invoice.totalAmount = findNumberInRow(row);
            }

            // Parse payment method rows
            if (rowStr.indexOf('Tiền mặt:') !== -1 || rowStr.indexOf('Tiền mặt :') !== -1) {
                var cashAmount = extractPaymentAmount(row, 'Tiền mặt');
                if (cashAmount > 0) {
                    invoice.paymentMethods.push({ method: 'Tiền mặt', amount: cashAmount });
                }
            }
            if (rowStr.indexOf('Chuyển khoản:') !== -1 || rowStr.indexOf('Chuyển khoản :') !== -1) {
                var transferAmount = extractPaymentAmount(row, 'Chuyển khoản');
                if (transferAmount > 0) {
                    invoice.paymentMethods.push({ method: 'Chuyển khoản', amount: transferAmount });
                }
            }
            if (rowStr.indexOf('Thẻ:') !== -1 || rowStr.indexOf('Thẻ :') !== -1) {
                var cardAmount = extractPaymentAmount(row, 'Thẻ');
                if (cardAmount > 0) {
                    invoice.paymentMethods.push({ method: 'Thẻ', amount: cardAmount });
                }
            }

            // Check for end of invoice block (empty rows or next invoice)
            if (isEndOfInvoice(data, i, startRow)) {
                invoice._endRow = i;
                break;
            }
        }

        invoice._endRow = i;
        i++;
    }

    // Only return if we found actual items
    if (invoice.items.length === 0) return null;

    // Determine payment method if not found in footer
    if (invoice.paymentMethods.length === 0 && invoice.totalAmount > 0) {
        invoice.paymentMethods.push({ method: 'Tiền mặt', amount: invoice.totalAmount });
    }

    return invoice;
}

function parseInfoRow(row, invoice) {
    var rowStr = row.map(function (c) { return String(c || ''); }).join('|||');

    // Parse invoice number: "Số: 2602000218"
    var numMatch = rowStr.match(/Số[:\s]*\s*([A-Z0-9]+)/i);
    if (numMatch) invoice.number = numMatch[1].trim();

    // Parse date: "Ngày: 01/04/2026 (18:23 - 19:21)"
    var dateMatch = rowStr.match(/Ngày[:\s]*\s*(\d{1,2}\/\d{1,2}\/\d{2,4})/);
    if (dateMatch) invoice.date = dateMatch[1].trim();

    var timeMatch = rowStr.match(/\((\d{1,2}:\d{2}\s*-\s*\d{1,2}:\d{2})\)/);
    if (timeMatch) invoice.timeRange = timeMatch[1].trim();

    // Parse cashier: "Thu ngân: THU NGÂN"
    var cashierMatch = rowStr.match(/Thu ngân[:\s]*\s*([^|]+)/i);
    if (cashierMatch) invoice.cashier = cashierMatch[1].trim();

    // Parse server: "Phục vụ: Nguyễn Văn Hoà"
    var serverMatch = rowStr.match(/Phục vụ[:\s]*\s*([^|]+)/i);
    if (serverMatch) invoice.server = serverMatch[1].trim();

    // Parse table: "Bàn: A.18"
    var tableMatch = rowStr.match(/Bàn[:\s]*\s*([^|]+)/i);
    if (tableMatch) invoice.table = tableMatch[1].trim();

    // Parse guest count: "Số người: 3"
    var guestMatch = rowStr.match(/Số người[:\s]*\s*(\d+)/i);
    if (guestMatch) invoice.guestCount = guestMatch[1].trim();

    // Parse customer: "KH: xxx"
    var customerMatch = rowStr.match(/KH[:\s]*\s*([^|]*)/i);
    if (customerMatch) {
        var name = customerMatch[1].trim();
        if (name && name !== 'KH:' && name !== '') {
            invoice.customerName = name;
        }
    }

    // Parse phone: "ĐT: xxx"
    var phoneMatch = rowStr.match(/ĐT[:\s]*\s*([^|]*)/i);
    if (phoneMatch) {
        var phone = phoneMatch[1].trim();
        if (phone && phone !== 'ĐT:' && phone !== '' && /\d/.test(phone)) {
            invoice.customerPhone = phone;
        }
    }

    // Parse customer card: "Thẻ KH: xxx"
    var cardMatch = rowStr.match(/Thẻ KH[:\s]*\s*([^|]*)/i);
    if (cardMatch) {
        var card = cardMatch[1].trim();
        if (card && card !== '') invoice.customerCard = card;
    }
}

function isItemHeaderRow(row) {
    var cells = row.map(function (c) { return String(c || '').trim().toLowerCase(); });
    // Check if row contains typical column headers
    var hasSTT = cells.indexOf('stt') !== -1;
    var hasTenMon = cells.some(function (c) { return c.indexOf('tên món') !== -1 || c.indexOf('ten mon') !== -1 || c === 'tên hàng'; });
    var hasDVT = cells.some(function (c) { return c === 'đvt' || c === 'dvt'; });
    return hasSTT && (hasTenMon || hasDVT);
}

function parseItemRow(row) {
    // CukCuk item row: STT, Tên món, ĐVT, Số lượng, Đơn giá, Tiền hàng, Tiền KM, % VAT, Thành tiền
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

        // Validate - must have a name at minimum
        if (!item.name) return null;

        // If amount is 0 but we have qty and price, calculate
        if (item.amount === 0 && item.quantity > 0 && item.unitPrice > 0) {
            item.amount = item.quantity * item.unitPrice;
        }

        // If total is 0, calculate
        if (item.total === 0 && item.amount > 0) {
            item.total = item.amount - item.discount;
            if (item.vatPercent > 0) {
                item.total += Math.round((item.amount - item.discount) * item.vatPercent / 100);
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
    var str = String(val).trim();

    // Remove quotes
    str = str.replace(/"/g, '');

    // Handle Vietnamese number format: "3,00" -> 3, "3.000" -> 3000
    // If contains comma and no dot: treat comma as decimal separator (e.g., "3,00" -> 3)
    // If contains dot: treat dot as thousands separator (e.g., "3.000" -> 3000)

    if (str.indexOf(',') !== -1 && str.indexOf('.') === -1) {
        // Comma as decimal: "3,00" -> "3.00"
        str = str.replace(',', '.');
    } else if (str.indexOf('.') !== -1) {
        // Dot as thousands separator: "3.000" -> "3000"
        // But also handle "3.000,00" -> "3000.00"
        if (str.indexOf(',') !== -1) {
            str = str.replace(/\./g, '').replace(',', '.');
        } else {
            // Check if dot is thousands separator (number after dot is 3 digits)
            var parts = str.split('.');
            if (parts.length >= 2) {
                var lastPart = parts[parts.length - 1];
                if (lastPart.length === 3) {
                    // Thousands separator
                    str = str.replace(/\./g, '');
                }
                // Otherwise treat dot as decimal
            }
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

function extractPaymentAmount(row, methodName) {
    var foundMethod = false;
    for (var i = 0; i < row.length; i++) {
        var cell = String(row[i] || '').trim();
        if (cell.indexOf(methodName) !== -1) {
            foundMethod = true;
        }
        if (foundMethod) {
            var val = parseCukCukNumber(row[i]);
            if (val > 0) return val;
            // Also check next cell
            if (i + 1 < row.length) {
                val = parseCukCukNumber(row[i + 1]);
                if (val > 0) return val;
            }
        }
    }
    return 0;
}

function isEndOfInvoice(data, currentRow, startRow) {
    // End of invoice if:
    // 1. We're at the end of data
    if (currentRow >= data.length - 1) return true;

    // 2. We've gone at least 8 rows into the invoice and hit 2+ consecutive empty rows
    if (currentRow - startRow > 8) {
        var currentEmpty = isRowEmpty(data[currentRow]);
        var nextEmpty = currentRow + 1 < data.length ? isRowEmpty(data[currentRow + 1]) : true;
        if (currentEmpty && nextEmpty) return true;
    }

    // 3. Next row starts a new invoice
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

function tryParseWholeSheet(data, fileName, sheetName) {
    // Fallback: try to parse the entire sheet as one invoice
    var invoice = parseInvoiceBlock(data, 0);
    if (invoice) {
        invoice.fileName = fileName;
        invoice.sheetName = sheetName;
    }
    return invoice;
}

// ============================================
// CukCuk to Sapo Converter
// ============================================
function convertToSapo(invoices) {
    var sapoRows = [];

    invoices.forEach(function (invoice) {
        var paymentMethod = determinePaymentMethod(invoice);

        invoice.items.forEach(function (item, itemIndex) {
            var row = new Array(36).fill('');

            // First item row gets the invoice header info
            if (itemIndex === 0) {
                row[SAPO_COL.MA_CHUNG_TU] = invoice.number;
                row[SAPO_COL.NGUOI_MUA] = invoice.customerName || '';
                row[SAPO_COL.SDT] = invoice.customerPhone || '';
                row[SAPO_COL.HINH_THUC_TT] = paymentMethod;
                row[SAPO_COL.LOAI_TIEN] = 'VND';
                row[SAPO_COL.TY_GIA] = 1;
            }

            // Item info (always filled)
            row[SAPO_COL.TINH_CHAT] = 'Hàng hóa dịch vụ';
            row[SAPO_COL.TEN_HANG] = item.name;
            row[SAPO_COL.DVT] = item.unit;
            row[SAPO_COL.SO_LUONG] = item.quantity;
            row[SAPO_COL.DON_GIA] = item.unitPrice;
            row[SAPO_COL.THANH_TIEN] = item.amount;

            // Discount
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

    // Multiple payment methods
    var hasCash = invoice.paymentMethods.some(function (p) { return p.method === 'Tiền mặt'; });
    var hasTransfer = invoice.paymentMethods.some(function (p) { return p.method === 'Chuyển khoản'; });

    if (hasCash && hasTransfer) return 'TM/CK';
    return 'TM/CK';
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
            '<polyline points="14,2 14,8 20,8"/>' +
            '</svg></div>' +
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

    // Header
    thead.innerHTML =
        '<tr>' +
        '<th>Hóa đơn</th>' +
        '<th>Ngày</th>' +
        '<th>HTTT</th>' +
        '<th>STT</th>' +
        '<th>Tên món</th>' +
        '<th>ĐVT</th>' +
        '<th>SL</th>' +
        '<th class="amount-cell">Đơn giá</th>' +
        '<th class="amount-cell">Tiền hàng</th>' +
        '<th class="amount-cell">Tiền KM</th>' +
        '<th>%VAT</th>' +
        '<th class="amount-cell">Thành tiền</th>' +
        '</tr>';

    // Body
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

    // Update stats
    document.getElementById('previewStats').innerHTML =
        '<span class="stat-badge"><strong>' + totalInvoices + '</strong> hóa đơn</span>' +
        '<span class="stat-badge"><strong>' + totalItems + '</strong> dòng</span>' +
        '<span class="stat-badge"><strong>' + formatNumber(grandTotal) + '</strong> VNĐ</span>';
}

function renderSapoPreview() {
    var thead = document.getElementById('sapoTableHead');
    var tbody = document.getElementById('sapoTableBody');

    // Build 3-row header
    var headerHtml = '';

    // Row 1
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

    // Row 2 - detailed column names
    headerHtml += '<tr class="header-row-2">';
    SAPO_HEADERS.row2.forEach(function (h) {
        headerHtml += '<th>' + h + '</th>';
    });
    headerHtml += '</tr>';

    thead.innerHTML = headerHtml;

    // Body
    tbody.innerHTML = '';
    AppState.sapoData.forEach(function (row) {
        var tr = document.createElement('tr');
        if (row[SAPO_COL.MA_CHUNG_TU]) tr.className = 'invoice-header-row';

        var html = '';
        row.forEach(function (cell, colIdx) {
            var cls = '';
            if ([25, 26, 27, 29, 31, 34, 35].indexOf(colIdx) !== -1) cls = ' class="amount-cell"';
            var displayVal = cell;
            if (typeof cell === 'number' && cell > 0) {
                if ([24].indexOf(colIdx) !== -1) {
                    displayVal = formatQuantity(cell);
                } else if ([25, 26, 27, 29, 31, 34, 35].indexOf(colIdx) !== -1) {
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
// Export Functions
// ============================================
function exportToXlsx() {
    try {
        var wb = XLSX.utils.book_new();
        var wsData = buildSapoSheetData();
        var ws = XLSX.utils.aoa_to_sheet(wsData);

        // Set column widths
        ws['!cols'] = [
            { wch: 12 }, // Ký hiệu
            { wch: 16 }, // Mã chứng từ
            { wch: 10 }, // Mã khách
            { wch: 16 }, // MST
            { wch: 25 }, // Tên đơn vị
            { wch: 30 }, // Địa chỉ
            { wch: 20 }, // Người mua
            { wch: 14 }, // SĐT
            { wch: 20 }, // Email
            { wch: 20 }, // Người nhận
            { wch: 20 }, // Email nhận
            { wch: 15 }, // HTTT
            { wch: 15 }, // Ngân hàng
            { wch: 16 }, // STK
            { wch: 8 },  // Loại tiền
            { wch: 8 },  // Tỷ giá
            { wch: 10 }, // CK giá trị
            { wch: 6 },  // CK %
            { wch: 8 },  // VAT HĐ
            { wch: 3 },  // spacer
            { wch: 16 }, // Tính chất
            { wch: 12 }, // Mã hàng
            { wch: 30 }, // Tên HH
            { wch: 8 },  // ĐVT
            { wch: 8 },  // SL
            { wch: 12 }, // Đơn giá
            { wch: 12 }, // Thành tiền
            { wch: 10 }, // CK/SP GiaTri
            { wch: 6 },  // CK/SP %
            { wch: 12 }, // Tiền CK
            { wch: 8 },  // %Thuế
            { wch: 14 }, // TT đã trừ CK
            { wch: 12 }, // Thuế giảm
            { wch: 8 },  // VAT%
            { wch: 12 }, // Tiền VAT
            { wch: 14 }  // Tổng tiền
        ];

        // Merge cells for header row 1
        ws['!merges'] = [
            // Thông tin người mua: C1:I1
            { s: { r: 0, c: 2 }, e: { r: 0, c: 8 } },
            // Thông tin người nhận: J1:K1
            { s: { r: 0, c: 9 }, e: { r: 0, c: 10 } },
            // Thông tin giao dịch: L1:P1
            { s: { r: 0, c: 11 }, e: { r: 0, c: 15 } },
            // Chiết khấu cả HĐ: Q1:R1
            { s: { r: 0, c: 16 }, e: { r: 0, c: 17 } },
            // Thông tin HH, DV: T1:AJ1
            { s: { r: 0, c: 19 }, e: { r: 0, c: 35 } },
            // Chiết khấu/SP header: AB2:AC2
            { s: { r: 1, c: 27 }, e: { r: 1, c: 28 } }
        ];

        XLSX.utils.book_append_sheet(wb, ws, 'Sapo_Import');

        var now = new Date();
        var dateStr = now.getFullYear().toString() +
            ('0' + (now.getMonth() + 1)).slice(-2) +
            ('0' + now.getDate()).slice(-2) +
            '_' + ('0' + now.getHours()).slice(-2) +
            ('0' + now.getMinutes()).slice(-2);
        var fileName = 'Sapo_Import_' + dateStr + '.xlsx';

        XLSX.writeFile(wb, fileName);
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
                // Escape quotes and wrap in quotes if contains comma, quote or newline
                if (str.indexOf(',') !== -1 || str.indexOf('"') !== -1 || str.indexOf('\n') !== -1) {
                    str = '"' + str.replace(/"/g, '""') + '"';
                }
                return str;
            }).join(',');
        }).join('\r\n');

        // Add BOM for UTF-8
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
        a.click();
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
    data.push(SAPO_HEADERS.row1);
    data.push(SAPO_HEADERS.row2);
    data.push(SAPO_HEADERS.row3);

    // Data rows
    AppState.sapoData.forEach(function (row) {
        data.push(row.slice());
    });

    return data;
}

// ============================================
// Tab Management
// ============================================
function switchTab(tabId) {
    AppState.currentTab = tabId;

    document.querySelectorAll('.tab').forEach(function (t) {
        t.classList.remove('active');
    });
    document.querySelectorAll('.tab-content').forEach(function (c) {
        c.style.display = 'none';
    });

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

        if (i < step) {
            el.classList.add('completed');
        } else if (i === step) {
            el.classList.add('active');
        }
    }

    // Update step lines
    var lines = document.querySelectorAll('.step-line');
    lines.forEach(function (line, idx) {
        if (idx < step - 1) {
            line.classList.add('completed');
        } else {
            line.classList.remove('completed');
        }
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
    // Show decimal only if fractional
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

    // Reset animation
    toast.classList.remove('toast-out');

    clearTimeout(toast._timeout);
    toast._timeout = setTimeout(function () {
        toast.classList.add('toast-out');
        setTimeout(function () {
            toast.style.display = 'none';
        }, 300);
    }, 3000);
}

function removeFile(index) {
    var removedFile = AppState.files[index];

    // Remove invoices from this file/sheet
    AppState.invoices = AppState.invoices.filter(function (inv) {
        return !(inv.fileName === removedFile.name && inv.sheetName === removedFile.sheet);
    });

    AppState.files.splice(index, 1);

    if (AppState.files.length === 0) {
        resetApp();
        return;
    }

    // Reconvert
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
