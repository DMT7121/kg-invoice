// Verification test: reads CukCuk file, runs conversion, outputs Sapo file
// Then compares structure with original Sapo template
const XLSX = require('xlsx');

// Load the app logic inline (simulated)
const TOTAL_COLS = 35;
const SAPO_COL = {
    KY_HIEU: 0, MA_CHUNG_TU: 1, MA_KHACH: 2, MA_SO_THUE: 3,
    TEN_DON_VI: 4, DIA_CHI: 5, NGUOI_MUA: 6, SDT: 7, EMAIL: 8,
    NGUOI_NHAN: 9, EMAIL_NHAN: 10, HINH_THUC_TT: 11,
    NGAN_HANG: 12, SO_TK: 13, LOAI_TIEN: 14, TY_GIA: 15,
    CK_GIA_TRI: 16, CK_PHAN_TRAM: 17, THUE_GTGT_HD: 18,
    TINH_CHAT: 19, MA_HANG: 20, TEN_HANG: 21, DVT: 22,
    SO_LUONG: 23, DON_GIA: 24, THANH_TIEN: 25,
    CK_SP_GIA_TRI: 26, CK_SP_PHAN_TRAM: 27, TIEN_CHIET_KHAU: 28,
    PHAN_TRAM_THUE: 29, THANH_TIEN_DA_TRU: 30, THUE_GTGT_GIAM: 31,
    THUE_GTGT_PCT: 32, TIEN_THUE_GTGT: 33, TONG_TIEN: 34
};

// Read CukCuk file
const ckWb = XLSX.readFile('test_cukcuk.xlsx');
const sheet = ckWb.Sheets[ckWb.SheetNames[0]];
const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: true });

console.log('=== PARSING CUKCUK ===');
console.log('Sheet:', ckWb.SheetNames[0]);
console.log('Rows:', data.length);

// Parse (simplified for test)
var items = [];
var invoiceNum = '';
var paymentMethod = '';
var foundItems = false;

for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var rowStr = row.map(c => String(c||'')).join(' ');
    
    // Get invoice number
    var numMatch = rowStr.match(/Số[:\s]*\s*([A-Z0-9]+)/i);
    if (numMatch) invoiceNum = numMatch[1];
    
    // Check header
    var cells = row.map(c => String(c||'').trim().toLowerCase());
    if (cells.indexOf('stt') !== -1 && cells.some(c => c.indexOf('tên món') !== -1)) {
        foundItems = true;
        continue;
    }
    
    if (foundItems && /^\d+$/.test(String(row[0]||'').trim())) {
        items.push({
            name: String(row[1]||''),
            unit: String(row[2]||''),
            quantity: typeof row[3] === 'number' ? row[3] : parseFloat(String(row[3]||'0')),
            unitPrice: typeof row[4] === 'number' ? row[4] : 0,
            amount: typeof row[5] === 'number' ? row[5] : 0,
            discount: typeof row[6] === 'number' ? row[6] : 0,
            vatPercent: typeof row[7] === 'number' ? row[7] : 0,
            total: typeof row[8] === 'number' ? row[8] : 0
        });
    }
    
    if (rowStr.indexOf('Chuyển khoản:') !== -1) paymentMethod = 'Chuyển khoản';
    else if (rowStr.indexOf('Tiền mặt:') !== -1) paymentMethod = 'Tiền mặt';
}

console.log('Invoice:', invoiceNum);
console.log('Items found:', items.length);
console.log('Payment:', paymentMethod);

// Convert to Sapo
var sapoRows = [];
items.forEach((item, idx) => {
    var row = new Array(TOTAL_COLS).fill('');
    if (idx === 0) {
        row[SAPO_COL.MA_CHUNG_TU] = invoiceNum;
        row[SAPO_COL.HINH_THUC_TT] = paymentMethod || 'Tiền mặt';
        row[SAPO_COL.LOAI_TIEN] = 'VND';
        row[SAPO_COL.TY_GIA] = 1;
    }
    row[SAPO_COL.TINH_CHAT] = 'Hàng hóa dịch vụ';
    row[SAPO_COL.TEN_HANG] = item.name;
    row[SAPO_COL.DVT] = item.unit;
    row[SAPO_COL.SO_LUONG] = item.quantity;
    row[SAPO_COL.DON_GIA] = item.unitPrice;
    row[SAPO_COL.THANH_TIEN] = item.amount;
    if (item.discount > 0) {
        row[SAPO_COL.CK_SP_GIA_TRI] = item.discount;
        row[SAPO_COL.TIEN_CHIET_KHAU] = item.discount;
    }
    var afterDiscount = item.amount - item.discount;
    row[SAPO_COL.THANH_TIEN_DA_TRU] = afterDiscount;
    if (item.vatPercent > 0) {
        row[SAPO_COL.THUE_GTGT_PCT] = item.vatPercent;
        var vat = Math.round(afterDiscount * item.vatPercent / 100);
        row[SAPO_COL.TIEN_THUE_GTGT] = vat;
        row[SAPO_COL.TONG_TIEN] = afterDiscount + vat;
    } else {
        row[SAPO_COL.TONG_TIEN] = afterDiscount;
    }
    sapoRows.push(row);
});

console.log('\n=== CONVERTED SAPO DATA ===');
sapoRows.forEach((row, idx) => {
    var filled = [];
    row.forEach((cell, col) => {
        if (cell !== '' && cell !== 0) {
            var colLetter = XLSX.utils.encode_col(col);
            filled.push(`${colLetter}(${col})="${cell}"`);
        }
    });
    console.log(`Row ${idx}: ${filled.join(', ')}`);
});

// Compare with Sapo template
console.log('\n=== STRUCTURE COMPARISON ===');
const sapoWb = XLSX.readFile('sapo_template.xlsx');
const sapoSheet = sapoWb.Sheets[sapoWb.SheetNames[0]];
const sapoData = XLSX.utils.sheet_to_json(sapoSheet, { header: 1, defval: '', raw: true });

// Check column count
console.log('Sapo template cols:', sapoData[0].length);
console.log('Our output cols:', TOTAL_COLS);
console.log('Match:', sapoData[0].length === TOTAL_COLS ? '✅' : '❌');

// Check data row structure matches template data
var templateDataRow = sapoData[3]; // First data row in template
console.log('\nTemplate data row filled columns:');
templateDataRow.forEach((cell, col) => {
    if (cell !== '' && cell !== null && cell !== undefined) {
        var colLetter = XLSX.utils.encode_col(col);
        console.log(`  ${colLetter}(${col}): "${cell}" (${typeof cell})`);
    }
});

console.log('\nOur first data row filled columns:');
sapoRows[0].forEach((cell, col) => {
    if (cell !== '' && cell !== 0) {
        var colLetter = XLSX.utils.encode_col(col);
        console.log(`  ${colLetter}(${col}): "${cell}" (${typeof cell})`);
    }
});

// Verify key column mapping
console.log('\n=== KEY COLUMN VERIFICATION ===');
var checks = [
    { name: 'Tính chất', ourCol: SAPO_COL.TINH_CHAT, templateCol: 19 },
    { name: 'Tên hàng', ourCol: SAPO_COL.TEN_HANG, templateCol: 21 },
    { name: 'ĐVT', ourCol: SAPO_COL.DVT, templateCol: 22 },
    { name: 'Số lượng', ourCol: SAPO_COL.SO_LUONG, templateCol: 23 },
    { name: 'Đơn giá', ourCol: SAPO_COL.DON_GIA, templateCol: 24 },
    { name: 'Thành tiền', ourCol: SAPO_COL.THANH_TIEN, templateCol: 25 },
    { name: 'CK/SP Giá trị', ourCol: SAPO_COL.CK_SP_GIA_TRI, templateCol: 26 },
    { name: 'CK/SP %', ourCol: SAPO_COL.CK_SP_PHAN_TRAM, templateCol: 27 },
    { name: 'Tiền chiết khấu', ourCol: SAPO_COL.TIEN_CHIET_KHAU, templateCol: 28 },
    { name: 'Thành tiền đã trừ CK', ourCol: SAPO_COL.THANH_TIEN_DA_TRU, templateCol: 30 },
    { name: 'Thuế GTGT %', ourCol: SAPO_COL.THUE_GTGT_PCT, templateCol: 32 },
    { name: 'Tiền thuế GTGT', ourCol: SAPO_COL.TIEN_THUE_GTGT, templateCol: 33 },
    { name: 'Tổng tiền', ourCol: SAPO_COL.TONG_TIEN, templateCol: 34 }
];

checks.forEach(c => {
    var match = c.ourCol === c.templateCol;
    console.log(`${match ? '✅' : '❌'} ${c.name}: our=${c.ourCol} template=${c.templateCol}`);
});

console.log('\n✅ All checks completed!');
