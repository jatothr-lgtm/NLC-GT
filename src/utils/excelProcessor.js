import * as XLSX from 'xlsx';
import { addMonths, endOfMonth, parse, format, isValid } from 'date-fns';

const normalizeHeader = (header) => {
    if (!header) return '';
    return String(header).trim().replace(/\s+/g, ' ').toLowerCase();
};

// Robust key for matching Item Names / SKUs across files. Strips the stray "Â"
// left by a double-encoded non-breaking space and collapses all whitespace
// (including nbsp  ) so mojibake variants of the same name match. Without
// this, a GT SKU whose name differs only by such artifacts is treated as new
// and appended as a duplicate row.
const normalizeKey = (v) => {
    if (v == null) return '';
    return String(v)
        .replace(/Â/g, ' ')
        .replace(/\s+/g, ' ')
        .trim()
        .toLowerCase();
};

// Default per-Item-Code uplift master (percent). These are the seed values for
// the editable "Uplift % Master" in the app UI. code = ERP Item Code, pct =
// uplift % applied to the GT-derived cost columns (Logistics excluded).
export const DEFAULT_UPLIFT_ROWS = [
    { code: 'Dates_4-3011',   name: 'Barkat dates Farmley Standee Pouch 250 g', pct: 13 },
    { code: 'Dates_4-2472',   name: 'Barkat Dates Farmley Standee Pouch 500g', pct: 13 },
    { code: 'Seeds_11-2932',  name: 'Basil Seeds Farmley Standee pouch 300g', pct: 8 },
    { code: 'Seeds_11-2693',  name: 'Premium Chia Seeds Farmley Jar 1 Kg', pct: 8 },
    { code: 'Seeds_11-2541',  name: 'Premium Chia Seeds Farmley Standee Pouch 200 g', pct: 8 },
    { code: 'Seeds_11-2550',  name: 'Premium Flax Seeds Farmley Standee Pouch 200 g', pct: 8 },
    { code: 'Seeds_11-2542',  name: 'Premium Jumbo Pumpkin Seeds Farmley Standee Pouch 200 g', pct: 8 },
    { code: 'Dates_4-2469',   name: 'Premium Omani Fard Dates Farmley standee pouch 400 g', pct: 8 },
    { code: 'Seeds_11-2549',  name: 'Premium Sunflower Seeds Farmley Standee Pouch 200 g', pct: 8 },
    { code: 'Seeds_11-20155', name: 'Premium Watermelon Seeds Farmley Standee Pouch 100 gms', pct: 8 },
    { code: 'Seeds_11-2867',  name: 'Quinoa seeds Farmley Jar 1kg', pct: 8 },
    { code: 'Seeds_11-2868',  name: 'Quinoa seeds Farmley Standee Pouch 500g', pct: 8 },
    { code: 'Seeds_11-2953',  name: 'Watermelon Seeds Farmley Standee Pouch 500g', pct: 8 },
];

// Built-in uplift map: lowercased Item Code -> factor (1 + pct/100). Used when
// the UI does not supply an override master. Every other GT-matched row = 1.05.
const ITEM_CODE_UPLIFT = Object.fromEntries(
    DEFAULT_UPLIFT_ROWS.map(r => [String(r.code).trim().toLowerCase(), 1 + r.pct / 100])
);

export const processExcelFiles = async (gtFile, nlcFile, itemFile, options = {}) => {
    // Optional uplift override master from the UI: { lowercasedItemCode: factor }.
    // When provided (non-empty) it fully drives the per-item uplift; otherwise
    // the built-in ITEM_CODE_UPLIFT defaults apply. Rows not listed use 1.05.
    const overrideMap = (options && options.upliftOverrides) || null;
    const activeUplift = (overrideMap && Object.keys(overrideMap).length) ? overrideMap : ITEM_CODE_UPLIFT;
    const readFile = (file, sheetSelector) => {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                const data = new Uint8Array(e.target.result);
                // Read without auto-parsing dates so we can extract exact strings
                const workbook = XLSX.read(data, { type: 'array', cellDates: false });
                const sheetName = (sheetSelector && sheetSelector(workbook)) || workbook.SheetNames[0];
                const worksheet = workbook.Sheets[sheetName];
                
                // Get raw values for perfect math, and formatted values to read dates exactly as displayed
                const rawJson = XLSX.utils.sheet_to_json(worksheet, { defval: '', raw: true });
                const formattedJson = XLSX.utils.sheet_to_json(worksheet, { defval: '', raw: false, dateNF: 'dd-MM-yyyy' });
                
                const json = rawJson.map((row, index) => {
                    row._formatted = formattedJson[index];
                    return row;
                });
                
                resolve(json);
            };
            reader.onerror = (e) => reject(e);
            reader.readAsArrayBuffer(file);
        });
    };

    // For multi-sheet costing workbooks, pick the GT channel sheet instead of
    // blindly using the first sheet. Prefer a sheet literally named "GT";
    // otherwise choose the sheet whose header area best matches the GT costing
    // schema (SKU + pricing columns). Falls back to the first sheet.
    const pickGtSheet = (workbook) => {
        const names = workbook.SheetNames || [];
        if (names.length <= 1) return names[0];

        const named = names.find(n => normalizeHeader(n) === 'gt');
        if (named) return named;

        const signature = ['sku', 'ex-factory', 'total cost (per kg)', 'nlc per kg (sale basis)', 'nlc per kg (cost basis)'];
        let best = names[0];
        let bestScore = -1;
        for (const name of names) {
            const rows = XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, defval: '', raw: false });
            const found = new Set();
            for (let i = 0; i < Math.min(15, rows.length); i++) {
                const row = rows[i] || [];
                for (const cell of row) {
                    const nc = normalizeHeader(cell);
                    if (signature.includes(nc)) found.add(nc);
                }
            }
            if (found.size > bestScore) {
                bestScore = found.size;
                best = name;
            }
        }
        return best;
    };

    const resolveColumnKey = (data, possibleNames) => {
        if (!data || !data.length) return possibleNames[0];
        
        // 1. Try to find in the actual object keys
        const keys = Object.keys(data[0]);
        for (const key of keys) {
            const normalizedKey = normalizeHeader(key);
            if (possibleNames.some(name => normalizeHeader(name) === normalizedKey)) {
                return key;
            }
        }
        
        // 2. Search first 15 rows for the column header text
        for (let i = 0; i < Math.min(15, data.length); i++) {
            const rowKeys = Object.keys(data[i]);
            for (const key of rowKeys) {
                const val = String(data[i][key]).trim();
                if (possibleNames.some(name => normalizeHeader(name) === normalizeHeader(val))) {
                    return key; // This key points to the column!
                }
            }
        }
        
        return possibleNames[0]; // Fallback
    };

    try {
        const gtData = await readFile(gtFile, pickGtSheet);
        const nlcData = await readFile(nlcFile);
        const itemData = await readFile(itemFile);

        // NLC File Columns
        const nlcItemNameKey = resolveColumnKey(nlcData, ['Item Name']);
        const nlcFromDateKey = resolveColumnKey(nlcData, ['From Date']);
        const nlcToDateKey = resolveColumnKey(nlcData, ['To Date']);
        const nlcMonthKey = resolveColumnKey(nlcData, ['Month']);
        const nlcCustomerGroupKey = resolveColumnKey(nlcData, ['Customer Group']);
        const nlcGstPercentKey = resolveColumnKey(nlcData, ['GST %']);
        const nlcItemCodeKey = resolveColumnKey(nlcData, ['Item Code']);
        const nlcItemGroupKey = resolveColumnKey(nlcData, ['Item Group']);
        const nlcUomKey = resolveColumnKey(nlcData, ['UOM']);
        const nlcExFactoryKey = resolveColumnKey(nlcData, ['Ex-Factory Cost Per Kg', 'Ex-Factory']);
        const nlcLogisticsKey = resolveColumnKey(nlcData, ['Logistics Cost']);
        const nlcMarginPctKey = resolveColumnKey(nlcData, ['Margin Percentage']);
        const nlcMarginValKey = resolveColumnKey(nlcData, ['Margin']);
        const nlcGstAmountKey = resolveColumnKey(nlcData, ['GST Amount']);
        const nlcNonGstFinalKey = resolveColumnKey(nlcData, ['Non-GST Final (Ex-Factory + Capital cost + Logistics + Margin) (Per kg)', 'Non-GST Final \r\n(Ex-Factory + Capital cost\r\n+ Logistics + Margin) \r\n(Per kg)', 'Non-GST Final']);
        const nlcGrandFinalKey = resolveColumnKey(nlcData, ['Grand Final (Ex-Factory + Capital cost + Logistics + Margin + GST) (Per kg)', 'Grand Final\r\n(Ex-Factory + Capital cost +\r\nLogistics + Margin +\r\nGST) \r\n(Per kg)', 'Grand Final']);
        const nlcCostBasisKey = resolveColumnKey(nlcData, ['Ex-Factory+ Capital Cost+ Logistics+ GST (per kg) For Margin calculation only', 'Ex-Factory+ Capital Cost+ Logistics+ GST \r\n(per kg) \r\nFor Margin calculation only']);
        const nlcMrpKey = resolveColumnKey(nlcData, ['MRP']);
        const nlcEanKey = resolveColumnKey(nlcData, ['EAN Code']);

        // GT File Columns
        const gtSkuKey = resolveColumnKey(gtData, ['SKU']);
        const gtGroupKey = resolveColumnKey(gtData, ['Group', 'Item Group']);
        const gtUomKey = resolveColumnKey(gtData, ['UOM (G)', 'UOM']);
        const gtExFactoryKey = resolveColumnKey(gtData, ['Ex-Factory']);
        const gtLogisticsKey = resolveColumnKey(gtData, ['Logistics Cost']);
        const gtMarginKey = resolveColumnKey(gtData, ['Margin']);
        const gtTotalCostKey = resolveColumnKey(gtData, ['Total Cost (per KG)']);
        const gtNlcSaleKey = resolveColumnKey(gtData, ['NLC PER KG (SALE BASIS)']);
        const gtNlcCostKey = resolveColumnKey(gtData, ['NLC PER KG (COST BASIS)']);

        // Item File Columns
        const itemItemNameKey = resolveColumnKey(itemData, ['Item Name']);
        const itemItemCodeKey = resolveColumnKey(itemData, ['Item Code']);
        const itemItemGroupKey = resolveColumnKey(itemData, ['Item Group']);
        const itemMrpKey = resolveColumnKey(itemData, ['Mrp', 'MRP']);
        const itemEanKey = resolveColumnKey(itemData, ['Barcode (Item Barcode)', 'EAN Code']);

        // Map Item Data (VLOOKUP by Item Name)
        const itemByName = {};
        itemData.forEach(item => {
            const name = item[itemItemNameKey];
            if (name) itemByName[normalizeKey(name)] = item;
        });

        // Map GT Data
        const gtBySku = {};
        gtData.forEach(gt => {
            const sku = gt[gtSkuKey];
            if (sku && String(sku).trim().toUpperCase() !== 'SKU') {
                gtBySku[normalizeKey(sku)] = gt;
            }
        });

        // Track existing NLC items
        const nlcItems = new Set();
        nlcData.forEach(row => {
            const name = row[nlcItemNameKey];
            if (name) nlcItems.add(normalizeKey(name));
        });

        // Step 1: Append missing SKUs
        const refNlcRow = nlcData[0] || {};
        gtData.forEach(gtRow => {
            const skuVal = gtRow[gtSkuKey];
            if (!skuVal) return;
            const sku = String(skuVal).trim();
            const skuLower = normalizeKey(sku);

            if (sku.toUpperCase() !== 'SKU' && !nlcItems.has(skuLower)) {
                const newRow = {};
                
                // Copy dates from reference row
                newRow[nlcFromDateKey] = refNlcRow[nlcFromDateKey];
                newRow[nlcToDateKey] = refNlcRow[nlcToDateKey];
                newRow[nlcMonthKey] = refNlcRow[nlcMonthKey];
                if (refNlcRow._formatted) newRow._formatted = { ...refNlcRow._formatted };
                
                newRow[nlcItemNameKey] = skuVal;
                newRow[nlcItemGroupKey] = gtRow[gtGroupKey] || '';
                // UOM = GT "UOM (G)" / 1000
                const appendUom = parseFloat(gtRow[gtUomKey]);
                newRow[nlcUomKey] = isNaN(appendUom) ? '' : appendUom / 1000;
                newRow[nlcCustomerGroupKey] = 'GT';
                
                const itemMatch = itemByName[skuLower];
                if (itemMatch) {
                    newRow[nlcItemCodeKey] = itemMatch[itemItemCodeKey] || '';
                }
                
                nlcData.push(newRow);
                nlcItems.add(skuLower);
            }
        });

        // Process all NLC rows
        const rowMeta = [];
        const processedNlcData = nlcData.map((row, rowIndex) => {
            const newRow = { ...row };
            const formattedRow = row._formatted || {};
            delete newRow._formatted;

            const meta = { hasGt: false };

            // Step 2: Date Math
            if (newRow[nlcFromDateKey]) {
                let rawDateStr = formattedRow[nlcFromDateKey] || newRow[nlcFromDateKey];
                let rawDateNum = newRow[nlcFromDateKey];
                let parsedDate = null;

                if (rawDateStr) {
                    let str = String(rawDateStr).trim();
                    const match = str.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})/);
                    if (match) {
                        let d = parseInt(match[1], 10);
                        let m = parseInt(match[2], 10) - 1;
                        let y = parseInt(match[3], 10);
                        if (y < 100) y += 2000;
                        parsedDate = new Date(y, m, d);
                    } else if (typeof rawDateNum === 'number') {
                        // Excel serial numbers are unambiguous: decode directly to
                        // y/m/d. (A previous day<=12 "swap" heuristic corrupted every
                        // date with day-of-month <= 12, e.g. turning 01-08 into 08-01.)
                        const utc_days  = Math.floor(rawDateNum - 25569);
                        const utc_value = utc_days * 86400;
                        const date_info = new Date(utc_value * 1000);
                        const m = date_info.getUTCMonth();
                        const d = date_info.getUTCDate();
                        const y = date_info.getUTCFullYear();
                        parsedDate = new Date(y, m, d);
                    } else {
                        parsedDate = new Date(str);
                    }
                }

                if (parsedDate && isValid(parsedDate)) {
                    const nextMonthDate = addMonths(parsedDate, 1);
                    const endOfNextMonth = endOfMonth(nextMonthDate);
                    
                    newRow[nlcFromDateKey] = format(nextMonthDate, 'dd-MM-yyyy');
                    newRow[nlcToDateKey] = format(endOfNextMonth, 'dd-MM-yyyy');
                    
                    if (nlcMonthKey || newRow['Month'] !== undefined) {
                        const actualMonthKey = nlcMonthKey || 'Month';
                        newRow[actualMonthKey] = format(nextMonthDate, 'MMMM');
                    }
                }
            }

            // Step 3: STATIC VALUES (GUARANTEED FOR ALL ROWS)
            newRow[nlcCustomerGroupKey] = 'GT';
            newRow[nlcGstPercentKey] = 0.05;

            // Step 4 & 5: Lookups
            const itemNameVal = newRow[nlcItemNameKey];
            if (itemNameVal) {
                const itemNameLower = normalizeKey(itemNameVal);
                
                // Item lookup (VLOOKUP Item Name: NLC <-> Item List) -> Item Code, Item Group, MRP, Barcode/EAN
                const itemMatch = itemByName[itemNameLower];
                if (itemMatch) {
                    newRow[nlcItemCodeKey] = newRow[nlcItemCodeKey] || itemMatch[itemItemCodeKey];
                    newRow[nlcItemGroupKey] = newRow[nlcItemGroupKey] || itemMatch[itemItemGroupKey];
                    newRow[nlcMrpKey] = itemMatch[itemMrpKey] || newRow[nlcMrpKey];
                    newRow[nlcEanKey] = itemMatch[itemEanKey] || newRow[nlcEanKey];
                }

                // GT lookup (VLOOKUP Item Name <-> GT SKU)
                const gtMatch = gtBySku[itemNameLower];
                if (gtMatch) {
                    // UOM = GT "UOM (G)" / 1000
                    const gtUomG = parseFloat(gtMatch[gtUomKey]);
                    if (!isNaN(gtUomG)) {
                        newRow[nlcUomKey] = gtUomG / 1000;
                    }

                    const gtExFactory = parseFloat(gtMatch[gtExFactoryKey]) || 0;
                    const gtLogistics = parseFloat(gtMatch[gtLogisticsKey]) || 0;
                    const gtTotalCost = parseFloat(gtMatch[gtTotalCostKey]) || 0;
                    const gtNlcSale = parseFloat(gtMatch[gtNlcSaleKey]) || 0;
                    const gtNlcCost = parseFloat(gtMatch[gtNlcCostKey]) || 0;
                    
                    let gtMarginRaw = gtMatch[gtMarginKey];
                    let gtMarginPct = 0;
                    if (typeof gtMarginRaw === 'string' && gtMarginRaw.includes('%')) {
                        gtMarginPct = parseFloat(gtMarginRaw.replace('%', '')) / 100;
                    } else if (gtMarginRaw) {
                        gtMarginPct = parseFloat(gtMarginRaw);
                        if (gtMarginPct > 1) gtMarginPct = gtMarginPct / 100;
                    }

                    // Uplift factor: per-Item-Code override table, else 5% default.
                    const itemCodeKey = String(newRow[nlcItemCodeKey] == null ? '' : newRow[nlcItemCodeKey]).trim().toLowerCase();
                    const upliftFactor = activeUplift[itemCodeKey] !== undefined ? activeUplift[itemCodeKey] : 1.05;

                    newRow[nlcExFactoryKey] = gtExFactory * upliftFactor;
                    newRow[nlcLogisticsKey] = gtLogistics; // kept as-is, no uplift
                    newRow[nlcMarginPctKey] = gtMarginPct;
                    newRow[nlcNonGstFinalKey] = gtTotalCost * upliftFactor;
                    newRow[nlcGrandFinalKey] = gtNlcSale * upliftFactor;
                    newRow[nlcCostBasisKey] = gtNlcCost * upliftFactor;

                    // Stash bases + factor for building formulas in the output sheet
                    meta.hasGt = true;
                    meta.factor = upliftFactor;
                    meta.exBase = gtExFactory;
                    meta.logBase = gtLogistics;
                    meta.totalBase = gtTotalCost;
                    meta.saleBase = gtNlcSale;
                    meta.costBase = gtNlcCost;
                }
            }
            
            // Step 6 & 7: Calculations
            const marginPct = parseFloat(newRow[nlcMarginPctKey]) || 0;
            const costBasis = parseFloat(newRow[nlcCostBasisKey]) || 0;
            const nonGstFinal = parseFloat(newRow[nlcNonGstFinalKey]) || 0;

            newRow[nlcMarginValKey] = marginPct * costBasis;
            newRow[nlcGstAmountKey] = 0.05 * nonGstFinal;

            meta.marginPct = marginPct;
            meta.costBasis = costBasis;
            meta.nonGstFinal = nonGstFinal;
            meta.gstPct = 0.05;
            rowMeta[rowIndex] = meta;

            return newRow;
        });

        // Generate Output file
        const newWorkbook = XLSX.utils.book_new();
        const newWorksheet = XLSX.utils.json_to_sheet(processedNlcData);

        // Inject live Excel formulas (with cached values) so the sheet recalculates in-app.
        // - Uplift columns: GT base * factor (1.05)
        // - Margin      : Margin % * Cost Basis   (in-sheet cell references)
        // - GST Amount  : GST %   * Non-GST Final (in-sheet cell references)
        const range = XLSX.utils.decode_range(newWorksheet['!ref']);
        const colByHeader = {};
        for (let c = range.s.c; c <= range.e.c; c++) {
            const headerCell = newWorksheet[XLSX.utils.encode_cell({ r: 0, c })];
            if (headerCell && headerCell.v != null) colByHeader[String(headerCell.v)] = c;
        }
        const colLetter = (headerKey) => {
            const c = colByHeader[headerKey];
            return c == null ? null : XLSX.utils.encode_col(c);
        };
        const setFormula = (r, headerKey, formula, cachedValue) => {
            const c = colByHeader[headerKey];
            if (c == null || !isFinite(cachedValue)) return;
            newWorksheet[XLSX.utils.encode_cell({ r, c })] = { t: 'n', f: formula, v: cachedValue };
        };

        for (let i = 0; i < processedNlcData.length; i++) {
            const meta = rowMeta[i];
            if (!meta) continue;
            const r = i + 1;         // row 0 is the header row
            const excelRow = r + 1;  // 1-based row number for A1-style refs

            if (meta.hasGt) {
                const f = meta.factor;
                setFormula(r, nlcExFactoryKey, `${meta.exBase}*${f}`, meta.exBase * f);
                // Logistics Cost is kept as-is (no uplift) -> plain value, no formula
                setFormula(r, nlcNonGstFinalKey, `${meta.totalBase}*${f}`, meta.totalBase * f);
                setFormula(r, nlcGrandFinalKey, `${meta.saleBase}*${f}`, meta.saleBase * f);
                setFormula(r, nlcCostBasisKey, `${meta.costBase}*${f}`, meta.costBase * f);
            }

            const pctCol = colLetter(nlcMarginPctKey);
            const costCol = colLetter(nlcCostBasisKey);
            if (pctCol && costCol) {
                setFormula(r, nlcMarginValKey, `${pctCol}${excelRow}*${costCol}${excelRow}`, meta.marginPct * meta.costBasis);
            }

            const gstPctCol = colLetter(nlcGstPercentKey);
            const nonGstCol = colLetter(nlcNonGstFinalKey);
            if (gstPctCol && nonGstCol) {
                setFormula(r, nlcGstAmountKey, `${gstPctCol}${excelRow}*${nonGstCol}${excelRow}`, meta.gstPct * meta.nonGstFinal);
            }
        }

        XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "NLC_Processed");
        
        // Write to buffer and trigger download
        const excelBuffer = XLSX.write(newWorkbook, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([excelBuffer], {type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
        
        return blob;

    } catch (error) {
        console.error("Error processing files:", error);
        throw error;
    }
};

// ---------------------------------------------------------------------------
// Master 2 helper: given the GT file and a list of { code, name, mrp, rate }
// rows, compute the uplift % required so the app output matches each new rate.
// Required % = (new rate / GT "NLC per pkt (sale basis)") - 1, matched by SKU
// name. currentPercents is { lowercasedItemCode: pct } used to flag changes.
// Returns [{ code, name, mrp, rate, base, pct, current, needsChange, matched }].
// ---------------------------------------------------------------------------
export const computeRateSuggestions = (gtFile, rows, currentPercents = {}) => {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array', cellDates: false });
                const names = wb.SheetNames || [];

                // Pick the GT channel sheet (prefer one named "GT", else score).
                let sheetName = names.find(n => normalizeHeader(n) === 'gt');
                if (!sheetName) {
                    const sig = ['sku', 'nlc per pkt (sale basis)', 'nlc per kg (sale basis)'];
                    let best = names[0], bestScore = -1;
                    names.forEach(nm => {
                        const rws = XLSX.utils.sheet_to_json(wb.Sheets[nm], { header: 1, defval: '', raw: false });
                        const found = new Set();
                        for (let i = 0; i < Math.min(15, rws.length); i++) {
                            (rws[i] || []).forEach(c => { const n = normalizeHeader(c); if (sig.includes(n)) found.add(n); });
                        }
                        if (found.size > bestScore) { bestScore = found.size; best = nm; }
                    });
                    sheetName = best;
                }

                const data = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: '', raw: true });
                const resolveKey = (poss) => {
                    if (!data.length) return poss[0];
                    for (const k of Object.keys(data[0])) {
                        if (poss.some(p => normalizeHeader(p) === normalizeHeader(k))) return k;
                    }
                    for (let i = 0; i < Math.min(15, data.length); i++) {
                        for (const k of Object.keys(data[i])) {
                            if (poss.some(p => normalizeHeader(p) === normalizeHeader(data[i][k]))) return k;
                        }
                    }
                    return poss[0];
                };
                const skuKey = resolveKey(['SKU']);
                const pktSaleKey = resolveKey(['NLC PER PKT (SALE BASIS)', 'NLC per pkt (sale basis)']);

                const bySku = {};
                data.forEach(r => {
                    const s = r[skuKey];
                    if (s && String(s).trim().toUpperCase() !== 'SKU') bySku[normalizeKey(s)] = r;
                });

                const out = (rows || []).map(row => {
                    const gt = bySku[normalizeKey(row.name)];
                    let base = gt ? parseFloat(gt[pktSaleKey]) : NaN;
                    if (isNaN(base)) base = null;
                    const rate = parseFloat(row.rate);
                    let pct = (base && !isNaN(rate)) ? ((rate / base) - 1) * 100 : null;
                    if (pct != null) pct = Math.round(pct * 100) / 100;
                    const codeKey = String(row.code || '').trim().toLowerCase();
                    const current = Object.prototype.hasOwnProperty.call(currentPercents, codeKey) ? currentPercents[codeKey] : null;
                    const needsChange = pct != null && (current == null || Math.abs(pct - current) > 0.5);
                    return { code: row.code, name: row.name, mrp: row.mrp, rate: row.rate, base, pct, current, needsChange, matched: !!gt };
                });
                resolve(out);
            } catch (err) {
                reject(err);
            }
        };
        reader.onerror = (err) => reject(err);
        reader.readAsArrayBuffer(gtFile);
    });
};
