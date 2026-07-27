import * as XLSX from 'xlsx';
import { addMonths, endOfMonth, parse, format, isValid } from 'date-fns';

const normalizeHeader = (header) => {
    if (!header) return '';
    return String(header).trim().replace(/\s+/g, ' ').toLowerCase();
};

export const processExcelFiles = async (gtFile, nlcFile, itemFile) => {
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
            if (name) itemByName[String(name).trim().toLowerCase()] = item;
        });

        // Map GT Data
        const gtBySku = {};
        gtData.forEach(gt => {
            const sku = gt[gtSkuKey];
            if (sku && String(sku).trim().toUpperCase() !== 'SKU') {
                gtBySku[String(sku).trim().toLowerCase()] = gt;
            }
        });

        // Track existing NLC items
        const nlcItems = new Set();
        nlcData.forEach(row => {
            const name = row[nlcItemNameKey];
            if (name) nlcItems.add(String(name).trim().toLowerCase());
        });

        // Step 1: Append missing SKUs
        const refNlcRow = nlcData[0] || {};
        gtData.forEach(gtRow => {
            const skuVal = gtRow[gtSkuKey];
            if (!skuVal) return;
            const sku = String(skuVal).trim();
            const skuLower = sku.toLowerCase();
            
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
                        const utc_days  = Math.floor(rawDateNum - 25569);
                        const utc_value = utc_days * 86400;
                        const date_info = new Date(utc_value * 1000);
                        let m = date_info.getUTCMonth();
                        let d = date_info.getUTCDate();
                        let y = date_info.getUTCFullYear();
                        
                        if (d <= 12) {
                            parsedDate = new Date(y, d - 1, m + 1);
                        } else {
                            parsedDate = new Date(y, m, d);
                        }
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
                const itemNameLower = String(itemNameVal).trim().toLowerCase();
                
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

                    const upliftFactor = 1.05;
                    newRow[nlcExFactoryKey] = gtExFactory * upliftFactor;
                    newRow[nlcLogisticsKey] = gtLogistics * upliftFactor;
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
                setFormula(r, nlcLogisticsKey, `${meta.logBase}*${f}`, meta.logBase * f);
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
