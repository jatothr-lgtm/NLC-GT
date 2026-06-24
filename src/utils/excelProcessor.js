import * as XLSX from 'xlsx';
import { addMonths, endOfMonth, parse, format, isValid } from 'date-fns';

const normalizeHeader = (header) => {
    if (!header) return '';
    return String(header).trim().replace(/\s+/g, ' ').toLowerCase();
};

export const processExcelFiles = async (gtFile, nlcFile, itemFile) => {
    const readFile = (file) => {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                const data = new Uint8Array(e.target.result);
                // Read without auto-parsing dates so we can extract exact strings
                const workbook = XLSX.read(data, { type: 'array', cellDates: false });
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                
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
        const gtData = await readFile(gtFile);
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
        const itemUomKey = resolveColumnKey(itemData, ['Conversion Factor (UOM Conversion Detail)']);
        const itemMrpKey = resolveColumnKey(itemData, ['Mrp', 'MRP']);
        const itemEanKey = resolveColumnKey(itemData, ['Barcode (Item Barcode)', 'EAN Code']);

        // Map Item Data
        const itemByName = {};
        const itemByCode = {};
        itemData.forEach(item => {
            const name = item[itemItemNameKey];
            const code = item[itemItemCodeKey];
            if (name) itemByName[String(name).trim().toLowerCase()] = item;
            if (code) itemByCode[String(code).trim().toLowerCase()] = item;
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
                newRow[nlcUomKey] = gtRow[gtUomKey] || '';
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
        const processedNlcData = nlcData.map(row => {
            const newRow = { ...row };
            const formattedRow = row._formatted || {};
            delete newRow._formatted;

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
                
                // Item lookup
                const itemMatch = itemByName[itemNameLower];
                if (itemMatch) {
                    newRow[nlcItemCodeKey] = newRow[nlcItemCodeKey] || itemMatch[itemItemCodeKey];
                    newRow[nlcItemGroupKey] = newRow[nlcItemGroupKey] || itemMatch[itemItemGroupKey];
                    newRow[nlcUomKey] = newRow[nlcUomKey] || itemMatch[itemUomKey];
                }
                
                // GT lookup
                const gtMatch = gtBySku[itemNameLower];
                if (gtMatch) {
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

                    newRow[nlcExFactoryKey] = gtExFactory * 1.05;
                    newRow[nlcLogisticsKey] = gtLogistics * 1.05;
                    newRow[nlcMarginPctKey] = gtMarginPct;
                    newRow[nlcNonGstFinalKey] = gtTotalCost * 1.05;
                    newRow[nlcGrandFinalKey] = gtNlcSale * 1.05;
                    newRow[nlcCostBasisKey] = gtNlcCost * 1.05;
                }
            }
            
            // Step 6 & 7: Calculations
            const marginPct = parseFloat(newRow[nlcMarginPctKey]) || 0;
            const costBasis = parseFloat(newRow[nlcCostBasisKey]) || 0;
            const nonGstFinal = parseFloat(newRow[nlcNonGstFinalKey]) || 0;

            newRow[nlcMarginValKey] = marginPct * costBasis;
            newRow[nlcGstAmountKey] = 0.05 * nonGstFinal;

            // Step 8: Second Item lookup (by Item Code)
            const itemCodeVal = newRow[nlcItemCodeKey];
            if (itemCodeVal) {
                const itemByCodeMatch = itemByCode[String(itemCodeVal).trim().toLowerCase()];
                if (itemByCodeMatch) {
                    newRow[nlcMrpKey] = itemByCodeMatch[itemMrpKey] || newRow[nlcMrpKey];
                    newRow[nlcEanKey] = itemByCodeMatch[itemEanKey] || newRow[nlcEanKey];
                }
            }

            return newRow;
        });

        // Generate Output file
        const newWorkbook = XLSX.utils.book_new();
        const newWorksheet = XLSX.utils.json_to_sheet(processedNlcData);
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
