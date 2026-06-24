import * as XLSX from 'xlsx';
import { addMonths, endOfMonth, parse, format, isValid } from 'date-fns';

const normalizeHeader = (header) => {
    if (!header) return '';
    return header.toString().trim().replace(/\s+/g, ' ');
};

const findColumnKey = (row, possibleNames) => {
    const keys = Object.keys(row);
    for (const key of keys) {
        const normalizedKey = normalizeHeader(key);
        if (possibleNames.some(name => normalizeHeader(name) === normalizedKey)) {
            return key;
        }
    }
    return possibleNames[0]; // Fallback to the first possible name
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

    try {
        const gtData = await readFile(gtFile);
        const nlcData = await readFile(nlcFile);
        const itemData = await readFile(itemFile);

        // Map Item Data by Item Name and Item Code for quick lookup
        const itemByName = {};
        const itemByCode = {};
        itemData.forEach(item => {
            const itemNameKey = findColumnKey(item, ['Item Name']);
            const itemCodeKey = findColumnKey(item, ['Item Code']);
            if (item[itemNameKey]) itemByName[item[itemNameKey]] = item;
            if (item[itemCodeKey]) itemByCode[item[itemCodeKey]] = item;
        });

        // Map GT Data by SKU
        const gtBySku = {};
        gtData.forEach(gt => {
            const skuKey = findColumnKey(gt, ['SKU']);
            if (gt[skuKey]) gtBySku[gt[skuKey]] = gt;
        });

        // Track existing NLC items
        const nlcItems = new Set();
        const nlcItemNameKey = findColumnKey(nlcData[0] || {}, ['Item Name']);
        nlcData.forEach(row => {
            if (row[nlcItemNameKey]) nlcItems.add(row[nlcItemNameKey]);
        });

        // Step 1: Append missing SKUs from GT to NLC
        const gtSkuKey = findColumnKey(gtData[0] || {}, ['SKU']);
        const gtGroupKey = findColumnKey(gtData[0] || {}, ['Group', 'Item Group']);
        const gtUomKey = findColumnKey(gtData[0] || {}, ['UOM (G)', 'UOM']);
        
        const refNlcRow = nlcData[0] || {};
        const nlcFromDateKey = findColumnKey(refNlcRow, ['From Date']);
        const nlcToDateKey = findColumnKey(refNlcRow, ['To Date']);
        const nlcMonthKey = findColumnKey(refNlcRow, ['Month']);
        const nlcItemGroupKey = findColumnKey(refNlcRow, ['Item Group']);
        const nlcUomKey = findColumnKey(refNlcRow, ['UOM']);
        const nlcCustomerGroupKey = findColumnKey(refNlcRow, ['Customer Group']);
        const nlcItemCodeKey = findColumnKey(refNlcRow, ['Item Code']);
        
        gtData.forEach(gtRow => {
            const sku = gtRow[gtSkuKey];
            if (sku && !nlcItems.has(sku)) {
                const newRow = {};
                
                // Copy reference dates so date processing works seamlessly
                if (refNlcRow[nlcFromDateKey]) newRow[nlcFromDateKey] = refNlcRow[nlcFromDateKey];
                if (refNlcRow[nlcToDateKey]) newRow[nlcToDateKey] = refNlcRow[nlcToDateKey];
                if (refNlcRow[nlcMonthKey]) newRow[nlcMonthKey] = refNlcRow[nlcMonthKey];
                if (refNlcRow._formatted) newRow._formatted = { ...refNlcRow._formatted };
                
                // Set explicitly requested fields
                newRow[nlcItemNameKey] = sku;
                newRow[nlcItemGroupKey] = gtRow[gtGroupKey] || '';
                newRow[nlcUomKey] = gtRow[gtUomKey] || '';
                newRow[nlcCustomerGroupKey] = 'GT';
                
                // Fetch Item Code from itemData
                const itemMatch = itemByName[sku];
                if (itemMatch && nlcItemCodeKey) {
                    const itemCodeVal = itemMatch[findColumnKey(itemMatch, ['Item Code'])];
                    newRow[nlcItemCodeKey] = itemCodeVal || '';
                }
                
                nlcData.push(newRow);
                nlcItems.add(sku);
            }
        });

        // Process all NLC rows
        const processedNlcData = nlcData.map(row => {
            const newRow = { ...row };
            const formattedRow = row._formatted || {};
            delete newRow._formatted;

            // Step 2: Date Math
            const fromDateKey = findColumnKey(newRow, ['From Date']);
            const toDateKey = findColumnKey(newRow, ['To Date']);
            const monthKey = findColumnKey(newRow, ['Month']);
            
            if (newRow[fromDateKey]) {
                let rawDateStr = formattedRow[fromDateKey] || newRow[fromDateKey];
                let rawDateNum = newRow[fromDateKey];
                let parsedDate = null;

                if (rawDateStr) {
                    let str = String(rawDateStr).trim();
                    // Aggressive match for DD-MM-YYYY even if there's trailing time/spaces
                    const match = str.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})/);
                    if (match) {
                        let d = parseInt(match[1], 10);
                        let m = parseInt(match[2], 10) - 1;
                        let y = parseInt(match[3], 10);
                        if (y < 100) y += 2000;
                        parsedDate = new Date(y, m, d);
                    } else if (typeof rawDateNum === 'number') {
                        // Failsafe for unformatted serial numbers
                        const utc_days  = Math.floor(rawDateNum - 25569);
                        const utc_value = utc_days * 86400;
                        const date_info = new Date(utc_value * 1000);
                        let m = date_info.getUTCMonth();
                        let d = date_info.getUTCDate();
                        let y = date_info.getUTCFullYear();
                        
                        // INDIAN LOCALE FIX: If Excel stored it as Jan 6th (m=0, d=6) instead of June 1st, auto-swap it.
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
                    
                    newRow[fromDateKey] = format(nextMonthDate, 'dd-MM-yyyy');
                    newRow[toDateKey] = format(endOfNextMonth, 'dd-MM-yyyy');
                    
                    if (monthKey || newRow['Month'] !== undefined) {
                        const actualMonthKey = monthKey || 'Month';
                        newRow[actualMonthKey] = format(nextMonthDate, 'MMMM');
                    }
                }
            }

            // Step 3: Static values
            const customerGroupKey = findColumnKey(newRow, ['Customer Group']);
            const gstPercentKey = findColumnKey(newRow, ['GST %']);
            newRow[customerGroupKey] = 'GT';
            newRow[gstPercentKey] = 0.05;

            // Step 4: First Item lookup (by Item Name)
            const itemName = newRow[nlcItemNameKey];
            const itemMatch = itemByName[itemName];
            if (itemMatch) {
                const itemCodeTarget = findColumnKey(newRow, ['Item Code']);
                const itemGroupTarget = findColumnKey(newRow, ['Item Group']);
                const uomTarget = findColumnKey(newRow, ['UOM']);
                
                newRow[itemCodeTarget] = newRow[itemCodeTarget] || itemMatch[findColumnKey(itemMatch, ['Item Code'])];
                newRow[itemGroupTarget] = newRow[itemGroupTarget] || itemMatch[findColumnKey(itemMatch, ['Item Group'])];
                newRow[uomTarget] = newRow[uomTarget] || itemMatch[findColumnKey(itemMatch, ['Conversion Factor (UOM Conversion Detail)'])];
            }

            // Step 5: GT lookup (by SKU = Item Name)
            const gtMatch = gtBySku[itemName];
            
            const exFactoryTarget = findColumnKey(newRow, ['Ex-Factory Cost Per Kg']);
            const logisticsTarget = findColumnKey(newRow, ['Logistics Cost']);
            const marginPctTarget = findColumnKey(newRow, ['Margin Percentage']);
            const nonGstFinalTarget = findColumnKey(newRow, ['Non-GST Final (Ex-Factory + Capital cost + Logistics + Margin) (Per kg)', 'Non-GST Final \r\n(Ex-Factory + Capital cost\r\n+ Logistics + Margin) \r\n(Per kg)', 'Non-GST Final']);
            const grandFinalTarget = findColumnKey(newRow, ['Grand Final (Ex-Factory + Capital cost + Logistics + Margin + GST) (Per kg)', 'Grand Final\r\n(Ex-Factory + Capital cost +\r\nLogistics + Margin +\r\nGST) \r\n(Per kg)', 'Grand Final']);
            const costBasisTarget = findColumnKey(newRow, ['Ex-Factory+ Capital Cost+ Logistics+ GST (per kg) For Margin calculation only', 'Ex-Factory+ Capital Cost+ Logistics+ GST \r\n(per kg) \r\nFor Margin calculation only']);
            
            if (gtMatch) {
                const gtExFactory = parseFloat(gtMatch[findColumnKey(gtMatch, ['Ex-Factory'])]) || 0;
                const gtLogistics = parseFloat(gtMatch[findColumnKey(gtMatch, ['Logistics Cost'])]) || 0;
                const gtTotalCost = parseFloat(gtMatch[findColumnKey(gtMatch, ['Total Cost (per KG)'])]) || 0;
                const gtNlcSale = parseFloat(gtMatch[findColumnKey(gtMatch, ['NLC PER KG (SALE BASIS)'])]) || 0;
                const gtNlcCost = parseFloat(gtMatch[findColumnKey(gtMatch, ['NLC PER KG (COST BASIS)'])]) || 0;
                
                let gtMarginRaw = gtMatch[findColumnKey(gtMatch, ['Margin'])];
                let gtMarginPct = 0;
                if (typeof gtMarginRaw === 'string' && gtMarginRaw.includes('%')) {
                    gtMarginPct = parseFloat(gtMarginRaw.replace('%', '')) / 100;
                } else if (gtMarginRaw) {
                    gtMarginPct = parseFloat(gtMarginRaw);
                    if (gtMarginPct > 1) gtMarginPct = gtMarginPct / 100;
                }

                newRow[exFactoryTarget] = gtExFactory * 1.05;
                newRow[logisticsTarget] = gtLogistics * 1.05;
                newRow[marginPctTarget] = gtMarginPct;
                newRow[nonGstFinalTarget] = gtTotalCost * 1.05;
                newRow[grandFinalTarget] = gtNlcSale * 1.05;
                newRow[costBasisTarget] = gtNlcCost * 1.05;
            }

            // Step 6 & 7: Calculations
            const marginTarget = findColumnKey(newRow, ['Margin']);
            const gstAmountTarget = findColumnKey(newRow, ['GST Amount']);
            
            const marginPct = parseFloat(newRow[marginPctTarget]) || 0;
            const costBasis = parseFloat(newRow[costBasisTarget]) || 0;
            const nonGstFinal = parseFloat(newRow[nonGstFinalTarget]) || 0;

            newRow[marginTarget] = marginPct * costBasis;
            newRow[gstAmountTarget] = 0.05 * nonGstFinal;

            // Step 8: Second Item lookup (by Item Code)
            const itemCode = newRow[findColumnKey(newRow, ['Item Code'])];
            const itemByCodeMatch = itemByCode[itemCode];
            
            if (itemByCodeMatch) {
                const mrpTarget = findColumnKey(newRow, ['MRP']);
                const eanTarget = findColumnKey(newRow, ['EAN Code']);
                
                newRow[mrpTarget] = itemByCodeMatch[findColumnKey(itemByCodeMatch, ['Mrp'])] || newRow[mrpTarget];
                newRow[eanTarget] = itemByCodeMatch[findColumnKey(itemByCodeMatch, ['Barcode (Item Barcode)'])] || newRow[eanTarget];
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
