const XLSX = require('xlsx');

// Input and output file paths
const inputFile = 'HCI Merged Records Cleanup Workbook.xlsx';
const outputFile = 'Updated_HCI_Merged_Records_v3.xlsx';

function processWorkbook() {
    try {
        console.log('Reading workbook...');
        
        // Read the workbook with full formatting preservation
        const workbook = XLSX.readFile(inputFile, {
            cellStyles: true,
            cellFormula: true,
            cellNF: true,
            cellText: true,
            raw: true
        });

        // Get both sheets
        const tab1 = workbook.Sheets['EE Portal Merged Record IDs'];
        const tab2 = workbook.Sheets['ISSA Portal'];

        // Convert sheets to JSON for easier processing
        const tab1Data = XLSX.utils.sheet_to_json(tab1, {header: 1, raw: true});
        const tab2Data = XLSX.utils.sheet_to_json(tab2, {header: 1, raw: true});

        console.log('Processing records...');

        // Store matches first
        const matches = [];
        const nonMatches = [];
        
        // Keep the header row
        const headerRow = tab2Data[0];

        // Process each row in Tab 2
        for (let i = 1; i < tab2Data.length; i++) {
            const row = tab2Data[i];
            if (!row) {
                nonMatches.push(row);
                continue;
            }

            const eeContactRecordId = row[2]; // Column C
            if (!eeContactRecordId) {
                nonMatches.push(row);
                continue;
            }

            // Search for matching ID in Tab 1
            let matchedMergedId = null;
            for (let j = 1; j < tab1Data.length; j++) {
                const tab1Row = tab1Data[j];
                if (!tab1Row) continue;

                // Search through columns C through M
                for (let k = 2; k <= 12; k++) { // Column C (2) through Column13 (12)
                    if (tab1Row[k] === eeContactRecordId) {
                        matchedMergedId = tab1Row[2]; // Get Merged Contact ID from Column C
                        break;
                    }
                }
                if (matchedMergedId) break;
            }

            // If match found, add to matches array, otherwise add to nonMatches
            if (matchedMergedId) {
                const rowCopy = [...row];
                rowCopy[3] = matchedMergedId; // Set Column D
                matches.push(rowCopy);
            } else {
                nonMatches.push(row);
            }
        }

        // Combine arrays: header + matches + non-matches
        const newData = [headerRow, ...matches, ...nonMatches];

        // Convert back to sheet
        const newTab2 = XLSX.utils.aoa_to_sheet(newData);

        // Copy the original sheet's properties and styling
        newTab2['!cols'] = tab2['!cols'];
        newTab2['!rows'] = tab2['!rows'];
        newTab2['!merges'] = tab2['!merges'];

        // Ensure the range includes Column D
        const newRange = XLSX.utils.decode_range(newTab2['!ref']);
        if (newRange.e.c < 3) newRange.e.c = 3;
        newTab2['!ref'] = XLSX.utils.encode_range(newRange);

        // Update the workbook
        workbook.Sheets['ISSA Portal'] = newTab2;

        console.log('Writing updated workbook...');
        
        // Write the workbook with all formatting preserved
        XLSX.writeFile(workbook, outputFile, {
            bookSST: true,
            cellStyles: true,
            compression: true
        });

        console.log('\nProcessing Summary:');
        console.log('------------------');
        console.log(`Total matches found and written: ${matches.length}`);
        console.log(`Matches consolidated at rows 2-${matches.length + 1}`);
        console.log(`Process completed successfully!`);
        console.log(`Output file created: ${outputFile}`);

    } catch (error) {
        console.error('Error processing workbook:', error);
    }
}

// Execute the process
processWorkbook(); 