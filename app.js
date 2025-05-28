document.addEventListener('DOMContentLoaded', () => {
    const jsonFileInput = document.getElementById('jsonFile');
    const xlsxFileInput = document.getElementById('xlsxFile');
    const customTestEmailsInput = document.getElementById('customTestEmails'); // New input
    const processButton = document.getElementById('processButton');
    const jsonResultP = document.getElementById('jsonResult');
    const xlsxResultP = document.getElementById('xlsxResult');
    const statusP = document.getElementById('status');
    const resultsContainer = document.getElementById('resultsContainer');

    // Base list of test emails (normalized to lowercase)
    const BASE_TEST_EMAILS_CONFIG = [
        "ching.test.email@gmail.com"
    ].map(email => email.toLowerCase());

    processButton.addEventListener('click', async () => {
        processButton.disabled = true;
        statusP.textContent = 'Processing...';
        jsonResultP.textContent = '- Import: N/A';
        xlsxResultP.textContent = '- Sent: N/A';

        const jsonFile = jsonFileInput.files[0];
        const xlsxFile = xlsxFileInput.files[0];
        const customEmailsString = customTestEmailsInput.value;

        // --- Combine base test emails with user-provided ones ---
        let currentRunTestEmails = [...BASE_TEST_EMAILS_CONFIG];
        if (customEmailsString.trim() !== "") {
            const userEnteredEmails = customEmailsString
                .split(/[\s,;\n]+/) // Split by comma, semicolon, newline, or spaces
                .map(email => email.trim().toLowerCase())
                .filter(email => email !== "" && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)); // Basic email format check

            userEnteredEmails.forEach(email => {
                if (!currentRunTestEmails.includes(email)) {
                    currentRunTestEmails.push(email);
                }
            });
        }
        // --- End of test email combination ---

        if (!jsonFile && !xlsxFile) {
            statusP.textContent = 'Please select at least one file.';
            processButton.disabled = false;
            return;
        }

        let importedEmailsCount = null;
        let totalSentEmailsXLSX = 0;
        let totalTestEmailsFoundInXLSX = 0;

        if (jsonFile) {
            statusP.textContent = 'Processing JSON file...';
            try {
                const jsonData = await readFileAsText(jsonFile);
                const parsedJson = JSON.parse(jsonData);
                const count = countEmailKeysInJson(parsedJson);
                importedEmailsCount = count;
                jsonResultP.textContent = `- Import: ${count}`;
                statusP.textContent = 'JSON processing complete.';
            } catch (error) {
                console.error("Error processing JSON:", error);
                jsonResultP.textContent = '- Import: Error';
                statusP.textContent = `Error processing JSON: ${error.message}`;
                importedEmailsCount = null;
            }
        } else {
             jsonResultP.textContent = '- Import: N/A (No file selected)';
        }

        if (xlsxFile) {
            statusP.textContent = 'Processing XLSX file (this might take a moment for large files)...';
            try {
                const arrayBuffer = await readFileAsArrayBuffer(xlsxFile);
                const workbook = XLSX.read(arrayBuffer, { type: 'array' });
                const firstSheetName = workbook.SheetNames[0];
                if (!firstSheetName) throw new Error("XLSX file contains no sheets.");
                const worksheet = workbook.Sheets[firstSheetName];
                const data = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });

                if (data.length === 0) {
                    xlsxResultP.textContent = '- Sent: 0 (Sheet is empty)';
                    totalSentEmailsXLSX = 0;
                } else {
                    const headerRow = data[0];
                    let emailColumnIndex = -1;
                    for (let i = 0; i < headerRow.length; i++) {
                        if (String(headerRow[i]).trim().toLowerCase() === 'email') {
                            emailColumnIndex = i;
                            break;
                        }
                    }

                    if (emailColumnIndex === -1) {
                        xlsxResultP.textContent = '- Sent: 0 ("email" column not found)';
                        totalSentEmailsXLSX = 0;
                    } else {
                        let currentXlsxEmailCount = 0;
                        let testEmailOccurrences = {};
                        // Initialize occurrences for all test emails in this run
                        currentRunTestEmails.forEach(email => testEmailOccurrences[email] = 0);

                        for (let i = 1; i < data.length; i++) {
                            const row = data[i];
                            if (row && row[emailColumnIndex] !== undefined) {
                                const cellValue = String(row[emailColumnIndex]).trim();
                                if (cellValue !== "") {
                                    currentXlsxEmailCount++;
                                    const lowerCellValue = cellValue.toLowerCase();
                                    // Check against the combined list for this run
                                    if (currentRunTestEmails.includes(lowerCellValue)) {
                                        testEmailOccurrences[lowerCellValue]++;
                                    }
                                }
                            }
                        }
                        totalSentEmailsXLSX = currentXlsxEmailCount;

                        let testEmailsDetailsParts = [];
                        totalTestEmailsFoundInXLSX = 0;
                        // Iterate through all potential test emails for this run for the "Except for" string
                        currentRunTestEmails.forEach(testEmail => {
                            if (testEmailOccurrences[testEmail] > 0) {
                                testEmailsDetailsParts.push(`${testEmail} (${testEmailOccurrences[testEmail]})`);
                                totalTestEmailsFoundInXLSX += testEmailOccurrences[testEmail];
                            }
                        });

                        let xlsxDisplayText = `- Sent: ${totalSentEmailsXLSX}`;
                        if (testEmailsDetailsParts.length > 0) {
                            xlsxDisplayText += ` (Except for ${testEmailsDetailsParts.join(', ')})`;
                        }
                        xlsxResultP.textContent = xlsxDisplayText;
                    }
                }
                statusP.textContent = 'XLSX processing complete.';
            } catch (error) {
                console.error("Error processing XLSX:", error);
                xlsxResultP.textContent = '- Sent: Error';
                statusP.textContent = `Error processing XLSX: ${error.message}`;
            }
        } else {
            xlsxResultP.textContent = '- Sent: N/A (No file selected)';
        }

        if (xlsxFile && !xlsxResultP.textContent.includes('Error') && !xlsxResultP.textContent.includes('N/A')) {
            let completionStatusText;
            if (importedEmailsCount === null) {
                completionStatusText = "incompleted (JSON data unavailable)";
            } else {
                const effectiveSentCount = totalSentEmailsXLSX - totalTestEmailsFoundInXLSX;
                if (importedEmailsCount === effectiveSentCount) {
                    completionStatusText = "completed";
                } else {
                    completionStatusText = "incompleted";
                }
            }
            xlsxResultP.textContent += ` ---> ${completionStatusText}`;
        }

        if(jsonFile && xlsxFile && !statusP.textContent.toLowerCase().includes('error')) {
            statusP.textContent = 'All processing complete.';
        } else if (jsonFile && !statusP.textContent.toLowerCase().includes('error')) {
            // status already set
        } else if (xlsxFile && !statusP.textContent.toLowerCase().includes('error')) {
            // status already set
        }
        
        processButton.disabled = false;
    });

    function readFileAsText(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = event => resolve(event.target.result);
            reader.onerror = error => reject(error);
            reader.readAsText(file);
        });
    }

    function readFileAsArrayBuffer(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = event => resolve(event.target.result);
            reader.onerror = error => reject(error);
            reader.readAsArrayBuffer(file);
        });
    }

    function countEmailKeysInJson(data) {
        let count = 0;
        const toVisit = [data];
        while (toVisit.length > 0) {
            const current = toVisit.pop();
            if (typeof current === 'object' && current !== null) {
                if (Array.isArray(current)) {
                    for (let i = current.length - 1; i >= 0; i--) toVisit.push(current[i]);
                } else {
                    for (const key in current) {
                        if (current.hasOwnProperty(key)) {
                            if (key === 'email') count++;
                            if (typeof current[key] === 'object' && current[key] !== null) {
                                toVisit.push(current[key]);
                            }
                        }
                    }
                }
            }
        }
        return count;
    }

    function copyResultsToClipboard() {
        const jsonText = jsonResultP.textContent;
        const xlsxText = xlsxResultP.textContent;

        if ((jsonText.includes(': N/A') || jsonText.includes(': Error')) &&
            (xlsxText.includes(': N/A') || xlsxText.includes(': Error'))) {
            statusP.textContent = "Nothing valid to copy.";
            return;
        }
        const combinedText = `${jsonText}\n${xlsxText}`;
        navigator.clipboard.writeText(combinedText).then(() => {
            const originalStatus = statusP.textContent;
            resultsContainer.style.backgroundColor = '#d4edda';
            statusP.textContent = `Copied all results!`;
            setTimeout(() => {
                statusP.textContent = originalStatus;
                resultsContainer.style.backgroundColor = '';
            }, 2500);
        }).catch(err => {
            console.error('Failed to copy with navigator.clipboard: ', err);
            try {
                const textArea = document.createElement("textarea");
                textArea.value = combinedText;
                textArea.style.position = "fixed"; textArea.style.opacity = "0";
                document.body.appendChild(textArea);
                textArea.focus(); textArea.select();
                document.execCommand('copy');
                document.body.removeChild(textArea);
                const originalStatus = statusP.textContent;
                resultsContainer.style.backgroundColor = '#d4edda';
                statusP.textContent = `Copied all results (fallback)!`;
                setTimeout(() => {
                    statusP.textContent = originalStatus;
                    resultsContainer.style.backgroundColor = '';
                }, 2500);
            } catch (execErr) {
                console.error('Fallback copy method failed: ', execErr);
                statusP.textContent = 'Failed to copy text.';
            }
        });
    }
    resultsContainer.addEventListener('click', copyResultsToClipboard);
});
