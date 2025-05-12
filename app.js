const newman = require('newman');
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');  // Import exceljs

newman.run({
    collection: 'collection/Dandys_commerce_collection.json',
    environment: 'collection/dandy_env.json',
    iterationCount: 1,
    reporters: ['htmlextra', 'json'],
    reporter: {
        htmlextra: {
            export: 'index.html',
            logs: true,
            showOnlyFails: false,
            noSyntaxHighlighting: false,
            browserTitle: "Dandy's Commerce API Report",
            title: "Dandy's Commerce Test Report",
            titleSize: 4,
            omitHeaders: false,
            skipHeaders: "Authorization",
            omitRequestBodies: false,
            omitResponseBodies: false,
            showEnvironmentData: true,
            skipEnvironmentVars: ["API_KEY", "SECRET"],
            showGlobalData: true,
            skipGlobalVars: ["API_TOKEN"],
            skipSensitiveData: true,
            showMarkdownLinks: true,
            showFolderDescription: true,
            timezone: "UTC",
            displayProgressBar: true
        },
        json: {
            export: 'report.json'
        }
    }
}, function (err, summary) {
    if (err) {
        console.error("🚨 Newman encountered an error:", err);
        process.exit(1);
    }

    console.log("✅ Newman run completed successfully!");

    if (summary.run.failures.length > 0) {
        console.error(`⚠️ Some tests failed (${summary.run.failures.length} failures). Check the report for details.`);
    } else {
        console.log("🎉 All tests passed successfully!");
    }

    // === JSON Data to XLSX conversion starts here ===
    const jsonReportPath = path.join(__dirname, 'report.json');
    const xlsxReportPath = path.join(__dirname, 'report.xlsx');

    fs.readFile(jsonReportPath, 'utf8', (err, data) => {
        if (err) {
            console.error("❌ Failed to read JSON report for XLSX export:", err);
            return process.exit(1);
        }

        const report = JSON.parse(data);
        const executions = report.run.executions;

        const workbook = new ExcelJS.Workbook();
        const sheet = workbook.addWorksheet('DandyApp AP! Test Report')

        // Define column headers for the Excel sheet
        sheet.columns = [
            { header: 'Request Name', key: 'name', width: 40 },
            { header: 'Method', key: 'method', width: 10 },
            { header: 'URL', key: 'url', width: 10},
            { header: 'Status Code', key: 'statusCode', width: 10},
            { header: 'Test Result', key: 'testResults', width: 30 },
            { header: 'Response Time (ms)', key: 'responseTime', width: 20 },
            { header: 'Body Size (bytes)', key: 'bodySize', width: 20 },
            { header: 'Headers', key: 'headers', width: 60 },
            { header: 'Developer Comment', key: 'devComment', width: 20 },
            { header: 'QA Comment', key: 'qaComment', width: 20 }
        ];

        // Add data to the Excel sheet
        executions.forEach(exec => {
            const name = exec.item.name.replace(/,/g, '');
            const method = exec.request.method;
            const url = exec.request.url.raw;
            const statusCode = exec.response.code;
            const testResults = exec.assertions?.map(a => `${a.assertion}: ${a.error ? '❌ Failed' : '✅ Passed'}`).join(' | ') || 'No Tests';
            const responseTime = exec.response.responseTime;
            const bodySize = exec.response.responseSize;
            const headers = exec.response.headers ? 
                exec.response.headers.map(h => `${h.key}: ${h.value}`).join('; ').replace(/"/g, "'") : 
                'No Headers';


            const devComment = 'Todo';  // Default value for Developer comment
            const qaComment = 'Unresolved';  // Default value for QA comment

            // Add the row
            sheet.addRow({
                name, method, url, statusCode, testResults, responseTime, bodySize, headers, devComment, qaComment
            });
        });

        // Add dropdown validation for Developer and QA Comments
        const devCommentsRange = sheet.getColumn('devComment');
        const qaCommentsRange = sheet.getColumn('qaComment');
        
        const devComments = ['Done', 'In-progress', 'Todo'];
        const qaComments = ['Resolved', 'Unresolved', 'In-progress'];

        devCommentsRange.eachCell((cell, rowNumber) => {
            if (rowNumber > 1) {  // Skip the header row
                cell.dataValidation = {
                    type: 'list',
                    formula1: `"${devComments.join(',')}"`,
                    showErrorMessage: true
                };
            }
        });

        qaCommentsRange.eachCell((cell, rowNumber) => {
            if (rowNumber > 1) {  // Skip the header row
                cell.dataValidation = {
                    type: 'list',
                    formula1: `"${qaComments.join(',')}"`,
                    showErrorMessage: true
                };
            }
        });

        // Write the Excel file
        workbook.xlsx.writeFile(xlsxReportPath)
            .then(() => {
                console.log(`📄 XLSX report saved as ${xlsxReportPath}`);
                process.exit(0);  // Normal exit
            })
            .catch((error) => {
                console.error("❌ Error saving XLSX file:", error);
                process.exit(1);
            });
    });
});
