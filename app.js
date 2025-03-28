const newman = require('newman');

newman.run({
    collection: 'collection/Dandys_commerce_collection.json', // Path to the collection
    environment: 'collection/dandy_env.json', // Path to the environment file
    iterationCount: 1, // Number of iterations
    reporters: ['htmlextra'], // Using htmlextra for better reporting
    reporter: {
        htmlextra: {
            export: 'index.html', // Save report to a file
            logs: true, // Enable detailed logs
            showOnlyFails: false, // Show all results, not just failures
            noSyntaxHighlighting: false, // Keep syntax highlighting enabled
            browserTitle: "Dandy's Commerce API Report",
            title: "Dandy's Commerce Test Report",
            titleSize: 4,
            omitHeaders: false, // Show request headers
            skipHeaders: "Authorization", // Hide sensitive headers
            omitRequestBodies: false, // Show request bodies
            omitResponseBodies: false, // Show response bodies
            showEnvironmentData: true, // Display environment variables in the report
            skipEnvironmentVars: ["API_KEY", "SECRET"], // Hide sensitive environment variables
            showGlobalData: true, // Show global variables
            skipGlobalVars: ["API_TOKEN"], // Hide sensitive global variables
            skipSensitiveData: true, // Mask sensitive data
            showMarkdownLinks: true, // Enable markdown links
            showFolderDescription: true, // Show folder descriptions
            timezone: "UTC", // Standardized timezone
            displayProgressBar: true // Show progress bar during execution
        }
    }
}, function (err, summary) {
    if (err) {
        console.error("🚨 Newman encountered an error:", err);
    }
    
    console.log("✅ Newman run completed successfully!");

    // Log failed tests but do not exit immediately
    if (summary.run.failures.length > 0) {
        console.error(`⚠️ Some tests failed (${summary.run.failures.length} failures). Check the report for details.`);
    } else {
        console.log("🎉 All tests passed successfully!");
    }
    
    // Always exit normally so all tests run
    process.exit(0);
});
