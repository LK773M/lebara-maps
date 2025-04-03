const azureKey = "1M6mP4Z6ka1nry2LjLZHv3DZWJTb7xeJtzggRuOkB5RC6osufI3ZJQQJ99BCAC5RqLJOJqSjAAAgAZMP2H79";

// Cache for storing validated addresses
const addressCache = new Map();
const BATCH_SIZE = 10;
const RATE_LIMIT_DELAY = 1000; // 1 second delay between batches

// Global map variables
let map = null;
let datasource = null;
let popup = null;

// Fuzzy matching function (Levenshtein distance for more accuracy)
function levenshteinDistance(a, b) {
    if (a.length === 0) return b.length;
    if (b.length === 0) return a.length;
    
    const matrix = [];

    // increment along the first column of each row
    for (let i = 0; i <= b.length; i++) {
        matrix[i] = [i];
    }

    // increment each column in the first row
    for (let j = 0; j <= a.length; j++) {
        matrix[0][j] = j;
    }

    // Fill in the rest of the matrix
    for (let i = 1; i <= b.length; i++) {
        for (let j = 1; j <= a.length; j++) {
            if (b.charAt(i - 1) == a.charAt(j - 1)) {
                matrix[i][j] = matrix[i - 1][j - 1];
            } else {
                matrix[i][j] = Math.min(
                    matrix[i - 1][j - 1] + 1, // substitution
                    matrix[i][j - 1] + 1,     // insertion
                    matrix[i - 1][j] + 1      // deletion
                );
            }
        }
    }

    return matrix[b.length][a.length];
}

function fuzzyMatch(str1, str2, threshold = 0.7) {
    if (!str1 || !str2) return 0;
    str1 = str1.toLowerCase().trim();
    str2 = str2.toLowerCase().trim();
    if (str1 === str2) return 1;
    
    const distance = levenshteinDistance(str1, str2);
    const maxLength = Math.max(str1.length, str2.length);
    const score = 1 - distance / maxLength;
    
    return score >= threshold ? score : 0;
}

// Enhanced Address Parsing - More robust regex and flag setting
function parseInputAddress(streetNumName, postCode, town) {
    // Ensure inputs are strings
    const postCodeStr = String(postCode || '');
    const townStr = String(town || '');
    const streetNumNameStr = String(streetNumName || '');
    
    const components = {
        rawInput: `${streetNumNameStr}, ${postCodeStr} ${townStr}`.trim().replace(/\s+/g, ' '), // Normalize spaces
        streetNumber: '',
        streetName: '',
        postalCode: postCodeStr.trim() || '',
        city: townStr.trim() || '',
        country: 'France',
        extraInfo: '',
        hasNumberRange: false // Initialize flag
    };

    if (streetNumNameStr) {
        let currentStreetInput = streetNumNameStr.trim();
        
        // 1. Extract common prefixes (e.g., "Chez ...")
        const extraMatch = currentStreetInput.match(/^(Chez\s.+?)(?=\s+\d+\s|\s+$|$)/i);
        if (extraMatch) {
            components.extraInfo = extraMatch[1].trim();
            currentStreetInput = currentStreetInput.substring(extraMatch[1].length).trim();
        }
        
        // 2. Enhanced Regex for Number/Range Extraction
        // Handles: "123", "123 bis", "123A", "19-21", "19 / 21", "19 à 21"
        // Requires a space after the number part before the street name.
        const streetMatch = currentStreetInput.match(
            // Group 1: Number part (range, number+suffix, or number)
            /^(\d+\s*[-–—\/à]\s*\d+\b|\d+\s*(?:bis|ter|quater|A|B|C|D)\b|\d+)\s+(.+)$/i
            // Group 2: Street name part
        );

        if (streetMatch && streetMatch[1] && streetMatch[2]) {
            components.streetNumber = streetMatch[1].trim().replace(/\s+/g, ' '); // Normalize spaces
            components.streetName = streetMatch[2].trim().replace(/\s+/g, ' '); // Normalize spaces
            
            // Check if the extracted number is specifically a range
            components.hasNumberRange = /^\d+\s*[-–—\/à]\s*\d+\b/i.test(components.streetNumber);
            
        } else {
             // If no number pattern matched at the start, assume the whole input is the street name
             components.streetName = currentStreetInput.replace(/\s+/g, ' ').trim();
             components.hasNumberRange = false;
        }
    }

    // Basic validation: Ensure core components needed for a meaningful query exist
    components.hasRequiredFields = !!(components.streetName && components.postalCode && components.city);

    return components;
}

document.getElementById("fileInput").addEventListener("change", async (e) => {
  const file = e.target.files[0];
    const fileNameDisplay = document.getElementById("fileNameDisplay"); // Get the span
    if (!file) {
        if (fileNameDisplay) fileNameDisplay.textContent = "No file chosen"; // Reset span if no file
        return;
    }

    // --- ADD FILE TYPE CHECK ---
    if (!file.name.toLowerCase().endsWith('.xlsx')) {
        alert("Invalid file type. Please upload an .xlsx file.");
        e.target.value = null; // Clear the selected file
        if (fileNameDisplay) fileNameDisplay.textContent = "No file chosen"; // Reset display
        return; // Stop processing
    }
    // --- END FILE TYPE CHECK ---
    
    // Update the span with the selected file name
    if (fileNameDisplay) fileNameDisplay.textContent = file.name;

    // *** Clear the cache on new file upload ***
    addressCache.clear();
    console.log("Address cache cleared.");
    // *** End Cache Clear ***

    const loadingDiv = document.getElementById("loading");
    const resultsDiv = document.getElementById("results");
    const dashboardContainer = document.getElementById("dashboardContainer");
    
    // Reset UI state
    resultsDiv.innerHTML = ''; 
    dashboardContainer.style.display = 'none';
    loadingDiv.style.display = 'block'; // Show loading indicator

    try {
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, { type: "array" });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet);
        
        if (!rows || rows.length === 0) {
            loadingDiv.style.display = 'none';
            resultsDiv.innerHTML = '<p>No data found in the Excel file. Please check the file and try again.</p>';
            return;
        }
        
        resultsDiv.innerHTML = '<div class="progress">Starting address validation...</div>'; 

  const updatedRows = [];
        const batches = [];
        
        // Split rows into batches
        for (let i = 0; i < rows.length; i += BATCH_SIZE) {
            batches.push(rows.slice(i, i + BATCH_SIZE));
        }
        
        let processedCount = 0;
        // Process batches with rate limiting
        for (const batch of batches) {
            const batchPromises = batch.map(async (row) => {
                const streetInput = row["Street Number and Street Name"] || "";
                const postalInput = row["Post Code"] || "";
                const cityInput = row["Town"] || "";
                
                // Keep original names
                const firstName = row["First Name"] || ""; 
                const lastName = row["Last Name"] || "";
                
                const inputComponents = parseInputAddress(streetInput, postalInput, cityInput);
                const result = await validateAddress(inputComponents);
                
                // Add names to the final result object
                return { 
                    ...row, // Include original row data
                    ...result, // Include validation results
                    "_FirstName": firstName, // Add specific key for first name
                    "_LastName": lastName   // Add specific key for last name
                };
            });
            
            const batchResults = await Promise.all(batchPromises);
            updatedRows.push(...batchResults);
            processedCount += batch.length;
            
            // Rate limiting delay
            if (batches.indexOf(batch) < batches.length - 1) { 
                await new Promise(resolve => setTimeout(resolve, RATE_LIMIT_DELAY));
            }
        }
        
        // Display enhanced dashboard and results
  displayResults(updatedRows);
        
    } catch (error) {
        console.error("Error processing file:", error);
        loadingDiv.style.display = 'none';
        resultsDiv.innerHTML = `<p>Error processing file: ${error.message}</p>`;
    }
});

// --- Validation Function (Correct & Found / Incorrect or Not Found) ---
async function validateAddress(inputComponents) {
    // Construct query parts
    const queryParts = [
        inputComponents.streetNumber,
        inputComponents.streetName,
        inputComponents.postalCode,
        inputComponents.city,
        'France'
    ];
    const query = queryParts.filter(part => part && String(part).trim()).join(', ').trim();
    const cacheKey = `correct_${query.toLowerCase()}`; // Use different cache prefix

    // 1. Check Cache
    if (addressCache.has(cacheKey)) {
        return addressCache.get(cacheKey);
    }

    // 2. Initial Input Checks
    if (!inputComponents.hasRequiredFields) {
        return cacheAndReturn(cacheKey, createCorrectnessResult("Incorrect / Not Found", "Missing required fields", inputComponents));
    }
    if (!query) {
        return cacheAndReturn(cacheKey, createCorrectnessResult("Incorrect / Not Found", "Empty search query", inputComponents));
    }

    // 3. Call Azure Maps API
    const url = `https://atlas.microsoft.com/search/address/json?api-version=1.0&subscription-key=${azureKey}&query=${encodeURIComponent(query)}&countrySet=FR&limit=1`;

  try {
    const res = await fetch(url);
        if (!res.ok) {
            console.error(`Azure Maps HTTP Error: ${res.status} ${res.statusText} for query: ${query}`);
            return createCorrectnessResult("Incorrect / Not Found", `Azure Maps API Error: ${res.status}`, inputComponents);
        }
    const data = await res.json();

        // 4. Handle No Results Found
    if (!data.results || data.results.length === 0) {
            return cacheAndReturn(cacheKey, createCorrectnessResult("Incorrect / Not Found", "Not found in Azure Maps", inputComponents));
    }

        // 5. Found a Result - Now Compare Components
    const match = data.results[0];
    const matched = match.address || {};
        const matchedComponents = {
            number: String(matched.streetNumber || '').trim(),
            street: String(matched.streetName || '').trim(),
            postal: String(matched.postalCode || '').trim(),
            city: String(matched.municipality || '').trim(),
            freeform: String(matched.freeformAddress || '').trim()
        };

        const comparison = compareAddressComponents(inputComponents, matchedComponents);

        // 6. Determine Final Status: Correct & Found vs. Incorrect / Not Found
        let finalStatus = "";
        let finalDetails = "";

        // Conditions for "Correct & Found"
        if (comparison.isPostalMatch && 
            comparison.isCityMatch && 
            comparison.isStreetMatch && 
            comparison.isNumberMatch && 
            !comparison.isNumberMismatchSignificant) 
        {
            finalStatus = "Correct & Found";
            finalDetails = "Input address confirmed by Azure Maps.";
             // Add specific detail if number was missing from input
            if (comparison.details.some(d => d.includes("Input number missing"))) {
                 finalDetails += ` (${comparison.details.find(d => d.includes("Input number missing"))})`;
            }
    } else {
            finalStatus = "Incorrect / Not Found";
            // Concatenate reasons for mismatch
            if (!comparison.isPostalMatch) finalDetails += "Postal code mismatch. ";
            if (!comparison.isCityMatch) finalDetails += "City mismatch. ";
            if (!comparison.isStreetMatch) finalDetails += "Street name mismatch. ";
            if (!comparison.isNumberMatch) finalDetails += "Street number mismatch/issue. ";
            if (comparison.isNumberMismatchSignificant) finalDetails += "Unrealistic street number. ";
            if (finalDetails === "") finalDetails = "Component comparison failed, but no specific mismatch reason identified."; // Fallback
            finalDetails = finalDetails.trim();
        }
        
        return cacheAndReturn(cacheKey, createCorrectnessResult(finalStatus, finalDetails, inputComponents, matchedComponents.freeform, matchedComponents));

  } catch (err) {
        console.error("Address validation internal error:", err);
        return createCorrectnessResult("Incorrect / Not Found", `Validation Error: ${err.message}`, inputComponents);
    }
}

// Helper to create the Correctness result object
function createCorrectnessResult(status, details, input, matched = "N/A", matched_components = {}) {
    let statusIcon = status === "Correct & Found" ? "&check;" : "&cross;";
    return {
        // Input Data
        "Input Address Raw": input.rawInput,
        "Input Street Parsed": input.streetName,
        "Input Number Parsed": input.streetNumber,
        "Input Postal Code": input.postalCode,
        "Input Town": input.city,
        // Validation Outcome
        "Validation Status": `${statusIcon} ${status}`,
        "Matched Address": matched, // Still useful to show what Azure *did* find
        "Details": details,
        // Matched Components (for details view)
        "_Matched Number": matched_components.number || 'N/A',
        "_Matched Street": matched_components.street || 'N/A',
        "_Matched Postal": matched_components.postal || 'N/A',
        "_Matched City": matched_components.city || 'N/A'
    };
}

// Reintroduce the missing cacheAndReturn helper function
function cacheAndReturn(key, result) {
    // Optional: Add debugging log for specific addresses if needed
    // if (result["Input Address Raw"] && result["Input Address Raw"].includes("...")) {
    //     console.log("DEBUG (cacheAndReturn - Correctness): Final result for ...:", JSON.stringify(result, null, 2));
    // }
    addressCache.set(key, result);
    return result;
}

// New renderer for compact stats in the list header
function renderCompactHeaderStats(stats) {
    const container = document.getElementById('listHeaderStatsContainer');
    if (!container) {
        console.error("List header stats container not found!");
        return;
    }

    // Use check and cross HTML entities
    const check = '&check;'; 
    const cross = '&cross;';

    container.innerHTML = `
        <span class="header-stat total-stat">Total: <strong>${stats.total}</strong></span>
        <span class="header-stat correct-stat">${check} Correct: <strong>${stats.correct}</strong></span>
        <span class="header-stat incorrect-stat">${cross} Incorrect: <strong>${stats.incorrectOrNotFound}</strong></span>
    `;
}

// --- OLD STATS RENDERER - NO LONGER NEEDED ---
/*
function renderCorrectnessSummaryStats(stats) {
    const container = document.getElementById('summaryStats');
    if (!container) return;

    container.innerHTML = `
        <div class="stat-box total-stat"><h3>Total Processed</h3><p>${stats.total}</p></div>
        <div class="stat-box valid-stat"><h3>&#10003; Correct & Found</h3><p>${stats.correct}</p></div>
        <div class="stat-box invalid-stat"><h3>&#10007; Incorrect / Not Found</h3><p>${stats.incorrectOrNotFound}</p></div>
    `;
}
*/

// --- Display Results Function ---
function displayResults(data) {
    const resultsContainer = document.getElementById("results");
    const loadingDiv = document.getElementById("loading");
    const downloadButton = document.getElementById("downloadButton");
    const dashboardContainer = document.getElementById("dashboardContainer");
    const uploadContainer = document.querySelector(".upload-container");
    const newFileButton = document.getElementById("newFileBtn");
    
    loadingDiv.style.display = 'none';
    if (uploadContainer) uploadContainer.style.display = 'none';
    if (dashboardContainer) dashboardContainer.style.display = 'block';
    if (newFileButton) newFileButton.style.display = 'inline-block';
    if (downloadButton) downloadButton.style.display = 'inline-block';
    
    if (!data || data.length === 0) {
        resultsContainer.innerHTML = "<p>No data to display.</p>";
        if (downloadButton) downloadButton.style.display = 'none';
        return;
    }

    // Calculate Correctness stats
    const stats = calculateStats(data);
    
    // Render the NEW compact header stats
    renderCompactHeaderStats(stats); 

    // Call other render functions as needed (e.g., results table via filter)
    filterResultsTable(data); // This call handles the initial render

    // Show download button
    downloadButton.onclick = () => downloadUpdatedExcel(data);
    
    // Set up filter event listeners AFTER initial render
    setupTableFiltering(data);
}

// --- Add Event Listener for New File Button ---
document.addEventListener('DOMContentLoaded', (event) => {
    const newFileButton = document.getElementById("newFileBtn");
    const fileInput = document.getElementById("fileInput");
    const uploadContainer = document.querySelector(".upload-container");
    const dashboardContainer = document.getElementById("dashboardContainer");
  const resultsDiv = document.getElementById("results");
    const downloadButton = document.getElementById("downloadButton");
    const listHeaderStatsContainer = document.getElementById('listHeaderStatsContainer');
    const filterResultCount = document.getElementById('filterResultCount');
    const fileNameDisplay = document.getElementById("fileNameDisplay"); // Get the span

    if (newFileButton) {
        newFileButton.addEventListener('click', () => {
            // Reset UI
            if (uploadContainer) uploadContainer.style.display = 'block';
            if (dashboardContainer) dashboardContainer.style.display = 'none';
            if (resultsDiv) resultsDiv.innerHTML = '';
            if (downloadButton) downloadButton.style.display = 'none';
            if (newFileButton) newFileButton.style.display = 'none';
            if (listHeaderStatsContainer) listHeaderStatsContainer.innerHTML = ''; // Clear stats header
            if (filterResultCount) filterResultCount.innerHTML = ''; // Clear filter count
            
            // Clear the file input
            if (fileInput) fileInput.value = null;
            
            // Clear the cache
            addressCache.clear();
            console.log("Address cache cleared for new file upload.");
            
            // Optional: Reset filters
            const statusFilter = document.getElementById("statusFilter");
            const tableSearch = document.getElementById("tableSearch");
            if (statusFilter) statusFilter.value = 'all';
            if (tableSearch) tableSearch.value = '';
            
            // Reset the file name display span
            if (fileNameDisplay) fileNameDisplay.textContent = "No file chosen";
        });
    }
});

// --- Dashboard Rendering (Simplified for Correctness Check) ---

function renderDashboard(data, stats) {
    renderCorrectnessSummaryStats(stats); // Use correctness stats renderer
    // renderCorrectnessStatusChart(stats); // Use correctness chart renderer - COMMENTED OUT
    // Keep Quality Meter & Advanced Metrics hidden
    const qualityMeter = document.getElementById('qualityBar')?.parentElement;
    if (qualityMeter) qualityMeter.style.display = 'none';
    const advancedMetrics = document.getElementById('advancedMetrics')?.parentElement;
    if (advancedMetrics) advancedMetrics.style.display = 'none';
}

// Render a simplified status chart (Correctness)
function renderCorrectnessStatusChart(stats) {
    const statusCtx = document.getElementById('statusChart');
    if (!statusCtx) return;
    
    if (statusCtx._chart) {
        statusCtx._chart.destroy();
    }
    
    const chart = new Chart(statusCtx, {
        type: 'pie',
        data: {
            labels: ['Correct & Found', 'Incorrect / Not Found'],
            datasets: [{
                data: [stats.correct, stats.incorrectOrNotFound],
                backgroundColor: [
                    '#27ae60', // Correct - Green
                    '#e74c3c', // Incorrect - Red
                ]
            }]
        },
        options: { 
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { position: 'right' },
                title: { display: true, text: 'Address Correctness Status' },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            const value = context.raw;
                            const percentage = ((value / stats.total) * 100).toFixed(1);
                            return `${context.label}: ${value} (${percentage}%)`;
                        }
                    }
                }
            }
         }
    });
    statusCtx._chart = chart;
}

// --- Calculate Stats (Simplified for Correctness Check) ---
function calculateStats(data) {
    let correctCount = 0;
    let incorrectCount = 0;
    
    data.forEach(row => {
        const status = row["Validation Status"] || "";
        if (status.includes("Correct & Found")) {
            correctCount++;
        } else { // Includes "Incorrect / Not Found" and any potential "Error" status
            incorrectCount++;
        }
    });
    
    return {
        total: data.length,
        correct: correctCount,
        incorrectOrNotFound: incorrectCount
    };
}

// --- Re-added Advanced Metrics Renderer ---
function renderAdvancedMetrics(data) {
    const container = document.getElementById('advancedMetrics');
    const parentContainer = container?.parentElement;
    if (!container || !parentContainer) {
        console.warn("Advanced metrics container not found, skipping render.");
        return;
    }

    if (!data || data.length === 0) {
        parentContainer.style.display = 'none'; // Hide if no data
        return;
    }

    const stats = calculateStats(data); // Reuse stats calculation
    const qualityScore = stats.total > 0 ? (stats.correct / stats.total) : 0;
    const qualityPercentage = (qualityScore * 100).toFixed(0);

    // Basic analysis of the main issue (most frequent keyword in Details)
    const issueCounts = {};
  data.forEach(row => {
        const details = row['Details'] || '';
        if (details.includes("Postal code mismatch")) issueCounts["Postal Code Mismatch"] = (issueCounts["Postal Code Mismatch"] || 0) + 1;
        else if (details.includes("Street name mismatch")) issueCounts["Street Name Mismatch"] = (issueCounts["Street Name Mismatch"] || 0) + 1;
        else if (details.includes("City mismatch")) issueCounts["City Mismatch"] = (issueCounts["City Mismatch"] || 0) + 1;
        else if (details.includes("number mismatch") || details.includes("Number outside range") || details.includes("Input number missing")) issueCounts["Number Issues"] = (issueCounts["Number Issues"] || 0) + 1;
        else if (details.includes("Not found")) issueCounts["Not Found"] = (issueCounts["Not Found"] || 0) + 1;
    });

    let mainIssue = "N/A";
    let maxCount = 0;
    for (const issue in issueCounts) {
        if (issueCounts[issue] > maxCount) {
            mainIssue = issue;
            maxCount = issueCounts[issue];
        }
    }

    // Determine quality class for styling
    let qualityClass = 'average';
    if (qualityScore >= 0.9) qualityClass = 'good';
    else if (qualityScore < 0.6) qualityClass = 'poor';

    container.innerHTML = `
        <div class="minimal-metrics">
            <div class="key-metric ${qualityClass}">
                <div class="key-metric-value">${qualityPercentage}%</div>
                <div class="key-metric-label">Overall Correctness</div>
                <div class="key-metric-counts">(${stats.correct} / ${stats.total})</div>
            </div>
            ${maxCount > 0 ? `
            <div class="key-metric issues">
                <div class="issue-name">Main Issue Identified:</div>
                <div class="key-metric-issue">
                    <div class="issue-name">${mainIssue}</div>
                    <div class="issue-value">(${maxCount} affected)</div>
                </div>
            </div>
            ` : '<div class="key-metric">No specific major issue identified.</div>'}
        </div>
    `;
    parentContainer.style.display = 'block'; // Ensure it's visible
}

// Show enhanced address details with component comparison
function showAddressDetails(row) {
    const primaryDetailsContainer = document.getElementById('addressPrimaryDetails');
    const comparisonGridContainer = document.getElementById('addressComparisonGridContainer');
    if (!primaryDetailsContainer || !comparisonGridContainer) {
        console.error('Required detail containers not found');
        return;
    }
    
    const mainArea = document.querySelector('.main-content-area');
    if (!mainArea) return;
    
    // Determine if the address was marked as correct
    const isCorrect = (row["Validation Status"] || "").includes("Correct & Found");

    // --- Generate Primary Details HTML ---
    let primaryHtml = `
        <div style="display: flex; justify-content: space-between; align-items: start; margin-bottom: 15px;">
           <div> 
                <p class="details-field"><strong>Input Address:</strong> ${row["Input Address Raw"] || "N/A"}</p>
                <p class="details-field"><strong>Status:</strong> <span class="status-${isCorrect ? 'valid' : 'invalid'}">${row["Validation Status"] || "N/A"}</span></p>
                <!-- Removed Details/Reason -->
           </div>
            <button id="backToListBtn" class="back-button" data-index="${row._originalIndex}">&larr; Back to List</button>
        </div>
    `;
    primaryDetailsContainer.innerHTML = primaryHtml;

    // --- Generate Comparison HTML ---
    let comparisonHtml = '<div class="comparison-grid">';
    const fieldsToCompare = [
        { label: "Street Number", inputKey: "Input Number Parsed", matchedKey: "_Matched Number" },
        { label: "Street Name", inputKey: "Input Street Parsed", matchedKey: "_Matched Street" },
        { label: "Postal Code", inputKey: "Input Postal Code", matchedKey: "_Matched Postal" },
        { label: "City / Town", inputKey: "Input Town", matchedKey: "_Matched City" }
    ];

    fieldsToCompare.forEach(field => {
        const inputValue = row[field.inputKey] || "(empty)";
        const matchedValue = row[field.matchedKey] || "(N/A)";
        const isMatch = String(inputValue).trim().toLowerCase() === String(matchedValue).trim().toLowerCase() || 
                      (field.label === "Street Number" && row["_Number Match"] === "Within Range") || // Consider range match
                      (field.label === "Street Number" && row["_Number Match"] === "Partial (Suffix)") || // Consider suffix match
                      (field.label === "Street Number" && row["_Number Match"] === "Input Missing, Found"); // Consider number found when input missing ok
                      
        comparisonHtml += `
            <div class="comparison-row ${isMatch ? 'match' : 'mismatch'}">
                <div class="comparison-label">${field.label}</div>
                <div class="comparison-input">${inputValue}</div>
                <div class="comparison-matched">${isMatch ? '&check;' : '&rightarrow;'} ${matchedValue}</div>
            </div>
        `;
    });
    comparisonHtml += '</div>'; // Close comparison-grid
    comparisonGridContainer.innerHTML = comparisonHtml;
    // --- End Comparison HTML ---

    // Back button listener (remains the same)
    document.getElementById('backToListBtn')?.addEventListener('click', function() {
        mainArea.classList.remove('show-details');
    });

    // Show the details panel
    mainArea.classList.add('show-details');
    
    // Initialize map (remains the same, uses Matched Address)
    initializeMapWithAddress(row); 
}

// Render the results table - For Correctness Check
function renderResultsTable(data, container) {
    if (!container) return; // Ensure container exists
    
    // *** Clear the container FIRST ***
    container.innerHTML = ''; 
    // ***
    
    const tableContainer = document.createElement("div");
    tableContainer.className = "address-cards-container";

    if (data.length === 0) {
        // If no data, add the 'no results' message directly to the container
        container.innerHTML = '<div class="no-results">No addresses match your filter criteria</div>';
        return;
    }

    data.forEach((row, index) => {
        // Store the original index on the row object for back-navigation
        row._originalIndex = index; 
        
        const statusString = row["Validation Status"] || "";
        const isCorrect = statusString.includes("Correct & Found");

        let statusClass = isCorrect ? "status-valid" : "status-invalid"; // Green for Correct, Red for Incorrect/Not Found
        let statusIcon = isCorrect ? "&check;" : "&cross;"; // USE HTML ENTITIES

        // Get customer name
        const customerName = `${row["_FirstName"] || ''} ${row["_LastName"] || ''}`.trim();

        // Create card
        const card = document.createElement("div");
        card.className = `address-card ${statusClass}`;
        card.dataset.index = index;

        card.innerHTML = `
            <div class="card-status-badge ${statusClass}">${statusIcon}</div>
            <div class="card-content">
                <div class="address-main">
                    <div class="address-text">${row["Input Address Raw"] || ""}</div>
                    ${customerName ? `<div class="customer-name">${customerName}</div>` : ''}
                    <div class="address-location">
                        <span class="postal-code">${row["Input Postal Code"] || ""}</span>
                        <span class="town">${row["Input Town"] || ""}</span>
                    </div>
                </div>

                ${isCorrect ? 
                    // For Correct: Show confirmed match
                    `<div class="address-match-preview">
                        <span class="match-icon">✓</span> ${row["Matched Address"] || "Confirmed"}
                     </div>` : 
                    // For Incorrect: Show mismatch details
                    `<div class="no-match mismatch-warning general-warning"> 
                         <span class="warning-icon">⚠</span> ${row["Details"] || "Address incorrect or not found."}
                     </div>`
                }
                
                <div class="card-actions">
                   <button class="view-details-btn">View Details</button> 
                </div>
            </div>
        `;

        // ADD click listener to the entire card
        card.addEventListener('click', () => {
            showAddressDetails(row);
        });

        tableContainer.appendChild(card);
    });

    // Append the container holding all cards to the main container
    container.appendChild(tableContainer);
}

// Initialize or update the map with the address details
function initializeMapWithAddress(row) {
    const mapContainer = document.getElementById('addressMapContainer');
    const mapControlButtons = document.querySelectorAll('.map-controls button');
    
    if (!mapContainer) return;
    
    // Check if we have valid coordinates to display
    const matchedAddress = row["Matched Address"];
    if (!matchedAddress || matchedAddress === "N/A" || matchedAddress === "Not found in Azure Maps") {
        mapContainer.innerHTML = '<div style="padding: 20px; text-align: center;">No map data available for this address.</div>';
        return;
    }
    
    // If map is already initialized, clear it
    if (map) {
        // Clear existing data
        if (datasource) {
            datasource.clear();
        }
        
        // Immediately query for coordinates (map is already initialized)
        queryAddressCoordinates(matchedAddress, row);
    } else {
        // Initialize the map
        map = new atlas.Map('addressMapContainer', {
            language: 'fr-FR',
            view: 'Auto',
            // Use the same Azure Maps key
            authOptions: {
                authType: 'subscriptionKey',
                subscriptionKey: azureKey
            },
            enableAccessibility: true,
            center: [2.3522, 48.8566], // Default to Paris
            zoom: 4 // Start zoomed out
        });
        
        // Wait until the map resources are ready.
        map.events.add('ready', function() {
            console.log("Map ready event fired");
            
            // Create a data source for the pin
            datasource = new atlas.source.DataSource();
            map.sources.add(datasource);
            
            // Create a layer for rendering the pin
            const pinLayer = new atlas.layer.SymbolLayer(datasource, null, {
                iconOptions: {
                    image: 'pin-round-darkblue',
                    anchor: 'center',
                    allowOverlap: true
                }
            });
            
            map.layers.add(pinLayer);
            
            // Create a popup but leave it closed initially
            popup = new atlas.Popup({
                pixelOffset: [0, -10],
                closeButton: false
            });
            
            // Show the buttons
            mapControlButtons.forEach(btn => btn.style.display = 'block');
            
            // Add event handlers for map controls
            document.getElementById('toggleSatelliteBtn').addEventListener('click', function() {
                const style = map.getStyle().style;
                if (style === 'road') {
                    map.setStyle({ style: 'satellite' });
                    this.textContent = 'Show Road Map';
                } else {
                    map.setStyle({ style: 'road' });
                    this.textContent = 'Show Satellite';
                }
            });
            
            document.getElementById('zoomInBtn').addEventListener('click', function() {
                map.setCamera({ 
                    zoom: map.getCamera().zoom + 1,
                    duration: 500, // Add animation duration
                    type: 'ease' // Add animation type
                });
            });
            
            document.getElementById('zoomOutBtn').addEventListener('click', function() {
                map.setCamera({ 
                    zoom: map.getCamera().zoom - 1,
                    duration: 500, // Add animation duration
                    type: 'ease' // Add animation type
                });
            });
            
            // Query for coordinates AFTER map is fully loaded
            queryAddressCoordinates(matchedAddress, row);
        });
    }
    
    // Since we've re-rendered the map container, re-attach event handlers
    mapControlButtons.forEach(btn => btn.style.display = 'block');
}

// Query for coordinates and then display them on the map
async function queryAddressCoordinates(address, row) {
    console.log("Querying coordinates for address:", address);
    
    // Use Azure Maps to geocode the address
    const url = `https://atlas.microsoft.com/search/address/json?api-version=1.0&subscription-key=${azureKey}&query=${encodeURIComponent(address)}&limit=1`;
    
    try {
        const response = await fetch(url);
        
        if (!response.ok) {
            throw new Error(`HTTP error! Status: ${response.status}`);
        }
        
        const data = await response.json();
        
        console.log("Azure Maps geocoding response:", data);
        
        if (data.results && data.results.length > 0) {
            const result = data.results[0];
            const position = result.position;
            
            console.log("Found coordinates:", position);
            
            if (map && datasource) {
                // Clear any existing data
                datasource.clear();
                
                // Add the pin to the map
                const point = new atlas.data.Point([position.lon, position.lat]);
                const feature = new atlas.data.Feature(point, {
                    title: `${row["Input Street Parsed"] || "Address"}`,
                    description: address,
                    status: row["Validation Status"] || "Unknown",
                    score: row["Overall Score"] || "N/A"
                });
                
                datasource.add(feature);
                
                // Remove any existing event handlers to prevent duplicates
                map.events.remove('click', datasource);
                
                // Add a popup to the pin
                map.events.add('click', datasource, (e) => {
                    if (e.shapes && e.shapes.length > 0) {
                        const properties = e.shapes[0].getProperties();
                        
                        // Create content for popup
                        const content = `
                            <div class="poi-bubble">
                                <div class="poi-title">${properties.title}</div>
                                <div class="poi-address">${properties.description}</div>
                                <div>Status: ${properties.status}</div>
                                <div>Score: ${properties.score}</div>
                            </div>
                        `;
                        
                        popup.setOptions({
                            content: content,
                            position: e.position
                        });
                        
                        popup.open(map);
                    }
                });
                
                // Set the map's camera to the address with animation
                map.setCamera({
                    center: [position.lon, position.lat],
                    zoom: 15,
                    type: 'fly'
                });
                
                // Simulate a click to show the popup initially
                setTimeout(() => {
                    popup.setOptions({
                        content: `
                            <div class="poi-bubble">
                                <div class="poi-title">${row["Input Street Parsed"] || "Address"}</div>
                                <div class="poi-address">${address}</div>
                                <div>Status: ${row["Validation Status"] || "Unknown"}</div>
                                <div>Score: ${row["Overall Score"] || "N/A"}</div>
                            </div>
                        `,
                        position: [position.lon, position.lat]
                    });
                    popup.open(map);
                }, 1000);
            } else {
                console.error('Map or datasource not initialized');
            }
        } else {
            // No coordinates found
            console.error('No coordinates found for address:', address);
            const mapContainer = document.getElementById('addressMapContainer');
            if (mapContainer) {
                mapContainer.innerHTML = `
                    <div style="padding: 20px; text-align: center;">
                        <p>Could not find coordinates for: ${address}</p>
                        <p>Try a different address or check the format.</p>
                    </div>`;
            }
        }
    } catch (error) {
        console.error('Error fetching coordinates:', error);
        const mapContainer = document.getElementById('addressMapContainer');
        if (mapContainer) {
            mapContainer.innerHTML = `
                <div style="padding: 20px; text-align: center;">
                    <p>Error loading map for address: ${address}</p>
                    <p>Error details: ${error.message}</p>
                </div>`;
        }
    }
}

// Setup download button
function setupDownloadButton(data) {
    const downloadBtn = document.getElementById("downloadButton");
    if (downloadBtn) {
        // Remove previous listener if any
        downloadBtn.replaceWith(downloadBtn.cloneNode(true));
        const newDownloadBtn = document.getElementById("downloadButton"); // Get the new clone
        newDownloadBtn.onclick = () => downloadUpdatedExcel(data);
        newDownloadBtn.style.display = 'block'; // Make button visible
    }
}

function downloadUpdatedExcel(data) {
    // Keep original columns + add the new one
    const exportData = data.map(row => { 
        // 1. Start with a copy of the original row
        let exportRow = { ...row };

        // 2. Determine the value for the new column
        const isValid = (row["Validation Status"] || "").includes("Correct & Found");
        exportRow["address check"] = isValid ? "Address found" : "Address not found";

        // 3. Remove internal/intermediate fields we added during processing
        delete exportRow["Input Address Raw"];
        delete exportRow["Input Street Parsed"];
        delete exportRow["Input Number Parsed"];
        delete exportRow["Input Postal Code"];
        delete exportRow["Input Town"];
        delete exportRow["Validation Status"]; // Remove original status field
        delete exportRow["Matched Address"];
        delete exportRow["Details"];
        delete exportRow["_FirstName"];
        delete exportRow["_LastName"];
        delete exportRow["_Matched Number"];
        delete exportRow["_Matched Street"];
        delete exportRow["_Matched Postal"];
        delete exportRow["_Matched City"];
        delete exportRow["Overall Score"]; // If this was being added
        delete exportRow["_originalIndex"]; // Remove helper index
        
        // Clean up the original status string from the result object
        // if(exportRow["Validation Status"]) {
        //     exportRow["Validation Status"] = exportRow["Validation Status"].replace(/^(&.+?;\s*)/, ''); 
        // }

        return exportRow;
    });

    // Check if there is data to export
    if (!exportData || exportData.length === 0) {
        alert("No data available to download.");
        return;
    }

    // Get headers from the first object keys (will include original + new column)
    const headers = Object.keys(exportData[0]);

    const ws = XLSX.utils.json_to_sheet(exportData, { header: headers });
  const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Address Check Results");
    
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const fileName = `address_check_results_${timestamp}.xlsx`;
    XLSX.writeFile(wb, fileName);
}

async function uploadToFTP() {
  // IMPORTANT: Secure FTP upload from the browser is not recommended or straightforward.
  // This typically requires a server-side component (e.g., Node.js, Python, PHP)
  // to handle credentials securely and perform the transfer.
  // This function remains a placeholder.
  console.log("FTP Upload placeholder triggered. Requires backend integration for actual upload to specific folder.");
}

// Setup table filtering and searching
function setupTableFiltering(data) {
    const statusFilter = document.getElementById('statusFilter');
    const tableSearch = document.getElementById('tableSearch');
    
    if (statusFilter) {
        statusFilter.addEventListener('change', () => {
            filterResultsTable(data, statusFilter.value, tableSearch.value);
        });
    }
    
    if (tableSearch) {
        tableSearch.addEventListener('input', () => {
            filterResultsTable(data, statusFilter.value, tableSearch.value);
        });
    }
}

// Filter the results table based on status and search term
function filterResultsTable(data, statusFilterValue, searchTermValue) {
    const resultsDiv = document.getElementById("results");
    const countElement = document.getElementById("filterResultCount");
    const mainArea = document.querySelector('.main-content-area'); // Get main area
    
    // --- FIX: Get current filter values from DOM if not passed as arguments ---
    const currentStatusFilter = statusFilterValue !== undefined ? statusFilterValue : document.getElementById('statusFilter')?.value || 'all';
    const currentSearchTerm = searchTermValue !== undefined ? searchTermValue : document.getElementById('tableSearch')?.value || '';
    // --- End FIX ---
    
    // Hide details panel when filtering
    if (mainArea) {
        mainArea.classList.remove('show-details'); 
    }
    
    if (!resultsDiv) return;
    
    // Clear the current table
    resultsDiv.innerHTML = '';
    
    // Filter the data using currentStatusFilter and currentSearchTerm
    const filteredData = data.filter(row => {
        // Status filter
        if (currentStatusFilter && currentStatusFilter !== 'all') {
            const status = row["Validation Status"] || "";
            if (currentStatusFilter === 'correct' && !status.includes("Correct & Found")) return false;
            if (currentStatusFilter === 'incorrect' && !status.includes("Incorrect / Not Found")) return false;
        }
        
        // Search filter
        if (currentSearchTerm && currentSearchTerm.trim() !== '') {
            const term = currentSearchTerm.toLowerCase();
            const address = (row["Input Address Raw"] || "").toLowerCase();
            const postal = (row["Input Postal Code"] || "").toLowerCase();
            const town = (row["Input Town"] || "").toLowerCase();
            const matched = (row["Matched Address"] || "").toLowerCase();
            
            return address.includes(term) || 
                   postal.includes(term) || 
                   town.includes(term) || 
                   matched.includes(term);
        }
        
        return true;
    });
    
    // Update count display
    if (countElement) {
        if (filteredData.length === data.length) {
            countElement.textContent = `Showing all ${data.length} addresses`;
        } else {
            countElement.textContent = `Showing ${filteredData.length} of ${data.length} addresses`;
        }
    }
    
    // Render the filtered table
    renderResultsTable(filteredData, resultsDiv);
}

// --- Component Comparison Logic --- (Reinstated for Correct/Incorrect Check)
function compareAddressComponents(input, matched) {
    const results = {
        isPostalMatch: false,
        isCityMatch: false,
        isStreetMatch: false, // True if a reasonable street match exists
        isNumberMatch: false, // True if number is exact, range, suffix, or both missing
        isNumberMismatchSignificant: false, // True for large mismatches or unrealistic numbers
        details: [] // Store details about mismatches
    };

    // 1. Postal Code: Exact match required
    results.isPostalMatch = (input.postalCode && matched.postal && input.postalCode === matched.postal);
    if (!results.isPostalMatch && input.postalCode) {
        results.details.push(`Postal mismatch (${input.postalCode} vs ${matched.postal || 'N/A'})`);
    }

    // 2. City: High confidence fuzzy match required
    const cityScore = fuzzyMatch(input.city, matched.city, 0.85); // High threshold
    results.isCityMatch = cityScore > 0.85;
    if (!results.isCityMatch && input.city) {
         results.details.push(`City mismatch (${input.city} vs ${matched.city || 'N/A'}) - Score: ${cityScore.toFixed(2)}`);
    }

    // 3. Street Name: Handle abbreviations and reasonable fuzzy match
    const inputStreetLower = input.streetName.toLowerCase();
    const matchedStreetLower = matched.street.toLowerCase();
    let streetScore = 0;
    let isAbbreviation = false;

    if (inputStreetLower === matchedStreetLower) {
        streetScore = 1.0;
    } else {
        // Abbreviation Check (Simplified)
        const abbreviations = { /* ... same as before ... */ }; 
        const inputFirstWord = inputStreetLower.split(' ')[0];
        const matchedFirstWord = matchedStreetLower.split(' ')[0];
        let expandedInput = inputStreetLower;

        if (abbreviations[inputFirstWord] === matchedFirstWord) {
            isAbbreviation = true;
            expandedInput = inputStreetLower.replace(new RegExp(`^${inputFirstWord}\\b`), abbreviations[inputFirstWord]);
        } else if (abbreviations[matchedFirstWord] === inputFirstWord) {
            isAbbreviation = true;
        }

        if (isAbbreviation) {
            const restOfInput = expandedInput.substring(expandedInput.indexOf(' ') + 1);
            const restOfMatched = matchedStreetLower.substring(matchedStreetLower.indexOf(' ') + 1);
            if (restOfInput === restOfMatched) {
                streetScore = 0.95; 
            } else {
                streetScore = Math.max(fuzzyMatch(expandedInput, matchedStreetLower, 0.7), fuzzyMatch(inputStreetLower, matchedStreetLower, 0.7));
            }
        } else {
            streetScore = fuzzyMatch(inputStreetLower, matchedStreetLower, 0.7); // Threshold for acceptable match
        }
    }
    results.isStreetMatch = streetScore >= 0.7;
    if (!results.isStreetMatch && input.streetName) {
         results.details.push(`Street mismatch (${input.streetName} vs ${matched.street || 'N/A'}) - Score: ${streetScore.toFixed(2)}`);
    }

    // 4. Number: Check for acceptable matches or significant mismatches
    const inputNumStr = input.streetNumber;
    const matchedNumStr = matched.number;
    const inputNumPresent = !!inputNumStr;
    const matchedNumPresent = !!matchedNumStr;

    if (inputNumPresent && matchedNumPresent) {
        // Check for impossible/unrealistic number first
        const inputNumericCheck = parseInt(inputNumStr.match(/^\d+/)?.[0] || "NaN");
        if (inputNumericCheck > 9999) {
             results.isNumberMismatchSignificant = true;
             results.details.push(`Input number unrealistic (${inputNumStr})`);
        }
        // Check Range
        else if (input.hasNumberRange) {
            const rangeParts = inputNumStr.match(/^(\d+)\s*[-–—\/à]\s*(\d+)/i);
            const matchedNum = parseInt(matchedNumStr.match(/^\d+/)?.[0] || "NaN");
            if (rangeParts && !isNaN(matchedNum)) {
                const rangeStart = parseInt(rangeParts[1]);
                const rangeEnd = parseInt(rangeParts[2]);
                if (matchedNum >= rangeStart && matchedNum <= rangeEnd) {
                    results.isNumberMatch = true;
                } else {
                     results.details.push(`Number outside range (${inputNumStr} vs ${matchedNumStr})`);
                }
            } else {
                 results.details.push(`Range parse error`);
            }
        } 
        // Check Exact Match
        else if (inputNumStr === matchedNumStr) {
            results.isNumberMatch = true;
        } 
        // Check Suffix Match (e.g., 123 vs 123 bis)
        else { 
            const inputNumeric = parseInt(inputNumStr.match(/^\d+/)?.[0] || "NaN");
            const matchedNumeric = parseInt(matchedNumStr.match(/^\d+/)?.[0] || "NaN");
            if (!isNaN(inputNumeric) && inputNumeric === matchedNumeric) {
                results.isNumberMatch = true; // Consider suffix match acceptable
            } else {
                // Definite Mismatch
                 results.details.push(`Number mismatch (${inputNumStr} vs ${matchedNumStr})`);
                 if (!isNaN(inputNumeric) && !isNaN(matchedNumeric)) {
                     const diff = Math.abs(inputNumeric - matchedNumeric);
                     if (diff > 100 && diff > matchedNumeric * 2) {
                         results.isNumberMismatchSignificant = true; // Flag impossible mismatch
                         results.details.push(`Impossible number difference`);
                     }
                 }
            }
        }
    } else if (inputNumPresent && !matchedNumPresent) {
        results.details.push(`Input number provided but not found in match (${inputNumStr})`);
        const inputNumericCheck = parseInt(inputNumStr.match(/^\d+/)?.[0] || "NaN");
        if (inputNumericCheck > 9999) { // Check if input itself was unrealistic
             results.isNumberMismatchSignificant = true;
             results.details.push(`Input number unrealistic`);
        }
    } else if (!inputNumPresent && matchedNumPresent) {
        // Generally acceptable if input didn't provide one, but matched address has one
        results.isNumberMatch = true; 
        results.details.push(`Input number missing, matched has ${matchedNumStr}`); // Informative detail
    } else { // Both missing
        results.isNumberMatch = true; // Acceptable if both lack a number
    }

    return results;
}

// --- Setup Download Button --- 
function setupDownload(data) {
    const downloadButton = document.getElementById('downloadButton');
    if (!downloadButton) return;

    downloadButton.onclick = () => {
        console.log("Download button clicked. Preparing Excel file...");
        
        // 1. Define headers we want in the Excel file
        const headers = [
            "Input Address Raw",
            "Input Street Parsed",
            "Input Number Parsed",
            "Input Postal Code",
            "Input Town",
            "Validation Status",
            "Matched Address",
            "Details"
            // Add any other relevant fields from the 'row' object you want to export
            // E.g., "Overall Score" if you were using the scoring model
        ];

        // 2. Prepare data rows based on headers
        const dataToExport = data.map(row => {
            const exportRow = {};
            headers.forEach(header => {
                // Handle potential variations in status string (remove HTML icon)
                if (header === "Validation Status") {
                    exportRow[header] = (row[header] || "").replace(/^(&.+?;\s*)/, ''); 
                } else {
                    exportRow[header] = row[header] !== undefined ? row[header] : "N/A";
                }
            });
            return exportRow;
        });

        try {
            // 3. Create worksheet and workbook
            const worksheet = XLSX.utils.json_to_sheet(dataToExport, { header: headers });
            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, worksheet, "Validation Results");

            // 4. Trigger download
            XLSX.writeFile(workbook, "address_validation_results.xlsx");
            console.log("Excel file generated and download triggered.");

        } catch (error) {
            console.error("Error generating Excel file:", error);
            alert("Failed to generate Excel file. Check console for details.");
        }
    };
}
