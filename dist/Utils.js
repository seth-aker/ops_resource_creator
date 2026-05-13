"use strict";
const DEFAULT_BATCH_SIZE = 100;
const MAX_RETRIES = 5;
function getSpreadSheetData(spreadsheet) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(spreadsheet);
    if (!sheet)
        throw new Error(`Could not find spreadsheet: "${spreadsheet}"`);
    const dataRange = sheet.getDataRange(); // Get data
    const data = dataRange.getValues(); // create 2D array
    // Process data (e.g., converting to JSON format for API)
    const headers = data[0];
    if (!headers)
        throw new Error(`No headers found in row`);
    const jsonData = [];
    for (let rowIndex = 1; rowIndex < data.length; rowIndex++) {
        const row = {};
        for (let colIndex = 0; colIndex < headers.length; colIndex++) {
            const rowData = data[rowIndex];
            if (!rowData)
                continue;
            let value = rowData[colIndex];
            // Trim whitespace if the value is a string
            if (typeof value === 'string') {
                value = value.trim();
            }
            row[headers[colIndex]] = value;
        }
        jsonData.push(row);
    }
    return jsonData;
}
function createHeaders(token) {
    return {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json'
    };
}
function batchFetch(batchOptions) {
    const sliceCount = Math.ceil(batchOptions.length / DEFAULT_BATCH_SIZE);
    const responses = [];
    logBatchProgress({
        failedCount: 0,
        totalItems: batchOptions.length,
        completedItems: 0,
        totalBatches: sliceCount,
        completedBatches: 0
    });
    let failedCount = 0;
    for (let i = 0; i < sliceCount; i++) {
        logEvent(`Starting batch ${i + 1} of ${sliceCount}`);
        const batchRes = _fetchAllWithRetries(batchOptions.slice(i * DEFAULT_BATCH_SIZE, (i + 1) * DEFAULT_BATCH_SIZE));
        batchRes.forEach((res) => {
            if (res.getResponseCode() > 299) {
                failedCount++;
            }
            responses.push(...batchRes);
        });
        logBatchProgress({
            failedCount,
            completedItems: responses.length - failedCount,
            completedBatches: i + 1
        });
        logEvent(`Batch ${i + 1} complete`);
    }
    return responses;
}
function _fetchAllWithRetries(batchOptions, retryCount = 0) {
    const retries = [];
    const responses = UrlFetchApp.fetchAll(batchOptions);
    const retryIdxs = [];
    responses.forEach((res, idx) => {
        const code = res.getResponseCode();
        if (code === 500) {
            retries.push(batchOptions[idx]);
            retryIdxs.push(idx);
        }
    });
    if (retryCount <= MAX_RETRIES && retries.length > 0) {
        const batchProgress = getBatchProgress();
        logEvent(`${retries.length} items timed out in batch ${batchProgress.completedBatches + 1}. Retrying...`);
        const retryResponses = _fetchAllWithRetries(retries, retryCount + 1);
        retryResponses.forEach((res, idx) => {
            responses[retryIdxs[idx]] = res;
        });
    }
    return responses;
}
function logEvent(message) {
    const userService = PropertiesService.getUserProperties();
    const raw = userService.getProperty('scriptEvents');
    const events = raw ? JSON.parse(raw) : [];
    events.push(message);
    userService.setProperty('scriptEvents', JSON.stringify(events));
}
function getBatchProgress() {
    const userService = PropertiesService.getUserProperties();
    const raw = userService.getProperty('batchProgress');
    const current = raw ? JSON.parse(raw) : {};
    return current;
}
function logBatchProgress(progress) {
    const userService = PropertiesService.getUserProperties();
    const raw = userService.getProperty('batchProgress');
    const current = raw ? JSON.parse(raw) : {};
    userService.setProperty('batchProgress', JSON.stringify({ ...current, ...progress }));
}
function getScriptProgress() {
    const userService = PropertiesService.getUserProperties();
    const properties = userService.getProperties();
    const batchProgress = properties.batchProgress ? JSON.parse(properties.batchProgress) : {};
    const scriptEvents = properties.scriptEvents ? JSON.parse(properties.scriptEvents) : [];
    const scriptFinished = properties.scriptFinished ? JSON.parse(properties.scriptFinished) : false;
    return {
        batchProgress,
        scriptEvents,
        scriptFinished
    };
}
function clearScriptProgress() {
    const userService = PropertiesService.getUserProperties();
    userService.deleteProperty('batchProgress');
    userService.deleteProperty('scriptEvents');
}
function openProgressSidebar(title) {
    const html = HtmlService.createHtmlOutputFromFile("ScriptProgressSidebar")
        .setTitle(title);
    SpreadsheetApp.getUi().showSidebar(html);
}
function highlightRows(rowIndices, color) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const lastColumn = sheet.getLastColumn();
    const rowGroups = new Map();
    let groupStart = rowIndices[0];
    let groupEnd = rowIndices[rowIndices.length - 1];
    rowGroups.set(groupStart, groupEnd);
    for (let i = 0; i < rowIndices.length - 1; i++) {
        if (rowIndices[i + 1] !== rowIndices[i] + 1) { // if there are entries between rows that did not fail
            groupEnd = rowIndices[i];
            rowGroups.set(groupStart, groupEnd);
            groupStart = rowIndices[i + 1];
        }
    }
    // set the last group
    rowGroups.set(groupStart, rowIndices[rowIndices.length - 1]);
    const groupStarts = Array.from(rowGroups.keys()).sort((a, b) => a - b);
    groupStarts.forEach(rowStart => {
        if (rowStart >= 0) {
            const groupSize = rowGroups.get(rowStart) - rowStart + 1;
            sheet.getRange(rowStart, 1, groupSize, lastColumn).setBackground(color);
        }
    });
}
function setIsScriptFinished(isFinished) {
    PropertiesService.getUserProperties().setProperty('scriptFinished', JSON.stringify(isFinished));
}
