# GAS-helper-functions

A Google Apps Script library for frequently used simple tasks. Can be used with minimal understanding of Javascript.

## Table of content

## Installation

## Features

## Usage

### Methods

| Method | Return type | Brief description |
| --- | --- | --- |
| **archive** (sourceSheet, targetSheet, condition, headerRow, spreadsheetId) | <code>**void**</code> | copy rows that match condition from sourceSheet to targetSheet and delete these rows from sourceSheet
| **compare** (firstDataObj, secondDataObj, comparisonObj) | <code>**object**</code> | compares firstDataObj to secondDataObj based on property names and returns an object that contains 
| **importCsv** (searchString, searchPlace, fileName, parseConfig, deleteBool, safeMode, encoding, manipulationFunc) | <code>**Array[Array]**</code> | [**! requires PapaParse !**](https://github.com/mholt/PapaParse) reads and parses a CSV file from either Gmail Inbox or Drive and returns a 2-dimensional array |
| **addDataToSheet** (data, type, sheetName, spreadsheetId, startColumn, removeHeaderBool, headerRow, countRowsByColumn, useAdvancedServices) | <code>**void**</code> | writes data from a 2-dimensional array to a sheet either replacing existing data or adding to it / can utilize the Advanced Sheets service for large datasets |
| **sendMailOnEdit** (eventObj, sheetName, conditionHeader, condition, emailTextObj, emailSubjectObj, recipientHeader, emailOptionsObj, headerRow) | <code>**void**</code> | sends an E-mail if the edit meets the condition |
| **moveValues** (sourceSheet, targetSheet, condition, headerRows, startColumn, filterArray , sourceSpreadsheet, targetSpreadsheet)  | <code>**void**</code> | copy rows that match condition from sourceSheet to targetSheet / similar to archive but doesn't delete matching rows and can be used with  different Spreadsheets