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
| **importCsv** (searchString, searchPlace, fileName, parseConfig, deleteBool, safeMode, encoding) | <code>**Array[Array]**</code> | [**! requires PapaParse !**](https://github.com/mholt/PapaParse) reads and parses a CSV file from either Gmail Inbox or Drive and returns a 2-dimensional array |
| **addDataToSheet** (data, sheetName, spreadsheetId, type, startColumn, removeHeaderBool, headerRow, countRowsByColumn, useAdvancedServices) | <code>**void**</code> | writes data from a 2-dimensional array to a sheet either replacing existing data or adding to it / can utilize the Advanced Sheets service for large datasets |
| **sendMailOnEdit** (eventObj, sheetName, conditionHeader, condition, emailTextObj, emailSubjectObj, recipientHeader, emailOptionsObj, headerRow) | <code>**void**</code> | sends an E-mail if the edit meets the condition |
| **moveValues** (sourceSheet, targetSheet, condition, headerRows, startColumn, filterArray , sourceSpreadsheet, targetSpreadsheet)  | <code>**void**</code> | copy rows that match condition from sourceSheet to targetSheet / similar to archive but doesn't delete matching rows and can be used with  different Spreadsheets


### Detailed Description

#### ``importCsv (searchString, searchPlace, fileName, parseConfig, deleteBool, safeMode, encoding)``

[**! requires PapaParse !**](https://github.com/mholt/PapaParse)

##### Parameters
Gets and parses a CSV-file returning the data as a 2-dimensional array. It get's the first result for the provided ``searchString`` from either the Gmail Inbox or the Google Drive and parses the file or attachment.

| Name | Type | Description |
| --- | --- | --- |
| **``searchString``** | String | String to find the E-mail with the relevant attachment or the relevant file in Drive. The search criteria for Drive are detailed in the [Google Drive SDK documentation](https://developers.google.com/drive/api/v3/search-parameters). Takes the first match. |
| **``searchPlace``** | String | (***optional***) Sets the location to search in. Can take the values of ``"m"`` for Mail and ``"d"`` for Drive. Defaults to ``"m"``. |
| **``fileName``** | String | (***optional***) Partial or full name of the file to use. Only works for searches in Inbox. Is required if the E-mail contains more than one attachement. |
| **``parseConfig``** | Object | (***optional***) Configuration object for PapaParse ([Documentation](https://www.papaparse.com/docs#config)). Defaults to ``{delimiter: ",", skipEmptyLines: true}``. |
| **``deleteBool``** | Boolean | (***optional***) If true the E-Mail or file will be moved to the trash after parsing. Defaults to ``true``. |
| **``safeMode``** | Boolean | (***optional***) Handles not properly [escaped double quotes](https://en.wikipedia.org/wiki/Comma-separated_values#Basic_rules) (according to [RFC4180](https://tools.ietf.org/html/rfc4180)) in the CSV-file. Significantly increases runtime and will time out for large data sets. Defaults to ``false``. |
| **``encoding``** | Boolean | (***optional***) Sets the encoding of the CSV-string, e.g. ``"ansi"``, ``"utf8"``, ``"cp1252"`` etc. . Defaults to ``"cp1252"``. |

##### Return

``Array[Array]`` 2-dimensional array representing the CSV-data


#### ``addDataToSheet (data, sheetName, spreadsheetId, type, startColumn, removeHeaderBool, headerRow, countRowsByColumn, useAdvancedServices)``

##### Parameters
| Name | Type | Description |
| --- | --- | --- |
| **``data``** | Array[Array] | The data to write into the sheet as a 2-dimensional array |
| **``sheetName``** | String | The name of the Sheet the data should be written to. |
| **``spreadsheetId``** | String | (***optional***) The ID of the  Spreadsheet the data should be written to. Can be omitted if the Script is bound to a Spreadsheet, in which case it will default to the bound Spreadsheet. |
| **``type``** | String | (***optional***) Determines if the data will overwrite the existing data or add it after it. Can take the values of ``"r"`` for "replace" and ``"a"`` for "add". Defaults to ``"r"``. |
| **``startColumn``** | Number \| String | (***optional***) The column the data start to be written in. Can be provided as the column number or the identifier in A1-Notation. Defaults to ``1`` or ``"A"`` respectively.|
| **``removeHeaderBool``** | Boolean | (***optional***) Will remove the first inner array from data if set to ``true``. Defaults to ``false``.|
| **``headerRow``** | Number | (***optional***) Sets the last row containing headers. Only needed if type is ``"r"``. The data will be written after the ``headerRow``. Defaults to ``1`` |
| **``countRowsByColumn``** | String | (***optional***) Column to use for finding the last row. Only becomes necessary if the Sheet contains formulas that span further down than the rest of the data. |
| **``useAdvancedServices``** | Boolean | (***optional***) If set to ``true`` the script will use the [Advanced Sheets Service](https://developers.google.com/apps-script/advanced/sheets) to write the data. Highly recommended for large data sets. Requires activating the Advanced Sheets Service and Sheets API. Defaults to ``false`` |