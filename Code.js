/**
 * Tayba Foundation - Student Folder Script
 * @fileoverview Script to sort students depending on whether they have a folder or not.
 * @author Muhammad Conn <muhammad.conn@icloud.com>, Tayba Foundation
 *
 * # Usage
 *
 */

// Column indices (0-based)
const firstNameColIndex = 0; // Column A
const lastNameColIndex = 1;  // Column B
const stateColIndex = 2; // Column C
const prisonIdColIndex = 3; // Column D
const inactiveColIndex = 4; // Column E
const folderColIndex = 5; // Column F
const releasedColIndex = 6; // Column G
const contactIdColIndex = 7// Column H
const createdDateColIndex = 8; // Column I
const lastModifiedColIndex = 9; // Column J

/*------------------------------------------------------------------
  Main Function - Sort Students
-------------------------------------------------------------------*/
function sortStudents() {
    // Initialize constants
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const studentsToSearchSheet = ss.getSheetByName("Students to search");
    const studentsWithFolderSheet = ss.getSheetByName("Students with folder but no link in SF");
    const studentsWithoutFolderSheet = ss.getSheetByName("Students without Folder");
    const startRow = 2; // First row is header
    const startCol = 1;
    const numRows = studentsToSearchSheet.getLastRow() - startRow + 1;
    const numCols = studentsToSearchSheet.getLastColumn() - startCol + 1;



    // Get all values in the range
    const values = studentsToSearchSheet.getRange(startRow, startCol, numRows, numCols).getValues();

    // Extract all folders at once to avoid repeated calls to DriveApp, and create index
    const parentFolder = DriveApp.getFolderById("1gwv4_UYxNld1rTwkdqdcYq-C4QnDkqbJ");
    const allStudentFolders = [];
    const letterFolders = parentFolder.getFolders();
    while (letterFolders.hasNext()) {
        const letterFolder = letterFolders.next();
        const studentSubfolders = letterFolder.getFolders();
        while (studentSubfolders.hasNext()) {
            const subfolder = studentSubfolders.next();
            // Store all the info you need in an object
            allStudentFolders.push({
                id: subfolder.getId(),
                name: subfolder.getName()
            });
        }
    }
    Logger.log("Built index of " + allStudentFolders.length + " total student folders.");


    // Iterate through each row of data

    let rowsForFoundSheet = [];
    let rowsForNotFoundSheet = [];
    for (const row of values) {
        const student = createStudentObject(row);
        Logger.log("Processing student: " + JSON.stringify(student));

        // Skip if the student already has a folder
        if (student.folder) {
            Logger.log("Skipping student (has folder): " + student.firstName + " " + student.lastName);
            continue;
        }

        // Search for the student's folder
        for (const folder of allStudentFolders) {
            // If it's found, save it and break the loop
            if (detectStudentFolderMatch(student.firstName, student.lastName, student.prisonId, folder.name)) {
                student.folder = "https://drive.google.com/drive/u/folders/" + folder.id;
                Logger.log("Found folder for student: " + student.firstName + " " + student.lastName + " -> " + folder.name);
                rowsForFoundSheet.push(row);
                break;
            }
            /**
            else {
                Logger.log("No match in folder: " + folder.name);
            }*/
        }
        if (!student.folder) {
            Logger.log("No folder found for student: " + student.firstName + " " + student.lastName);
            rowsForNotFoundSheet.push(row);
        }

    }

    // Batch write results //
    if (rowsForFoundSheet.length > 0) {
        studentsWithFolderSheet.getRange(studentsWithFolderSheet.getLastRow() + 1, 1, rowsForFoundSheet.length, rowsForFoundSheet[0].length)
            .setValues(rowsForFoundSheet);
    }
    if (rowsForNotFoundSheet.length > 0) {
        studentsWithoutFolderSheet.getRange(studentsWithoutFolderSheet.getLastRow() + 1, 1, rowsForNotFoundSheet.length, rowsForNotFoundSheet[0].length)
            .setValues(rowsForNotFoundSheet);
    }

}


/**
 * Creates a student object from a row of data.
 * @param row {Array} An array representing a row of student data.
 * @returns {{firstName: string, lastName: string, inactive: string, folder: string, createdDate: *, contactId: string, prisonId: string, state: string, lastModified: *, released: string}}
 */
function createStudentObject(row) {
    return {
        firstName: row[firstNameColIndex].toString().trim(),
        lastName: row[lastNameColIndex].toString().trim(),
        state: row[stateColIndex].toString().trim(),
        prisonId: row[prisonIdColIndex].toString().trim(),
        inactive: row[inactiveColIndex].toString().trim(),
        folder: row[folderColIndex].toString().trim(),
        released: row[releasedColIndex].toString().trim(),
        contactId: row[contactIdColIndex].toString().trim(),
        createdDate: row[createdDateColIndex],
        lastModified: row[lastModifiedColIndex]
    };
}


/**
 * Writes a student's data to the specified sheet.
 * @param sheet {Sheet} The sheet to write to.
 * @param student {Object} The student object containing their data.
 * @param hasFolder {boolean} Whether the student has a folder.
 */
function writeStudentToSheet(sheet, student, hasFolder) {
    const newRow = [
        student.firstName,
        student.lastName,
        student.state,
        student.prisonId,
        student.inactive,
        hasFolder ? student.folder : "No Folder",
        student.released,
        student.contactId,
        student.createdDate,
        student.lastModified
    ]

    sheet.appendRow(newRow);
}

/**
 * Cleans text by removing diacritics and non-alphanumeric characters (except spaces and dashes) and converting to lowercase.
 * @param text {string} The text to clean.
 * @returns {string} The cleaned text.
 */
function cleanText(text) {
    return text.normalize('NFD').replace(/\p{Diacritic}/gu, '').replace(/[^a-zA-Z0-9 -]/g, '').toLowerCase();
}


/**
 * Detects if a student's folder name matches their name and ID using confidence scoring
 * @param studentFirstName first name of the student
 * @param studentLastName last name of the student
 * @param studentId ID of the student
 * @param folderName name of the folder to check
 * @returns {boolean} true if the folder name matches the student, false otherwise
 */
function detectStudentFolderMatch(studentFirstName, studentLastName, studentId, folderName) {

    // Step 1 - prepare student data //

    // Clean texts for comparison
    const cleanedFirstName = cleanText(studentFirstName);
    const cleanedLastName = cleanText(studentLastName);


    // Divide name and folder name into parts by spaces and dashes
    const firstNameParts = cleanedFirstName.split(/[\s-]+/);
    const lastNameParts = cleanedLastName.split(/[\s-]+/);


    // Define primary name parts by their position.
    const primaryFirstName = firstNameParts[0] || '';
    const primaryLastName = lastNameParts.length > 0 ? lastNameParts[lastNameParts.length - 1] : '';

    // Collect all other name parts as extra parts.
    const extraNameParts = [
        ...firstNameParts.slice(1), // All but the first part of the first name
        ...lastNameParts.slice(0, -1) // All but the last part of the last name
    ].filter(p => p); // Ensure no blank parts are included.


    // Sometimes students have 2 ids separated by a slash, so we need to get both of them
    let studentIds = [];
    if (studentId && studentId.trim() !== '') {
        studentIds = studentId.split("/")
            .map(id => id.trim().replace(/^0+/, '')); // Remove leading zeros and trim whitespace;
    }


    // Step 2 - prepare folder data //
    const cleanedFolderName = cleanText(folderName);
    const folderNameParts = cleanedFolderName.split(/[\s-]+/);


    // step 3 - scoring //

    let score = 0;

    // Add 10 points if the ID matches
    if (studentIds.length > 0) {
        for (const id of studentIds) {
            if (folderNameParts.includes(id)) {
                score += 10;
                break; // Only add points once even if multiple IDs match
            }
        }
    }

    // Primary name score
    if (primaryFirstName && folderNameParts.includes(primaryFirstName)) {
        score += 8;
    }
    if (primaryLastName && folderNameParts.includes(primaryLastName)) {
        score += 6;
    }

    // +1 point for every extra name part that matches
    for (const part of extraNameParts) {
        if (folderNameParts.includes(part)) {
            score += 1;
        }
    }

    // Testing //
    /**Logger.log("Cleaned First Name: " + cleanedFirstName);
    Logger.log("Cleaned Last Name: " + cleanedLastName);
    Logger.log("Cleaned Folder Name: " + cleanText(folderName));

    Logger.log("Main first name: " + primaryFirstName);
    Logger.log("Main last name: " + primaryLastName);
    Logger.log("Extra name parts: " + JSON.stringify(extraNameParts));
    Logger.log("Student IDs: " + JSON.stringify(studentIds));
    Logger.log("Score: " + score);*/

    // Determine if the score meets the threshold for a match
    const threshold = 14;
    return score >= threshold;
}

function testDetectStudentFolderMatch() {
    const isMatch = detectStudentFolderMatch("John A.", "Doe-Smith", "01234/56789", "John Doe 1234");
    Logger.log("Match result: " + isMatch); // Expected output: true
}