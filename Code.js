/**
 * Tayba Foundation - Student Folder Script
 * @fileoverview Script to sort students depending on whether they have a folder or not.
 * @author Muhammad Conn <muhammad.conn@icloud.com>, Tayba Foundation
 *
 * # Usage
 *
 */



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

    // Column indices (0-based)
    const firstNameColIndex = 0; // Column A
    const lastNameColIndex = 1;  // Column B
    const stateColIndex = 2; // Column C
    const prisonIdColIndex = 3; // Column D
    const inactiveColIndex = 4; // Column E
    const folderColIndex = 5; // Column F
    const releasedColIndex = 6; // Column G
    const contactIdColIndex = 7; // Column H
    const createdDateColIndex = 8; // Column I
    const lastModifiedColIndex = 9; // Column J



    // Get all values in the range
    const values = studentsToSearchSheet.getRange(startRow, startCol, numRows, numCols).getValues();

    // Map letters to their folder IDs
    const parentFolder = DriveApp.getFolderById("1gwv4_UYxNld1rTwkdqdcYq-C4QnDkqbJ");
    const folderLetterMap = createFolderLetterMap(parentFolder);

    Logger.log("Folder map created: " + JSON.stringify(folderLetterMap));

    // Iterate through each row of data
    for (const row of values) {
        const student = createStudentObject(row);
        Logger.log("Processing student: " + JSON.stringify(student));

        // Skip if the student already has a folder or is inactive
        if (student.folder ) {
            Logger.log("Skipping student (has folder): " + student.firstName + " " + student.lastName);
        }
        else {
            const studentFolder = searchStudentFolders(student, folderLetterMap);
            if (studentFolder) {
                // Student folder found
                Logger.log("Folder found for student: " + student.firstName + " " + student.lastName + " Folder ID: " + studentFolder);
                student.folder = `https://drive.google.com/drive/u/0/folders/${studentFolder}`;
                writeStudentToSheet(studentsWithFolderSheet, student, true);
            }
            else {
                // Student folder not found
                Logger.log("No folder found for student: " + student.firstName + " " + student.lastName);
                writeStudentToSheet(studentsWithoutFolderSheet, student, false);
            }
        }

    }


}

/**
 * Creates a map of letter folders to their corresponding folder IDs.
 * @param parentFolder {Folder} The parent folder containing letter folders.
 * @returns {{}} A map where keys are letters and values are folder IDs.
 */
function createFolderLetterMap(parentFolder) {
    const folders = parentFolder.getFolders();

    // Create a map of folder names to folder objects for quick lookup
    const folderMap = {};
    const regex = /^[A-Z](-[A-Z]){0,2}$/; // Tests for letter folders
    while (folders.hasNext()) {
        const folder = folders.next();
        const folderName = folder.getName().trim();

        // If the folder name  is in the pattern of A, A-B, or A-B-C
        if (regex.test(folderName)) {

            // Store the ID of the folder under the letter key
            // Since multiple folders may exist for a single letter, we store an array of IDs
            const folderId = folder.getId();
            const letters = folderName.split("-");
            for (const letter of letters) {
                if (folderMap[letter]) {
                    folderMap[letter].push(folderId);
                }
                else {
                    folderMap[letter] = [folderId];
                }
            }
        }
    }
    return folderMap;
}

/**
 * Creates a student object from a row of data.
 * @param row {Array} An array representing a row of student data.
 * @returns {{firstName: string, lastName: string, inactive: string, folder: string, createdDate: *, contactId: string, prisonId: string, state: string, lastModified: *, released: string}}
 */
function createStudentObject(row) {
    return {
        firstName: row[0].toString().trim(),
        lastName: row[1].toString().trim(),
        state: row[2].toString().trim(),
        prisonId: row[3].toString().trim(),
        inactive: row[4].toString().trim(),
        folder: row[5].toString().trim(),
        released: row[6].toString().trim(),
        contactId: row[7].toString().trim(),
        createdDate: row[8],
        lastModified: row[9]
    };
}

/**
 * Searches for a student's folder based on their last name initial and ID.
 * @param student {Object} The student object containing firstName, lastName, and prisonId.
 * @param folderLetterMap {Object} A map of letters to folder IDs.
 * @returns {undefined|string} The ID of the found folder, or undefined if not found.
 */
function searchStudentFolders(student, folderLetterMap) {

    const lastInitial = student.lastName.charAt(0).toUpperCase();
    const studentId = student.prisonId;
    const firstName = student.firstName;
    const lastName = student.lastName;

    if (folderLetterMap.hasOwnProperty(lastInitial)) {
        const folders = folderLetterMap[lastInitial];

        // Since multiple folders may exist for a single letter, we need to search each one
        for (const folder of folders) {
            // Search each subfolder for the student folder
            const letterFolder = DriveApp.getFolderById(folder);
            const foundFolderId = searchWithinLetterFolder(letterFolder, studentId, firstName, lastName);
            if (foundFolderId) {
                return foundFolderId; // Return the found folder ID
            }
        }
        return undefined; // Not found in any of the letter folders
    }
    return undefined; // No folder for this initial
}

/**
 * Searches within a letter folder for a student folder by ID or name.
 * @param folder {Folder} The letter folder to search within.
 * @param studentId {string} The student ID to search for.
 * @param firstName {string} The student's first name.
 * @param lastName {string} The student's last name.
 * @returns {string} The ID of the found folder, or undefined if not found.
 */
function searchWithinLetterFolder(folder, studentId, firstName, lastName) {
    const subfolders = folder.getFolders();
    while (subfolders.hasNext()) {
        const subfolder = subfolders.next();
        const folderName = subfolder.getName().trim();

        // Clean texts for comparison
        const cleanedFolderName = cleanText(folderName);
        const cleanedFirstName = removeMiddleInitials(firstName, lastName).firstName
        const cleanedLastName = removeMiddleInitials(firstName, lastName).lastName;
        const cleanedStudentId = cleanText(studentId);

        // Check if the folder name contains the student ID
        if (cleanedFolderName.includes(cleanedStudentId)) {
            return subfolder.getId();
        }


        // If the student's last name contains a dash, we do 2 searches with full last name and only the first part
        const lastNameParts = lastName.split("-");
        if (lastNameParts.length > 1) {
            const shortLastName = lastNameParts[0].trim();
            if (cleanedFolderName.includes(cleanedFirstName) && cleanedFolderName.includes(cleanText(shortLastName))) {
                return subfolder.getId();
            }
        }

        if (cleanedFolderName.includes(cleanedFirstName) && cleanedFolderName.includes(cleanedLastName)) {
            return subfolder.getId();
        }
    }
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
 * Cleans text by removing diacritics and converting to lowercase.
 * @param text {string} The text to clean.
 * @returns {string} The cleaned text.
 */
function cleanText(text) {
    return text.normalize('NFD').replace(/\p{Diacritic}/gu, '').toLowerCase();
}

/**
 * Removes middle initials from first and last names.
 * @param firstName
 * @param lastName
 * @returns {{firstName: string, lastName: string}}
 */
function removeMiddleInitials(firstName, lastName) {
    // Remove middle initials from first name
    let cleanedFirstName = firstName.replace(/\s+[A-Z]\.?(\s+|$)/g, ' ').trim();

    // Remove middle initials from last name
    let cleanedLastName = lastName.replace(/\s+[A-Z]\.?(\s+|$)/g, ' ').trim();

    return {
        firstName: cleanedFirstName,
        lastName: cleanedLastName
    };
}