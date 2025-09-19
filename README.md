# Student Folders Sorting Script
## Description
This script uses Google Appscript to take a spreadsheet of students as well as a set of google drive folders and sort students based on whether a match can be found between a student in the spreadsheet and a folder in their name.

## Features
- Uses confidence scoring system for best accuracy
- Can handle irregular names, middle initials, and suffixes
- If a folder is found, it adds the link to the sheet
- Sorts students into their own sheets based on three categories:
-   Student folder already listed
-   Student folder exists but isn't listed
-   Student folder does not exist

## Requirements
- A Google Sheet formatted as follows:
-   Row 1 - header
-   Column A - First names
-   Column B - Last names
-   Column D - Ids
-   Column F - Folder links
-   Contains sheets named "Students to search", "Students with folder but no link in SF", "Students without folder".
- A Google drive folder containing subfolders with letter names and student folders within each of those subfolders that contain student names/ids

## Usage
1. Open the Google Sheet containing the students
2. Open the Apps Script editor (Extensions > Apps Script).
3. Copy and paste the provided JS file as well as the appsscript.json file into the editor, save.
4. Return to your sheet, click "Tayba > Sort Students for Folders"
<img width="1271" height="176" alt="image" src="https://github.com/user-attachments/assets/90a0480f-c147-4727-9d80-59242640001f" />
5. Open "Students with folder but no link in SF" and "Students without folder" sheets to view results.
