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

### Scoring system
The script matches students to their folders based on their IDs and names. However, not all students have their IDs listed, or the IDs won't match the ID in the folder. In addition, some students have complex names containing dashed compound names, middle initials, suffixes, etc. To overcome these issues, the script employes a scoring system to dynamically compare student data and folder names to find a match.

Each student is assigned a 'main' first name and a 'main' last name. For example, John Smith-Doe Jr. would have the main first name "John" and main last name "Doe." This is determined by picking the very first item in the first name and the second-to-last item in the last name.

Everything that isn't a main first or last name is designated as an 'extra' piece of the name.

The script then compares the student to each folder name, scoring as follows (where 15 points are needed for a match):
- If the folder contains the student's ID, add 10 points (almost guarenteed match)
- If it contains the student's first name, add 8 points
- If it contains the student's last name, add 6 points
- If it contains any 'extra' piece of the name, add 1 point
- If only one folder is a possible match at all, add 1 point for uniqueness

This ensures that if a student is missing their ID, a match can still be made so long as there aren't two folders that share the same student name.

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
