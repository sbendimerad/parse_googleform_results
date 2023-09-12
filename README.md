# Code README

## Introduction

This code is a Google Apps Script intended for use within Google Sheets. It calculates total scores based on responses in a specified sheet. The script assumes that the responses are organized in columns, with each column representing a different question or category.

## How to Use

### 1. Open Google Sheets

Make sure you have your Google Sheet document open in which you want to calculate scores.

### 2. Access Script Editor

- Click on `Extensions` in the top menu.
- Select `Apps Script` to open the Google Apps Script editor.

### 3. Paste Code

- Delete any existing code in the script editor.
- Paste the provided code into the script editor.

### 4. Modify Sheet Names (if needed)

If your sheets are named differently than "Reponse" and "Correspondence," replace these names with the actual sheet names you want to work with.

### 5. Run the Script

- Save your script if necessary.
- Click the play button (▶️) in the toolbar to run the `calculateScores` function.


## Code Explanation

### `calculateScores()`

This function is responsible for calculating total scores for each row in the specified sheet. It does this by:

- Fetching the necessary data from the active spreadsheet.
- Adding a new column for total scores.
- Looping through each row (excluding the header row).
- Looping through each column (starting from the third column) and calculating the total score for each row.

### `getCorrespondenceData(correspondenceSheet)`

This helper function retrieves and organizes the correspondence data from the "Correspondence" sheet into an object for efficient lookup. It is used by the `calculateScores` function to map response options to point values.

### `calculateCellScore(listAnswers, correspondenceData)`

Another helper function used by `calculateScores`, this function calculates the score for a single cell based on a list of response options and the correspondence data.

## Exemple of Input

Suppose you have a Google Sheet named "Reponse" with the following structure:
| Name  | Question 1          | Question 2     | Question 3 |
|-------|---------------------|---------------|------------|
| Alice | Yes, No, Yes       | No, Yes, No   | Yes, Yes   |
| Bob   | No, Yes, No        | Yes, No, Yes  | No, No     |
| Carol | Yes, Yes, No       | No, No, Yes   | Yes, Yes   |

And you have another sheet named "Correspondence" with the following data:
| Response Option | Point Value |
|-----------------|-------------|
| Yes             | 5           |
| No              | 2           |

## Exemple of Output

| Name  | Question 1          | Question 2     | Question 3 | Total Score |
|-------|---------------------|---------------|------------|-------------|
| Alice | Yes, No, Yes       | No, Yes, No   | Yes, Yes   | 24          |
| Bob   | No, Yes, No        | Yes, No, Yes  | No, No     | 12          |
| Carol | Yes, Yes, No       | No, No, Yes   | Yes, Yes   | 22          |

