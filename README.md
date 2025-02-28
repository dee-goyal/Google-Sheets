# Web Application Mimicking Google Sheets

## Overview
This is a basic web application that mimics the core functionalities of Google Sheets. The application provides a spreadsheet interface with essential mathematical operations, basic data quality functions, and simple UI interactions.

## Features Implemented
- **Spreadsheet Interface**:  
  - Basic grid layout resembling Google Sheets  
  - Ability to enter and edit data in cells  
  - Simple cell formatting (bold, italics, font color)  

- **Mathematical Functions**:  
  - `SUM(range)`: Calculates the sum of selected cells  
  - `AVERAGE(range)`: Computes the average of selected cells  
  - `MAX(range)`: Returns the maximum value from selected cells  
  - `MIN(range)`: Returns the minimum value from selected cells  
  - `COUNT(range)`: Counts numeric values in the selected cells  

- **Data Quality Functions**:  
  - `TRIM(text)`: Removes leading and trailing whitespace  
  - `UPPER(text)`: Converts text to uppercase  
  - `LOWER(text)`: Converts text to lowercase  

- **Data Entry and Validation**:  
  - Users can input numbers and text  
  - Basic validation for numeric cells  


## Tech Stack Used
- **Frontend**: HTML, CSS, JavaScript  
- **Backend**: (Not implemented yet)  
- **Data Storage**: Local storage (for temporary data retention)  

## Data Structures Used
- **2D Array**: Used to store cell data and track user input  
- **Objects**: Used to store cell properties (e.g., formatting)  

## How to Run
1. Clone this repository:  
   ```sh
   git clone <repo-url>
