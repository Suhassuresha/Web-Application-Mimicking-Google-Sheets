# Web Application Mimicking Google Sheets

Welcome to the Web Application Mimicking Google Sheets! This project aims to provide a fully functional web-based spreadsheet application similar to Google Sheets, with a focus on mathematical and data quality functions, data entry, and key UI interactions.

## Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Tech Stack](#tech-stack)
- [Data Structures](#data-structures)
- [Usage Examples](#usage-examples)
- [Getting Started](#getting-started)
- [Contributing](#contributing)
- [License](#license)

## Overview

The Web Application Mimicking Google Sheets is designed to replicate some of the core functionalities of Google Sheets. Users can add and delete rows and columns, apply formulas, and style cells. The application is built with accessibility and ease of use in mind, providing a familiar interface for users who are accustomed to traditional spreadsheet software.

[Check out the live project here](http://your-live-project-link.com)

## Features

- Add and delete rows and columns
- Apply mathematical and text-based formulas
- Basic cell styling (bold, italic, font size, color)
- Save and load spreadsheets (upcoming feature)
- Data visualization (upcoming feature)
- Data quality functions (e.g., remove duplicates, find and replace)

## Tech Stack

The application is built using the following technologies:

- **React**: A JavaScript library for building user interfaces. React's component-based architecture makes it easy to create reusable UI components.
- **XLSX**: A powerful library to parse and write spreadsheet files. It is used to handle file operations for loading and saving spreadsheets.
- **CSS**: For styling the application and ensuring a responsive design.

### Why These Technologies?

- **React**: React is well-suited for building dynamic and interactive user interfaces. Its virtual DOM and component-based architecture make it efficient and maintainable.
- **XLSX**: This library provides robust support for Excel file formats, making it a natural choice for handling spreadsheet data.
- **CSS**: CSS is essential for styling and ensuring that the application is visually appealing and user-friendly.

## Data Structures

The application uses a few key data structures to manage spreadsheet data and operations:

- **2D Array**: The spreadsheet data is stored in a 2D array, where each element represents a cell. This structure allows for efficient access and manipulation of cell data.
- **Object**: Cell styles are managed using objects, where each key represents a cell's coordinates, and the value is an object containing style properties.
- **Set**: Used for operations like removing duplicates, ensuring that each element is unique.

### Why These Data Structures?

- **2D Array**: Provides a straightforward way to represent the grid structure of a spreadsheet, allowing for easy access and updates to cell data.
- **Object**: Allows for flexible and dynamic management of cell styles, enabling easy updates and retrieval of style properties.
- **Set**: Ensures uniqueness in operations like removing duplicates, making such operations efficient and straightforward.

## Usage Examples

### Mathematical Functions

1. **SUM**: Calculates the sum of a range of cells.
   ```plaintext
   =SUM(A1:A10)
   ```
   This formula will add all the values from cells A1 to A10.

2. **AVERAGE**: Calculates the average of a range of cells.
   ```plaintext
   =AVERAGE(B1:B10)
   ```
   This formula will calculate the average of the values from cells B1 to B10.

3. **MAX**: Returns the maximum value from a range of cells.
   ```plaintext
   =MAX(C1:C10)
   ```
   This formula will return the maximum value from cells C1 to C10.

4. **MIN**: Returns the minimum value from a range of cells.
   ```plaintext
   =MIN(D1:D10)
   ```
   This formula will return the minimum value from cells D1 to D10.

5. **COUNT**: Counts the number of cells containing numerical values in a range.
   ```plaintext
   =COUNT(E1:E10)
   ```
   This formula will count the number of cells with numerical values from cells E1 to E10.

### Data Quality Functions

1. **TRIM**: Removes leading and trailing whitespace from a cell.
   ```plaintext
   =TRIM(F1)
   ```
   This formula will remove any leading and trailing whitespace from the value in cell F1.

2. **UPPER**: Converts the text in a cell to uppercase.
   ```plaintext
   =UPPER(G1)
   ```
   This formula will convert the text in cell G1 to uppercase.

3. **LOWER**: Converts the text in a cell to lowercase.
   ```plaintext
   =LOWER(H1)
   ```
   This formula will convert the text in cell H1 to lowercase.

## Getting Started

To get started with the Web Application Mimicking Google Sheets, follow these steps:

1. Clone the repository:
   ```bash
   git clone https://github.com/Suhassuresha/Web-Application-Mimicking-Google-Sheets.git
   cd Web-Application-Mimicking-Google-Sheets
   ```

2. Install dependencies:
   ```bash
   npm install
   ```

3. Run the application:
   ```bash
   npm start
   ```

The application will be available at `http://localhost:3000`.

## Contributing

We welcome contributions from the community. If you have suggestions or find any issues, please open an issue or submit a pull request. For major changes, please discuss them in an issue first.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.