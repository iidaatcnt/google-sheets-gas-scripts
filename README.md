# Google Sheets GAS Scripts for Construction Management

This repository contains Google Apps Script (GAS) files designed to help manage construction or service-related businesses using Google Sheets. It provides scripts for setting up a structured spreadsheet and inserting sample data.

## Overview

These scripts automate the creation of essential sheets for managing customer information, construction/service history, and complaints/notes. They are ideal for small businesses or individuals who want to leverage Google Sheets for simple CRM and project tracking without complex database setups.

## Features

-   **`setupSheets.gs`:**
    -   Creates three main sheets: "顧客マスタシート" (Customer Master Sheet), "施工履歴シート" (Construction History Sheet), and "クレーム・注意点シート" (Complaint/Notes Sheet).
    -   Sets up predefined headers for each sheet.
    -   Applies data validation rules (e.g., dropdowns for customer type, prefectures, work types, and status).
    -   Freezes header rows for better readability.
-   **`insertSampleData.gs`:**
    -   Inserts sample data into the "顧客マスタシート," "施工履歴シート," and "クレーム・注意点シート" for quick testing and demonstration.
-   **`onOpen()` Function:**
    -   Automatically creates a custom menu named "カスタムメニュー" (Custom Menu) in the Google Sheet.
    -   Adds menu items to easily run `setupConstructionManagementSheets` and `insertSampleData` functions directly from the spreadsheet.

## How to Use

1.  **Create a new Google Sheet:** Go to Google Sheets and create a new blank spreadsheet.
2.  **Open the Script Editor:** From the new Google Sheet, go to `Extensions > Apps Script`.
3.  **Copy and Paste Scripts:**
    -   In the Apps Script editor, delete any existing code (e.g., `Code.gs`).
    -   Create new script files (e.g., `insertSampleData.gs`, `setupSheets.gs`) and copy the content from the respective files in this repository into them.
    -   Ensure the `onOpen()` function is also included in one of the `.gs` files (e.g., `setupSheets.gs`).
4.  **Save the Project:** Click the save icon (floppy disk) in the Apps Script editor.
5.  **Run `setupSheets.gs`:**
    -   Go back to your Google Sheet.
    -   A new custom menu named "カスタムメニュー" should appear in the menu bar.
    -   Click `カスタムメニュー > シートをセットアップ` to create and configure the sheets.
6.  **Run `insertSampleData.gs` (Optional):**
    -   After setting up the sheets, click `カスタムメニュー > サンプルデータを挿入` to populate them with sample data.

## Technologies Used

-   Google Apps Script (JavaScript-based)
-   Google Sheets

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details.
