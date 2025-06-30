# HACCP Data Management & Compliance Automation System (Excel-VBA)

## Overview

This repository showcases a robust Microsoft Excel-VBA system designed to automate the critical recordkeeping and data management processes required for Hazard Analysis and Critical Control Points (HACCP) compliance in a food processing environment. The system ensures data integrity, streamlines user entry, and provides a secure, auditable record of all compliance activities, significantly reducing the risk of record-based noncompliance reports (NRs).

## Problem Solved

Manual or inefficient HACCP recordkeeping can lead to data inaccuracies, compliance gaps, and significant administrative burden. The challenge was to create a centralized, user-friendly, and secure system that automates the documentation of critical control points and sanitation procedures, ensuring regulatory adherence and immediate audit readiness.

## Solution & Key Features

This system is built around a clear separation of user entry and secure recordkeeping, powered by Excel's native features and strategic VBA automation:

* **Dedicated Entry Worksheets:**
    The system utilizes five dedicated, unlocked worksheets for daily user data entry:
    * **'SSOP'** (Standard Sanitation Operating Procedures)
    * **'CCP'** (Critical Control Points)
    * **'DCA'** (Deviations and Corrective Actions)
    * **'TVP'** (Thermometer Verification Procedure)
    * **'Salmonella Testing'**
    Each sheet provides clear instructions for documentation.

* **Secure & Hidden Recordkeeping Worksheets:**
    Corresponding to each entry sheet, there are five respective recordkeeping sheets:
    * **'SSOP Record'**
    * **'CCP Record'**
    * **'DCA Record'**
    * **'TVP Record'**
    * **'ST Record'** (Salmonella Testing Record)
    - These sheets are locked and typically hidden, serving as a permanent, immutable record of all submitted data. They are accessible only by authorized managers or appointed users for audit purposes, allowing specific sections or time periods to be printed as needed for inspectors.

* **Automated Data Transfer & Validation (VBA Macros):**
    Each entry sheet features a command button that triggers a VBA macro (e.g., `MoveSSOP` for the 'SSOP' sheet) to securely transfer data to its respective recordkeeping sheet. The core logic for these macros is consistent across all entry types, ensuring uniformity and reliability.

    * **User Verification Prompt:** A critical step in the macro workflow is a `MsgBox` prompt that requires the user to verify the accuracy and completeness of their inputs before submission. This emphasizes user accountability and acts as a crucial last line of defense against errors, directly aiming to prevent NRs.
        ```vba
        'Prompt user to verify inputs
        strMsg = MsgBox("Do you verify that the SSOP information entered is accurate and complete?" & vbCrLf & vbCrLf & _
                        "Clicking yes will transfer the current SSOP information to a permanent record. " & _
                        "A printed copy will be available for the inspector upon confirmation. Requests to " & _
                        "view previous records must be processed through the current HACCP Administrator." & vbCrLf & vbCrLf & _
                        "Reminder: This submission is permanent. Invalid, incomplete, or missing records could likely result in an NR.", _
                        vbYesNo, "Verify Inputs")
        ```
    * **Secure Data Transfer:** Upon user verification, the macro temporarily unprotects the record sheet (using a password), copies the data from the entry sheet's `ListObject` (`tblSSOP` in the example) to the next available row in the record sheet's `ListObject` (`tblSSOPRecord`), clears the entry sheet's contents, and then immediately re-protects the record sheet.
        ```vba
        'Copy data to protected table and clear contents/clipboard
        wsRecord.Unprotect Password:="HACCP"
        wsSSOP.Range("tblSSOP").Copy Destination:=wsRecord.Cells(LastRow + 1, sh3Col)
        Application.CutCopyMode = False
        wsRecord.Protect Password:="HACCP"
        'Sheets("SSOP").PrintOut ' Converted to comment for GitHub showcase
        wsSSOP.Range("tblSSOP").ClearContents
        ```
    * **Automated Printing (Commented for Demo):** Originally, the system automatically printed a copy of the submitted entry sheet for immediate inspector review upon successful transfer. This line has been commented out (`'Sheets("SSOP").PrintOut`) for the purpose of this public demonstration.

* **Robust Workbook Save Logic (`wbSave` Macro):**
    The system includes a `wbSave` macro designed for robust saving and potentially backup operations. While commented out for this demonstration, its intended activation point is within the `Workbook_BeforeClose` event, ensuring critical data handling upon workbook closure.

## Impact & Results

This HACCP Data Management & Compliance Automation System delivered substantial benefits:

* **Ensured Regulatory Compliance:** Provided a structured and automated framework for documenting all necessary HACCP and sanitation records, significantly reducing the likelihood of record-based noncompliance reports (NRs) during inspections.
* **Enhanced Data Integrity & Security:** Automated data transfer to locked, hidden record sheets prevented accidental modification or corruption of historical compliance data, ensuring audit readiness.
* **Streamlined Recordkeeping:** Minimized manual effort and human error in daily data entry and record management.
* **Improved Audit Readiness:** Provided quick and secure access to historical records for inspectors, facilitating efficient and transparent audits.
* **Increased User Accountability:** The explicit verification prompts fostered a culture of meticulous data entry among users.

## Technologies Used

* **Microsoft Excel:** Excel Tables (ListObjects), Named Ranges, Worksheet Protection, Data Validation.
* **VBA (Visual Basic for Applications):** Macros (`Sub` procedures), `MsgBox` for user interaction, `Worksheet` Object Model, `Workbook_BeforeClose` Event, `Application.ScreenUpdating` optimization, `Range.Copy`, `Range.ClearContents`, Password Protection via VBA.

## Getting Started / How to Use

1.  **Download & Open:** Clone this repository or download the `excel-vba-haccp-compliance.xlsm` file and open it in Microsoft Excel (ensure macros are enabled).
2.  **Data Entry:** Navigate to any of the five entry sheets ('SSOP', 'CCP', 'DCA', 'TVP', 'Salmonella Testing').
3.  **Input Data:** Enter the relevant HACCP compliance data into the designated fields on the active entry sheet.
4.  **Submit Record:** Click the command button on the entry sheet. Carefully review the verification prompt before confirming submission.
5.  **Access Records:** (For authorized users) Access the hidden record sheets (e.g., 'SSOP Record') to view historical compliance data. Use the password "HACCP" to unprotect sheets

*Note: The workbook has been pre-populated with anonymized dummy data to fully convey its functionality.*

## Anonymization Note

Please note that all sensitive and proprietary data from the original project, including plant names and personal identifiers, has been replaced with dummy data to protect confidentiality. The structure, formulas, and core VBA functionality remain fully intact, demonstrating the complete capabilities.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
