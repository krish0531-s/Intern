# Document Automation System

A Python-based document automation tool that streamlines client onboarding and compliance document management.

The system automatically creates client folder structures, generates Excel-based document mapping sheets, and routes uploaded documents to the correct destination folders based on user selections.

## Features

* Configurable document templates using `checklist.txt`
* Automatic creation of client-specific folder structures
* Support for deeply nested folders
* Dynamic Excel dropdown generation from uploaded documents
* Human-in-the-loop document verification and classification
* Automatic document renaming and routing
* Optional custom naming convention (`Name_Particular.ext`)
* Support for reusing the same source document multiple times
* Standalone Windows executable for non-technical users

## Workflow

1. Run `DocumentAutomation.exe`
2. Enter the client name
3. Select the folder containing client documents
4. The application:

   * Creates the client folder structure
   * Copies all documents into a `Dump` folder
   * Generates an Excel mapping sheet
5. Select documents from the dropdown list in Excel
6. Optionally enter a custom name
7. Save and close Excel
8. Documents are automatically renamed and placed in the correct folders

## Example

Input:

```
scan1.pdf
scan2.pdf
```

Excel Mapping:

| Particular   | File      | Name  |
| ------------ | --------- | ----- |
| PAN Card     | scan1.pdf | Krish |
| Aadhaar Card | scan2.pdf |       |

Output:

```
KYC/
├── Krish_PAN Card.pdf
└── Aadhaar Card.pdf
```

## Folder Template Format

```
PAN Card|KYC
Aadhaar Card|KYC
GST Certificate|Auditor/GST
Bank Statement|Auditor/Financial
Director PAN|Auditor/KYC/Directors
```

Supports unlimited folder nesting.

## Technologies

* Python
* OpenPyXL
* Tkinter
* PyInstaller
* Excel Automation
* File System Automation

## Future Improvements

* GUI-based client name input
* Drag-and-drop document upload
* PDF preview during mapping
* Document validation rules
* Audit logs and reporting
