# Intelligent Email-Based Data Entry Automation System

## 1. Project Overview

The Intelligent Email-Based Data Entry Automation System is designed to automate the process of collecting, extracting, organizing, validating, and entering customer-provided information into Excel spreadsheets.

The system allows customers to provide raw data through email without manually creating or editing Excel files.

The automation processes the received information, identifies individual data fields, organizes the information into suitable rows and columns, performs basic data validation, and generates a new Excel workbook containing the structured data.

The completed Excel file is automatically delivered to the customer through email.

## 2. Main Purpose

The primary purpose is to replace repetitive manual data entry tasks with an intelligent automated system that can transform unstructured or semi-structured customer information into organized Excel data.

## 3. Core Functionalities

- **Email Data Collection:** Receive customer information through email.
- **Data Extraction:** Extract relevant information from email messages.
- **Data Understanding:** Identify field names, values, and relationships between data items.
- **Data Structuring:** Organize information into appropriate rows and columns.
- **Data Validation:** Identify missing, invalid, or inconsistent values.
- **Data Cleaning:** Remove unnecessary spaces and standardize suitable data values.
- **Excel Data Entry:** Automatically insert processed data into a new Excel workbook.
- **Excel Generation:** Create a separate Excel file for each customer request.
- **Automated Delivery:** Send the completed Excel file back to the original customer.

## 4. System Workflow

Customer Sends Raw Data Through Email  
↓  
n8n Email Trigger  
↓  
Email Content Extraction  
↓  
AI-Based Data Understanding and Extraction  
↓  
Data Cleaning and Validation  
↓  
Structured JSON Generation  
↓  
Python API  
↓  
OpenPyXL Excel Data Entry  
↓  
New Excel Workbook Creation  
↓  
Email the Completed Excel File to Customer

## 5. Example Scenario

**Customer Email:**

Please enter the following customer information into Excel:

Ali Khan, 03001234567, Lahore, 22  
Ahmed Raza, 03123456789, Islamabad, 25  
Sara Ahmed, 03211234567, Karachi, 23

**Expected Excel Output:**

| Name | Phone | City | Age |
|---|---|---|---|
| Ali Khan | 03001234567 | Lahore | 22 |
| Ahmed Raza | 03123456789 | Islamabad | 25 |
| Sara Ahmed | 03211234567 | Karachi | 23 |

The system identifies the records, creates the required columns, enters all three records, and sends the newly generated Excel file to the customer.

## 6. Technology Stack

| Technology | Responsibility |
|---|---|
| n8n Cloud | Workflow automation |
| AI Model / LLM | Understanding and extracting customer data |
| Python | Data processing and validation |
| FastAPI | Communication between n8n and Python |
| OpenPyXL | Excel workbook creation and data entry |
| GitHub | Source code management |
| Render | Python API hosting |
| IMAP | Receiving customer emails |
| SMTP | Sending completed Excel files |

## 7. Project Scope

### Initial Version
- Receive data through email.
- Extract and structure customer information.
- Support multiple records in one email.
- Perform basic data validation.
- Generate a new Excel file for every email.
- Automatically send the completed file to the customer.

### Future Enhancements
- Understand different data formats and natural-language instructions.
- Support PDF, Word, CSV, and Excel data sources.
- Detect duplicate records.
- Support customer-specific Excel templates.
- Handle multiple sheets and complex data structures.
- Provide intelligent error detection and correction suggestions.
- Develop a web-based customer interface.

## 8. Expected Outcome

A working intelligent data entry automation service that converts customer-provided information into structured Excel spreadsheets without requiring customers to perform manual data entry.

## 9. Conclusion

This project aims to make data entry faster, easier, and more accurate by combining AI-powered data understanding, n8n workflow automation, Python data processing, and OpenPyXL Excel generation.
