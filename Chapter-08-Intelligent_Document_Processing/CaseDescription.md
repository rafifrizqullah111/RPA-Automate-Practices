# 📘 Intelligent Document Processing (IDP)

This chapter focuses on automating document understanding and data extraction from semi-structured and unstructured documents.  
The cases are designed to cover different aspects of reading, parsing, validating, and integrating extracted document data into downstream systems.

---

## 📂 Cases Overview

### 01. Extract Data from Invoices (Template-Based)
- **Objective:** Extract invoice number, date, and total amount from structured invoice PDFs.
- **Hands-on Tasks:**
  - Load sample invoice PDF.
  - Use OCR or text extraction to capture key fields.
  - Store extracted data into Excel.
- **Skills Covered:** OCR basics, text extraction, template-based document parsing, Excel integration.

### 02. Process Semi-Structured Purchase Orders
- **Objective:** Extract supplier, items, and total from varying PO formats.
- **Hands-on Tasks:**
  - Read PDF/Word purchase orders with different layouts.
  - Apply document understanding models to identify fields.
  - Output normalized data into a structured table.
- **Skills Covered:** Semi-structured data extraction, classification, data normalization.

### 03. Validate Extracted Data Against Master Records
- **Objective:** Perform validation of extracted invoice/PO data against an ERP master file.
- **Hands-on Tasks:**
  - Import master vendor list from Excel.
  - Compare extracted VendorID/Name with master data.
  - Flag mismatches for manual review.
- **Skills Covered:** Data validation, lookup operations, exception handling.

### 04. Automate Document Classification
- **Objective:** Classify incoming documents (Invoice, PO, Contract, Others).
- **Hands-on Tasks:**
  - Monitor an input folder with mixed documents.
  - Apply classification rules/models.
  - Move documents into respective folders (Invoices, POs, Contracts).
- **Skills Covered:** Document classification, file management, rule-based and ML-based classification.

### 05. End-to-End IDP Workflow
- **Objective:** Build a complete pipeline for Intelligent Document Processing.
- **Hands-on Tasks:**
  - Watch input folder for new files.
  - Extract, classify, and validate document data.
  - Push results to an output system (Excel/Database).
  - Generate an exception report for failed documents.
- **Skills Covered:** End-to-end automation design, exception handling, reporting.

---

## 🎯 Goal
By completing these hands-on cases, you will learn:

- OCR and text extraction from structured and semi-structured documents.  
- How to classify, parse, and validate business documents.  
- Data normalization and integration with downstream systems (Excel, DB, ERP).  
- Building end-to-end document processing automation pipelines.  
