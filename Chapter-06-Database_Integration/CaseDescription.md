# 📘 Database Integration (PAD)

This chapter focuses on integrating Power Automate Desktop with databases to retrieve, manipulate, and store data efficiently.  
The cases cover different database operations without overlapping tasks.

---

## 📂 Cases Overview

### 01. Connect to a SQL Database
- **Objective:** Establish a connection to a SQL database and validate connectivity.
- **Hands-on Tasks:**
  - Configure a database connection in PAD using credentials.
  - Test the connection and handle possible errors.
  - Retrieve basic metadata (e.g., list of tables).
- **Skills Covered:** Database connections, authentication, error handling.

### 02. Retrieve Data from a Table
- **Objective:** Extract data from a specific table and store it in Excel.
- **Hands-on Tasks:**
  - Execute a `SELECT` query to fetch table data.
  - Loop through the results and write them to an Excel file.
- **Skills Covered:** SQL queries, data extraction, Excel integration.

### 03. Insert Data into a Database
- **Objective:** Add new records into a database table automatically.
- **Hands-on Tasks:**
  - Prepare sample data in PAD variables or Excel.
  - Execute an `INSERT` query dynamically for each record.
  - Validate that data has been successfully inserted.
- **Skills Covered:** Data insertion, query parameterization, database validation.

### 04. Update Existing Records
- **Objective:** Update specific rows in a database based on a condition.
- **Hands-on Tasks:**
  - Identify records that meet a condition (e.g., status = 'Pending').
  - Execute an `UPDATE` query to modify the records.
  - Log the number of rows affected.
- **Skills Covered:** Conditional updates, dynamic queries, logging changes.

### 05. Delete Records and Handle Transactions
- **Objective:** Delete records from a table safely with transaction handling.
- **Hands-on Tasks:**
  - Identify rows to delete based on specific criteria.
  - Execute `DELETE` queries within a transaction.
  - Commit or rollback based on validation.
- **Skills Covered:** Database transactions, safe deletion, error handling.

---

## 🎯 Goal
By completing these 5 hands-on cases, you will learn:

- Connecting to and authenticating with databases.
- Retrieving, inserting, updating, and deleting data.
- Executing SQL queries dynamically within PAD.
- Combining database operations with Excel and logging.
- Handling transactions and errors safely for real-world scenarios.
