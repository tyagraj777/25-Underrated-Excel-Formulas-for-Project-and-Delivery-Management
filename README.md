# 25 Underrated Excel Formulas for Project and Delivery Management

In project and delivery management, Excel and other spreadsheet tools are indispensable for tracking, analyzing, and reporting data. However, many **underrated formulas** are often overlooked, despite their potential to save time, improve accuracy, and provide deeper insights. This README file outlines **25 underrated formulas** that can significantly enhance project and delivery management workflows.

---

## **Table of Contents**
1. [Introduction](#introduction)
2. [List of Underrated Formulas](#list-of-underrated-formulas)
   - [1. `IFERROR`](#1-iferror)
   - [2. `SUMIFS`](#2-sumifs)
   - [3. `COUNTIFS`](#3-countifs)
   - [4. `INDEX` + `MATCH`](#4-index--match)
   - [5. `XLOOKUP`](#5-xlookup)
   - [6. `NETWORKDAYS`](#6-networkdays)
   - [7. `WORKDAY`](#7-workday)
   - [8. `TEXTJOIN`](#8-textjoin)
   - [9. `UNIQUE`](#9-unique)
   - [10. `FILTER`](#10-filter)
   - [11. `SUBTOTAL`](#11-subtotal)
   - [12. `EDATE`](#12-edate)
   - [13. `DATEDIF`](#13-datedif)
   - [14. `ROUNDUP` / `ROUNDDOWN`](#14-roundup--rounddown)
   - [15. `CONCATENATE` (or `&`)](#15-concatenate-or-)
   - [16. `LEFT` / `RIGHT` / `MID`](#16-left--right--mid)
   - [17. `LEN`](#17-len)
   - [18. `ISBLANK`](#18-isblank)
   - [19. `CHOOSE`](#19-choose)
   - [20. `INDIRECT`](#20-indirect)
   - [21. `TRANSPOSE`](#21-transpose)
   - [22. `AVERAGEIFS`](#22-averageifs)
   - [23. `EOMONTH`](#23-eomonth)
   - [24. `SWITCH`](#24-switch)
   - [25. `LET`](#25-let)
3. [Why These Formulas Matter](#why-these-formulas-matter)

---

## **Introduction**
This document highlights **25 underrated Excel formulas** that are often forgotten but can significantly improve project and delivery management. These formulas help automate tasks, reduce errors, and provide deeper insights into project performance.

---

## **List of Underrated Formulas**

### **1. `IFERROR`**
- **Purpose:** Handles errors gracefully by replacing them with a custom value.
- **Use Case:** Avoid displaying errors in reports (e.g., `#DIV/0!` or `#VALUE!`).
- **Example:**  
  ```excel
  =IFERROR(A2/B2, "N/A")
  ```

---

### **2. `SUMIFS`**
- **Purpose:** Sums values based on multiple criteria.
- **Use Case:** Calculate total hours spent on specific tasks or projects.
- **Example:**  
  ```excel
  =SUMIFS(Hours, Project, "Cloud Migration", Status, "Completed")
  ```

---

### **3. `COUNTIFS`**
- **Purpose:** Counts cells that meet multiple criteria.
- **Use Case:** Track the number of tasks completed by a specific team member.
- **Example:**  
  ```excel
  =COUNTIFS(Assignee, "John", Status, "Completed")
  ```

---

### **4. `INDEX` + `MATCH`**
- **Purpose:** A powerful alternative to `VLOOKUP` for flexible lookups.
- **Use Case:** Retrieve specific data from a large table.
- **Example:**  
  ```excel
  =INDEX(DataRange, MATCH("Task Name", TaskColumn, 0))
  ```

---

### **5. `XLOOKUP`**
- **Purpose:** A modern replacement for `VLOOKUP` and `HLOOKUP`.
- **Use Case:** Look up values in any direction with fewer limitations.
- **Example:**  
  ```excel
  =XLOOKUP("Task Name", TaskColumn, DataRange)
  ```

---

### **6. `NETWORKDAYS`**
- **Purpose:** Calculates working days between two dates, excluding weekends and holidays.
- **Use Case:** Estimate project timelines or delivery dates.
- **Example:**  
  ```excel
  =NETWORKDAYS(StartDate, EndDate, Holidays)
  ```

---

### **7. `WORKDAY`**
- **Purpose:** Adds a specified number of working days to a date.
- **Use Case:** Calculate deadlines or milestones.
- **Example:**  
  ```excel
  =WORKDAY(StartDate, 10, Holidays)
  ```

---

### **8. `TEXTJOIN`**
- **Purpose:** Combines text from multiple cells with a delimiter.
- **Use Case:** Create a comma-separated list of task owners.
- **Example:**  
  ```excel
  =TEXTJOIN(", ", TRUE, AssigneeRange)
  ```

---

### **9. `UNIQUE`**
- **Purpose:** Extracts unique values from a range.
- **Use Case:** List unique project names or team members.
- **Example:**  
  ```excel
  =UNIQUE(ProjectRange)
  ```

---

### **10. `FILTER`**
- **Purpose:** Filters data based on specified criteria.
- **Use Case:** Display only tasks with a specific status.
- **Example:**  
  ```excel
  =FILTER(TaskRange, StatusRange = "In Progress")
  ```

---

### **11. `SUBTOTAL`**
- **Purpose:** Performs calculations (e.g., sum, average) on filtered data.
- **Use Case:** Calculate totals for visible rows after applying filters.
- **Example:**  
  ```excel
  =SUBTOTAL(9, HoursRange)
  ```

---

### **12. `EDATE`**
- **Purpose:** Adds a specified number of months to a date.
- **Use Case:** Calculate future milestones or deadlines.
- **Example:**  
  ```excel
  =EDATE(StartDate, 6)
  ```

---

### **13. `DATEDIF`**
- **Purpose:** Calculates the difference between two dates in days, months, or years.
- **Use Case:** Track project duration or team member tenure.
- **Example:**  
  ```excel
  =DATEDIF(StartDate, EndDate, "d")
  ```

---

### **14. `ROUNDUP` / `ROUNDDOWN`**
- **Purpose:** Rounds numbers up or down to a specified number of decimal places.
- **Use Case:** Simplify budget or timeline calculations.
- **Example:**  
  ```excel
  =ROUNDUP(TotalCost, 0)
  ```

---

### **15. `CONCATENATE` (or `&`)**
- **Purpose:** Combines text from multiple cells.
- **Use Case:** Create unique task IDs or labels.
- **Example:**  
  ```excel
  =A2 & "-" & B2
  ```

---

### **16. `LEFT` / `RIGHT` / `MID`**
- **Purpose:** Extracts specific parts of a text string.
- **Use Case:** Parse task codes or project IDs.
- **Example:**  
  ```excel
  =LEFT(TaskCode, 3)
  ```

---

### **17. `LEN`**
- **Purpose:** Returns the length of a text string.
- **Use Case:** Validate data entry (e.g., task descriptions).
- **Example:**  
  ```excel
  =LEN(Description)
  ```

---

### **18. `ISBLANK`**
- **Purpose:** Checks if a cell is empty.
- **Use Case:** Identify missing data in project plans.
- **Example:**  
  ```excel
  =IF(ISBLANK(A2), "Missing", "Complete")
  ```

---

### **19. `CHOOSE`**
- **Purpose:** Selects a value from a list based on an index.
- **Use Case:** Assign priority levels or statuses.
- **Example:**  
  ```excel
  =CHOOSE(Priority, "Low", "Medium", "High")
  ```

---

### **20. `INDIRECT`**
- **Purpose:** Creates a dynamic reference to a cell or range.
- **Use Case:** Build flexible dashboards or reports.
- **Example:**  
  ```excel
  =SUM(INDIRECT("A" & B2 & ":A" & C2))
  ```

---

### **21. `TRANSPOSE`**
- **Purpose:** Flips rows and columns in a range.
- **Use Case:** Reformat data for better visualization.
- **Example:**  
  ```excel
  =TRANSPOSE(DataRange)
  ```

---

### **22. `AVERAGEIFS`**
- **Purpose:** Calculates the average of values that meet multiple criteria.
- **Use Case:** Analyze team performance or task completion times.
- **Example:**  
  ```excel
  =AVERAGEIFS(Hours, Project, "Cloud Migration", Status, "Completed")
  ```

---

### **23. `EOMONTH`**
- **Purpose:** Returns the last day of a month after adding a specified number of months.
- **Use Case:** Calculate end-of-month deadlines or reporting dates.
- **Example:**  
  ```excel
  =EOMONTH(StartDate, 0)
  ```

---

### **24. `SWITCH`**
- **Purpose:** Evaluates a value against a list of conditions and returns a corresponding result.
- **Use Case:** Simplify nested `IF` statements for status or priority mapping.
- **Example:**  
  ```excel
  =SWITCH(Priority, 1, "Low", 2, "Medium", 3, "High")
  ```

---

### **25. `LET`**
- **Purpose:** Assigns a name to a calculation or value for reuse in a formula.
- **Use Case:** Simplify complex formulas and improve readability.
- **Example:**  
  ```excel
  =LET(x, A2+B2, x*C2)
  ```

---

## **Why These Formulas Matter**
These formulas are often underutilized but can significantly enhance project and delivery management by:
- **Saving Time:** Automating repetitive tasks and calculations.
- **Improving Accuracy:** Reducing manual errors in data analysis.
- **Enhancing Insights:** Providing deeper visibility into project performance.
- **Increasing Flexibility:** Adapting to dynamic project requirements.

By incorporating these formulas into your workflows, you can streamline processes, make data-driven decisions, and deliver projects more effectively.

--- 

Feel free to explore and use these formulas to elevate your project and delivery management practices!
