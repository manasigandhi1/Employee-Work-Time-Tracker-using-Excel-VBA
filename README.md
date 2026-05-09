# ⏱️ Employee Work Time Tracker using Excel VBA

## 📌 Project Overview

Developed an automated work-time tracking system in Excel using VBA macros to record employee clock-in and clock-out activities.

The project was designed to simplify daily work-hour logging, reduce manual tracking errors, and create structured time-log data that can later be used for reporting and productivity analysis.

The system automatically records timestamps, validates entries, calculates work duration, and prevents incomplete or duplicate records.

[![Project Demo](https://img.youtube.com/vi/p6K5bJ27bjY/0.jpg)](https://www.youtube.com/watch?v=p6K5bJ27bjY)

---

## 🎯 Project Objective

Manual work-hour tracking often leads to:

- Missing timestamps
- Duplicate entries
- Incorrect duration calculations
- Incomplete attendance records
- Difficulty in productivity tracking

This project was built to automate time logging and improve accuracy using Excel VBA.

The system ensures clean and structured attendance data while minimizing manual intervention.

---

## 🛠️ Tools & Technologies Used

- **Excel VBA** → Automation and business logic
- **Excel Tables** → Structured data storage
- **Macros** → Interactive clock-in and clock-out system

---

## 🔧 VBA Logic & Automation Implemented

This project demonstrates practical VBA automation and workflow validation techniques.

### ✅ Dynamic Column Detection

Created a reusable VBA function to dynamically identify column positions based on header names.

```vba
Function GetColumn(ws As Worksheet, header As String)
```

### ✅ Automated Clock-In System

The Clock-In macro:

- Automatically inserts current date and timestamp
- Detects incomplete previous entries
- Prevents multiple active sessions
- Reuses empty rows when available
- Adds new table rows dynamically when required

### ✅ Automated Clock-Out System

The Clock-Out macro:

- Validates whether Clock-In exists
- Prevents duplicate Clock-Out actions
- Uses confirmation popups before submission
- Automatically records Clock-Out timestamp
- Calculates work duration in decimal hours
- Generates formatted `hh:mm` duration display

### ✅ Data Validation & Error Handling

Implemented validation checks using:

- `If conditions`
- `MsgBox alerts`
- `WorksheetFunction.CountA`
- Table object handling (`ListObject`)
- Row-level validation logic

### ✅ Automated Duration Calculation

The system calculates:

- Decimal duration for analytics and Power BI integration
- Formatted time duration for user readability

```vba
Round((ClockOut - ClockIn) * 24, 2)
```

### ✅ Structured Table-Based Logging

Used Excel Tables (`ListObject`) to:

- Maintain structured records
- Support scalability
- Improve automation reliability
- Enable future dashboard integration

---

## 📈 Features of the System

- One-click Clock-In button
- One-click Clock-Out button
- Automatic timestamp recording
- Work duration calculation
- Validation against incomplete entries
- Duplicate prevention logic
- User-friendly popup messages
- Structured time-log database

---

## 🌍 Real-World Importance of This Project

This project demonstrates how VBA can automate repetitive operational tasks commonly performed in workplaces.

### Real-world applications include:

- Employee attendance tracking
- Freelancer work-hour logging
- Productivity monitoring
- Payroll support systems
- Shift management
- Small business workforce tracking
- Work-from-home monitoring systems

### Business Benefits

- Reduces manual entry errors
- Saves administrative time
- Improves attendance accuracy
- Creates analytics-ready structured data
- Helps track employee productivity
- Supports reporting and dashboard creation

---

## 🚀 Skills Demonstrated

### Excel VBA Skills

- VBA Macros
- Functions & Procedures
- Dynamic Table Handling
- Date & Time Automation
- Conditional Logic
- Error Handling
- Row Validation
- User Interaction using MsgBox
- Automation Workflow Design

### Analytical & Operational Skills

- Process Automation
- Data Structuring
- Productivity Tracking
- Attendance System Logic
- Business Workflow Optimization

---

## 👨‍💻 Author

[Manasi Gandhi](https://manasigandhiportfolio.lovable.app/)

**Excel VBA | Automation | Macros | Process Automation | Productivity Tracking**
