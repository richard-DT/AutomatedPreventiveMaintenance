# Automated Maintenance System [Developed by: Richard De Los Trinos]

##  Overview

This project is an end-to-end preventive maintenance monitoring system built using Excel VBA. It combines data visualization, automation, and user interaction to help track machine health, manage maintenance activities, and reduce manual monitoring.

---

##  Objectives

* Monitor overall and per-section machine health
* Automate daily maintenance alerts for items due within 1 week
* Provide an interactive dashboard for users
* Allow users to update and track maintenance activities through drill-down interaction
* Auto-refresh data every minute for real-time monitoring

---

## System Architecture

### 1. Dashboard (UI Layer)

* Displays machine health status using color indicators

  * 🔴 Critical / Due for maintenance
  * 🟢 Normal
* Interactive sections that users can click (drill-down)
* Auto-refreshes every minute for real-time updates

### 2. Drill-Down Functionality with User Input

* Users can click a section to view:

  * List of machines
  * Pending maintenance items
* Further clicking on a machine or task opens an **input form** where users can:

  * Mark maintenance as completed
  * Input details of work performed
  * Update records directly
* This creates a **multi-level drill-down with interactive user input** (summary → pending items → maintenance update)

### 3. Maintenance Tracking (CRUD Operations)

* Tracks both pending and completed tasks
* Integrates directly with user input from the drill-down UI

### 4. Automation (RPA-like Behavior)

* System automatically:

  * Analyzes historical data
  * Determines maintenance schedules
  * Sends **daily email notifications** for upcoming or pending maintenance within one week

### 5. Decision Logic

* Rule-based system to:

  * Identify due maintenance within a week
  * Assign machine health status
  * Trigger alerts

---

## Technologies Used

* Microsoft Excel
* VBA (Visual Basic for Applications)
* Outlook (for email automation)

---

## Key Features

* Interactive dashboard with real-time status (auto-refresh every minute)
* **Multi-level drill-down navigation with user input**
* Maintenance tracking with data input/update capability
* Automated **daily email alert system** for due and pending maintenance
* Historical data analysis for maintenance prediction

---

## Impact

* Reduced manual monitoring effort
* Improved maintenance scheduling accuracy
* Faster response to critical issues
* Centralized system for tracking maintenance activities in real-time

---

## Screenshots

Include screenshots from your PowerPoint:

1. Main Dashboard (overall status)
2. Section Drill-Down View
3. Maintenance Details Form / User Input Form
4. Daily Email Notification Example

---

## Future Improvements

* Integration with external dashboards (e.g., Power BI)
* User access control and logging
* Migration to a web-based system for wider accessibility

---

## Developer Overview

I developed a preventive maintenance monitoring system using Excel VBA that features an interactive dashboard with multi-level drill-down functionality. Users can view machine health, click into critical sections to see pending maintenance tasks, and further click on tasks to update completion status and work details. The system auto-refreshes data every minute and sends **daily email alerts** for due or pending maintenance within one week based on historical data analysis.

---

