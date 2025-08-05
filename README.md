# Patient Visit Log System

A simple yet effective desktop database application built with Microsoft Access and VBA to manage and track patient visit records. This project was developed as a beginner-friendly portfolio piece to demonstrate core database management and automation skills.

## 📋 Project Overview

This application provides a user-friendly interface for a small clinic or a solo practitioner to log patient information and their visit details. The system simplifies data entry and allows for quick generation of a summary report, organized by patient, for easy tracking and analysis.

## 📸 Screenshots

*(**Trainer's Note:** Replace these links after you take your own screenshots! Create a folder named `screenshots` in your project directory and save your images there.)*

**1. Main Data Entry Form** - Clean and intuitive form for adding or viewing patient records.


**2. Patient Visit Summary Report** - A professional report grouping all visits by patient name.


## ✨ Key Features

* **Intuitive Data Entry:** A user-friendly form (`frmPatientEntry`) for seamless data input.
* **VBA-Powered Automation:** Form controls are automated with VBA for key functions:
    * **Add New Patient:** Clears the form for a new record.
    * **Save Record:** Commits the current record to the database.
    * **Close Form:** Safely closes the data entry form.
* **Structured Data Storage:** Utilizes a well-defined Access table (`tblPatientVisits`) to ensure data integrity.
* **Dynamic Summary Reporting:** Generates a clean, grouped, and sorted report (`rptPatientVisitSummary`) of all patient visits for easy review.

## 🛠️ Technology Stack

* **Database:** Microsoft Access
* **Automation/Logic:** VBA (Visual Basic for Applications)

## 🚀 Getting Started

To run this application on your local machine, please follow these steps.

### Prerequisites

* Microsoft Access (2013 or later) installed on your machine.

### Installation & Setup

1.  **Clone or Download the Repository:**
    * Click the `Code` button on GitHub and select `Download ZIP`.
    * Extract the ZIP file to a location on your computer.

2.  **Enable Content:**
    * Open the `HealthcareLog.accdb` file.
    * Microsoft Access may show a security warning yellow bar at the top. Click **"Enable Content"** to allow the VBA code to run. For a better experience, save the database in a "Trusted Location".

## 📖 How to Use the Application

1.  From the main Navigation Pane, double-click **`frmPatientEntry`** to open the data entry form.
2.  Click the **`Add New Patient`** button to prepare a blank form for a new entry.
3.  Fill in the patient's details (`First Name`, `Last Name`, `DOB`) and the visit information (`VisitDate`, `ReasonForVisit`).
4.  Click the **`Save Record`** button to save the information to the database.
5.  To view a summary of all records, open the **`rptPatientVisitSummary`** report from the Navigation Pane.
6.  When finished, click the **`Close Form`** button on the form.

## 👤 Author

* **[Ardhi Mahesh]**
* **LinkedIn:** `[linkedin.com/in/ardhi-mahesh-40b743122]`




---
