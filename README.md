# Inspection Report Automation using VBA

Automated the generation of Microsoft Word thermography inspection reports by inserting thermal and digital images based on image filenames using VBA.

##  Project Impact

-  Successfully automated reports containing **300+ images** in real-world use.
-  Reduced report preparation time from **3–5 hours to under 10 minutes**.
-  One-click automation for image insertion, resizing, and alignment.
-  Eliminated repetitive manual formatting and image placement.

---

## Overview

Preparing thermography inspection reports manually requires inserting, resizing, and aligning hundreds of thermal and digital images into Microsoft Word tables.

This project automates the entire workflow by reading image filenames from a predefined table, locating the corresponding images, inserting them into the report, resizing them while preserving aspect ratio, and aligning them within the designated table cells.

> **Note:** The demonstration included in this repository uses a small sample dataset (8–10 images). The same automation was successfully used to generate actual inspection reports containing **300+ images**.


# thermography-word-report-automation
Automating Microsoft Word thermography inspection reports by inserting thermal and digital images using VBA.

**Overview**

Preparing thermography inspection reports manually requires inserting, resizing, and aligning hundreds of thermal and digital images into Microsoft Word tables.

This project automates that entire process using Microsoft Word VBA.

The macro reads image filenames from a predefined table, searches the selected image folder, automatically inserts matching images, resizes them while preserving aspect ratio, and aligns them inside table cells.

Although the demonstration uses a small sample report, the automation was developed and successfully used to prepare actual inspection reports containing 300+ images, reducing report preparation time from several hours to only a few minutes.


**Features**

✔ Automatic thermal image insertion

✔ Automatic digital image insertion

✔ Folder picker

✔ Automatic JPG detection

✔ Preserve aspect ratio

✔ Fit images within table cells

✔ Center images automatically

✔ Detect missing images

✔ Support large reports (300+ images)


**Workflow**

Prepare Word Table (Digital Image cells 2.5/2" & Thermal Image cells 2/2")

↓

Enter Image File Names

↓

Run VBA Macro (Alt + F8)

↓

Select Image Folder (Give Path of Folder containing sample omages)

↓

Images Inserted Automatically (see magic happening :))

↓

Completed Inspection Report (Images inserted, now modify content, add header/footer as per your requirement)


**Technologies**

Microsoft Word
VBA (Visual Basic for Applications)
Microsoft Office 365


**Demo**

See: Thermography_Inspection_Report_Automation_Demo.1.mp4


**Project Structure**

Thermography-report-automation

│
├── README.md
├── VBA
│     InsertThermalDigitalImages.bas
│
├── Sample Report
│     SampleReport.docm
│
├── Demo
│     Thermography_Inspection_Report_Automation_Demo.1.mp4



**Results**

| Before | After |
| -------| ------|
|Manual image insertion|Automatic resizing|
|Automatic resizing|Automatic alignment|
|Hours of work|Few minutes|


**Notes**

The demonstration files contain sample data only. Code can be utilized efficiently for 300+ images.

