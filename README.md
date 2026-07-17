# Thermography Inspection Report Automation using VBA

Automated the generation of Microsoft Word thermography inspection reports by inserting thermal and digital images based on image filenames using VBA.

##  Project Impact

-  Successfully automated reports containing **300+ images** in real-world use.
-  Reduced report preparation time from **3–5 hours to under 10 minutes**.
-  One-click automation for image insertion, resizing, and alignment.
-  Eliminated repetitive manual formatting and image placement.
  

**Overview**

Preparing thermography inspection reports manually requires inserting, resizing, and aligning hundreds of thermal and digital images into Microsoft Word tables then manually numbering each row.

This project automates that entire process using Microsoft Word VBA.

The macro reads image filenames from a predefined table, searches the selected image folder, automatically inserts matching images, resizes them while preserving aspect ratio, and aligns them inside table cells.

Macro is provided for automatic serial numbering also. It was used for numbering 200+ rows of report.

Although the demonstration uses a small sample report, the automation was developed and successfully used to prepare actual inspection reports containing 300+ images, reducing report preparation time from several hours to only a few minutes.


**Features**

✔ Automatic thermal image insertion

✔ Automatic digital image insertion

✔ Automatic image resizing and alignment

✔ Automatic serial numbering for inspection tables

✔ Missing image detection

✔ One-click report generation

✔ Supports inspection reports with 300+ images

✔ Preserves image aspect ratio

✔ Compatible with Microsoft Word 365


## Workflow

Prepare the Word table
- Set the required cell dimensions (Thermal Image: **2.5" × 1.88"**, Digital Image: **2" × 1.88"**).
- Enter the Thermal and Digital image filenames in the respective columns.

↓

(Optional) Renumber the table
- Press **Alt + F11** → **Insert → Module** → import **AutoSerialNumbering.bas**.
- Run **AutoSerialNumbering.bas** to automatically update serial numbers after adding or deleting rows.

↓

Run the image insertion macro
- Press **Alt + F11** → **Insert → Module** → import **InsertThermalDigitalImages.bas**.
- Press **Alt + F8** and execute **InsertThermalDigitalImages**.


↓

Select the image folder
- Browse to the folder containing the thermal and digital **.jpg** images.

↓

Automatic report generation
- The macro locates matching images, inserts them into the corresponding table cells, resizes them, preserves aspect ratio, centers them, and reports any missing files.

↓

Finalize the report
- Review the report, update observations and recommendations, and customize the header, footer, and formatting as required.


**Technologies**

Microsoft Word
VBA (Visual Basic for Applications)
Microsoft Office 365


**Demo**

See: Thermography_Inspection_Report_Automation_Demo.1.mp4


## Project Structure

Thermography-report-automation

```text

Inspection-Report-Automation
│
├── README.md
├── LICENSE
│
├── VBA
│     ├── InsertThermalDigitalImages.bas
│     └── AutoSerialNumbering.bas
│
├── Sample Report
│     └── SampleReport.docm
│
├── Demo
│     ├── Thermography_Inspection_Report_Automation_Demo.mp4
│     └── Auto_Serial_Numbering_Inside_Report_Demo.mp4
│
│
└── Screenshots
      ├── Before.png
      └── After.png

```


## Results

| Before | After |
| -------| ------|
|Manual image insertion|Automatic resizing|
|Automatic resizing|Automatic alignment|
|Manual serial numbering|Automatic serial number insertion|
|Hours of work|Few minutes|



**Notes**

The demonstration files contain sample data only. Code can be utilized efficiently for 300+ images.

