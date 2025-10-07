🚀 Features

✅ Activity Submission Form

Google Apps Script HTML form for entering new activity records.

Automatically saves entries to the Activity sheet.

Includes fields for:

Date Start / End

Time Start / End

Title

Target Participants

Office Facilitator (with color code)

Remarks

Link

✅ Color Code Automation

Each division in the Division sheet has a unique color code.

When you select a division in the form, its corresponding color code is stored and can color the activity cell in the main sheet.

✅ Auto Archive Process

A scheduled script runs weekly to:

Check the Date End column in the Activity sheet.

Automatically move all past activities (those with end date before today) to ActivityArchive.

Preserve all data, including the color code.

✅ Weekly Auto Trigger

Runs automatically once per week (e.g., every Sunday at 2:00 AM).

You can adjust the schedule anytime using Apps Script Triggers.

🧩 Sheets Overview
Sheet Name	Purpose
Form Activity	Stores the HTML form inputs if using bound mode (optional).
Activity	Main sheet where current and upcoming activities are stored.
ActivityArchive	Stores past activities automatically.
Division	Lookup sheet containing Division names and color codes.

Example (Division Sheet):

Division Name	Color Code
Administrative and Finance Division	#E6B0AA
Agribusiness and Marketing Assistance Division	#AED6F1
Field Operations Division	#82E0AA
Planning, Monitoring and Evaluation Division	#F9E79F
...	...
🧠 Script Summary
1️⃣ saveToActivityFromHtml(data)

Stores submitted data from the form into the Activity sheet and applies the corresponding color.

2️⃣ movePastActivitiesToArchive()

Checks all rows in Activity and moves past events to ActivityArchive automatically.

3️⃣ getOffices()

Fetches division names and color codes for the dropdown in the HTML form.

4️⃣ whenhaveValue(e)

Automatically updates cell background color in Activity when the Office Facilitator column changes.

⏱ Setting Up Auto Trigger

Go to Extensions → Apps Script

Click the clock icon (Triggers) → Add Trigger

Choose:

Function: movePastActivitiesToArchive

Event Source: Time-driven

Type: Week timer

Day of week: Sunday (or your preferred day)

Time of day: 2 AM

Save ✅

🗂 Example Data Format
Date Start	Time Start	Date End	Time End	Title	Target Participants	Color	Office Facilitator	Status	Remarks	Link
Jan 1, 2025	8:00 AM	Jan 3, 2025	5:00 PM	Example	All PA, All MAO/CAO, FOD Rice	#81C744	Field Operations Division	In Progress	—	PDF Link
🧰 Tech Used

Google Sheets

Google Apps Script

Bootstrap 5.3 (Form UI)

HTML + JavaScript
