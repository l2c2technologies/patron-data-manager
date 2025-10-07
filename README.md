# L2C2 Patron Data Manager for Google Sheets

A custom toolkit for managing pre-import status patron data for Koha ILS available in Google Sheet format. It provides a set of tools to clean, fix, organize, and export your data.

***IMPORTANT*** Always use this AppScript on a copy of your data and ***never*** directly on the master copy of your data. This code is shared on AS-IS basis without any assurance about it safety and usability. This was written as an in-house tool, thus YMMV.

---

## Get the Latest Version 🌐

The most up-to-date version of this script is always available on our GitHub repository.

* **Download from GitHub:** [https://github.com/l2c2technologies/patron-data-manager](https://github.com/l2c2technologies/patron-data-manager)

---

## What It Can Do ✨

Once installed, you'll get a new menu called **"L2C2 Patron Data Manager"** with these tools:

### 🧼 Cleaning Tools
* **Remove Line Breaks in Column:** Fixes a whole column by getting rid of messy line breaks and extra spaces.
* **Advanced Cleanup in Range:** A deep clean for a specific block of cells. It removes spaces before commas, extra spaces between words, and trims spaces from the beginning and end of cells.

### 🏗️ Structural Tools
* **Add Column with Preset Value:** Instantly adds a new column and fills every row with a default value you choose (like "Batch 2025").
* **Replicate Column:** Copies an entire column and its data into a new one.
* **Rename Column Header:** Lets you quickly rename any column.
* **Delete Column:** Deletes a column for good (it'll ask you first!).

### ✨ Transformation Tools
* **Conditional Population:** Fills out a column for you based on another column's value. For example, it can put "Regular Admission" in Column P for every row where Column N says "First Year".

### 🕵️‍♂️ Validation Tools
* **Find & Handle Duplicates:** Scans your columns for duplicate data. It's flexible: you can decide for each duplicate whether to **remove the whole row**, just **clear the cell**, or **skip it**. You can also tell it to apply one action to all duplicates at once.
* **Validate & Clean Mobile Numbers:** Checks a column for valid 10-digit Indian mobile numbers. It automatically formats good numbers and removes bad ones.
* **Validate & Clean Emails:** Checks emails for two things: 1) if the format is correct (like `name@domain.com`), and 2) if the domain is real and can receive mail. It removes any bad emails.

### 📝 Documentation Tools
* **Generate Koha Field Map:** An interactive tool that helps you document how your sheet's columns map to the official Koha patron import fields. It creates a new sheet called "Field Mapping" with the results.

### 📤 Export Tools
* **Export Filtered Data as CSV:** Lets you pull just the data you need. For example, you can export a CSV file of only the rows where the "Category" column is "BE-SC-ST".

### ❓ Help
* **About this Tool:** A simple guide inside the menu that explains what each function does.

**Note:** All actions that change your data are automatically recorded in a sheet named **"Action Log"** for full traceability.

---

## 🚀 How to Install It (One-Time Setup)

Adding this to your Google Sheet is easy. Just follow these steps once.

1.  **Get the Code:**
    Copy the entire Apps Script code from the latest version on GitHub or from the file provided.

2.  **Open Your Google Sheet:**
    Open the spreadsheet you want to add the tools to.

3.  **Open the Script Editor:**
    From the menu at the top, click **Extensions > Apps Script**.

4.  **Paste the Code:**
    A code editor window will open. Delete any code that's already there and paste in the entire script.

5.  **Save the Project:**
    Click the **💾 (Save project)** icon in the toolbar. Give the project a name if it asks, like "DataManager".

6.  **Refresh Your Sheet:**
    Go back to your spreadsheet tab and reload the page. **This is important!** The new menu won't show up until you do.

That's it! You will now see the **"L2C2 Patron Data Manager"** menu at the top of your sheet.

---

## 💡 How to Use the Tools

1.  Click on the **"L2C2 Patron Data Manager"** menu.
2.  Find the tool you need in the submenus (like "Cleaning Tools").
3.  Click the tool. A small box will pop up asking for info, like which column letter to work on (`A`, `B`, `C`...) or a cell range (`A2:D50`).
4.  Follow the simple prompts, and a confirmation message will appear when the job is done.

### A Quick Word on Permissions
The first time you use any tool, Google will pop up a window asking for your permission. This is a normal security step. Here’s what it needs and why:

* **Access this spreadsheet:** So it can actually read and change your data as you command.
* **Connect to the internet:** Only for the email validation tool, so it can check if a domain like `some-company.com` is real.
* **Access your Google Drive:** Only for the export tool, so it can save the CSV file you create.

---

## License and Author

* **Author:** Indranil Das Gupta <indradg@l2c2.co.in>
* **Copyright:** (c) 2023 - 2025 L2C2 Technologies
* **License:** AGPL v3+
