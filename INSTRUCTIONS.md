# Newsletter System Setup Instructions

This guide will help you set up a complete newsletter system using Google Sheets, Google Forms, Google Docs, and a custom script.

## Step 1: Create the Google Sheet

1.  Create a new Google Sheet. name it something like **"My Newsletter System"**.
2.  At the bottom, rename the first tab (sheet) to **`Subscribers`**.
3.  Click the `+` button to add a new tab and rename it to **`Campaigns`**.

**Important:** The names must be exactly `Subscribers` and `Campaigns` (case-sensitive) for the script to work.

## Step 2: Create the Subscriber Form

1.  In your Google Sheet, go to **Tools > Create a new form**.
2.  Name the form **"Subscriber Sign-up"**.
3.  Add the following questions:
    *   **Name** (Short answer) - *Required*
    *   **Email** (Short answer) - *Required*
    *   **Phone** (Short answer) - *Optional*
    *   **Address** (Paragraph) - *Optional*
    *   **Category** (Dropdown) - *Required* (Essential for targeting)
        *   Option 1: **VIP**
        *   Option 2: **General**
        *   Option 3: **Government officials**
4.  Once created, go back to your Google Sheet. You will see a new tab named **"Form Responses 1"**.
5.  **Rename this tab to `Subscribers`**.
    *   *Note: If you already renamed the first tab to `Subscribers` in Step 1, delete the empty one and rename the form responses tab to `Subscribers`.*
6.  Ensure the columns in the `Subscribers` sheet are in this order (drag columns if needed):
    *   Column A: Timestamp
    *   Column B: Name
    *   Column C: Email
    *   Column D: Phone
    *   Column E: Address
    *   Column F: Category
    *   *Note: If your form added them in a different order, just ensure `Name` is column B, `Email` is column C, and `Category` is column F. If they are different, you must update the `COL_` numbers in the script.*

## Step 3: Create the Newsletter Authoring Form

1.  Create another new Google Form (directly from Drive or duplicate the previous one).
2.  Name it **"Newsletter Authoring Form"**.
3.  Add the following questions:
    *   **Subject** (Short answer) - *Required*
    *   **Google Doc Link** (Short answer) - *Required* (Paste the shareable link of your newsletter Doc here)
    *   **Target Category** (Dropdown) - *Required* (Use the exact same options: VIP, General, Government officials)
4.  Go to the **Responses** tab in the Form editor.
5.  Click the **Google Sheets icon** (Create Spreadsheet).
6.  Select **"Select existing spreadsheet"**.
7.  Choose your **"My Newsletter System"** sheet (from Step 1).
8.  Go back to your Google Sheet. You will see a new tab (e.g., "Form Responses 2").
9.  **Rename this tab to `Campaigns`**.
    *   *Note: If you created an empty `Campaigns` tab in Step 1, delete the empty one first, then rename the form responses tab to `Campaigns`.*
10. Ensure the columns in the `Campaigns` sheet are in this order:
    *   Column A: Timestamp
    *   Column B: Subject
    *   Column C: Google Doc Link
    *   Column D: Target Category
    *   Column E: (Leave empty, the script will write "Sent" here)

## Step 4: Install the Script

1.  Open your Google Sheet ("My Newsletter System").
2.  Go to **Extensions > Apps Script**.
3.  Delete any code currently in the script editor (e.g., `function myFunction() {...}`).
4.  Copy the entire code from the `GOOGLE_APPS_SCRIPT.js` file provided in this repository.
5.  Paste it into the script editor.
6.  Click the **Save** icon (floppy disk). Name the project "Newsletter Script".
7.  **Reload the Google Sheet page**.
8.  After a few seconds, you should see a new menu item at the top called **"Newsletter Manager"**.

## Step 5: Create Newsletter Content

1.  Create a standard Google Doc.
2.  Write your newsletter content.
    *   You can use **Bold**, *Italic*, and standard text formatting.
    *   You can insert **Inline Images** directly into the doc.
    *   **Personalization:** Use `{{Name}}` anywhere in the doc to insert the subscriber's name automatically.
        *   Example: "Dear {{Name}}, welcome to our newsletter!"
3.  **Share settings:** Ensure the Google Doc is accessible. (Usually "Anyone with the link can view" or just ensuring your account has access is enough).

## Step 6: Sending the Newsletter

1.  **Draft a Campaign:**
    *   Open your **"Newsletter Authoring Form"**.
    *   Enter the **Subject Line**.
    *   Paste the **Google Doc Link** (from Step 5).
    *   Select the **Target Category**.
    *   Submit the form.
2.  **Verify:**
    *   Go to your Google Sheet.
    *   Check the `Campaigns` tab. You should see a new row with your submission.
3.  **Send:**
    *   Click **Newsletter Manager** in the menu bar.
    *   Select **Send Pending Newsletter**.
    *   The script will ask for permission the first time. Authorize it.
    *   The script will read the last campaign, find subscribers in that category, and send the emails.
    *   Once finished, it will mark the campaign as "Sent" in Column E.

## Troubleshooting

*   **"Error: Missing sheets":** Check that your tabs are named exactly `Subscribers` and `Campaigns`.
*   **"No campaigns found":** Submit a response to the Newsletter Authoring Form.
*   **Emails not sending:** Check if you have subscribers in the selected Category.
*   **Images not showing:** Ensure images are pasted directly into the Google Doc (not linked from external sites).
