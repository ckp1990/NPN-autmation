/**
 * GOOGLE APPS SCRIPT FOR NEWSLETTER AUTOMATION
 *
 * INSTRUCTIONS:
 * 1. Open your Google Sheet.
 * 2. Go to Extensions > Apps Script.
 * 3. Delete any code there and paste this entire script.
 * 4. Save the project.
 * 5. Reload your Google Sheet to see the "Newsletter Manager" menu.
 */

// CONFIGURATION
const SUBSCRIBERS_SHEET_NAME = "Subscribers"; // Name of the sheet with subscriber data
const CAMPAIGNS_SHEET_NAME = "Campaigns";     // Name of the sheet with newsletter drafts
// Column Indices (1-based)
// Ensure these match your Form Response columns
const COL_SUB_NAME = 2;      // Column B: Name
const COL_SUB_EMAIL = 3;     // Column C: Email
// Note: If you have "Phone" and "Address" columns in between, adjust this index.
// Based on instructions: A=Timestamp, B=Name, C=Email, D=Phone, E=Address, F=Category
const COL_SUB_CATEGORY = 6;  // Column F: Category (VIP, General, Government officials)

const COL_CAMP_SUBJECT = 2;  // Column B: Subject
const COL_CAMP_DOC_LINK = 3; // Column C: Google Doc Link
const COL_CAMP_CATEGORY = 4; // Column D: Target Category
const COL_CAMP_STATUS = 5;   // Column E: Status (Script will write here)

/**
 * Creates a custom menu in the Google Sheet.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Newsletter Manager')
    .addItem('Send Pending Newsletter', 'sendNewsletter')
    .addToUi();
}

/**
 * Main function to send the newsletter.
 * It looks for the last submission in the "Campaigns" sheet.
 */
function sendNewsletter() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const subscribersSheet = ss.getSheetByName(SUBSCRIBERS_SHEET_NAME);
  const campaignsSheet = ss.getSheetByName(CAMPAIGNS_SHEET_NAME);

  if (!subscribersSheet || !campaignsSheet) {
    ui.alert("Error: Missing sheets. Please ensure you have 'Subscribers' and 'Campaigns' sheets.");
    return;
  }

  // Get the last campaign submission
  const lastRow = campaignsSheet.getLastRow();
  if (lastRow < 2) {
    ui.alert("No campaigns found.");
    return;
  }

  const status = campaignsSheet.getRange(lastRow, COL_CAMP_STATUS).getValue();
  if (status === "Sent") {
    const response = ui.alert("The last campaign is already marked as 'Sent'. Do you want to send it again?", ui.ButtonSet.YES_NO);
    if (response !== ui.Button.YES) return;
  }

  // Read Campaign Details
  const subject = campaignsSheet.getRange(lastRow, COL_CAMP_SUBJECT).getValue();
  const docLink = campaignsSheet.getRange(lastRow, COL_CAMP_DOC_LINK).getValue();
  const targetCategory = campaignsSheet.getRange(lastRow, COL_CAMP_CATEGORY).getValue();

  if (!docLink || !subject) {
    ui.alert("Error: Missing Subject or Doc Link in the last campaign row.");
    return;
  }

  // Extract Doc ID from Link
  const docId = getIdFromUrl(docLink);
  if (!docId) {
    ui.alert("Error: Invalid Google Doc Link.");
    return;
  }

  // Convert Doc to HTML
  let emailContent;
  try {
    emailContent = convertGoogleDocToHtml(docId);
  } catch (e) {
    ui.alert("Error accessing Google Doc: " + e.message);
    return;
  }

  // Get Subscribers
  const subscribers = subscribersSheet.getDataRange().getValues();
  // Remove header row
  subscribers.shift();

  let sentCount = 0;

  // Iterate through subscribers
  for (let i = 0; i < subscribers.length; i++) {
    const row = subscribers[i];
    const name = row[COL_SUB_NAME - 1]; // Array is 0-indexed
    const email = row[COL_SUB_EMAIL - 1];
    const category = row[COL_SUB_CATEGORY - 1];

    // Filter by Category
    if (category === targetCategory && email) {
      // Personalize Content
      // Replace {{Name}} with actual name
      // We do a global replace for {{Name}}
      let personalizedHtml = emailContent.html.replace(/{{Name}}/g, name || "Subscriber");

      try {
        MailApp.sendEmail({
          to: email,
          subject: subject,
          htmlBody: personalizedHtml,
          inlineImages: emailContent.images
        });
        sentCount++;
      } catch (e) {
        console.error("Failed to send to " + email + ": " + e.message);
      }
    }
  }

  // Update Status
  campaignsSheet.getRange(lastRow, COL_CAMP_STATUS).setValue("Sent");
  campaignsSheet.getRange(lastRow, COL_CAMP_STATUS).setBackground("#d9ead3"); // Light green

  ui.alert("Newsletter sent to " + sentCount + " subscribers in category '" + targetCategory + "'.");
}

/**
 * Extracts the file ID from a Google Docs URL.
 */
function getIdFromUrl(url) {
  const match = url.match(/[-\w]{25,}/);
  return match ? match[0] : null;
}

/**
 * Converts a Google Doc to HTML and extracts inline images.
 * Returns object: { html: string, images: object }
 */
function convertGoogleDocToHtml(docId) {
  const doc = DocumentApp.openById(docId);
  const body = doc.getBody();
  const numChildren = body.getNumChildren();
  let html = "";
  const images = {};

  // Simple parser for Paragraphs, Lists, and Images
  for (let i = 0; i < numChildren; i++) {
    const child = body.getChild(i);
    const type = child.getType();

    if (type === DocumentApp.ElementType.PARAGRAPH) {
      html += processParagraph(child, images);
    } else if (type === DocumentApp.ElementType.LIST_ITEM) {
      // Simplified list handling: convert to bullet points
      html += processListItem(child, images);
    } else if (type === DocumentApp.ElementType.TABLE) {
      // Basic table support could be added here
      html += "<p>[Table content not supported in this simple script]</p>";
    }
  }

  return { html: html, images: images };
}

function processParagraph(paragraph, images) {
  const numChildren = paragraph.getNumChildren();
  let pContent = "";

  // If paragraph is empty, return a break
  if (numChildren === 0) return "<br/>";

  for (let i = 0; i < numChildren; i++) {
    const child = paragraph.getChild(i);
    const type = child.getType();

    if (type === DocumentApp.ElementType.TEXT) {
      let text = child.getText();
      // Handle basic formatting
      if (child.isBold()) text = "<b>" + text + "</b>";
      if (child.isItalic()) text = "<i>" + text + "</i>";
      if (child.getLinkUrl()) text = "<a href='" + child.getLinkUrl() + "'>" + text + "</a>";

      pContent += text;
    } else if (type === DocumentApp.ElementType.INLINE_IMAGE) {
      const blob = child.getBlob();
      const imageId = "image" + Object.keys(images).length;
      images[imageId] = blob;
      pContent += "<img src='cid:" + imageId + "' style='max-width:100%;' />";
    }
  }

  // Handle Paragraph Alignment
  let style = "";
  const align = paragraph.getAlignment();
  if (align === DocumentApp.HorizontalAlignment.CENTER) style = "text-align: center;";
  if (align === DocumentApp.HorizontalAlignment.RIGHT) style = "text-align: right;";

  return "<p style='" + style + "'>" + pContent + "</p>";
}

function processListItem(listItem, images) {
  const numChildren = listItem.getNumChildren();
  let liContent = "";

  for (let i = 0; i < numChildren; i++) {
    const child = listItem.getChild(i);
    const type = child.getType();

    if (type === DocumentApp.ElementType.TEXT) {
      let text = child.getText();
      if (child.isBold()) text = "<b>" + text + "</b>";
      if (child.isItalic()) text = "<i>" + text + "</i>";
      liContent += text;
    }
  }
  return "<div>&bull; " + liContent + "</div>";
}
