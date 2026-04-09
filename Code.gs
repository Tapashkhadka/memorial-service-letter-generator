/**
 * Memorial Service Letter Generator
 * Google Apps Script Web App
 */

/**
 * Entry point - serves the web app
 */
function doGet(e) {
  var page = e && e.parameter && e.parameter.page;

  if (page === "program") {
    return HtmlService.createHtmlOutputFromFile("FuneralProgram")
      .setTitle("In Loving Memory of Challi Khadka Ghimiray")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("Memorial Service Letter Generator")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Returns the web app URL so client-side JS can build links
 */
function getAppUrl() {
  return ScriptApp.getService().getUrl();
}

/**
 * Generate a memorial service letter based on form data
 * @param {Object} formData - The form data from the client
 * @returns {Object} - Result with success status and letter content
 */
function generateLetter(formData) {
  try {
    const letter = createLetterContent(formData);
    return {
      success: true,
      letter: letter,
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString(),
    };
  }
}

/**
 * Create the letter content based on template and data
 * @param {Object} data - Form data
 * @returns {String} - Formatted letter HTML
 */
function createLetterContent(data) {
  const templates = {
    announcement: createAnnouncementLetter,
    thankYou: createThankYouLetter,
    invitation: createInvitationLetter,
    donation: createDonationLetter,
    workExcuse: createWorkExcuseLetter,
  };

  const templateFunc = templates[data.letterType] || createAnnouncementLetter;
  return templateFunc(data);
}

/**
 * Memorial Service Announcement Letter Template
 */
function createAnnouncementLetter(data) {
  const date = formatDate(data.serviceDate);
  const time = data.serviceTime || "Time to be announced";

  return `
    <div class="letter-header">
      <h1>Memorial Service Announcement</h1>
    </div>
    
    <div class="letter-body">
      <p class="date">${formatDate(new Date())}</p>
      
      <p>Dear Friends and Family,</p>
      
      <p>It is with heavy hearts that we announce the passing of <strong>${data.deceasedName}</strong>, 
      who left us on <strong>${formatDate(data.dateOfDeath)}</strong>${data.age ? " at the age of " + data.age : ""}.</p>
      
      <p>${data.bio || ""}</p>
      
      <div class="service-details">
        <h3>Memorial Service Details</h3>
        <p><strong>Date:</strong> ${date}</p>
        <p><strong>Time:</strong> ${time}</p>
        <p><strong>Location:</strong> ${data.location || "Location to be announced"}</p>
        ${data.address ? `<p><strong>Address:</strong> ${data.address}</p>` : ""}
      </div>
      
      ${
        data.reception
          ? `
      <div class="reception-details">
        <h3>Reception to Follow</h3>
        <p>${data.reception}</p>
      </div>
      `
          : ""
      }
      
      ${
        data.donationInfo
          ? `
      <div class="donation-info">
        <p>In lieu of flowers, donations may be made to: ${data.donationInfo}</p>
      </div>
      `
          : ""
      }
      
      <p>With love and remembrance,</p>
      <p class="signature">${data.senderName || "The Family"}</p>
      
      ${data.contactInfo ? `<p class="contact">Contact: ${data.contactInfo}</p>` : ""}
    </div>
  `;
}

/**
 * Thank You Letter Template
 */
function createThankYouLetter(data) {
  return `
    <div class="letter-header">
      <h1>Thank You</h1>
    </div>
    
    <div class="letter-body">
      <p class="date">${formatDate(new Date())}</p>
      
      <p>Dear ${data.recipientName || "Friends and Family"},</p>
      
      <p>We would like to express our sincere gratitude for your kindness and support during this difficult time 
      following the loss of our beloved <strong>${data.deceasedName}</strong>.</p>
      
      ${data.specificGift ? `<p>Your thoughtful gift of ${data.specificGift} meant so much to us.</p>` : ""}
      
      ${data.attendedService ? `<p>Thank you for attending the memorial service and sharing in our remembrance of ${data.deceasedName}.</p>` : ""}
      
      <p>Your condolences, prayers, and presence have been a great comfort to our family. We are deeply touched 
      by your thoughtfulness and care.</p>
      
      <p>With heartfelt appreciation,</p>
      <p class="signature">${data.senderName || "The Family"}</p>
    </div>
  `;
}

/**
 * Memorial Service Invitation Template
 */
function createInvitationLetter(data) {
  const date = formatDate(data.serviceDate);

  return `
    <div class="letter-header">
      <h1>Memorial Service Invitation</h1>
    </div>
    
    <div class="letter-body">
      <p>You are cordially invited to join us in celebrating the life of</p>
      
      <h2 class="deceased-name">${data.deceasedName}</h2>
      
      <p class="service-type">${data.serviceType || "Memorial Service"}</p>
      
      <div class="service-details">
        <p><strong>Date:</strong> ${date}</p>
        <p><strong>Time:</strong> ${data.serviceTime}</p>
        <p><strong>Location:</strong> ${data.location}</p>
        ${data.address ? `<p>${data.address}</p>` : ""}
      </div>
      
      ${data.specialRequests ? `<p class="special-requests"><em>${data.specialRequests}</em></p>` : ""}
      
      <p>We hope you can join us in honoring and remembering ${data.deceasedName}.</p>
      
      <p class="signature">${data.senderName || "The Family"}</p>
    </div>
  `;
}

/**
 * Donation Request Letter Template
 */
function createDonationLetter(data) {
  return `
    <div class="letter-header">
      <h1>In Memory of ${data.deceasedName}</h1>
    </div>
    
    <div class="letter-body">
      <p class="date">${formatDate(new Date())}</p>
      
      <p>Dear ${data.recipientName || "Friends and Family"},</p>
      
      <p>As we mourn the loss of <strong>${data.deceasedName}</strong>, we would like to honor their memory 
      by supporting a cause that was dear to their heart.</p>
      
      ${data.charityDescription ? `<p>${data.charityDescription}</p>` : ""}
      
      <div class="donation-details">
        <h3>Donation Information</h3>
        <p><strong>Charity/Organization:</strong> ${data.charityName}</p>
        ${data.charityAddress ? `<p><strong>Address:</strong> ${data.charityAddress}</p>` : ""}
        ${data.donationInstructions ? `<p><strong>Instructions:</strong> ${data.donationInstructions}</p>` : ""}
        ${data.onlineLink ? `<p><strong>Online:</strong> <a href="${data.onlineLink}">${data.onlineLink}</a></p>` : ""}
      </div>
      
      <p>Your donation in ${data.deceasedName}'s memory would mean a great deal to our family and would 
      continue their legacy of ${data.legacy || "giving and compassion"}.</p>
      
      <p>Thank you for considering this request.</p>
      
      <p class="signature">${data.senderName || "The Family"}</p>
    </div>
  `;
}

/**
 * Save letter to Google Drive
 * @param {String} letterHtml - The letter HTML content
 * @param {String} fileName - Name for the file
 * @returns {Object} - Result with file URL
 */
function saveToDrive(letterHtml, fileName) {
  try {
    const blob = Utilities.newBlob(
      letterHtml,
      MimeType.HTML,
      fileName + ".html",
    );
    const file = DriveApp.createFile(blob);

    return {
      success: true,
      fileUrl: file.getUrl(),
      fileId: file.getId(),
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString(),
    };
  }
}

/**
 * Send letter via email
 * @param {String} recipientEmail - Email address
 * @param {String} subject - Email subject
 * @param {String} letterHtml - Letter HTML content
 * @returns {Object} - Result status
 */
function sendEmail(recipientEmail, subject, letterHtml) {
  try {
    MailApp.sendEmail({
      to: recipientEmail,
      subject: subject,
      htmlBody: letterHtml,
    });

    return {
      success: true,
      message: "Email sent successfully",
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString(),
    };
  }
}

/**
 * Helper function to format dates
 * @param {String|Date} date - Date to format
 * @returns {String} - Formatted date string
 */
function formatDate(date) {
  if (!date) return "";
  const d = new Date(date);
  const options = { year: "numeric", month: "long", day: "numeric" };
  return d.toLocaleDateString("en-US", options);
}

/**
 * Funeral Work Excuse Letter Template - Professional Bereavement Leave Request
 */
function createWorkExcuseLetter(data) {
  const today = formatDate(new Date());
  const serviceDate = formatDate(data.serviceDate);
  const deathDate = formatDate(data.dateOfDeath || data.deathDate);
  const leaveStart = data.leaveStartDate
    ? formatDate(data.leaveStartDate)
    : serviceDate;
  const returnDate = data.returnDate
    ? formatDate(data.returnDate)
    : "to be determined";

  // Build employee info block
  const employeeInfo = [];
  if (data.employeeTitle) employeeInfo.push(data.employeeTitle);
  if (data.department) employeeInfo.push(data.department);
  if (data.employeeId) employeeInfo.push(`ID: ${data.employeeId}`);

  return `
    <div class="work-excuse-letter">
      <!-- Letter Header -->
      <div class="letter-header" style="border-bottom: 2px solid #333; padding-bottom: 20px; margin-bottom: 30px;">
        <h1 style="font-size: 24px; margin: 0 0 10px 0; color: #333;">Bereavement Leave Request</h1>
        <p style="margin: 0; color: #666; font-size: 14px;">Date: ${today}</p>
      </div>
      
      <!-- Recipient Block -->
      <div class="recipient-block" style="margin-bottom: 25px;">
        <p style="margin: 5px 0;"><strong>To:</strong> ${data.managerName || "Human Resources Department / Direct Manager"}</p>
        <p style="margin: 5px 0;"><strong>From:</strong> ${data.employeeName || data.signatoryName}</p>
        <p style="margin: 5px 0;"><strong>Subject:</strong> Request for Bereavement Leave - ${data.deceasedName}</p>
      </div>
      
      <!-- Salutation -->
      <p style="margin-top: 25px;">Dear ${data.managerName ? data.managerName.split(" ")[0] : "Manager"},</p>
      
      <!-- Opening Paragraph -->
      <p style="line-height: 1.6; margin: 15px 0;">
        I am writing to formally request bereavement leave following the passing of my ${data.relationship || "family member"}, 
        <strong>${data.deceasedName}</strong>, on ${deathDate}. I am deeply saddened by this loss and need time to 
        attend the funeral service, make necessary arrangements, and be with my family during this difficult period.
      </p>
      
      <!-- Leave Details Section -->
      <div class="leave-details" style="background-color: #f9f9f9; padding: 20px; margin: 20px 0; border-left: 4px solid #666;">
        <h3 style="margin-top: 0; color: #333; font-size: 16px;">Leave Request Details</h3>
        <table style="width: 100%; border-collapse: collapse;">
          <tr>
            <td style="padding: 8px 0; width: 40%;"><strong>Leave Start Date:</strong></td>
            <td style="padding: 8px 0;">${leaveStart}</td>
          </tr>
          <tr>
            <td style="padding: 8px 0;"><strong>Expected Return Date:</strong></td>
            <td style="padding: 8px 0;">${returnDate}</td>
          </tr>
          <tr>
            <td style="padding: 8px 0;"><strong>Number of Days:</strong></td>
            <td style="padding: 8px 0;">${data.daysRequested || "3-5"} days</td>
          </tr>
        </table>
      </div>
      
      <!-- Service Information -->
      ${
        serviceDate || data.serviceLocation
          ? `
      <div class="service-info" style="margin: 20px 0;">
        <h3 style="color: #333; font-size: 16px; margin-bottom: 10px;">Service Information</h3>
        ${serviceDate ? `<p style="margin: 5px 0;"><strong>Service Date:</strong> ${serviceDate}</p>` : ""}
        ${data.serviceLocation ? `<p style="margin: 5px 0;"><strong>Location:</strong> ${data.serviceLocation}</p>` : ""}
        ${data.serviceDetails ? `<p style="margin: 5px 0;">${data.serviceDetails}</p>` : ""}
      </div>
      `
          : ""
      }
      
      <!-- Work Coverage Plan -->
      <div class="coverage-plan" style="margin: 20px 0;">
        <h3 style="color: #333; font-size: 16px; margin-bottom: 10px;">Work Coverage During Absence</h3>
        <p style="line-height: 1.6; margin: 10px 0;">
          I have made arrangements to ensure continuity of my responsibilities during my absence:
        </p>
        <ul style="line-height: 1.6; margin: 10px 0;">
          ${data.coveragePlan ? `<li>${data.coveragePlan}</li>` : "<li>All pending tasks have been completed or delegated appropriately</li>"}
          <li>My email will be monitored for urgent matters ${data.contactDuringLeave ? `or you may contact me at ${data.contactDuringLeave}` : "if necessary"}</li>
          <li>I will provide a comprehensive handover document before my departure</li>
        </ul>
      </div>
      
      <!-- Closing Paragraph -->
      <p style="line-height: 1.6; margin: 20px 0;">
        I understand the importance of minimizing disruption to the team and have taken steps to ensure 
        a smooth transition during my absence. I will be available for any urgent matters that require 
        my immediate attention, though I appreciate your understanding that my availability may be limited.
      </p>
      
      <p style="line-height: 1.6; margin: 20px 0;">
        Please let me know if you require any additional documentation, such as a death certificate or 
        funeral program, to process this leave request. I am happy to provide any necessary paperwork 
        upon my return.
      </p>
      
      <p style="line-height: 1.6; margin: 20px 0;">
        Thank you for your understanding and support during this difficult time. I appreciate your 
        consideration of this request.
      </p>
      
      <!-- Closing -->
      <p style="margin-top: 30px;">Respectfully,</p>
      
      <!-- Signature Block -->
      <div class="signature-block" style="margin-top: 40px; margin-left: 0;">
        <p style="margin: 5px 0; font-size: 18px; font-family: 'Brush Script MT', cursive; color: #333;">
          ${data.employeeName || data.signatoryName}
        </p>
        <p style="margin: 5px 0; color: #666; font-size: 14px;">
          ${employeeInfo.join(" | ")}
        </p>
        ${data.phoneNumber ? `<p style="margin: 5px 0; color: #666; font-size: 14px;">Contact: ${data.phoneNumber}</p>` : ""}
      </div>
      
      <!-- Footer -->
      <div class="letter-footer" style="margin-top: 50px; padding-top: 20px; border-top: 1px solid #ddd; font-size: 12px; color: #999; text-align: center;">
        <p>This letter is submitted as a formal request for bereavement leave in accordance with company policy.</p>
      </div>
    </div>
  `;
}

/**
 * Get available letter templates
 * @returns {Array} - List of template options
 */
function getTemplates() {
  return [
    {
      id: "announcement",
      name: "Service Announcement",
      description: "Announce memorial service details to friends and family",
    },
    {
      id: "thankYou",
      name: "Thank You Letter",
      description: "Express gratitude for support and condolences",
    },
    {
      id: "invitation",
      name: "Service Invitation",
      description: "Formal invitation to memorial service",
    },
    {
      id: "donation",
      name: "Donation Request",
      description: "Request donations in lieu of flowers",
    },
    {
      id: "workExcuse",
      name: "Funeral Work Excuse",
      description: "Request bereavement leave from employer",
    },
  ];
}
