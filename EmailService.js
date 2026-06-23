/**
 * =========================================================
 * EMAIL SERVICE MODULE (EmailService.gs)
 * =========================================================
 * Manages all email-driven automation:
 * - Parsing recruitment confirmation emails
 * - Sending candidate welcome/magic-link emails
 * - Document rejection notifications
 * - Package submission notifications
 * - Escalation alerts for overdue items
 * =========================================================
 */

const RECRUITMENT_SENDER_DOMAIN = '@yourcompany.com'; // Filter emails from this domain only
const HR_COORDINATOR_EMAIL_DEFAULT = 'hr.coordinator@yourcompany.com';
const HR_OPS_HEAD_EMAIL = 'hr.ops.head@yourcompany.com';
const FROM_NAME = 'Enterprise HR Mobilization System';
const SYSTEM_EMAIL = Session.getActiveUser().getEmail();

// ─────────────────────────────────────────────
// EMAIL PARSER (Time-Driven Trigger)
// ─────────────────────────────────────────────

/**
 * Triggered automatically by a Google Apps Script time-driven trigger.
 * Scans for unread recruitment confirmation emails and processes them.
 *
 * To activate: In Apps Script editor go to Triggers > Add Trigger 
 * > function: checkRecruitmentEmails, Event: Time-driven, Every 1 hour.
 */
function checkRecruitmentEmails() {
  const threads = GmailApp.search('is:unread subject:"Candidate Accepted Offer" from:' + RECRUITMENT_SENDER_DOMAIN);

  threads.forEach(thread => {
    const message = thread.getMessages()[0];
    const body = message.getPlainBody();

    // Extract candidate data from structured email body
    const parsed = parseRecruitmentEmail_(body);
    if (!parsed.success) {
      Logger.log('Email parse failed: ' + message.getSubject());
      // Email admin about parse failure
      GmailApp.sendEmail(
        SYSTEM_EMAIL,
        '⚠️ HR System: Email Parse Failure',
        'Could not parse recruitment email. Subject: ' + message.getSubject() + '\n\nManual entry required.'
      );
      return;
    }

    // Create candidate and send welcome email
    const result = api_createCandidate(parsed.data);
    if (result.success) {
      api_createCandidateFolder(result.candidateId, parsed.data.fullName);
      sendCandidateWelcomeEmail_(parsed.data.email, parsed.data.fullName, result.candidateId);
      thread.markRead();
      Logger.log('Candidate auto-created: ' + parsed.data.fullName);
    }
  });
}

/**
 * Parses a structured recruitment confirmation email body.
 * The recruitment team must follow this exact format in their emails.
 *
 * Expected format example:
 *   CANDIDATE_NAME: Ahmed Ali
 *   POSITION: Civil Engineer
 *   DEPARTMENT: Projects
 *   EMAIL: ahmed.ali@gmail.com
 *   PHONE: +201012345678
 *   NATIONALITY: Egyptian
 *   SALARY_OMR: 350
 */
function parseRecruitmentEmail_(body) {
  try {
    const extract = (key) => {
      const match = body.match(new RegExp(key + ':\\s*(.+)', 'i'));
      return match ? match[1].trim() : null;
    };

    const data = {
      fullName: extract('CANDIDATE_NAME'),
      position: extract('POSITION'),
      department: extract('DEPARTMENT'),
      email: extract('EMAIL'),
      phone: extract('PHONE'),
      nationality: extract('NATIONALITY'),
      salary: extract('SALARY_OMR'),
      coordinatorEmail: HR_COORDINATOR_EMAIL_DEFAULT
    };

    // Validate required fields
    if (!data.fullName || !data.email || !data.position) {
      return { success: false, error: 'Missing required fields in email.' };
    }
    return { success: true, data: data };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ─────────────────────────────────────────────
// EMAIL SENDERS
// ─────────────────────────────────────────────

/**
 * Public wrapper for the frontend to trigger the candidate welcome email.
 * @param {string} email
 * @param {string} fullName
 * @param {string} candidateId
 */
function api_sendWelcomeEmail(email, fullName, candidateId) {
  try {
    sendCandidateWelcomeEmail_(email, fullName, candidateId);
    return { success: true };
  } catch (e) {
    Logger.log('api_sendWelcomeEmail error: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * Sends a welcome email with a secure magic link to the candidate.
 * The magic link contains the candidate's UUID as a query parameter.
 * @param {string} toEmail
 * @param {string} candidateName
 * @param {string} candidateId
 */
function sendCandidateWelcomeEmail_(toEmail, candidateName, candidateId) {
  Logger.log({
    toEmail: toEmail,
    candidateName: candidateName,
    candidateId: candidateId,
    emailType: typeof toEmail
  });

  if (typeof toEmail !== 'string') {
    throw new Error(
      'sendCandidateWelcomeEmail_: Expected string email but received: ' +
      JSON.stringify(toEmail)
    );
  }

  // Get the published Web App URL
  const appUrl = ScriptApp.getService().getUrl();
  const magicLink = appUrl + '?view=candidatePortal&cid=' + candidateId;

  const subject = 'Welcome to the Oman Mobilization Process - Action Required';
  const htmlBody = `
    <div style="font-family: Segoe UI, sans-serif; max-width: 600px; margin: 0 auto; padding: 24px; border: 1px solid #edebe9;">
      <h2 style="color: #0f6abe;">Welcome, ${candidateName}!</h2>
      <p>Congratulations on accepting your offer. To begin your Oman mobilization process, you must upload the following required documents.</p>
      <ul>
        <li>Passport Copy (valid, clear scan)</li>
        <li>Personal Photo (white background)</li>
        <li>Graduation Certificate (certified)</li>
        <li>Medical Examination Report</li>
        <li>Medical Analysis Report</li>
      </ul>
      <p><a href="${magicLink}" style="background:#0f6abe; color:#fff; padding: 12px 24px; text-decoration:none; border-radius:4px; display:inline-block; margin-top:16px;">Upload My Documents →</a></p>
      <p style="color:#888; font-size:12px; margin-top:24px;">This link is unique to you. Do not share it. Expires in 48 hours.</p>
    </div>
  `;

  GmailApp.sendEmail(toEmail, subject, '', {
    name: FROM_NAME,
    htmlBody: htmlBody
  });
}

/**
 * Notifies a candidate that a document was rejected and requires re-upload.
 * @param {string} candidateEmail
 * @param {string} candidateName
 * @param {string} docType
 * @param {string} reason
 * @param {string} candidateId
 */
function api_sendRejectionEmail(candidateEmail, candidateName, docType, reason, candidateId) {
  const appUrl = ScriptApp.getService().getUrl();
  const magicLink = appUrl + '?view=candidatePortal&cid=' + candidateId;

  const subject = 'Action Required: Please Re-upload Your ' + docType;
  const htmlBody = `
    <div style="font-family: Segoe UI, sans-serif; max-width: 600px; margin: 0 auto; padding: 24px; border: 1px solid #edebe9;">
      <h2 style="color: #a4262c;">Document Update Required</h2>
      <p>Dear ${candidateName},</p>
      <p>Your <strong>${docType}</strong> document needs to be re-uploaded for the following reason:</p>
      <blockquote style="border-left: 4px solid #d83b01; margin: 16px 0; padding: 8px 16px; background: #fff4ce;">${reason}</blockquote>
      <p><a href="${magicLink}" style="background:#0f6abe; color:#fff; padding: 12px 24px; text-decoration:none; border-radius:4px; display:inline-block; margin-top:16px;">Re-upload Document →</a></p>
    </div>
  `;

  GmailApp.sendEmail(candidateEmail, subject, '', {
    name: FROM_NAME,
    htmlBody: htmlBody
  });
}

/**
 * Notifies the HR Ops Section Head that a package is ready for final audit.
 * @param {string} candidateName
 * @param {string} coordinatorName
 */
function api_sendPackageSubmissionAlert(candidateName, coordinatorName) {
  const subject = '📦 HR System: New Package Ready for Final Audit';
  const body = `
    Dear HR Operations Team,\n\nA new candidate package is ready for your final review.\n\n
    Candidate: ${candidateName}\nSubmitted by: ${coordinatorName}\n\n
    Please log in to the HR Mobilization System to review and approve.\n\nRegards,\nHR System`;

  GmailApp.sendEmail(HR_OPS_HEAD_EMAIL, subject, body, { name: FROM_NAME });
}

/**
 * Escalation alert for overdue candidates (pending > 7 days).
 * Designed to be run via a weekly time-driven trigger.
 */
function checkOverdueCandidates() {
  const result = api_getAllCandidates();
  if (!result.success) return;

  const sevenDaysAgo = new Date(Date.now() - 7 * 24 * 60 * 60 * 1000);

  const overdue = result.data.filter(c => {
    const updated = new Date(c.UpdatedAt);
    return updated < sevenDaysAgo &&
      c.CurrentStatus !== 'Mobilized' &&
      c.CurrentStatus !== 'Closed';
  });

  if (overdue.length === 0) return;

  const list = overdue.map(c => `- ${c.FullName} (${c.CurrentStatus})`).join('\n');
  GmailApp.sendEmail(HR_OPS_HEAD_EMAIL,
    `⚠️ HR Alert: ${overdue.length} Overdue Mobilization Cases`,
    `The following candidates have not progressed in over 7 days:\n\n${list}`,
    { name: FROM_NAME }
  );
}
