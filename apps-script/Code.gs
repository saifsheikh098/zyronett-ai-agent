const T = {
  leads: 'Leads',
  projects: 'Projects',
  profit: 'Profit',
  dash: 'Dashboard',
  sop: 'SOPs'
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('ZyroNett')
    .addItem('Setup workspace', 'setupWorkspace')
    .addItem('Generate DM drafts', 'generateDmDrafts')
    .addItem('Email today follow-ups', 'sendFollowupDigest')
    .addItem('Enable daily 8 AM digest', 'installDailyDigestTrigger')
    .addToUi();
}

function setupWorkspace() {
  const ss = SpreadsheetApp.getActive();

  const defs = {
    Leads: [
      'Lead ID', 'Company / Agent', 'Contact Name', 'Role',
      'City / Market', 'Instagram URL', 'Website',
      'Public Business Email', 'Followers', 'Property Segment',
      'Lead Source', 'Fit Score (1-5)', 'Status',
      'Last Contact Date', 'Next Follow-up Date', 'Touch Count',
      'Personalized Observation', 'DM Draft', 'Reply Status',
      'Potential Monthly Value (₹)', 'Notes'
    ],
    Projects: [
      'Project ID', 'Client', 'Package', 'Start Date', 'Deadline',
      'Deliverables', 'Footage Received?', 'Assigned Editor',
      'Project Status', 'Revision Notes', 'Final Delivery Link',
      'Invoice Status', 'Next Action', 'Notes'
    ],
    Profit: [
      'Date', 'Client', 'Package', 'Reels Delivered',
      'Revenue (₹)', 'Editor Cost (₹)', 'Other Variable Costs (₹)',
      'Gross Profit (₹)', 'Profit Margin', 'Invoice Status',
      'Payment Date', 'Notes'
    ],
    Dashboard: ['Metric', 'Value'],
    SOPs: ['Stage', 'Founder-led SOP', 'Owner', 'Definition of Done']
  };

  Object.keys(defs).forEach(name => {
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    sh.getRange(1, 1, 1, defs[name].length)
      .setValues([defs[name]])
      .setFontWeight('bold')
      .setBackground('#172554')
      .setFontColor('#ffffff')
      .setWrap(true);
    sh.setFrozenRows(1);
    sh.setRowHeight(1, 40);
  });

  const leads = ss.getSheetByName('Leads');

  valid_(leads, 10, [
    'Residential', 'Ultra-luxury', 'Commercial',
    'Land/Plots', 'Developer Project', 'Other'
  ]);

  valid_(leads, 11, [
    'Instagram', 'Meta Ads Library', 'Website', 'Referral', 'Other'
  ]);

  valid_(leads, 13, [
    'New', 'Qualified', 'Contacted', 'Replied',
    'Call Booked', 'Proposal Sent', 'Won', 'Lost', 'Not a Fit'
  ]);

  valid_(leads, 19, [
    'Not contacted', 'No reply', 'Interested',
    'Not interested', 'Follow up later'
  ]);

  leads.getRange('L2:L501').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireNumberBetween(1, 5)
      .setAllowInvalid(false)
      .build()
  );

  const projects = ss.getSheetByName('Projects');
  valid_(projects, 7, ['Yes', 'No']);

  valid_(projects, 9, [
    'Not started', 'Awaiting footage', 'Editing',
    'Internal review', 'Client review', 'Revisions',
    'Delivered', 'Completed', 'On hold'
  ]);

  valid_(projects, 12, [
    'Not invoiced', 'Invoiced', 'Partially paid',
    'Paid', 'Overdue'
  ]);

  const profit = ss.getSheetByName('Profit');

  for (let r = 2; r <= 501; r++) {
    profit.getRange(r, 8)
      .setFormula(`=IF(E${r}="","",E${r}-F${r}-G${r})`);

    profit.getRange(r, 9)
      .setFormula(`=IFERROR(H${r}/E${r},0)`)
      .setNumberFormat('0.0%');
  }

  const dash = ss.getSheetByName('Dashboard');
  dash.clearContents();

  dash.getRange(1, 1, 12, 2).setValues([
    ['Metric', 'Value'],
    ['Lead names entered', '=COUNTIF(Leads!B2:B,"<>")'],
    ['Qualified', '=COUNTIF(Leads!M2:M,"Qualified")'],
    ['Contacted', '=COUNTIF(Leads!M2:M,"Contacted")'],
    ['Replies', '=COUNTIF(Leads!S2:S,"Interested")+COUNTIF(Leads!S2:S,"Not interested")+COUNTIF(Leads!S2:S,"Follow up later")'],
    ['Calls booked', '=COUNTIF(Leads!M2:M,"Call Booked")'],
    ['Proposals sent', '=COUNTIF(Leads!M2:M,"Proposal Sent")'],
    ['Clients won', '=COUNTIF(Leads!M2:M,"Won")'],
    ['Potential monthly pipeline (₹)', '=SUM(Leads!T2:T)'],
    ['Won monthly revenue (₹)', '=SUMIF(Leads!M2:M,"Won",Leads!T2:T)'],
    ['Follow-ups due today', '=COUNTIFS(Leads!O2:O,"<="&TODAY(),Leads!O2:O,">0",Leads!M2:M,"<>Won",Leads!M2:M,"<>Lost")'],
    ['Total gross profit (₹)', '=SUM(Profit!H2:H)']
  ]);

  dash.getRange(1, 1, 1, 2)
    .setFontWeight('bold')
    .setBackground('#172554')
    .setFontColor('#ffffff');

  const sop = ss.getSheetByName('SOPs');

  sop.getRange(2, 1, 8, 4).setValues([
    ['Lead sourcing', 'Find relevant real-estate agents, brokers and developers. Verify public business details.', 'Founder', 'Relevant profile and recent property activity checked.'],
    ['Qualification', 'Score activity, listing value, content gap and contactability from 1–5.', 'Founder', 'Specific observation recorded; prioritize scores 4–5.'],
    ['Outreach', 'Send up to 20 personalized messages daily; avoid generic bulk spam.', 'Founder', 'Contact date, status and next follow-up logged.'],
    ['Follow-up', 'Follow up politely after 3–4 days, then once after 5–7 days. Stop on opt-out.', 'Founder', 'Open leads have a next action and date.'],
    ['Proposal', 'Confirm deliverables, revisions, deadline and payment terms.', 'Founder', 'Written scope approved.'],
    ['Production', 'Brief editor with footage, references, hook and deadline.', 'Founder + editor', 'Editor confirms scope and ETA.'],
    ['Quality control', 'Check hook, pacing, color, text, spelling, audio, crop and export.', 'Founder', 'Final export passes quality checks.'],
    ['Delivery & profit', 'Record delivery, feedback, revenue, editor costs and other costs.', 'Founder', 'Next action and actual margin recorded.']
  ]);

  [leads, projects, profit, dash, sop].forEach(sh => {
    sh.setFrozenRows(1);
    sh.getDataRange().setVerticalAlignment('middle');
  });

  SpreadsheetApp.getUi().alert(
    'ZyroNett workspace ready! Add verified leads to the Leads tab.'
  );
}

function valid_(sheet, col, options) {
  sheet.getRange(2, col, 500, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(options, true)
      .setAllowInvalid(false)
      .build()
  );
}

function generateDmDrafts() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getActiveRange();

  if (sheet.getName() !== 'Leads' ||
      !range || range.getRow() < 2) {
    SpreadsheetApp.getUi().alert(
      'Leads tab mein lead ki row select karo.'
    );
    return;
  }

  for (
    let row = range.getRow();
    row < range.getRow() + range.getNumRows();
    row++
  ) {
    const a = sheet.getRange(row, 1, 1, 21).getValues()[0];

    const company = a[1] || 'there';
    const role = a[3] || 'real estate';
    const market = a[4] || 'your market';
    const observation = a[16] ||
      '[add a specific observation about their recent property video]';

    const message =
      `Hi ${company} — I came across your ${role} content in ${market}. ` +
      `${observation}. I run ZyroNett Studio and help real estate ` +
      `professionals turn property footage into polished cinematic reels. ` +
      `Would you be open to seeing a sample approach for one of your listings?`;

    sheet.getRange(row, 18).setValue(message);
  }

  SpreadsheetApp.getUi().alert(
    'DM drafts created. Review and personalize before sending.'
  );
}

function sendFollowupDigest() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Leads');

  if (!sheet) {
    throw new Error('Run setupWorkspace first.');
  }

  const values = sheet.getDataRange().getValues();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const due = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const date = row[14];
    const status = String(row[12] || '');

    if (
      date instanceof Date &&
      date <= today &&
      !['Won', 'Lost', 'Not a Fit'].includes(status)
    ) {
      due.push(
        `${row[0] || ''} | ${row[1] || '(company missing)'} | ` +
        `${status || 'New'} | ${row[5] || ''}`
      );
    }
  }

  const email = Session.getEffectiveUser().getEmail();

  if (!email) {
    throw new Error('Could not detect Google account email.');
  }

  MailApp.sendEmail(
    email,
    'ZyroNett Studio — Follow-ups Due Today',
    due.length
      ? `Follow-ups due (${due.length}):\n\n${due.join('\n')}`
      : 'Aaj koi follow-up due nahi hai.'
  );

  SpreadsheetApp.getUi().alert(
    `Digest email sent. Due leads: ${due.length}`
  );
}

function installDailyDigestTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'sendFollowupDigest')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('sendFollowupDigest')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();

  SpreadsheetApp.getUi().alert(
    'Daily digest enabled. Google may run it between 8–9 AM.'
  );
}
