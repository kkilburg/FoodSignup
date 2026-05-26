/**
 * CBS Food Signup builder.
 *
 * How to use:
 * 1) In script.google.com create a new Apps Script project.
 * 2) Paste this file into Code.gs.
 * 3) Update CONFIG.adminEmail if desired.
 * 4) Run buildCbsFoodSignupSystem().
 *
 * The script creates:
 * - A Google Spreadsheet with tabs: Events, Slots, Signups, Helper_OpenSlots, Public_View.
 * - A Google Form linked to Signups sheet with required fields and edit-after-submit.
 */

const CONFIG = {
  timezone: 'America/Chicago',
  defaultEventTitle: 'Weekly CBS Night',
  spreadsheetNamePrefix: 'CBS Food Signup',
  formNamePrefix: 'CBS Food Signup',
  adminEmail: Session.getActiveUser().getEmail() || '',
  sampleEvents: [
    { id: 'evt_2026_09_10', date: '2026-09-10', title: 'Weekly CBS Night', description: '' },
    { id: 'evt_2026_09_17', date: '2026-09-17', title: 'Weekly CBS Night', description: '' }
  ],
  sampleSlots: [
    { id: 'slt_evt_2026_09_10_main_1', event_id: 'evt_2026_09_10', category: 'Main', item_name: 'Pasta', max_signups: 1 },
    { id: 'slt_evt_2026_09_10_side_1', event_id: 'evt_2026_09_10', category: 'Side', item_name: 'Salad', max_signups: 1 },
    { id: 'slt_evt_2026_09_10_dessert_1', event_id: 'evt_2026_09_10', category: 'Dessert', item_name: '', max_signups: 1 },
    { id: 'slt_evt_2026_09_17_main_1', event_id: 'evt_2026_09_17', category: 'Main', item_name: 'Soup', max_signups: 1 },
    { id: 'slt_evt_2026_09_17_drinks_1', event_id: 'evt_2026_09_17', category: 'Drinks', item_name: '', max_signups: 2 }
  ]
};

function buildCbsFoodSignupSystem() {
  const now = Utilities.formatDate(new Date(), CONFIG.timezone, 'yyyy-MM-dd HH:mm');
  const spreadsheet = SpreadsheetApp.create(`${CONFIG.spreadsheetNamePrefix} (${now})`);
  spreadsheet.setSpreadsheetTimeZone(CONFIG.timezone);

  const eventsSheet = getOrCreateSheet_(spreadsheet, 'Events');
  const slotsSheet = getOrCreateSheet_(spreadsheet, 'Slots');
  const signupsSheet = getOrCreateSheet_(spreadsheet, 'Signups');
  const helperSheet = getOrCreateSheet_(spreadsheet, 'Helper_OpenSlots');
  const publicViewSheet = getOrCreateSheet_(spreadsheet, 'Public_View');

  setupEventsSheet_(eventsSheet);
  setupSlotsSheet_(slotsSheet);
  setupSignupsSheet_(signupsSheet);
  setupHelperSheet_(helperSheet);
  setupPublicViewSheet_(publicViewSheet);

  seedEvents_(eventsSheet);
  seedSlots_(slotsSheet);
  applySlotsFormulas_(slotsSheet);
  applyEventsFormula_(eventsSheet);

  const form = FormApp.create(`${CONFIG.formNamePrefix} (${now})`);
  form.setDescription('Sign up to bring food for an upcoming CBS event. No account required.');
  form.setAllowResponseEdits(true);
  form.setCollectEmail(true);
  form.setConfirmationMessage('Thanks for signing up! You can use your edit link to update your response later.');

  // Link form responses to this spreadsheet.
  form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheet.getId());

  setupFormQuestions_(form);
  // Default simple dropdown mode (v1):
  refreshFormChoicesFromSheets_();
  // Optional upgraded UX mode (v2): uncomment next line to build event-specific sections.
  // buildSectionBasedFormFromSheets_();

  protectCalculatedColumns_(eventsSheet, slotsSheet, signupsSheet);
  addConditionalFormatting_(slotsSheet);

  if (CONFIG.adminEmail) {
    MailApp.sendEmail({
      to: CONFIG.adminEmail,
      subject: 'CBS Food Signup system created',
      htmlBody: `Your CBS system is ready.<br><br>Spreadsheet: <a href="${spreadsheet.getUrl()}">${spreadsheet.getUrl()}</a><br>Form (edit): <a href="${form.getEditUrl()}">${form.getEditUrl()}</a><br>Form (live): <a href="${form.getPublishedUrl()}">${form.getPublishedUrl()}</a>`
    });
  }

  Logger.log('Spreadsheet: ' + spreadsheet.getUrl());
  Logger.log('Form (live): ' + form.getPublishedUrl());
  Logger.log('Form (edit): ' + form.getEditUrl());

  return {
    spreadsheetUrl: spreadsheet.getUrl(),
    formUrl: form.getPublishedUrl(),
    formEditUrl: form.getEditUrl()
  };
}

function refreshFormChoicesFromSheets_() {
  const form = FormApp.getActiveForm();
  const ss = SpreadsheetApp.getActive();
  const events = readRows_(ss.getSheetByName('Events'));
  const slots = readRows_(ss.getSheetByName('Slots'));

  const eventLabels = events
    .filter(r => r.id)
    .map(r => r.display_label || `${r.date} — ${r.title || CONFIG.defaultEventTitle}`);

  const openSlots = slots
    .filter(r => r.id && String(r.status).toUpperCase() !== 'FULL')
    .map(r => {
      const eventLabel = events.find(e => e.id === r.event_id)?.display_label || r.event_id;
      const baseSlot = r.item_name ? `${r.category} — ${r.item_name}` : r.category;
      return `${eventLabel} | ${baseSlot} | ${r.id}`;
    });

  const items = form.getItems();
  const eventItem = items.find(i => i.getTitle() === 'Event Date');
  const slotItem = items.find(i => i.getTitle() === 'Food Slot');

  if (!eventItem || !slotItem) {
    throw new Error('Form is missing expected Event Date or Food Slot fields.');
  }

  eventItem.asListItem().setChoiceValues(eventLabels.length ? eventLabels : ['No events configured']);
  slotItem.asListItem().setChoiceValues(openSlots.length ? openSlots : ['No open slots available']);
}


/**
 * Upgraded UX mode (v2): rebuild the form into event-specific sections.
 * Users pick event first, then are routed to a section showing only that event's open slots.
 */
function buildSectionBasedFormFromSheets_() {
  const form = FormApp.getActiveForm();
  const ss = SpreadsheetApp.getActive();
  const events = readRows_(ss.getSheetByName('Events')).filter(r => r.id);
  const slots = readRows_(ss.getSheetByName('Slots')).filter(r => r.id && String(r.status).toUpperCase() !== 'FULL');

  resetFormKeepingBasics_(form);

  form.addTextItem().setTitle('Name').setRequired(true);
  form.addTextItem().setTitle('Email').setRequired(true).setHelpText('Use the best email to receive updates.');

  const eventPicker = form.addMultipleChoiceItem().setTitle('Event Date').setRequired(true)
    .setHelpText('Choose your CBS date; you will then pick a food slot for that date.');

  const sectionByEventId = {};
  const eventChoiceMap = [];

  // Build one section per event with only its open slots.
  events.forEach(evt => {
    const page = form.addPageBreakItem().setTitle(evt.display_label || `${evt.date} — ${evt.title || CONFIG.defaultEventTitle}`);
    sectionByEventId[evt.id] = page;

    const openForEvent = slots.filter(s => s.event_id === evt.id);
    const slotQ = form.addListItem().setTitle('Food Slot').setRequired(true)
      .setHelpText('Pick one available item.');

    const options = openForEvent.map(s => {
      const base = s.item_name ? `${s.category} — ${s.item_name}` : s.category;
      return `${base} | ${s.id}`;
    });

    slotQ.setChoiceValues(options.length ? options : ['No open slots available for this date']);
    form.addParagraphTextItem().setTitle('Notes').setRequired(false);

    eventChoiceMap.push({ label: evt.display_label || `${evt.date} — ${evt.title || CONFIG.defaultEventTitle}`, page });
  });

  const submitPage = form.addPageBreakItem().setTitle('Submit');

  const choices = eventChoiceMap.map(({label, page}) => eventPicker.createChoice(label, page));
  if (choices.length === 0) {
    eventPicker.setChoiceValues(['No events configured']);
  } else {
    eventPicker.setChoices(choices);
  }

  // Ensure every event section submits after notes.
  const items = form.getItems();
  for (let i = 0; i < items.length; i++) {
    if (items[i].getType() === FormApp.ItemType.PAGE_BREAK && items[i].asPageBreakItem().getTitle() !== 'Submit') {
      items[i].asPageBreakItem().setGoToPage(submitPage);
    }
  }
}

function resetFormKeepingBasics_(form) {
  const items = form.getItems();
  for (let i = items.length - 1; i >= 0; i--) {
    form.deleteItem(items[i]);
  }
}
function setupFormQuestions_(form) {
  form.addTextItem().setTitle('Name').setRequired(true);
  form.addTextItem().setTitle('Email').setRequired(true).setHelpText('Use the best email to receive updates.');
  form.addListItem().setTitle('Event Date').setRequired(true);
  form.addListItem().setTitle('Food Slot').setRequired(true);
  form.addParagraphTextItem().setTitle('Notes').setRequired(false);
}

function setupEventsSheet_(sheet) {
  sheet.clear();
  const headers = ['id', 'date', 'title', 'description', 'display_label'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader_(sheet, headers.length);
}

function setupSlotsSheet_(sheet) {
  sheet.clear();
  const headers = ['id', 'event_id', 'category', 'item_name', 'max_signups', 'slot_label', 'current_signup_count', 'open_spots', 'status'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader_(sheet, headers.length);
}

function setupSignupsSheet_(sheet) {
  sheet.clear();
  const headers = ['form_timestamp', 'name', 'email', 'slot_id', 'notes', 'signup_id', 'event_id', 'created_at', 'updated_at'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader_(sheet, headers.length);
}

function setupHelperSheet_(sheet) {
  sheet.clear();
  sheet.getRange('A1').setValue('Open slots only (auto-updated)');
  sheet.getRange('A2').setFormula('=FILTER(Slots!A:I, Slots!I:I="OPEN")');
}

function setupPublicViewSheet_(sheet) {
  sheet.clear();
  sheet.getRange('A1').setValue('Read-only public slot status');
  sheet.getRange('A2').setFormula('=QUERY({Events!A:E,Slots!A:I},"select Col2,Col3,Col8,Col7,Col9 where Col6 = Col1 label Col2 \"date\", Col3 \"title\", Col8 \"slot\", Col7 \"open_spots\", Col9 \"status\"")');
}

function seedEvents_(sheet) {
  const rows = CONFIG.sampleEvents.map(e => [e.id, e.date, e.title, e.description, '']);
  if (rows.length) sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

function seedSlots_(sheet) {
  const rows = CONFIG.sampleSlots.map(s => [s.id, s.event_id, s.category, s.item_name, s.max_signups, '', '', '', '']);
  if (rows.length) sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

function applyEventsFormula_(sheet) {
  const last = Math.max(sheet.getLastRow(), 2);
  sheet.getRange(2, 5, last - 1, 1).setFormulaR1C1('=IF(RC[-3]="","",TEXT(RC[-3],"yyyy-mm-dd")&" — "&IF(RC[-2]="","CBS Night",RC[-2]))');
}

function applySlotsFormulas_(sheet) {
  const last = Math.max(sheet.getLastRow(), 2);
  sheet.getRange(2, 6, last - 1, 1).setFormulaR1C1('=IF(RC[-2]="",RC[-3],RC[-3]&" — "&RC[-2])');
  sheet.getRange(2, 7, last - 1, 1).setFormulaR1C1('=COUNTIF(Signups!C4,RC[-6])');
  sheet.getRange(2, 8, last - 1, 1).setFormulaR1C1('=MAX(0,RC[-3]-RC[-1])');
  sheet.getRange(2, 9, last - 1, 1).setFormulaR1C1('=IF(RC[-1]<=0,"FULL","OPEN")');
}

function protectCalculatedColumns_(eventsSheet, slotsSheet, signupsSheet) {
  [eventsSheet.getRange('E:E'), slotsSheet.getRange('F:I'), signupsSheet.getRange('F:I')].forEach(rng => {
    const protection = rng.protect();
    protection.setWarningOnly(true);
  });
}

function addConditionalFormatting_(slotsSheet) {
  const dataRange = slotsSheet.getRange('A2:I');
  const rules = [
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$I2="FULL"').setBackground('#f4cccc').setRanges([dataRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$H2=1').setBackground('#fff2cc').setRanges([dataRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$H2>1').setBackground('#d9ead3').setRanges([dataRange]).build()
  ];
  slotsSheet.setConditionalFormatRules(rules);
}

function styleHeader_(sheet, width) {
  const header = sheet.getRange(1, 1, 1, width);
  header.setFontWeight('bold').setBackground('#d9e1f2');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, width);
}

function getOrCreateSheet_(spreadsheet, name) {
  const existing = spreadsheet.getSheetByName(name);
  return existing || spreadsheet.insertSheet(name);
}

function readRows_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(h => String(h).trim());
  return values.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}
