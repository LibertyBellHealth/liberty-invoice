// Regression: a STANDALONE MDHHS-6064-P (the "Provider Time and Task Management" form the ASW now
// sends on its own) never produced an effective date, and any ASW email containing a digit was
// dropped. Both were reported from the field on 2026-09-01; the line shapes below are taken from
// real forms, with the client/worker details replaced.
const test = require('node:test');
const assert = require('node:assert');
const { loadApp } = require('./harness');
const w = loadApp();

// Section 2 is a two-column table: labels on one line, values on the next.
const STANDALONE_6064 = [
  ['MDHHS - Adult Services', 'Liberty Home Care Assistance'],
  [
    'MDHHS-6064-P, PROVIDER TIME AND TASK MANAGEMENT',
    '(New 9-25)',
    'SECTION 1 - CLIENT INFORMATION',
    'Client Name Client ID Number',
    'Jane Doe 23518223',
    'County Name Case Number',
    '82-WAYNE 698084-1',
    'SECTION 2 - ADULT SERVICES INFORMATION',
    'Adult Services Worker (ASW) Name ASW Phone Number',
    'T Coleman 313-804-6694',
    'ASW Email Address Date',
    'ColemanT1@michigan.gov 08/06/2026',
    'SECTION 3 - TASKS',
    'Provider Name Provider Pay Rate',
    'Liberty Home Care Assistance $ 27.00',
    'Authorized Tasks Time / Day Number of Days Time /Month Amount',
    'Bathing 00:05 7 days per week 02:30 $67.72',
    'Medication 00:02 7 days per week 01:00 $27.09',
    'Total per month 03:30 $ 94.81',
  ],
];

// The full packet: DHS-1210-A cover letter carries the real service start date.
const FULL_PACKET = [
  ['MDHHS OFFICE'],
  [
    'DHS-1210-A, SERVICES AND PAYMENT APPROVAL NOTICE',
    'Adult Services Worker Email Telephone Number',
    'A Sawyer SawyerA2@michigan.gov 586-770-9560',
    'are approved for Home Help Services effective 03/01/2026. Services have been approved for 41 Hours',
    'and 16 Minutes per month.',
  ],
  [
    'MDHHS-6064-P, PROVIDER TIME AND TASK MANAGEMENT',
    // One row carrying the whole authorization, so the packet reconciles and warns about nothing.
    'Meal Preparation 01:20 7 days per week 41:16 $1,114.56',
    'Total per month 41:16 $ 1,114.56',
  ],
];

test('standalone MDHHS-6064: effective date is read from the value line, not the label', () => {
  const r = w.parseDHS1210(STANDALONE_6064);
  assert.strictEqual(r.formType, 'MDHHS-6064');
  assert.strictEqual(r.effectiveDate, '08/06/2026');
  // It is the ASW's signature date, not a stated start of service — say so rather than imply certainty.
  assert.strictEqual(r.effectiveDateGuessed, true);
  assert.ok(r.warnings.some(x => /effective date \(read from/.test(x)),
    'expected a verify-this warning, got: ' + JSON.stringify(r.warnings));
  // The old code emitted the bare "effective date" miss; that must be gone.
  assert.ok(!r.warnings.includes('effective date'));
});

test('standalone MDHHS-6064: hours still come from the "Total per month" row', () => {
  const r = w.parseDHS1210(STANDALONE_6064);
  assert.strictEqual(r.hours, 3);
  assert.strictEqual(r.minutes, 30);
  assert.strictEqual(r.tasks.length, 2);
});

test('ASW email containing a digit is captured', () => {
  assert.strictEqual(w.parseDHS1210(STANDALONE_6064).aswEmail, 'ColemanT1@michigan.gov');
  assert.strictEqual(w.parseDHS1210(FULL_PACKET).aswEmail, 'SawyerA2@michigan.gov');
});

test('full DHS-1210-A packet: the stated effective date wins and is NOT flagged as a guess', () => {
  const r = w.parseDHS1210(FULL_PACKET);
  assert.strictEqual(r.effectiveDate, '03/01/2026');
  assert.strictEqual(r.effectiveDateGuessed, undefined);
  assert.strictEqual(r.hours, 41);
  assert.strictEqual(r.minutes, 16);
  assert.strictEqual(JSON.stringify([...r.warnings]), '[]');
});

test('a form with no date at all still reports the miss', () => {
  const r = w.parseDHS1210([['MDHHS-6064-P, PROVIDER TIME AND TASK MANAGEMENT', 'Total per month 03:30 $ 94.81']]);
  assert.strictEqual(r.effectiveDate, undefined);
  assert.ok(r.warnings.includes('effective date'));
});
