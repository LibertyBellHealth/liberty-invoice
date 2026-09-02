// DHS-4771 (FICA withholding authorisation) and BPHASA-2421 (live-in caregiver EVV exemption) were
// retired from the Forms UI and their maps deleted. Owner asked for both back, 2026-09-01. Both
// templates are AcroForms, so they fill by field name — the risk is filling the WRONG box, which
// these tests pin. Field names below were read out of the shipped PDFs.
const test = require('node:test');
const assert = require('node:assert');
const { loadApp } = require('./harness');
const w = loadApp();

const DICT = {
  client_name: 'Jane Doe', client_first_name: 'Jane', client_last_name: 'Doe',
  medicaid_id: '1234567', client_county: '50-MACOMB', case_number: '',
  client_address: '1 Client St', client_city: 'Sterling Heights', client_state: 'MI', client_zip: '48314',
  worker_name: 'A Sawyer', worker_phone: '586-770-9560', today_date: '09/01/2026',
  caregiver_first_name: 'Sam', caregiver_last_name: 'Carer',
  caregiver_address: '2 Carer Ave', caregiver_city: 'Warren', caregiver_state: 'MI', caregiver_zip: '48089',
  caregiver_email: 'sam@example.com', caregiver_phone: '313-555-0100', caregiver_champs_id: '6221933',
};
const fill = (type, field) => w._matchAcroFormField(field, DICT, type);

test('both forms are registered and point at the shipped templates', () => {
  assert.strictEqual(w.STATE_FORM_OVERLAYS.dhs4771.file, '/forms/DHS-4771.pdf');
  assert.strictEqual(w.STATE_FORM_OVERLAYS.bphasa2421.file, '/forms/BPHASA-2421.pdf');
  // Neither is auto-signed: the client signs the 4771, the caregiver the 2421.
  assert.strictEqual(w.STATE_FORM_OVERLAYS.dhs4771.signature, null);
  assert.strictEqual(w.STATE_FORM_OVERLAYS.bphasa2421.signature, null);
});

test('DHS-4771: the address block is the CLIENT\'s, not the agency\'s', () => {
  assert.strictEqual(fill('dhs4771', 'Address'), '1 Client St');
  assert.strictEqual(fill('dhs4771', 'City'), 'Sterling Heights');
  assert.strictEqual(fill('dhs4771', 'Zip Code'), '48314');
  // The beneficiary signs it, so the printed name is theirs.
  assert.strictEqual(fill('dhs4771', 'Printed Name'), 'Jane Doe');
});

test('DHS-4771: caseworker fields carry the caseworker, and Client ID the Medicaid ID', () => {
  assert.strictEqual(fill('dhs4771', 'Adult Services Worker ASW'), 'A Sawyer');
  assert.strictEqual(fill('dhs4771', 'ASW Telephone Number'), '586-770-9560');
  assert.strictEqual(fill('dhs4771', 'Client ID'), '1234567');
  assert.strictEqual(fill('dhs4771', 'County'), '50-MACOMB');
});

test('an explicitly mapped field with no value stays BLANK instead of being guessed', () => {
  // case_number is empty in the dict; the old matcher fell through to keyword rules here.
  assert.strictEqual(fill('dhs4771', 'Case Number'), '');
});

test('BPHASA-2421: Section 1 is the caregiver, Section 2 is the beneficiary', () => {
  assert.strictEqual(fill('bphasa2421', 'First NameRow1'), 'Sam');
  assert.strictEqual(fill('bphasa2421', 'Last NameRow1'), 'Carer');
  assert.strictEqual(fill('bphasa2421', 'Street AddressRow1'), '2 Carer Ave');
  assert.strictEqual(fill('bphasa2421', 'CHAMPS Provider ID NumberRow1'), '6221933');
  assert.strictEqual(fill('bphasa2421', 'First NameRow1_2'), 'Jane');
  assert.strictEqual(fill('bphasa2421', 'Last NameRow1_2'), 'Doe');
  assert.strictEqual(fill('bphasa2421', 'Street AddressRow1_2'), '1 Client St');
  assert.strictEqual(fill('bphasa2421', 'Medicaid ID NumberRow1'), '1234567');
});

test('BPHASA-2421: the two address blocks never collapse into one', () => {
  assert.notStrictEqual(fill('bphasa2421', 'Street AddressRow1'), fill('bphasa2421', 'Street AddressRow1_2'));
  assert.notStrictEqual(fill('bphasa2421', 'CityRow1'), fill('bphasa2421', 'CityRow1_2'));
  assert.notStrictEqual(fill('bphasa2421', 'Zip CodeRow1'), fill('bphasa2421', 'Zip CodeRow1_2'));
});

test('no checkbox is auto-ticked on either form', () => {
  assert.strictEqual(w.STATE_FORM_CHECKS.dhs4771, undefined);
  assert.strictEqual(w.STATE_FORM_CHECKS.bphasa2421, undefined);
  // The DHS-390's one allowed tick is untouched.
  assert.strictEqual(w.STATE_FORM_CHECKS.dhs390['Home Help'], true);
});

test('the existing forms still map as before', () => {
  assert.strictEqual(w._matchAcroFormField('Full Name', DICT, 'dhs390'), 'Jane Doe');
  assert.strictEqual(w._matchAcroFormField('Client Name', DICT, 'msa4676'), 'Jane Doe');
  assert.strictEqual(w._matchAcroFormField('Agency Provider Name', DICT, 'msa4676'), '');
});
