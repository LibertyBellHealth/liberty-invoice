'use strict';
// Agency info (editable provider details auto-filled onto state forms). Stored in lhca_agency,
// merged over defaults so an older/partial save still yields every field.
const { test, beforeEach } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage } = require('./harness');

const w = loadApp();
beforeEach(() => resetStorage(w));

test('getAgencyInfo: returns the defaults when nothing is saved', () => {
  const a = w.getAgencyInfo();
  assert.strictEqual(a.agency_provider_name, 'Thomas Jaboro');
  assert.strictEqual(a.agency_provider_id, '6221933');
  assert.strictEqual(a.agency_state, 'MI');
});

test('saveAgencyInfo -> getAgencyInfo round-trips edits', () => {
  w.saveAgencyInfo({
    agency_provider_name: 'New Name LLC', agency_provider_id: '9999999',
    agency_address: '1 Main St', agency_city: 'Detroit', agency_state: 'MI',
    agency_zip: '48201', agency_phone: '(313) 555-1212',
  });
  const a = w.getAgencyInfo();
  assert.strictEqual(a.agency_provider_name, 'New Name LLC');
  assert.strictEqual(a.agency_provider_id, '9999999');
  assert.strictEqual(a.agency_city, 'Detroit');
});

test('getAgencyInfo: a partial saved object still fills every field from defaults', () => {
  w.localStorage.setItem('lhca_agency', JSON.stringify({ agency_provider_name: 'Only Name' }));
  const a = w.getAgencyInfo();
  assert.strictEqual(a.agency_provider_name, 'Only Name', 'saved value used');
  assert.strictEqual(a.agency_provider_id, '6221933', 'missing field falls back to default');
  assert.strictEqual(a.agency_zip, '48314', 'missing field falls back to default');
});
