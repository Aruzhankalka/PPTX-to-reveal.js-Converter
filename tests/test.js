const { validate } = require('../src/ir/validator');
const fs = require('fs');
const path = require('path');

const fixture = JSON.parse(fs.readFileSync(path.join(__dirname, 'fixtures/minimal-ir.json'), 'utf8'));
const result = validate(fixture);

console.log('Validating minimal-ir.json...');
if (result.valid) {
  console.log('PASS - Valid IR document');
} else {
  console.log('FAIL - Invalid:');
  console.log(JSON.stringify(result.errors, null, 2));
}

// Negative test 1: missing required slideset
console.log('\nValidating broken doc (missing slideset)...');
const broken = { foo: 'bar' };
const r2 = validate(broken);
console.log(r2.valid ? 'FAIL - Should have been rejected' : 'PASS - Correctly rejected');

// Negative test 2: media missing required file-link
console.log('\nValidating broken doc (media without file-link)...');
const broken2 = JSON.parse(JSON.stringify(fixture));
delete broken2.slideset.slides[1].contents.media[0]['file-link'];
const r3 = validate(broken2);
console.log(r3.valid ? 'FAIL - Should have been rejected' : 'PASS - Correctly rejected');
if (!r3.valid) console.log('  Errors:', r3.errors.map(e => e.message).join('; '));

// Positive test: Sprint 2 fields shouldn't break Sprint 1 validation
console.log('\nValidating doc with Sprint 2 fields (shapes, master)...');
const sprint2ish = JSON.parse(JSON.stringify(fixture));
sprint2ish.slideset.master = { 'aspect-ratio': '16:9' };
sprint2ish.slideset.slides[0].contents.shapes = [{ id: 's1', type: 'rectangle' }];
const r4 = validate(sprint2ish);
console.log(r4.valid ? 'PASS - Forward-compatible' : 'FAIL: ' + JSON.stringify(r4.errors));

// Negative test 3: invalid enum value
console.log('\nValidating broken doc (invalid media-type)...');
const broken3 = JSON.parse(JSON.stringify(fixture));
broken3.slideset.slides[1].contents.media[0]['media-type'] = 'audio';
const r5 = validate(broken3);
console.log(r5.valid ? 'FAIL - Should have been rejected' : 'PASS - Correctly rejected');
if (!r5.valid) console.log('  Errors:', r5.errors.map(e => `${e.instancePath} ${e.message}`).join('; '));
