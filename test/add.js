const code = import('../Code.js');

if (typeof QUnit !== 'undefined') {
QUnit.module('code');

   QUnit.test("another example", function( assert ) {
    assert.equal(10, 210);
   });
}

