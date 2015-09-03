'use strict'

/*
These tests use the test spreadsheet accessible at
https://docs.google.com/spreadsheets/d/148tpVrZgcc-ReSMRXiQaqf9hstgT8HTzyPeKx6f399Y/edit#gid=0

In order to allow other devs to test both read and write funcitonality,
the doc must be public read/write which means if someone feels like it,
they could mess up the sheet which would mess up the tests.
Please don't do that...
*/

var GoogleSpreadsheet = require('../index.js')
var doc = new GoogleSpreadsheet('148tpVrZgcc-ReSMRXiQaqf9hstgT8HTzyPeKx6f399Y')
var creds = require('./test-creds')
var sheet
var test = require('tape')
// var Q = require('bluebird')

test('get info', function (t) {
  t.plan(2)
  doc.getInfo().then(function (info) {
    // even with public read/write, I think sheet author should stay constant
    t.equal(info.author.email, 'theozero@gmail.com', 'can read sheet info from google doc')

    sheet = info.worksheets[0]
    t.equal(sheet.title, 'Sheet1', 'can read sheet names from doc')

    t.end()
  })
})

test('check init auth', function (t) {
  doc.useServiceAccountAuth(creds).then(function () {
    t.end()
  })
})

test('clear sheet', function (t) {
  sheet.getRows().then(function (rows) {
    // these must be cleared one at a time or errors occur
    function nextRow () {
      if (!rows.length) return
      return rows.pop().del().then(function () { return nextRow() })
    }

    return nextRow()
  }).then(function () {
    t.end()
  })
})

test('check delete', function (t) {
  t.plan(1)

  sheet.getRows().then(function (rows) {
    t.equal(rows.length, 0, 'sheet should be empty after delete calls')
    t.end()
  })
})

test('basic write and read', function (t) {
  t.plan(2)

  // NOTE -- key and val are arbitrary headers.
  // These are the column headers in the first row of the spreadsheet.
  sheet.addRow({col1: 'test-col1', col2: 'test-col2'}).then(function () {
    return sheet.getRows()
  }).then(function (rows) {
    t.equal(rows[0].col1, 'test-col1', 'newly written value should match read value')
    t.equal(rows[0].col2, 'test-col2', 'newly written value should match read value')
  })
})

test('check newlines read', function (t) {
  t.plan(2)

  sheet.addRow({col1: 'Newline\ntest', col2: 'Double\n\nnewline test'}).then(function () {
    return sheet.getRows()
  }).then(function (rows) {
    // this was an issue before with an older version of xml2js
    t.ok(rows[1].col1.indexOf('\n') > 0, 'newline is read from sheet')
    t.ok(rows[1].col2.indexOf('\n\n') > 0, 'double newline is read from sheet')
  })
})
// TODO - test cell based feeds
