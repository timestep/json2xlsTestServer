'use strict'

var json2xls = require('json2xls');
var fs = require('fs');

var json = [
  {
    foo: 'bar',
    qux: 'moo',
    poo: 123,
    stux: new Date(),
    no: true,
    chikle: 'hwown'
  },
  {
    foo: 'nope',
    qux: 'qweqw',
    poo: 12323232,
    stux: new Date() + 1,
    no: true,
    chikle: 'hwown'
  },
  {
    foo: 'nope',
    qux: 'qweqw',
    poo: 12323232,
    stux: new Date() + 1,
    no: true,
    chikle: 'hwown'
  }
]

var xls = json2xls(json);

fs.writeFile('data.xlsx', xls, 'binary',function (err) {
  if (err) throw err;
  console.log('It\'s saved!');
});
