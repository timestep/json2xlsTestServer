'use strict'

var json2xls = require('json2xls');
var fs = require('fs');
var express = require('express');
var app = express();

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

app.use(json2xls.middleware);

app.get('/', function (req, res) {

  var xls = json2xls(json);
  var fileName = 'data.xlsx'

  fs.writeFile(fileName, xls, 'binary',function (err) {
    if (err) throw err;
    console.log('It\'s saved!');

    res.sendFile(fileName, {root:'/Users/kevin/Documents/json2xls/'}, function (err) {
        if (err) {
          console.log(err);
          return res.status(err.status).end();
        }
        else {
          return console.log('Sent:', fileName);
        }
      });
  });

});

app.get('/stream',function(req,res) {
  return res.xls('data.xlsx', json); // this is the ideal solution as it doesn't require to save on the server
});

var server = app.listen(3000, function () {

  var host = server.address().address;
  var port = server.address().port;

  console.log('Example app listening at http://%s:%s', host, port);

});
