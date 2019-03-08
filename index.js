const express = require('express');
const app = express();
const path = require('path');
const docx = require('docx');
const fs = require('fs');
var numeral = require('numeral');
const bodyParser = require("body-parser");
const router = express.Router();

//include the routes file
var demand2 = require('./routes/demand2');
var demand1 = require('./routes/demand1');
var overduecc = require('./routes/overduecc');
var suspensioncc = require('./routes/suspensioncc');
var prelistingcc = require('./routes/prelistingcc');
var prelisting = require('./routes/prelisting');
var postlistingunsecured = require('./routes/postlistingunsecured');
var postlistingsecured = require('./routes/postlistingsecured');
var day40 = require('./routes/day40');
var day90 = require('./routes/day90');

////////
app.use('/demand2', demand2);
app.use('/demand1', demand1);
app.use('/overduecc', overduecc);
app.use('/suspensioncc', suspensioncc);
app.use('/prelistingcc', prelistingcc);
app.use('/prelisting', prelisting);
app.use('/postlistingunsecured', postlistingunsecured);
app.use('/postlistingsecured', postlistingsecured);
app.use('/day40', day40);
app.use('/day90', day90);


router.get('/', function (req, res) {
  res.json({ message: 'hooray! welcome to our rest video api!' }); 
});

app.use(bodyParser.urlencoded({
  extended: true
}));


//add the router
app.use('/', router);
app.listen(process.env.port || 8002);

console.log('Running at Port 8002');