const express = require('express');
var multer = require('multer')();
const path = require('path');
const app = express();
const main = require('./script');
const port = 3000;

app.use(multer.any());

app.get('*', function(req, res) {
    res.sendFile(path.join(__dirname + '/index.html'));
});

app.post('/', function(req, res) {
    main(req, res);
});

app.listen(port, () => console.log(`Excel splitter is running on http://localhost:${port}`));