const express = require('express');
var multer = require('multer')();
const path = require('path');
const app = express();
const port = process.env.PORT || 3000;
const main = require('./src/script');

app.use(multer.any());
app.use(express.static('./static/'));

app.post('/api/process', function(req, res) {
    res.setHeader('Content-Disposition', 'attachment; filename=result.zip');
    main(req, res);
});

app.listen(port, () => console.log(`Excel splitter is running on http://localhost:${port}`));
