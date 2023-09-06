const express = require('express');
const path = require('path');
const bodyParser = require('body-parser');
const mainRouter = require('./Routers/main');

const app = express();

app.use(bodyParser.urlencoded({ extended: false }));
app.use(mainRouter);

app.listen(3000, () => {
  console.log('Sever is listening to port 3000');
});
