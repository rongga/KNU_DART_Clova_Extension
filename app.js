const express = require('express');
const app = express();
const router = require('./routes/router')(app);

//app.set('host', process.env.HOST || '0.0.0.0');
app.use('/',router);

app.listen(3000,function(){
  console.log('3000');
});
