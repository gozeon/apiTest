const express = require('express');
const morgan = require('morgan')
const bodyParser = require('body-parser')
const app = express();

app.use(morgan('dev'))
app.use(bodyParser.json())

app.use('*', function (req, res) {
    console.log(req.body)
    res.json({
        "status": 0,
        "msg": "ok",
        error: '',
        "data": [],
    });
})



app.listen(8080)
