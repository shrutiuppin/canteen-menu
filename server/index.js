var express = require('express')
var app = express()
var nodemailer = require('nodemailer');
var xls = require('excel');
var https = require('https');
var fs = require('fs');

//var file = fs.createWriteStream("file.xlsx");
// var request = https.get(" ", function(response) {
//   response.pipe(file);
// });

app.get('/', function (req, res) {
    console.log("hello");
    getmenu(new Date());
    res.send('Hello World')
})

app.listen(3000)

function getmenu(today) {
   console.log("worked");
  //  var today = new Date(second);
    var startDay = new Date('1900-01-01');

    var day = Math.round((today - startDay) / (1000 * 60 * 60 * 24)) + 2;
    
    var flag = 0,
        iter = 0;
   
    xls('file.xlsx', function (err, data) {
        if (err) throw err;
        // data is an array of arrays
        console.log(data);
        for (var i = 3; i < data.length - 5; i += 2) {
            if (data[i][0] == day) {
                flag = 1;
                iter = i;
            }
        }

        var transporter = nodemailer.createTransport({
            service: 'gmail',
            auth: {
                user: 'abcdummy54@gmail.com',
                pass: 'abc@123!'
            }
        });

        if (flag == 1) {
            var mailOptions = {
                from: 'abcdummy54@gmail.com',
                to: 'shrutiuppin149@gmail.com',
                subject: 'Menu for ' + data[iter][1],
                html: '<br>' + '<table border="1">' +
                    '<tr>' +
                    '<th>' + "Breakfast" + '</th>' +
                    '<th>' + "Lunch" + '</th>' +
                    '<th>' + "Snacks" + '</th>' +
                    '</tr>' +
                    '<tr>' +
                    '<td align="center">' + data[iter][2] + '</td>' +
                    '<td align="center">' + data[iter][3] + '<br>' +
                    data[iter][4] + '<br>' +
                    data[iter][5] + '<br>' +
                    data[iter][6] + '<br>' +
                    data[iter][7] + '<br>' +
                    data[iter][8] + '<br>' +
                    data[iter][9] + '<br>' +
                    data[iter][10] + '<br>' +
                    data[iter][11] +
                    '<td align="center">' + data[iter][12] + '</td>' +
                    '</tr>' +
                    '</table>'
            };
        } else {
            console.log("not ");
            var mailOptions = {
                from: 'abcdummy54@gmail.com',
                to: 'shrutiuppin149@gmail.com',
                subject: 'Menu for day does not exist',
                html: '<h2>Try ' + data[1][0] + ' and except Sundays</h2>'
            };
        }

        transporter.sendMail(mailOptions, function (error, info) {
            if (error) {
                console.log(error);
            } else {
                console.log('Email sent: ' + info.response);
            }
        });
    });
}

//getmenu('2015-8-6');