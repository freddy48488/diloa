const express = require('express');
const app = express();
const hostname = '127.0.0.1'
const port = 8080
var bodyParser = require('body-parser');
var urlencodedParser = bodyParser.urlencoded({ extended: false })
const MongoClient = require('mongodb').MongoClient
const dbname = 'SIM_CARD'
const connectstr = 'mongodb://jacob129:w920913s@cluster0-shard-00-00-4jze1.mongodb.net:27017,cluster0-shard-00-01-4jze1.mongodb.net:27017,cluster0-shard-00-02-4jze1.mongodb.net:27017/test?ssl=true&replicaSet=Cluster0-shard-0&authSource=admin&retryWrites=true&w=majority'
const mongoose = require('mongoose')
        Schema = mongoose.Schema,
        objectId = Schema.objectId,
        mongoose.Promise = global.Promise
var xl = require('excel4node');
var wb = new xl.Workbook();
var ws = wb.addWorksheet('Sheet 1');

//style create start
var fonts_style = wb.createStyle({
    font: {
      color: '#000000',
      size: 12,
    },
    alignment: {
        horizontal: 'center',
        vertical: 'center'
    },
});

var border_style = wb.createStyle({
    border:{
        left:{
            style: 'medium',
            color: 'black'
        },
        right:{
            style: 'medium',
            color: "black"
        },
        top:{
            style: 'medium',
            color: 'black'
        },
        bottom:{
          style: 'medium',
          color: 'black'
        },
        outline: false,
    },
})
//style create end

//codes
app.use(bodyParser.urlencoded({ extended: true }))
app.use(bodyParser.json())

app.get('/', function(req, res){
    res.sendFile(__dirname + '/index.html')
})

app.post('/', function(req, res){
    res.sendFile(__dirname + '/index.html')
})

// app.get('/secret', function(req, res){
//     res.sendFile(__dirname + '/secret.html')
// })

// app.post('/secret', function(req, res){
//     res.sendFile(__dirname + '/secret.html')
//})

app.get('/accounting', function(req, res){
    res.sendFile(__dirname + '/accounting.html')
})

app.post('/accounting', function(req, res){
    res.setHeader('Content-Type', 'text/html')
    res.sendFile(__dirname + '/accounting.html')
    const ary = []
    const ary1 = []
    let num1 = parseFloat(req.body.plan1)
    if(isNaN(num1)) num1 = 0
    let num2 = parseFloat(req.body.plan2)
    if(isNaN(num2)) num2 = 0
    let num3 = parseFloat(req.body.plan3)
    if(isNaN(num3)) num3 = 0
    let num4 = parseFloat(req.body.plan4)
    if(isNaN(num4)) num4 = 0
    const totalnum = parseInt(num1) + parseInt(num2) + parseInt(num3) + parseInt(num4)
    const insert = {
        開新卡三個月數量: num1,
        新卡續費三個月數量: num2,
        開卡一個月數量: num3,
        續費一個月數量: num4,
        總數量: totalnum
    }
    const getvalue1 = req.body.plan1 *= 6000
    const getvalue2 = req.body.plan2 *= 5000 
    const getvalue3 = req.body.plan3 *= 3000
    const getvalue4 = req.body.plan4 *= 2000
    const totalvalue = getvalue1 + getvalue2 + getvalue3 + getvalue4
    const insert1 = {
        開新卡三個月總額: getvalue1,
        新卡續費三個月總額: getvalue2,
        開卡一個月總額: getvalue3,
        續費一個月總額: getvalue4,
        總金額: totalvalue
    }
    const getrmb = parseFloat(req.body.rmb)
    MongoClient.connect(connectstr, function(err, client){
        var dbcol = client.db(dbname)
        ary.push(insert)
        ary1.push(insert1)
        console.log(ary)
        console.log(ary1)
        quotesCollection = dbcol.collection('accounting')
        quotesCollection.insertMany(ary)
        quotesCollection.insertMany(ary1)
    })
    let count1 = parseFloat(totalnum) * 882
    let count = count1 * parseFloat(getrmb).toFixed(2)
    ws.cell(1, 1)
    .string('方案')
    .style({font:{size: 12}});
    //space
    ws.cell(2, 1)
    .string('時間(拿卡/開卡)')
    .style({font:{size: 12}});
    //space
    ws.cell(3, 1, 8, 1, true)
    .string('收入')
    .style(fonts_style)
    .style(border_style)
    //space
    ws.cell(3, 2, 8, 6)
    .style(border_style)
    //space
    ws.cell(4, 2)
    .string('新卡三個月')
    .style(fonts_style)
    //space
    ws.cell(5, 2)
    .string('續卡三個月')
    .style(fonts_style)
    //space
    ws.cell(6, 2)
    .string('新卡一個月')
    .style(fonts_style)
    //space
    ws.cell(7, 2)
    .string('續卡一個月')
    .style(fonts_style)
    //space
    ws.cell(8, 2)
    .string('總收入')
    .style(fonts_style)
    //space
    ws.cell(3, 3)
    .string('金額')
    .style(fonts_style)
    //space
    ws.cell(3, 5)
    .string('數量')
    .style(fonts_style)
    //space
    ws.cell(3, 6)
    .string('總額')
    .style(fonts_style)
    //space
    ws.cell(4, 3)
    .string('6000')
    .style(fonts_style)
    //space
    ws.cell(5, 3)
    .string('5000')
    .style(fonts_style)
    //space
    ws.cell(6, 3)
    .string('3000')
    .style(fonts_style)
    //space
    ws.cell(7, 3)
    .string('2000')
    .style(fonts_style)
    //space
    ws.cell(4, 4, 7, 4, true)
    .string('X')
    .style(fonts_style)
    //space
    ws.cell(10, 1, 11, 1, true)
    .string('成本')
    .style(fonts_style)
    .style(border_style)
    //space
    ws.cell(10, 2, 11, 6)
    .style(border_style)
    //space
    ws.cell(10, 2)
    .string('數量 X')
    .style(fonts_style)
    //space
    ws.cell(10, 3)
    .string('人民幣')
    .style(fonts_style)
    //space
    ws.cell(10, 4, 11, 4, true)
    .string('X')
    .style(fonts_style)
    //space
    ws.cell(10, 5)
    .string('匯率')
    .style(fonts_style)
    //space
    ws.cell(10, 6)
    .string('總額')
    .style(fonts_style)
    //space
    ws.cell(13, 1, 14, 6)
    .style(border_style)
    //space
    ws.cell(13, 1)
    .string('毛利')
    .style(fonts_style)
    //space
    ws.cell(14, 1)
    .string('毛利率')
    .style(fonts_style)
    //space
    ws.cell(13, 2)
    .string('收入總額-成本總額')
    .style(fonts_style)
    //space
    ws.cell(14, 2)
    .string('(收入總額-成本總額)/收入總額')
    .style(fonts_style)
    //space
    ws.column(1).setWidth(20)
    ws.column(2).setWidth(30)
    ws.column(4).setWidth(5)
    //space
    let getplan = req.body.plan
    ws.cell(1, 2)
    .string(JSON.stringify(getplan))
    .style(fonts_style)
    //space
    let gettake = req.body.takedate
    ws.cell(2, 2)
    .string(JSON.stringify(gettake))
    .style(fonts_style)
    //space
    let getact = req.body.dateactive
    ws.cell(2, 3)
    .string(JSON.stringify(getact))
    .style(fonts_style)
    //space
    ws.cell(4, 6)
    .string(JSON.stringify(getvalue1))
    .style(fonts_style)
    //space
    ws.cell(5, 6)
    .string(JSON.stringify(getvalue2))
    .style(fonts_style)
    //space
    ws.cell(6, 6)
    .string(JSON.stringify(getvalue3))
    .style(fonts_style)
    //space
    ws.cell(7, 6)
    .string(JSON.stringify(getvalue4))
    .style(fonts_style)
    //space
    ws.cell(8, 6)
    .string(JSON.stringify(totalvalue))
    .style(fonts_style)
    //space
    ws.cell(11, 2)
    .string(JSON.stringify(totalnum))
    .style(fonts_style)
    //space
    ws.cell(4, 5)
    .string(JSON.stringify(num1))
    .style(fonts_style)
    //space
    ws.cell(5, 5)
    .string(JSON.stringify(num2))
    .style(fonts_style)
    //space
    ws.cell(6, 5)
    .string(JSON.stringify(num3))
    .style(fonts_style)
    //space
    ws.cell(7, 5)
    .string(JSON.stringify(num4))
    .style(fonts_style)
    //space
    ws.cell(11, 3)
    .string('882')
    .style(fonts_style)
    //space
    ws.cell(11, 5)
    .string(JSON.stringify(getrmb))
    .style(fonts_style)
    //space
    ws.cell(11, 6)
    .string(JSON.stringify(count))
    .style(fonts_style)
    //space
    ws.cell(13, 4)
    .string('-')
    .style(fonts_style)
    //space
    ws.cell(14, 4)
    .string('/')
    .style(fonts_style)
    //space
    ws.cell(13, 3)
    .string(JSON.stringify(totalvalue))
    .style(fonts_style)
    //space
    ws.cell(13, 5)
    .string(JSON.stringify(count))
    .style(fonts_style)
    //space
    let result1 = parseFloat(totalvalue).toFixed(2) - parseFloat(count).toFixed(2)
    ws.cell(13, 6)
    .string(JSON.stringify(result1))
    .style(fonts_style)
    //space
    ws.cell(14, 3)
    .string(JSON.stringify(result1))
    .style(fonts_style)
    //space
    ws.cell(14, 5)
    .string(JSON.stringify(totalvalue))
    .style(fonts_style)
    //space
    let result2 = parseFloat(result1).toFixed(2) / parseFloat(totalvalue).toFixed(2)
    ws.cell(14, 6)
    .string(JSON.stringify(result2))
    .style(fonts_style)
    //space
    wb.write('Excel.xlsx');
})

app.get('/search', function(req, res){
    res.sendFile(__dirname + '/search.html')
})

app.post('/search', function(req, res){
    res.setHeader('Content-Type', 'text/html');
    res.sendFile(__dirname + '/search.html')
    MongoClient.connect(connectstr, function(err, client){
        var dbcol = client.db(dbname).collection('cards')
        search = {
            cardsnumber: req.body.cardsnum
        }
        JSON.stringify(search).toUpperCase()
        dbcol.findOne(search, function(err, item){
            if(err){
                throw err
            }
            else{
                console.log(item)
            }
        })
    })
})

app.get('/add', function(req, res){
    res.sendFile(__dirname + '/add.html')
})

app.post('/add', urlencodedParser, function(req, res, next){
    res.sendFile(__dirname + '/add.html')
    userdata = {
        名稱: req.body.username,
        卡號: req.body.cardsnum,
        拿卡日期: req.body.takedate,
        開卡日期: req.body.dateactive,
        方案: req.body.plan,
        是否繳錢: req.body.money,
        賣價: req.body.price,
        出入價: req.body.price1,
        卡片成本: req.body.cost,
        付款日期: req.body.paydate,
        付款金額: req.body.pay
    }
    var asname = req.body.username
    MongoClient.connect(connectstr, function(err, client){
        const dbo = client.db(dbname)
        const str = JSON.stringify(asname).toUpperCase().replace("\"", "")
        const str1 = str.replace("\"", "")
        quotesCollection = dbo.collection(str1)
        quotesCollection.insertOne(userdata)
        quotesCollection1 = dbo.collection('cards')
        quotesCollection1.insertOne(userdata)
    })
})

app.listen(port, hostname, () => {
    console.log(`Server running at http://${hostname}:${port}`)
})