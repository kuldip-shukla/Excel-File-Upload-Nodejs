const express = require('express');
const app = express();
const multer = require('multer');
const path = require('path');
const xlsx = require('xlsx');
const mongoose = require("mongoose");
const config = require("./config");
const bodyParser = require("body-parser");
const schema = mongoose.Schema;

app.use(bodyParser.urlencoded({extended:false}));
app.use(bodyParser.json());
mongoose.connect(config.path);

const dataSchema = new schema({
    Name:{type:String},
    Question_No:{type:Number},
    Option_1:{type:String},
    Option_2:{type:String},
    Option_3:{type:String},
    Option_4:{type:String},
    Priority:{type:String},
    Answer:{type:String}
})

var upload = multer({dest: __dirname+'/upload/'});
var storage = multer.diskStorage({ 
    destination: function (req, file, cb) { 
        cb(null, __dirname+'/upload') 
    }, 
    filename: function (req, file, cb) { 
        console.log(file.originalname)
        cb(null,file.originalname) 
    } 
})
var upload = multer({ storage: storage }).single('file');

app.get('/',(req,res)=>{
    res.sendFile(__dirname+'/file.html');
});

app.post('/upload',(req,res)=>{
    upload(req,res,function(err) 
    {
        if(err) 
        {
            console.log(err)
            return res.end("Error in File Uploading and Saving");
        }
        var workbook = xlsx.readFile(__dirname+'/upload/practical.xlsx');
        var sheet_name_list = workbook.SheetNames; 
        sheet_name_list.forEach(function(y) 
        {
            var worksheet = workbook.Sheets[y];
            var headers = {};
            var data = [];
            var myData = data;
            for(z in worksheet) 
            {
                if(z[0] === '!') continue;
                var tt = 0;
                for (var i = 0; i < z.length; i++) 
                {
                    if (!isNaN(z[i])) 
                    {
                        tt = i;
                        break;
                    }
                }
                var col = z.substring(0,tt);
                var row = parseInt(z.substring(tt));
                var value = worksheet[z].v;
                if(row == 1 && value) 
                {
                    headers[col] = value;
                    continue;
                }

                if(!data[row]) data[row]={};
                data[row][headers[col]] = value;
            }
            data.shift();
            data.shift();
            var arr1 = [];
            var arr2 = [];
            var all = [];
            for (var i = 0; i < data.length; i++) 
            {
                if(!data[i].Priority){
                    data[i].Priority = "null"
                    arr1.push(data[i])
                }
                else{
                    arr2.push(data[i])
                }
            }
            all = arr1
            var min = Number.MIN_SAFE_INTEGER
            var sorted = arr2.sort(function (a,b) {
                return (a.Priority || min) - (b.Priority || min)
            })
            
            for (var i = 0; i < arr2.length; i++) 
            {
                all.splice(arr2[i].Priority-1 ,0,arr2[i])
            }
            console.log(all)
            var temp = mongoose.model('filedata', dataSchema)
            temp.insertMany(all)
            module.exports = temp
        });
    })  
    res.end("File is Uploading and Saving Successfully"); 
});
app.listen(8888);