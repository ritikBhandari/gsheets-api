const express = require('express');
const app = express();
app.use(express.json());
const fetch = require('node-fetch');
const { stringify } = require('querystring');
const {google} = require('googleapis');

const port = process.env.port || 8000;

const CLIENT_ID = ""
const CLIENT_SECRET = ""
const REDIRECT_URI = `http://localhost:${port}/`


let userCredential;
var bearer = '';

const oauth2Client = new google.auth.OAuth2(
    CLIENT_ID,
    CLIENT_SECRET,
    REDIRECT_URI
)

const scopes = [
    'https://www.googleapis.com/auth/spreadsheets' 
]


const authorizationUrl = oauth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: scopes,
    include_granted_scopes: true
})

app.get('/login', (req, res)=>{
    res.redirect(authorizationUrl);
})

app.get('/', async (req, res)=>{

    try{
        const {tokens} = await oauth2Client.getToken(req.query.code);
        oauth2Client.setCredentials(tokens);
        userCredential = tokens.access_token;
        console.log(userCredential);
        res.json("Copy the access token in your console for authorization purposes!");
    }

    catch(err){
        console.log(err);
    }

})

app.get('/spreadsheet/:id', async (req, res)=>{

   const response = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${req.params.id}?includeGridData=true`, {
    method: 'GET',
    headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${bearer}`
    }
   });

   const data = await response.json();
   

    let userData = {};

    try{
        for (let index=0; index<data.sheets.length; index++){
            userData[data.sheets[index].properties.title] = [];
        
            for(row=0; row<data.sheets[index].data[0].rowData.length; row++){
    
                var obj = {}
    
                let column = 0
    
                if (data.sheets[index].data[0].rowData[row].values==null) {
                    continue;
                } else {
    
                while (column < data.sheets[index].data[0].rowData[row].values.length) {
                if (data.sheets[index].data[0].rowData[row].values[column].formattedValue==null) {
                    data.sheets[index].data[0].rowData[row].values[column].formattedValue=" ";
                }
                obj = {...obj, [column]: data.sheets[index].data[0].rowData[row].values[column].formattedValue};
                column++;
                }
    
                userData[data.sheets[index].properties.title].push(obj);
    
            }}}
            
            res.json({"success" : true, data: userData});
    }

    catch(err){
        if(err.message=="Cannot read properties of undefined (reading 'length')")
        res.json("Data isn't available for this sheet. Either refresh your access token or try another sheet");
    }

})

app.post('/spreadsheet/update', async (req, res)=>{

        const response = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${req.body.id}/values:batchUpdateByDataFilter`, {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${bearer}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({  
                "data": [
                    {
                      "dataFilter": {
                        "gridRange": {
                          "endColumnIndex": req.body.column_number,
                          "endRowIndex": req.body.row_number,
                          "sheetId": req.body.sheetId,
                          "startColumnIndex": req.body.column_number-1,
                          "startRowIndex": req.body.row_number-1
                        }
                      },
                      "values": [
                        [
                          req.body.values
                        ]
                      ],
                      "majorDimension": "ROWS"
                    }
                  ],
                  "valueInputOption": "RAW"
        })
        })
    
        if(response.ok){
            res.json({"Success": true})
        }
        else{
            res.json({"Success": false})
        }
    })


app.listen(port, ()=>{
    console.log(`Server listening at ${port}`);
})
