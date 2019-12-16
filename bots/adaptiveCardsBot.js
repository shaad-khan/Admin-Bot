// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { ActivityHandler, CardFactory } = require('botbuilder');

// Import AdaptiveCard content.
//const cake = require('../resources/cake.json');
const fs=require('fs');
//const nodemailer = require('nodemailer');
//const choco = require('../resources/choco.json');
//const ImageGalleryCard = require('../resources/ImageGalleryCard.json');
//const LargeWeatherCard = require('../resources/LargeWeatherCard.json');
//const RestaurantCard = require('../resources/RestaurantCard.json');
//const SolitaireCard = require('../resources/SolitaireCard.json');
//const selection = require('../resources/selection.json');
//const sendemail = require('../resources/sendemail.json');
const d_1003 = require('../resources/1003.json');
const d_1005 = require('../resources/1005.json');
//const createcase = require('../resources/createcase.json');
// Create array of AdaptiveCard content, this will be used to send a random card to the user.
/*const CARDS = [
    cake,
    choco,
    selection,
    sendemail,
    createcase,
    d_1003,
    d_1005,
    ImageGalleryCard,
    LargeWeatherCard,
    RestaurantCard,
    SolitaireCard
];*/

const WELCOME_TEXT = 'This bot will try to answer all your system admin related queries...!!';

class AdaptiveCardsBot extends ActivityHandler {
    constructor() {
        super();
        this.onMembersAdded(async (context, next) => {
            const membersAdded = context.activity.membersAdded;
            for (let cnt = 0; cnt < membersAdded.length; cnt++) {
                if (membersAdded[cnt].id !== context.activity.recipient.id) {
               
                   // await context.sendActivity(`Welcome to Continuserve Bot. ${ WELCOME_TEXT }`);
                  // const randomlySelectedCard = CARDS[0];
                   await context.sendActivity({
                       text: '',
                       attachments: [CardFactory.adaptiveCard(require('../resources/1003.json'))]
                   });
               //     await context.sendActivity(`Welcome to Continuserve Bot  ${ membersAdded[cnt].name }. ${ WELCOME_TEXT }`);
                }
            }

            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });

        this.onMessage(async (context, next) => {
           // console.log(JSON.stringify(context._activity));
           fs.writeFileSync("synchronous.txt", JSON.stringify(context._activity));
          var text;
            //const randomlySelectedCard = CARDS[Math.floor((Math.random() * CARDS.length - 1) + 1)];
            if(context._activity.text)
            {
               text=context._activity.text;
              // console.log(JSON.stringify(text));
            /*if(context._activity.text.includes('reset') || context._activity.text.includes('password'))
            {
                const randomlySelectedCard = CARDS[0];
                await context.sendActivity({
                    text: 'Please select the application for which password need to be reset?',
                    attachments: [CardFactory.adaptiveCard(randomlySelectedCard)]
                });
            }
            if(context._activity.text.includes('email'))
            {
                const randomlySelectedCard = CARDS[3];
                console.log(JSON.stringify(randomlySelectedCard));
                await context.sendActivity({
                    text: 'Please enter the email body to be send !!',
                    attachments: [CardFactory.adaptiveCard(randomlySelectedCard)]
                });
            }*/
            if(context._activity.text.includes('hi'))
            {
                //const randomlySelectedCard = CARDS[5];
                //console.log(JSON.stringify(randomlySelectedCard));
                await context.sendActivity({
                    text: '',
                    attachments: [CardFactory.adaptiveCard(require('../resources/1003.json'))]
                });
            }
            
           /* if(context._activity.text.includes('outlook loading'))
            {
                const randomlySelectedCard = CARDS[2];
                randomlySelectedCard.actions[0].data.id="Q-112";
                await context.sendActivity({
                    text: 'Select the option',
                    attachments: [CardFactory.adaptiveCard(randomlySelectedCard)]
                });
            }
            
            /*else
            {
              
                    console.log("here");
                  //  const randomlySelectedCard = CARDS[3];
                    await context.sendActivity({
                        text: 'How i can help you today? I will try to answer all your system admin related queries...!!'
                       // attachments: [CardFactory.adaptiveCard(randomlySelectedCard)]
                    });   
               
    
            }*/
           }
           else if(context._activity.value)
           {
           
       console.log(JSON.stringify(context._activity));

await context.sendActivity({
                    text: '',
                    
                    attachments: [CardFactory.adaptiveCard(require(`../resources/${context._activity.value.value}.json`))]
                });
                if(context._activity.value.value=="1008"||context._activity.value.value=="1009"||context._activity.value.value=="1012")

{
    await context.sendActivity({
    text: '',
    
    attachments: [CardFactory.adaptiveCard(require(`../resources/1001.json`))]
});

}                /* for(var i=0;i<CARDS.length;i++)
            {
                console.log(CARDS[i]+"d_"+context._activity.value);
                if(CARDS[i].includes(context._activity.value))
                {
                    const randomlySelectedCard = CARDS[i];
                
            } 
            }*/
              
        
            
            console.log(JSON.stringify(context));
          if(context._activity.value.email)
          {
             var x=context._activity.value.email;
            nodemailer.createTestAccount((err, account) => {
                // create reusable transporter object using the default SMTP transport
                let transporter = nodemailer.createTransport({
                    host: 'smtp.office365.com',
                    port: 587,
                    secure: false, // true for 465, false for other ports
                    auth: {
                        user: "CS_Connect@continuserve.com", // generated ethereal user
                        pass: "bgbfgqlgtxdknwqn" // generated ethereal password
                    }
                });
        
                // setup email data with unicode symbols
                let mailOptions = {
                    from: 'CS_Connect@continuserve.com', // sender address
                    to: `shadab.k@continuserve.com`, // list of receivers
                    subject: 'Email From CSBOT', // Subject line
                    //text: 'Expense applied by'+ result.EmployeeName, // plain text body
                    html: `<html>
        <head>
            <title>CSConnect :: Multi Factor Authentication</title>
            <style type="text/css">
                table.gridtable {
                    font-family: Franklin Gothic Book,arial,sans-serif;
                    font-size: 13px;
                    color: #000;
                    border-width: 1px;
                    border-color: #337ab7;
                    border-collapse: collapse;
                }
        
                    table.gridtable th {
                        border-width: 1px;
                        padding: 8px;
                        border-style: solid;
                        border-color: #337ab7;
                        background-color: #d9edf7;
                    }
        
                    table.gridtable td {
                        border-width: 1px;
                        padding: 8px;
                        border-style: solid;
                        border-color: #337ab7;
                        background-color: #ffffff;
                    }
        
                table.gridtableWhite {
                    font-family: Franklin Gothic Book,arial,sans-serif;
                    font-size: 13px;
                    color: #333333;
                    border-width: 1px;
                    border-color: #666666;
                    border-collapse: collapse;
                }
        
                    table.gridtableWhite th {
                        border-width: 1px;
                        padding: 8px;
                        border-style: solid;
                        border-color: #666666;
                        background-color: #FFF;
                    }
        
                    table.gridtableWhite td {
                        border-width: 1px;
                        padding: 8px;
                        border-style: solid;
                        border-color: #FFF;
                        background-color: #ffffff;
                    }
        
                .fontStyle {
                    font-family: Franklin Gothic Book,arial,sans-serif;
                }
        
                .text-success {
                    color: #1FD41C;
                    font-weight: bolder;
                }
        
                .main-footer {
                    background: #fff;
                    padding: 15px;
                    color: #444;
                    border-top: 1px solid #d2d6de;
                }
            </style>
        </head>
        <body>
           <h1>${x}
        </body>
        </html>`
                };
        
                // send mail with defined transport object
                transporter.sendMail(mailOptions, (error, info) => {
                    if (error) {
                      
                      
                        return console.log(error);
                    }
                    console.log('Message sent: %s', info.messageId);
                    // Preview only available when sending through an Ethereal account
                    console.log('Preview URL: %s', nodemailer.getTestMessageUrl(info));
        
                 
                    // Message sent: <b658f8ca-6296-ccf4-8306-87d57a0b4321@example.com>
                    // Preview URL: https://ethereal.email/message/WaQKMgKddxQDoou...
                });

          });
          await context.sendActivity({
            text: 'Email was sent successfully'
           // attachments: [CardFactory.adaptiveCard(randomlySelectedCard)]
        });   
        }

          else  if(context._activity.value.CompactSelectVal == 1)
            {
                const randomlySelectedCard = CARDS[2];
                randomlySelectedCard.actions[0].data.id="Q-111";
                await context.sendActivity({
                    text: 'you have selected Office 365',
                    
                    attachments: [CardFactory.adaptiveCard(randomlySelectedCard)]
                }); 
            }
            if(context._activity.value.CompactSelectVal == 2)
            {
                const randomlySelectedCard = CARDS[2];
                randomlySelectedCard.actions[0].data.id="Q-112";
                await context.sendActivity({
                    text: 'you have selected Office',
                    
                    attachments: [CardFactory.adaptiveCard(randomlySelectedCard)]
                }); 
            }
            else if(context._activity.value.CompactSelectVal == 4)
            {
              //  console.log(JSON.stringify(context._activity.value));
              
var data=fs.readFileSync('./resources/createcase.json', 'utf8');
                const randomlySelectedCard = JSON.parse(data);
                console.log(randomlySelectedCard.id);
                await context.sendActivity({
                    text: 'Create a case',
                    attachments: [CardFactory.adaptiveCard(randomlySelectedCard)]
                });   
            }
            else if(context._activity.value.CompactSelectVal == 3 && context._activity.value.id=="Q-111")
            {
                console.log("text"+JSON.stringify(context._activity.value));
                //const randomlySelectedCard = CARDS[3];
                await context.sendActivity({
                    text: 'Follow this link: https://passwordreset.microsoftonline.com/'
                   // attachments: [CardFactory.adaptiveCard(randomlySelectedCard)]
                });   
            }
            else if(context._activity.value.CompactSelectVal == 3 && context._activity.value.id=="Q-112")
            {

                console.log("text"+JSON.stringify(context._activity.value));
                const randomlySelectedCard = CARDS[3];
                delete randomlySelectedCard.body;
                randomlySelectedCard.body=[];
                randomlySelectedCard.body.push({
                    "type": "TextBlock",
                    "text": "Follow this link: https://Office.com/",
                    "weight": "bolder", "separator": true,
                    "isSubtle": false
                });
                randomlySelectedCard.body.push({
                    "type": "TextBlock",
                    "text": "Step 1: Enter the Continuserve email Id and password and logon.",
                    "weight": "bolder",
                    "isSubtle": false, "separator": true,
                    "wrap": true

                });
                randomlySelectedCard.body.push({
                    "type": "TextBlock",
                    "text": "Step 2: Click on your profile->My account->Additional security verification",
                    "weight": "bolder",
                    "isSubtle": false, "separator": true,
                    "wrap": true
                });
                randomlySelectedCard.body.push({
                    "type": "TextBlock",
                    "text": "->Create and manage app passwordsàdelete old password and create new.",
                    "weight": "bolder",
                    "isSubtle": false, "separator": true,
                    "wrap": true
                });
                randomlySelectedCard.body.push({
                    "type": "TextBlock",
                    "text": "Step 3: Copy the new Password some where and save that file.",
                    "weight": "bolder",
                    "isSubtle": false, "separator": true,
                    "wrap": true
                });
                await context.sendActivity({
                    text: 'Do it Yourself',
                    attachments: [CardFactory.adaptiveCard(randomlySelectedCard)]
                });   
            }
            else if(context._activity.value.CompactSelectVal == 3 && context._activity.value.id=="Q-113" )
            {
                console.log("helre"+JSON.stringify(context._activity.value));
                //const randomlySelectedCard = CARDS[3];
                await context.sendActivity({
                    text: 'Follow this link: https://passwordreset.microsoftonline.com/'
                   // attachments: [CardFactory.adaptiveCard(randomlySelectedCard)]
                });   
            }
            
           
            
           }
          
           

            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });
    }
}

module.exports.AdaptiveCardsBot = AdaptiveCardsBot;
