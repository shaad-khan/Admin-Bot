// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { ActivityHandler, CardFactory } = require('botbuilder');

// Import AdaptiveCard content.
const cake = require('../resources/cake.json');
const fs=require('fs');
const nodemailer = require('nodemailer');
const choco = require('../resources/choco.json');
const ImageGalleryCard = require('../resources/ImageGalleryCard.json');
const LargeWeatherCard = require('../resources/LargeWeatherCard.json');
const RestaurantCard = require('../resources/RestaurantCard.json');
const SolitaireCard = require('../resources/SolitaireCard.json');
const selection = require('../resources/selection.json');
const sendemail = require('../resources/sendemail.json');
const createcase = require('../resources/createcase.json');
// Create array of AdaptiveCard content, this will be used to send a random card to the user.
const CARDS = [
    cake,
    choco,
    selection,
    sendemail,
    createcase,
    ImageGalleryCard,
    LargeWeatherCard,
    RestaurantCard,
    SolitaireCard
];

const WELCOME_TEXT = 'This bot will try to answer all your system admin related queries...!!';

class AdaptiveCardsBot extends ActivityHandler {
    constructor() {
        super();
        this.onMembersAdded(async (context, next) => {
            const membersAdded = context.activity.membersAdded;
            for (let cnt = 0; cnt < membersAdded.length; cnt++) {
                if (membersAdded[cnt].id !== context.activity.recipient.id) {
               
                    await context.sendActivity(`Welcome to Continuserve Bot. ${ WELCOME_TEXT }`);
               
               //     await context.sendActivity(`Welcome to Continuserve Bot  ${ membersAdded[cnt].name }. ${ WELCOME_TEXT }`);
                }
            }

            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });

        this.onMessage(async (context, next) => {
            //console.log(context);
          var text;
            //const randomlySelectedCard = CARDS[Math.floor((Math.random() * CARDS.length - 1) + 1)];
            if(context._activity.text)
            {
               text=context._activity.text;
               console.log(JSON.stringify(text));
            if(context._activity.text.includes('reset') || context._activity.text.includes('password'))
            {
                const randomlySelectedCard = CARDS[0];
                await context.sendActivity({
                    text: 'Please select the application for which password need to be reset?',
                    attachments: [CardFactory.adaptiveCard(randomlySelectedCard)]
                });
            }
            if(context._activity.text.includes('Send Email'))
            {
                const randomlySelectedCard = CARDS[3];
                await context.sendActivity({
                    text: 'Please select the application for which password need to be reset?',
                    attachments: [CardFactory.adaptiveCard(randomlySelectedCard)]
                });
            }
            if(context._activity.text.includes('outlook loading'))
            {
                const randomlySelectedCard = CARDS[2];
                randomlySelectedCard.actions[0].data.id="Q-112";
                await context.sendActivity({
                    text: 'Select the option',
                    attachments: [CardFactory.adaptiveCard(randomlySelectedCard)]
                });
            }
            
            else
            {
              
                    console.log("here");
                  //  const randomlySelectedCard = CARDS[3];
                    await context.sendActivity({
                        text: 'How i can help you today? I will try to answer all your system admin related queries...!!'
                       // attachments: [CardFactory.adaptiveCard(randomlySelectedCard)]
                    });   
               
    
            }
           }
           else if(context._activity.value)
           {
            console.log(JSON.stringify(context));

            if(context._activity.value.CompactSelectVal == 1)
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
