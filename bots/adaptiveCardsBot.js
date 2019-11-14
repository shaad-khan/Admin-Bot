// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { ActivityHandler, CardFactory } = require('botbuilder');

// Import AdaptiveCard content.
const cake = require('../resources/cake.json');

const choco = require('../resources/choco.json');
const ImageGalleryCard = require('../resources/ImageGalleryCard.json');
const LargeWeatherCard = require('../resources/LargeWeatherCard.json');
const RestaurantCard = require('../resources/RestaurantCard.json');
const SolitaireCard = require('../resources/SolitaireCard.json');
const selection = require('../resources/selection.json');
const createcase = require('../resources/createcase.json');
// Create array of AdaptiveCard content, this will be used to send a random card to the user.
const CARDS = [
    cake,
    choco,
    selection,
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
          
            //const randomlySelectedCard = CARDS[Math.floor((Math.random() * CARDS.length - 1) + 1)];
            if(context._activity.text)
            {
               
            if(context._activity.text.includes('reset') || context._activity.text.includes('password'))
            {
                const randomlySelectedCard = CARDS[0];
                await context.sendActivity({
                    text: 'Please select the application for which password need to be reset?',
                    attachments: [CardFactory.adaptiveCard(randomlySelectedCard)]
                });
            }
            
            else
            {
              
                    console.log("here");
                  //  const randomlySelectedCard = CARDS[3];
                    await context.sendActivity({
                        text: 'How i can help you today?'
                       // attachments: [CardFactory.adaptiveCard(randomlySelectedCard)]
                    });   
               
    
            }
           }
           else if(context._activity.value)
           {
            console.log(JSON.stringify(context._activity.value.CompactSelectVal));

            if(context._activity.value.CompactSelectVal == 1)
            {
                const randomlySelectedCard = CARDS[2];
                await context.sendActivity({
                    text: 'you have selected Office 365',
                    attachments: [CardFactory.adaptiveCard(randomlySelectedCard)]
                }); 
            }
            else if(context._activity.value.CompactSelectVal == 4)
            {
                console.log(JSON.stringify(context._activity.value));
                const randomlySelectedCard = CARDS[3];
                await context.sendActivity({
                    text: 'Create a case',
                    attachments: [CardFactory.adaptiveCard(randomlySelectedCard)]
                });   
            }
            else if(context._activity.value.CompactSelectVal == 3)
            {
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
