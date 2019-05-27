// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { ActivityHandler,MessageFactory } = require('botbuilder');
// const choiceProp = 'choiceProp';

class EchoBot extends ActivityHandler {
    /**
     *
     * @param {UserState} User state to persist boolean flag to indicate
     *                    if the bot had already welcomed the user
     */
    constructor(userState,conversationState,dialog,logger) {
        super();
        if (!conversationState) throw new Error('[DialogBot]: Missing parameter. conversationState is required');
        if (!userState) throw new Error('[DialogBot]: Missing parameter. userState is required');
        if (!dialog) throw new Error('[DialogBot]: Missing parameter. dialog is required');
        if (!logger) {
            logger = console;
            logger.log('[DialogBot]: logger not passed in, defaulting to console');
        }
        // this.choiceProp = userState.createProperty(choiceProp);

        this.conversationState = conversationState;
        this.userState = userState;
        this.dialog = dialog;
        this.logger = logger;
        this.dialogState = this.conversationState.createProperty('DialogState');

        // See https://aka.ms/about-bot-activity-message to learn more about the message and other activity types.
        this.onMessage(async (context, next) => {
            this.logger.log('Running dialog with Message Activity.');
    
            await this.dialog.run(context, this.dialogState);
            // if(flag==='Ask' && context.activity.text==='Ask')
            // {
            //     await context.sendActivity('Ask me for sure!');
            // }
            // else if (flag==='Search' && context.activity.text==='Search')
            // {
            //     await context.sendActivity('Search me for sure!');
            // }
            // else
            // {
            //     await context.sendActivity(context.activity.text);
            // }
            // // By calling next() you ensure that the next BotHandler is run.
            // await this.userState.saveChanges(context);
            await next();
        });

        this.onDialog(async (context, next) => {
            // Save any state changes. The load happened during the execution of the Dialog.
            await this.conversationState.saveChanges(context, false);
            await this.userState.saveChanges(context, false);
            await next();
        });

        this.onMembersAdded(async (context, next) => {
            const membersAdded = context.activity.membersAdded;
            console.log(context.activity);
            for (let cnt = 0; cnt < membersAdded.length; ++cnt) {
                console.log(membersAdded[cnt].id);
                console.log(context.activity.recipient.id);
                if (membersAdded[cnt].id !== context.activity.recipient.id) {
                    await this.dialog.run(context, this.dialogState);
                    // await context.sendActivity('Hello and welcome!');
                    // var reply = MessageFactory.suggestedActions(['Ask', 'Search'], 'What do you want to do?');
                    // await context.sendActivity(reply);
                }
            }
            
            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });
    }
}

module.exports.EchoBot = EchoBot;
