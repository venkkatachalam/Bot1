// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.
const { ActivityTypes } = require('botbuilder');
const {
    ChoiceFactory,
    ChoicePrompt,
    ComponentDialog,
    DialogSet,
    DialogTurnStatus,
    TextPrompt,
    WaterfallDialog,
    ConfirmPrompt
} = require('botbuilder-dialogs');
const { UserProfile } = require('./userProfile');

const CONFIRM_PROMPT = 'CONFIRM_PROMPT';
const CHOICE_PROMPT = 'CHOICE_PROMPT';
const NAME_PROMPT = 'NAME_PROMPT';
const USER_PROFILE = 'USER_PROFILE';
const WATERFALL_DIALOG = 'WATERFALL_DIALOG';

class UserProfileDialog extends ComponentDialog {
    constructor(userState, logger) {
        super('userProfileDialog');

        this.userProfile = userState.createProperty(USER_PROFILE);

        this.logger = logger;

        this.addDialog(new TextPrompt(NAME_PROMPT));
        this.addDialog(new ChoicePrompt(CHOICE_PROMPT));
        this.addDialog(new ConfirmPrompt(CONFIRM_PROMPT));


        this.addDialog(new WaterfallDialog(WATERFALL_DIALOG, [
            this.AskorSearchStep.bind(this),
            this.textStep.bind(this),
            // this.textConfirmStep.bind(this),
            this.summaryStep.bind(this)
        ]));

        this.initialDialogId = WATERFALL_DIALOG;
    }

    /**
     * The run method handles the incoming activity (in the form of a TurnContext) and passes it through the dialog system.
     * If no dialog is active, it will start the default dialog.
     * @param {*} turnContext
     * @param {*} accessor
     */
    async run(turnContext, accessor) {
        const dialogSet = new DialogSet(accessor);
        dialogSet.add(this);

        const dialogContext = await dialogSet.createContext(turnContext);
        const results = await dialogContext.continueDialog();
        if (results.status === DialogTurnStatus.empty) {
            await dialogContext.beginDialog(this.id);
        }
    }

    async AskorSearchStep(step) {
        // WaterfallStep always finishes with the end of the Waterfall or with another dialog; here it is a Prompt Dialog.
        // Running a prompt here means the next WaterfallStep will be run when the users response is received.
        return await step.prompt(CHOICE_PROMPT, {
            prompt: 'Choose your option wisely!',
            choices: ChoiceFactory.toChoices(['Ask', 'Search'])
        });
    }

    async textStep(step) {
        step.values.Qtype = step.result.value;
        return await step.prompt(NAME_PROMPT, `What do you want to ${step.result.value} for?`);
    }
    async textConfirmStep(step) {
        step.values.Ques = step.result;

        // We can send messages to the user at any point in the WaterfallStep.
        await step.context.sendActivity(`Please wait while we search for ${ step.result }.`);

        // WaterfallStep always finishes with the end of the Waterfall or with another dialog; here it is a Prompt Dialog.
        return await step.prompt(CONFIRM_PROMPT, 'Confirm?', ['yes', 'no']);
    }

    async summaryStep(step) {
        if (step.result) {
            // console.log(step.result.value);
            // console.log(step.values.Qtype);

            step.values.Ques = step.result;
            // Get the current profile object from user state.
            const userProfile = await this.userProfile.get(step.context, new UserProfile());

            userProfile.Qtype = step.values.Qtype;
            userProfile.Ques = step.values.Ques;
            const reply = { type: ActivityTypes.Message };
            reply.attachments = [{
                name: 'pdf.pdf',
                contentType: 'doc/pdf',
                contentUrl: 'http://www.pdf995.com/samples/pdf.pdf'
            }];

            let msg = `You ${ userProfile.Qtype }ed for ${ userProfile.Ques }.`;

            await step.context.sendActivity(reply);
        } else {
            await step.context.sendActivity('Thanks. Your profile will not be kept.');
        }

        // WaterfallStep always finishes with the end of the Waterfall or with another dialog, here it is the end.
        return await step.endDialog();
    }
}

module.exports.UserProfileDialog = UserProfileDialog;
