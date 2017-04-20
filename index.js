/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/*jslint es6 this:true */
"use strict";

// Amazon Alexa SDK
// https://github.com/alexa/alexa-skills-kit-sdk-for-nodejs
var Alexa = require("alexa-sdk");

// Microsoft Graph JavaScript SDK
// npm install msgraph-sdk-javascript
var MicrosoftGraph = require("msgraph-sdk-javascript");

// Email helper
var emailHelper = require("./emailHelper.js");

// Graph client
var client = {};

// Message when the skill is first called
var WelcomeMessage = "Welcome to gina office. . Here are some things you can say. . . Send Mail . . Send a test message . . or Send me mail";
var WelcomeMessageCard = 'Here are some things you can say:\n\n "Send Mail" \n "Send a test message" \n "Send me mail"';
var WelcomeMessageCardTitle = "Welcome to Graph Bot"

// Message for help intent
var HelpMessage = "Here are some things you can say. . . Send Mail . . Send a test message . . or Send me mail";
var HelpMessageCard = 'Here are some things you can say:\n\n "Send Mail" \n "Send a test message" \n "Send me mail"';
var HelpMessageCardTitle = "Graph Bot Help"

// Used to tell user skill is closing
var shutdownMessage = "Ok see you again soon.";

// Used to tell user to link their account to the skill
var linkAccountMessage = "Please link your Microsoft account to use this skill.";

// Response Card images
// An image cannot be larger than 2 MB
var responseCardImages = {
    // 720w x 480h
    smallImageUrl: "https://raw.githubusercontent.com/PaulStubbs/nodejs-alexa-connect-sample/master/skill/AlexaGraphBotCard720x480.png",
    // 1200w x 800h
    largeImageUrl: "https://raw.githubusercontent.com/PaulStubbs/nodejs-alexa-connect-sample/master/skill/AlexaGraphBotCard1200x800.png"
};

// Adding Telemetry
// http://VoiceLabs.co
var VoiceInsights = require('voice-insights-sdk');
// SETUP NOTE: Add Lambda Environment Variable called VI_APP_TOKEN
var VI_APP_TOKEN = process.env.VI_APP_TOKEN;

// Adding logging levels
// error -  Other runtime errors or unexpected conditions. 
// warn -   'Almost' errors, other runtime situations that are undesirable 
//          or unexpected, but not necessarily "wrong".
// info -   Interesting runtime events (startup/shutdown/Intents). 
// debug -  Detailed information on the flow through the system. 
var logLevels = {error: 3, warn: 2, info: 1, debug: 0};

// get the current log level from the current environment if set, else set to info
// SETUP NOTE: Add Lambda Environment Variable called LOG_LEVEL
var currLogLevel = process.env.LOG_LEVEL != null ? process.env.LOG_LEVEL : 'debug';

var replyMessage = '';
// print the log statement, only if the requested log level is greater than the current log level
function log(statement, logLevel) {

    // no loglevel, set to debug
    if(!logLevel){
        logLevel = logLevels.debug;
    }
    // output log if greater then log level
    if(logLevel >= logLevels[currLogLevel] ) {
        console.log(statement);
    }
}

// Entry point for Alexa
exports.handler = function(event, context, callback) {

    // Initialize Telemetry
    VoiceInsights.initialize(event.session, VI_APP_TOKEN);

    // DEBUG: return all the environment varibles
    log(process.env, logLevels.debug);
    // DEBUG: return the Node version info
    log(process.versions, logLevels.debug);

    // Verify that the Request is Intended for Your Service
    // This value is set on the server or in the .env file
    // SETUP NOTE: Add Lambda Environment Variable called APPLICATION_ID
    var appId = process.env.APPLICATION_ID;

    var alexa = Alexa.handler(event, context);
    alexa.appId = appId;
    alexa.registerHandlers(handlers);

    // Get the OAuth 2.0 Bearer Token from the linked account
    var token = event.session.user.accessToken;

    // validate the Auth Token
    if (token) {
        log("Auth Token: " + token, logLevels.debug);
        // TODO: validate the token

        // Initialize the Microsoft Graph client
        client = MicrosoftGraph.Client.init({
            authProvider: (done) => {
                done(null, token);
            }
        });

        // Handle the intent
        alexa.execute();

    } else {
        // no token! display card and let user know they need to sign in
        log("No Auth Token", logLevels.warn);
        var speechOutput = linkAccountMessage;

        VoiceInsights.track("NoAuthToken", null, null, (error, response) => {
            alexa.emit(":tellWithLinkAccountCard", speechOutput);
        });
    }
};


var handlers = {
    "LaunchRequest": function () {
        log("LaunchRequest", logLevels.info);
        var speechOutput = WelcomeMessage;
        var repromptSpeech = WelcomeMessage;
        var cardTitle = WelcomeMessageCardTitle;
        var cardContent = WelcomeMessageCard;
        VoiceInsights.track("LaunchRequest", null, null, (error, response) => {
            this.emit(":askWithCard", speechOutput, repromptSpeech, cardTitle, cardContent, responseCardImages);
        });
    },
    "SessionEndedRequest": function () {
        log("SessionEndedRequest", logLevels.info);
        var speechOutput = shutdownMessage;
        VoiceInsights.track("SessionEndedRequest", null, null, (error, response) => {
            this.emit(":tell", speechOutput);
        });
    },
    "SendMailIntent": function () {
        log("SendMailIntent", logLevels.info);

        SendMailIntent(this);

    },
    "WhoAmIIntent": function () {
        log("WhoAmIIntent", logLevels.info);

        WhoAmIIntent(this);

    },
    "WhatsNextIntent": function () {
        log("WhatsNextIntent", logLevels.info);

        WhatsNextIntent(this);

    },
    "AMAZON.StopIntent": function() {
        log("StopIntent", logLevels.info);
        var speechOutput = shutdownMessage;
        VoiceInsights.track("StopIntent", null, null, (error, response) => {
            this.emit(":tell", speechOutput);
        });
    },
    // Let the user completely exit the skill
    "AMAZON.CancelIntent": function() {
        log("CancelIntent", logLevels.info);
        VoiceInsights.track("CancelIntent", null, null, (error, response) => {
            this.emit(":tell", shutdownMessage);
        });
    },
    // Provide help about how to use the skill
    "AMAZON.HelpIntent": function () {
        log("HelpIntent", logLevels.info);
        var speechOutput = HelpMessage;
        var repromptSpeech = HelpMessage;
        var cardTitle = HelpMessageCardTitle;
        var cardContent = HelpMessageCard;
        VoiceInsights.track("HelpIntent", null, null, (error, response) => {
           this.emit(":askWithCard", speechOutput, repromptSpeech, cardTitle, cardContent, responseCardImages);
        });
    },
    // Catch everything else
    "Unhandled": function () {
        log("UnhandledIntent", logLevels.info);
        var speechOutput = HelpMessage;
        var repromptSpeech = HelpMessage;
        VoiceInsights.track("UnhandledIntent", null, null, (error, response) => {
            this.emit(":ask", speechOutput, repromptSpeech);
        });
    }
};

function WhoAmIIntent(alexaResponse){

        //return the results to Alexa
        VoiceInsights.track("WhatsNextIntent", null, null, (error, response) => {
            return alexaResponse.emit(":tell", "What Next is not yet implemented");
        });
}

function WhoAmIIntent(alexaResponse){
        // get the authenticated user info
        getUser(alexaResponse)
        // **
        // handle the getUser results
        .then(function(user){
            //check if the user is valid
            if(!user) throw "There is no user returned ";

            var displayName = user.displayName;

            //return the results to Alexa
            VoiceInsights.track("WhoAmIIntent", null, null, (error, response) => {
                return alexaResponse.emit(":tell", "The linked account belongs to " + displayName);
            });
        })
        .catch(function(err){
            log("WhoAmIIntent getUser Error: " + JSON.stringify(err), logLevels.error);
            alexaResponse.emit(":tell", "There was an error. " + err.message)
            // re-throw the error so the chain of promises don't continue
            throw "There was a getuser catch error: " + JSON.stringify(err);
        })
}

function SendMailIntent(alexaResponse){
        // get the authenticated user info
        getUser(alexaResponse)
        // **
        // handle the getUser results
        .then(function(user){
            //check if the user is valid
            if(!user) throw "There is no user returned ";

            // then send a mail to the current user           
            return sendMail(user);
        })
        .catch(function(err){
            log("SendMailIntent getUser Error: " + JSON.stringify(err), logLevels.error);
            alexaResponse.emit(":tell", "There was an error. " + err.message)
            // re-throw the error so the chain of promises don't continue
            throw "1111There was a getuser catch error: " + JSON.stringify(err);
        })

        // handle the sendMail results
        .then(function(mail){
            // check if the sendmail succeded
            if(!mail) throw "2222There was an error sending mail";

            // then send confirmation back to alexa
            // var mailSubject = mail.Message.Subject;
            log("Mail Sent: " + JSON.stringify(mail), logLevels.debug);
            //return the results to Alexa
            VoiceInsights.track("sendMailIntent", null, null, (error, response) => {
                return alexaResponse.emit(":tell", mail);
            });
        })
        .catch(function(err){
            log("sendMail Error: " + JSON.stringify(err), logLevels.error);
            alexaResponse.emit(":tell", "3333There was an error sending the mail");
            // re-throw the error so the chain of promises don't continue
            throw "There was an sendmail catch error: " + JSON.stringify(err);
        })

}

function getUser(){
    log("getUser", logLevels.debug)
    //Make a call to the Graph API, this returns a Promise
    return client
            .api("/me")
            .get();
}

function sendMail(user){
    log("sendMail: " + JSON.stringify(user), logLevels.debug)

    return new Promise(function(resolve, reject){

    var destinationEmailAddress = user.userPrincipalName;
    var displayName = user.displayName;

    log("sendMail: email: " + destinationEmailAddress + " name: " + displayName, logLevels.info);
    var mail = emailHelper.generateMailBody(
        displayName,
        destinationEmailAddress
    );

    //DEBUG: log the user
    log("displayName: " + displayName + 
                " destinationEmailAddress: " + destinationEmailAddress); 

    //Make a call to the Graph API
    // client
    //     .api("/me/sendMail")
    //     .post({message: mail.Message}, (err, res) => {
    //         if(err){
    //             log("sendMail Error: " + JSON.stringify(err));
    //             reject(err);
    //         }else{
    //             // log the sendMail results
    //             log("sendMail successful: ");
    //             // return the mail that was sent
    //             VoiceInsights.track("getUser", null, JSON.stringify(user), (error, response) => {
    //                 resolve(mail);
    //             });
    //         }
    // });

    // Find my top 5 contacts on the beta endpoint
    // client
    // .api('/me/people')
    // .version('beta')
    // .top(5)
    // .select("displayName")
    // .get((err, res) => {
    //     if (err) {
    //         console.log(err)
    //         return;
    //     }
    //     const topContacts = res.value.map((u) => {return u.displayName});
    //     console.log("Your top contacts are", topContacts.join(", "));
    // });    


    // GET 3 of my events

    // var url = '/me/calendar/calendarView?startDateTime=2017-01-01T19:00:00.0000000&endDateTime=2017-01-07T19:00:00.0000000';

var Moment = require('moment-timezone');
var today = Moment().tz('Asian/Taipei').startOf('hour').add(8, 'hours').format('YYYY-MM-DD');
var startDate = today+'T'+'00:00:00.0000000';
var endDate = today+'T'+'23:59:59.0000000';

    console.log('type '+ typeof(startDate));
    var url = '/me/calendar/calendarView?startDateTime='+ startDate.toString() + '&'+'endDateTime='+endDate.toString();
    
    console.log('date:    '+new Date());
    // var url = '/me/calendar/calendarView?startDateTime=2017-04-21T00:00:00.0000000&endDateTime=2017-04-21T23:59:59.0000000';
    console.log('first'+url.toString());

    // var url = '/me/calendar/calendarView?startDateTime=2017-04-20T00:00:00.0000000&endDateTime=2017-04-20T19:00:00.0000000';

    // var url = '/me/calendar/calendarView?startDateTime=2017-04-21T00:00:00.0000000&endDateTime=2017-04-21T23:00:00.0000000';
    client
    .api(url.toString())
    .header("Prefer", 'outlook.timezone="Asia/Taipei"')
    .top(3)
    .get((err, res) => {
        if (err) {
            console.log(err)
            return;
        }else{
            console.log(url);
            var upcomingEventNames = []

            
            for (var i=0; i<res.value.length; i++) {
                upcomingEventNames.push(JSON.stringify( res.value[i]));
            }
            
            replyMessage = 'you have '+upcomingEventNames.length+' meeting today. . ';
            
            for(var i=1; i<=upcomingEventNames.length; i++){
                replyMessage += i+'. ' + res.value[i-1].subject + ' at ' + res.value[i-1].start.dateTime.substring(res.value[i-1].start.dateTime.lastIndexOf("T")+1,res.value[i-1].start.dateTime.lastIndexOf("."))+'. . ';
            }
            if(upcomingEventNames.length>=3){
                replyMessage += 'for more, please check your alexa app';
            }
            


            console.log(JSON.stringify(res));
            
            // VoiceInsights.track("sendMailIntent", null, replyMessage, (error, response) =>{

            // });
            VoiceInsights.track("sendMailIntent", null, replyMessage, (error, response) => {
                    resolve(replyMessage);
            });
        }

        
    })



    }) //end Promise
}

