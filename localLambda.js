// import the Lambda function 
var skill = require('./index.js');

// Mock the context input from Alexa
var context =
            {
            succeed: function () {
                console.log('------------');
                console.log('Context succeed');
            },
            fail: function (err) {
                console.log('------------');
                console.log('Context fail');
                console.log('   error:', err);
            },            
            callbackWaitsForEmptyEventLoop: true,
            logGroupName: "/aws/lambda/GraphBot",
            logStreamName: "2016/09/12/[$LATEST]87cc665c85fe4299ad5e0d5e424be2d2",
            functionName: "Graph-Snippet-Skill",
            memoryLimitInMB: "128",
            functionVersion: "$LATEST",
            invokeid: "0e347316-53d0-11e6-a3c3-b9f00dad8a1b",
            awsRequestId: "0e347316-53d0-11e6-a3c3-b9f00dad8a1b",
            invokedFunctionArn: "arn:aws:lambda:us-east-1:906973470133:function:Graph-Snippet-Skill"
            };

// Mock the event input from Alexa
var event = 
{
  "session": {
    "sessionId": "SessionId.bc74fffc-3073-4f38-ac16-6c229e3bd81e",
    "application": {
      "applicationId": "amzn1.ask.skill.b1f438ec-284c-4bb4-ae35-9f198083aeec"
    },
    "attributes": {},
    "user": {
      "userId": "amzn1.ask.account.AFCQUG25DSVPISOJRQ2AVMI4KUSWRIQXCDNLYUNQMWCPRDSNFTHIBOGVZJZIF4PYHOIRTWN5ZMZM22FWASNWGOHN5LQU54OLBAYWLRFSJMCWHLHEIFQR46INMHWDVCTERU3F46E4IB2LRQB2XWK2R63T6D6Y36BIY7V3H5TL7KD2NPI7DLFJZMNABMRJP73MBPMVLOCJFOGJBMQ",
      "accessToken": "eyJ0eXAiOiJKV1QiLCJub25jZSI6IkFRQUJBQUFBQUFEUk5ZUlEzZGhSU3JtLTRLLWFkcENKRURheS1OMFlBNXdqUXhwS1V5cG1Ka3IwZHdObmVseXhza3FHQWdEMVIwMTc5cnlTci1MN0R2Z3dFUkpvZWlreTF0c2k5MWUyQXp5bXpVYjM1VHRJMmlBQSIsImFsZyI6IlJTMjU2IiwieDV0IjoiX1VncVhHX3RNTGR1U0oxVDhjYUh4VTdjT3RjIiwia2lkIjoiX1VncVhHX3RNTGR1U0oxVDhjYUh4VTdjT3RjIn0.eyJhdWQiOiJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20iLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC83MmY5ODhiZi04NmYxLTQxYWYtOTFhYi0yZDdjZDAxMWRiNDcvIiwiaWF0IjoxNDg3ODI5NzM4LCJuYmYiOjE0ODc4Mjk3MzgsImV4cCI6MTQ4NzgzMzYzOCwiYWNyIjoiMSIsImFtciI6WyJwd2QiLCJtZmEiXSwiYXBwX2Rpc3BsYXluYW1lIjoiR3JhcGggQm90IEFsZXhhIFNraWxsIiwiYXBwaWQiOiIxYmEwMTg5Zi1lMTgzLTQ0YjgtYTE2Yy1jNzM3OGVmMmRhMmEiLCJhcHBpZGFjciI6IjEiLCJmYW1pbHlfbmFtZSI6IlN0dWJicyIsImdpdmVuX25hbWUiOiJQYXVsIiwiaW5fY29ycCI6InRydWUiLCJpcGFkZHIiOiIxNjcuMjIwLjYwLjE3NyIsIm5hbWUiOiJQYXVsIFN0dWJicyIsIm9pZCI6IjQ4ZWIzM2Q2LWYwMDktNDkxMy05OGQzLTZkZDg1MTU4NzZkYSIsIm9ucHJlbV9zaWQiOiJTLTEtNS0yMS0yMTI3NTIxMTg0LTE2MDQwMTI5MjAtMTg4NzkyNzUyNy00NzYzMDQ1IiwicGxhdGYiOiI1IiwicHVpZCI6IjEwMDMwMDAwODAxQzEzMDEiLCJzY3AiOiJNYWlsLlJlYWQgTWFpbC5TZW5kIFVzZXIuUmVhZCIsInN1YiI6ImxuVHNZamJPUjZJdV9OTlBIMnVCTjZCdWo4TkFOUWhqYTFGcU1uYXVCdzQiLCJ0aWQiOiI3MmY5ODhiZi04NmYxLTQxYWYtOTFhYi0yZDdjZDAxMWRiNDciLCJ1bmlxdWVfbmFtZSI6InBzdHViYnNAbWljcm9zb2Z0LmNvbSIsInVwbiI6InBzdHViYnNAbWljcm9zb2Z0LmNvbSIsInZlciI6IjEuMCJ9.PVaXIx8efaq_-coSKpwWPuxa-WS-jgfAZxgWYXWDWFSH7CeruHyZSWczY9u9hcHGC2bi_eFUhppVFt46EhFKFOJuze0yquCojgbddOwPBfke0AJeqhy9aRbedpDu60xEbHRu0G4l44ShR54xUdaJexGq-0vnvZgNEzgYyt0U54GHulnfbVFojDLnyHEYyQ50ynHzlBr2ts8e_gmhyg0LA6Eo1wLPRAiJkQbJTdq9DMvOLD4h_2A9tkiBwnaSuiALlHnbL8rl6vDyrPW4Xz84ElN_y1BTBeSOq53mXCb63occxaNZKPPwQmUcWRHbEV_ji6mOqxUlRgWtZiCIBeEWQA"
    },
    "new": true
  },
  "request": {
    "type": "IntentRequest",
    "requestId": "EdwRequestId.ec11af31-a531-4d62-962c-f630c52c2f32",
    "locale": "en-US",
    "timestamp": "2017-02-10T17:25:39Z",
    "intent": {
      "name": "WhoAmIIntent",
      "slots": {}
    }
  },
  "version": "1.0"
};

// Call the Lambda function 
skill.handler(event, context);
