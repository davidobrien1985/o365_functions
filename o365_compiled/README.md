# o365_slackbot

>NOTE: `master` branch is leveraging "compiled" web jobs function which is currently "under development" (yeah, I know, in master branch).
> `dev` branch uses `csx` files.

This function is written in C# and leverages Azure Functions to execute REST calls against the Microsoft Graph API.

## How it works

The Function parses the Slack payload and expects a valid UPN (User Principal Name) of a user in your Office 365 tenant.
The function will then check your current E3 license status and depending on the number of available licenses will it then try to assign an E3 license to that user and remove the E1 license.

> Currently only assignment of E3 and removal of E1 is supported.

## Configuration / Dependencies

This code expects to be triggered by a [Slack Slash command](https://api.slack.com/slash-commands). 
In your Azure Functions App Settings you will need to have added the following three variables:

```
   string clientId = GetEnvironmentVariable("clientId");
   string clientSecret = GetEnvironmentVariable("clientSecret");
   string tenantId = GetEnvironmentVariable("tenantId");
```
Check the documentation to learn how to get these values for your custom "application": (https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-integrating-applications)

## Known issues

- this function suffers from cold-start issues
	Slack expects a response back in less than 3,000ms, which a cold-start won't be able to achieve
	see TODO, will convert to compiled library

## TODO

- query different license SKUs
- convert to compiled library
- support assignment of other SKUs
- support disabling of certain plans
- add way more error handling