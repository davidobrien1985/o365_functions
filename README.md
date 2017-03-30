# o365_slackbot

>NOTE: `master` branch is leveraging "compiled" web jobs function which is currently "under development" (yeah, I know, in master branch).
> `dev` branch uses `csx` files.

This function is written in C# and leverages Azure Functions to execute REST calls against the Microsoft Graph API.

## How it works

This Azure Function App supports several "functions" through one defined HTTPS endpoint accessed from Slack.

The main entry function is `receiveJobFromSlack`, it parses the Slack payload and expects a valid UPN (User Principal Name) of a user in your Office 365 tenant.
Example: `/<functionname> david.obrien@outlook.com`

### Available Slack commands
`/geto365userlicense david.obrien@outlook.com`
This function retrieves the currently assigned license SKU for the user you provided the Slack command with.

`/allocateE3O365 david.obrien@outlook.com`
This function will check your current E3 license status and depending on the number of available licenses will it then try to assign an E3 license to that user and remove the E1 license.

`/deallocateE3O365 david.obrien@outlook.com`
This function will check your current license status and depending on the number of available E1 licenses will it then try to assign an E1 license to that user and remove the E3 license.

### Workflow

The way this Function app works is quite easy:

Slack Slash command: All three set up to call the `receiveJobFromSlack` function.
The code will take the Slack payload and write it to an Azure Storage Queue depending on the actual command that used from Slack. There is one Storage Queue for each function.
Adding a message to a queue will trigger of a next function which will then actually compute the change you desire and on completion send a message back to Slack.

## Configuration / Dependencies

This code expects to be triggered by a [Slack Slash command](https://api.slack.com/slash-commands). 
In your Azure Functions App Settings you will need to have added the following five variables that will be used in your code:

```
   string clientId = GetEnvironmentVariable("clientId");
   string clientSecret = GetEnvironmentVariable("clientSecret");
   string tenantId = GetEnvironmentVariable("tenantId");
   double graphApiVersion = double.Parse(GetEnvironmentVariable("graphApiVersion"));
   string allowedChannelName = GetEnvironmentVariable("allowedChannelName");
```
Check the documentation to learn how to get these values for your custom "application": (https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-integrating-applications)

`allowedChannelName` is the Slack channel that is allowed to execute this function.
`graphApiVersion` is currently (23/03/2017) at 1.6

In addition to this the following App Settings need to also be configured for Function bindings:

`versent_slack_DocumentDB`
This is used as the "logging database" based on Azure DocumentDB. All calls from Slack to the `receiveJobFromSlack` function will be logged to this database.
The setting will have a value of `AccountEndpoint=<Endpoint for your account>;AccountKey=<Your primary access key>`



## Known issues

- this function suffers from cold-start issues
	- these should've gotten better now, but occasionally still occur

## TODO

- query different license SKUs
- convert to compiled library (DONE! on `master` branch)
- support assignment of other SKUs
- support disabling of certain plans
- add way more error handling