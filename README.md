# Calendar Summary SPFx Web Part Experiment

(or "Ed's first ever Sharepoint component")

This is an experimental Sharepoint Framework (SPFx) web part that:

- Loads your Outlook calendar for the rest of the day via the Graph API
- Uses the OpenAI API to generate you a nice summary of your day
- Streams the result to the screen

This web part was generated using the spfx yeoman template, and the main logic for this generation of the summary is in [CalendarSummaryWebPart.tsx](/src/webparts/calendarSummaryWebPart/components/CalendarSummaryWebPart.tsx).

Examples of the results:

Two events

![image of two events](docs/sharepoint-1.gif)

One event

![image of one events](docs/sharepoint-2.gif)

Empty calendar

![Empty calendar](docs/sharepoint-3.gif)

## How to use / install

- Download the .sppkg file from the [Releases](https://github.com/edandersen/calendar-summary-spfx/releases/) page
- Upload to Sharepoint App catalog
- Add as a web part on a site
- On the Properties panel of the web part, add a valid OpenAI API key

![Property setting](docs/sharepoint-4.png)

## Potential improvements

This is a proof of concept - ideally the API Key would not be exposed to the client app. You would want to create an app or Azure function protected by Entra ID that the client connects to, which then securely makes the Open AI connection. 

Even better would be removing the Open AI dependency entirely, but:

- Azure Open AI service is invite only at the moment
- Somehow call the Bing Copilot API for "chat" when a Microsoft 365 tenant has access to it, using the current user's credentials. This would be the best as then the solution would not require an API key or middleware.

