import * as React from 'react';
import styles from './CalendarSummaryWebPart.module.scss';
import type { ICalendarSummaryWebPartProps } from './ICalendarSummaryWebPartProps';
import { escape } from '@microsoft/sp-lodash-subset';
import * as MicrosoftGraph from "@microsoft/microsoft-graph-types"

export interface ICalendarSummaryWebPartState {
  eventsSummary: string;
}

export interface NormalizedOutlookEvent {
  subject: string;
  attendees: string[];
  startDateTime: string;
  endDateTime: string;
  location: string;
}

export default class CalendarSummaryWebPart extends React.Component<ICalendarSummaryWebPartProps, ICalendarSummaryWebPartState> {

  constructor(props: ICalendarSummaryWebPartProps)
  {
    super(props);
    this.state = { eventsSummary: "Loading..."};
  }

  componentDidMount(): void {
    this._getCalendarSummary();
  }

  private async _getCalendarSummary(): Promise<void> {

    // check that the API Key has been added to the web part properties
    if (!this.props.apiKey)
    {
      this.setState({eventsSummary: "Add your Open AI API Key to the Properties / options of this " +
      "web part to get a nice summary of your calendar events for the day!"});
      return;
    }

    const client = await this.props.context.msGraphClientFactory.getClient('3');
    const startOfDay = new Date(); // now
    const endOfDay = new Date();
    endOfDay.setDate(endOfDay.getDate() + 1); // Set to end of today

    const start = startOfDay.toISOString();
    const end = endOfDay.toISOString();

    // first get the user's calendar events for the rest of the day
    await client
      .api(`/me/calendar/events`)
      .select('subject,start,end,location,attendees')
      .filter(`start/dateTime gt '${start}' and end/dateTime lt '${end}'`)
      .orderby('start/dateTime')
      .get(async (error, response) => {
        if (error) {
          this.setState({eventsSummary: "There was a problem getting your calendar events. Make sure this web part has Calendars.Read permission."})
          console.error("Error fetching calendar events:", error);
          return;
        }

        // normalize the Outlook event data into something more reasonable and easier for ChatGPT to parse
        const normalizedEvents: NormalizedOutlookEvent[] = [];
        const returnedEvents = response.value as MicrosoftGraph.Event[];

        returnedEvents.forEach(outlookEvent => {
            normalizedEvents.push({
              subject: outlookEvent.subject ?? "",
              location: outlookEvent.location?.displayName ?? "",
              // convert Outlook UTC time to local time in a nice textual format
              startDateTime: outlookEvent.start?.dateTime ? new Date(outlookEvent.start?.dateTime + "Z").toTimeString().split(" ")[0] : "",
              endDateTime: outlookEvent.end?.dateTime ? new Date(outlookEvent.end?.dateTime + "Z").toTimeString().split(" ")[0] : "",
              attendees: outlookEvent.attendees?.map(attendee => attendee.emailAddress?.name ?? "") ?? []
            })
        })

        // call ChatGPT with the events for the day
        const apiKey = this.props.apiKey;
        const endpoint = 'https://api.openai.com/v1/chat/completions';
        const data = {
            model: "gpt-3.5-turbo",
            messages: [
              {
                role: "user",
                content: "I will provide a list of calendar events for the rest of today in JSON format. " +
                "Summarize the schedule into around 80 words. If there are no events, say so and provide a motivating message." +
                "If there are no attendees listed, don't mention it." +
                JSON.stringify(normalizedEvents)
              }
            ],
            stream: true
        };

        const gptResponse = await fetch(endpoint, {
          method: 'POST',
          headers: {
              'Content-Type': 'application/json',
              'Authorization': `Bearer ${apiKey}`
          },
          body: JSON.stringify(data),
        });

        if (!gptResponse.ok) {
            this.setState({eventsSummary: "There was a problem summarising your calendar for the day. Maybe check your API Key."})
            return;
        }

        // stream the tokens from the ChatGPT response, updating the state as the response is generated
        const streamingReader = gptResponse.body?.pipeThrough(new TextDecoderStream()).getReader();
        if (!streamingReader) return;

        let allResponseText : string = "";

        while (true) {
          const { value, done } = await streamingReader.read();
          if (done) break;
          let dataDone = false;
          const streamingResponseLines = value.split('\n');
          streamingResponseLines.forEach((data) => {
            if (data.length === 0) return; 
            if (data[0] == ':') return; 
            if (data === 'data: [DONE]') {
              dataDone = true;
              return;
            }
            const json = JSON.parse(data.substring(6));

            if (json.choices[0].delta.content)
            {
              allResponseText += json.choices[0].delta.content; 

              this.setState({eventsSummary: allResponseText});
            }
          });
          if (dataDone) break;
        }
      });

  }

  public render(): React.ReactElement<ICalendarSummaryWebPartProps> {
    const {
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.calendarSummaryWebPart} ${hasTeamsContext ? styles.teams : ''}`}>
        <h1>Hi {escape(userDisplayName)}!</h1>
        <p>{this.state.eventsSummary}</p>
      </section>
    );
  }
}
