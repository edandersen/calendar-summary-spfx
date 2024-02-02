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
    const client = await this.props.context.msGraphClientFactory.getClient('3');
    const startOfDay = new Date(); // now
    const endOfDay = new Date();
    endOfDay.setDate(endOfDay.getDate() + 1); // Set to end of today

    const start = startOfDay.toISOString();
    const end = endOfDay.toISOString();

    console.log('fetching calendars');

    await client
      .api(`/me/calendar/events`)
      .select('subject,start,end,location,attendees')
      .filter(`start/dateTime gt '${start}' and end/dateTime lt '${end}'`)
      .orderby('start/dateTime')
      
      .get(async (error, response) => {
        if (error) {
          console.error("Error fetching calendar events:", error);
          return;
        }

        const normalizedEvents: NormalizedOutlookEvent[] = [];
        const returnedEvents = response.value as MicrosoftGraph.Event[];

        returnedEvents.forEach(outlookEvent => {
            normalizedEvents.push({
              subject: outlookEvent.subject ?? "",
              location: outlookEvent.location?.displayName ?? "",
              // convert Outlook UTC time to local time
              startDateTime: outlookEvent.start?.dateTime ? new Date(outlookEvent.start?.dateTime + "Z").toTimeString().split(" ")[0] : "",
              endDateTime: outlookEvent.end?.dateTime ? new Date(outlookEvent.end?.dateTime + "Z").toTimeString().split(" ")[0] : "",
              attendees: outlookEvent.attendees?.map(attendee => attendee.emailAddress?.name ?? "") ?? []
            })
        })

        console.log(returnedEvents);
        console.log(normalizedEvents);

        const apiKey = '';
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
            ]
        };

        const gptResponse = await fetch(endpoint, {
          method: 'POST',
          headers: {
              'Content-Type': 'application/json',
              'Authorization': `Bearer ${apiKey}`
          },
          body: JSON.stringify(data)
        });

        if (!gptResponse.ok) {
            throw new Error(`Error: ${gptResponse.status}`);
        }

        var textResponse = (await gptResponse.json()).choices[0].message.content;

        this.setState({eventsSummary: textResponse});

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
        <h3>{this.state.eventsSummary}</h3>
      </section>
    );
  }
}
