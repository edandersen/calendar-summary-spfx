import * as React from 'react';
import styles from './CalendarSummaryWebPart.module.scss';
import type { ICalendarSummaryWebPartProps } from './ICalendarSummaryWebPartProps';
import { escape } from '@microsoft/sp-lodash-subset';

export interface ICalendarSummaryWebPartState {
  eventsSummary: string;
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
    const startOfDay = new Date();
    startOfDay.setHours(0,0,0,0); // Set to start of today
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
      .get((error, response) => {
        if (error) {
          console.error("Error fetching calendar events:", error);
          return;
        }

        if (response.value.length == 0) this.setState({eventsSummary: "No events today."})
        else {
          this.setState({eventsSummary: "Next event is: " + response.value[0].subject})
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
        <h2>{escape(this.state.eventsSummary)}</h2>
      </section>
    );
  }
}
