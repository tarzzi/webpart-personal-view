// acts as a mail item component, that gets the mail item data from the graph api response
import * as React from "react";
import styles from "./CalendarItem.module.scss";

const returnEventJson = (event: Date) => {
  const year = event.getFullYear();
  const month = event.getMonth() + 1;
  const day = event.getDate();
  const hour = event.getHours();
  const minute = event.getMinutes();
  let minuteString = minute.toString();
  if (minute < 10) {
    minuteString = "0" + minute.toString();
  }
  // make into json format
  const json = {
    year: year,
    month: month,
    day: day,
    hour: hour,
    minute: minuteString,
  };
  return json;
}

export default function CalendarItem(event: any): JSX.Element {
  const item = event.event;
  let notSameYear = false;
  console.log(item);
  const eventStart = returnEventJson(new Date(item.start.dateTime));
  const eventEnd = returnEventJson(new Date(item.end.dateTime));
  const thisYear = new Date().getFullYear();

  // show year if needed
  if(eventStart.year !== eventEnd.year) {
    console.log("event spanning");
    notSameYear = true;
  }
  if(eventStart.year !== thisYear || eventEnd.year !== thisYear) {
    console.log("event not fully in this year");
    notSameYear = true;
  }
  // check if the event is a full day event
  const fullDayEvent = item.isAllDay;


  // check if the event is on the same day
  const startDate = `${eventStart.day}/${eventStart.month}`;
  const endDate = `${eventEnd.day}/${eventEnd.month}`;

  if (startDate === endDate) {
    // Both on same day and month  so show only one date with time
    return (
      <div className={styles.container}>
        <p>
        <span className={styles.day}>{eventStart.day} / </span>
        <span className={styles.month}>{eventStart.month}</span>
        {notSameYear && <span className={styles.year}> - {eventStart.year}</span>}
        </p>
        {!fullDayEvent && <p className={styles.time}>{eventStart.hour}:{eventStart.minute} - {eventEnd.hour}:{eventEnd.minute}</p>}
        <p className={styles.bold}>{item.subject}</p>
      </div>
    );
  }
  if(startDate !== endDate) {
    // show start and end date
    return (
      <div className={styles.container}>
       <p><span className={styles.day}>{eventStart.day}/</span><span className={styles.month}>{eventStart.month}</span><span className={styles.year}>{notSameYear && `/${eventEnd.year}`}</span></p>
       {!fullDayEvent && <p className={styles.time}>{eventStart.hour}:{eventStart.minute} - {eventEnd.hour}:{eventEnd.minute}</p>}
        <p className={styles.bold}>{item.subject}</p>
      </div>
    );
  }

  return (
    <div className={styles.container}>
      <p><span className={styles.day}>{eventStart.day}/</span><span className={styles.month}>{eventStart.month}</span><span className={styles.year}>{notSameYear && `/${eventEnd.year}`}</span></p>
        {!fullDayEvent && <p>{eventStart.hour}:{eventStart.minute} - {eventEnd.hour}:{eventEnd.minute}</p>}
      <p className={styles.bold}>{item.subject}</p>
    </div>
  );


}
