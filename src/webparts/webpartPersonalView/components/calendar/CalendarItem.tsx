// acts as a mail item component, that gets the mail item data from the graph api response
import * as React from "react";
import styles from "./CalendarItem.module.scss";

export default function CalendarItem(event: any): JSX.Element {
  const item = event.event;
  let isSameDay = false;
  
  const startTimeFullDateTime = new Date(item.start.dateTime).toLocaleString();
  const endTimeFullDateTime = new Date(item.end.dateTime).toLocaleString();
  const startDate =  new Date(item.start.dateTime).toDateString();
  const endDate = new Date(item.end.dateTime).toDateString();
  const startTime =  new Date(item.start.dateTime).toLocaleTimeString();
  const endTime = new Date(item.end.dateTime).toLocaleTimeString();

  if(startDate === endDate){
    // Both on same day so only show time
    isSameDay = true;
    }


  return (
    <div className={styles.container}>
        <p className={styles.bold}>{item.subject}</p>
        {isSameDay ? <p>{startDate} : {startTime} - {endTime}</p> : <p>{startTimeFullDateTime} - {endTimeFullDateTime}</p>}
    </div>
  );
}
