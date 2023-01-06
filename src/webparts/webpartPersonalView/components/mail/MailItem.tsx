// acts as a mail item component, that gets the mail item data from the graph api response
import * as React from "react";
import styles from "./MailItem.module.scss";

export default function MailItem(mailItem: any): JSX.Element {
  const item = mailItem.mailItem;
  console.log(item);

  return (
    <div className={styles.container}>
      <a href={item.webLink} target="_blank" rel="noreferrer">
        <p>
          {item.sender.emailAddress.name} - {item.receivedDateTime}
        </p>
        <p className={styles.bold}> {item.subject} </p>
        <hr className={styles.line} />
        <p className={styles.italic}> {item.bodyPreview} </p>
      </a>
    </div>
  );
}
