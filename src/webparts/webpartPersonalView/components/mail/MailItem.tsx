// acts as a mail item component, that gets the mail item data from the graph api response
import * as React from "react";
import styles from "./MailItem.module.scss";
import { Stack, IStackTokens } from "@fluentui/react/lib/Stack";
import { Text } from "@fluentui/react/lib/Text";

export default function MailItem(mailItem: any): JSX.Element {
  const item = mailItem.mailItem;
  const receieved = new Date(item.receivedDateTime).toLocaleString();

  const stackTokens: IStackTokens = {
    childrenGap: 5,
    padding: 10,
  };
  const bodyPreviewSuffixed =
    item.bodyPreview.length < 250
      ? item.bodyPreview
      : item.bodyPreview.substring(0, 250) + "...";
  console.log(item);
  return (
    <a href={item.webLink} target="_blank" rel="noreferrer">
      <Stack horizontal tokens={stackTokens} className={styles.container}>
        <Stack.Item className={styles.info} disableShrink>
          <Stack tokens={stackTokens}>
            <Text className={styles.bold}>
              {item.sender.emailAddress.name}
            </Text>
            <Text className={styles.italic}> {receieved} </Text>
          </Stack>
        </Stack.Item>
        <Stack.Item align="center" tokens={stackTokens}>
          <Text> {bodyPreviewSuffixed} </Text>
        </Stack.Item>
      </Stack>
    </a>
  );
}
