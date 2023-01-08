// acts as a mail item component, that gets the mail item data from the graph api response
import * as React from "react";
import styles from "./DriveItem.module.scss";
import { Text } from "@fluentui/react/lib/Text";
import { Stack } from "@fluentui/react/lib/Stack";
import { IStackTokens } from "@fluentui/react/lib/Stack";
import { Image } from "office-ui-fabric-react";

export default function DriveItem(file: any): JSX.Element {
  const url = file.file.resourceReference.webUrl;
  file = file.file.resourceVisualization;
  const stackTokens: IStackTokens = {
    childrenGap: 5,
    padding: 10,
  };

/*   //check if image is valid, if not, use a placeholder

  let a = document.createElement("a");
  a.href = file.previewImageUrl;
  const isValid = a.host && a.host !== window.location.host;

  if(!isValid){
    file.previewImageUrl = "https://placekitten.com/200/200";
  } */

  return (
    <div className={styles.container}>
      <Stack tokens={stackTokens}>
        <Stack.Item>
          <Stack tokens={stackTokens}>
            <Text  className={styles.bold}>{file.title}</Text>
            <Text className={styles.italic}>{file.type}</Text>
          </Stack>
        </Stack.Item>
        <Stack.Item align="center" tokens={stackTokens}>
          <a href={url} target="_blank" rel="noreferrer">
            <Image
              onError={({ currentTarget }) => {
                currentTarget.onerror = null; // prevents looping
                currentTarget.src = "https://placekitten.com/115/150";
              }}
              height="150px"
              src={file.previewImageUrl}
            />
          </a>
        </Stack.Item>
      </Stack>
    </div>
  );
}
