// acts as a mail item component, that gets the mail item data from the graph api response
import * as React from "react";
import styles from "./DriveItem.module.scss";
import { Text } from "@fluentui/react/lib/Text";
import { Stack } from "@fluentui/react/lib/Stack";
import { IStackTokens } from "@fluentui/react/lib/Stack";
import { Image } from "office-ui-fabric-react";
import { Icon } from "@fluentui/react/lib/Icon";
import { getFileTypeIconProps} from "@fluentui/react-file-type-icons";
// import icons for mail, calendar, files and todo
3

export default function DriveItem(file: any): JSX.Element {
  const lastUsed = file.file.lastUsed;
  const lastAccessedDateTime = new Date(
    lastUsed.lastAccessedDateTime
  ).toLocaleString();
  const lastModifiedDateTime = new Date(
    lastUsed.lastModifiedDateTime
  ).toLocaleString();
  const lastSeenDateTime =
    lastAccessedDateTime > lastModifiedDateTime
      ? lastAccessedDateTime
      : lastModifiedDateTime;

  
  const url = file.file.resourceReference.webUrl;
  file = file.file.resourceVisualization;
  let fileType: string = file.type;
  fileType = fileType.toLowerCase();
  // set icon based on file type

  switch (fileType) {
    case "pdf":
      fileType = "pdf";
      break;
    case "excel":
      fileType = "xls";
      break;
    case "word":
      fileType = "doc";
      break;
    case "powerpoint":
      fileType = "ppt";
      break;
    case "image":
      fileType = "png";
      break;
    case "video":
      fileType = "mp4";
      break;
    case "audio":
      fileType = "mp3";
      break;
    case "text":
      fileType = "txt";
      break;
    case "zip":
      fileType = "zip";
      break;
    default:
      fileType = "docx";
      break;
  }

  const stackTokens: IStackTokens = {
    childrenGap: 5,
  };
  const imageStackTokens: IStackTokens = {
    childrenGap: 5,
  };
  const textStackTokens: IStackTokens = {
    padding: 15,
  };

  // check if hovering over container to show preview image
  const [hovering, setHovering] = React.useState(false);

  const handleMouseEnter = () : void =>  {
    setHovering(true);
  };
  const handleMouseLeave = () : void => {
    setHovering(false);
  };
  

  /*   //check if image is valid, if not, use a placeholder

  let a = document.createElement("a");
  a.href = file.previewImageUrl;
  const isValid = a.host && a.host !== window.location.host;

  if(!isValid){
    file.previewImageUrl = "https://placekitten.com/200/200";
  } */

  return (
    <div className={styles.container} onMouseEnter={handleMouseEnter} onMouseLeave={handleMouseLeave}>
      <Stack horizontal grow={3} tokens={stackTokens}>
        <Stack.Item className={`${styles.fullWidth} ${styles.relative}`} tokens={textStackTokens} >
          <Stack tokens={stackTokens}>
            <Icon
              {...getFileTypeIconProps({
                extension: fileType,
                size: 40,
                imageFileType: "png",
              })}
              className={styles.fileIcon}
            />
            <Text className={styles.bold}>{file.title}</Text>
            <Text className={styles.italic}>{file.type}</Text>
            <Text className={styles.italic}>Edited: {lastSeenDateTime}</Text>
          </Stack>
        </Stack.Item>
        <Stack.Item className={hovering ? styles.show : styles.hide} align="end" tokens={imageStackTokens}>
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
