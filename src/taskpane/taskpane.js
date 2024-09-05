Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("logStyleContentButton").onclick = getListInfoFromSelection;
  }
});

async function getListInfoFromSelection() {
  try {
    await Word.run(async (context) => {
      console.log("Getting list info and styles from selection");

      const selection = context.document.getSelection();
      const selectionRange = selection.getRange();
      const paragraphs = selectionRange.paragraphs;
      paragraphs.load("items");
      await context.sync();

      console.log(`Total paragraphs in the selection: ${paragraphs.items.length}`);

      let currentList = [];
      let currentLevel = -1;
      let clipboardData = []; // Array to store data for clipboard

      for (let i = 0; i < paragraphs.items.length; i++) {
        const paragraph = paragraphs.items[i];
        paragraph.load("text,style,isListItem");
        await context.sync();

        const style = paragraph.style;
        const text = paragraph.text.trim();
        const isListItem = paragraph.isListItem;

        if (isListItem) {
          paragraph.listItem.load("level,listString");
          await context.sync();

          const level = paragraph.listItem.level;
          const listString = paragraph.listItem.listString || "";

          if (level <= currentLevel && currentList.length > 0) {
            clipboardData.push(currentList.join("\n"));
            currentList = [];
          }

          const indent = "  ".repeat(level);
          const formattedItem = `${listString} ${text}`;
          currentList.push(`${indent}${formattedItem}`);
          currentLevel = level;
        } else {
          if (currentList.length > 0) {
            clipboardData.push(currentList.join("\n"));
            currentList = [];
            currentLevel = -1;
          }
          clipboardData.push(text);
        }
      }

      if (currentList.length > 0) {
        clipboardData.push(currentList.join("\n"));
      }

      // Join all data into a single string and copy to clipboard
      const clipboardString = clipboardData.join("\n");
      copyToClipboard(clipboardString);
      
      console.log("All data copied to clipboard:");
      console.log(clipboardString);
    });
  } catch (error) {
    console.error("An error occurred:", error);
    if (error instanceof OfficeExtension.Error) {
      console.error("Debug info:", error.debugInfo);
    }
  }
}

function copyToClipboard(text) {
  // Create a temporary textarea element
  const textArea = document.createElement("textarea");
  textArea.value = text;
  
  // Make the textarea out of viewport
  textArea.style.position = "fixed";
  textArea.style.left = "-999999px";
  textArea.style.top = "-999999px";
  document.body.appendChild(textArea);
  
  // Select the text
  textArea.focus();
  textArea.select();

  try {
    // Execute the copy command
    const successful = document.execCommand('copy');
    const msg = successful ? 'successful' : 'unsuccessful';
    console.log('Copying text was ' + msg);
  } catch (err) {
    console.error('Unable to copy to clipboard', err);
  }

  // Remove the temporary element
  document.body.removeChild(textArea);
}
