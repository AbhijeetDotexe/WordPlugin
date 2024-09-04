Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("logStyleContentButton").onclick = logStyleContent;
    document.getElementById("displayAllStylesButton").onclick = displayAllStyles;
    document.getElementById("displayHeadingsWithNumberingButton").onclick = displayHeadingsWithNumbering;
  }
});

// Function to log content of any element with a specific style in the entire document
async function logStyleContent() {
  const styleName = document.getElementById("styleInput").value.trim();

  if (!styleName) {
    console.error("Please enter a style name.");
    return;
  }

  await Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items/style,items/text");

    await context.sync();

    let matchingParagraphs = '';

    paragraphs.items.forEach((paragraph, index) => {
      if (paragraph.style === styleName) {
        console.log(`Paragraph ${index + 1}: Style - ${paragraph.style}, Text - "${paragraph.text}"`);
        matchingParagraphs += `Paragraph ${index + 1}: ${paragraph.text}\n`;
      }
    });

    if (!matchingParagraphs) {
      console.log(`No content found with style "${styleName}".`);
    }

    await context.sync();
  }).catch(errorHandler);
}

// Function to display all text in the document along with their styles
async function displayAllStyles() {
  await Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items/style,items/text");

    await context.sync();

    paragraphs.items.forEach((paragraph, index) => {
      console.log(`Paragraph ${index + 1}: Style - ${paragraph.style}, Text - "${paragraph.text}"`);
    });

    await context.sync();
  }).catch(errorHandler);
}

// Function to display all headings with their numbering like the navigation pane
async function displayHeadingsWithNumbering() {
  await Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items/style,items/text,items/listItemInfo");

    await context.sync();

    let headingDetails = '';

    paragraphs.items.forEach((paragraph, index) => {
      const style = paragraph.style;
      const listItemInfo = paragraph.listItemInfo;
      const paragraphText = paragraph.text.trim();

      // Ensure the paragraph has both a heading style and numbering
      console.log(style)
      if (style && style.toLowerCase().includes("heading") && listItemInfo && listItemInfo.levelString) {
        const numbering = listItemInfo.levelString; // Get the numbering like "1.1", "2.3.1", etc.
        headingDetails += `Heading ${index + 1}: ${numbering} - ${paragraphText}\n`;
      }
    });

    if (headingDetails) {
      console.log("Headings with Numbering:\n" + headingDetails);
    } else {
      console.log("No headings with numbering found.");
    }

    await context.sync();
  }).catch(errorHandler);
}

// Error handling function
function errorHandler(error) {
  console.error("Error: " + error);
  if (error instanceof OfficeExtension.Error) {
    console.error("Debug info: " + JSON.stringify(error.debugInfo));
  }
}
