let addinClipboard = "";

// Office initialization
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("selectNumberingButton").onclick = selectNumbering;
    document.getElementById("selectParagraphButton").onclick = selectParagraph;
  }
});

// Function to select and copy numbering from selected text
async function selectNumbering() {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const paragraphs = selection.paragraphs;
    paragraphs.load("items");

    await context.sync();

    let found = false;
    for (let i = 0; i < paragraphs.items.length; i++) {
      const paragraph = paragraphs.items[i];
      paragraph.load("text");

      await context.sync();

      const numbering = extractNumbering(paragraph.text);

      if (numbering) {
        const startIndex = paragraph.text.indexOf(numbering);
        const endIndex = startIndex + numbering.length;

        const range = paragraph.getRange();
        const numberingRange = range.getRange("Start").expandTo(startIndex, endIndex);
        numberingRange.select();

        addinClipboard = numbering;
        console.log("Numbering selected and copied to add-in's clipboard.");

        // Update UI to show copied content
        const clipboardContentElement = document.getElementById("clipboardContent");
        if (clipboardContentElement) {
          clipboardContentElement.textContent = addinClipboard;
        } else {
          console.warn("Element with ID 'clipboardContent' not found.");
        }

        // Copy to clipboard using fallback method
        fallbackCopyToClipboard(addinClipboard);

        found = true;
        break;
      }
    }

    if (!found) {
      console.log("No numbering found in the current paragraph.");
    }

    await context.sync();
  }).catch(errorHandler);
}

// Function to select and copy entire paragraph
async function selectParagraph() {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const paragraphs = selection.paragraphs;
    paragraphs.load("items");

    await context.sync();

    if (paragraphs.items.length > 0) {
      const paragraph = paragraphs.items[0];
      paragraph.select();
      
      paragraph.load("text");
      await context.sync();

      addinClipboard = paragraph.text;
      console.log("Paragraph selected and copied to add-in's clipboard.");

      // Update UI to show copied content
      const clipboardContentElement = document.getElementById("clipboardContent");
      if (clipboardContentElement) {
        clipboardContentElement.textContent = addinClipboard;
      } else {
        console.warn("Element with ID 'clipboardContent' not found.");
      }

      // Copy to clipboard using fallback method
      fallbackCopyToClipboard(addinClipboard);
      
    } else {
      console.log("No paragraph found at the current cursor position.");
    }

    await context.sync();
  }).catch(errorHandler);
}

// Function to extract numbering from text
function extractNumbering(text) {
  const match = text.match(/\d+(\.\d+)*|\d+/); // Match numbering like "1.", "1.2", "1.2.3" etc.
  return match ? match[0] : null;
}

// Error handling function
function errorHandler(error) {
  console.error("Error: " + error);
  if (error instanceof OfficeExtension.Error) {
    console.error("Debug info: " + JSON.stringify(error.debugInfo));
  }
}

// Fallback function to copy text to clipboard
function fallbackCopyToClipboard(text) {
  // Create a temporary textarea element
  const textArea = document.createElement("textarea");
  textArea.value = text;

  // Make the textarea invisible to the user
  textArea.style.position = "fixed";
  textArea.style.opacity = 0;
  
  // Append the textarea to the body
  document.body.appendChild(textArea);

  // Select the text inside the textarea
  textArea.focus();
  textArea.select();

  try {
    // Execute the copy command
    const successful = document.execCommand("copy");
    const msg = successful ? "successful" : "unsuccessful";
    console.log(`Fallback: Copying text command was ${msg}`);
    alert("Text copied to clipboard using fallback method!");
  } catch (err) {
    console.error("Fallback: Could not copy text: ", err);
  }

  // Remove the textarea from the document
  document.body.removeChild(textArea);
}
