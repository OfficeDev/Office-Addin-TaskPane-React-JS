/* global Office console */

const insertText = async (text) => {
  // Write text to the selected slide.
  try {
    Office.context.document.setSelectedDataAsync(text, (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        throw asyncResult.error.message;
      }
    });
  } catch (error) {
    console.log("Error: " + error);
  }
};

export default insertText;
