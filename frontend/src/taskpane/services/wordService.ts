
export const insertTranscribedText = async (text: string) => {
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
    
      const range = selection.insertText(text + " ", Word.InsertLocation.replace);
     
      range.font.size = 13;
      range.font.italic = true;
      range.font.name = "Calibri"; 


      range.font.color = null; 
      range.font.highlightColor = null;

      await context.sync();
    });
  } catch (error) {
    console.error("Transcription Insert Error:", error);
  }
};


export const insertImageInWord = async (base64Image: string) => {
  try {
    await Word.run(async (context) => {
      const cleanBase64 = base64Image.split(",")[1] || base64Image;
      const selection = context.document.getSelection();
      const image = selection.insertInlinePictureFromBase64(
        cleanBase64,
        Word.InsertLocation.replace
      );
      image.width = 400;
      image.height = 300;
      await context.sync();
    });
  } catch (error) {
    console.error("Image Insert Error:", error);
  }
};

export const finalizeReport = async () => {
  try {
    return await Word.run(async (context) => {
      const contentControls = context.document.contentControls;
      context.load(contentControls, "items");
      await context.sync();

      let count = 0;
      for (let i = contentControls.items.length - 1; i >= 0; i--) {
        const cc = contentControls.items[i];
        const range = cc.getRange();
        context.load(range, "font/hidden");
        await context.sync();

        if (range.font.hidden) {
          cc.delete(true);
          count++;
        }
      }
      await context.sync();
      return { success: true, count };
    });
  } catch (e: any) {
    return { success: false, error: e.message };
  }
};
