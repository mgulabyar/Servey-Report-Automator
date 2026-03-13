/* global Word */

export const insertImageInWord = async (base64Image: string) => {
  try {
    await Word.run(async (context) => {
      const cleanBase64 = base64Image.split(",")[1] || base64Image;
      const selection = context.document.getSelection();
      const image = selection.insertInlinePictureFromBase64(
        cleanBase64,
        Word.InsertLocation.replace
      );
      image.width = 250;
      image.height = 180;
      await context.sync();
    });
  } catch (error) {
    console.error("Image Error:", error);
  }
};

export const insertTranscribedText = async (text: string) => {
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();

      const range = selection.insertText(text, Word.InsertLocation.after);

      range.font.name = "Arial";
      range.font.size = 10;
      range.font.bold = false;

      range.select(Word.SelectionMode.end);

      await context.sync();
    });
  } catch (error) {
    console.warn("Word Sync Error:", error);
  }
};
// export const finalizeReport = async (): Promise<{ success: boolean; count: number; error?: string }> => {
//   try {
//     return await Word.run(async (context) => {
//       const contentControls = context.document.contentControls;

//       context.load(contentControls, "items/tag, items/font/hidden");
//       await context.sync();

//       const items = contentControls.items;
//       let removedCount = 0;

//       for (let i = items.length - 1; i >= 0; i--) {
//         const cc = items[i] as any;
//         const tag = (cc.tag || "").toLowerCase().trim();

//         if (tag.startsWith("sec_")) {
//           if (cc.font.hidden === true) {
//             cc.delete(false);
//             removedCount++;
//           }
//         }
//       }

//       await context.sync();
//       return { success: true, count: removedCount };
//     });
//   } catch (error: any) {
//     return { success: false, count: 0, error: error.message };
//   }
// };

//////////////////////////// iu is  sucessful ha hiden succestion delete kr ny k lie.//////
export const finalizeReport = async (): Promise<{ success: boolean; count: number; error?: string }> => {
  try {
    return await Word.run(async (context) => {
      const contentControls = context.document.contentControls;
      
      // 1. Tags aur Hidden status load karein
      context.load(contentControls, "items/tag, items/font/hidden");
      await context.sync();

      const items = contentControls.items;
      const idsToDelete: string[] = [];

      // STEP 1: Pehle sirf Hidden Sections ki ID nikaalein (Ye list banana zaroori hai)
      for (let i = 0; i < items.length; i++) {
        const cc = items[i];
        const tag = (cc.tag || "").toLowerCase().trim();

        // Agar Paragraph hidden hai, to uski ID note karlo
        if (tag.startsWith("sec_") && cc.font.hidden === true) {
          const id = tag.split("_")[1]; 
          if (id && !idsToDelete.includes(id)) {
            idsToDelete.push(id);
          }
        }
      }

      console.log("IDs Marked for full deletion:", idsToDelete);

      // STEP 2: Ab in IDs ke 'chk_' aur 'sec_' dono ko delete karein
      let removedCount = 0;
      
      // Reverse loop taake Word ka index kharab na ho
      for (let i = items.length - 1; i >= 0; i--) {
        const cc = items[i];
        const tag = (cc.tag || "").toLowerCase().trim();
        const parts = tag.split("_");
        
        if (parts.length > 1) {
          const prefix = parts[0]; 
          const id = parts[1];

          // AGAR ID match ho jaye, to chahye checkbox ho ya section, delete kar do!
          if (idsToDelete.includes(id)) {
            if (prefix === "chk" || prefix === "sec") {
              // SIRF ye ek line kafi hai, range wala panga khatam
              cc.delete(true); 
              removedCount++;
            }
          }
        }
      }

      await context.sync();
      return { success: true, count: removedCount };
    });
  } catch (error: any) {
    console.error("Finalize Error:", error);
    return { success: false, count: 0, error: error.message };
  }
};
