// export const insertTranscribedText = async (text: string) => {
//   try {
//     await Word.run(async (context) => {
//       const selection = context.document.getSelection();

//       const range = selection.insertText(text + " ", Word.InsertLocation.replace);

//       range.font.size = 13;
//       range.font.italic = true;
//       range.font.name = "Calibri";

//       range.font.color = null;
//       range.font.highlightColor = null;

//       await context.sync();
//     });
//   } catch (error) {
//     console.error("Transcription Insert Error:", error);
//   }
// };

// export const insertImageInWord = async (base64Image: string) => {
//   try {
//     await Word.run(async (context) => {
//       const cleanBase64 = base64Image.split(",")[1] || base64Image;
//       const selection = context.document.getSelection();
//       const image = selection.insertInlinePictureFromBase64(
//         cleanBase64,
//         Word.InsertLocation.replace
//       );
//       image.width = 400;
//       image.height = 300;
//       await context.sync();
//     });
//   } catch (error) {
//     console.error("Image Insert Error:", error);
//   }
// };

// export const finalizeReport = async () => {
//   try {
//     return await Word.run(async (context) => {
//       const contentControls = context.document.contentControls;
//       context.load(contentControls, "items");
//       await context.sync();

//       let count = 0;
//       for (let i = contentControls.items.length - 1; i >= 0; i--) {
//         const cc = contentControls.items[i];
//         const range = cc.getRange();
//         context.load(range, "font/hidden");
//         await context.sync();

//         if (range.font.hidden) {
//           cc.delete(true);
//           count++;
//         }
//       }
//       await context.sync();
//       return { success: true, count };
//     });
//   } catch (e: any) {
//     return { success: false, error: e.message };
//   }
// };

/* global Word */

// export const insertTranscribedText = async (text: string) => {
//   try {
//     await Word.run(async (context) => {
//       const selection = context.document.getSelection();
//       const range = selection.insertText(text + " ", Word.InsertLocation.replace);
//       range.font.size = 13;
//       range.font.italic = true;
//       range.font.name = "Calibri";
//       range.font.color = null;
//       await context.sync();
//     });
//   } catch (error) {
//     console.error("Transcription Insert Error:", error);
//   }
// };

// export const insertImageInWord = async (base64Image: string) => {
//   try {
//     await Word.run(async (context) => {
//       const cleanBase64 = base64Image.split(",")[1] || base64Image;
//       const selection = context.document.getSelection();
//       const image = selection.insertInlinePictureFromBase64(
//         cleanBase64,
//         Word.InsertLocation.replace
//       );
//       image.width = 400;
//       image.height = 300;
//       await context.sync();
//     });
//   } catch (error) {
//     console.error("Image Insert Error:", error);
//   }
// };

// export const finalizeReport = async () => {
//   console.log("Starting finalizeReport process...");
//   try {
//     return await Word.run(async (context) => {
//       const contentControls = context.document.contentControls;
//       // Pehle sirf tags load karein check karne k liye
//       context.load(contentControls, "items/tag");
//       await context.sync();

//       const items = contentControls.items;
//       console.log(`Found ${items.length} total content controls.`);

//       let count = 0;

//       for (let i = items.length - 1; i >= 0; i--) {
//         const cc = items[i] as any;
//         const tag = (cc.tag || "").toLowerCase();

//         if (tag.startsWith("sec_")) {
//           // Range load karein specifically visibility check karne k liye
//           const range = cc.getRange();
//           context.load(range, "font/hidden");
//           await context.sync(); // Her item k liye sync zaroori hai agar nested issues hon

//           if (range.font.hidden) {
//             cc.delete(false);
//             count++;
//           } else {
//             cc.delete(true);
//           }
//         } else if (tag.startsWith("chk_")) {
//           cc.delete(false);
//         } else if (tag.startsWith("val_") || tag.startsWith("rate_") || tag.startsWith("act_")) {
//           cc.delete(true);
//         }
//       }

//       await context.sync();
//       console.log("Finalize process completed.");
//       return { success: true, count: count };
//     });
//   } catch (e: any) {
//     console.error("Word API Error in Finalize:", e);
//     return { success: false, error: e.message };
//   }
// };

/* global Word */

export const insertTranscribedText = async (text: string) => {
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();

      // Reset logic: Hidden text ko false karein taake text nazar aaye
      selection.font.hidden = false;
      selection.font.color = null;

      // Text insert karein
      const range = selection.insertText(text + " ", Word.InsertLocation.replace);

      // Formatting apply karein
      range.font.size = 13;
      range.font.italic = true;
      range.font.name = "Calibri";
      range.font.color = null; // Automatic color for Dark/Light mode
      range.font.hidden = false;

      await context.sync();
    });
  } catch (error) {
    console.error("Transcription Error:", error);
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
    console.error("Image Error:", error);
  }
};

export const finalizeReport = async () => {
  try {
    return await Word.run(async (context) => {
      const contentControls = context.document.contentControls;
      // 'cannotEdit' aur 'appearance' load karein taake state ka pata chale
      context.load(contentControls, "tag, text, appearance, length");
      await context.sync();

      let count = 0;

      // Reverse loop for safe deletion
      for (let i = contentControls.items.length - 1; i >= 0; i--) {
        const cc = contentControls.items[i];
        const tag = (cc.tag || "").toLowerCase();

        // 1. Agar ye wahi Section hai jise delete hona chahiye
        if (tag.startsWith("sec_")) {
          // Check karein agar content empty hai ya aapne koi specific tag lagaya hai delete karne k liye
          if (cc.text.trim() === "" || cc.text.includes("Click or tap here")) {
            cc.delete(false); // Delete content and control
            count++;
          } else {
            cc.delete(true); // Keep text, remove boundary
          }
        }

        // 2. Checkboxes ko hamesha clean karein
        else if (tag.startsWith("chk_")) {
          cc.delete(false);
        }

        // 3. Data fields ki sirf boundary hatayein
        else if (tag.startsWith("val_") || tag.startsWith("rate_") || tag.startsWith("act_")) {
          cc.delete(true);
        }
      }

      await context.sync();
      return { success: true, count: count };
    });
  } catch (e: any) {
    return { success: false, error: e.message };
  }
};
