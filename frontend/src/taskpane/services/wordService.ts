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

// export const insertTranscribedText = async (text: string) => {
//   try {
//     await Word.run(async (context) => {
//       const selection = context.document.getSelection();

//       // Reset logic: Hidden text ko false karein taake text nazar aaye
//       selection.font.hidden = false;
//       selection.font.color = null;

//       // Text insert karein
//       const range = selection.insertText(text + " ", Word.InsertLocation.replace);

//       // Formatting apply karein
//       range.font.size = 13;
//       range.font.italic = true;
//       range.font.name = "Calibri";
//       range.font.color = null; // Automatic color for Dark/Light mode
//       range.font.hidden = false;

//       await context.sync();
//     });
//   } catch (error) {
//     console.error("Transcription Error:", error);
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
//     console.error("Image Error:", error);
//   }
// };

// export const finalizeReport = async () => {
//   try {
//     return await Word.run(async (context) => {
//       const contentControls = context.document.contentControls;
//       // 'cannotEdit' aur 'appearance' load karein taake state ka pata chale
//       context.load(contentControls, "tag, text, appearance, length");
//       await context.sync();

//       let count = 0;

//       // Reverse loop for safe deletion
//       for (let i = contentControls.items.length - 1; i >= 0; i--) {
//         const cc = contentControls.items[i];
//         const tag = (cc.tag || "").toLowerCase();

//         // 1. Agar ye wahi Section hai jise delete hona chahiye
//         if (tag.startsWith("sec_")) {
//           // Check karein agar content empty hai ya aapne koi specific tag lagaya hai delete karne k liye
//           if (cc.text.trim() === "" || cc.text.includes("Click or tap here")) {
//             cc.delete(false); // Delete content and control
//             count++;
//           } else {
//             cc.delete(true); // Keep text, remove boundary
//           }
//         }

//         // 2. Checkboxes ko hamesha clean karein
//         else if (tag.startsWith("chk_")) {
//           cc.delete(false);
//         }

//         // 3. Data fields ki sirf boundary hatayein
//         else if (tag.startsWith("val_") || tag.startsWith("rate_") || tag.startsWith("act_")) {
//           cc.delete(true);
//         }
//       }

//       await context.sync();
//       return { success: true, count: count };
//     });
//   } catch (e: any) {
//     return { success: false, error: e.message };
//   }
// };

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

// export const finalizeReport = async () => {
//   console.log("--- Finalize Started ---");
//   try {
//     return await Word.run(async (context) => {
//       const contentControls = context.document.contentControls;

//       // Bohat Aham: Items ki properties directly load karein
//       context.load(contentControls, "items/tag, items/text, items/font/hidden");
//       await context.sync();

//       const items = contentControls.items;
//       console.log(`Total Controls Found: ${items.length}`);

//       let deletedCount = 0;

//       // Reverse loop taake deletion mein index miss na ho
//       for (let i = items.length - 1; i >= 0; i--) {
//         const cc = items[i];
//         const tag = (cc.tag || "").toLowerCase().trim();

//         // 1. Check for Sections (sec_...)
//         if (tag.startsWith("sec_")) {
//           // Agar paragraph hidden hai (Macro se grey/hidden kiya gaya tha)
//           if (cc.font.hidden === true) {
//             console.log(`Deleting Hidden Section: ${tag}`);
//             cc.delete(false); // Poora content delete
//             deletedCount++;
//           } else {
//             // Agar visible hai to sirf blue border hatao
//             // Placeholder text (Click here...) saaf karein
//             if (cc.text.includes("Click or tap here") || cc.text.trim() === "") {
//               cc.insertText(" ", "Replace");
//             }
//             cc.delete(true);
//           }
//         }
//         // 2. Check for Checkboxes (chk_...) - Inhein hamesha khatam karna hai
//         else if (tag.startsWith("chk_")) {
//           console.log(`Removing Checkbox: ${tag}`);
//           cc.delete(false);
//         }
//         // 3. Check for Input/Table tags (val_, rate_, act_) - Sirf borders hatao
//         else if (tag.startsWith("val_") || tag.startsWith("rate_") || tag.startsWith("act_")) {
//           cc.delete(true);
//         }
//       }

//       await context.sync();
//       console.log(`--- Finalize Done. Removed ${deletedCount} sections ---`);
//       return { success: true, count: deletedCount };
//     });
//   } catch (e: any) {
//     console.error("Finalize API Error:", e);
//     return { success: false, error: e.message };
//   }
// };

// export const finalizeReport = async () => {
//   try {
//     return await Word.run(async (context) => {
//       const contentControls = context.document.contentControls;
//       // Bohat Aham: font/hidden load karna zaroori hai
//       context.load(contentControls, "items/tag, items/font/hidden");
//       await context.sync();

//       const items = contentControls.items;
//       let count = 0;

//       // Reverse loop to safely delete
//       for (let i = items.length - 1; i >= 0; i--) {
//         const cc = items[i];
//         const tag = (cc.tag || "").toLowerCase().trim();

//         // 1. Handle Sections (sec_...)
//         if (tag.startsWith("sec_")) {
//           if (cc.font.hidden === true) {
//             // Agar hidden hai (Untick tha) to Heading+Para sab delete
//             cc.delete(false); 
//             count++;
//           } else {
//             // AGAR VISIBLE HAI: James ne kaha heading delete kardo, sirf para rehne do
//             // To hum poora control wrapper urha denge. 
//             // NOTE: Agar heading sec_ tag k andar hai, to wo bhi delete ho jayegi.
//             cc.delete(true); 
//           }
//         } 
//         // 2. Handle Checkboxes (chk_...) - James wants these GONE 100%
//         else if (tag.startsWith("chk_")) {
//           // cc.delete(false) control aur uska symbol (icon) dono urha dega
//           cc.delete(false);
//         } 
//         // 3. Metadata boxes (val, rate, act) - Sirf wrapper hatao
//         else if (tag.startsWith("val_") || tag.startsWith("rate_") || tag.startsWith("act_")) {
//           cc.delete(true); 
//         }
//       }

//       await context.sync();
//       return { success: true, count: count };
//     });
//   } catch (e: any) {
//     return { success: false, error: e.message };
//   }
// };


/* global Word */

export const finalizeReport = async (): Promise<{ success: boolean; count: number; error?: string }> => {
  console.log("--- Safe Finalize Started ---");
  try {
    return await Word.run(async (context) => {
      const contentControls = context.document.contentControls;
      
      // Load properties
      context.load(contentControls, "items/tag, items/font/hidden, items/font/color");
      await context.sync();

      const items = contentControls.items;
      let removedCount = 0;

      // Reverse loop for safe deletion
      for (let i = items.length - 1; i >= 0; i--) {
        const cc = items[i] as any;
        const tag = (cc.tag || "").toLowerCase().trim();

        // SIRF 'sec_' wale tags ko check karein
        if (tag.startsWith("sec_")) {
          const isHidden = cc.font.hidden === true;
          const color = (cc.font.color || "").toUpperCase();
          const isGrey = color === "#A9A9A9" || color === "#D3D3D3" || color === "#C0C0C0" || color === "#808080";

          // Agar section hidden ya grey hai tabhi delete karein
          if (isHidden || isGrey) {
            console.log(`[Action] Deleting Hidden Section: ${tag}`);
            cc.delete(false); // Poora content (Heading + Para) delete
            removedCount++;
          }
          // ELSE wala part nikal dia hai taake visible tags mehfooz rahein
        }
        // chk_, val_, rate_, act_ ko ab ye code touch bhi nahi karega
      }

      await context.sync();
      console.log(`--- Finalize Done. Removed: ${removedCount} ---`);
      return { success: true, count: removedCount };
    });
  } catch (e: any) {
    console.error("Finalize Error:", e);
    return { success: false, count: 0, error: e.message };
  }
};

// ... baqi functions (insertImage, insertTranscribedText, syncTableData) waisay hi rehne dein
// ye kam kr rhi ha
// export const insertTranscribedText = async (text: string) => {
//   try {
//     await Word.run(async (context) => {
//       // Current selection (cursor position) pakrein
//       const selection = context.document.getSelection();
      
//       // Text ko cursor k aage insert karein (taake real-time typing lage)
//       const range = selection.insertText(text, Word.InsertLocation.end);
      
//       // James's Formatting
//       range.font.size = 11;
//       // range.font.italic = true;
//       range.font.name = "Calibri";
//       range.font.color = null; // Automatic based on theme
      
//       // Cursor ko naye text k aakhir mein le jayein taake agla lafz agay aye
//       range.select(Word.SelectionMode.end);
      
//       await context.sync();
//     });
//   } catch (error) {
//     // Console error ko ignore karein taake user disturb na ho
//     console.warn("Word API Syncing...");
//   }
// };

export const insertTranscribedText = async (text: string) => {
  try {
    await Word.run(async (context) => {
      // 1. Current cursor position hasil karein
      const selection = context.document.getSelection();
      
      // 2. Text ko cursor ke baad (after) insert karein
      const range = selection.insertText(text, Word.InsertLocation.after);
      
      // 3. Formatting apply karein
      range.font.name = "Arial";
      range.font.size = 10;
      range.font.bold = false;

      // 4. SAB SE ZAROORI: Cursor ko naye insert kiye gaye text ke end par move karein
      // Is se agla word hamesha iske agay ayega
      range.select(Word.SelectionMode.end);
      
      // Word ko foran update karein
      await context.sync();
    });
  } catch (error) {
    console.warn("Word Sync Error:", error);
  }
};

