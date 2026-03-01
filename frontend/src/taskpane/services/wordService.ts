// export const getDocumentSections = async () => {
//   return await Word.run(async (context) => {
//     const contentControls = context.document.contentControls;

//     context.load(contentControls, "items/tag, items/title, items/text, items/range/font/hidden");

//     try {
//       await context.sync();
//     } catch (e) {
//       console.error("Scanning Error: Document might have corrupt controls.", e);
//       return [];
//     }

//     const uniqueTags = new Set();
//     const sections: any[] = [];

//     for (let i = 0; i < contentControls.items.length; i++) {
//       const item = contentControls.items[i] as any;

//       try {
//         const tag = item.tag ? item.tag.toLowerCase() : "";

//         if (tag.startsWith("section_") && !uniqueTags.has(tag)) {
//           uniqueTags.add(tag);

//           let preview = "No preview available";
//           if (item.text) {
//             preview = item.text.substring(0, 90).replace(/[\r\n\t]+/g, " ") + "...";
//           }

//           sections.push({
//             title: item.title || "Untitled Section",
//             tag: item.tag,
//             text: preview,
//             isVisible: item.range ? !item.range.font.hidden : true,
//           });
//         }
//       } catch (err) {
//         console.warn("Skipping a problematic control at index " + i);
//         continue;
//       }
//     }
//     return sections;
//   });
// };

// export const toggleVisibility = async (tag: string, show: boolean) => {
//   try {
//     return await Word.run(async (context) => {
//       const controls = context.document.contentControls.getByTag(tag);
//       context.load(controls, "items");
//       await context.sync();

//       if (controls.items.length === 0) return { success: false, error: "Section not found" };

//       controls.items.forEach((cc: Word.ContentControl) => {
//         const range = cc.getRange();
//         range.font.hidden = !show;
//       });

//       await context.sync();
//       return { success: true };
//     });
//   } catch (e: any) {
//     return { success: false, error: e.message };
//   }
// };

// export const finalizeReport = async () => {
//   try {
//     return await Word.run(async (context) => {
//       const contentControls = context.document.contentControls;
//       context.load(contentControls, "items");
//       await context.sync();

//       let count = 0;
//       for (let i = 0; i < contentControls.items.length; i++) {
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

// // export const insertTranscribedText = async (text: string) => {
// //   await Word.run(async (context) => {
// //     const selection = context.document.getSelection();
// //     const range = selection.insertText(text + " ", Word.InsertLocation.replace);
// //     range.font.italic = true;
// //     range.font.size = 13;
// //     range.font.color = "#fff";
// //     await context.sync();
// //   });
// // };

// export const insertImageInWord = async (base64Image: string) => {
//   await Word.run(async (context) => {
//     const cleanBase64 = base64Image.split(",")[1] || base64Image;
//     const selection = context.document.getSelection();
//     const image = selection.insertInlinePictureFromBase64(cleanBase64, Word.InsertLocation.replace);
//     image.width = 400;
//     image.height = 300;
//     await context.sync();
//   });
// };

// export const syncTableData = async () => {
//   try {
//     await Word.run(async (context) => {
//       const allControls = context.document.contentControls;
//       allControls.load("tag, text");
//       await context.sync();

//       const elements = [
//         "Const",
//         "Chimney",
//         "MainRoof",
//         "SecRoof",
//         "Drainage",
//         "Eaves",
//         "Walls",
//         "Vent",
//         "Damp",
//         "WinDoor",
//         "ExtJoin",
//         "RoofConst",
//         "Ceilings",
//         "Plaster",
//         "Fireplace",
//         "Floors",
//         "Dampness",
//         "Timber",
//         "Basement",
//         "IntJoin",
//         "Sanitary",
//         "IntDecor",
//         "IntDrain",
//         "ColdWater",
//         "Gas",
//         "Electric",
//         "Heating",
//         "Insulation",
//         "Garages",
//         "Outbuild",
//         "Patios",
//         "Fences",
//       ];

//       let syncCount = 0;

//       elements.forEach((el) => {
//         const bRating = allControls.items.find(
//           (c) => c.tag?.toLowerCase() === `b_rating_${el.toLowerCase()}`
//         );
//         const bAction = allControls.items.find(
//           (c) => c.tag?.toLowerCase() === `b_action_${el.toLowerCase()}`
//         );

//         const tRating = allControls.items.find(
//           (c) => c.tag?.toLowerCase() === `t_rating_${el.toLowerCase()}`
//         );
//         const tAction = allControls.items.find(
//           (c) => c.tag?.toLowerCase() === `t_action_${el.toLowerCase()}`
//         );

//         if (bRating && tRating) {
//           const rVal = bRating.text ? bRating.text.trim() : " ";
//           tRating.insertText(rVal === "" ? " " : rVal, "Replace");
//           syncCount++;
//         }
//         if (bAction && tAction) {
//           const aVal = bAction.text ? bAction.text.trim() : " ";
//           tAction.insertText(aVal === "" ? " " : aVal, "Replace");
//           syncCount++;
//         }
//       });

//       await context.sync();
//     });
//   } catch (error) {
//     console.error("Master Sync Error:", error);
//     throw error;
//   }
// };

// export const insertTranscribedText = async (text: string) => {
//   await Word.run(async (context) => {
//     const selection = context.document.getSelection();
//     const range = selection.insertText(text + " ", Word.InsertLocation.replace);
//     range.font.italic = true;
//     range.font.color = "#fff";
//     range.font.size = 13;
//     await context.sync();
//   });
// };

/* global Word */

export const getDocumentSections = async () => {
  return await Word.run(async (context) => {
    const contentControls = context.document.contentControls;
    context.load(contentControls, "items/tag, items/title, items/text, items/range/font/hidden");
    try {
      await context.sync();
    } catch (e) {
      return [];
    }
    const uniqueTags = new Set();
    const sections: any[] = [];
    for (let i = 0; i < contentControls.items.length; i++) {
      const item = contentControls.items[i] as any;
      try {
        const tag = item.tag ? item.tag.toLowerCase() : "";
        if (tag.startsWith("section_") && !uniqueTags.has(tag)) {
          uniqueTags.add(tag);
          let preview = item.text
            ? item.text.substring(0, 90).replace(/[\r\n\t]+/g, " ") + "..."
            : "No preview";
          sections.push({
            title: item.title || "Untitled Section",
            tag: item.tag,
            text: preview,
            isVisible: item.range ? !item.range.font.hidden : true,
          });
        }
      } catch (err) {
        continue;
      }
    }
    return sections;
  });
};

export const toggleVisibility = async (tag: string, show: boolean) => {
  try {
    return await Word.run(async (context) => {
      const controls = context.document.contentControls.getByTag(tag);
      context.load(controls, "items");
      await context.sync();
      controls.items.forEach((cc: Word.ContentControl) => {
        const range = cc.getRange();
        range.font.hidden = !show;
      });
      await context.sync();
      return { success: true };
    });
  } catch (e: any) {
    return { success: false, error: e.message };
  }
};

export const finalizeReport = async () => {
  try {
    return await Word.run(async (context) => {
      const contentControls = context.document.contentControls;
      context.load(contentControls, "items");
      await context.sync();
      let count = 0;
      for (let i = 0; i < contentControls.items.length; i++) {
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

export const insertTranscribedText = async (text: string) => {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const range = selection.insertText(text + " ", Word.InsertLocation.replace);
    range.font.italic = true;
    range.font.color = "#FFFFFF";
    range.font.size = 13;
    await context.sync();
  });
};

export const insertImageInWord = async (base64Image: string) => {
  await Word.run(async (context) => {
    const cleanBase64 = base64Image.split(",")[1] || base64Image;
    const selection = context.document.getSelection();
    const image = selection.insertInlinePictureFromBase64(cleanBase64, Word.InsertLocation.replace);
    image.width = 400;
    image.height = 300;
    await context.sync();
  });
};

export const syncTableData = async () => {
  try {
    return await Word.run(async (context) => {
      const allControls = context.document.contentControls;
      allControls.load("tag, text");
      await context.sync();
      const elements = [
        "Const",
        "Chimney",
        "MainRoof",
        "SecRoof",
        "Drainage",
        "Eaves",
        "Walls",
        "Vent",
        "Damp",
        "WinDoor",
        "ExtJoin",
        "RoofConst",
        "Ceilings",
        "Plaster",
        "Fireplace",
        "Floors",
        "Dampness",
        "Timber",
        "Basement",
        "IntJoin",
        "Sanitary",
        "IntDecor",
        "IntDrain",
        "ColdWater",
        "Gas",
        "Electric",
        "Heating",
        "Insulation",
        "Garages",
        "Outbuild",
        "Patios",
        "Fences",
      ];
      let syncCount = 0;
      elements.forEach((el) => {
        const bRating = allControls.items.find(
          (c) => c.tag?.toLowerCase() === `b_rating_${el.toLowerCase()}`
        );
        const bAction = allControls.items.find(
          (c) => c.tag?.toLowerCase() === `b_action_${el.toLowerCase()}`
        );
        const tRating = allControls.items.find(
          (c) => c.tag?.toLowerCase() === `t_rating_${el.toLowerCase()}`
        );
        const tAction = allControls.items.find(
          (c) => c.tag?.toLowerCase() === `t_action_${el.toLowerCase()}`
        );
        if (bRating && tRating) {
          tRating.insertText(bRating.text || " ", "Replace");
          syncCount++;
        }
        if (bAction && tAction) {
          tAction.insertText(bAction.text || " ", "Replace");
          syncCount++;
        }
      });
      await context.sync();
      return { success: true, count: syncCount };
    });
  } catch (error) {
    return { success: false };
  }
};
