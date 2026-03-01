/* global Word */

export const insertReport = async (text: string) => {
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      // Pehle se mojood text clear karna hai ya naya page?
      // Filhal hum end mein insert kar rahy hain
      body.insertText(text, Word.InsertLocation.end);

      await context.sync();
    });
  } catch (error) {
    console.error("Error: " + error);
  }
};
