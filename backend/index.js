// const express = require("express");
// const cors = require("cors");
// const multer = require("multer");
// const axios = require("axios");
// const FormData = require("form-data");
// const fs = require("fs");
// require("dotenv").config();

// const app = express();
// app.use(cors()); 
// app.use(express.json());

// const upload = multer({ dest: "uploads/" });

// app.post("/api/transcribe", upload.single("file"), async (req, res) => {
//   try {
//     if (!req.file) return res.status(400).json({ error: "No audio file" });

//     const formData = new FormData();
//     formData.append(
//       "file",
//       fs.createReadStream(req.file.path),
//       "recording.wav",
//     );
//     formData.append("model", "whisper-1");

//     const response = await axios.post(
//       "https://api.openai.com/v1/audio/transcriptions",
//       formData,
//       {
//         headers: {
//           ...formData.getHeaders(),
//           Authorization: `Bearer ${process.env.OPENAI_API_KEY}`,
//         },
//       },
//     );

//     fs.unlinkSync(req.file.path);

//     res.json({ text: response.data.text });
//   } catch (error) {
//     console.error("Backend Error:", error.message);
//     res.status(500).json({ error: "AI Transcription Failed" });
//   }
// });

// const PORT = process.env.PORT || 5000;
// app.listen(PORT, () => console.log(`Server running on port ${PORT}`));


const express = require("express");
const cors = require("cors");
const multer = require("multer");
const axios = require("axios");
const FormData = require("form-data");
const fs = require("fs");
const path = require("path");
const os = require("os"); // OS module temp directory k liye
require("dotenv").config();

const app = express();

// 1. FIXED CORS: Office Add-in k liye explicit permissions
app.use(cors({
  origin: "*", 
  methods: ["POST", "GET", "OPTIONS"],
  allowedHeaders: ["Content-Type", "Authorization"]
}));

app.use(express.json());

// 2. FIXED MULTER: Vercel par sirf '/tmp' folder writable hota hai
const upload = multer({ dest: os.tmpdir() }); 

app.post("/api/transcribe", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) {
      console.error("No file received");
      return res.status(400).json({ error: "No audio file" });
    }

    console.log("File received:", req.file.path);

    const formData = new FormData();
    // OpenAI ko file bhejte waqt stream aur filename lazmi hai
    formData.append("file", fs.createReadStream(req.file.path), "recording.wav");
    formData.append("model", "whisper-1");

    const response = await axios.post(
      "https://api.openai.com/v1/audio/transcriptions",
      formData,
      {
        headers: {
          ...formData.getHeaders(),
          Authorization: `Bearer ${process.env.OPENAI_API_KEY}`,
        },
      }
    );

    // Temp file delete karein
    if (fs.existsSync(req.file.path)) {
      fs.unlinkSync(req.file.path);
    }

    console.log("Transcription successful");
    res.json({ text: response.data.text });

  } catch (error) {
    // Detailed error logging
    console.error("Backend Error:", error.response ? error.response.data : error.message);
    res.status(500).json({ 
      error: "AI Transcription Failed", 
      details: error.response ? error.response.data : error.message 
    });
  }
});

// Root path for testing
app.get("/", (req, res) => res.send("Survey Automator API is Live!"));

const PORT = process.env.PORT || 5000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));