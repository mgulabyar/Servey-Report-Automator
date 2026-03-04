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
const os = require("os"); // Temp folder access k liye
require("dotenv").config();

const app = express();

// 1. CORS FIX: Sirf aapke frontend URL ko allow karein
app.use(cors({
  origin: "https://survey-report-automator.vercel.app",
  methods: ["POST", "GET", "OPTIONS"],
  allowedHeaders: ["Content-Type", "Authorization"]
}));

app.use(express.json());

// 2. VERCEL FIX: Vercel par 'uploads/' folder nahi chalta, OS temp directory use karein
const upload = multer({ dest: os.tmpdir() });

app.post("/api/transcribe", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) {
      console.error("No file received in request");
      return res.status(400).json({ error: "No audio file" });
    }

    const formData = new FormData();
    // File stream create karein
    formData.append("file", fs.createReadStream(req.file.path), "recording.wav");
    formData.append("model", "whisper-1");

    // OpenAI API call
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

    // Temp file delete karein (Cleanup)
    if (fs.existsSync(req.file.path)) {
      fs.unlinkSync(req.file.path);
    }

    // Success Response
    res.json({ text: response.data.text });

  } catch (error) {
    console.error("Backend Error Details:", error.response ? error.response.data : error.message);
    
    // Agar error aye tab bhi temp file delete karein
    if (req.file && fs.existsSync(req.file.path)) {
      fs.unlinkSync(req.file.path);
    }

    res.status(500).json({ 
      error: "AI Transcription Failed", 
      details: error.response ? error.response.data : error.message 
    });
  }
});

// Root route verification k liye
app.get("/", (req, res) => res.send("Survey Report API is running and CORS is configured!"));

const PORT = process.env.PORT || 5000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));