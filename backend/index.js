const express = require("express");
const cors = require("cors");
const multer = require("multer");
const axios = require("axios");
const FormData = require("form-data");
const fs = require("fs");
require("dotenv").config();

const app = express();
app.use(cors()); 
app.use(express.json());

const upload = multer({ dest: "uploads/" });

app.post("/api/transcribe", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: "No audio file" });

    const formData = new FormData();
    formData.append(
      "file",
      fs.createReadStream(req.file.path),
      "recording.wav",
    );
    formData.append("model", "whisper-1");

    const response = await axios.post(
      "https://api.openai.com/v1/audio/transcriptions",
      formData,
      {
        headers: {
          ...formData.getHeaders(),
          Authorization: `Bearer ${process.env.OPENAI_API_KEY}`,
        },
      },
    );

    fs.unlinkSync(req.file.path);

    res.json({ text: response.data.text });
  } catch (error) {
    console.error("Backend Error:", error.message);
    res.status(500).json({ error: "AI Transcription Failed" });
  }
});

const PORT = process.env.PORT || 5000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
