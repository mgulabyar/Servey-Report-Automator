// /* global Office */
// import React, { useState, useRef } from "react";
// import axios from "axios";
// import {
//   Box,
//   Typography,
//   Button,
//   Paper,
//   CircularProgress,
//   AppBar,
//   Toolbar,
//   Snackbar,
//   Alert,
//   Tabs,
//   Tab,
//   Stack,
// } from "@mui/material";
// import MicIcon from "@mui/icons-material/Mic";
// import StopIcon from "@mui/icons-material/Stop";
// import PhotoCameraIcon from "@mui/icons-material/PhotoCamera";
// import AlternateEmailIcon from "@mui/icons-material/AlternateEmail";
// import CheckCircleIcon from "@mui/icons-material/CheckCircle";
// import { finalizeReport, insertImageInWord, insertTranscribedText } from "../services/wordService";

// const NAVY = "#123048";
// const BACKEND_URL = "https://survey-report-api.vercel.app/api/transcribe";

// const App: React.FC = () => {
//   const [tabValue, setTabValue] = useState(0);
//   const [toast, setToast] = useState({ open: false, msg: "", severity: "success" as any });
//   const [isRecording, setIsRecording] = useState(false);
//   const [actionLoading, setActionLoading] = useState(false);

//   const mediaRecorderRef = useRef<MediaRecorder | null>(null);
//   const audioChunksRef = useRef<Blob[]>([]);

//   const showToast = (msg: string, severity: "success" | "error" = "success") => {
//     setToast({ open: true, msg, severity });
//   };

//   const startRecording = async () => {
//     try {
//       const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
//       const mediaRecorder = new MediaRecorder(stream);
//       mediaRecorderRef.current = mediaRecorder;
//       audioChunksRef.current = [];
//       mediaRecorder.ondataavailable = (e) => audioChunksRef.current.push(e.data);
//       mediaRecorder.onstop = async () => {
//         const blob = new Blob(audioChunksRef.current, { type: "audio/wav" });
//         await sendToBackend(blob);
//         stream.getTracks().forEach((t) => t.stop());
//       };
//       mediaRecorder.start();
//       setIsRecording(true);
//     } catch (err) {
//       showToast("Microphone access denied", "error");
//     }
//   };

//   const stopRecording = () => {
//     if (mediaRecorderRef.current) {
//       mediaRecorderRef.current.stop();
//       setIsRecording(false);
//       setActionLoading(true);
//     }
//   };

//   const sendToBackend = async (audioBlob: Blob) => {
//     const formData = new FormData();
//     formData.append("file", audioBlob, "recording.wav");
//     try {
//       const response = await axios.post(BACKEND_URL, formData, {
//         headers: { "Content-Type": "multipart/form-data" },
//       });
//       if (response.data && response.data.text) {
//         await insertTranscribedText(response.data.text);
//         showToast("Dictation inserted successfully");
//       }
//     } catch (e) {

//     } finally {
//       setActionLoading(false);
//     }
//   };

//   const handleImageUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
//     const file = event.target.files?.[0];
//     if (file) {
//       const reader = new FileReader();
//       reader.onload = async (e) => {
//         await insertImageInWord(e.target?.result as string);
//         showToast("Photo added successfully");
//       };
//       reader.readAsDataURL(file);
//     }
//   };

//   const handleFinalize = async () => {
//     setActionLoading(true);
//     const res: any = await finalizeReport();
//     if (res && res.success) showToast(`Report Cleaned! ${res.count} items removed`);
//     setActionLoading(false);
//   };

//   return (
//     <Box sx={{ bgcolor: "#F4F7F9", minHeight: "100vh" }}>
//       <AppBar
//         position="sticky"
//         elevation={0}
//         sx={{ bgcolor: "white", borderBottom: "1px solid #ddd" }}
//       >
//         <Toolbar variant="dense">
//           <Typography sx={{ color: NAVY, fontWeight: 900, fontSize: "14px" }}>
//             SURVEYOR TOOLS
//           </Typography>
//         </Toolbar>
//         <Tabs value={tabValue} onChange={(_, v) => setTabValue(v)} variant="fullWidth">
//           <Tab label="MEDIA" sx={{ fontWeight: 700 }} />
//           <Tab label="WORKFLOW" sx={{ fontWeight: 700 }} />
//         </Tabs>
//       </AppBar>

//       <Box p={2}>
//         {tabValue === 0 ? (
//           <Stack spacing={3}>
//             <Box>
//               <Typography
//                 sx={{
//                   fontSize: "11px",
//                   fontWeight: 900,
//                   color: NAVY,
//                   mb: 1,
//                   textTransform: "uppercase",
//                 }}
//               >
//                 AI Voice Dictation
//               </Typography>
//               <Paper
//                 variant="outlined"
//                 sx={{ p: 2, textAlign: "center", borderStyle: "dashed", bgcolor: "#fff" }}
//               >
//                 <Button
//                   fullWidth
//                   variant="contained"
//                   onClick={isRecording ? stopRecording : startRecording}
//                   startIcon={isRecording ? <StopIcon /> : <MicIcon />}
//                   sx={{ bgcolor: isRecording ? "#d32f2f" : NAVY, py: 1.2 }}
//                 >
//                   {isRecording ? "Stop & Transcribe" : "Start Speaking"}
//                 </Button>
//                 {actionLoading && <CircularProgress size={20} sx={{ mt: 1 }} />}
//               </Paper>
//             </Box>

//             <Box>
//               <Typography
//                 sx={{
//                   fontSize: "11px",
//                   fontWeight: 900,
//                   color: NAVY,
//                   mb: 1,
//                   textTransform: "uppercase",
//                 }}
//               >
//                 Visual Documentation
//               </Typography>
//               <Paper
//                 variant="outlined"
//                 sx={{ p: 2, textAlign: "center", borderStyle: "dashed", bgcolor: "#fff" }}
//               >
//                 <Button
//                   fullWidth
//                   variant="outlined"
//                   component="label"
//                   startIcon={<PhotoCameraIcon />}
//                   sx={{ color: NAVY, borderColor: NAVY, py: 1.2 }}
//                 >
//                   Insert Site Photo{" "}
//                   <input type="file" hidden accept="image/*" onChange={handleImageUpload} />
//                 </Button>
//               </Paper>
//             </Box>
//           </Stack>
//         ) : (
//           <Stack spacing={3}>
//             <Box>
//               <Typography
//                 sx={{
//                   fontSize: "11px",
//                   fontWeight: 900,
//                   color: NAVY,
//                   mb: 1,
//                   textTransform: "uppercase",
//                 }}
//               >
//                 Finalize Report
//               </Typography>
//               <Paper
//                 variant="outlined"
//                 sx={{ p: 2, borderLeft: `4px solid ${NAVY}`, bgcolor: "#fff" }}
//               >
//                 <Typography variant="body2" sx={{ mb: 2, fontSize: "13px" }}>
//                   Permanently delete hidden sections to clean the report.
//                 </Typography>
//                 <Button
//                   fullWidth
//                   variant="contained"
//                   onClick={handleFinalize}
//                   startIcon={<CheckCircleIcon />}
//                   sx={{ bgcolor: NAVY, fontWeight: 700 }}
//                 >
//                   Finalize & Clean
//                 </Button>
//               </Paper>
//             </Box>

//             <Box>
//               <Typography
//                 sx={{
//                   fontSize: "11px",
//                   fontWeight: 900,
//                   color: "#d32f2f",
//                   mb: 1,
//                   textTransform: "uppercase",
//                 }}
//               >
//                 Submission
//               </Typography>
//               <Paper
//                 variant="outlined"
//                 sx={{ p: 2, borderLeft: `4px solid #d32f2f`, bgcolor: "#fff" }}
//               >
//                 <Button
//                   fullWidth
//                   variant="contained"
//                   onClick={() => (window.location.href = `mailto:typist@completesurveys.co.uk`)}
//                   startIcon={<AlternateEmailIcon />}
//                   sx={{ bgcolor: "#d32f2f", fontWeight: 700 }}
//                 >
//                   Email to Typist
//                 </Button>
//               </Paper>
//             </Box>
//           </Stack>
//         )}
//       </Box>

//       <Snackbar
//         open={toast.open}
//         autoHideDuration={2000}
//         onClose={() => setToast({ ...toast, open: false })}
//         anchorOrigin={{ vertical: "bottom", horizontal: "right" }}
//       >
//         <Alert severity={toast.severity} variant="filled" sx={{ fontSize: "11px" }}>
//           {toast.msg}
//         </Alert>
//       </Snackbar>
//     </Box>
//   );
// };

// export default App;

/* global Office */
import React, { useState, useRef, useEffect } from "react";
import axios from "axios";
import {
  Box,
  Typography,
  Button,
  Paper,
  CircularProgress,
  AppBar,
  Toolbar,
  Snackbar,
  Alert,
  Tabs,
  Tab,
  Stack,
} from "@mui/material";
import MicIcon from "@mui/icons-material/Mic";
import StopIcon from "@mui/icons-material/Stop";
import PhotoCameraIcon from "@mui/icons-material/PhotoCamera";
import AlternateEmailIcon from "@mui/icons-material/AlternateEmail";
import CheckCircleIcon from "@mui/icons-material/CheckCircle";
import { finalizeReport, insertImageInWord, insertTranscribedText } from "../services/wordService";

const NAVY = "#123048";
const BACKEND_URL = "https://survey-report-api.vercel.app/api/transcribe";

const App: React.FC = () => {
  const [tabValue, setTabValue] = useState(0);
  const [toast, setToast] = useState({ open: false, msg: "", severity: "success" as any });
  const [isRecording, setIsRecording] = useState(false);
  const [actionLoading, setActionLoading] = useState(false);

  const mediaRecorderRef = useRef<MediaRecorder | null>(null);
  const audioChunksRef = useRef<Blob[]>([]);

  const showToast = (msg: string, severity: "success" | "error" = "success") => {
    setToast({ open: true, msg, severity });
  };

  // const handleFinalize = async () => {
  //   console.log("Finalize Button Clicked");
  //   setActionLoading(true);
  //   try {
  //     const res: any = await finalizeReport();
  //     console.log("Finalize Result:", res);
  //     if (res && res.success) {
  //       showToast(`Report Cleaned! ${res.count} items removed`);
  //     } else {
  //       showToast("Finalize failed or no items to clean", "error");
  //     }
  //   } catch (err) {
  //     console.error("Finalize Error:", err);
  //     showToast("Error executing finalize", "error");
  //   } finally {
  //     setActionLoading(false);
  //   }
  // };
  const handleFinalize = async () => {
    console.log("UI Action: Finalize Report Triggered");
    setActionLoading(true);

    try {
      const res = await finalizeReport();

      if (res && res.success) {
        console.log("Cleanup detailed response:", res);
        // Success toast with removal count
        showToast(`Report Cleaned! ${res.count} sections removed successfully.`, "success");
      } else {
        console.warn("Cleanup process finished with zero deletions or minor failure.");
        showToast(res.error || "No hidden sections found to clean.");
      }
    } catch (err) {
      console.error("Frontend exception during finalize:", err);
      showToast("Finalize command failed. Please check document logs.", "error");
    } finally {
      setActionLoading(false);
    }
  };

  const startRecording = async () => {
    try {
      const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
      const mediaRecorder = new MediaRecorder(stream);
      mediaRecorderRef.current = mediaRecorder;
      audioChunksRef.current = [];
      mediaRecorder.ondataavailable = (e) => audioChunksRef.current.push(e.data);
      mediaRecorder.onstop = async () => {
        const blob = new Blob(audioChunksRef.current, { type: "audio/wav" });
        await sendToBackend(blob);
        stream.getTracks().forEach((t) => t.stop());
      };
      mediaRecorder.start();
      setIsRecording(true);
    } catch (err) {
      showToast("Mic Access Denied", "error");
    }
  };

  const stopRecording = () => {
    if (mediaRecorderRef.current) {
      mediaRecorderRef.current.stop();
      setIsRecording(false);
      setActionLoading(true);
    }
  };

  const sendToBackend = async (audioBlob: Blob) => {
    const formData = new FormData();
    formData.append("file", audioBlob, "recording.wav");
    try {
      const response = await axios.post(BACKEND_URL, formData, {
        headers: { "Content-Type": "multipart/form-data" },
      });
      if (response.data && response.data.text) {
        await insertTranscribedText(response.data.text);
        showToast("Dictation inserted successfully");
      }
    } catch (e) {
      console.error("Backend Error:", e);
    } finally {
      setActionLoading(false);
    }
  };

  const handleImageUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = async (e) => {
        await insertImageInWord(e.target?.result as string);
        showToast("Photo added successfully");
      };
      reader.readAsDataURL(file);
    }
  };

  return (
    <Box sx={{ bgcolor: "#F4F7F9", minHeight: "100vh" }}>
      <AppBar
        position="sticky"
        elevation={0}
        sx={{ bgcolor: "white", borderBottom: "1px solid #ddd" }}
      >
        <Toolbar variant="dense">
          <Typography sx={{ color: NAVY, fontWeight: 900, fontSize: "14px" }}>
            SURVEYOR TOOLS
          </Typography>
        </Toolbar>
        <Tabs value={tabValue} onChange={(_, v) => setTabValue(v)} variant="fullWidth">
          <Tab label="MEDIA" sx={{ fontWeight: 700 }} />
          <Tab label="WORKFLOW" sx={{ fontWeight: 700 }} />
        </Tabs>
      </AppBar>

      <Box p={2}>
        {tabValue === 0 ? (
          <Stack spacing={3}>
            <Box>
              <Typography
                sx={{
                  fontSize: "11px",
                  fontWeight: 900,
                  color: NAVY,
                  mb: 1,
                  textTransform: "uppercase",
                }}
              >
                AI Voice Dictation
              </Typography>
              <Paper
                variant="outlined"
                sx={{ p: 2, textAlign: "center", borderStyle: "dashed", bgcolor: "#fff" }}
              >
                <Button
                  fullWidth
                  variant="contained"
                  onClick={isRecording ? stopRecording : startRecording}
                  startIcon={isRecording ? <StopIcon /> : <MicIcon />}
                  sx={{ bgcolor: isRecording ? "#d32f2f" : NAVY, py: 1.2 }}
                >
                  {isRecording ? "Stop & Transcribe" : "Start Speaking"}
                </Button>
                {actionLoading && <CircularProgress size={20} sx={{ mt: 1 }} />}
              </Paper>
            </Box>
            <Box>
              <Typography
                sx={{
                  fontSize: "11px",
                  fontWeight: 900,
                  color: NAVY,
                  mb: 1,
                  textTransform: "uppercase",
                }}
              >
                Visual Documentation
              </Typography>
              <Paper
                variant="outlined"
                sx={{ p: 2, textAlign: "center", borderStyle: "dashed", bgcolor: "#fff" }}
              >
                <Button
                  fullWidth
                  variant="outlined"
                  component="label"
                  startIcon={<PhotoCameraIcon />}
                  sx={{ color: NAVY, borderColor: NAVY, py: 1.2 }}
                >
                  Insert Site Photo{" "}
                  <input type="file" hidden accept="image/*" onChange={handleImageUpload} />
                </Button>
              </Paper>
            </Box>
          </Stack>
        ) : (
          <Stack spacing={3}>
            <Box>
              <Typography
                sx={{
                  fontSize: "11px",
                  fontWeight: 900,
                  color: NAVY,
                  mb: 1,
                  textTransform: "uppercase",
                }}
              >
                Finalize Report
              </Typography>
              <Paper
                variant="outlined"
                sx={{ p: 2, borderLeft: `4px solid ${NAVY}`, bgcolor: "#fff" }}
              >
                <Typography variant="body2" sx={{ mb: 2, fontSize: "13px" }}>
                  Permanently delete hidden sections to clean the report.
                </Typography>
                <Button
                  fullWidth
                  variant="contained"
                  onClick={handleFinalize}
                  startIcon={<CheckCircleIcon />}
                  sx={{ bgcolor: NAVY, fontWeight: 700 }}
                >
                  Finalize & Clean
                </Button>
              </Paper>
            </Box>
            <Box>
              <Typography
                sx={{
                  fontSize: "11px",
                  fontWeight: 900,
                  color: "#d32f2f",
                  mb: 1,
                  textTransform: "uppercase",
                }}
              >
                Submission
              </Typography>
              <Paper
                variant="outlined"
                sx={{ p: 2, borderLeft: `4px solid #d32f2f`, bgcolor: "#fff" }}
              >
                <Button
                  fullWidth
                  variant="contained"
                  onClick={() => (window.location.href = `mailto:typist@completesurveys.co.uk`)}
                  startIcon={<AlternateEmailIcon />}
                  sx={{ bgcolor: "#d32f2f", fontWeight: 700 }}
                >
                  Email to Typist
                </Button>
              </Paper>
            </Box>
          </Stack>
        )}
      </Box>

      <Snackbar
        open={toast.open}
        autoHideDuration={2000}
        onClose={() => setToast({ ...toast, open: false })}
        anchorOrigin={{ vertical: "bottom", horizontal: "right" }}
      >
        <Alert severity={toast.severity} variant="filled" sx={{ fontSize: "11px" }}>
          {toast.msg}
        </Alert>
      </Snackbar>
    </Box>
  );
};

export default App;
