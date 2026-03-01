// /* global Office */
// import React, { useEffect, useState, useRef } from "react";
// import axios from "axios";
// import {
//   Box,
//   Typography,
//   Button,
//   Checkbox,
//   FormControlLabel,
//   Paper,
//   CircularProgress,
//   AppBar,
//   Toolbar,
//   IconButton,
//   Snackbar,
//   Alert,
//   Tabs,
//   Tab,
//   Stack,
// } from "@mui/material";
// import RefreshIcon from "@mui/icons-material/Refresh";
// import MicIcon from "@mui/icons-material/Mic";
// import StopIcon from "@mui/icons-material/Stop";
// import PhotoCameraIcon from "@mui/icons-material/PhotoCamera";
// import AlternateEmailIcon from "@mui/icons-material/AlternateEmail";
// import SyncAltIcon from "@mui/icons-material/SyncAlt";
// import {
//   getDocumentSections,
//   toggleVisibility,
//   finalizeReport,
//   insertImageInWord,
//   insertTranscribedText,
//   syncTableData,
// } from "../services/wordService";

// const NAVY = "#123048";
// // const BACKEND_URL = "http://localhost:5000/api/transcribe";
// const BACKEND_URL = "https://survey-report-api.vercel.app/api/transcribe";
// const App: React.FC = () => {
//   const [sections, setSections] = useState<any[]>([]);
//   const [selectedTags, setSelectedTags] = useState<string[]>([]);
//   const [loading, setLoading] = useState(true);
//   const [tabValue, setTabValue] = useState(0);
//   const [toast, setToast] = useState({ open: false, msg: "", severity: "success" as any });
//   const [isRecording, setIsRecording] = useState(false);
//   const [actionLoading, setActionLoading] = useState(false);

//   const mediaRecorderRef = useRef<MediaRecorder | null>(null);
//   const audioChunksRef = useRef<Blob[]>([]);

//   const loadData = async () => {
//     setLoading(true);
//     try {
//       const data = await getDocumentSections();
//       setSections(data || []);
//       setSelectedTags(data ? data.filter((s) => s.isVisible).map((s) => s.tag) : []);
//     } catch (e) {
//       showToast("Error loading sections", "error");
//     } finally {
//       setLoading(false);
//     }
//   };

//   useEffect(() => {
//     Office.onReady(() => loadData());
//   }, []);

//   const showToast = (msg: string, severity: "success" | "error" = "success") => {
//     setToast({ open: true, msg, severity });
//   };

//   const handleToggle = async (tag: string) => {
//     const isVisible = !selectedTags.includes(tag);
//     setSelectedTags((prev) => (isVisible ? [...prev, tag] : prev.filter((t) => t !== tag)));
//     await toggleVisibility(tag, isVisible);
//   };

//   const handleSync = async () => {
//     setActionLoading(true);
//     try {
//       const res: any = await syncTableData();
//       if (res && res.success) {
//         showToast("Table synced successfully", "success");
//       } else {
//         showToast("Sync completed with no changes", "success");
//       }
//     } catch (err) {
//       showToast("Sync failed. Check Word tags", "error");
//     } finally {
//       setActionLoading(false);
//     }
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
//     formData.append("file", audioBlob);
//     try {
//       const response = await axios.post(BACKEND_URL, formData, {
//         headers: { "Content-Type": "multipart/form-data" },
//       });
//       if (response.data.text) {
//         await insertTranscribedText(response.data.text);
//         showToast("Dictation inserted", "success");
//       }
//     } catch (e) {
//       showToast("AI transcription error", "error");
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
//         showToast("Photo added", "success");
//       };
//       reader.readAsDataURL(file);
//     }
//   };

//   const emailTypist = () => {
//     window.location.href = `mailto:typist@completesurveys.co.uk?subject=Report Ready&body=The report basis is ready for finalization.`;
//   };

//   const filteredSections = sections.filter((s) => {
//     const t = s.tag.toLowerCase();
//     if (tabValue === 0) return t.includes("legal") && !t.includes("tc_");
//     if (tabValue === 1) return t.includes("inst");
//     if (tabValue === 2) return t.includes("struct");
//     if (tabValue === 3) return t.includes("ext") && !t.includes("extjoin");
//     if (tabValue === 4) return t.includes("wall");
//     if (tabValue === 5)
//       return (
//         t.includes("int") &&
//         !t.includes("intjoin") &&
//         !t.includes("intdecor") &&
//         !t.includes("intdrain")
//       );
//     if (tabValue === 6) return t.includes("haz");
//     if (tabValue === 7) return t.includes("join");
//     if (tabValue === 8) return t.includes("drain");
//     if (tabValue === 9) return t.includes("util");
//     if (tabValue === 10) return t.includes("heat");
//     if (tabValue === 11) return t.includes("ground");
//     if (tabValue === 12) return t.includes("env");
//     if (tabValue === 13) return t.includes("app");
//     if (tabValue === 14) return t.includes("glo");
//     if (tabValue === 15) return t.includes("mnt");
//     if (tabValue === 16) return t.includes("tc_");
//     return false;
//   });

//   return (
//     <Box sx={{ bgcolor: "#F4F7F9", minHeight: "100vh" }}>
//       <AppBar
//         position="sticky"
//         elevation={0}
//         sx={{ bgcolor: "white", borderBottom: "1px solid #ddd" }}
//       >
//         <Toolbar variant="dense" sx={{ justifyContent: "space-between" }}>
//           <Typography sx={{ color: NAVY, fontWeight: 900, fontSize: "12px" }}>
//             AUTO-SURVEY PRO
//           </Typography>
//           <IconButton onClick={loadData} size="small">
//             <RefreshIcon fontSize="small" />
//           </IconButton>
//         </Toolbar>
//         <Tabs
//           value={tabValue}
//           onChange={(_, v) => setTabValue(v)}
//           variant="scrollable"
//           scrollButtons="auto"
//           sx={{
//             minHeight: "38px",
//             "& .MuiTab-root": { minHeight: "38px", fontSize: "9px", fontWeight: 700 },
//           }}
//         >
//           <Tab label="Legal" />
//           <Tab label="Inst." />
//           <Tab label="Struct" />
//           <Tab label="Ext." />
//           <Tab label="Walls" />
//           <Tab label="Int." />
//           <Tab label="Hazards" />
//           <Tab label="Joinery" />
//           <Tab label="Drain" />
//           <Tab label="Util" />
//           <Tab label="Heat" />
//           <Tab label="Ground" />
//           <Tab label="Env" />
//           <Tab label="Appx" />
//           <Tab label="Gloss" />
//           <Tab label="Maint" />
//           <Tab label="T&C" />
//           <Tab label="MEDIA" sx={{ color: NAVY, fontWeight: 800 }} />
//           <Tab label="WORKFLOW" sx={{ color: "#d32f2f", fontWeight: 800 }} />
//         </Tabs>
//       </AppBar>

//       <Box p={2} sx={{ pb: 15 }}>
//         {loading ? (
//           <Box textAlign="center" mt={10}>
//             <CircularProgress size={24} sx={{ color: NAVY }} />
//           </Box>
//         ) : tabValue === 17 ? (
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
//                 AI Dictation Assistant
//               </Typography>
//               <Paper variant="outlined" sx={{ p: 2, textAlign: "center", borderStyle: "dashed" }}>
//                 <Button
//                   fullWidth
//                   variant="contained"
//                   onClick={isRecording ? stopRecording : startRecording}
//                   startIcon={isRecording ? <StopIcon /> : <MicIcon />}
//                   sx={{ bgcolor: isRecording ? "#d32f2f" : NAVY }}
//                 >
//                   {isRecording ? "Stop and Transcribe" : "Start Voice Notes"}
//                 </Button>
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
//                 Photo Integration
//               </Typography>
//               <Paper variant="outlined" sx={{ p: 2, textAlign: "center", borderStyle: "dashed" }}>
//                 <Button
//                   fullWidth
//                   variant="outlined"
//                   component="label"
//                   startIcon={<PhotoCameraIcon />}
//                   sx={{ color: NAVY, borderColor: NAVY }}
//                 >
//                   Upload Site Photo{" "}
//                   <input type="file" hidden accept="image/*" onChange={handleImageUpload} />
//                 </Button>
//               </Paper>
//             </Box>
//           </Stack>
//         ) : tabValue === 18 ? (
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
//                 Data Synchronization
//               </Typography>
//               <Paper variant="outlined" sx={{ p: 2, borderLeft: `4px solid ${NAVY}` }}>
//                 <Typography variant="body2" sx={{ mb: 2, fontSize: "12px" }}>
//                   Sync ratings and actions from report body to Summary Table.
//                 </Typography>
//                 <Button
//                   fullWidth
//                   variant="outlined"
//                   onClick={handleSync}
//                   startIcon={actionLoading ? <CircularProgress size={20} /> : <SyncAltIcon />}
//                   sx={{ color: NAVY, borderColor: NAVY, fontWeight: 700 }}
//                 >
//                   Update Summary Table
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
//                 Final Submission
//               </Typography>
//               <Paper variant="outlined" sx={{ p: 2, borderLeft: `4px solid #d32f2f` }}>
//                 <Typography variant="body2" sx={{ mb: 2, fontSize: "12px" }}>
//                   Notify your typist that the report basis is ready.
//                 </Typography>
//                 <Button
//                   fullWidth
//                   variant="contained"
//                   onClick={emailTypist}
//                   startIcon={<AlternateEmailIcon />}
//                   sx={{ bgcolor: "#d32f2f", fontWeight: 700 }}
//                 >
//                   Email Report
//                 </Button>
//               </Paper>
//             </Box>
//           </Stack>
//         ) : (
//           filteredSections.map((s) => (
//             <Paper
//               key={s.tag}
//               sx={{ p: 1, mb: 1, border: "1px solid #E0E4E8", borderRadius: "6px" }}
//               elevation={0}
//             >
//               <FormControlLabel
//                 sx={{ alignItems: "flex-start", m: 0 }}
//                 control={
//                   <Checkbox
//                     size="small"
//                     checked={selectedTags.includes(s.tag)}
//                     onChange={() => handleToggle(s.tag)}
//                     sx={{ p: 0.5, "&.Mui-checked": { color: NAVY } }}
//                   />
//                 }
//                 label={
//                   <Box sx={{ ml: 0.5, mt: 0.3 }}>
//                     <Typography sx={{ fontSize: "11px", fontWeight: 800, color: NAVY }}>
//                       {s.title}
//                     </Typography>
//                     <Typography
//                       sx={{
//                         fontSize: "9px",
//                         color: "#777",
//                         display: "-webkit-box",
//                         WebkitLineClamp: 2,
//                         WebkitBoxOrient: "vertical",
//                         overflow: "hidden",
//                       }}
//                     >
//                       {s.text}
//                     </Typography>
//                   </Box>
//                 }
//               />
//             </Paper>
//           ))
//         )}
//       </Box>

//       <Box
//         sx={{
//           position: "fixed",
//           bottom: 0,
//           width: "100%",
//           p: 2,
//           bgcolor: "white",
//           borderTop: "1px solid #ddd",
//           zIndex: 100,
//         }}
//       >
//         <Button
//           variant="contained"
//           fullWidth
//           onClick={async () => {
//             setActionLoading(true);
//             const res: any = await finalizeReport();
//             if (res && res.success) {
//               showToast("Report finalized and cleaned", "success");
//               await loadData();
//             }
//             setActionLoading(false);
//           }}
//           sx={{ bgcolor: NAVY, fontWeight: 800 }}
//         >
//           {actionLoading ? <CircularProgress size={20} color="inherit" /> : "FINALIZE REPORT"}
//         </Button>
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
import React, { useEffect, useState, useRef } from "react";
import axios from "axios";
import {
  Box,
  Typography,
  Button,
  Checkbox,
  FormControlLabel,
  Paper,
  CircularProgress,
  AppBar,
  Toolbar,
  IconButton,
  Snackbar,
  Alert,
  Tabs,
  Tab,
  Stack,
} from "@mui/material";
import RefreshIcon from "@mui/icons-material/Refresh";
import MicIcon from "@mui/icons-material/Mic";
import StopIcon from "@mui/icons-material/Stop";
import PhotoCameraIcon from "@mui/icons-material/PhotoCamera";
import AlternateEmailIcon from "@mui/icons-material/AlternateEmail";
import SyncAltIcon from "@mui/icons-material/SyncAlt";
import {
  getDocumentSections,
  toggleVisibility,
  finalizeReport,
  insertImageInWord,
  insertTranscribedText,
  syncTableData,
} from "../services/wordService";

const NAVY = "#123048";
const BACKEND_URL = "https://survey-report-api.vercel.app/api/transcribe";

const App: React.FC = () => {
  const [sections, setSections] = useState<any[]>([]);
  const [selectedTags, setSelectedTags] = useState<string[]>([]);
  const [loading, setLoading] = useState(true);
  const [tabValue, setTabValue] = useState(0);
  const [toast, setToast] = useState({ open: false, msg: "", severity: "success" as any });
  const [isRecording, setIsRecording] = useState(false);
  const [actionLoading, setActionLoading] = useState(false);

  const mediaRecorderRef = useRef<MediaRecorder | null>(null);
  const audioChunksRef = useRef<Blob[]>([]);

  const loadData = async () => {
    setLoading(true);
    try {
      const data = await getDocumentSections();
      setSections(data || []);
      setSelectedTags(data ? data.filter((s) => s.isVisible).map((s) => s.tag) : []);
    } catch (e) {
      console.error(e);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    Office.onReady(() => loadData());
  }, []);

  const showToast = (msg: string, severity: "success" | "error" = "success") => {
    setToast({ open: true, msg, severity });
  };

  const handleToggle = async (tag: string) => {
    const isVisible = !selectedTags.includes(tag);
    setSelectedTags((prev) => (isVisible ? [...prev, tag] : prev.filter((t) => t !== tag)));
    await toggleVisibility(tag, isVisible);
  };

  const handleSync = async () => {
    setActionLoading(true);
    try {
      const res: any = await syncTableData();
      if (res && res.success) showToast("Table synced successfully");
    } catch (err) {
      showToast("Sync failed", "error");
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
      console.error(err);
    }
  };

  const stopRecording = () => {
    if (mediaRecorderRef.current) {
      mediaRecorderRef.current.stop();
      setIsRecording(false);
      setActionLoading(true);
    }
  };

  // const sendToBackend = async (audioBlob: Blob) => {
  //   const formData = new FormData();
  //   formData.append("file", audioBlob);
  //   try {
  //     const response = await axios.post(BACKEND_URL, formData, {
  //       headers: { "Content-Type": "multipart/form-data" },
  //     });
  //     if (response.data.text) await insertTranscribedText(response.data.text);
  //   } catch (e) {
  //     console.error(e);
  //   } finally {
  //     setActionLoading(false);
  //   }
  // };

// const sendToBackend = async (audioBlob: Blob) => {
//   const formData = new FormData();
//   formData.append("file", audioBlob, "rec.wav");

//   try {
//     const response = await axios.post(BACKEND_URL, formData, {
//       headers: { "Content-Type": "multipart/form-data" },
//       timeout: 25000, 
//     });

//     if (response?.data?.text) {
//       const transcribedText = response.data.text;
      
//       try {
//         await insertTranscribedText(transcribedText);
//         showToast("Dictation inserted successfully", "success");
//       } catch (insertError) {
//         // Sirf console mein rakhein, user ko disturb na karein
//         console.error("Insertion silent error:", insertError);
//       }
//     }
//   } catch (e: any) {
//     // 1. Agar backend ne error bheja (400, 500 etc) tabhi toast dikhao
//     if (e.response) {
//       showToast("Server error, please try again", "error");
//     } 
//     // 2. Agar timeout hua
//     else if (e.code === 'ECONNABORTED') {
//       showToast("Request timed out", "error");
//     }
//     // 3. Baki sab (Tracking Prevention, network flickers) ko ignore kar dein
//     else {
//       console.warn("Silent ignore: Network/Tracking flicker.");
//     }
//   } finally {
//     setActionLoading(false);
//   }
// };

const sendToBackend = async (audioBlob: Blob) => {
  const formData = new FormData();
  formData.append("file", audioBlob, "rec.wav");

  try {
    const response = await axios.post(BACKEND_URL, formData, {
      headers: { "Content-Type": "multipart/form-data" },
      timeout: 30000, 
    });

    if (response?.data?.text) {
      const text = response.data.text;
      
      try {
        await insertTranscribedText(text);
        showToast("Dictation inserted successfully", "success");
      } catch (insertError) {
        console.error("Insertion failed:", insertError);
      }
      
      
      setActionLoading(false);
      return; 
    }

  } catch (e: any) {
  
    if (e.response) {
      showToast(`Server Error: ${e.response.status}`, "error");
    } else {
  
      console.warn("Silent ignore of network flicker");
    }
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
        showToast("Photo added");
      };
      reader.readAsDataURL(file);
    }
  };

  const emailTypist = () => {
    window.location.href = `mailto:typist@completesurveys.co.uk?subject=Report Ready&body=The report basis is ready.`;
  };

  const filteredSections = sections.filter((s) => {
    const t = s.tag.toLowerCase();
    if (tabValue === 0) return t.includes("legal") && !t.includes("tc_");
    if (tabValue === 1) return t.includes("inst");
    if (tabValue === 2) return t.includes("struct");
    if (tabValue === 3) return t.includes("ext") && !t.includes("extjoin");
    if (tabValue === 4) return t.includes("wall");
    if (tabValue === 5)
      return (
        t.includes("int") &&
        !t.includes("intjoin") &&
        !t.includes("intdecor") &&
        !t.includes("intdrain")
      );
    if (tabValue === 6) return t.includes("haz");
    if (tabValue === 7) return t.includes("join");
    if (tabValue === 8) return t.includes("drain");
    if (tabValue === 9) return t.includes("util");
    if (tabValue === 10) return t.includes("heat");
    if (tabValue === 11) return t.includes("ground");
    if (tabValue === 12) return t.includes("env");
    if (tabValue === 13) return t.includes("app");
    if (tabValue === 14) return t.includes("glo");
    if (tabValue === 15) return t.includes("mnt");
    if (tabValue === 16) return t.includes("tc_");
    return false;
  });

  return (
    <Box sx={{ bgcolor: "#F4F7F9", minHeight: "100vh" }}>
      <AppBar
        position="sticky"
        elevation={0}
        sx={{ bgcolor: "white", borderBottom: "1px solid #ddd" }}
      >
        <Toolbar variant="dense" sx={{ justifyContent: "space-between" }}>
          <Typography sx={{ color: NAVY, fontWeight: 900, fontSize: "12px" }}>
            AUTO-SURVEY PRO
          </Typography>
          <IconButton onClick={loadData} size="small">
            <RefreshIcon fontSize="small" />
          </IconButton>
        </Toolbar>
        <Tabs
          value={tabValue}
          onChange={(_, v) => setTabValue(v)}
          variant="scrollable"
          scrollButtons="auto"
          sx={{
            minHeight: "38px",
            "& .MuiTab-root": { minHeight: "38px", fontSize: "9px", fontWeight: 700 },
          }}
        >
          <Tab label="Legal" />
          <Tab label="Inst." />
          <Tab label="Struct" />
          <Tab label="Ext." />
          <Tab label="Walls" />
          <Tab label="Int." />
          <Tab label="Hazards" />
          <Tab label="Joinery" />
          <Tab label="Drain" />
          <Tab label="Util" />
          <Tab label="Heat" />
          <Tab label="Ground" />
          <Tab label="Env" />
          <Tab label="Appx" />
          <Tab label="Gloss" />
          <Tab label="Maint" />
          <Tab label="T&C" />
          <Tab label="MEDIA" sx={{ color: NAVY, fontWeight: 800 }} />
          <Tab label="WORKFLOW" sx={{ color: "#d32f2f", fontWeight: 800 }} />
        </Tabs>
      </AppBar>

      <Box p={2} sx={{ pb: 15 }}>
        {loading ? (
          <Box textAlign="center" mt={10}>
            <CircularProgress size={24} sx={{ color: NAVY }} />
          </Box>
        ) : tabValue === 17 ? (
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
                AI Dictation Assistant
              </Typography>
              <Paper variant="outlined" sx={{ p: 2, textAlign: "center", borderStyle: "dashed" }}>
                <Button
                  fullWidth
                  variant="contained"
                  onClick={isRecording ? stopRecording : startRecording}
                  startIcon={isRecording ? <StopIcon /> : <MicIcon />}
                  sx={{ bgcolor: isRecording ? "#d32f2f" : NAVY }}
                >
                  {isRecording ? "Stop & Transcribe" : "Start Voice Notes"}
                </Button>
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
                Photo Integration
              </Typography>
              <Paper variant="outlined" sx={{ p: 2, textAlign: "center", borderStyle: "dashed" }}>
                <Button
                  fullWidth
                  variant="outlined"
                  component="label"
                  startIcon={<PhotoCameraIcon />}
                  sx={{ color: NAVY, borderColor: NAVY }}
                >
                  Upload Site Photo{" "}
                  <input type="file" hidden accept="image/*" onChange={handleImageUpload} />
                </Button>
              </Paper>
            </Box>
          </Stack>
        ) : tabValue === 18 ? (
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
                Data Synchronization
              </Typography>
              <Paper variant="outlined" sx={{ p: 2, borderLeft: `4px solid ${NAVY}` }}>
                <Typography variant="body2" sx={{ mb: 2, fontSize: "12px" }}>
                  Sync ratings and actions from body to table.
                </Typography>
                <Button
                  fullWidth
                  variant="outlined"
                  onClick={handleSync}
                  startIcon={actionLoading ? <CircularProgress size={20} /> : <SyncAltIcon />}
                  sx={{ color: NAVY, borderColor: NAVY }}
                >
                  Update Summary Table
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
                Final Submission
              </Typography>
              <Paper variant="outlined" sx={{ p: 2, borderLeft: `4px solid #d32f2f` }}>
                <Button
                  fullWidth
                  variant="contained"
                  onClick={emailTypist}
                  startIcon={<AlternateEmailIcon />}
                  sx={{ bgcolor: "#d32f2f" }}
                >
                  Email Completed Report
                </Button>
              </Paper>
            </Box>
          </Stack>
        ) : (
          filteredSections.map((s) => (
            <Paper
              key={s.tag}
              sx={{ p: 1, mb: 1, border: "1px solid #E0E4E8", borderRadius: "6px" }}
              elevation={0}
            >
              <FormControlLabel
                sx={{ alignItems: "flex-start", m: 0 }}
                control={
                  <Checkbox
                    size="small"
                    checked={selectedTags.includes(s.tag)}
                    onChange={() => handleToggle(s.tag)}
                    sx={{ p: 0.5, "&.Mui-checked": { color: NAVY } }}
                  />
                }
                label={
                  <Box sx={{ ml: 0.5, mt: 0.3 }}>
                    <Typography sx={{ fontSize: "11px", fontWeight: 800, color: NAVY }}>
                      {s.title}
                    </Typography>
                    <Typography
                      sx={{
                        fontSize: "9px",
                        color: "#777",
                        display: "-webkit-box",
                        WebkitLineClamp: 2,
                        WebkitBoxOrient: "vertical",
                        overflow: "hidden",
                      }}
                    >
                      {s.text}
                    </Typography>
                  </Box>
                }
              />
            </Paper>
          ))
        )}
      </Box>

      <Box
        sx={{
          position: "fixed",
          bottom: 0,
          width: "100%",
          p: 2,
          bgcolor: "white",
          borderTop: "1px solid #ddd",
          zIndex: 100,
        }}
      >
        <Button
          variant="contained"
          fullWidth
          onClick={async () => {
            setActionLoading(true);
            const res: any = await finalizeReport();
            if (res && res.success) showToast("Report finalized");
            await loadData();
            setActionLoading(false);
          }}
          sx={{ bgcolor: NAVY, fontWeight: 800 }}
        >
          {actionLoading ? <CircularProgress size={20} color="inherit" /> : "FINALIZE REPORT"}
        </Button>
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
