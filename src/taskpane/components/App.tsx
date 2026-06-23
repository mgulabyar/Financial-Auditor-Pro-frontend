
// import React, { useState } from "react";
// import {
//   Box,
//   Typography,
//   Button,
  
//   Card,
//   Divider,
//   Paper,
//   CircularProgress,
//   Fade,
//   Stack,
//   LinearProgress,
//   Alert,
// } from "@mui/material";
// import {
//   Security,
//   Analytics,
//   Assignment,
//   Gavel,
//   Search,
//   Difference,
//   AccountBalanceWallet,
//   ArrowForwardIos,
// } from "@mui/icons-material";
// import axios from "axios";

// const App: React.FC = () => {
//   const [loading, setLoading] = useState(false);
//   const [results, setResults] = useState<any>(null);

//   const startAudit = async () => {
//     setLoading(true);
//     try {
//       // @ts-ignore
//       await Excel.run(async (context) => {
//         const sheet = context.workbook.worksheets.getActiveWorksheet();
//         const range = sheet.getUsedRange();
//         range.load("values");
//         await context.sync();

//         const headers = range.values[0];
//         const rows = range.values.slice(1).map((row: any) => {
//           let obj: any = {};
//           headers.forEach((h: any, i: any) => (obj[h] = row[i]));
//           return obj;
//         });

//         const res = await axios.post("http://127.0.0.1:8001/api/audit/scan", { data: rows });
//         setResults(res.data);
//       });
//     } catch (e) {
//       alert("System Error: Ensure Excel has data and Backend is running on 8001.");
//     }
//     setLoading(false);
//   };

//   // Helper function to clean AI response from symbols
//   const cleanSummary = (text: string) => {
//     return text.replace(/[*#]/g, "").trim();
//   };

//   return (
//     <Box sx={{ height: "100vh", bgcolor: "#f1f5f9", overflowY: "auto" }}>
//       {/* Header */}
//       <Paper elevation={0} sx={{ p: 2, bgcolor: "#0f172a", color: "white", borderRadius: 0, position:'sticky', top:0, zIndex: 1999}}>
//         <Stack direction="row" spacing={1.5} alignItems="center">
//           <Security sx={{ color: "#ef4444" }} />
//           <Typography variant="h6" fontWeight={800} sx={{ fontSize: "1.1rem" }}>
//             AUDIT PRO AI
//           </Typography>
//         </Stack>
//       </Paper>

//       <Box sx={{ p: 2 }}>
//         <Typography
//           fontWeight={800}
//           color="textSecondary"
//           sx={{ mb: 1.5, display: "block", letterSpacing: 1, fontSize: "12px" }}
//         >
//           AUDIT ENGINE TOOLS
//         </Typography>
//         <Stack spacing={1} sx={{ mb: 3 }}>
//           {[
//             { label: "Tax Compliance", icon: <Gavel fontSize="small" /> },
//             { label: "Fraud Detection", icon: <Search fontSize="small" /> },
//             { label: "Duplicate Finder", icon: <Difference fontSize="small" /> },
//             { label: "Balance Sheet AI", icon: <AccountBalanceWallet fontSize="small" /> },
//           ].map((tool, i) => (
//             <Paper
//               key={i}
//               variant="outlined"
//               sx={{
//                 p: 1.2,
//                 cursor: "pointer",
//                 "&:hover": { bgcolor: "#f8fafc" },
//                 borderRadius: 2,
//                 display: "flex",
//                 alignItems: "center",
//                 justifyContent: "space-between",
//               }}
//             >
//               <Stack direction="row" alignItems="center" spacing={1.5}>
//                 <Box sx={{ color: "#64748b", display: "flex" }}>{tool.icon}</Box>
//                 <Typography variant="body2" fontWeight={600} color="#1e293b">
//                   {tool.label}
//                 </Typography>
//               </Stack>
//               <ArrowForwardIos sx={{ fontSize: 10, color: "#94a3b8" }} />
//             </Paper>
//           ))}
//         </Stack>

//         <Button
//           fullWidth
//           variant="contained"
//           color="error"
//           startIcon={loading ? <CircularProgress size={18} color="inherit" /> : <Analytics />}
//           onClick={startAudit}
//           disabled={loading}
//           sx={{ py: 1, fontWeight: 700, borderRadius: 2, mb: 3 }}
//         >
//           {loading ? "SCANNING..." : "START SYSTEM AUDIT"}
//         </Button>

//         {results && (
//           <Fade in={true}>
//             <Box>
//               {/* Integrity Score */}
//               <Card
//                 sx={{ p: 2, mb: 2, borderRadius: 3, boxShadow: "0 4px 6px -1px rgb(0 0 0 / 0.1)" }}
//               >
//                 <Stack direction="row" justifyContent="space-between" sx={{ mb: 1 }}>
//                   <Typography variant="subtitle2" fontWeight={700} color="textSecondary">
//                     Integrity Score
//                   </Typography>
//                   <Typography
//                     variant="h6"
//                     fontWeight={800}
//                     color={results.statistics.health_score > 80 ? "success.main" : "error.main"}
//                   >
//                     {results.statistics.health_score}%
//                   </Typography>
//                 </Stack>
//                 <LinearProgress
//                   variant="determinate"
//                   value={results.statistics.health_score}
//                   color={results.statistics.health_score > 80 ? "success" : "error"}
//                   sx={{ height: 10, borderRadius: 5 }}
//                 />
//               </Card>

//               {/* Stats Cards */}
//               <Box sx={{ display: "flex", gap: 1.5, mb: 2 }}>
//                 {[
//                   { label: "ROWS", val: results.statistics.total_rows, color: "inherit" },
//                   {
//                     label: "DUPES",
//                     val: results.statistics.duplicates,
//                     color: results.statistics.duplicates > 0 ? "#f59e0b" : "inherit",
//                   },
//                   {
//                     label: "MISSING",
//                     val: results.statistics.missing_values,
//                     color: results.statistics.missing_values > 0 ? "#ef4444" : "inherit",
//                   },
//                 ].map((item, i) => (
//                   <Paper
//                     key={i}
//                     variant="outlined"
//                     sx={{ flex: 1, p: 1.5, textAlign: "center", borderRadius: 2 }}
//                   >
//                     <Typography variant="caption" fontWeight={700} color="textSecondary">
//                       {item.label}
//                     </Typography>
//                     <Typography variant="h6" fontWeight={800} sx={{ color: item.color }}>
//                       {item.val}
//                     </Typography>
//                   </Paper>
//                 ))}
//               </Box>

//               {/* Anomalies */}
//               {results.statistics.anomalies.length > 0 && (
//                 <Box sx={{ mb: 2 }}>
//                   {results.statistics.anomalies.map((a: any, i: number) => (
//                     <Alert
//                       severity="warning"
//                       key={i}
//                       sx={{ mb: 1, borderRadius: 2, fontSize: "0.8rem" }}
//                     >
//                       Anomaly in <strong>{a.column}</strong>. {a.suggestion}
//                     </Alert>
//                   ))}
//                 </Box>
//               )}

//               {/* AI Report Box - Cleaned Text */}
//               <Card sx={{ p: 2, borderRadius: 3, border: "1px solid #e2e8f0" }}>
//                 <Typography
//                   variant="subtitle2"
//                   fontWeight={800}
//                   sx={{ mb: 1.5, display: "flex", alignItems: "center", gap: 1 }}
//                 >
//                   <Assignment color="primary" fontSize="small" /> AI AUDITOR REPORT
//                 </Typography>
//                 <Divider sx={{ mb: 2 }} />
//                 <Typography
//                   variant="body2"
//                   sx={{
//                     fontSize: "0.85rem",
//                     lineHeight: 1.8,
//                     color: "#334155",
//                     whiteSpace: "pre-wrap",
//                   }}
//                 >
//                   {cleanSummary(results.ai_analysis)}
//                 </Typography>
//               </Card>



import React, { useState } from "react";
import {
  Box,
  Typography,
  Button,
  
  Card,
  Divider,
  Paper,
  CircularProgress,
  Fade,
  Stack,
  LinearProgress,
  Alert,
} from "@mui/material";
import {
  Security,
  Analytics,
  Assignment,
  Gavel,
  Search,
  Difference,
  AccountBalanceWallet,
  ArrowForwardIos,
} from "@mui/icons-material";
import axios from "axios";

const App: React.FC = () => {
  const [loading, setLoading] = useState(false);
  const [results, setResults] = useState<any>(null);

  const startAudit = async () => {
    setLoading(true);
    try {
      // @ts-ignore
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getUsedRange();
        range.load("values");
        await context.sync();

        const headers = range.values[0];
        const rows = range.values.slice(1).map((row: any) => {
          let obj: any = {};
          headers.forEach((h: any, i: any) => (obj[h] = row[i]));
          return obj;
        });

        const res = await axios.post("http://127.0.0.1:8001/api/audit/scan", { data: rows });
        setResults(res.data);
      });
    } catch (e) {
      alert("System Error: Ensure Excel has data and Backend is running on 8001.");
    }
    setLoading(false);
  };

  // Helper function to clean AI response from symbols
  const cleanSummary = (text: string) => {
    return text.replace(/[*#]/g, "").trim();
  };

  return (
    <Box sx={{ height: "100vh", bgcolor: "#f1f5f9", overflowY: "auto" }}>
      {/* Header */}
      <Paper elevation={0} sx={{ p: 2, bgcolor: "#0f172a", color: "white", borderRadius: 0, position:'sticky', top:0, zIndex: 1999}}>
        <Stack direction="row" spacing={1.5} alignItems="center">
          <Security sx={{ color: "#ef4444" }} />
          <Typography variant="h6" fontWeight={800} sx={{ fontSize: "1.1rem" }}>
            AUDIT PRO AI
          </Typography>
        </Stack>
      </Paper>

      <Box sx={{ p: 2 }}>
        <Typography
          fontWeight={800}
          color="textSecondary"
          sx={{ mb: 1.5, display: "block", letterSpacing: 1, fontSize: "12px" }}
        >
          AUDIT ENGINE TOOLS
        </Typography>
        <Stack spacing={1} sx={{ mb: 3 }}>
          {[
            { label: "Tax Compliance", icon: <Gavel fontSize="small" /> },
            { label: "Fraud Detection", icon: <Search fontSize="small" /> },
            { label: "Duplicate Finder", icon: <Difference fontSize="small" /> },
            { label: "Balance Sheet AI", icon: <AccountBalanceWallet fontSize="small" /> },
          ].map((tool, i) => (
            <Paper
              key={i}
              variant="outlined"
              sx={{
                p: 1.2,
                cursor: "pointer",
                "&:hover": { bgcolor: "#f8fafc" },
                borderRadius: 2,
                display: "flex",
                alignItems: "center",
                justifyContent: "space-between",
              }}
            >
              <Stack direction="row" alignItems="center" spacing={1.5}>
                <Box sx={{ color: "#64748b", display: "flex" }}>{tool.icon}</Box>
                <Typography variant="body2" fontWeight={600} color="#1e293b">
                  {tool.label}
                </Typography>
              </Stack>
              <ArrowForwardIos sx={{ fontSize: 10, color: "#94a3b8" }} />
            </Paper>
          ))}
        </Stack>

        <Button
          fullWidth
          variant="contained"
          color="error"
          startIcon={loading ? <CircularProgress size={18} color="inherit" /> : <Analytics />}
          onClick={startAudit}
          disabled={loading}
          sx={{ py: 1, fontWeight: 700, borderRadius: 2, mb: 3 }}
        >
          {loading ? "SCANNING..." : "START SYSTEM AUDIT"}
        </Button>

        {results && (
          <Fade in={true}>
            <Box>
              {/* Integrity Score */}
              <Card
                sx={{ p: 2, mb: 2, borderRadius: 3, boxShadow: "0 4px 6px -1px rgb(0 0 0 / 0.1)" }}
              >
                <Stack direction="row" justifyContent="space-between" sx={{ mb: 1 }}>
                  <Typography variant="subtitle2" fontWeight={700} color="textSecondary">
                    Integrity Score
                  </Typography>
                  <Typography
                    variant="h6"
                    fontWeight={800}
                    color={results.statistics.health_score > 80 ? "success.main" : "error.main"}
                  >
                    {results.statistics.health_score}%
                  </Typography>
                </Stack>
                <LinearProgress
                  variant="determinate"
                  value={results.statistics.health_score}
                  color={results.statistics.health_score > 80 ? "success" : "error"}
                  sx={{ height: 10, borderRadius: 5 }}
                />
              </Card>

              {/* Stats Cards */}
              <Box sx={{ display: "flex", gap: 1.5, mb: 2 }}>
                {[
                  { label: "ROWS", val: results.statistics.total_rows, color: "inherit" },
                  {
                    label: "DUPES",
                    val: results.statistics.duplicates,
                    color: results.statistics.duplicates > 0 ? "#f59e0b" : "inherit",
                  },
                  {
                    label: "MISSING",
                    val: results.statistics.missing_values,
                    color: results.statistics.missing_values > 0 ? "#ef4444" : "inherit",
                  },
                ].map((item, i) => (
                  <Paper
                    key={i}
                    variant="outlined"
                    sx={{ flex: 1, p: 1.5, textAlign: "center", borderRadius: 2 }}
                  >
                    <Typography variant="caption" fontWeight={700} color="textSecondary">
                      {item.label}
                    </Typography>
                    <Typography variant="h6" fontWeight={800} sx={{ color: item.color }}>
                      {item.val}
                    </Typography>
                  </Paper>
                ))}
              </Box>

              {/* Anomalies */}
              {results.statistics.anomalies.length > 0 && (
                <Box sx={{ mb: 2 }}>
                  {results.statistics.anomalies.map((a: any, i: number) => (
                    <Alert
                      severity="warning"
                      key={i}
                      sx={{ mb: 1, borderRadius: 2, fontSize: "0.8rem" }}
                    >
                      Anomaly in <strong>{a.column}</strong>. {a.suggestion}
                    </Alert>
                  ))}
                </Box>
              )}

              {/* AI Report Box - Cleaned Text */}
              <Card sx={{ p: 2, borderRadius: 3, border: "1px solid #e2e8f0" }}>
                <Typography
                  variant="subtitle2"
                  fontWeight={800}
                  sx={{ mb: 1.5, display: "flex", alignItems: "center", gap: 1 }}
                >
                  <Assignment color="primary" fontSize="small" /> AI AUDITOR REPORT
                </Typography>
                <Divider sx={{ mb: 2 }} />
                <Typography
                  variant="body2"
                  sx={{
                    fontSize: "0.85rem",
                    lineHeight: 1.8,
                    color: "#334155",
                    whiteSpace: "pre-wrap",
                  }}
                >
                  {cleanSummary(results.ai_analysis)}
                </Typography>
              </Card>

              <Typography
                variant="caption"
                sx={{ mt: 3, display: "block", textAlign: "center", color: "#94a3b8" }}
              >
                Report ID: {results.report_id}
              </Typography>
            </Box>
          </Fade>
        )}
      </Box>
    </Box>
  );
};

export default App;
