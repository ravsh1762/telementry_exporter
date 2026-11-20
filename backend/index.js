require("dotenv").config();
const cors = require("cors");
const express = require("express");

const router = require("./routes/routes");
const toolsRouter = require("./routes/toolsRoutes");

const PORT = process.env.PORT || 4000;

const app = express();

// the browser needs to read these headers on the Excel tool downloads
app.use(cors({ exposedHeaders: ["Content-Disposition", "X-Tool-Summary"] }));
app.use(express.json());

// telemetry generator (POST /telemetry/generate)
app.use(router);

// Excel tools (POST /tools/<name>)
app.use("/tools", toolsRouter);

// upload / unexpected errors -> JSON
app.use((err, _req, res, _next) => {
  res.status(400).json({ success: false, error: err.message || "Something went wrong." });
});

app.listen(PORT, () => console.log(`Server is listening at Port: ${PORT}`));
