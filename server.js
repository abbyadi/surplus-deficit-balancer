const https = require("https");
const express = require("express");
const devCerts = require("office-addin-dev-certs");

const PORT = 3000;

async function startServer() {
  const app = express();
  app.use(express.static(__dirname));

  const httpsOptions = await devCerts.getHttpsServerOptions();
  https.createServer(httpsOptions, app).listen(PORT, () => {
    console.log(`Server running at https://localhost:${PORT}`);
    console.log("Press Ctrl+C to stop.");
  });
}

startServer().catch((err) => {
  console.error("Failed to start server:", err);
  process.exit(1);
});
