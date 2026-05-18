const path = require("path");
const https = require("https");
const express = require("express");
const devCerts = require("office-addin-dev-certs");

const PORT = Number(process.env.PORT || 3000);

async function main() {
  const app = express();
  app.use((req, res, next) => {
    res.setHeader("Access-Control-Allow-Origin", "*");
    next();
  });
  app.use(express.static(path.join(__dirname)));

  const httpsOptions = await devCerts.getHttpsServerOptions();
  https.createServer(httpsOptions, app).listen(PORT, () => {
    console.log(`LaTeX Equation Add-in server running at https://localhost:${PORT}`);
    console.log("Sideload manifest.xml in PowerPoint, then open the task pane from the Home ribbon.");
  });
}

main().catch((error) => {
  console.error("Failed to start HTTPS server.");
  console.error(error);
  process.exit(1);
});
