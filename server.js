const http = require("http");
const fs = require("fs");
const path = require("path");

const PORT = 3000;

const MIME = {
  ".html": "text/html",
  ".css": "text/css",
  ".js": "application/javascript",
  ".png": "image/png",
  ".json": "application/json",
};

http.createServer((req, res) => {
  const url = req.url === "/" ? "/taskpane.html" : req.url;
  const filePath = path.join(__dirname, url.split("?")[0]);
  const ext = path.extname(filePath);
  const type = MIME[ext] || "text/plain";

  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.writeHead(404);
      res.end("Not found");
      return;
    }
    res.writeHead(200, {
      "Content-Type": type,
      "Access-Control-Allow-Origin": "*",
    });
    res.end(data);
  });
}).listen(PORT, () => {
  console.log(`StyleHelp dev server: http://localhost:${PORT}`);
  console.log("Sideload manifest.xml in Word to test.");
});
