let port = 8080;
let express = require("express");
let app = express();

app.get("/", (req, res) => {
  res.send("Hello there!!");
});
app.listen(port, () => {
  console.log("Running on port: " + port);
});
