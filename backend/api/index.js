const express = require("express");
// const cors = require("cors");
const axios = require("axios");

const app = express();
// app.use(cors());
app.use(express.json());
const PORT = 5000;
const { GOOGLE_API_KEY } = process.env;

app.get("/api", async (req, res) => {
  console.log("request from Google");

  const { origins, destinations, arrival_time } = req.query;
  try {
    const response = await axios.get(
      `https://maps.googleapis.com/maps/api/distancematrix/json`,
      {
        params: {
          origins,
          destinations,
          arrival_time,
          key: GOOGLE_API_KEY,
        },
      }
    );
    res.json(response.data);
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "Error fetching distance data" });
  }
});

app.listen(PORT, () => {
  console.log(`Server is running on http://localhost:${PORT}`);
});
