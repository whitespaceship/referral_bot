const express = require("express");
const cors = require("cors");

const app = express();

app.use(cors({
  origin: ["https://atomicbot.ai", "https://atomicbot-ai.webflow.io"],
}));
app.use(express.json());

app.post("/signup", async (req, res) => {
  const { email, kid } = req.body;

  if (!email) {
    return res.status(400).json({ error: "email is required" });
  }

  const campaignId = process.env.KOL_CAMPAIGN_ID;
  const apiKey = process.env.KOL_API_KEY;

  if (!campaignId || !apiKey) {
    return res.status(500).json({ error: "Server misconfigured" });
  }

  const body = {
    api_key: apiKey,
    email,
    ip: req.headers["x-forwarded-for"] || req.socket.remoteAddress,
    __source: "website",
  };

  if (kid) {
    body.social_id = kid;
  }

  try {
    const response = await fetch(
      `https://api.kickofflabs.com/v2/${campaignId}`,
      {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body),
      }
    );

    const data = await response.json();

    res.json({
      code: data.social_id,
      url: data.social_url,
      referrals: data.referrals,
    });
  } catch (err) {
    console.error("KickoffLabs API error:", err.message);
    res.status(502).json({ error: "Failed to reach KickoffLabs API" });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
