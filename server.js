const express = require("express");
const cors = require("cors");
const { Resend } = require("resend");

const resend = new Resend(process.env.RESEND_API_KEY);
const AUDIENCE_ID = process.env.RESEND_AUDIENCE_ID;

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

    const referralUrl = "https://atomicbot.ai?kid=" + data.social_id;

    // Add contact to Resend audience
    if (AUDIENCE_ID) {
      try {
        await resend.contacts.create({
          email,
          unsubscribed: false,
          audienceId: AUDIENCE_ID,
        });
      } catch (contactErr) {
        console.error("Resend contact error:", contactErr.message);
      }
    }

    // Send welcome email with referral link via Resend REST API
    try {
      const emailRes = await fetch("https://api.resend.com/emails", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: "Bearer " + process.env.RESEND_API_KEY,
        },
        body: JSON.stringify({
          from: "Atomic Bot <welcome@atomicbot.ai>",
          to: email,
          subject: "✅ +1 Atomic Bot! You're on the early access list.",
          template: {
            id: "accesscode",
            variables: { referral_url: referralUrl },
          },
        }),
      });
      const emailResult = await emailRes.json();
      console.log("Email sent:", emailResult);
    } catch (emailErr) {
      console.error("Resend email error:", emailErr);
    }

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
