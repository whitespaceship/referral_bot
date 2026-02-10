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

function buildEmailHtml(referralUrl, unsubscribeUrl) {
  return `<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html dir="ltr" xmlns="http://www.w3.org/1999/xhtml" xmlns:o="urn:schemas-microsoft-com:office:office" lang="en">
<head>
  <meta charset="UTF-8">
  <meta content="width=device-width, initial-scale=1" name="viewport">
  <meta name="x-apple-disable-message-reformatting">
  <meta http-equiv="X-UA-Compatible" content="IE=edge">
  <meta content="telephone=no" name="format-detection">
  <title>AtomicBot Invite</title>
  <!--[if (mso 16)]><style type="text/css">@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;700&display=swap');a {text-decoration: none;}</style><![endif]-->
  <!--[if gte mso 9]><style>sup { font-size: 100% !important; }</style><![endif]-->
  <!--[if gte mso 9]><noscript><xml><o:OfficeDocumentSettings><o:AllowPNG></o:AllowPNG><o:PixelsPerInch>96</o:PixelsPerInch></o:OfficeDocumentSettings></xml></noscript><![endif]-->
  <!--[if mso]><xml><w:WordDocument xmlns:w="urn:schemas-microsoft-com:office:word"><w:DontUseAdvancedTypographyReadingMail/></w:WordDocument></xml><![endif]-->
  <style type="text/css">
    body{font-family:'Inter',-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif!important;background-color:#F3F4F6!important;color:#4B5563!important}
    h1,h2,h3,h4,h5,h6{font-family:'Inter',-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif!important;color:#111827!important;letter-spacing:-0.025em!important}
    .es-content-body{background-color:#FFFFFF!important;border-radius:16px!important;box-shadow:0 10px 15px -3px rgba(0,0,0,0.1),0 4px 6px -2px rgba(0,0,0,0.05)!important}
    #outlook a{padding:0}
    span.MsoHyperlink,span.MsoHyperlinkFollowed{color:inherit;mso-style-priority:99}
    a.es-button{mso-style-priority:100!important;text-decoration:none!important}
    a[x-apple-data-detectors],#MessageViewBody a{color:inherit!important;text-decoration:none!important;font-size:inherit!important;font-family:inherit!important;font-weight:inherit!important;line-height:inherit!important}
    @media only screen and (max-width:600px){
      .es-m-p0r{padding-right:0!important}.es-m-p0l{padding-left:0!important}
      p,a{line-height:150%!important}
      h1,h1 a{line-height:120%!important;font-size:30px!important;text-align:center}
      h2,h2 a{line-height:120%!important;font-size:24px!important;text-align:center}
      .es-m-txt-c,.es-m-txt-c h1,.es-m-txt-c h2{text-align:center!important}
      .es-content table,.es-header table,.es-footer table,.es-content,.es-footer,.es-header{width:100%!important;max-width:600px!important}
      .img-1447{width:50px!important}
      .es-text-2405 .es-text-mobile-size-24,.es-text-2405 .es-text-mobile-size-24 *{font-size:24px!important}
      .es-text-5729 .es-text-mobile-size-20,.es-text-5729 .es-text-mobile-size-20 *{font-size:20px!important}
    }
  </style>
</head>
<body class="body" style="width:100%;height:100%;font-family:'Inter',-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif;-webkit-text-size-adjust:100%;-ms-text-size-adjust:100%;padding:0;Margin:0">
  <div style="display:none;font-size:1px;color:#333333;line-height:1px;max-height:0px;max-width:0px;opacity:0;overflow:hidden;">Your exclusive invite is inside. Get early access now.&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;</div>
  <div dir="ltr" class="es-wrapper-color" lang="en" style="background-color:#F3F4F6">
    <!--[if gte mso 9]><v:background xmlns:v="urn:schemas-microsoft-com:vml" fill="t"><v:fill type="tile" color="#F3F4F6"></v:fill></v:background><![endif]-->
    <table width="100%" cellspacing="0" cellpadding="0" class="es-wrapper" role="none" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px;padding:0;Margin:0;width:100%;height:100%;background-repeat:repeat;background-position:center top;background-color:#F3F4F6">
      <tr>
        <td valign="top" style="padding:0;Margin:0">
          <table cellpadding="0" cellspacing="0" align="center" class="es-content" role="none" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px;width:100%;table-layout:fixed!important">
            <tr>
              <td align="center" style="padding:0;Margin:0;padding-top:20px;padding-bottom:20px">
                <table bgcolor="#ffffff" align="center" cellpadding="0" cellspacing="0" class="es-content-body" role="none" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px;background-color:#FFFFFF;width:600px;border-radius:16px;box-shadow:0 10px 15px -3px rgba(0,0,0,0.1),0 4px 6px -2px rgba(0,0,0,0.05);overflow:hidden">
                  <tr>
                    <td align="left" style="Margin:0;padding-top:30px;padding-right:20px;padding-bottom:10px;padding-left:20px">
                      <table cellpadding="0" cellspacing="0" align="left" class="es-left" role="none" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px;float:left">
                        <tr>
                          <td align="center" valign="top" style="padding:0;Margin:0;width:560px">
                            <table cellpadding="0" cellspacing="0" width="100%" role="presentation" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px">
                              <tr>
                                <td align="center" style="padding:0;Margin:0;padding-bottom:24px;padding-top:24px;font-size:0px"><img src="https://fyiuhmm.stripocdn.email/content/guids/CABINET_5b4cbac9f587a1f0e716af8db283aa7ec20b0f69ac71ed604f4e7263766a26cd/images/128.png" alt="" width="80" class="img-1447" style="display:block;font-size:14px;border:0;outline:none;text-decoration:none;margin:0;opacity:0.85"></td>
                              </tr>
                              <tr>
                                <td align="center" style="padding:0;Margin:0;padding-bottom:16px">
                                  <h1 class="es-m-txt-c" style="Margin:0;font-family:'Inter',-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif;mso-line-height-rule:exactly;letter-spacing:-0.025em;font-size:36px;font-style:normal;font-weight:800;line-height:42px;color:#111827">Welcome to early access!</h1>
                                </td>
                              </tr>
                              <tr>
                                <td align="center" class="es-m-p0r es-m-p0l" style="Margin:0;padding-top:8px;padding-right:48px;padding-bottom:8px;padding-left:48px">
                                  <p style="Margin:0;mso-line-height-rule:exactly;font-family:'Inter',-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif;line-height:26px;letter-spacing:0;color:#4B5563;font-size:16px">Early access to Atomic Bot confirmed.</p>
                                  <p style="Margin:0;mso-line-height-rule:exactly;font-family:'Inter',-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif;line-height:26px;letter-spacing:0;color:#4B5563;font-size:16px">Share your link with friends to get access faster.</p>
                                </td>
                              </tr>
                            </table>
                          </td>
                        </tr>
                      </table>
                    </td>
                  </tr>
                  <tr>
                    <td style="padding:0;Margin:0">
                      <table cellspacing="0" cellpadding="0" align="center" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px;width:600px;border-bottom:12px solid transparent" role="none">
                        <tr>
                          <td align="left" style="Margin:0;padding-right:20px;padding-bottom:10px;padding-left:20px;padding-top:10px">
                            <table cellpadding="0" cellspacing="0" width="100%" role="none" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px">
                              <tr>
                                <td align="center" valign="top" style="padding:0;Margin:0;width:560px">
                                  <table cellpadding="0" cellspacing="0" width="100%" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:separate;border-spacing:0px;border-radius:12px;background-color:#F9FAFB;border:1px solid #E5E7EB" role="presentation">
                                    <tr>
                                      <td align="center" class="es-text-5729" style="padding:0;Margin:0;padding-right:20px;padding-left:20px;padding-top:24px">
                                        <h2 class="es-m-txt-c es-text-mobile-size-20" style="Margin:0;font-family:'Inter',-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif;mso-line-height-rule:exactly;letter-spacing:0.05em;font-size:14px;font-style:normal;font-weight:600;line-height:20px;color:#6B7280;text-transform:uppercase">Your invite link</h2>
                                      </td>
                                    </tr>
                                    <tr>
                                      <td align="center" class="es-text-2405" style="Margin:0;padding-right:20px;padding-left:20px;padding-top:12px;padding-bottom:24px">
                                        <h1 class="es-m-txt-c es-text-mobile-size-24" style="Margin:0;font-family:ui-monospace,SFMono-Regular,Menlo,Monaco,Consolas,'Liberation Mono','Courier New',monospace;mso-line-height-rule:exactly;letter-spacing:0;font-size:20px;font-style:normal;font-weight:500;line-height:1.2;color:#4F46E5;background-color:#EEF2FF;display:inline-block;padding:8px 16px;border-radius:6px">${referralUrl}</h1>
                                      </td>
                                    </tr>
                                  </table>
                                </td>
                              </tr>
                            </table>
                          </td>
                        </tr>
                      </table>
                    </td>
                  </tr>
                </table>
              </td>
            </tr>
          </table>
          <table cellpadding="0" cellspacing="0" align="center" class="es-footer" role="none" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px;width:100%;table-layout:fixed!important;background-color:transparent;background-repeat:repeat;background-position:center top">
            <tr>
              <td align="center" style="padding:0;Margin:0">
                <table align="center" cellpadding="0" cellspacing="0" class="es-footer-body" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px;background-color:transparent;width:600px" role="none">
                  <tr>
                    <td align="left" bgcolor="#F3F4F6" style="padding:0;Margin:0;padding-right:20px;padding-left:20px;padding-top:20px;background-color:#F3F4F6">
                      <table cellpadding="0" cellspacing="0" width="100%" role="none" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px">
                        <tr>
                          <td align="left" style="padding:0;Margin:0;width:560px">
                            <table width="100%" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px">
                              <tr>
                                <td align="center" style="padding:0;Margin:0;padding-top:15px;padding-bottom:15px;font-size:0">
                                  <table cellpadding="0" cellspacing="0" class="es-table-not-adapt es-social" role="presentation" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px">
                                    <tr>
                                      <td valign="top" align="center" style="padding:0;Margin:0;padding-right:40px"><a target="_blank" href="https://x.com/atomicbot_ai" style="mso-line-height-rule:exactly;text-decoration:underline;color:#333333;font-size:12px"><img alt="X" title="X" src="https://cdn.prod.website-files.com/6981cca3fe2c3f562a2ad751/698af2542595cb0f85cbc958_x%20(2).png" height="32" style="display:block;font-size:14px;border:0;outline:none;text-decoration:none;margin:0;height:32px;width:auto"></a></td>
                                      <td valign="top" align="center" style="padding:0;Margin:0;padding-right:40px"><a target="_blank" href="https://www.instagram.com/atomicbot.ai/" style="mso-line-height-rule:exactly;text-decoration:underline;color:#333333;font-size:12px"><img title="Instagram" src="https://cdn.prod.website-files.com/6981cca3fe2c3f562a2ad751/698af254f8c8b9bb51bb61fb_inst%20(2).png" alt="Inst" height="32" style="display:block;font-size:14px;border:0;outline:none;text-decoration:none;margin:0;height:32px;width:auto"></a></td>
                                      <td align="center" valign="top" style="padding:0;Margin:0"><a href="https://discord.gg/2TXafRV69m" target="_blank" style="mso-line-height-rule:exactly;text-decoration:underline;color:#333333;font-size:12px"><img title="Discord" src="https://cdn.prod.website-files.com/6981cca3fe2c3f562a2ad751/698af2544c5acd7d65ba8f22_discord%20(2).png" height="32" alt="Discord" style="display:block;font-size:14px;border:0;outline:none;text-decoration:none;margin:0;height:32px;width:auto"></a></td>
                                    </tr>
                                  </table>
                                </td>
                              </tr>
                              <tr>
                                <td style="padding:0;Margin:0">
                                  <table width="100%" cellpadding="0" cellspacing="0" class="es-menu" role="presentation" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px">
                                    <tr class="links">
                                      <td valign="top" width="50%" align="center" style="Margin:0;border:0;padding-bottom:5px;padding-top:5px;padding-right:5px;padding-left:5px">
                                        <a target="_blank" href="https://atomicbot.ai/terms-of-service" style="mso-line-height-rule:exactly;text-decoration:none;font-family:arial,'helvetica neue',helvetica,sans-serif;display:block;color:#999999;font-size:12px">Terms of Service</a>
                                      </td>
                                      <td valign="top" width="50%" align="center" style="Margin:0;border:0;padding-bottom:5px;padding-top:5px;padding-right:5px;padding-left:5px">
                                        <a target="_blank" href="${unsubscribeUrl}" style="mso-line-height-rule:exactly;text-decoration:none;font-family:arial,'helvetica neue',helvetica,sans-serif;display:block;color:#999999;font-size:12px">Unsubscribe</a>
                                      </td>
                                    </tr>
                                  </table>
                                </td>
                              </tr>
                            </table>
                          </td>
                        </tr>
                      </table>
                    </td>
                  </tr>
                </table>
              </td>
            </tr>
          </table>
        </td>
      </tr>
    </table>
  </div>
</body>
</html>`;
}

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

    // Send welcome email with referral link
    try {
      const unsubscribeUrl = "https://referralbot-production.up.railway.app/unsubscribe?email=" + encodeURIComponent(email);
      const { data: emailData, error: emailError } = await resend.emails.send({
        from: "Atomic Bot <welcome@atomicbot.ai>",
        to: email,
        subject: "\u2705 +1 Atomic Bot! You're on the early access list.",
        html: buildEmailHtml(referralUrl, unsubscribeUrl),
      });
      if (emailError) {
        console.error("Resend email error:", emailError);
      } else {
        console.log("Email sent:", emailData);
      }
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

// Unsubscribe endpoint
app.get("/unsubscribe", async (req, res) => {
  const { email } = req.query;
  if (!email) return res.send("Invalid link");

  try {
    if (AUDIENCE_ID) {
      const contacts = await resend.contacts.list({ audienceId: AUDIENCE_ID });
      const contact = contacts.data?.data?.find((c) => c.email === email);
      if (contact) {
        await resend.contacts.update({
          id: contact.id,
          audienceId: AUDIENCE_ID,
          unsubscribed: true,
        });
      }
    }
  } catch (err) {
    console.error("Unsubscribe error:", err);
  }

  res.send(`<!DOCTYPE html><html><head><meta charset="utf-8"><title>Unsubscribed</title><style>body{font-family:system-ui;max-width:400px;margin:100px auto;text-align:center;color:#111}</style></head><body><h1>Unsubscribed</h1><p>You won't receive more emails from Atomic Bot.</p></body></html>`);
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
