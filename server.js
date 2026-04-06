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

function buildEmailHtml(unsubscribeUrl) {
  return `<!DOCTYPE html
  PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html dir="ltr" xmlns="http://www.w3.org/1999/xhtml" xmlns:o="urn:schemas-microsoft-com:office:office" lang="en">

<head>
  <meta charset="UTF-8">
  <meta content="width=device-width, initial-scale=1" name="viewport">
  <meta name="x-apple-disable-message-reformatting">
  <meta http-equiv="X-UA-Compatible" content="IE=edge">
  <meta content="telephone=no" name="format-detection">
  <title>Download AtomicBot</title><!--[if (mso 16)]>
  <style type="text/css">
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;700&display=swap');
    a {text-decoration: none;}
    </style>
    <![endif]--><!--[if gte mso 9]><style>sup { font-size: 100% !important; }</style><![endif]--><!--[if gte mso 9]>
<noscript>
         <xml>
           <o:OfficeDocumentSettings>
           <o:AllowPNG></o:AllowPNG>
           <o:PixelsPerInch>96</o:PixelsPerInch>
           </o:OfficeDocumentSettings>
         </xml>
      </noscript>
<![endif]--><!--[if mso]><xml>
    <w:WordDocument xmlns:w="urn:schemas-microsoft-com:office:word">
      <w:DontUseAdvancedTypographyReadingMail/>
    </w:WordDocument>
    </xml><![endif]-->
  <style type="text/css">
    body {
      font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif !important;
      background-color: #F3F4F6 !important;
      color: #4B5563 !important;
    }
    h1, h2, h3, h4, h5, h6 {
      font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif !important;
      color: #111827 !important;
      letter-spacing: -0.025em !important;
    }
    .es-content-body {
      background-color: #FFFFFF !important;
      border-radius: 16px !important;
      box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1), 0 4px 6px -2px rgba(0, 0, 0, 0.05) !important;
    }
    .referral-box {
      background-color: #F9FAFB !important;
      border: 1px solid #E5E7EB !important;
      border-radius: 12px !important;
    }
    .referral-code {
      font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace !important;
      color: #4F46E5 !important;
      background-color: #EEF2FF !important;
      padding: 4px 8px;
      border-radius: 6px;
      -webkit-user-select: all;
      user-select: all;
      cursor: text;
    }
    #outlook a { padding: 0; }
    span.MsoHyperlink, span.MsoHyperlinkFollowed { color: inherit; mso-style-priority: 99; }
    a.es-button { mso-style-priority: 100 !important; text-decoration: none !important; }
    a[x-apple-data-detectors], #MessageViewBody a { color: inherit !important; text-decoration: none !important; font-size: inherit !important; font-family: inherit !important; font-weight: inherit !important; line-height: inherit !important; }
    @media only screen and (max-width:600px) {
      .es-m-p0r { padding-right: 0px !important }
      .es-m-p0l { padding-left: 0px !important }
      p, a { line-height: 150% !important }
      h1, h1 a { line-height: 120% !important }
      h2, h2 a { line-height: 120% !important }
      h3, h3 a { line-height: 120% !important }
      h4, h4 a { line-height: 120% !important }
      h5, h5 a { line-height: 120% !important }
      h6, h6 a { line-height: 120% !important }
      h1 { font-size: 30px !important; text-align: center }
      h2 { font-size: 24px !important; text-align: center }
      h3 { font-size: 20px !important; text-align: center }
      h4 { font-size: 24px !important; text-align: left }
      h5 { font-size: 20px !important; text-align: left }
      h6 { font-size: 16px !important; text-align: left }
      .es-header-body h1 a, .es-content-body h1 a, .es-footer-body h1 a { font-size: 30px !important }
      .es-header-body h2 a, .es-content-body h2 a, .es-footer-body h2 a { font-size: 24px !important }
      .es-header-body h3 a, .es-content-body h3 a, .es-footer-body h3 a { font-size: 20px !important }
      .es-header-body h4 a, .es-content-body h4 a, .es-footer-body h4 a { font-size: 24px !important }
      .es-header-body h5 a, .es-content-body h5 a, .es-footer-body h5 a { font-size: 20px !important }
      .es-header-body h6 a, .es-content-body h6 a, .es-footer-body h6 a { font-size: 16px !important }
      .es-menu td a { font-size: 12px !important }
      .es-header-body p, .es-header-body a { font-size: 14px !important }
      .es-content-body p, .es-content-body a { font-size: 16px !important }
      .es-footer-body p, .es-footer-body a { font-size: 14px !important }
      .es-infoblock p, .es-infoblock a { font-size: 12px !important }
      .es-m-txt-c, .es-m-txt-c h1, .es-m-txt-c h2, .es-m-txt-c h3, .es-m-txt-c h4, .es-m-txt-c h5, .es-m-txt-c h6 { text-align: center !important }
      .es-m-txt-r, .es-m-txt-r h1, .es-m-txt-r h2, .es-m-txt-r h3, .es-m-txt-r h4, .es-m-txt-r h5, .es-m-txt-r h6 { text-align: right !important }
      .es-m-txt-j, .es-m-txt-j h1, .es-m-txt-j h2, .es-m-txt-j h3, .es-m-txt-j h4, .es-m-txt-j h5, .es-m-txt-j h6 { text-align: justify !important }
      .es-m-txt-l, .es-m-txt-l h1, .es-m-txt-l h2, .es-m-txt-l h3, .es-m-txt-l h4, .es-m-txt-l h5, .es-m-txt-l h6 { text-align: left !important }
      .es-m-txt-r img, .es-m-txt-c img, .es-m-txt-l img { display: inline !important }
      .es-m-txt-r .rollover:hover .rollover-second, .es-m-txt-c .rollover:hover .rollover-second, .es-m-txt-l .rollover:hover .rollover-second { display: inline !important }
      .es-m-txt-r .rollover span, .es-m-txt-c .rollover span, .es-m-txt-l .rollover span { line-height: 0 !important; font-size: 0 !important; display: block }
      .es-spacer { display: inline-table }
      a.es-button, button.es-button { font-size: 20px !important; padding: 10px 20px 10px 20px !important; line-height: 120% !important }
      a.es-button, button.es-button, .es-button-border { display: inline-block !important }
      .es-m-fw, .es-m-fw.es-fw, .es-m-fw .es-button { display: block !important }
      .es-m-il, .es-m-il .es-button, .es-social, .es-social td, .es-menu.es-table-not-adapt { display: inline-block !important }
      .es-adaptive table, .es-left, .es-right { width: 100% !important }
      .es-content table, .es-header table, .es-footer table, .es-content, .es-footer, .es-header { width: 100% !important; max-width: 600px !important }
      .adapt-img { width: 100% !important; height: auto !important }
      .es-adapt-td { display: block !important; width: 100% !important }
      .es-mobile-hidden, .es-hidden { display: none !important }
      .es-container-hidden { display: none !important }
      .es-desk-hidden { width: auto !important; overflow: visible !important; float: none !important; max-height: inherit !important; line-height: inherit !important }
      tr.es-desk-hidden { display: table-row !important }
      table.es-desk-hidden { display: table !important }
      td.es-desk-hidden { display: table-cell !important }
      td.es-desk-menu-hidden { display: table-cell !important }
      .es-menu td { width: 1% !important }
      table.es-table-not-adapt, .esd-block-html table { width: auto !important }
      .h-auto { height: auto !important }
      .es-m-w-100 { width: -100% !important }
      .img-1447 { width: 50px !important }
      .es-text-2405 .es-text-mobile-size-24, .es-text-2405 .es-text-mobile-size-24 * { font-size: 24px !important }
      .es-text-5729 .es-text-mobile-size-20, .es-text-5729 .es-text-mobile-size-20 * { font-size: 20px !important }
    }
    @media screen and (max-width:384px) { .mail-message-content { width: 414px !important } }
  </style>
</head>

<body class="body"
  style="width:100%;height:100%;font-family:'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;-webkit-text-size-adjust:100%;-ms-text-size-adjust:100%;padding:0;Margin:0">
  <div
    style="display:none;font-size:1px;color:#333333;line-height:1px;max-height:0px;max-width:0px;opacity:0;overflow:hidden;">
    Welcome to Atomic Bot. Here is your access.
    &zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;&zwnj;&nbsp;
  </div>
  <div dir="ltr" class="es-wrapper-color" lang="en" style="background-color:#F3F4F6"><!--[if gte mso 9]>
			<v:background xmlns:v="urn:schemas-microsoft-com:vml" fill="t">
				<v:fill type="tile" color="#F3F4F6"></v:fill>
			</v:background>
		<![endif]-->
    <table width="100%" cellspacing="0" cellpadding="0" class="es-wrapper" role="none"
      style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px;padding:0;Margin:0;width:100%;height:100%;background-repeat:repeat;background-position:center top;background-color:#F3F4F6">
      <tr>
        <td valign="top" style="padding:0;Margin:0">
          <table cellpadding="0" cellspacing="0" align="center" class="es-content" role="none"
            style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px;width:100%;table-layout:fixed !important">
            <tr>
              <td align="center" style="padding:0;Margin:0;padding-top:20px;padding-bottom:20px">
                <table bgcolor="#ffffff" align="center" cellpadding="0" cellspacing="0" class="es-content-body"
                  role="none"
                  style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px;background-color:#FFFFFF;width:600px;border-radius:16px;box-shadow:0 10px 15px -3px rgba(0, 0, 0, 0.1), 0 4px 6px -2px rgba(0, 0, 0, 0.05);overflow:hidden">
                  <tr>
                    <td align="left"
                      style="Margin:0;padding-top:30px;padding-right:20px;padding-bottom:10px;padding-left:20px">
                      <table cellpadding="0" cellspacing="0" align="left" class="es-left" role="none"
                        style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px;float:left">
                        <tr>
                          <td align="center" valign="top" style="padding:0;Margin:0;width:560px">
                            <table cellpadding="0" cellspacing="0" width="100%" role="presentation"
                              style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px">
                              <tr>
                                <td align="center"
                                  style="padding:0;Margin:0;padding-bottom:24px;padding-top:24px;font-size:0px"><img
                                    src="https://fyiuhmm.stripocdn.email/content/guids/CABINET_5b4cbac9f587a1f0e716af8db283aa7ec20b0f69ac71ed604f4e7263766a26cd/images/128.png"
                                    alt="" width="80" class="img-1447"
                                    style="display:block;font-size:14px;border:0;outline:none;text-decoration:none;margin:0;opacity:0.85">
                                </td>
                              </tr>
                              <tr>
                                <td align="center" style="padding:0;Margin:0;padding-bottom:16px">
                                  <h1 class="es-m-txt-c"
                                    style="Margin:0;font-family:'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;mso-line-height-rule:exactly;letter-spacing:-0.025em;font-size:36px;font-style:normal;font-weight:800;line-height:42px;color:#111827">
                                    Welcome to Atomic Bot</h1>
                                </td>
                              </tr>
                              <!-- Cloud Version Block -->
                              <table width="100%" cellpadding="0" cellspacing="0" role="presentation">
                                <tr>
                                  <td align="center" style="padding:0;Margin:0;padding-bottom:10px;padding-top:20px">
                                    <h2
                                      style="Margin:0;font-family:'Inter', sans-serif;font-size:20px;font-weight:600;color:#111827">
                                      Web Version (Cloud)</h2>
                                  </td>
                                </tr>
                                <tr>
                                  <td align="center"
                                    style="padding:0;Margin:0;padding-bottom:20px;padding-left:20px;padding-right:20px">
                                    <p
                                      style="Margin:0;font-family:'Inter', sans-serif;font-size:15px;line-height:22px;color:#4B5563">
                                      Run your AI Assistant in one click. Launch in the cloud, no setup required.</p>
                                  </td>
                                </tr>
                                <tr>
                                  <td align="center" style="padding:0;Margin:0;padding-bottom:30px">
                                    <span class="es-button-border"
                                      style="border-style:solid;border-color:#111827;background:#111827;border-width:0px;display:inline-block;border-radius:100px;width:auto">
                                      <a href="https://atomicbot.ai?utm_source=email&utm_medium=welcome_email&utm_campaign=early_access"
                                        class="es-button" target="_blank"
                                        style="mso-style-priority:100 !important;text-decoration:none;font-family:'Inter', sans-serif;font-size:18px;color:#FFFFFF;border-style:solid;border-color:#111827;border-width:12px 24px;display:inline-block;background:#111827;border-radius:100px;font-weight:600">Start
                                        Now</a>
                                    </span>
                                  </td>
                                </tr>
                              </table>

                              <!-- Mac Version Block -->
                              <table width="100%" cellpadding="0" cellspacing="0" role="presentation">
                                <tr>
                                  <td align="center"
                                    style="padding:0;Margin:0;padding-bottom:10px;padding-top:10px;border-top:1px solid #F3F4F6">
                                    <h2
                                      style="Margin:0;font-family:'Inter', sans-serif;font-size:20px;font-weight:600;color:#111827;padding-top:20px">
                                      Available on MacOS</h2>
                                  </td>
                                </tr>
                                <tr>
                                  <td align="center"
                                    style="padding:0;Margin:0;padding-bottom:20px;padding-left:20px;padding-right:20px">
                                    <p
                                      style="Margin:0;font-family:'Inter', sans-serif;font-size:15px;line-height:22px;color:#4B5563">
                                      Open the installer and run locally on your Mac.</p>
                                  </td>
                                </tr>
                                <tr>
                                  <td align="center" style="padding:0;Margin:0;padding-bottom:30px">
                                    <span class="es-button-border"
                                      style="border-style:solid;border-color:#111827;background:#111827;border-width:0px;display:inline-block;border-radius:100px;width:auto">
                                      <a href="https://github.com/AtomicBot-ai/atomicbot/releases/download/v1.0.94/Atomic.Bot-1.0.94-arm64.dmg"
                                        class="es-button" target="_blank"
                                        style="mso-style-priority:100 !important;text-decoration:none;font-family:'Inter', sans-serif;font-size:18px;color:#FFFFFF;border-style:solid;border-color:#111827;border-width:12px 24px;display:inline-block;background:#111827;border-radius:100px;font-weight:600">Download
                                        for Mac</a>
                                    </span>
                                  </td>
                                </tr>
                              </table>

                              <!-- Windows Version Block -->
                              <table width="100%" cellpadding="0" cellspacing="0" role="presentation">
                                <tr>
                                  <td align="center"
                                    style="padding:0;Margin:0;padding-bottom:10px;padding-top:10px;border-top:1px solid #F3F4F6">
                                    <h2
                                      style="Margin:0;font-family:'Inter', sans-serif;font-size:20px;font-weight:600;color:#111827;padding-top:20px">
                                      Available on Windows</h2>
                                  </td>
                                </tr>
                                <tr>
                                  <td align="center"
                                    style="padding:0;Margin:0;padding-bottom:20px;padding-left:20px;padding-right:20px">
                                    <p
                                      style="Margin:0;font-family:'Inter', sans-serif;font-size:15px;line-height:22px;color:#4B5563">
                                      Open the installer and run locally on your Windows PC.</p>
                                  </td>
                                </tr>
                                <tr>
                                  <td align="center" style="padding:0;Margin:0;padding-bottom:30px">
                                    <span class="es-button-border"
                                      style="border-style:solid;border-color:#111827;background:#111827;border-width:0px;display:inline-block;border-radius:100px;width:auto">
                                      <a href="https://github.com/AtomicBot-ai/atomicbot/releases/download/v1.0.94/Atomic-Bot-1.0.94-x64-win.exe"
                                        class="es-button" target="_blank"
                                        style="mso-style-priority:100 !important;text-decoration:none;font-family:'Inter', sans-serif;font-size:18px;color:#FFFFFF;border-style:solid;border-color:#111827;border-width:12px 24px;display:inline-block;background:#111827;border-radius:100px;font-weight:600">Download
                                        for Windows</a>
                                    </span>
                                  </td>
                                </tr>
                              </table>

                              <!-- Coming Soon Block -->
                              <table width="100%" cellpadding="0" cellspacing="0" role="presentation">
                                <tr>
                                  <td align="center"
                                    style="padding:0;Margin:0;padding-bottom:10px;padding-top:10px;border-top:1px solid #F3F4F6">
                                    <h2
                                      style="Margin:0;font-family:'Inter', sans-serif;font-size:20px;font-weight:600;color:#111827;padding-top:20px">
                                      Coming Soon</h2>
                                  </td>
                                </tr>
                                <tr>
                                  <td align="center"
                                    style="padding:0;Margin:0;padding-bottom:30px;padding-left:20px;padding-right:20px">
                                    <p
                                      style="Margin:0;font-family:'Inter', sans-serif;font-size:15px;line-height:22px;color:#4B5563">
                                      One-click installers for iOS, Android and Linux coming soon.</p>
                                  </td>
                                </tr>
                              </table>

                              <!-- Social Icons -->
                              <table width="100%" cellpadding="0" cellspacing="0" role="presentation">
                                <tr>
                                  <td align="center"
                                    style="padding:0;Margin:0;padding-bottom:20px;padding-top:10px;border-top:1px solid #F3F4F6">
                                    <table cellpadding="0" cellspacing="0" class="es-table-not-adapt es-social"
                                      role="presentation"
                                      style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px;display:inline-block">
                                      <tr>
                                        <td valign="top" align="center"
                                          style="padding:0;Margin:0;padding-right:20px;padding-top:20px"><a
                                            target="_blank" href="https://x.com/atomicbot_ai"
                                            style="text-decoration:none"><img alt="X" title="X"
                                              src="https://cdn.prod.website-files.com/6981cca3fe2c3f562a2ad751/698af2542595cb0f85cbc958_x%20(2).png"
                                              height="24" style="display:block;border:0;height:24px;width:auto"></a>
                                        </td>
                                        <td align="center" valign="top" style="padding:0;Margin:0;padding-top:20px"><a
                                            href="https://discord.gg/2TXafRV69m" target="_blank"
                                            style="text-decoration:none"><img title="Discord"
                                              src="https://cdn.prod.website-files.com/6981cca3fe2c3f562a2ad751/698af2544c5acd7d65ba8f22_discord%20(2).png"
                                              height="24" alt="Discord"
                                              style="display:block;border:0;height:24px;width:auto"></a></td>
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
    </table>

    <!-- Footer with Terms and Unsubscribe -->
    <table cellpadding="0" cellspacing="0" align="center" role="none" style="width:100%;max-width:600px;margin:0 auto;">
      <tr>
        <td align="center" style="padding:20px;">
          <a target="_blank" href="https://atomicbot.ai/terms-of-service"
            style="text-decoration:none;font-family:'Inter', sans-serif;color:#9CA3AF;font-size:12px">Terms of
            Service</a>
          <span style="color:#9CA3AF;font-size:12px;font-family:'Inter', sans-serif"> &bull; </span>
          <a target="_blank" href="${unsubscribeUrl}"
            style="text-decoration:none;font-family:'Inter', sans-serif;color:#9CA3AF;font-size:12px">Unsubscribe</a>
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
        subject: "\u2705 +1 Atomic Bot! Here's how to get started",
        html: buildEmailHtml(unsubscribeUrl),
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
