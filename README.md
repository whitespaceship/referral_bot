# referral_bot

Express proxy between Webflow forms and Resend. Handles signup and welcome email delivery. Designed to run on Railway.

## API

**POST /signup** — accepts `{ email }`, adds the contact to the Resend audience, and sends the welcome-email template.

## Environment Variables

| Variable | Description |
|---|---|
| `RESEND_API_KEY` | Resend API key (required) |
| `RESEND_AUDIENCE_ID` | Resend audience to add signups to (optional) |
| `PORT` | Server port (default: 3000) |
