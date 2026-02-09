# referral_bot

Express proxy between Webflow forms and KickoffLabs API. Designed to run on Railway.

## API

**POST /signup** — accepts `{ email, kid }`, proxies the request to KickoffLabs and returns the referral code, share URL, and referral count.

## Environment Variables

| Variable | Description |
|---|---|
| `KOL_CAMPAIGN_ID` | KickoffLabs campaign ID |
| `KOL_API_KEY` | KickoffLabs API key |
| `PORT` | Server port (default: 3000) |
