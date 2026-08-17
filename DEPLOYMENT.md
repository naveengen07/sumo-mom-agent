# Deployment

This project is split into two services because meeting recordings are too
large to upload through a Vercel Function.

- **Vercel** serves `public/` only.
- **Render** (or another long-running Docker host) runs the Flask API in
  `app.py`.

## 1. Deploy the API to Render

1. Push this repository to GitHub.
2. In Render, choose **New +** > **Blueprint** and select the repository.
   Render reads `render.yaml` and builds the included `Dockerfile`. FFmpeg is
   included, which is required to extract and compress recording audio.
3. Add these environment variables in the Render service:

   ```text
   FLASK_SECRET_KEY=<long random secret>
   GOOGLE_CLIENT_ID=<Google OAuth client ID>
   GOOGLE_CLIENT_SECRET=<Google OAuth client secret>
   GOOGLE_REDIRECT_URI=https://<your-render-service>.onrender.com/api/google/callback
   FRONTEND_ORIGINS=https://sumo-mom-agent-gamma.vercel.app
   GOOGLE_TOKEN_JSON=<the complete JSON from local google_drive_token.json>
   ```

4. In Google Cloud Console, add the same Render callback URL as an authorized
   redirect URI for the OAuth client.
5. Confirm `https://<your-render-service>.onrender.com/api/health` returns
   JSON with `"ok": true`.

`GOOGLE_TOKEN_JSON` is a backend secret. It lets the API refresh its Google
Drive access token after container restarts; never put it in Vercel or commit
it to the repository.

## 2. Point Vercel at the API

Edit `public/config.js` and replace the empty value with the public Render URL:

```js
window.MOM_API_BASE_URL = "https://<your-render-service>.onrender.com";
```

Commit and push that change, then redeploy the Vercel project. The API URL is
public and is safe to include in this file; it is not a secret.

## 3. Local development

Copy `.env.example` to `.env`, fill in the local Google OAuth values, and run:

```powershell
python "backend geo.py"
```

The original command remains supported. It now launches the deployable Flask
application in `app.py`.

## Important upload limit

Do not deploy the recording-processing API as a Vercel Function. Vercel
Functions reject request bodies over 4.5 MB, while meeting recordings commonly
exceed that limit. The browser sends the recording directly to the Render API
instead.
