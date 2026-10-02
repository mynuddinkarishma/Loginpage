# Loginpage

## Deploy on Railway

1. Create a Railway project from this GitHub repository and deploy the `main` branch. Railway runs `npm run build` and starts the app with `npm start`.
2. Add a persistent volume to the service and mount it at `/app/data`. The server stores its SQLite database on that volume.
3. Generate a public domain for the service. The same service serves the React app and its API; `/health` is the health-check endpoint.

The local development commands are `npm run dev` for the frontend and `npm run server` for the API.
