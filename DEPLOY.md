# Deploy Slide Discreet Craft

## Render

1. Put this `presentation-maker` folder in a GitHub repository.
2. Go to Render and create a new Web Service.
3. Connect the repository.
4. Use these settings:
   - Build command: `npm install`
   - Start command: `npm start`
   - Node version: `20` or newer
5. Deploy.

The app reads `PORT` automatically, so Render/Railway can choose the port.

## Railway

1. Create a new Railway project from GitHub.
2. Select this app folder.
3. Railway should detect Node automatically.
4. If it asks, set:
   - Install command: `npm install`
   - Start command: `npm start`

## Local Run

```bash
npm install
npm start
```

Open `http://localhost:4173`.
