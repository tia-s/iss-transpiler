# iss-transpiler
A transpiler to convert IDEAScript to Pandas. Additional target frameworks/languages can be configured by implementing a new Translator.
```
isstranspiler/
├── client/
│   ├── public/
│   │   └── favicon.ico
│   ├── src/
│   │   ├── services/
│   │   │   └── transpiler.ts     # axios calls to server API
│   │   ├── components/           # UI elements
│   │   │   ├── EditorPane.vue
│   │   │   ├── ConvertButton.vue
│   │   │   └── ArrowDivider.vue
│   │   ├── composables/
│   │   │   └── useTranspiler.ts
│   │   ├── types/
│   │   │   └── index.ts          # shared TS interfaces
│   │   ├── App.vue
│   │   ├── main.ts
│   │   └── styles/
│   │       └── main.scss
│   ├── vite.config.ts
│   └── tsconfig.json
└── server/
    └── api/
```

## Local Development

### Client
```bash
cd client
npm install
cp .env.example .env
npm run dev
```

### Server
```bash
cd server
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
uvicorn main:app --reload
```

The FastAPI docs are available at `http://localhost:8000/docs` once the server is running.

## Environment Variables (.env.examples)

| Variable | Description |
|---|---|
| `VITE_API_URL` | Base URL of the FastAPI server e.g. `http://localhost:8000` |
