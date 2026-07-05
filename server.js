const express = require('express');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;
const ROOT_DIR = __dirname;

app.use(express.static(ROOT_DIR));

app.get('/health', (_req, res) => {
  res.json({ status: 'ok', app: 'account-contact-web', timestamp: new Date().toISOString() });
});

app.get('*', (_req, res) => {
  res.sendFile(path.join(ROOT_DIR, 'index.html'));
});

app.listen(PORT, () => {
  console.log(`Account & Contact Quality Check is running on http://localhost:${PORT}`);
});
