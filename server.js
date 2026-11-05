import express from 'express';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = process.env.PORT || 3000;
const HOST = '0.0.0.0';

// Serve static files from root directory
app.use(express.static(__dirname, {
  extensions: ['html'],
  index: 'index.html'
}));

// Route fallback for booking modal
app.get('/booking-modal', (req, res) => {
  res.sendFile(path.join(__dirname, 'booking-modal.html'));
});

// Single-page / static fallback
app.get('*', (req, res) => {
  if (path.extname(req.path)) {
    return res.status(404).send('Not Found');
  }
  res.sendFile(path.join(__dirname, 'index.html'));
});

app.listen(PORT, HOST, () => {
  console.log(`AutoGlancis website server listening on http://${HOST}:${PORT}`);
});
