// Process entry point (npm start/dev). Kept separate from app.js so tests
// can require the Express app directly with supertest, without binding a
// real port.
const app = require('./app');

const PORT = 3000;

app.listen(PORT, () => {
  console.log(`Converter running on http://localhost:${PORT}`);
});
