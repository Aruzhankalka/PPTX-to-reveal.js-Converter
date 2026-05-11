const app = require('./app');

const PORT = 3000;

app.listen(PORT, () => {
  console.log(`Converter running on http://localhost:${PORT}`);
});
