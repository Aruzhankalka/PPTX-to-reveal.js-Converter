module.exports = {
  testDir: './tests/e2e',
  testMatch: '**/*.spec.js',
  timeout: 30000,
  use: {
    baseURL: 'http://localhost:3000',
  },
  projects: [
    { name: 'Chrome', use: { browserName: 'chromium' } },
    { name: 'Firefox', use: { browserName: 'firefox' } },
    { name: 'Edge', use: { browserName: 'chromium', channel: 'msedge' } },
  ],
};