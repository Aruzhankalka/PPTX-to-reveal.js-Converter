const results = new Map();

const TTL_MS = 30 * 60 * 1000; // 30 minutes

function saveResult(resultId, data) {
  results.set(resultId, {
    ...data,
    createdAt: Date.now()
  });

  setTimeout(() => {
    results.delete(resultId);
  }, TTL_MS);
}

function getResult(resultId) {
  return results.get(resultId);
}

module.exports = {
  saveResult,
  getResult
};
