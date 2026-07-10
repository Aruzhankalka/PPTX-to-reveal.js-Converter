/**
 * In-memory conversion-result store — download.js/upload.js's shared cache
 * of {ir, html, media, warnings} results keyed by resultId, so a slow
 * conversion is done once and served many times (preview, ZIP download,
 * individual media). Not persisted: a process restart drops everything.
 */

const results = new Map();

const TTL_MS = 30 * 60 * 1000; // 30 minutes

/**
 * Store a result under resultId and schedule its automatic deletion after
 * TTL_MS (30 minutes). Overwrites any existing entry for the same id
 * silently (last write wins) — but does NOT cancel that entry's own
 * expiry timer, so calling this twice for the same resultId schedules two
 * independent deletions; whichever fires first deletes whatever is
 * currently stored, which can be earlier than 30 minutes after the second
 * save.
 *
 * @param {string} resultId - unique id for this conversion result
 * @param {object} data - result payload (e.g. {ir, html, media, warnings})
 * @returns {void}
 */
function saveResult(resultId, data) {
  results.set(resultId, {
    ...data,
    createdAt: Date.now()
  });

  setTimeout(() => {
    results.delete(resultId);
  }, TTL_MS);
}

/**
 * Look up a stored result by id.
 * @param {string} resultId
 * @returns {object|undefined} the stored {..., createdAt} object, or
 *   undefined if resultId was never saved or has since expired (TTL_MS).
 */
function getResult(resultId) {
  return results.get(resultId);
}

module.exports = {
  saveResult,
  getResult
};
