'use strict';

/**
 * FR-14 Animation parser — best-effort PPTX timing tree walker.
 *
 * Walks the p:timing / p:tnLst tree and emits IR Animation objects.
 * Effects that can be mapped produce supported:true entries.
 * Effects that cannot be mapped produce supported:false entries + a warning.
 * When the timing tree is absent the function returns an empty array
 * (schema-valid, not an error).
 *
 * The PPTX animation model:
 *   p:timing
 *     p:tnLst
 *       p:par (root)
 *         p:cTn[nodeType=tmRoot]
 *           p:childTnLst
 *             p:seq[mainSeq]     ← the click-sequence container
 *               p:cTn
 *                 p:childTnLst
 *                   p:par        ← one "click group" per click
 *                     p:cTn[nodeType=click]
 *                       p:childTnLst
 *                         p:par  ← one effect node per animation
 *                           p:cTn[presetID, presetClass, nodeType=clickEffect|withEffect|afterEffect]
 *                             p:childTnLst → actual tween nodes (not parsed here)
 *                             p:stCondLst  → may carry delay
 *
 * The target shape id comes from the first descendant <p:spTgt spid="N"/>.
 * It is a PPTX integer sp id (nvSpPr cNvPr id), which we emit as a string
 * to match the IR convention.  slide.js integration will remap these to the
 * stable 'shp-N' ids; for now we emit what we read.
 */

const { asArray } = require('./xml');

// ---------------------------------------------------------------------------
// Preset class mapping
// ---------------------------------------------------------------------------

// OOXML presetClass → IR effect.class
const PRESET_CLASS_MAP = {
  entr:       'entrance',
  emph:       'emphasis',
  exit:       'exit',
  motionPath: 'motionPath',
};

// ---------------------------------------------------------------------------
// Preset ID mapping (ECMA-376 Annex P.1, partial)
// Generator groups map these strings to CSS animation names.
// ---------------------------------------------------------------------------

// Only the most common presets are mapped here; everything else falls through
// to 'unknown-<id>' which produces supported:false.
const PRESET_ID_MAP = {
  // Entrance
  1:  'appear',
  2:  'flyIn',
  5:  'blinds',
  10: 'wipe',
  11: 'box',
  14: 'checkerboard',
  21: 'fade',
  22: 'dissolve',
  26: 'split',
  27: 'strips',
  // Emphasis
  5:  'spin',   // shared id in different presetClass contexts; resolved per-class
  // Exit (same ids as entrance in PPTX spec — resolved per presetClass)
};

// ---------------------------------------------------------------------------
// Trigger mapping (nodeType → IR trigger)
// ---------------------------------------------------------------------------

const NODE_TYPE_TO_TRIGGER = {
  clickEffect:  'onClick',
  withEffect:   'withPrevious',
  afterEffect:  'afterPrevious',
  // These appear on the click-group par, not on individual effect pars:
  click:        'onClick',
};

// ---------------------------------------------------------------------------
// Timing extraction
// ---------------------------------------------------------------------------

/**
 * Read delay from the first <p:stCondLst><p:cond delay="N"/> descendant.
 * Returns 0 when absent.
 */
function extractDelay(cTn) {
  const stCondLst = cTn['p:stCondLst'];
  if (!stCondLst) return 0;
  for (const cond of asArray(stCondLst['p:cond'])) {
    const delay = cond['@_delay'];
    if (delay != null && delay !== 'indefinite') {
      const ms = Number(delay);
      if (!Number.isNaN(ms)) return ms;
    }
  }
  return 0;
}

/**
 * Read duration from the cTn @dur attribute (PPTX duration in ms, or
 * 'indefinite').  Returns 0 when absent or indefinite.
 */
function extractDuration(cTn) {
  const dur = cTn['@_dur'];
  if (!dur || dur === 'indefinite') return 0;
  const ms = Number(dur);
  return Number.isNaN(ms) ? 0 : ms;
}

// ---------------------------------------------------------------------------
// Target extraction
// ---------------------------------------------------------------------------

/**
 * Find the first <p:spTgt spid="N"/> in the subtree of a cTn node and return
 * its spid as a string, or null when absent.
 *
 * The depth of nesting varies by effect type so we do a breadth-first search
 * over the childTnLst → behaviour chain.
 */
function findTargetSpid(cTn) {
  // Direct tgtEl paths: p:childTnLst → <effect> → p:cBhvr / p:cMediaNode → p:tgtEl → p:spTgt
  const childList = cTn['p:childTnLst'];
  if (!childList) return null;

  // The first-level children are the actual tween nodes (p:animEffect, p:set, p:anim, etc.)
  const tweenTagNames = [
    'p:animEffect', 'p:set', 'p:anim', 'p:animClr', 'p:animMotion',
    'p:animRot', 'p:animScale', 'p:cmd', 'p:audio', 'p:video',
  ];

  for (const tweenTag of tweenTagNames) {
    for (const tween of asArray(childList[tweenTag])) {
      const tgtEl = tween['p:cBhvr']
        ? tween['p:cBhvr']['p:tgtEl']
        : (tween['p:cMediaNode'] ? tween['p:cMediaNode']['p:tgtEl'] : null);
      if (!tgtEl) continue;
      const spTgt = tgtEl['p:spTgt'];
      if (spTgt && spTgt['@_spid'] != null) {
        return String(spTgt['@_spid']);
      }
    }
  }

  // Fallback: walk nested par children for the first spTgt
  for (const par of asArray(childList['p:par'])) {
    const innerCTn = par['p:cTn'];
    if (!innerCTn) continue;
    const spid = findTargetSpid(innerCTn);
    if (spid) return spid;
  }

  return null;
}

// ---------------------------------------------------------------------------
// Effect par walker
// ---------------------------------------------------------------------------

/**
 * Parse a single effect par node (the p:par that directly wraps a p:cTn with
 * presetClass/presetID) into an IR Animation entry.
 *
 * @param {object}   cTn      - parsed p:cTn node inside the effect par
 * @param {string}   trigger  - IR trigger value resolved from the click group
 * @param {number}   order    - sequence index (0-based, mutable counter)
 * @param {number}   animIdx  - unique animation counter for id generation
 * @param {string[]} warnings - mutable warnings array
 * @returns {object} IR Animation object (always; supported may be false)
 */
function parseEffectCTn(cTn, trigger, order, animIdx, warnings) {
  const presetClass = cTn['@_presetClass'];
  const presetID    = cTn['@_presetID'];

  const effectClass = presetClass ? (PRESET_CLASS_MAP[presetClass] || null) : null;
  const effectPreset = presetID != null ? (PRESET_ID_MAP[Number(presetID)] || null) : null;

  const targetSpid = findTargetSpid(cTn);
  const targetId   = targetSpid ? `spid-${targetSpid}` : 'unknown';

  const delayMs    = extractDelay(cTn);
  const durationMs = extractDuration(cTn);

  const supported = !!(effectClass && effectPreset && targetSpid);

  if (!supported) {
    const why = !targetSpid
      ? 'no target element found'
      : !effectClass
        ? `unknown presetClass "${presetClass}"`
        : `unknown presetID ${presetID} for class "${presetClass}"`;
    warnings.push(`animation anim-${animIdx}: ${why} — skipping in generator`);
  }

  return {
    id:        `anim-${animIdx}`,
    targetId,
    trigger,
    sequence: order,
    effect: {
      class:  effectClass  || 'entrance',
      preset: effectPreset || `unknown-${presetID ?? 'n/a'}`,
    },
    timing: {
      delayMs,
      durationMs,
    },
    supported,
  };
}

// ---------------------------------------------------------------------------
// Click group walker
// ---------------------------------------------------------------------------

/**
 * Walk one "click group" par (nodeType=click) and collect effect animations.
 *
 * Each direct child par of the click group has a p:cTn with nodeType
 * clickEffect / withEffect / afterEffect.  The trigger for the first child is
 * 'onClick'; subsequent children with withEffect / afterEffect get their own
 * trigger values.
 */
function walkClickGroup(clickCTn, orderCounter, animIdx, warnings) {
  const results = [];
  const childTnLst = clickCTn['p:childTnLst'];
  if (!childTnLst) return results;

  for (const par of asArray(childTnLst['p:par'])) {
    const cTn = par['p:cTn'];
    if (!cTn) continue;

    const nodeType = cTn['@_nodeType'];
    const trigger  = NODE_TYPE_TO_TRIGGER[nodeType] || 'onClick';

    const anim = parseEffectCTn(cTn, trigger, orderCounter.value, animIdx.value, warnings);
    results.push(anim);
    orderCounter.value++;
    animIdx.value++;
  }

  return results;
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Parse the animation timing tree of a slide into IR Animation objects.
 *
 * @param {object}   sld      - parsed <p:sld> root node (the outermost PPTX
 *                              slide object, NOT the spTree).
 * @param {string[]} [warnings] - mutable warnings array (created internally
 *                                when not provided so callers can omit it).
 * @returns {{ animations: object[], warnings: string[] }}
 *   animations — array of IR Animation objects (may be empty)
 *   warnings   — human-readable strings describing unmapped effects
 */
function parseAnimations(sld, warnings) {
  const w = warnings || [];

  const timing = sld && sld['p:timing'];
  if (!timing) return { animations: [], warnings: w };

  // Walk: p:timing → p:tnLst → p:par → p:cTn[tmRoot] → p:childTnLst →
  //        p:seq[mainSeq] → p:cTn → p:childTnLst → p:par[click groups]
  const tnLst = timing['p:tnLst'];
  if (!tnLst) return { animations: [], warnings: w };

  const rootPar  = tnLst['p:par'];
  const rootCTn  = rootPar && rootPar['p:cTn'];
  if (!rootCTn) return { animations: [], warnings: w };

  const rootChild = rootCTn['p:childTnLst'];
  if (!rootChild) return { animations: [], warnings: w };

  // The main sequence is the p:seq with nodeType=mainSeq
  const seqNodes = asArray(rootChild['p:seq']);
  let mainSeqCTn = null;
  for (const seq of seqNodes) {
    const cTn = seq['p:cTn'];
    if (cTn && cTn['@_nodeType'] === 'mainSeq') {
      mainSeqCTn = cTn;
      break;
    }
  }
  // Fallback: first seq if no mainSeq tag
  if (!mainSeqCTn && seqNodes.length > 0) {
    mainSeqCTn = seqNodes[0]['p:cTn'];
  }
  if (!mainSeqCTn) return { animations: [], warnings: w };

  const seqChildList = mainSeqCTn['p:childTnLst'];
  if (!seqChildList) return { animations: [], warnings: w };

  const animations = [];
  const orderCounter = { value: 0 };
  const animIdx     = { value: 0 };

  for (const clickPar of asArray(seqChildList['p:par'])) {
    const clickCTn = clickPar['p:cTn'];
    if (!clickCTn) continue;

    const group = walkClickGroup(clickCTn, orderCounter, animIdx, w);
    animations.push(...group);
  }

  return { animations, warnings: w };
}

module.exports = { parseAnimations };
