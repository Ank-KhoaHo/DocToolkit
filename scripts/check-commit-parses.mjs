#!/usr/bin/env node
//
// Fail when the commit message a squash merge WILL create cannot be parsed by release-please.
//
// WHY THIS EXISTS, and why the two regex checks beside it were not enough.
//
// This repository squash-merges with `squash_merge_commit_title: PR_TITLE` and
// `squash_merge_commit_message: PR_BODY`, so a pull request becomes one commit whose subject is
// the PR title and whose BODY IS THE PR DESCRIPTION - arbitrary markdown, written for humans,
// handed to a machine parser.
//
// On 2026-08-20 that cost a shipped feature. #312 merged as
// `feat(extensions): configure PDF fonts once, on DocToolkitOptions`, a flawless Conventional
// Commit subject that both sibling checks passed. Its body contained a ```csharp fence opening
// with `services.AddDocToolkit(o =>`, and release-please reported:
//
//     commit could not be parsed: 0b546a2 feat(extensions): configure PDF fonts once ... (#312)
//     error message: Error: unexpected token '\n' at 8:28, valid tokens [)]
//     No user facing commits found since 65bb53b - skipping
//
// An unclosed `(` in the body made the parser abandon the ENTIRE commit, subject included. That
// `feat:` was the only bumping commit since 0.32.0, so no Release PR was opened at all and the
// feature sat on main proposing nothing.
//
// THE FAILURE IS AN ABSENCE, WHICH IS WHY A HUMAN CANNOT BE THE CHECK. A spurious changelog entry
// gets noticed while curating a Release PR. A missing one does not, because there is no Release PR
// to read. Worse, CLAUDE.md tells readers that `commit could not be parsed` is routine noise from
// merge commits and "not worth fixing" - true for 25 runs, and the sentence standing between a
// reader and this.
//
// IT USES RELEASE-PLEASE'S OWN PARSER RATHER THAN A SECOND IMPLEMENTATION OF THE GRAMMAR.
// `@conventional-commits/parser` is what release-please 17.6.0 depends on, and reproducing
// #312 locally with it returned a byte-identical error, position included. A hand-written
// approximation here would be a second thing to keep in step with a dependency - the same
// mistake `gen-third-party-notices.py` avoids by reading the lockfile.
//
// Usage:
//   node scripts/check-commit-parses.mjs            message on stdin
//   node scripts/check-commit-parses.mjs --self-test controls; see the bottom of this file

import { readFileSync } from 'node:fs';

/** Load the parser, or fail loudly. A guard that passes because its parser is missing is worse
 *  than no guard - the same standard check-readme-coverage.py holds when it parses zero types. */
async function loadParser() {
  try {
    const mod = await import('@conventional-commits/parser');
    if (typeof mod.parser !== 'function') throw new Error('no parser export');
    return mod.parser;
  } catch (err) {
    console.error('::error::Could not load @conventional-commits/parser: ' + err.message);
    console.error('::error::Install it first: npm install --no-save @conventional-commits/parser@0.4.1');
    console.error('::error::Refusing to report success - this check cannot run without it.');
    process.exit(2);
  }
}

/** Point at the offending line. The parser says "8:28", which alone sends a reader counting
 *  newlines in a message they cannot see. */
function locate(message, error) {
  const at = /at (\d+):(\d+)/.exec(error.message ?? '');
  if (!at) return null;
  const [line, column] = [Number(at[1]), Number(at[2])];
  const text = message.split('\n')[line - 1];
  if (text === undefined) return null;
  return { line, column, text };
}

function check(message, parser) {
  try {
    parser(message);
    return { ok: true };
  } catch (err) {
    return { ok: false, error: err, where: locate(message, err) };
  }
}

/**
 * GITHUB RE-WRAPS THE PULL REQUEST BODY WHEN IT BUILDS THE SQUASH COMMIT MESSAGE, so the text
 * this check is handed is NOT the text git receives. Checking only the authored form is a hole,
 * and it cost a second dropped release the day after this script shipped.
 *
 * #315 was a `fix:` whose body contained an inline code span reading DoesNotContain("control
 * character"). Authored, it parsed. Committed, GitHub had folded the line so the span ended a
 * line at the open bracket, and release-please reported `commit could not be parsed` again.
 * Proved by isolating the two differences between body and commit: adding the `(#315)` suffix
 * to the subject still parses, re-wrapping the body alone reproduces the failure exactly.
 *
 * The observed commit wrapped to 89 columns. THE WIDTH IS UNDOCUMENTED AND NOT RELIED ON - a
 * body that only survives one particular fold is a body waiting to break, so several widths are
 * tried and any failure fails the check.
 */
function wrapLine(line, width) {
  if (line.length <= width) return [line];
  const out = [];
  let current = '';
  for (const word of line.split(' ')) {
    if (current === '') current = word;
    else if ((current + ' ' + word).length <= width) current += ' ' + word;
    else { out.push(current); current = word; }
  }
  if (current !== '') out.push(current);
  return out;
}

function wrapBody(message, width) {
  const split = message.indexOf('\n\n');
  if (split < 0) return message;
  const subject = message.slice(0, split);
  const body = message.slice(split + 2);
  const folded = body.split('\n').flatMap(l => wrapLine(l, width)).join('\n');
  return subject + '\n\n' + folded;
}

/** Widths to fold at. 89 is what GitHub was measured doing; the rest bracket it. */
const WIDTHS = [72, 80, 89, 100];

const parser = await loadParser();

if (process.argv.includes('--self-test')) {
  // The negative control is #312's own shape - the case that used to pass every check here.
  const bad = [
    'feat(extensions): configure PDF fonts once, on DocToolkitOptions (#312)',
    '',
    'Closes A56 - the mirror for PdfFontOptions.',
    '',
    '```csharp',
    'services.AddDocToolkit(o =>',
    '    o.Fonts = new PdfFontOptions("Noto Sans", File.ReadAllBytes("f.ttf")));',
    '```',
  ].join('\n');

  // The positive control. Without it, a parser that rejected EVERYTHING would look like a
  // working guard - it would fail the bad message for the wrong reason and nothing would say so.
  const good = [
    'feat(core): clamp a negative paragraph indent so the document renders',
    '',
    'Word allows a paragraph outside the margin; the PDF renderer refuses it.',
    'Measured over 99 documents: 71/99 to 75/99.',
  ].join('\n');

  // The fold-only control: #315's own body, reduced. It PARSES as authored and BREAKS once
  // folded, which is precisely the case the first version of this script could not see. Without
  // it, reverting the folding logic passes the whole self-test.
  const foldOnly = [
    'fix(core): name the control character instead of leaving it to the inner exception',
    '',
    'diagnoses also end with that phrase, so it could no longer fail. It now asserts the generic',
    'wrapper own wording, `"See the inner exception for details"`, plus `DoesNotContain("control',
    'character")`.',
  ].join('\n');

  const b = check(bad, parser);
  const g = check(good, parser);
  const fAuthored = check(foldOnly, parser);
  const fFolded = WIDTHS.map(w => check(wrapBody(foldOnly, w), parser));
  let failed = false;

  if (!fAuthored.ok) {
    console.error('SELF-TEST FAILED: the fold-only control was expected to parse as authored.');
    failed = true;
  } else if (fFolded.every(r => r.ok)) {
    console.error('SELF-TEST FAILED: the fold-only control still parses at every width, so this');
    console.error('check cannot see the failure that dropped #315. The folding logic is not working.');
    failed = true;
  } else {
    const w = WIDTHS[fFolded.findIndex(r => !r.ok)];
    console.log(`ok   fold-only control: parses as authored, breaks folded at ${w} columns`);
  }

  if (b.ok) {
    console.error('SELF-TEST FAILED: the #312 message parsed, so this check would not have caught it.');
    failed = true;
  } else {
    console.log(`ok   negative control rejected: ${b.error.message}`);
    if (b.where) console.log(`     line ${b.where.line}: ${b.where.text}`);
  }

  if (!g.ok) {
    console.error(`SELF-TEST FAILED: an ordinary message was rejected: ${g.error.message}`);
    console.error('This check would fail every pull request. Do not ship it.');
    failed = true;
  } else {
    console.log('ok   positive control accepted');
  }

  process.exit(failed ? 1 : 0);
}

const message = readFileSync(0, 'utf8').replace(/\r\n/g, '\n').trimEnd();

if (!message.trim()) {
  console.error('::error::No commit message on stdin.');
  process.exit(2);
}

// As authored, then as GitHub will fold it. Any failure is a failure.
let result = check(message, parser);
let foldedAt = null;

if (result.ok) {
  for (const width of WIDTHS) {
    const attempt = check(wrapBody(message, width), parser);
    if (!attempt.ok) { result = attempt; foldedAt = width; break; }
  }
}

if (result.ok) {
  console.log(`the squashed commit message parses as release-please reads it, `
    + `as authored and folded at ${WIDTHS.join(', ')} columns`);
  process.exit(0);
}

if (foldedAt !== null) {
  console.error(`::error::This parses as you wrote it and BREAKS once GitHub folds it to ${foldedAt} columns.`);
  console.error('::error::GitHub re-wraps the pull request body when it builds the squash commit');
  console.error('::error::message, so the text release-please sees is not the text you typed. This is');
  console.error('::error::not hypothetical - it dropped #315 from its own release the day after this');
  console.error('::error::check shipped, and the check passed because it only looked at the authored form.');
  console.error('::error::');
}

console.error(`::error::release-please cannot parse the commit message this pull request will create: ${result.error.message}`);
if (result.where) {
  console.error(`::error::line ${result.where.line}, column ${result.where.column}: ${result.where.text}`);
  console.error('::error::' + ' '.repeat(Math.max(0, result.where.column - 1)) + '^');
}
console.error('::error::');
console.error('::error::The PR TITLE is fine. The PR DESCRIPTION becomes this commit\'s body, and');
console.error('::error::something in it broke the parser - most often an unclosed ( or [ inside a');
console.error('::error::code fence. release-please would discard the WHOLE commit, so every feat:');
console.error('::error::or fix: in this pull request would be missing from the changelog, and if it');
console.error('::error::is the only bumping commit, no Release PR would be opened at all.');
console.error('::error::');
console.error('::error::Fix the pull request description, then push or re-run this job.');
process.exit(1);
