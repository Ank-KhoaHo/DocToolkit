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

  const b = check(bad, parser);
  const g = check(good, parser);
  let failed = false;

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

const result = check(message, parser);

if (result.ok) {
  console.log('the squashed commit message parses as release-please reads it');
  process.exit(0);
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
