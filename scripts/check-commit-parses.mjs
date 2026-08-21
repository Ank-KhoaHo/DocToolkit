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
 * THE RULE, MEASURED RATHER THAN MODELLED, on 2026-08-22 - after this script's own comments
 * were found describing a guess as the mechanism. The shapes it was measured on are in the
 * agreement control at the bottom of this file, which re-checks every one of them against the
 * parser on each run. No count is written here: the control prints the authoritative one.
 *
 * A body line breaks the parse if, and only if, ALL of these hold:
 *
 *   1. the line begins with a NON-WHITESPACE character - any indentation immunises it, tab or
 *      spaces alike;
 *   2. it contains a `(` that is not the first character and has no whitespace before it;
 *   3. no `)` appears earlier in the line than that `(`;
 *   4. no `)` appears after it ON THE SAME LINE.
 *
 * The parser is lexing `<type>(<scope>)`, so what breaks a commit is a line that LOOKS LIKE A
 * CONVENTIONAL COMMIT HEADER with an unterminated scope. That explains every observation and
 * several that contradict the obvious guesses:
 *
 *   `foo(bar`                     FAILS    the shape, unterminated
 *   `foo (bar`                    parses   a space before `(` is not a scope
 *   `foo(a b c d e) trailing`     parses   closed on the same line; spaces inside are fine
 *   `a note foo(bar here`         parses   not at column 1
 *   `    foo(bar`                 parses   indented
 *   `foo[bar`  `foo{bar`          parses   ONLY `(` matters
 *   `(unclosed at the start`      parses   nothing precedes the bracket, so no type token
 *   `([#331](https://…/331))`     parses   release-please's own suffix is safe
 *
 * THREE THINGS THIS FILE USED TO CLAIM ARE WRONG, and each is why the rule is written out here
 * instead of summarised. It said the hazard was `(` OR `[` - it is only `(`. It said the hazard
 * was an inline code span - backticks are irrelevant, `a note `foo(bar` here` parses. It said
 * the hazard was an unbalanced bracket anywhere - only column 1 counts.
 *
 * WHY THE WIDTH SWEEP IS STILL HERE, AND WHY IT IS NOT CALLED A SIMULATION ANY MORE.
 *
 * GitHub re-wraps the pull request body when it builds the squash commit message, so the text
 * this check is handed is NOT the text git receives. Checking only the authored form is a hole,
 * and it cost a second dropped release the day after this script shipped: #315 parsed as
 * authored, and once folded its `DoesNotContain("control` landed at column 1 and did not.
 *
 * Folding at a few widths is how that case is reached. It is NOT a reproduction of GitHub's
 * wrapping - GitHub re-flows rather than folding line by line, and the exact width is
 * undocumented. It is a cheap way to move words to column 1, which is the only thing rule 1
 * cares about. Any width failing fails the check.
 *
 * INDENTATION IS FOLDED BOTH WAYS, PRESERVED AND STRIPPED, because rule 1 turns on it and the
 * evidence does not settle which GitHub does. Measured 2026-08-22 by diffing 22 authored pull
 * request bodies against the commits they became:
 *
 *   4-space indents SURVIVED    - 1 of 1 (#312's fenced block). One sample; thin.
 *   1-3 space indents were LOST - counts fell 5 to 3, 6 to 3, 4 to 1, and #315 lost both,
 *                                 consistent with markdown reflow absorbing bullet continuations.
 *
 * So neither answer is safe on its own, and this file has already been burned twice by modelling
 * GitHub from a single observation. Trying BOTH needs no model: whichever GitHub does, one of the
 * two variants is the text release-please gets.
 *
 * MEASURED COST: none. Replayed over the 140 merged pull requests #182-#331, the whole of this
 * change - both fold variants, the named token, the warning - alters ZERO pass/fail verdicts
 * against the version it replaced. The gate is what it was; only its reasoning and its message
 * are different.
 */
function wrapLine(line, width, keepIndent) {
  if (line.length <= width) return [line];
  const indent = keepIndent ? /^\s*/.exec(line)[0] : '';
  const out = [];
  let current = '';
  for (const word of line.slice(/^\s*/.exec(line)[0].length).split(' ')) {
    if (current === '') current = indent + word;
    else if ((current + ' ' + word).length <= width) current += ' ' + word;
    else { out.push(current); current = indent + word; }
  }
  if (current !== '') out.push(current);
  return out;
}

/** Does this string break the parse when it sits at the start of a body line? The rule above,
 *  in code. The agreement control asserts this against the parser rather than trusting it. */
function headerShapedAndUnterminated(s) {
  if (s === '' || /^\s/.test(s)) return false;
  const open = s.indexOf('(');
  if (open <= 0) return false;
  if (/\s/.test(s.slice(0, open))) return false;
  const close = s.indexOf(')');
  if (close !== -1 && close < open) return false;
  return !s.includes(')', open + 1);
}

/**
 * Words that would break the parse IF a fold put them at column 1. Reported as a WARNING and
 * never as a failure, and the severity is a measurement rather than a preference.
 *
 * Over 200 real commits on main this flags 19. Two of them are #312 and #315 - the two that
 * actually broke. The other 17 parsed, because GitHub's fold happened not to land there: they
 * are ordinary sentences in this repository's register, `ConvertAsync(html,` and
 * `WorkbookEditor.Protect(xlsx,` and `` `PdfFontOptions("Noto ``. Failing on those would block
 * 8.5% of good pull requests, and a check that does that is switched off within a month.
 *
 * They are still genuine LATENT hazards - they survived by where the fold landed, not by being
 * safe - so they are worth naming and not worth blocking. Indenting the line immunises it.
 */
function foldHazards(body) {
  const found = [];
  body.split('\n').forEach((line, i) => {
    if (headerShapedAndUnterminated(line)) return;   // already a hard failure; not a warning
    for (const word of line.split(/\s+/)) {
      if (word && headerShapedAndUnterminated(word)) found.push({ line: i + 1, word, text: line });
    }
  });
  return found;
}

function wrapBody(message, width, keepIndent = true) {
  const split = message.indexOf('\n\n');
  if (split < 0) return message;
  const subject = message.slice(0, split);
  const body = message.slice(split + 2);
  const folded = body.split('\n').flatMap(l => wrapLine(l, width, keepIndent)).join('\n');
  return subject + '\n\n' + folded;
}

/** Every folded form to check: each width, indentation kept and stripped. */
function variants(message) {
  const out = [];
  for (const width of WIDTHS) {
    out.push({ width, keepIndent: true, text: wrapBody(message, width, true) });
    out.push({ width, keepIndent: false, text: wrapBody(message, width, false) });
  }
  return out;
}

/** Widths to fold at. 89 is what GitHub was measured doing; the rest bracket it. */
const WIDTHS = [72, 80, 89, 100];

const parser = await loadParser();

if (process.argv.includes('--self-test')) {
  // THE AGREEMENT CONTROL, and the most load-bearing thing in this file.
  //
  // The comment above headerShapedAndUnterminated states a rule as fact. A rule written down
  // beside code it does not govern is precisely the defect this script was filed against, so it
  // is checked here against the parser itself rather than asserted.
  //
  // Every shape below was measured on 2026-08-22. If a dependency bump changes the grammar, this
  // control fails and names the shape - instead of the predicate quietly describing a parser
  // that no longer behaves that way.
  const SHAPES = [
    ['foo(bar', true],                       ['foo (bar', false],
    ['foo(bar)', false],                     ['a note foo(bar here', false],
    ['`foo(bar`', true],                     ['a note `foo(bar` here', false],
    ['foo[bar', false],                      ['foo{bar', false],
    ['(bar', false],                         ['x1(bar', true],
    ['a.b(bar', true],                       ['- foo(bar', false],
    ['  foo(bar', false],                    ['\tfoo(bar', false],
    ['foo)bar(baz', false],                  ['foo(a b c d e) trailing', false],
    ['foo(a(b', true],                       ['foo(a) then bar(b', false],
    ['x(y', true],                           ['x(', true],
    ['foo()', false],                        ['([#331](https://g.com/o/r/pull/331))', false],
    ['[the docs](https://example.com) and more', false],
    ['- `services.AddDocToolkit(o =>` here', false],
    ['`services.AddDocToolkit(o` here', true],
    ['(unclosed at the very start', false],
    ['services.AddDocToolkit(o =>', true],
    ['`DoesNotContain("control', true],
  ];

  let disagreed = 0;
  for (const [shape, expected] of SHAPES) {
    let parses = true;
    try { parser('fix(core): a subject that is fine\n\n' + shape); } catch { parses = false; }
    const predicted = headerShapedAndUnterminated(shape);
    if (parses === !expected && predicted === expected) continue;
    disagreed++;
    console.error(`SELF-TEST FAILED: ${JSON.stringify(shape)} - parser ${parses ? 'accepts' : 'rejects'}`
      + `, predicate says ${predicted ? 'hazard' : 'safe'}, expected ${expected ? 'hazard' : 'safe'}.`);
  }
  if (disagreed === 0) {
    console.log(`ok   the rule matches the parser on all ${SHAPES.length} measured shapes`);
  }

  // A fold hazard must WARN and never fail. Measured: hard-failing on this shape would have
  // blocked 17 of 200 good commits, which is how a check gets switched off.
  const hazardous = [
    'docs: show how to get a byte[] out of the process',
    '',
    'Call `ConvertAsync(html, fonts, ct)` and read the array back.',
  ].join('\n');

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
  const fFolded = variants(foldOnly).map(v => check(v.text, parser));
  let failed = false;

  if (!fAuthored.ok) {
    console.error('SELF-TEST FAILED: the fold-only control was expected to parse as authored.');
    failed = true;
  } else if (fFolded.every(r => r.ok)) {
    console.error('SELF-TEST FAILED: the fold-only control still parses at every width, so this');
    console.error('check cannot see the failure that dropped #315. The folding logic is not working.');
    failed = true;
  } else {
    const v = variants(foldOnly)[fFolded.findIndex(r => !r.ok)];
    console.log(`ok   fold-only control: parses as authored, breaks folded at ${v.width} columns`);
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

  const hz = check(hazardous, parser);
  const hzFolded = variants(hazardous).map(v => check(v.text, parser));
  const hzWords = foldHazards(hazardous.split('\n\n').slice(1).join('\n\n'));
  if (!hz.ok || !hzFolded.every(r => r.ok)) {
    console.error('SELF-TEST FAILED: the warning control was rejected. A latent fold hazard must');
    console.error('warn, not fail - blocking this shape blocks 8.5% of the good commits here.');
    failed = true;
  } else if (hzWords.length === 0) {
    console.error('SELF-TEST FAILED: the warning control raised no hazard, so the warning path is');
    console.error('dead and nothing would ever be reported. Expected `ConvertAsync(html,`.');
    failed = true;
  } else {
    console.log(`ok   latent fold hazard warns without failing: ${hzWords[0].word}`);
  }

  if (disagreed) failed = true;

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
  for (const v of variants(message)) {
    const attempt = check(v.text, parser);
    if (!attempt.ok) {
      result = attempt;
      foldedAt = `${v.width} columns with indentation ${v.keepIndent ? 'preserved' : 'stripped'}`;
      break;
    }
  }
}

/** Name the offending token, rather than leaving a reader to count columns. */
function diagnose(text) {
  const lines = text.split('\n');
  for (let i = 0; i < lines.length; i++) {
    if (!headerShapedAndUnterminated(lines[i])) continue;
    const token = lines[i].slice(0, lines[i].indexOf('(') + 1);
    return { line: i + 1, token, text: lines[i] };
  }
  return null;
}

const body = message.split('\n\n').slice(1).join('\n\n');
const latent = foldHazards(body);

if (result.ok) {
  console.log(`the squashed commit message parses as release-please reads it, `
    + `as authored and folded at ${WIDTHS.join(', ')} columns`);

  // Latent hazards WARN. They parsed only because GitHub's fold did not land on them, and the
  // measurement behind not failing here is in the comment on foldHazards.
  if (latent.length) {
    console.log('');
    console.log(`::warning::${latent.length} place(s) would break release-please if GitHub folded `
      + `the line there. This pull request is fine as written - the fold did not land on them.`);
    for (const h of latent.slice(0, 5)) {
      console.log(`::warning::line ${h.line}: ${h.word}`);
    }
    if (latent.length > 5) console.log(`::warning::...and ${latent.length - 5} more`);
    console.log('::warning::A line starting with something like `name(` and no `)` on that same');
    console.log('::warning::line reads as a Conventional Commits scope, and the whole commit is');
    console.log('::warning::discarded. Indenting the line immunises it, as does closing the');
    console.log('::warning::bracket in the same word.');
  }
  process.exit(0);
}

if (foldedAt !== null) {
  console.error(`::error::This parses as you wrote it and BREAKS once GitHub folds it to ${foldedAt}.`);
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

// WHEN THE FAILURE CAME FROM A FOLDED VARIANT, THE LINE NUMBER ABOVE IS IN TEXT NOBODY WROTE.
// The one message whose entire purpose is legibility was pointing at a line that does not exist
// in the pull request description. Name the authored line as well.
const authored = diagnose(message);
console.error('::error::');
if (foldedAt !== null && latent.length) {
  console.error(`::error::In YOUR text it is body line ${latent[0].line}: ${latent[0].text.slice(0, 90)}`);
  console.error(`::error::The token is  ${latent[0].word}`);
} else if (authored) {
  console.error(`::error::The token is  ${authored.token}  on line ${authored.line}.`);
}
console.error("::error::A line starting with something like `name(`, with no `)` on that same");
console.error('::error::line, reads as a Conventional Commits scope. Indenting the line immunises');
console.error('::error::it, as does closing the bracket within the same word.');
console.error('::error::');
console.error('::error::The PR TITLE is fine. The PR DESCRIPTION becomes this commit\'s body, and');
console.error('::error::something in it broke the parser. release-please discards the WHOLE');
console.error('::error::commit when that happens, so every feat: or fix: here would be missing');
console.error('::error::from the changelog - and if it is the only bumping commit here, no');
console.error('::error::Release PR would be opened at all.');
console.error('::error::');
console.error('::error::Fix the pull request description, then push or re-run this job.');
process.exit(1);
