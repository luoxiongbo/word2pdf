'use strict';

// Minimal ANSI colours — no dependency needed
const c = {
  reset:  '\x1b[0m',
  bold:   '\x1b[1m',
  dim:    '\x1b[2m',
  green:  '\x1b[32m',
  yellow: '\x1b[33m',
  red:    '\x1b[31m',
  cyan:   '\x1b[36m',
  grey:   '\x1b[90m',
};

// Respect NO_COLOR / non-TTY environments
const useColor = process.stdout.isTTY && !process.env.NO_COLOR;
const col = (code, text) => useColor ? `${code}${text}${c.reset}` : text;

const log = {
  info:       (msg)           => console.log(col(c.cyan,   `ℹ ${msg}`)),
  success:    (file)          => console.log(col(c.green,  `✔ ${file}`)),
  skip:       (file)          => console.log(col(c.grey,   `↷ skipped (exists): ${file}`)),
  warn:       (msg)           => console.warn(col(c.yellow, `⚠ ${msg}`)),
  error:      (msg)           => console.error(col(c.red,   `✖ ${msg}`)),
  converting: (file)          => process.stdout.write(col(c.dim,  `  converting ${file}…\r`)),
  summary:    (ok, fail) => {
    const okStr   = col(c.green, `${ok} converted`);
    const failStr = fail > 0 ? col(c.red, ` ${fail} failed`) : '';
    console.log(`\n${col(c.bold, 'Done.')} ${okStr}${failStr}`);
  },
};

module.exports = { log };
