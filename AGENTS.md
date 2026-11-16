# AGENTS.md

Rules for every AI coding agent working in this repo. Claude Code reads it through
[`CLAUDE.md`](CLAUDE.md), a one-line `@AGENTS.md` import; put rules here, not there.

## Hard rules

- **No AI authorship in commits.** Never add a `Co-Authored-By:` trailer naming
  an AI (Claude, Copilot, etc.), a "Generated with …" line, or any similar
  attribution to a commit message, PR title, or PR body. AI assistance is
  already disclosed once, at the project level, at the bottom of
  [`ReadMe.md`](ReadMe.md) — repeating it per commit is noise. Commits are
  authored by the human who owns the change. Write the message as a plain
  description of *what changed and why*.
