# AG Schuck Conference Tracker — repo instructions for the scheduled routine

This repo is maintained by a recurring scheduled Claude Code routine that
scouts quantum-photonics conferences/workshops/schools for AG Schuck
(integrated quantum photonics group, University of Münster). Each run
gets assigned a fresh, randomly-named work branch
(`claude/<adjective>-<name>-<hash>`) by the scheduler — that name is
throwaway infrastructure, not something to build a workflow around.

## Critical: always land your work on the default branch

**A run is not done when you push your branch — it's done when your
changes are merged into this repository's actual default branch.**

Do not assume the default branch is `main`. Check it (e.g. via the
GitHub API/MCP `get_file_contents` with no `ref`, or `git remote show
origin`) — historically this repo's default branch has been named
`claude/push-to-conferences-repo-OVzNZ` because it was itself created by
one of these scheduled sessions.

Concretely, at the end of every run:
1. Open a PR from your work branch into the actual default branch.
2. Merge it yourself. Don't leave it sitting open, and don't stop after
   just pushing the branch — an unmerged branch is invisible to anyone
   who just visits the repo on github.com.
3. Delete your work branch after it's merged.

**Why this matters:** between 2026-05-13 and 2026-09-13, this routine ran
~17 times. Every single run except the first pushed its snapshot to its
own branch and stopped — none of them merged. The result: 16+ weeks of
research sat invisible in orphaned branches while the repo's default
branch (and anyone looking at it) still only showed the original
2026-05-13 snapshot. One of those orphaned runs even tried to fix this
exact problem (adding auto-merge logic) — and that fix itself was never
merged, for the same reason. Those stray branches have since been
deleted; don't repeat the pattern.

## Always build on the latest MERGED snapshot, not a stale one

Before researching updates, fetch/read the default branch's current
`conferences_*.tsv` — that is this run's starting point. Do not
regenerate from an old baseline (e.g. always diffing against
2026-05-13) just because that's what an earlier run happened to use.
Each run's diff/changelog should be described relative to the actual
previous snapshot, not a fixed historical date.

## Keep the repo to a single current snapshot

The repo should contain exactly:
- `build_snapshot.py` — the generator script
- `conferences_YYYY-MM-DD.tsv` and `conferences_YYYY-MM-DD.xlsx` — the
  **current** snapshot only

When publishing a new snapshot, delete the previous dated
`conferences_*.tsv`/`.xlsx` pair in the same PR (rows are never lost —
they live on as `archived` rows inside the TSV itself, per the tracker's
own quality rules — so there is no need to keep old snapshot files
around as a separate history mechanism; git history already preserves
every prior version if it's ever needed).

## Recap of the tracker's data rules (see original task prompt for full detail)

- Never delete a conference's row; mark it `archived` instead once it has
  passed. This is how next year's edition gets spotted.
- Verify facts against official event websites, not aggregators.
- Never invent dates — use `TBA` when unknown.
- Prefer European events but always keep the global flagships (CLEO,
  OFC, Photonics West, ECOC, QIP, ECIO, SPW, GRC series, Benasque/Les
  Houches/Cargèse schools, etc.), including TBA placeholders for
  predictable next editions that haven't been announced yet.
