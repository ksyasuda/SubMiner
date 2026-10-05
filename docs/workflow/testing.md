<!-- read_when: writing, reviewing, or deleting tests -->

# Testing Guidelines

Status: active  
Last verified: 2026-10-04  
Owner: Kyle Yasuda  
Read when: adding or changing tests

Which lanes to run lives in [Verification](./verification.md). This page covers what a
good test looks like here.

## Test Behavior, Not Wiring

- Exercise a module through its public interface and assert outcomes: return values,
  resulting state, files written, rows in sqlite, messages sent to a fake process.
- Prefer real collaborators that are cheap: tmp dirs, in-memory or tmp sqlite, fake
  executables, pure parsers. Fake only the edges (Electron, mpv socket, network).
- Assert call order only when order is the behavior (startup sequencing, ABA races),
  and then assert relative order, not the whole call log.

## Do Not Write

- Pass-through tests: a builder or handler that only forwards deps (`*-main-deps.ts`,
  `*-runtime-handlers.ts`) needs no test. TypeScript already checks the shape.
- Source-text tests: reading `.ts`, `.css`, `.html`, `.lua`, `.swift`, or workflow YAML
  and asserting with regexes. If the invariant matters, extract the logic and run it
  (see `src/preload-ipc-listeners.ts`). Policy lints over workflow YAML are the
  exception and belong in one parsed-YAML check.
- "Stays deleted" tests for removed modules, flags, or selectors. `tsc` and review
  catch reintroduction.
- Tests that restate constants, defaults, labels, or window options.
- The same rule at several layers. Test a rule once where it lives; keep one
  end-to-end case (for the tokenizer, the golden corpus) for integration.

## Keep Tests Small

- Shared setup goes in a `makeDeps(overrides)` helper or a `*-test-harness.ts` file, not
  pasted into every test.
- Near-identical cases become a table: `for (const c of cases) test(c.name, ...)`.
- Inject timers, clocks, and loggers instead of patching globals or private members.
- A bug fix gets a regression test for the behavior that broke, named after the
  behavior, not the ticket.

## Known Debt

- `src/main/main-wiring.test.ts` and `src/renderer/renderer-init-order.test.ts` still
  regex `main.ts` / `renderer.ts` to guard startup ordering. Replace each test with a
  behavioral one as the logic moves out of `main.ts` closures.
