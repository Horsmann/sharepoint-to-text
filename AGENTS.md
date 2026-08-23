# 🤖 Agent Execution Protocol

> **This document is the source of truth for all agent behavior in this repository.**
> Tasks are not considered complete unless every criterion in this protocol is satisfied.
> No exceptions.

---

## Table of Contents

1. [Core Principles](#1-core-principles)
2. [Planning & Verifiable States](#2-planning--verifiable-states)
3. [Code Quality Standards](#3-code-quality-standards)
4. [Progress Tracking](#4-progress-tracking)
5. [Mandatory Validation](#5-mandatory-validation)
6. [Completion Criteria](#6-completion-criteria)
7. [Quick Reference Checklist](#7-quick-reference-checklist)

---

## 1. Core Principles

Every agent operating in this repository must embody the following values:

| Principle | Description |
|-----------|-------------|
| **Correctness over speed** | A slow, correct solution is always preferred over a fast, broken one. |
| **Verifiability** | Every completed step must have a concrete, checkable proof of completion. |
| **Transparency** | Progress and blockers must be visible in the checklist at all times. |
| **Clean code** | Code must be readable, documented, and maintainable by the next person who touches it. |

---

## 2. Planning & Verifiable States

**Before writing a single line of code**, the agent must produce a written plan.

### 2.1 What a Good Plan Looks Like

Each plan must contain:

- A **clear goal statement** — one sentence describing what done looks like.
- A **numbered list of steps** — granular enough that each step maps to a single concern.
- A **verifiable state** per step — an objective, testable condition that proves the step is complete.

### 2.2 Verifiable State Examples

| Step | ❌ Vague State | ✅ Verifiable State |
|------|---------------|---------------------|
| Add a function | "Function exists" | "Function returns expected output for all defined test cases" |
| Refactor module | "Code looks cleaner" | "`uv run mypy .` reports zero errors on the module" |
| Fix a bug | "Bug seems fixed" | "Previously failing test now passes; no regressions introduced" |
| Add an endpoint | "Endpoint added" | "Integration test hits the endpoint and receives HTTP 200 with correct schema" |

> **Rule:** If a verifiable state cannot be checked automatically or by inspection in under 60 seconds, make it more specific.

---

## 3. Code Quality Standards

All code produced or modified by an agent must meet these standards **before** marking any step complete.

### 3.1 Docstrings

Every public function, method, and class **must** have a docstring. Use the following format:

```python
def calculate_discount(price: float, rate: float) -> float:
    """Calculate the discounted price after applying a percentage rate.

    Args:
        price: The original price in the base currency unit.
        rate: The discount rate as a decimal between 0.0 and 1.0.

    Returns:
        The final price after the discount has been applied.

    Raises:
        ValueError: If rate is not between 0.0 and 1.0.

    Example:
        >>> calculate_discount(100.0, 0.2)
        80.0
    """
    if not 0.0 <= rate <= 1.0:
        raise ValueError(f"Rate must be between 0.0 and 1.0, got {rate}")
    return price * (1 - rate)
```

**Docstring checklist:**
- [ ] One-line summary (imperative mood: "Calculate...", not "Calculates...")
- [ ] `Args` section for every parameter (type + purpose)
- [ ] `Returns` section describing what is returned
- [ ] `Raises` section if any exceptions are raised
- [ ] `Example` block for non-trivial functions

### 3.2 Clean Code Rules

```
✅  DO                                  ❌  DON'T
─────────────────────────────────────────────────────────────
Use descriptive names                   Use single-letter variables (except loop counters)
Keep functions under 30 lines           Write monolithic functions
One responsibility per function         Mix I/O with business logic
Prefer explicit over implicit           Leave unexplained magic numbers/strings
Write self-documenting code             Rely on comments to explain bad names
Handle errors explicitly                Silence exceptions with bare `except:`
Use type hints everywhere               Leave function signatures untyped
```

### 3.3 Type Hints

All function signatures must include complete type hints:

```python
# ✅ Correct
def fetch_user(user_id: int, include_inactive: bool = False) -> User | None:
    ...

# ❌ Wrong — no type information
def fetch_user(user_id, include_inactive=False):
    ...
```

### 3.4 Unit Test Requirements

Every new behavior and every regression fix **must** be verified by one or more
new unit tests that fail without the corresponding implementation or fix.

Existing unit tests are protected regression coverage:

- **Do not remove existing unit tests.**
- **Do not remove, weaken, or otherwise alter the behavioral content of an
  existing unit test.** This includes its scenarios, inputs, assertions,
  expected outcomes, and error cases.
- Test scaffolding may be refactored when necessary. Changes to data-type
  shells, fixtures, helper methods, variable names, parameter shapes, or the
  mechanism used to provide information are permitted only when all of the
  original test information, scenarios, and behavioral assertions remain
  semantically unchanged.
- If removing or changing any existing test content is genuinely necessary,
  stop and explain exactly what content must change, why the change is
  necessary, and what coverage would be affected. Obtain explicit human
  confirmation before making the change. Agent judgment alone is not
  sufficient authorization.

> **Rule:** A production-code change is not complete when it only makes the
> existing suite pass. New behavior and regression fixes require new unit-test
> coverage, and existing coverage must remain intact unless a human explicitly
> approves a justified exception.

---

## 4. Progress Tracking

The agent maintains a live checklist throughout the task. This checklist is the single source of truth for what has been done and what remains.

### 4.1 Checklist Format

```markdown
## Task: <Goal Statement>

### Plan

- [ ] Step 1 — <description>
      Verifiable state: <how to confirm this is done>

- [ ] Step 2 — <description>
      Verifiable state: <how to confirm this is done>

- [x] Step 3 — <description> ✓ completed
      Verifiable state: all unit tests for this module pass

### Validation

- [ ] `uv run pytest` — all tests pass
- [ ] `uv run mypy .` — zero type errors
- [ ] `uv run pre-commit run --all-files` — all hooks pass
```

### 4.2 Checklist Rules

- Mark steps `[x]` **only** when their verifiable state is confirmed — not when the work feels done.
- Never skip or silently reorder steps. If a step needs to change, update the plan explicitly.
- If a blocker is encountered, record it in the checklist rather than working around it silently.

---

## 5. Mandatory Validation

**No task is complete without passing all three validation commands.** This is non-negotiable.

### 5.1 Run Tests

```bash
uv run pytest
```

All tests must pass. Zero failures. Zero errors.

### 5.2 Run Type Checks

```bash
uv run mypy .
```

Zero type errors. Zero unresolved imports on code the agent has touched.

### 5.3 Run Pre-commit Hooks

```bash
uv run pre-commit run --all-files
```

All hooks must pass against all files.

### 5.4 Failure Protocol

If any command fails, the agent must:

1. **Diagnose** — read the full error output and identify the root cause.
2. **Fix** — apply a targeted fix. Do not suppress errors with `# type: ignore` or skip/xfail tests unless there is an explicit, documented reason.
3. **Re-run** — execute all three commands again from scratch.
4. **Repeat** until all three commands exit cleanly.

> ⚠️ **Never declare completion while errors remain.** A passing checklist with failing tests is a lie.

---

## 6. Completion Criteria

A task is **complete** when all of the following are true:

```
✅  All planned steps are marked [x]
✅  Every verifiable state has been confirmed
✅  uv run pytest — passes with zero failures
✅  uv run mypy . — passes with zero errors
✅  uv run pre-commit run --all-files — all hooks pass
✅  All new/modified code has docstrings
✅  All new/modified functions have complete type hints
✅  New behavior and regression fixes have new unit-test coverage
✅  Existing test content remains semantically intact, or a justified change was explicitly approved by a human
```

If any item above is not satisfied, the task is **not complete**.

---

## 7. Quick Reference Checklist

Copy this block at the start of every task:

```markdown
## Task: <goal>

### Steps
- [ ] Step 1 — ...  |  Verifiable: ...
- [ ] Step 2 — ...  |  Verifiable: ...
- [ ] Step 3 — ...  |  Verifiable: ...

### Code Quality
- [ ] Docstrings added to all public functions/classes
- [ ] Type hints present on all signatures
- [ ] No magic numbers — constants are named and documented
- [ ] Error handling is explicit (no bare `except:`)

### Test Coverage
- [ ] New behavior and regression fixes have new unit tests
- [ ] Existing tests and their behavioral content remain intact
- [ ] Any necessary test-content change has an explanation and explicit human confirmation

### Validation
- [ ] `uv run pytest` passes
- [ ] `uv run mypy .` passes
- [ ] `uv run pre-commit run --all-files` passes
```

---

*Failure to follow this protocol means the task is incomplete. Always prefer correctness and verification over speed.*
