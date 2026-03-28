```skill
---
name: adding-math-feature
description: Add a new structured math feature to windeskapp with RichEdit visualization (GDI overlay) and evaluator support, following the same architecture as summation/integral/fraction.
---

## Purpose
Use this playbook when adding a new math construct (for example, a product symbol, limit notation, matrix-like block, or another operator similar to big sigma). The feature is complete only when **both** are implemented:
1. Visual editing and rendering in the RichEdit overlay.
2. Numeric/symbolic evaluation path in the math engine.

## Architecture Checklist
A new feature touches these layers:
- `src/math_types.h`: add a new `MathType` enum value.
- `src/math_editor.cpp`: add command trigger (e.g. `\\prod` + Space), typing-part navigation, and calculate trigger behavior.
- `src/math_renderer.cpp`: draw the notation with GDI and define hit-testing rectangles.
- `src/math_manager.cpp`: compute result formatting and dispatch to evaluator.
- `src/math_evaluator.cpp`: parse/evaluate the expression and helper syntax.
- Tests/scripts: add one evaluator test and one UI smoke script if applicable.

## Step-by-Step

### 1) Define data shape
In `MathObject`, reuse `part1/part2/part3` with a clear contract for the new type.
Example convention:
- `part1`: upper/first parameter
- `part2`: lower/second parameter
- `part3`: body/expression

Document this mapping in comments where you add the feature.

### 2) Add command creation in editor
In `WM_CHAR` command parsing (`g_currentCommand`), add a branch:
- Assign `obj.type = MathType::<NewType>`
- Set default `part` values
- Insert anchor text (`NBSP` or line anchors) with suitable length
- Initialize `state.activePart` to the most useful first-edit field

Also update:
- `VK_TAB` cycling logic (`maxP`)
- Enter/equals behavior if custom evaluation timing differs

### 3) Render it like existing notation (big sigma style)
In `MathRenderer::Draw`:
- Measure baseline/font metrics (`tmBase`/`tmL`)
- Draw primary symbol (Cambria Math where needed)
- Draw parts with `DrawPart(...)`
- Draw result text to the right in active color

For notation similar to summation, mirror the pattern already used by `MathType::Summation`:
- center symbol on `xCenter`,`yMid`
- upper/lower limits offset vertically
- expression to the right

In `MathRenderer::GetHitPart`:
- Add hit rectangles for each editable part
- Keep them generous enough for reliable mouse selection

### 4) Add evaluation path
In `MathManager::CalculateResult`:
- Add `else if (obj.type == MathType::<NewType>)`
- Evaluate inputs via `MathEvaluator`
- Return a numeric result (or 0 for invalid domain)

If result format is complex (like systems), add a dedicated formatter method and call it from editor update/trigger flow.

### 5) Extend parser semantics
In `MathEvaluator`:
- Add any new function/operator grammar in parse functions
- Preserve existing precedence rules
- Keep invalid/domain cases safe (return 0 instead of crashing)

For function-like operators, support both `(...)` and `{...}` argument grouping if consistent with existing behavior.

### 6) Verify end-to-end
Minimum verification:
- Build project successfully.
- Insert notation via command in app.
- Edit all parts with mouse + Tab.
- Trigger calculation (`=` or Enter path).
- Confirm result rendering and no overlap artifacts.

Recommended tests:
- `test_eval.cpp` case for expected value.
- One invalid-input case returns 0 (or explicit error text for specialized types).

## Practical Example Pattern (Sigma-like)
When adding a symbol that behaves like summation:
- Visual: big symbol + upper/lower limits + right-side expression.
- Evaluation: iterate variable from lower to upper, accumulate body expression.
- Syntax helper: parse lower limit with `var=start` fallback.

## Done Criteria
A feature is done only if all are true:
- Command insertion works from editor text input.
- Typing state and hit-test switching work for every part.
- Renderer shows stable, scaled notation at different zooms.
- Evaluation path returns correct result and handles invalid input safely.
- At least one automated or scripted test validates behavior.
```
