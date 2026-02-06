---
name: rich-edit-fraction
description: This skill describes how to implement a custom two-dimensional fraction rendering engine on top of a standard Windows Rich Edit control. It uses a "subclassing and overlay" architecture, where standard text act as anchors, and GDI is used to draw mathematical formatting directly over the control's surface.
---



## Core Architecture: Subclassing & Overlays

Instead of modifying the Rich Edit's internal layout engine (which is extremely complex), this approach:
1.  **Subclasses** the control to intercept low-level messages.
2.  **Uses "Anchor Characters"**: Special characters (like the horizontal box-drawing character `\u2500`) are inserted into the text stream to reserve space.
3.  **Draws GDI Overlays**: Numerators and denominators are rendered manually over those anchors during the paint cycle.

---

## Key Implementation Steps

### 1. Subclassing the Control
Use `SetWindowLongPtr` to intercept messages before the control handles them.
```cpp
g_originalProc = (WNDPROC)SetWindowLongPtr(hWnd, GWLP_WNDPROC, (LONG_PTR)FractionRichEditProc);
```

### 2. Detecting the Fraction Trigger (`WM_CHAR`)
Intercept the `/` key. When detected, identify the preceding digits, replace them with a horizontal bar, and enter an "active" typing state for the denominator.

```cpp
if (ch == L'/') {
    if (!g_currentNumber.empty()) {
        const LONG numLen = (LONG)g_currentNumber.size();
        const LONG barLen = std::max<LONG>(3, numLen);
        std::wstring bar((size_t)barLen, L'\u2500'); // Unicode Horizontal Bar

        // Replace the numerator digits with the bar anchor
        SendMessage(hwnd, EM_SETSEL, (WPARAM)(caretPos - numLen), (LPARAM)caretPos);
        SendMessage(hwnd, EM_REPLACESEL, (WPARAM)TRUE, (LPARAM)bar.c_str());

        // Store fraction metadata
        FractionSpan span;
        span.barStart = caretPos - numLen;
        span.barLen = barLen;
        span.numerator = g_currentNumber;
        g_fractions.push_back(std::move(span));
        g_state.active = true; // Now intercept next digits as denominator
    }
}
```

### 3. Maintaining Synchronized Positions
As the user types elsewhere in the document, the character indices of your fraction "anchors" will shift. You must adjust your metadata on every `WM_CHAR` or `WM_REPLACESEL`.

```cpp
static void ShiftFractionsAfter(LONG atPosInclusive, LONG delta) {
    for (auto& f : g_fractions) {
        if (f.barStart >= atPosInclusive)
            f.barStart += delta;
    }
}
```

### 4. Coordinate Mapping (`EM_POSFROMCHAR`)
To know where to draw, you must find the screen coordinates of the "anchor" characters. Standard `EM_POSFROMCHAR` behavior varies between Rich Edit versions; use a robust wrapper.

```cpp
static bool TryGetCharPos(HWND hEdit, LONG charIndex, POINT& outPt) {
    POINTL ptL = {};
    if (SendMessage(hEdit, EM_POSFROMCHAR, (WPARAM)&ptL, (LPARAM)charIndex) != -1) {
        outPt.x = (int)ptL.x;
        outPt.y = (int)ptL.y;
        return true;
    }
    return false;
}
```

### 5. Dynamic Scaling & Zoom Measurement
To ensure fractions scale exactly with the editor's zoom level, you must measure the distance between anchor characters on screen and compare it to the font's ideal metrics.

```cpp
renderScale = (double)(pN.x - p0.x) / (double)(unzoomed_char_width * f.barLen);
```
> **How it works**: By calculating the ratio between the *actual* rendered pixel width of the bar characters and their *expected* unzoomed width, you derive a precise scale factor that accounts for standard Rich Edit zoom (`EM_SETZOOM`), high-DPI displays, and font-smoothing.

### 6. Vertical Geometry & "Metric Locking"
To prevent coordinates from sliding relative to the horizontal bar during zoom, the vertical anchor must be tethered to a system-reported metric.

```cpp
const int yMid = ptStart.y + (int)(tmBase.tmHeight);
```
> **Accommodates to**: This single line ensures that the fraction's visual center (`yMid`) is locked to a fixed percentage of the current line's bounding box. Because `tmBase.tmHeight` is updated by Windows GDI on every zoom level change, the entire fraction block (numerator, bar, and denominator) moves as a single unit, maintaining zero "drift" relative to the editor text. By setting the anchor to the full line height, the fraction is positioned cleanly below the baseline of the main text flow.

### 7. Dynamic Margin Spacing
Separation must scale linearly with the font size to avoid looking cramped at small sizes or disconnected at large sizes.

```cpp
const int gap = std::max<int>(4, (int)(tmBase.tmHeight / 4));
```
> **Accommodates to**: By deriving the `gap` as a percentage of `tmBase.tmHeight`, the "clear space" between the numbers and the line remains visually identical whether the text is 10 pixels high or 100 pixels high. The `max` call ensures a minimum safe distance (4px) to prevent intersections when the bar rendering thickens at high zoom levels.

### 8. Coordinate Projection & Scaling
Zooming is handled by scaling the custom fonts based on the on-screen width of the anchor.

```cpp
renderScale = (pN.x - p0.x) / (unzoomedSize.cx * f.barLen);
```
> **Accommodates to**: This handles the transition between the Rich Edit's internal coordinate system and your custom GDI rendering. It measures the *actual* growth in horizontal pixels and applies it back to the vertical fonts, ensuring the fraction stays proportional in both dimensions simultaneously.
---

## Targeted Messages
- `WM_CHAR`: Handle input and fraction creation.
- `WM_PAINT`: Perform the custom GDI rendering.
- `WM_MOUSEWHEEL`: Invalidate the window to force a redraw when the user zooms.
- `WM_KEYDOWN`: Invalidate state if the user navigates away or deletes characters.

## Done Criteria
1. Typing `3/4` results in a vertically stacked fraction.
2. The fraction scales perfectly when the user zooms in or out.
3. The numerator and denominator do not overlap the horizontal middle line.


