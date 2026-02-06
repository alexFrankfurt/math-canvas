# Skill: Rich Edit Custom Fraction Rendering

## Description
This skill describes how to implement a custom two-dimensional fraction rendering engine on top of a standard Windows Rich Edit control. It uses a "subclassing and overlay" architecture, where standard text act as anchors, and GDI is used to draw mathematical formatting directly over the control's surface.

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
Intercept the `/` key. When detected:
- Identify the preceding digits (the numerator).
- Replace the numerator text with a series of horizontal bar characters (`\u2500`).
- Store the mathematical state (numerator/denominator strings) in a custom side-car structure.

### 3. Coordinate Mapping (`EM_POSFROMCHAR`)
To know where to draw, you must find the screen coordinates of the "anchor" characters.
```cpp
POINT ptStart = {};
SendMessage(hEdit, EM_POSFROMCHAR, (WPARAM)&ptStart, (LPARAM)charIndex);
```

### 4. Dynamic Scaling & Zoom Support
Rich Edit controls can be zoomed (`EM_SETZOOM`). To ensure fractions scale correctly, calculate a "Render Scale" by comparing the actual pixel width of the anchor characters to their ideal font width.
```cpp
double scale = (double)actual_pixel_width / (double)font_metrics_width;
```

### 5. Vertical Geometry & Layout
For a professional look, align the vertical midpoint of the fraction with the bottom of the logical text line.
```cpp
// Logic discovered for optimal alignment:
const int yMid = ptStart.y + tmBase.tmHeight; 
const int gap = tmBase.tmHeight / 4;

int yNum = yMid - gap - tmNum.tmHeight; // Numerator above
int yDen = yMid + gap;                 // Denominator below
```

### 6. Painting the Overlay (`WM_PAINT` / `WM_PRINTCLIENT`)
Call the original window procedure first so the Rich Edit draws its text, then use `GetDC` or the provided `wParam` (for `WM_PRINTCLIENT`) to draw your numbers over it.
- **SetBkMode(hdc, TRANSPARENT)** is essential so your numbers don't wipe out the horizontal bar.
- **Use High-Quality Fonts**: Create scaled fonts with `CreateFontIndirectW` using the Render Scale calculated in Step 4.

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
