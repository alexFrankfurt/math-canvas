#include "math_renderer.h"
#include "math_manager.h"
#include <richedit.h>
#include <algorithm>

// Undefine Windows min/max macros to use std::min/std::max
#ifdef max
#undef max
#endif
#ifdef min
#undef min
#endif

namespace {
    static const bool kDebugOverlay = false;

    struct NodeMetrics
    {
        int cx = 0;
        int ascent = 0;
        int descent = 0;

        int Height() const
        {
            return ascent + descent;
        }
    };

    static NodeMetrics MeasureMathNodeMetrics(HDC hdc, const MathNode& node, const TEXTMETRICW& tmBase);
    static NodeMetrics MeasureMathSequenceMetrics(HDC hdc, const std::vector<MathNode>& nodes, const TEXTMETRICW& tmBase);
    static SIZE MeasureMathNode(HDC hdc, const MathNode& node, const TEXTMETRICW& tmBase);
    static void DrawMathNode(HDC hdc, const MathNode& node, int x, int baseline, const TEXTMETRICW& tmBase, COLORREF color);
    static bool TryGetSequenceCaret(HDC hdc, const std::vector<MathNode>& nodes, const std::vector<size_t>& path, size_t pathOffset, int x, int baseline, const TEXTMETRICW& tmBase, POINT& outPt);
    static bool TryGetNodeCaret(HDC hdc, const MathNode& node, size_t nodeIndex, const std::vector<size_t>& path, size_t pathOffset, int x, int baseline, const TEXTMETRICW& tmBase, POINT& outPt);
    static bool HitTestMathNodeSequence(HDC hdc, const std::vector<MathNode>& nodes, int x, int baseline, const TEXTMETRICW& tmBase, POINT ptMouse, const std::vector<size_t>& prefix, std::vector<size_t>& outPath);
    static bool HitTestMathNode(HDC hdc, const MathNode& node, size_t nodeIndex, int x, int baseline, const TEXTMETRICW& tmBase, POINT ptMouse, const std::vector<size_t>& prefix, std::vector<size_t>& outPath);

    static SIZE MeasureDisplayText(HDC hdc, const std::wstring& text)
    {
        SIZE textSize = {};
        const wchar_t* display = text.empty() ? L"?" : text.c_str();
        const int length = text.empty() ? 1 : (int)text.size();
        GetTextExtentPoint32W(hdc, display, length, &textSize);
        return textSize;
    }

    static NodeMetrics MeasureDisplayTextMetrics(HDC hdc, const std::wstring& text, const TEXTMETRICW& tmBase)
    {
        NodeMetrics metrics = {};
        const SIZE size = MeasureDisplayText(hdc, text);
        metrics.cx = size.cx;
        metrics.ascent = tmBase.tmAscent;
        metrics.descent = tmBase.tmDescent;
        return metrics;
    }

    static bool NodeHasVisibleContent(const MathNode& node)
    {
        if (node.kind == MathNodeKind::Text)
            return !node.text.empty();

        if (node.kind == MathNodeKind::Group)
        {
            for (const auto& child : node.children)
            {
                if (NodeHasVisibleContent(child))
                    return true;
            }
            return false;
        }

        return true;
    }

    static bool SequenceHasVisibleContent(const std::vector<MathNode>& nodes)
    {
        for (const auto& node : nodes)
        {
            if (NodeHasVisibleContent(node))
                return true;
        }
        return false;
    }

    static bool SlotHasVisibleContent(const MathObject& obj, int partIndex)
    {
        const size_t slotIndex = MathObject::SlotIndexFromPart(partIndex);
        if (slotIndex < obj.slots.size())
        {
            const MathSlot& slot = obj.slots[slotIndex];
            return !slot.text.empty() || SequenceHasVisibleContent(slot.children);
        }
        return !obj.SlotText(partIndex).empty();
    }

    static void SetSequenceTailPath(const std::vector<MathNode>& nodes, const std::vector<size_t>& prefix, std::vector<size_t>& outPath)
    {
        outPath = prefix;
        outPath.push_back(nodes.empty() ? 0 : nodes.size() - 1);
    }

    static SIZE MeasureMathNodeSequence(HDC hdc, const std::vector<MathNode>& nodes, const TEXTMETRICW& tmBase)
    {
        const NodeMetrics metrics = MeasureMathSequenceMetrics(hdc, nodes, tmBase);
        SIZE total = { metrics.cx, metrics.Height() };
        return total;
    }

    static NodeMetrics MeasureMathSequenceMetrics(HDC hdc, const std::vector<MathNode>& nodes, const TEXTMETRICW& tmBase)
    {
        NodeMetrics total = {};
        for (const auto& node : nodes)
        {
            const NodeMetrics nodeMetrics = MeasureMathNodeMetrics(hdc, node, tmBase);
            total.cx += nodeMetrics.cx;
            total.ascent = (std::max)(total.ascent, nodeMetrics.ascent);
            total.descent = (std::max)(total.descent, nodeMetrics.descent);
        }
        return total;
    }

    static bool TryGetSequenceCaret(HDC hdc, const std::vector<MathNode>& nodes, const std::vector<size_t>& path, size_t pathOffset, int x, int baseline, const TEXTMETRICW& tmBase, POINT& outPt)
    {
        if (pathOffset >= path.size())
        {
            outPt.x = x + MeasureMathNodeSequence(hdc, nodes, tmBase).cx;
            outPt.y = baseline;
            return true;
        }

        const size_t targetNodeIndex = path[pathOffset];
        if (targetNodeIndex >= nodes.size())
            return false;

        int cursorX = x;
        for (size_t nodeIndex = 0; nodeIndex < targetNodeIndex; ++nodeIndex)
            cursorX += MeasureMathNode(hdc, nodes[nodeIndex], tmBase).cx;

        const MathNode& node = nodes[targetNodeIndex];
        if (pathOffset + 1 >= path.size())
        {
            outPt.x = cursorX + MeasureMathNode(hdc, node, tmBase).cx;
            outPt.y = baseline;
            return true;
        }

        return TryGetNodeCaret(hdc, node, targetNodeIndex, path, pathOffset + 1, cursorX, baseline, tmBase, outPt);
    }

    static bool TryGetNodeCaret(HDC hdc, const MathNode& node, size_t nodeIndex, const std::vector<size_t>& path, size_t pathOffset, int x, int baseline, const TEXTMETRICW& tmBase, POINT& outPt)
    {
        const int pad = (std::max<int>)(2, tmBase.tmHeight / 8);
        const int radicalW = (std::max<int>)(8, tmBase.tmAveCharWidth);
        const int overlineGap = (std::max<int>)(2, tmBase.tmHeight / 10);
        const int fracGap = (std::max<int>)(3, tmBase.tmHeight / 6);
        const int absPad = (std::max<int>)(3, tmBase.tmHeight / 7);
        const int superGap = (std::max<int>)(2, tmBase.tmAveCharWidth / 3);

        (void)nodeIndex;

        if (node.kind == MathNodeKind::Group)
            return TryGetSequenceCaret(hdc, node.children, path, pathOffset, x, baseline, tmBase, outPt);

        if (pathOffset >= path.size())
        {
            outPt.x = x + MeasureMathNode(hdc, node, tmBase).cx;
            outPt.y = baseline;
            return true;
        }

        if (!node.IsStructural())
        {
            SIZE textSize = MeasureDisplayText(hdc, node.text);
            outPt.x = x + textSize.cx;
            outPt.y = baseline;
            return true;
        }

        const size_t slotIndex = path[pathOffset];

        if (node.kind == MathNodeKind::SquareRoot)
        {
            const NodeMetrics childMetrics = MeasureMathSequenceMetrics(hdc, node.SlotNodes(0), tmBase);
            const NodeMetrics indexMetrics = MeasureMathSequenceMetrics(hdc, node.SlotNodes(1), tmBase);
            const int radTop = baseline - (std::max)((int)tmBase.tmAscent, childMetrics.ascent) - overlineGap - 1;
            const int xRadStart = x + indexMetrics.cx;
            const int xRadPeak = xRadStart + radicalW;
            const int xChild = xRadPeak + pad;
            if (slotIndex == 1)
                return TryGetSequenceCaret(hdc, node.SlotNodes(1), path, pathOffset + 1, x, radTop + indexMetrics.ascent, tmBase, outPt);
            if (slotIndex == 0)
                return TryGetSequenceCaret(hdc, node.SlotNodes(0), path, pathOffset + 1, xChild, baseline, tmBase, outPt);
            return false;
        }

        if (node.kind == MathNodeKind::Fraction)
        {
            const NodeMetrics numMetrics = MeasureMathSequenceMetrics(hdc, node.SlotNodes(0), tmBase);
            const NodeMetrics denMetrics = MeasureMathSequenceMetrics(hdc, node.SlotNodes(1), tmBase);
            const int barWidth = (std::max)(numMetrics.cx, denMetrics.cx) + pad * 2;
            const int numX = x + (barWidth - numMetrics.cx) / 2;
            const int denX = x + (barWidth - denMetrics.cx) / 2;
            if (slotIndex == 0)
                return TryGetSequenceCaret(hdc, node.SlotNodes(0), path, pathOffset + 1, numX, baseline - fracGap - numMetrics.descent, tmBase, outPt);
            if (slotIndex == 1)
                return TryGetSequenceCaret(hdc, node.SlotNodes(1), path, pathOffset + 1, denX, baseline + fracGap + denMetrics.ascent, tmBase, outPt);
            return false;
        }

        if (node.kind == MathNodeKind::Power)
        {
            const NodeMetrics baseMetrics = MeasureMathSequenceMetrics(hdc, node.SlotNodes(0), tmBase);
            const NodeMetrics expMetrics = MeasureMathSequenceMetrics(hdc, node.SlotNodes(1), tmBase);
            const int exponentBaseline = baseline - (std::max)((int)tmBase.tmAscent / 2, expMetrics.descent);
            if (slotIndex == 0)
                return TryGetSequenceCaret(hdc, node.SlotNodes(0), path, pathOffset + 1, x, baseline, tmBase, outPt);
            if (slotIndex == 1)
                return TryGetSequenceCaret(hdc, node.SlotNodes(1), path, pathOffset + 1, x + baseMetrics.cx + superGap, exponentBaseline, tmBase, outPt);
            return false;
        }

        if (node.kind == MathNodeKind::AbsoluteValue)
        {
            if (slotIndex == 0)
                return TryGetSequenceCaret(hdc, node.SlotNodes(0), path, pathOffset + 1, x + absPad * 2, baseline, tmBase, outPt);
            return false;
        }

        if (node.kind == MathNodeKind::Logarithm)
        {
            SIZE logSize = {};
            GetTextExtentPoint32W(hdc, L"log", 3, &logSize);
            const NodeMetrics baseMetrics = MeasureMathSequenceMetrics(hdc, node.SlotNodes(0), tmBase);
            const int xBase = x + logSize.cx + 1;
            const int baseBaseline = baseline + (std::max)((int)tmBase.tmDescent / 2, baseMetrics.ascent / 4);
            if (slotIndex == 0)
                return TryGetSequenceCaret(hdc, node.SlotNodes(0), path, pathOffset + 1, xBase, baseBaseline, tmBase, outPt);
            if (slotIndex == 1)
                return TryGetSequenceCaret(hdc, node.SlotNodes(1), path, pathOffset + 1, xBase + baseMetrics.cx + superGap + 1, baseline, tmBase, outPt);
            return false;
        }

        return false;
    }

    static SIZE MeasureMathNode(HDC hdc, const MathNode& node, const TEXTMETRICW& tmBase)
    {
        const NodeMetrics metrics = MeasureMathNodeMetrics(hdc, node, tmBase);
        SIZE size = { metrics.cx, metrics.Height() };
        return size;
    }

    static NodeMetrics MeasureMathNodeMetrics(HDC hdc, const MathNode& node, const TEXTMETRICW& tmBase)
    {
        const int pad = (std::max<int>)(2, tmBase.tmHeight / 8);
        const int radicalW = (std::max<int>)(8, tmBase.tmAveCharWidth);
        const int overlineGap = (std::max<int>)(2, tmBase.tmHeight / 10);
        const int fracGap = (std::max<int>)(3, tmBase.tmHeight / 6);
        const int absPad = (std::max<int>)(3, tmBase.tmHeight / 7);
        const int superGap = (std::max<int>)(2, tmBase.tmAveCharWidth / 3);

        if (node.kind == MathNodeKind::Group)
            return MeasureMathSequenceMetrics(hdc, node.children, tmBase);

        if (node.kind == MathNodeKind::SquareRoot)
        {
            const NodeMetrics childMetrics = MeasureMathSequenceMetrics(hdc, node.SlotNodes(0), tmBase);
            const NodeMetrics indexMetrics = MeasureMathSequenceMetrics(hdc, node.SlotNodes(1), tmBase);
            NodeMetrics metrics = {};
            metrics.cx = radicalW + pad + childMetrics.cx + pad + indexMetrics.cx;
            const int radicalAscent = (std::max)((int)tmBase.tmAscent, childMetrics.ascent) + overlineGap + 1;
            metrics.ascent = (std::max)(radicalAscent, indexMetrics.Height() + overlineGap);
            metrics.descent = (std::max)(childMetrics.descent, (int)tmBase.tmDescent + pad);
            return metrics;
        }

        if (node.kind == MathNodeKind::Fraction)
        {
            const NodeMetrics numMetrics = MeasureMathSequenceMetrics(hdc, node.SlotNodes(0), tmBase);
            const NodeMetrics denMetrics = MeasureMathSequenceMetrics(hdc, node.SlotNodes(1), tmBase);
            NodeMetrics metrics = {};
            metrics.cx = (std::max)(numMetrics.cx, denMetrics.cx) + pad * 2;
            metrics.ascent = numMetrics.Height() + fracGap + 1;
            metrics.descent = denMetrics.Height() + fracGap + 1;
            return metrics;
        }

        if (node.kind == MathNodeKind::Power)
        {
            const NodeMetrics baseMetrics = MeasureMathSequenceMetrics(hdc, node.SlotNodes(0), tmBase);
            const NodeMetrics expMetrics = MeasureMathSequenceMetrics(hdc, node.SlotNodes(1), tmBase);
            NodeMetrics metrics = {};
            metrics.cx = baseMetrics.cx + superGap + expMetrics.cx;
            const int exponentLift = (std::max)((int)tmBase.tmAscent / 2, expMetrics.descent);
            metrics.ascent = (std::max)(baseMetrics.ascent, exponentLift + expMetrics.ascent);
            metrics.descent = (std::max)(baseMetrics.descent, expMetrics.descent - exponentLift / 2);
            return metrics;
        }

        if (node.kind == MathNodeKind::AbsoluteValue)
        {
            const NodeMetrics exprMetrics = MeasureMathSequenceMetrics(hdc, node.SlotNodes(0), tmBase);
            NodeMetrics metrics = {};
            metrics.cx = exprMetrics.cx + absPad * 4;
            metrics.ascent = exprMetrics.ascent + absPad;
            metrics.descent = exprMetrics.descent + absPad;
            return metrics;
        }

        if (node.kind == MathNodeKind::Logarithm)
        {
            SIZE logSize = {};
            GetTextExtentPoint32W(hdc, L"log", 3, &logSize);
            const NodeMetrics baseMetrics = MeasureMathSequenceMetrics(hdc, node.SlotNodes(0), tmBase);
            const NodeMetrics argMetrics = MeasureMathSequenceMetrics(hdc, node.SlotNodes(1), tmBase);
            NodeMetrics metrics = {};
            metrics.cx = logSize.cx + baseMetrics.cx + argMetrics.cx + superGap * 2;
            metrics.ascent = (std::max)((int)tmBase.tmAscent, argMetrics.ascent);
            metrics.descent = (std::max)(argMetrics.descent, (int)tmBase.tmDescent / 2 + baseMetrics.descent);
            return metrics;
        }

        NodeMetrics textMetrics = MeasureDisplayTextMetrics(hdc, node.text, tmBase);
        if (!node.children.empty())
        {
            const NodeMetrics childMetrics = MeasureMathSequenceMetrics(hdc, node.children, tmBase);
            textMetrics.cx += childMetrics.cx;
            textMetrics.ascent = (std::max)(textMetrics.ascent, childMetrics.ascent);
            textMetrics.descent = (std::max)(textMetrics.descent, childMetrics.descent);
        }
        return textMetrics;
    }

    static void DrawMathNodeSequence(HDC hdc, const std::vector<MathNode>& nodes, int x, int baseline, const TEXTMETRICW& tmBase, COLORREF color)
    {
        int cursorX = x;

        for (const auto& node : nodes)
        {
            DrawMathNode(hdc, node, cursorX, baseline, tmBase, color);
            SIZE nodeSize = MeasureMathNode(hdc, node, tmBase);
            cursorX += nodeSize.cx;
        }
    }

    static void DrawMathNode(HDC hdc, const MathNode& node, int x, int baseline, const TEXTMETRICW& tmBase, COLORREF color)
    {
        const int pad = (std::max<int>)(2, tmBase.tmHeight / 8);
        const int radicalW = (std::max<int>)(8, tmBase.tmAveCharWidth);
        const int overlineGap = (std::max<int>)(2, tmBase.tmHeight / 10);
        const int penWidth = 1;
        const int fracGap = (std::max<int>)(3, tmBase.tmHeight / 6);
        const int absPad = (std::max<int>)(3, tmBase.tmHeight / 7);
        const int superGap = (std::max<int>)(2, tmBase.tmAveCharWidth / 3);

        if (node.kind == MathNodeKind::Group)
        {
            DrawMathNodeSequence(hdc, node.children, x, baseline, tmBase, color);
            return;
        }

        if (node.kind == MathNodeKind::SquareRoot)
        {
            const NodeMetrics childMetrics = MeasureMathSequenceMetrics(hdc, node.SlotNodes(0), tmBase);
            const NodeMetrics indexMetrics = MeasureMathSequenceMetrics(hdc, node.SlotNodes(1), tmBase);
            const int radTop = baseline - (std::max)((int)tmBase.tmAscent, childMetrics.ascent) - overlineGap - penWidth;
            const int radBot = baseline + tmBase.tmDescent + pad;
            const int radMid = radBot - (int)((radBot - radTop) * 0.35);
            const int xRadStart = x + indexMetrics.cx;
            const int xValley = xRadStart + radicalW / 3;
            const int xRadPeak = xRadStart + radicalW;
            const int xChild = xRadPeak + pad;
            const int xOverlineEnd = xChild + childMetrics.cx + pad;

            if (indexMetrics.cx > 0)
                DrawMathNodeSequence(hdc, node.SlotNodes(1), x, radTop + indexMetrics.ascent, tmBase, color);

            HPEN hPen = CreatePen(PS_SOLID, penWidth, color);
            HPEN oldPen = (HPEN)SelectObject(hdc, hPen);
            MoveToEx(hdc, xRadStart, radMid - (int)(tmBase.tmHeight * 0.05), NULL);
            LineTo(hdc, xValley, radBot);
            LineTo(hdc, xRadPeak, radTop);
            LineTo(hdc, xOverlineEnd, radTop);
            SelectObject(hdc, oldPen);
            DeleteObject(hPen);

            DrawMathNodeSequence(hdc, node.SlotNodes(0), xChild, baseline, tmBase, color);
            return;
        }

        if (node.kind == MathNodeKind::Fraction)
        {
            const NodeMetrics numMetrics = MeasureMathSequenceMetrics(hdc, node.SlotNodes(0), tmBase);
            const NodeMetrics denMetrics = MeasureMathSequenceMetrics(hdc, node.SlotNodes(1), tmBase);
            const int barWidth = (std::max)(numMetrics.cx, denMetrics.cx) + pad * 2;
            const int barY = baseline;
            const int numX = x + (barWidth - numMetrics.cx) / 2;
            const int denX = x + (barWidth - denMetrics.cx) / 2;
            const int numBaseline = barY - fracGap - numMetrics.descent;
            const int denBaseline = barY + fracGap + denMetrics.ascent;

            DrawMathNodeSequence(hdc, node.SlotNodes(0), numX, numBaseline, tmBase, color);
            DrawMathNodeSequence(hdc, node.SlotNodes(1), denX, denBaseline, tmBase, color);

            HPEN hPen = CreatePen(PS_SOLID, penWidth, color);
            HPEN oldPen = (HPEN)SelectObject(hdc, hPen);
            MoveToEx(hdc, x, barY, NULL);
            LineTo(hdc, x + barWidth, barY);
            SelectObject(hdc, oldPen);
            DeleteObject(hPen);
            return;
        }

        if (node.kind == MathNodeKind::Power)
        {
            const NodeMetrics baseMetrics = MeasureMathSequenceMetrics(hdc, node.SlotNodes(0), tmBase);
            const NodeMetrics expMetrics = MeasureMathSequenceMetrics(hdc, node.SlotNodes(1), tmBase);
            const int exponentBaseline = baseline - (std::max)((int)tmBase.tmAscent / 2, expMetrics.descent);
            DrawMathNodeSequence(hdc, node.SlotNodes(0), x, baseline, tmBase, color);
            DrawMathNodeSequence(hdc, node.SlotNodes(1), x + baseMetrics.cx + superGap, exponentBaseline, tmBase, color);
            return;
        }

        if (node.kind == MathNodeKind::AbsoluteValue)
        {
            SIZE exprSize = MeasureMathNodeSequence(hdc, node.SlotNodes(0), tmBase);
            const int barTop = baseline - tmBase.tmAscent - absPad;
            const int barBot = baseline + tmBase.tmDescent + absPad;
            const int xLeftBar = x + absPad;
            const int xExpr = xLeftBar + absPad;
            const int xRightBar = xExpr + exprSize.cx + absPad;

            HPEN hPen = CreatePen(PS_SOLID, penWidth + 1, color);
            HPEN oldPen = (HPEN)SelectObject(hdc, hPen);
            MoveToEx(hdc, xLeftBar, barTop, NULL);
            LineTo(hdc, xLeftBar, barBot);
            MoveToEx(hdc, xRightBar, barTop, NULL);
            LineTo(hdc, xRightBar, barBot);
            SelectObject(hdc, oldPen);
            DeleteObject(hPen);

            DrawMathNodeSequence(hdc, node.SlotNodes(0), xExpr, baseline, tmBase, color);
            return;
        }

        if (node.kind == MathNodeKind::Logarithm)
        {
            SIZE logSize = {};
            GetTextExtentPoint32W(hdc, L"log", 3, &logSize);
            const NodeMetrics baseMetrics = MeasureMathSequenceMetrics(hdc, node.SlotNodes(0), tmBase);
            const int baseBaseline = baseline + (std::max)((int)tmBase.tmDescent / 2, baseMetrics.ascent / 4);
            SetTextColor(hdc, color);
            TextOutW(hdc, x, baseline, L"log", 3);
            DrawMathNodeSequence(hdc, node.SlotNodes(0), x + logSize.cx + 1, baseBaseline, tmBase, color);
            DrawMathNodeSequence(hdc, node.SlotNodes(1), x + logSize.cx + baseMetrics.cx + superGap + 1, baseline, tmBase, color);
            return;
        }

        const wchar_t* text = node.text.empty() ? L"?" : node.text.c_str();
        const int len = node.text.empty() ? 1 : (int)node.text.size();
        SIZE textSize = {};
        GetTextExtentPoint32W(hdc, text, len, &textSize);
        SetTextColor(hdc, color);
        TextOutW(hdc, x, baseline, text, len);
        if (!node.children.empty())
            DrawMathNodeSequence(hdc, node.children, x + textSize.cx, baseline, tmBase, color);
    }

    static bool HitTestMathNodeSequence(HDC hdc, const std::vector<MathNode>& nodes, int x, int baseline, const TEXTMETRICW& tmBase, POINT ptMouse, const std::vector<size_t>& prefix, std::vector<size_t>& outPath)
    {
        int cursorX = x;
        for (size_t nodeIndex = 0; nodeIndex < nodes.size(); ++nodeIndex)
        {
            if (HitTestMathNode(hdc, nodes[nodeIndex], nodeIndex, cursorX, baseline, tmBase, ptMouse, prefix, outPath))
                return true;
            SIZE nodeSize = MeasureMathNode(hdc, nodes[nodeIndex], tmBase);
            cursorX += nodeSize.cx;
        }
        return false;
    }

    static bool HitTestMathNode(HDC hdc, const MathNode& node, size_t nodeIndex, int x, int baseline, const TEXTMETRICW& tmBase, POINT ptMouse, const std::vector<size_t>& prefix, std::vector<size_t>& outPath)
    {
        const int pad = (std::max<int>)(2, tmBase.tmHeight / 8);
        const int radicalW = (std::max<int>)(8, tmBase.tmAveCharWidth);
        const int overlineGap = (std::max<int>)(2, tmBase.tmHeight / 10);
        const int fracGap = (std::max<int>)(3, tmBase.tmHeight / 6);
        const int absPad = (std::max<int>)(3, tmBase.tmHeight / 7);
        const int superGap = (std::max<int>)(2, tmBase.tmAveCharWidth / 3);

        std::vector<size_t> nodePrefix = prefix;

        if (node.kind == MathNodeKind::Group)
            return HitTestMathNodeSequence(hdc, node.children, x, baseline, tmBase, ptMouse, prefix, outPath);

        if (node.kind == MathNodeKind::SquareRoot)
        {
            const auto& indexNodes = node.SlotNodes(1);
            const auto& radicandNodes = node.SlotNodes(0);
            const NodeMetrics childMetrics = MeasureMathSequenceMetrics(hdc, radicandNodes, tmBase);
            const NodeMetrics indexMetrics = MeasureMathSequenceMetrics(hdc, indexNodes, tmBase);
            const int radTop = baseline - (std::max)((int)tmBase.tmAscent, childMetrics.ascent) - overlineGap - 1;
            const int radBot = baseline + tmBase.tmDescent + pad;
            const int xRadStart = x + indexMetrics.cx;
            const int xRadPeak = xRadStart + radicalW;
            const int xChild = xRadPeak + pad;
            const int xOverlineEnd = xChild + childMetrics.cx + pad;
            if (indexMetrics.cx > 0)
            {
                RECT rcIndex = { x, radTop - 4, xRadStart - 1, baseline + tmBase.tmDescent + 4 };
                if (PtInRect(&rcIndex, ptMouse))
                {
                    nodePrefix.push_back(nodeIndex);
                    nodePrefix.push_back(1);
                    if (!HitTestMathNodeSequence(hdc, indexNodes, x, radTop + indexMetrics.ascent, tmBase, ptMouse, nodePrefix, outPath))
                        SetSequenceTailPath(indexNodes, nodePrefix, outPath);
                    return true;
                }
            }

            RECT rcRadicand = { xRadStart, radTop - 4, xOverlineEnd + 4, radBot + 4 };
            if (PtInRect(&rcRadicand, ptMouse))
            {
                nodePrefix.push_back(nodeIndex);
                nodePrefix.push_back(0);
                if (!HitTestMathNodeSequence(hdc, radicandNodes, xChild, baseline, tmBase, ptMouse, nodePrefix, outPath))
                    SetSequenceTailPath(radicandNodes, nodePrefix, outPath);
                return true;
            }
            return false;
        }

        if (node.kind == MathNodeKind::Fraction)
        {
            const auto& numeratorNodes = node.SlotNodes(0);
            const auto& denominatorNodes = node.SlotNodes(1);
            const NodeMetrics numMetrics = MeasureMathSequenceMetrics(hdc, numeratorNodes, tmBase);
            const NodeMetrics denMetrics = MeasureMathSequenceMetrics(hdc, denominatorNodes, tmBase);
            const int barWidth = (std::max)(numMetrics.cx, denMetrics.cx) + pad * 2;
            const int numX = x + (barWidth - numMetrics.cx) / 2;
            const int denX = x + (barWidth - denMetrics.cx) / 2;
            const int numBaseline = baseline - fracGap - numMetrics.descent;
            const int denBaseline = baseline + fracGap + denMetrics.ascent;
            RECT rcNum = { x - 4, numBaseline - numMetrics.ascent - 4, x + barWidth + 4, numBaseline + numMetrics.descent + 2 };
            if (PtInRect(&rcNum, ptMouse))
            {
                nodePrefix.push_back(nodeIndex);
                nodePrefix.push_back(0);
                if (!HitTestMathNodeSequence(hdc, numeratorNodes, numX, numBaseline, tmBase, ptMouse, nodePrefix, outPath))
                    SetSequenceTailPath(numeratorNodes, nodePrefix, outPath);
                return true;
            }
            RECT rcDen = { x - 4, denBaseline - denMetrics.ascent - 2, x + barWidth + 4, denBaseline + denMetrics.descent + 4 };
            if (PtInRect(&rcDen, ptMouse))
            {
                nodePrefix.push_back(nodeIndex);
                nodePrefix.push_back(1);
                if (!HitTestMathNodeSequence(hdc, denominatorNodes, denX, denBaseline, tmBase, ptMouse, nodePrefix, outPath))
                    SetSequenceTailPath(denominatorNodes, nodePrefix, outPath);
                return true;
            }
            return false;
        }

        if (node.kind == MathNodeKind::Power)
        {
            const auto& baseNodes = node.SlotNodes(0);
            const auto& exponentNodes = node.SlotNodes(1);
            const NodeMetrics baseMetrics = MeasureMathSequenceMetrics(hdc, baseNodes, tmBase);
            const NodeMetrics expMetrics = MeasureMathSequenceMetrics(hdc, exponentNodes, tmBase);
            const int exponentBaseline = baseline - (std::max)((int)tmBase.tmAscent / 2, expMetrics.descent);
            RECT rcBase = { x - 4, baseline - baseMetrics.ascent - 4, x + baseMetrics.cx + 6, baseline + baseMetrics.descent + 4 };
            if (PtInRect(&rcBase, ptMouse))
            {
                nodePrefix.push_back(nodeIndex);
                nodePrefix.push_back(0);
                if (!HitTestMathNodeSequence(hdc, baseNodes, x, baseline, tmBase, ptMouse, nodePrefix, outPath))
                    SetSequenceTailPath(baseNodes, nodePrefix, outPath);
                return true;
            }
            RECT rcExponent = { x + baseMetrics.cx + superGap - 4, exponentBaseline - expMetrics.ascent - 4, x + baseMetrics.cx + superGap + expMetrics.cx + 6, exponentBaseline + expMetrics.descent + 2 };
            if (PtInRect(&rcExponent, ptMouse))
            {
                nodePrefix.push_back(nodeIndex);
                nodePrefix.push_back(1);
                if (!HitTestMathNodeSequence(hdc, exponentNodes, x + baseMetrics.cx + superGap, exponentBaseline, tmBase, ptMouse, nodePrefix, outPath))
                    SetSequenceTailPath(exponentNodes, nodePrefix, outPath);
                return true;
            }
            return false;
        }

        if (node.kind == MathNodeKind::AbsoluteValue)
        {
            const auto& exprNodes = node.SlotNodes(0);
            SIZE exprSize = MeasureMathNodeSequence(hdc, exprNodes, tmBase);
            RECT rcExpr = { x - 4, baseline - tmBase.tmAscent - absPad - 4, x + exprSize.cx + absPad * 4 + 4, baseline + tmBase.tmDescent + absPad + 4 };
            if (PtInRect(&rcExpr, ptMouse))
            {
                nodePrefix.push_back(nodeIndex);
                nodePrefix.push_back(0);
                const int xExpr = x + absPad * 2;
                if (!HitTestMathNodeSequence(hdc, exprNodes, xExpr, baseline, tmBase, ptMouse, nodePrefix, outPath))
                    SetSequenceTailPath(exprNodes, nodePrefix, outPath);
                return true;
            }
            return false;
        }

        if (node.kind == MathNodeKind::Logarithm)
        {
            const auto& baseNodes = node.SlotNodes(0);
            const auto& argNodes = node.SlotNodes(1);
            SIZE logSize = {};
            GetTextExtentPoint32W(hdc, L"log", 3, &logSize);
            const NodeMetrics baseMetrics = MeasureMathSequenceMetrics(hdc, baseNodes, tmBase);
            const NodeMetrics argMetrics = MeasureMathSequenceMetrics(hdc, argNodes, tmBase);
            const int xBase = x + logSize.cx + 1;
            const int baseBaseline = baseline + (std::max)((int)tmBase.tmDescent / 2, baseMetrics.ascent / 4);
            RECT rcBase = { xBase - 3, baseBaseline - baseMetrics.ascent - 2, xBase + baseMetrics.cx + 4, baseBaseline + baseMetrics.descent + 4 };
            if (PtInRect(&rcBase, ptMouse))
            {
                nodePrefix.push_back(nodeIndex);
                nodePrefix.push_back(0);
                if (!HitTestMathNodeSequence(hdc, baseNodes, xBase, baseBaseline, tmBase, ptMouse, nodePrefix, outPath))
                    SetSequenceTailPath(baseNodes, nodePrefix, outPath);
                return true;
            }
            const int xArg = xBase + baseMetrics.cx + superGap + 1;
            RECT rcArg = { xArg - 4, baseline - argMetrics.ascent - 4, xArg + argMetrics.cx + 6, baseline + argMetrics.descent + 4 };
            if (PtInRect(&rcArg, ptMouse))
            {
                nodePrefix.push_back(nodeIndex);
                nodePrefix.push_back(1);
                if (!HitTestMathNodeSequence(hdc, argNodes, xArg, baseline, tmBase, ptMouse, nodePrefix, outPath))
                    SetSequenceTailPath(argNodes, nodePrefix, outPath);
                return true;
            }
            return false;
        }

        SIZE textSize = {};
        const wchar_t* text = node.text.empty() ? L"?" : node.text.c_str();
        const int len = node.text.empty() ? 1 : (int)node.text.size();
        GetTextExtentPoint32W(hdc, text, len, &textSize);
        RECT rcText = { x - 2, baseline - tmBase.tmAscent - 4, x + textSize.cx + 2, baseline + tmBase.tmDescent + 4 };
        if (PtInRect(&rcText, ptMouse))
        {
            outPath = prefix;
            outPath.push_back(nodeIndex);
            return true;
        }

        if (!node.children.empty())
            return HitTestMathNodeSequence(hdc, node.children, x + textSize.cx, baseline, tmBase, ptMouse, prefix, outPath);
        return false;
    }

    static COLORREF GetRichEditBkColor(HWND hEdit)
    {
        COLORREF prev = (COLORREF)SendMessage(hEdit, EM_SETBKGNDCOLOR, 0, (LPARAM)GetSysColor(COLOR_WINDOW));
        SendMessage(hEdit, EM_SETBKGNDCOLOR, 0, (LPARAM)prev);
        return prev;
    }

    static double ComputeRenderScale(HWND hEdit, HDC hdc, const MathObject& obj, HFONT unzoomedBaseFont)
    {
        if (obj.barLen <= 0 || !unzoomedBaseFont) return 1.0;
        POINT p0 = {}, p1 = {};
        if (MathRenderer::TryGetCharPos(hEdit, obj.barStart, p0) && MathRenderer::TryGetCharPos(hEdit, obj.barStart + 1, p1))
        {
            if (p0.y == p1.y && p1.x > p0.x)
            {
                HFONT old = (HFONT)SelectObject(hdc, unzoomedBaseFont);
                wchar_t buf[16] = {0};
                TEXTRANGEW tr = { {obj.barStart, obj.barStart + 1}, buf };
                SendMessage(hEdit, EM_GETTEXTRANGE, 0, (LPARAM)&tr);
                SIZE one = {};
                GetTextExtentPoint32W(hdc, buf, 1, &one);
                SelectObject(hdc, old);
                if (one.cx > 0) return (double)(p1.x - p0.x) / (double)one.cx;
            }
        }
        return 1.0;
    }

    static HFONT CreateScaledFont(HFONT baseFont, double scale, int percent)
    {
        LOGFONTW lf = {};
        if (!baseFont || GetObjectW(baseFont, sizeof(lf), &lf) != sizeof(lf))
        {
            return CreateFontW(-11, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
                OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH | FF_SWISS, L"Arial");
        }
        if (!(scale > 0.0)) scale = 1.0;
        const int sign = (lf.lfHeight < 0) ? -1 : 1;
        const int absH = (lf.lfHeight < 0) ? -lf.lfHeight : lf.lfHeight;
        const double scaled = (double)absH * scale * ((double)percent / 100.0);
        lf.lfHeight = sign * (std::max)(1, (int)(scaled + 0.5));
        return CreateFontIndirectW(&lf);
    }
}

COLORREF MathRenderer::GetDefaultTextColor(HWND hEdit)
{
    CHARFORMAT2W cf = {};
    cf.cbSize = sizeof(cf);
    SendMessage(hEdit, EM_GETCHARFORMAT, (WPARAM)SCF_DEFAULT, (LPARAM)&cf);
    if ((cf.dwMask & CFM_COLOR) == 0 || (cf.dwEffects & CFE_AUTOCOLOR) != 0)
        return GetSysColor(COLOR_WINDOWTEXT);
    return cf.crTextColor;
}

COLORREF MathRenderer::GetActiveColor(HWND hEdit)
{
    const COLORREF text = GetDefaultTextColor(hEdit);
    const int brightness = (GetRValue(text) * 299 + GetGValue(text) * 587 + GetBValue(text) * 114) / 1000;
    return (brightness > 128) ? RGB(100, 180, 255) : RGB(0, 102, 204);
}

bool MathRenderer::TryGetCharPos(HWND hEdit, LONG charIndex, POINT& outPt)
{
    POINTL ptL = {};
    if (SendMessage(hEdit, EM_POSFROMCHAR, (WPARAM)&ptL, (LPARAM)charIndex) != -1)
    {
        outPt.x = (int)ptL.x; outPt.y = (int)ptL.y;
        return true;
    }
    LRESULT xy = SendMessage(hEdit, EM_POSFROMCHAR, (WPARAM)charIndex, 0);
    if (xy != -1)
    {
        outPt.x = (short)LOWORD(xy); outPt.y = (short)HIWORD(xy);
        return true;
    }
    return false;
}

bool MathRenderer::TryGetObjectBounds(HWND hEdit, HDC hdc, const MathObject& obj, RECT& outRect)
{
    if (obj.barLen <= 0)
        return false;

    const std::wstring& part1 = obj.SlotText(1);
    const std::wstring& part2 = obj.SlotText(2);
    const std::wstring& part3 = obj.SlotText(3);
    const std::wstring& part4 = obj.SlotText(4);
    POINT ptStart = {}, ptEnd = {};
    if (!TryGetCharPos(hEdit, obj.barStart, ptStart)) return false;
    if (!TryGetCharPos(hEdit, obj.barStart + std::max<LONG>(0, obj.barLen - 1), ptEnd)) return false;

    HFONT baseFont = (HFONT)SendMessage(hEdit, WM_GETFONT, 0, 0);
    if (!baseFont) baseFont = (HFONT)GetStockObject(DEFAULT_GUI_FONT);

    double renderScale = ComputeRenderScale(hEdit, hdc, obj, baseFont);
    if (obj.type == MathType::Summation || obj.type == MathType::Integral
        || obj.type == MathType::SystemOfEquations || obj.type == MathType::Power
        || obj.type == MathType::Logarithm || obj.type == MathType::Sum
        || obj.type == MathType::Matrix || obj.type == MathType::Determinant)
        renderScale *= 0.5;

    HFONT renderBaseFont = CreateScaledFont(baseFont, renderScale, 100);
    HFONT limitFont = CreateScaledFont(baseFont, renderScale, 70);
    const int saved = SaveDC(hdc);
    HFONT oldFont = (HFONT)SelectObject(hdc, renderBaseFont);
    TEXTMETRICW tmBase = {};
    GetTextMetricsW(hdc, &tmBase);

    const int barWidth = (ptEnd.x - ptStart.x) + (int)tmBase.tmAveCharWidth;
    const int xCenter = ptStart.x + (barWidth / 2);
    const int yMid = ptStart.y + tmBase.tmAscent;

    RECT rc = { ptStart.x, ptStart.y, ptEnd.x + tmBase.tmAveCharWidth, ptStart.y + tmBase.tmHeight };
    auto IncludeRect = [&](int left, int top, int right, int bottom) {
        if (left < rc.left) rc.left = left;
        if (top < rc.top) rc.top = top;
        if (right > rc.right) rc.right = right;
        if (bottom > rc.bottom) rc.bottom = bottom;
    };

    auto MeasureSlotMetrics = [&](int partIndex) -> NodeMetrics {
        const size_t slotIndex = (size_t)(partIndex - 1);
        if (slotIndex < obj.slots.size() && !obj.slots[slotIndex].children.empty())
            return MeasureMathSequenceMetrics(hdc, obj.slots[slotIndex].children, tmBase);
        return MeasureDisplayTextMetrics(hdc, obj.SlotText(partIndex), tmBase);
    };

    SIZE resultSz = {};
    if (!obj.resultText.empty())
        resultSz = MeasureDisplayText(hdc, obj.resultText);

    if (obj.type == MathType::Fraction)
    {
        HFONT fracFont = CreateScaledFont(baseFont, renderScale, 105);
        SelectObject(hdc, fracFont);
        TEXTMETRICW tmL = {};
        GetTextMetricsW(hdc, &tmL);
        SelectObject(hdc, renderBaseFont);

        const NodeMetrics numMetrics = MeasureSlotMetrics(1);
        const NodeMetrics denMetrics = MeasureSlotMetrics(2);
        const int gap = 4;
        const int barY = ptStart.y + tmBase.tmHeight / 2;
        const int numLeft = xCenter - numMetrics.cx / 2;
        const int numBaseline = barY - gap - tmL.tmDescent;
        const int denLeft = xCenter - denMetrics.cx / 2;
        const int denBaseline = barY + gap + tmL.tmAscent;
        IncludeRect(ptStart.x - 2, barY - 2, ptStart.x + barWidth + 2, barY + 2);
        IncludeRect(numLeft, numBaseline - numMetrics.ascent, numLeft + numMetrics.cx, numBaseline + numMetrics.descent);
        IncludeRect(denLeft, denBaseline - denMetrics.ascent, denLeft + denMetrics.cx, denBaseline + denMetrics.descent);
        if (!obj.resultText.empty())
        {
            const int resultBaseline = barY + (tmBase.tmAscent - tmBase.tmDescent) / 2;
            IncludeRect(ptStart.x + barWidth + 4, resultBaseline - tmBase.tmAscent, ptStart.x + barWidth + 4 + resultSz.cx, resultBaseline + tmBase.tmDescent);
        }
        DeleteObject(fracFont);
    }
    else if (obj.type == MathType::Summation || obj.type == MathType::Product || obj.type == MathType::Integral)
    {
        SelectObject(hdc, limitFont);
        TEXTMETRICW tmL = {};
        GetTextMetricsW(hdc, &tmL);
        SelectObject(hdc, renderBaseFont);
        const NodeMetrics upper = MeasureSlotMetrics(1);
        const NodeMetrics lower = MeasureSlotMetrics(2);
        const NodeMetrics expr = MeasureSlotMetrics(3);
        IncludeRect(xCenter - upper.cx / 2, yMid - tmBase.tmAscent - 2 - upper.ascent, xCenter + upper.cx / 2, yMid - tmBase.tmAscent - 2 + upper.descent);
        IncludeRect(xCenter - lower.cx / 2, yMid + tmBase.tmDescent + tmL.tmAscent + 2 - lower.ascent, xCenter + lower.cx / 2, yMid + tmBase.tmDescent + tmL.tmAscent + 2 + lower.descent);
        const int exprX = ptEnd.x + (obj.type == MathType::Integral ? 6 : 4);
        IncludeRect(exprX, yMid - expr.ascent, exprX + expr.cx, yMid + expr.descent);
        if (!obj.resultText.empty())
            IncludeRect(exprX + expr.cx + 4, yMid - tmBase.tmAscent, exprX + expr.cx + 4 + resultSz.cx, yMid + tmBase.tmDescent);
    }
    else if (obj.type == MathType::Sum)
    {
        const NodeMetrics expr = MeasureSlotMetrics(1);
        const int exprX = ptStart.x + 4;
        IncludeRect(exprX, yMid - expr.ascent, exprX + expr.cx, yMid + expr.descent);
        if (!obj.resultText.empty())
            IncludeRect(exprX + expr.cx + 4, yMid - tmBase.tmAscent, exprX + expr.cx + 4 + resultSz.cx, yMid + tmBase.tmDescent);
    }
    else if (obj.type == MathType::SystemOfEquations)
    {
        const int braceW = (std::max)(10, (int)(tmBase.tmAveCharWidth * 1.2));
        const int lineH = tmBase.tmHeight + 4;
        const int eqCount = SlotHasVisibleContent(obj, 3) ? 3 : 2;
        const int totalH = lineH * eqCount;
        const int yTop = yMid - totalH / 2 + tmBase.tmAscent;
        const int blockTop = yTop - tmBase.tmAscent - 2;
        const int blockBot = blockTop + totalH + 4;
        const int eqX = ptStart.x + braceW + 6;
        IncludeRect(ptStart.x, blockTop, eqX + 10, blockBot);
        const NodeMetrics eq1 = MeasureSlotMetrics(1);
        const NodeMetrics eq2 = MeasureSlotMetrics(2);
        IncludeRect(eqX, yTop - eq1.ascent, eqX + eq1.cx, yTop + eq1.descent);
        IncludeRect(eqX, yTop + lineH - eq2.ascent, eqX + eq2.cx, yTop + lineH + eq2.descent);
        if (eqCount >= 3)
        {
            const NodeMetrics eq3 = MeasureSlotMetrics(3);
            IncludeRect(eqX, yTop + lineH * 2 - eq3.ascent, eqX + eq3.cx, yTop + lineH * 2 + eq3.descent);
        }
        if (!obj.resultText.empty())
        {
            const int eq1Width = eq1.cx;
            const int eq2Width = eq2.cx;
            const int eq3Width = eqCount >= 3 ? MeasureSlotMetrics(3).cx : 0;
            const int resultX = eqX + (std::max)((std::max)(eq1Width, eq2Width), eq3Width) + 10;
            IncludeRect(resultX, yMid - tmBase.tmAscent, resultX + resultSz.cx, yMid + tmBase.tmDescent);
        }
    }
    else if (obj.type == MathType::Matrix || obj.type == MathType::Determinant)
    {
        const int cellPadX = (std::max)(6, (int)tmBase.tmAveCharWidth / 2);
        const int cellPadY = (std::max)(4, (int)tmBase.tmHeight / 6);
        const int colGap = (std::max)(12, (int)tmBase.tmAveCharWidth);
        const int rowGap = (std::max)(8, (int)tmBase.tmHeight / 4);
        const int bracketInset = (std::max)(8, (int)(tmBase.tmAveCharWidth * 0.8));
        const NodeMetrics cell1 = MeasureSlotMetrics(1);
        const NodeMetrics cell2 = MeasureSlotMetrics(2);
        const NodeMetrics cell3 = MeasureSlotMetrics(3);
        const NodeMetrics cell4 = MeasureSlotMetrics(4);
        const int col0 = (std::max)(cell1.cx, cell3.cx);
        const int col1 = (std::max)(cell2.cx, cell4.cx);
        const int row0 = (std::max)(cell1.Height(), cell2.Height());
        const int row1 = (std::max)(cell3.Height(), cell4.Height());
        const int contentX = ptStart.x + bracketInset + 10;
        const int totalH = row0 + row1 + rowGap + cellPadY * 4;
        const int topY = yMid - totalH / 2;
        const int rightStrokeX = contentX + col0 + col1 + colGap + cellPadX * 4 + 8;
        IncludeRect(ptStart.x + bracketInset, topY, rightStrokeX, topY + totalH);
        if (!obj.resultText.empty())
            IncludeRect(rightStrokeX + 10, yMid - tmBase.tmAscent, rightStrokeX + 10 + resultSz.cx, yMid + tmBase.tmDescent);
    }
    else if (obj.type == MathType::SquareRoot)
    {
        const NodeMetrics radicand = MeasureSlotMetrics(1);
        const NodeMetrics index = MeasureSlotMetrics(2);
        const bool showIndex = SlotHasVisibleContent(obj, 2);
        const int pad = (std::max)(2, (int)(tmBase.tmHeight / 8));
        const int overlineGap = (std::max)(2, (int)(tmBase.tmHeight / 10));
        const int radicalW = (int)(tmBase.tmAveCharWidth * 1.0);
        const int indexGap = (std::max)(2, (int)(tmBase.tmAveCharWidth / 2));
        const int radTop = yMid - (std::max)((int)tmBase.tmAscent, radicand.ascent) - overlineGap - 1;
        const int radBot = yMid + tmBase.tmDescent + pad;
        const int xRadStart = ptStart.x + (showIndex ? index.cx + indexGap : 0);
        const int xExprStart = xRadStart + radicalW + pad;
        const int xOverlineEnd = xExprStart + radicand.cx + pad;
        IncludeRect(xRadStart, radTop, xOverlineEnd + 2, radBot + 2);
        IncludeRect(xExprStart, yMid - radicand.ascent, xExprStart + radicand.cx, yMid + radicand.descent);
        if (showIndex)
            IncludeRect(xRadStart - indexGap - index.cx, radTop, xRadStart - indexGap, radTop + index.Height());
        if (!obj.resultText.empty())
            IncludeRect(xOverlineEnd + 6, yMid - tmBase.tmAscent, xOverlineEnd + 6 + resultSz.cx, yMid + tmBase.tmDescent);
    }
    else if (obj.type == MathType::AbsoluteValue)
    {
        const NodeMetrics expr = MeasureSlotMetrics(1);
        const int pad = (std::max)(2, (int)(tmBase.tmHeight / 6));
        const int barExtend = (std::max)(3, (int)(tmBase.tmHeight / 5));
        const int penWidth = (std::max)(2, (int)(1.8 * renderScale));
        const int xLeftBar = ptStart.x + pad;
        const int xExprStart = xLeftBar + pad + penWidth;
        const int xRightBar = xExprStart + expr.cx + pad;
        const int barTop = yMid - expr.ascent - barExtend;
        const int barBot = yMid + expr.descent + barExtend;
        IncludeRect(xLeftBar, barTop, xRightBar + penWidth + pad, barBot);
        if (!obj.resultText.empty())
            IncludeRect(xRightBar + penWidth + pad + 4, yMid - tmBase.tmAscent, xRightBar + penWidth + pad + 4 + resultSz.cx, yMid + tmBase.tmDescent);
    }
    else if (obj.type == MathType::Power)
    {
        const NodeMetrics base = MeasureSlotMetrics(1);
        const NodeMetrics exponent = MeasureSlotMetrics(2);
        const int superGap = (std::max)(2, (int)tmBase.tmAveCharWidth / 3);
        const int xBase = ptStart.x + 2;
        const int exponentBaseline = yMid - (std::max)((int)tmBase.tmAscent / 2, exponent.descent);
        IncludeRect(xBase, yMid - base.ascent, xBase + base.cx, yMid + base.descent);
        IncludeRect(xBase + base.cx + superGap, exponentBaseline - exponent.ascent, xBase + base.cx + superGap + exponent.cx, exponentBaseline + exponent.descent);
        if (!obj.resultText.empty())
            IncludeRect(xBase + base.cx + superGap + exponent.cx + 8, yMid - tmBase.tmAscent, xBase + base.cx + superGap + exponent.cx + 8 + resultSz.cx, yMid + tmBase.tmDescent);
    }
    else if (obj.type == MathType::Logarithm)
    {
        const NodeMetrics base = MeasureSlotMetrics(1);
        const NodeMetrics arg = MeasureSlotMetrics(2);
        SIZE logSz = {};
        GetTextExtentPoint32W(hdc, L"log", 3, &logSz);
        const int xLog = ptStart.x + 2;
        const int xBase = xLog + logSz.cx + 1;
        const int baseBaseline = yMid + (std::max)((int)tmBase.tmDescent / 2, base.ascent / 4);
        const int xArg = xBase + base.cx + 2;
        IncludeRect(xLog, yMid - tmBase.tmAscent, xLog + logSz.cx, yMid + tmBase.tmDescent);
        IncludeRect(xBase, baseBaseline - base.ascent, xBase + base.cx, baseBaseline + base.descent);
        IncludeRect(xArg, yMid - arg.ascent, xArg + arg.cx, yMid + arg.descent);
        if (!obj.resultText.empty())
            IncludeRect(xArg + arg.cx + 6, yMid - tmBase.tmAscent, xArg + arg.cx + 6 + resultSz.cx, yMid + tmBase.tmDescent);
    }

    SelectObject(hdc, oldFont);
    RestoreDC(hdc, saved);
    DeleteObject(renderBaseFont);
    DeleteObject(limitFont);
    outRect = rc;
    return true;
}

bool MathRenderer::TryGetActiveCaretPoint(HWND hEdit, HDC hdc, const MathObject& obj, size_t objIndex, const MathTypingState& state, POINT& outPt)
{
    if (!(state.active && state.objectIndex == objIndex))
        return false;
    if (state.activePart <= 0)
        return false;

    const std::wstring& part1 = obj.SlotText(1);
    const std::wstring& part2 = obj.SlotText(2);
    const std::wstring& part3 = obj.SlotText(3);
    const std::wstring& part4 = obj.SlotText(4);

    POINT ptStart = {}, ptEnd = {};
    if (!TryGetCharPos(hEdit, obj.barStart, ptStart))
        return false;
    if (!TryGetCharPos(hEdit, obj.barStart + std::max<LONG>(0, obj.barLen - 1), ptEnd))
        return false;

    HFONT baseFont = (HFONT)SendMessage(hEdit, WM_GETFONT, 0, 0);
    if (!baseFont)
        baseFont = (HFONT)GetStockObject(DEFAULT_GUI_FONT);

    double renderScale = ComputeRenderScale(hEdit, hdc, obj, baseFont);
    if (obj.type == MathType::Summation || obj.type == MathType::Integral
        || obj.type == MathType::SystemOfEquations || obj.type == MathType::Power
        || obj.type == MathType::Logarithm || obj.type == MathType::Sum
        || obj.type == MathType::Matrix || obj.type == MathType::Determinant)
        renderScale *= 0.5;

    HFONT renderBaseFont = CreateScaledFont(baseFont, renderScale, 100);
    HFONT limitFont = CreateScaledFont(baseFont, renderScale, 70);

    const int saved = SaveDC(hdc);
    HFONT oldFont = (HFONT)SelectObject(hdc, renderBaseFont);
    TEXTMETRICW tmBase = {};
    GetTextMetricsW(hdc, &tmBase);

    const int barWidth = (ptEnd.x - ptStart.x) + (int)tmBase.tmAveCharWidth;
    const int xCenter = ptStart.x + (barWidth / 2);
    const int yMid = ptStart.y + tmBase.tmAscent;

    auto setSequenceCaret = [&](const std::vector<MathNode>& nodes, int baseline, int x, const std::vector<size_t>& path) {
        POINT caretPt = { x + MeasureMathNodeSequence(hdc, nodes, tmBase).cx, baseline };
        if (!path.empty())
            TryGetSequenceCaret(hdc, nodes, path, 0, x, baseline, tmBase, caretPt);
        outPt = caretPt;
        return true;
    };

    auto setLeftTextCaret = [&](const std::wstring& text, int baseline, int x) {
        outPt.x = x + MeasureDisplayText(hdc, text).cx;
        outPt.y = baseline;
        return true;
    };

    auto setCenteredTextCaret = [&](const std::wstring& text, int baseline, int centerX) {
        outPt.x = centerX + MeasureDisplayText(hdc, text).cx / 2;
        outPt.y = baseline;
        return true;
    };

    bool found = false;
    switch (obj.type)
    {
    case MathType::Fraction:
    {
        HFONT fracFont = CreateScaledFont(baseFont, renderScale, 105);
        SelectObject(hdc, fracFont);
        TEXTMETRICW tmL = {};
        GetTextMetricsW(hdc, &tmL);
        SelectObject(hdc, renderBaseFont);

        const int gap = 4;
        const int barY = ptStart.y + tmBase.tmHeight / 2;
        const int caretBaseline = (state.activePart == 1) ? (barY - gap - tmL.tmDescent) : (barY + gap + tmL.tmAscent);
        const size_t slotIndex = (size_t)(state.activePart - 1);
        if (slotIndex < obj.slots.size() && !obj.slots[slotIndex].children.empty())
        {
            const SIZE caretSlotSize = MeasureMathNodeSequence(hdc, obj.slots[slotIndex].children, tmBase);
            const int caretX = xCenter - caretSlotSize.cx / 2;
            found = setSequenceCaret(obj.slots[slotIndex].children, caretBaseline, caretX, state.activeNodePath);
        }
        else
        {
            found = setCenteredTextCaret(obj.SlotText(state.activePart), caretBaseline, xCenter);
        }

        DeleteObject(fracFont);
        break;
    }

    case MathType::Summation:
    case MathType::Product:
    {
        SelectObject(hdc, limitFont);
        TEXTMETRICW tmL = {};
        GetTextMetricsW(hdc, &tmL);
        const int upperBaseline = yMid - tmBase.tmAscent - 2;
        const int lowerBaseline = yMid + tmBase.tmDescent + tmL.tmAscent + 2;
        if (state.activePart == 1)
        {
            if (!obj.slots.empty() && SequenceHasVisibleContent(obj.slots[0].children))
            {
                const SIZE upperSz = MeasureMathNodeSequence(hdc, obj.slots[0].children, tmBase);
                found = setSequenceCaret(obj.slots[0].children, upperBaseline, xCenter - upperSz.cx / 2, state.activeNodePath);
            }
            else
            {
                found = setCenteredTextCaret(part1, upperBaseline, xCenter);
            }
        }
        else if (state.activePart == 2)
        {
            if (obj.slots.size() > 1 && SequenceHasVisibleContent(obj.slots[1].children))
            {
                const SIZE lowerSz = MeasureMathNodeSequence(hdc, obj.slots[1].children, tmBase);
                found = setSequenceCaret(obj.slots[1].children, lowerBaseline, xCenter - lowerSz.cx / 2, state.activeNodePath);
            }
            else
            {
                found = setCenteredTextCaret(part2, lowerBaseline, xCenter);
            }
        }
        else if (state.activePart == 3)
        {
            SelectObject(hdc, renderBaseFont);
            if (obj.slots.size() > 2 && !obj.slots[2].children.empty())
                found = setSequenceCaret(obj.slots[2].children, yMid, ptEnd.x + 4, state.activeNodePath);
            else
                found = setLeftTextCaret(part3, yMid, ptEnd.x + 4);
        }
        break;
    }

    case MathType::Sum:
        if (!obj.slots.empty() && !obj.slots[0].children.empty())
            found = setSequenceCaret(obj.slots[0].children, yMid, ptStart.x + 4, state.activeNodePath);
        else
            found = setLeftTextCaret(part1, yMid, ptStart.x + 4);
        break;

    case MathType::Integral:
    {
        SelectObject(hdc, limitFont);
        const int upperBaseline = yMid - tmBase.tmAscent + (int)(tmBase.tmAscent * 0.2);
        const int lowerBaseline = yMid + tmBase.tmDescent + 2;
        if (state.activePart == 1)
        {
            if (!obj.slots.empty() && SequenceHasVisibleContent(obj.slots[0].children))
                found = setSequenceCaret(obj.slots[0].children, upperBaseline, ptEnd.x - 2, state.activeNodePath);
            else
                found = setLeftTextCaret(part1, upperBaseline, ptEnd.x - 2);
        }
        else if (state.activePart == 2)
        {
            if (obj.slots.size() > 1 && SequenceHasVisibleContent(obj.slots[1].children))
                found = setSequenceCaret(obj.slots[1].children, lowerBaseline, ptEnd.x - 8, state.activeNodePath);
            else
                found = setLeftTextCaret(part2, lowerBaseline, ptEnd.x - 8);
        }
        else if (state.activePart == 3)
        {
            SelectObject(hdc, renderBaseFont);
            if (obj.slots.size() > 2 && !obj.slots[2].children.empty())
                found = setSequenceCaret(obj.slots[2].children, yMid, ptEnd.x + 6, state.activeNodePath);
            else
                found = setLeftTextCaret(part3, yMid, ptEnd.x + 6);
        }
        break;
    }

    case MathType::SystemOfEquations:
    {
        SelectObject(hdc, renderBaseFont);
        TEXTMETRICW tmEq = {};
        GetTextMetricsW(hdc, &tmEq);
        const int lineH = tmEq.tmHeight + 4;
        int eqCount = 3;
        if (!(obj.slots.size() > 2 && SequenceHasVisibleContent(obj.slots[2].children)) && part3.empty() && state.activePart != 3)
            eqCount = 2;
        const int totalH = lineH * eqCount;
        const int yTop = yMid - totalH / 2 + tmEq.tmAscent;
        int braceW = (int)(tmBase.tmAveCharWidth * 1.2);
        if (braceW < 10) braceW = 10;
        const int slotX = ptStart.x + braceW + 6;
        const int slotBaseline = yTop + lineH * (state.activePart - 1) + tmEq.tmAscent;
        const size_t slotIndex = (size_t)(state.activePart - 1);
        if (slotIndex < obj.slots.size() && SequenceHasVisibleContent(obj.slots[slotIndex].children))
            found = setSequenceCaret(obj.slots[slotIndex].children, slotBaseline, slotX, state.activeNodePath);
        else
            found = setLeftTextCaret(obj.SlotText(state.activePart), slotBaseline, slotX);
        break;
    }

    case MathType::Matrix:
    case MathType::Determinant:
    {
        SelectObject(hdc, renderBaseFont);
        TEXTMETRICW tmRow = {};
        GetTextMetricsW(hdc, &tmRow);
        const int cellPadX = (std::max)(6, (int)tmRow.tmAveCharWidth / 2);
        const int cellPadY = (std::max)(4, (int)tmRow.tmHeight / 6);
        const int colGap = (std::max)(12, (int)tmRow.tmAveCharWidth);
        const int rowGap = (std::max)(8, (int)tmRow.tmHeight / 4);
        int bracketInset = (std::max)(8, (int)(tmBase.tmAveCharWidth * 0.8));
        const bool hasNestedCells[] = {
            !obj.slots.empty() && !obj.slots[0].children.empty(),
            obj.slots.size() > 1 && !obj.slots[1].children.empty(),
            obj.slots.size() > 2 && !obj.slots[2].children.empty(),
            obj.slots.size() > 3 && !obj.slots[3].children.empty()
        };
        const std::wstring* cellParts[] = { &part1, &part2, &part3, &part4 };
        SIZE cellSz[4] = {};
        for (int cellIndex = 0; cellIndex < 4; ++cellIndex)
        {
            if (hasNestedCells[cellIndex])
                cellSz[cellIndex] = MeasureMathNodeSequence(hdc, obj.slots[(size_t)cellIndex].children, tmBase);
            else
                cellSz[cellIndex] = MeasureDisplayText(hdc, *cellParts[cellIndex]);
        }

        const int colWidths[2] = {
            (std::max)(cellSz[0].cx, cellSz[2].cx),
            (std::max)(cellSz[1].cx, cellSz[3].cx)
        };
        const int rowHeights[2] = {
            (std::max)(cellSz[0].cy, cellSz[1].cy),
            (std::max)(cellSz[2].cy, cellSz[3].cy)
        };
        const int contentX = ptStart.x + bracketInset + 10;
        const int colLeft[2] = {
            contentX + cellPadX,
            contentX + cellPadX * 3 + colWidths[0] + colGap
        };
        const int topY = yMid - (rowHeights[0] + rowHeights[1] + rowGap + cellPadY * 4) / 2;
        const int rowTop[2] = {
            topY + cellPadY,
            topY + cellPadY * 3 + rowHeights[0] + rowGap
        };
        const int baselineY[2] = {
            rowTop[0] + tmRow.tmAscent,
            rowTop[1] + tmRow.tmAscent
        };

        const int activeCellIndex = state.activePart - 1;
        if (activeCellIndex >= 0 && activeCellIndex < 4)
        {
            const int rowIndex = activeCellIndex / 2;
            const int colIndex = activeCellIndex % 2;
            const int drawX = colLeft[colIndex] + (colWidths[colIndex] - cellSz[activeCellIndex].cx) / 2;
            const int drawY = baselineY[rowIndex];
            if ((size_t)activeCellIndex < obj.slots.size() && !obj.slots[(size_t)activeCellIndex].children.empty())
                found = setSequenceCaret(obj.slots[(size_t)activeCellIndex].children, drawY, drawX, state.activeNodePath);
            else
                found = setLeftTextCaret(obj.SlotText(state.activePart), drawY, drawX);
        }
        break;
    }

    case MathType::SquareRoot:
    {
        SelectObject(hdc, renderBaseFont);
        TEXTMETRICW tmExpr = {};
        GetTextMetricsW(hdc, &tmExpr);
        SelectObject(hdc, limitFont);
        TEXTMETRICW tmIndex = {};
        GetTextMetricsW(hdc, &tmIndex);
        const bool hasNestedRadicand = !obj.slots.empty() && !obj.slots[0].children.empty();
        const bool hasNestedIndex = obj.slots.size() > 1 && !obj.slots[1].children.empty();
        SelectObject(hdc, renderBaseFont);
        SIZE exprSz = hasNestedRadicand ? MeasureMathNodeSequence(hdc, obj.slots[0].children, tmBase) : MeasureDisplayText(hdc, part1);
        const bool showIndex = !part2.empty() || state.activePart == 2;
        SIZE indexSz = {};
        if (showIndex)
        {
            SelectObject(hdc, limitFont);
            indexSz = hasNestedIndex ? MeasureMathNodeSequence(hdc, obj.slots[1].children, tmBase) : MeasureDisplayText(hdc, part2);
        }

        const int pad = (std::max)(2, (int)(tmBase.tmHeight / 8));
        const int overlineGap = (std::max)(2, (int)(tmBase.tmHeight / 10));
        const int radicalW = (int)(tmBase.tmAveCharWidth * 1.0);
        const int penWidth = (std::max)(1, (int)(1.2 * renderScale));
        const int indexGap = (std::max)(2, (int)(tmBase.tmAveCharWidth / 2));
        const int radTop = yMid - tmExpr.tmAscent - overlineGap - penWidth;
        const int xRadStart = ptStart.x + (showIndex ? indexSz.cx + indexGap : 0);
        const int xRadPeak = xRadStart + radicalW;
        const int xExprStart = xRadPeak + pad;
        const int xIndex = xRadStart - indexGap;
        const int yIndex = radTop + tmIndex.tmAscent + overlineGap;
        (void)exprSz;

        if (state.activePart == 2)
        {
            const int caretX = xIndex - indexSz.cx;
            if (hasNestedIndex)
                found = setSequenceCaret(obj.slots[1].children, yIndex, caretX, state.activeNodePath);
            else
                found = setLeftTextCaret(part2, yIndex, caretX);
        }
        else if (state.activePart == 1)
        {
            if (hasNestedRadicand)
                found = setSequenceCaret(obj.slots[0].children, yMid, xExprStart, state.activeNodePath);
            else
                found = setLeftTextCaret(part1, yMid, xExprStart);
        }
        break;
    }

    case MathType::AbsoluteValue:
    {
        SelectObject(hdc, renderBaseFont);
        TEXTMETRICW tmExpr = {};
        GetTextMetricsW(hdc, &tmExpr);
        const bool hasNestedExpr = !obj.slots.empty() && !obj.slots[0].children.empty();
        const int pad = (std::max)(2, (int)(tmBase.tmHeight / 6));
        const int penWidth = (std::max)(2, (int)(1.8 * renderScale));
        const int xLeftBar = ptStart.x + pad;
        const int xExprStart = xLeftBar + pad + penWidth;
        (void)tmExpr;
        if (hasNestedExpr)
            found = setSequenceCaret(obj.slots[0].children, yMid, xExprStart, state.activeNodePath);
        else
            found = setLeftTextCaret(part1, yMid, xExprStart);
        break;
    }

    case MathType::Power:
    {
        SelectObject(hdc, renderBaseFont);
        const bool hasNestedBase = !obj.slots.empty() && !obj.slots[0].children.empty();
        const bool hasNestedExponent = obj.slots.size() > 1 && !obj.slots[1].children.empty();
        SIZE baseSz = hasNestedBase ? MeasureMathNodeSequence(hdc, obj.slots[0].children, tmBase) : MeasureDisplayText(hdc, part1);
        const int xBase = ptStart.x + 2;
        const int yBase = yMid;
        const int xExp = xBase + baseSz.cx + 2;
        const int yExp = yBase - (tmBase.tmAscent / 2);

        if (state.activePart == 1)
        {
            if (hasNestedBase)
                found = setSequenceCaret(obj.slots[0].children, yBase, xBase, state.activeNodePath);
            else
                found = setLeftTextCaret(part1, yBase, xBase);
        }
        else if (state.activePart == 2)
        {
            if (hasNestedExponent)
                found = setSequenceCaret(obj.slots[1].children, yExp, xExp, state.activeNodePath);
            else
                found = setLeftTextCaret(part2, yExp, xExp);
        }
        break;
    }

    case MathType::Logarithm:
    {
        SelectObject(hdc, renderBaseFont);
        SIZE logSz = {};
        GetTextExtentPoint32W(hdc, L"log", 3, &logSz);
        const bool hasNestedBase = !obj.slots.empty() && !obj.slots[0].children.empty();
        const bool hasNestedArg = obj.slots.size() > 1 && !obj.slots[1].children.empty();
        SIZE baseSz = hasNestedBase ? MeasureMathNodeSequence(hdc, obj.slots[0].children, tmBase) : MeasureDisplayText(hdc, part1);
        const int xLog = ptStart.x + 2;
        const int yLog = yMid;
        const int xBase = xLog + logSz.cx + 1;
        const int yBase = yLog + (tmBase.tmDescent / 2);
        const int xArg = xBase + baseSz.cx + 2;

        if (state.activePart == 1)
        {
            if (hasNestedBase)
                found = setSequenceCaret(obj.slots[0].children, yBase, xBase, state.activeNodePath);
            else
                found = setLeftTextCaret(part1, yBase, xBase);
        }
        else if (state.activePart == 2)
        {
            if (hasNestedArg)
                found = setSequenceCaret(obj.slots[1].children, yMid, xArg, state.activeNodePath);
            else
                found = setLeftTextCaret(part2, yMid, xArg);
        }
        break;
    }

    default:
        break;
    }

    SelectObject(hdc, oldFont);
    RestoreDC(hdc, saved);
    DeleteObject(renderBaseFont);
    DeleteObject(limitFont);
    return found;
}

void MathRenderer::Draw(HWND hEdit, HDC hdc, const MathObject& obj, size_t objIndex, const MathTypingState& state)
{
    if (obj.barLen <= 0) return;
    const std::wstring& part1 = obj.SlotText(1);
    const std::wstring& part2 = obj.SlotText(2);
    const std::wstring& part3 = obj.SlotText(3);
    const std::wstring& part4 = obj.SlotText(4);
    POINT ptStart = {}, ptEnd = {};
    if (!TryGetCharPos(hEdit, obj.barStart, ptStart)) return;
    if (!TryGetCharPos(hEdit, obj.barStart + std::max<LONG>(0, obj.barLen - 1), ptEnd)) return;

    HFONT baseFont = (HFONT)SendMessage(hEdit, WM_GETFONT, 0, 0);
    if (!baseFont) baseFont = (HFONT)GetStockObject(DEFAULT_GUI_FONT);

    double renderScale = ComputeRenderScale(hEdit, hdc, obj, baseFont);
    // Non-fraction command objects use 2x-height anchors for vertical space,
    // but the renderScale picks up that 2x width too. Compensate so drawn
    // content stays at the correct size relative to zoom.
    if (obj.type == MathType::Summation || obj.type == MathType::Integral
        || obj.type == MathType::SystemOfEquations || obj.type == MathType::Power
        || obj.type == MathType::Logarithm || obj.type == MathType::Sum
        || obj.type == MathType::Matrix || obj.type == MathType::Determinant)
        renderScale *= 0.5;
    HFONT renderBaseFont = CreateScaledFont(baseFont, renderScale, 100);
    HFONT limitFont = CreateScaledFont(baseFont, renderScale, 70);

    const int saved = SaveDC(hdc);
    // Clip to the RichEdit's client rect so we never draw over the toolbar
    // or outside the control's visible area.
    RECT clientRect;
    GetClientRect(hEdit, &clientRect);
    HRGN hClip = CreateRectRgnIndirect(&clientRect);
    SelectClipRgn(hdc, hClip);
    DeleteObject(hClip);
    SetBkMode(hdc, TRANSPARENT);
    SetTextAlign(hdc, TA_BASELINE | TA_CENTER);
    HFONT oldFont = (HFONT)SelectObject(hdc, renderBaseFont);
    TEXTMETRICW tmBase = {};
    GetTextMetricsW(hdc, &tmBase);

    const int barWidth = (ptEnd.x - ptStart.x) + (int)(tmBase.tmAveCharWidth);
    const int xCenter = ptStart.x + (barWidth / 2);
    const int yMid = ptStart.y + tmBase.tmAscent;
    const COLORREF normalColor = GetDefaultTextColor(hEdit);
    const COLORREF activeColor = GetActiveColor(hEdit);

    auto DrawActiveCaret = [&](int x, int baseline) {
        if (!(state.active && state.objectIndex == objIndex))
            return;
        HPEN caretPen = CreatePen(PS_SOLID, 1, activeColor);
        HPEN oldCaretPen = (HPEN)SelectObject(hdc, caretPen);
        MoveToEx(hdc, x, baseline - tmBase.tmAscent - 1, NULL);
        LineTo(hdc, x, baseline + tmBase.tmDescent + 1);
        SelectObject(hdc, oldCaretPen);
        DeleteObject(caretPen);
    };

    auto DrawSequenceCaret = [&](const std::vector<MathNode>& nodes, int baseline, int x, const std::vector<size_t>& path) {
        POINT caretPt = { x + MeasureMathNodeSequence(hdc, nodes, tmBase).cx, baseline };
        if (!path.empty())
            TryGetSequenceCaret(hdc, nodes, path, 0, x, baseline, tmBase, caretPt);
        DrawActiveCaret(caretPt.x, caretPt.y);
    };

    auto DrawPart = [&](const std::wstring& text, int x, int y, int partIdx) {
        bool isActive = (state.active && state.objectIndex == objIndex && state.activePart == partIdx);
        SetTextColor(hdc, isActive ? activeColor : normalColor);
        
        if (kDebugOverlay)
        {
            SIZE sz = {};
            GetTextExtentPoint32W(hdc, text.empty() ? L"?" : text.c_str(), (int)std::max<size_t>(1, text.size()), &sz);
            RECT rc = { x - sz.cx/2, y - sz.cy, x + sz.cx/2, y };
            if (GetTextAlign(hdc) & TA_LEFT) {
                rc.left = x; rc.right = x + sz.cx;
            }
            HBRUSH hBr = CreateSolidBrush(isActive ? RGB(255, 0, 0) : RGB(0, 255, 0));
            FrameRect(hdc, &rc, hBr);
            DeleteObject(hBr);
        }

        if (text.empty() && isActive) TextOutW(hdc, x, y, L"?", 1);
        else TextOutW(hdc, x, y, text.c_str(), (int)text.size());
    };

    if (kDebugOverlay)
    {
        HPEN hPen = CreatePen(PS_DOT, 1, RGB(200, 200, 200));
        HPEN oldP = (HPEN)SelectObject(hdc, hPen);
        MoveToEx(hdc, ptStart.x - 20, yMid, NULL);
        LineTo(hdc, ptEnd.x + 40, yMid);
        MoveToEx(hdc, xCenter, yMid - 40, NULL);
        LineTo(hdc, xCenter, yMid + 40);
        SelectObject(hdc, oldP);
        DeleteObject(hPen);
    }

    // (Anchor characters are hidden via CHARFORMAT in math_editor.cpp,
    //  so we only need to paint-over for non-Fraction types that may
    //  still show raw anchor glyphs.)
    if (obj.type != MathType::Fraction)
    {
        COLORREF bk = GetRichEditBkColor(hEdit);
        SIZE charSize = {};
        GetTextExtentPoint32W(hdc, L"\u2500", 1, &charSize);
        int coverRight = ptEnd.x + (std::max)((int)charSize.cx, (int)tmBase.tmAveCharWidth) + 4;
        RECT rc = { ptStart.x - 2, ptStart.y - 4, coverRight, ptStart.y + tmBase.tmHeight + 4 };
        HBRUSH hBr = CreateSolidBrush(bk);
        FillRect(hdc, &rc, hBr);
        DeleteObject(hBr);
    }

    if (obj.type == MathType::Fraction)
    {
        // 50% bigger than the default limit font (70% * 1.5 ≈ 105%)
        HFONT fracFont = CreateScaledFont(baseFont, renderScale, 105);
        SelectObject(hdc, fracFont);
        TEXTMETRICW tmL = {}; GetTextMetricsW(hdc, &tmL);
        const bool hasNestedNumerator = !obj.slots.empty() && !obj.slots[0].children.empty();
        const bool hasNestedDenominator = obj.slots.size() > 1 && !obj.slots[1].children.empty();

        const int gap = 4;  // pixels between bar and text

        // Place the bar at the vertical center of the anchor cell.
        const int barY = ptStart.y + tmBase.tmHeight / 2;

        SIZE numeratorSz = hasNestedNumerator ? MeasureMathNodeSequence(hdc, obj.slots[0].children, tmBase) : SIZE{};
        SIZE denominatorSz = hasNestedDenominator ? MeasureMathNodeSequence(hdc, obj.slots[1].children, tmBase) : SIZE{};

        // Numerator: baseline so bottom of text is `gap` above the bar
        if (hasNestedNumerator)
        {
            const int numeratorX = xCenter - numeratorSz.cx / 2;
            const COLORREF numeratorColor = (state.active && state.objectIndex == objIndex && state.activePart == 1) ? activeColor : normalColor;
            DrawMathNodeSequence(hdc, obj.slots[0].children, numeratorX, barY - gap - tmL.tmDescent, tmBase, numeratorColor);
        }
        else DrawPart(part1, xCenter, barY - gap - tmL.tmDescent, 1);
        // Denominator: baseline so top of text is `gap` below the bar
        if (hasNestedDenominator)
        {
            const int denominatorX = xCenter - denominatorSz.cx / 2;
            const COLORREF denominatorColor = (state.active && state.objectIndex == objIndex && state.activePart == 2) ? activeColor : normalColor;
            DrawMathNodeSequence(hdc, obj.slots[1].children, denominatorX, barY + gap + tmL.tmAscent, tmBase, denominatorColor);
        }
        else DrawPart(part2, xCenter, barY + gap + tmL.tmAscent, 2);

        SelectObject(hdc, oldFont);
        DeleteObject(fracFont);

        // Draw the vinculum (fraction bar) via GDI
        HPEN hPen = CreatePen(PS_SOLID, (std::max)(1, (int)(1.2 * renderScale)), normalColor);
        HPEN oldPen = (HPEN)SelectObject(hdc, hPen);
        MoveToEx(hdc, ptStart.x, barY, NULL);
        LineTo(hdc, ptStart.x + barWidth, barY);
        SelectObject(hdc, oldPen);
        DeleteObject(hPen);

        // Draw result to the right, vertically centered on the bar
        if (!obj.resultText.empty()) {
            SetTextAlign(hdc, TA_BASELINE | TA_LEFT);
            HFONT boldFont = CreateScaledFont(baseFont, renderScale, 100);
            LOGFONTW lfB = {}; GetObjectW(boldFont, sizeof(lfB), &lfB);
            lfB.lfWeight = FW_BOLD;
            DeleteObject(boldFont);
            boldFont = CreateFontIndirectW(&lfB);
            HFONT prevF = (HFONT)SelectObject(hdc, boldFont);
            TEXTMETRICW tmR = {}; GetTextMetricsW(hdc, &tmR);
            // Align equals sign's horizontal stroke with the bar
            int resultBaseline = barY + (tmR.tmAscent - tmR.tmDescent) / 2;
            SetTextColor(hdc, activeColor);
            TextOutW(hdc, ptStart.x + barWidth + 4, resultBaseline, obj.resultText.c_str(), (int)obj.resultText.size());
            SelectObject(hdc, prevF);
            DeleteObject(boldFont);
        }

        if (state.active && state.objectIndex == objIndex)
        {
            const int caretBaseline = (state.activePart == 1) ? (barY - gap - tmL.tmDescent) : (barY + gap + tmL.tmAscent);
            if ((size_t)(state.activePart - 1) < obj.slots.size() && !obj.slots[(size_t)(state.activePart - 1)].children.empty())
            {
                const SIZE caretSlotSize = MeasureMathNodeSequence(hdc, obj.slots[(size_t)(state.activePart - 1)].children, tmBase);
                const int caretX = xCenter - caretSlotSize.cx / 2;
                DrawSequenceCaret(obj.slots[(size_t)(state.activePart - 1)].children, caretBaseline, caretX, state.activePart == 1 || state.activePart == 2 ? state.activeNodePath : std::vector<size_t>{});
            }
            else
            {
                const std::wstring& caretText = obj.SlotText(state.activePart);
                const SIZE caretTextSize = MeasureDisplayText(hdc, caretText);
                DrawActiveCaret(xCenter + caretTextSize.cx / 2, caretBaseline);
            }
        }
    }
    else if (obj.type == MathType::Summation)
    {
        const bool hasNestedUpper = obj.slots.size() > 0 && SequenceHasVisibleContent(obj.slots[0].children);
        const bool hasNestedLower = obj.slots.size() > 1 && SequenceHasVisibleContent(obj.slots[1].children);
        const bool hasNestedExpr = obj.slots.size() > 2 && !obj.slots[2].children.empty();
        // Draw sigma symbol via GDI in Cambria Math
        {
            LOGFONTW lfSym = {};
            GetObjectW(renderBaseFont, sizeof(lfSym), &lfSym);
            wcscpy_s(lfSym.lfFaceName, L"Cambria Math");
            HFONT symbolFont = CreateFontIndirectW(&lfSym);
            HFONT prevF = (HFONT)SelectObject(hdc, symbolFont);
            SetTextColor(hdc, normalColor);
            TextOutW(hdc, xCenter, yMid, L"\u2211", 1);
            SelectObject(hdc, prevF);
            DeleteObject(symbolFont);
        }

        SelectObject(hdc, limitFont);
        TEXTMETRICW tmL = {}; GetTextMetricsW(hdc, &tmL);

        const int upperBaseline = yMid - tmBase.tmAscent - 2;
        const int lowerBaseline = yMid + tmBase.tmDescent + tmL.tmAscent + 2;
        if (hasNestedUpper)
        {
            const SIZE upperSz = MeasureMathNodeSequence(hdc, obj.slots[0].children, tmBase);
            const COLORREF upperColor = (state.active && state.objectIndex == objIndex && state.activePart == 1) ? activeColor : normalColor;
            DrawMathNodeSequence(hdc, obj.slots[0].children, xCenter - upperSz.cx / 2, upperBaseline, tmBase, upperColor);
        }
        else DrawPart(part1, xCenter, upperBaseline, 1);

        if (hasNestedLower)
        {
            const SIZE lowerSz = MeasureMathNodeSequence(hdc, obj.slots[1].children, tmBase);
            const COLORREF lowerColor = (state.active && state.objectIndex == objIndex && state.activePart == 2) ? activeColor : normalColor;
            DrawMathNodeSequence(hdc, obj.slots[1].children, xCenter - lowerSz.cx / 2, lowerBaseline, tmBase, lowerColor);
        }
        else DrawPart(part2, xCenter, lowerBaseline, 2);
        
        SetTextAlign(hdc, TA_BASELINE | TA_LEFT);
        SelectObject(hdc, renderBaseFont);
        // Draw expression slightly to the right of the anchor spaces
        if (hasNestedExpr)
        {
            const COLORREF exprColor = (state.active && state.objectIndex == objIndex && state.activePart == 3) ? activeColor : normalColor;
            DrawMathNodeSequence(hdc, obj.slots[2].children, ptEnd.x + 4, yMid, tmBase, exprColor);
        }
        else DrawPart(part3, ptEnd.x + 4, yMid, 3);

        // Draw result (e.g. " ＝ 302") right after expression via GDI
        if (!obj.resultText.empty()) {
            SIZE exprSz = {};
            if (hasNestedExpr) exprSz = MeasureMathNodeSequence(hdc, obj.slots[2].children, tmBase);
            else GetTextExtentPoint32W(hdc, part3.c_str(), (int)part3.size(), &exprSz);
            SetTextColor(hdc, activeColor);
            HFONT boldFont = CreateScaledFont(baseFont, renderScale, 100);
            LOGFONTW lfB = {}; GetObjectW(boldFont, sizeof(lfB), &lfB);
            lfB.lfWeight = FW_BOLD;
            DeleteObject(boldFont);
            boldFont = CreateFontIndirectW(&lfB);
            HFONT prevF = (HFONT)SelectObject(hdc, boldFont);
            TextOutW(hdc, ptEnd.x + 4 + exprSz.cx + 4, yMid, obj.resultText.c_str(), (int)obj.resultText.size());
            SelectObject(hdc, prevF);
            DeleteObject(boldFont);
        }

        if (state.active && state.objectIndex == objIndex)
        {
            if (state.activePart == 1 && hasNestedUpper)
            {
                const SIZE upperSz = MeasureMathNodeSequence(hdc, obj.slots[0].children, tmBase);
                DrawSequenceCaret(obj.slots[0].children, upperBaseline, xCenter - upperSz.cx / 2, state.activeNodePath);
            }
            else if (state.activePart == 2 && hasNestedLower)
            {
                const SIZE lowerSz = MeasureMathNodeSequence(hdc, obj.slots[1].children, tmBase);
                DrawSequenceCaret(obj.slots[1].children, lowerBaseline, xCenter - lowerSz.cx / 2, state.activeNodePath);
            }
        }
    }
    else if (obj.type == MathType::Product)
    {
        const bool hasNestedUpper = obj.slots.size() > 0 && SequenceHasVisibleContent(obj.slots[0].children);
        const bool hasNestedLower = obj.slots.size() > 1 && SequenceHasVisibleContent(obj.slots[1].children);
        const bool hasNestedExpr = obj.slots.size() > 2 && !obj.slots[2].children.empty();
        // Draw product symbol via GDI in Cambria Math
        {
            LOGFONTW lfSym = {};
            GetObjectW(renderBaseFont, sizeof(lfSym), &lfSym);
            wcscpy_s(lfSym.lfFaceName, L"Cambria Math");
            HFONT symbolFont = CreateFontIndirectW(&lfSym);
            HFONT prevF = (HFONT)SelectObject(hdc, symbolFont);
            SetTextColor(hdc, normalColor);
            TextOutW(hdc, xCenter, yMid, L"\u03A0", 1);
            SelectObject(hdc, prevF);
            DeleteObject(symbolFont);
        }

        SelectObject(hdc, limitFont);
        TEXTMETRICW tmL = {}; GetTextMetricsW(hdc, &tmL);

        const int upperBaseline = yMid - tmBase.tmAscent - 2;
        const int lowerBaseline = yMid + tmBase.tmDescent + tmL.tmAscent + 2;
        if (hasNestedUpper)
        {
            const SIZE upperSz = MeasureMathNodeSequence(hdc, obj.slots[0].children, tmBase);
            const COLORREF upperColor = (state.active && state.objectIndex == objIndex && state.activePart == 1) ? activeColor : normalColor;
            DrawMathNodeSequence(hdc, obj.slots[0].children, xCenter - upperSz.cx / 2, upperBaseline, tmBase, upperColor);
        }
        else DrawPart(part1, xCenter, upperBaseline, 1);

        if (hasNestedLower)
        {
            const SIZE lowerSz = MeasureMathNodeSequence(hdc, obj.slots[1].children, tmBase);
            const COLORREF lowerColor = (state.active && state.objectIndex == objIndex && state.activePart == 2) ? activeColor : normalColor;
            DrawMathNodeSequence(hdc, obj.slots[1].children, xCenter - lowerSz.cx / 2, lowerBaseline, tmBase, lowerColor);
        }
        else DrawPart(part2, xCenter, lowerBaseline, 2);

        SetTextAlign(hdc, TA_BASELINE | TA_LEFT);
        SelectObject(hdc, renderBaseFont);
        // Draw expression slightly to the right of the anchor spaces
        if (hasNestedExpr)
        {
            const COLORREF exprColor = (state.active && state.objectIndex == objIndex && state.activePart == 3) ? activeColor : normalColor;
            DrawMathNodeSequence(hdc, obj.slots[2].children, ptEnd.x + 4, yMid, tmBase, exprColor);
        }
        else DrawPart(part3, ptEnd.x + 4, yMid, 3);

        // Draw result (e.g. " ＝ 302") right after expression via GDI
        if (!obj.resultText.empty()) {
            SIZE exprSz = {};
            if (hasNestedExpr) exprSz = MeasureMathNodeSequence(hdc, obj.slots[2].children, tmBase);
            else GetTextExtentPoint32W(hdc, part3.c_str(), (int)part3.size(), &exprSz);
            SetTextColor(hdc, activeColor);
            HFONT boldFont = CreateScaledFont(baseFont, renderScale, 100);
            LOGFONTW lfB = {}; GetObjectW(boldFont, sizeof(lfB), &lfB);
            lfB.lfWeight = FW_BOLD;
            DeleteObject(boldFont);
            boldFont = CreateFontIndirectW(&lfB);
            HFONT prevF = (HFONT)SelectObject(hdc, boldFont);
            TextOutW(hdc, ptEnd.x + 4 + exprSz.cx + 4, yMid, obj.resultText.c_str(), (int)obj.resultText.size());
            SelectObject(hdc, prevF);
            DeleteObject(boldFont);
        }

        if (state.active && state.objectIndex == objIndex)
        {
            if (state.activePart == 1 && hasNestedUpper)
            {
                const SIZE upperSz = MeasureMathNodeSequence(hdc, obj.slots[0].children, tmBase);
                DrawSequenceCaret(obj.slots[0].children, upperBaseline, xCenter - upperSz.cx / 2, state.activeNodePath);
            }
            else if (state.activePart == 2 && hasNestedLower)
            {
                const SIZE lowerSz = MeasureMathNodeSequence(hdc, obj.slots[1].children, tmBase);
                DrawSequenceCaret(obj.slots[1].children, lowerBaseline, xCenter - lowerSz.cx / 2, state.activeNodePath);
            }
        }
    }
    else if (obj.type == MathType::Sum)
    {
        const bool hasNestedExpr = !obj.slots.empty() && !obj.slots[0].children.empty();
        // Draw expression (numbers separated by operators) - no sigma symbol
        SetTextAlign(hdc, TA_BASELINE | TA_LEFT);
        SelectObject(hdc, renderBaseFont);
        // Draw expression starting at the anchor position (no symbol on the left)
        if (hasNestedExpr)
        {
            const COLORREF exprColor = (state.active && state.objectIndex == objIndex && state.activePart == 1) ? activeColor : normalColor;
            DrawMathNodeSequence(hdc, obj.slots[0].children, ptStart.x + 4, yMid, tmBase, exprColor);
        }
        else DrawPart(part1, ptStart.x + 4, yMid, 1);

        // Draw result (e.g., " ＝ 402365") right after expression via GDI
        if (!obj.resultText.empty()) {
            SIZE exprSz = {};
            if (hasNestedExpr) exprSz = MeasureMathNodeSequence(hdc, obj.slots[0].children, tmBase);
            else GetTextExtentPoint32W(hdc, part1.c_str(), (int)part1.size(), &exprSz);
            SetTextColor(hdc, activeColor);
            HFONT boldFont = CreateScaledFont(baseFont, renderScale, 100);
            LOGFONTW lfB = {}; GetObjectW(boldFont, sizeof(lfB), &lfB);
            lfB.lfWeight = FW_BOLD;
            DeleteObject(boldFont);
            boldFont = CreateFontIndirectW(&lfB);
            HFONT prevF = (HFONT)SelectObject(hdc, boldFont);
            TextOutW(hdc, ptEnd.x + 4 + exprSz.cx + 4, yMid, obj.resultText.c_str(), (int)obj.resultText.size());
            SelectObject(hdc, prevF);
            DeleteObject(boldFont);
        }
    }
    else if (obj.type == MathType::Integral)
    {
        const bool hasNestedUpper = obj.slots.size() > 0 && SequenceHasVisibleContent(obj.slots[0].children);
        const bool hasNestedLower = obj.slots.size() > 1 && SequenceHasVisibleContent(obj.slots[1].children);
        const bool hasNestedExpr = obj.slots.size() > 2 && !obj.slots[2].children.empty();
        // Draw integral symbol via GDI in Cambria Math
        {
            LOGFONTW lfSym = {};
            GetObjectW(renderBaseFont, sizeof(lfSym), &lfSym);
            wcscpy_s(lfSym.lfFaceName, L"Cambria Math");
            HFONT symbolFont = CreateFontIndirectW(&lfSym);
            HFONT prevF = (HFONT)SelectObject(hdc, symbolFont);
            SetTextColor(hdc, normalColor);
            TextOutW(hdc, xCenter, yMid, L"\u222B", 1);
            SelectObject(hdc, prevF);
            DeleteObject(symbolFont);
        }

        SetTextAlign(hdc, TA_BASELINE | TA_LEFT);
        SelectObject(hdc, limitFont);
        TEXTMETRICW tmL = {}; GetTextMetricsW(hdc, &tmL);

        const int upperBaseline = yMid - tmBase.tmAscent + (int)(tmBase.tmAscent * 0.2);
        const int lowerBaseline = yMid + tmBase.tmDescent + 2;
        if (hasNestedUpper)
        {
            const COLORREF upperColor = (state.active && state.objectIndex == objIndex && state.activePart == 1) ? activeColor : normalColor;
            DrawMathNodeSequence(hdc, obj.slots[0].children, ptEnd.x - 2, upperBaseline, tmBase, upperColor);
        }
        else DrawPart(part1, ptEnd.x - 2, upperBaseline, 1);

        if (hasNestedLower)
        {
            const COLORREF lowerColor = (state.active && state.objectIndex == objIndex && state.activePart == 2) ? activeColor : normalColor;
            DrawMathNodeSequence(hdc, obj.slots[1].children, ptEnd.x - 8, lowerBaseline, tmBase, lowerColor);
        }
        else DrawPart(part2, ptEnd.x - 8, lowerBaseline, 2);

        SelectObject(hdc, renderBaseFont);
        if (hasNestedExpr)
        {
            const COLORREF exprColor = (state.active && state.objectIndex == objIndex && state.activePart == 3) ? activeColor : normalColor;
            DrawMathNodeSequence(hdc, obj.slots[2].children, ptEnd.x + 6, yMid, tmBase, exprColor);
        }
        else DrawPart(part3, ptEnd.x + 6, yMid, 3);

        // Draw result right after expression via GDI
        if (!obj.resultText.empty()) {
            SIZE exprSz = {};
            if (hasNestedExpr) exprSz = MeasureMathNodeSequence(hdc, obj.slots[2].children, tmBase);
            else GetTextExtentPoint32W(hdc, part3.c_str(), (int)part3.size(), &exprSz);
            SetTextColor(hdc, activeColor);
            HFONT boldFont = CreateScaledFont(baseFont, renderScale, 100);
            LOGFONTW lfB = {}; GetObjectW(boldFont, sizeof(lfB), &lfB);
            lfB.lfWeight = FW_BOLD;
            DeleteObject(boldFont);
            boldFont = CreateFontIndirectW(&lfB);
            HFONT prevF = (HFONT)SelectObject(hdc, boldFont);
            TextOutW(hdc, ptEnd.x + 6 + exprSz.cx + 4, yMid, obj.resultText.c_str(), (int)obj.resultText.size());
            SelectObject(hdc, prevF);
            DeleteObject(boldFont);
        }

        if (state.active && state.objectIndex == objIndex)
        {
            if (state.activePart == 1 && hasNestedUpper)
                DrawSequenceCaret(obj.slots[0].children, upperBaseline, ptEnd.x - 2, state.activeNodePath);
            else if (state.activePart == 2 && hasNestedLower)
                DrawSequenceCaret(obj.slots[1].children, lowerBaseline, ptEnd.x - 8, state.activeNodePath);
        }
    }
    else if (obj.type == MathType::SystemOfEquations)
    {
        const bool hasNestedEq1 = obj.slots.size() > 0 && SequenceHasVisibleContent(obj.slots[0].children);
        const bool hasNestedEq2 = obj.slots.size() > 1 && SequenceHasVisibleContent(obj.slots[1].children);
        const bool hasNestedEq3 = obj.slots.size() > 2 && SequenceHasVisibleContent(obj.slots[2].children);
        // First measure the equations block to know total height
        SelectObject(hdc, renderBaseFont);
        TEXTMETRICW tmEq = {};
        GetTextMetricsW(hdc, &tmEq);
        int lineH = tmEq.tmHeight + 4;
        int eqCount = 3;
        if (!hasNestedEq3 && part3.empty() && !(state.active && state.objectIndex == objIndex && state.activePart == 3))
            eqCount = 2;
        int totalH = lineH * eqCount;
        // yTop is the baseline of the first equation line
        int yTop = yMid - totalH / 2 + tmEq.tmAscent;

        // Draw a left curly brace using GDI Bezier curves for precise height control
        {
            int blockTop = yTop - tmEq.tmAscent - 2;
            int blockBot = blockTop + totalH + 4;
            int blockMidY = (blockTop + blockBot) / 2;
            int braceW = (int)(tmBase.tmAveCharWidth * 1.2);
            if (braceW < 10) braceW = 10;
            int xRight = ptStart.x + braceW;           // where the arms end (top & bottom)
            int xMid   = ptStart.x + braceW / 2;       // vertical spine of the brace
            int xTip   = ptStart.x;                     // middle tip pointing left
            int armH   = (blockMidY - blockTop);        // half-height

            HPEN hPen = CreatePen(PS_SOLID, (std::max)(1, (int)(1.5 * renderScale)), normalColor);
            HPEN oldPen = (HPEN)SelectObject(hdc, hPen);
            HBRUSH oldBrush = (HBRUSH)SelectObject(hdc, GetStockObject(NULL_BRUSH));

            // Top hook: small curve from top endpoint inward
            // Goes from (xRight, blockTop) curving left to (xMid, blockTop + armH*0.25)
            POINT topHook[4] = {
                { xRight, blockTop },
                { xMid,   blockTop },
                { xMid,   blockTop },
                { xMid,   blockTop + (int)(armH * 0.25) }
            };
            PolyBezier(hdc, topHook, 4);

            // Top arm: straight-ish from spine down to the middle tip
            // Goes from (xMid, blockTop + armH*0.25) curving to (xTip, blockMidY)
            POINT topArm[4] = {
                { xMid, blockTop + (int)(armH * 0.25) },
                { xMid, blockMidY - (int)(armH * 0.15) },
                { xMid, blockMidY - (int)(armH * 0.05) },
                { xTip, blockMidY }
            };
            PolyBezier(hdc, topArm, 4);

            // Bottom arm: from middle tip curving back to spine
            POINT botArm[4] = {
                { xTip, blockMidY },
                { xMid, blockMidY + (int)(armH * 0.05) },
                { xMid, blockMidY + (int)(armH * 0.15) },
                { xMid, blockBot - (int)(armH * 0.25) }
            };
            PolyBezier(hdc, botArm, 4);

            // Bottom hook: from spine curving out to bottom endpoint
            POINT botHook[4] = {
                { xMid, blockBot - (int)(armH * 0.25) },
                { xMid, blockBot },
                { xMid, blockBot },
                { xRight, blockBot }
            };
            PolyBezier(hdc, botHook, 4);

            SelectObject(hdc, oldBrush);
            SelectObject(hdc, oldPen);
            DeleteObject(hPen);

            // Draw equations stacked vertically, left-aligned after the brace
            int eqX = ptStart.x + braceW + 6;
            SetTextAlign(hdc, TA_BASELINE | TA_LEFT);
            SelectObject(hdc, renderBaseFont);
            if (hasNestedEq1)
            {
                const COLORREF eqColor = (state.active && state.objectIndex == objIndex && state.activePart == 1) ? activeColor : normalColor;
                DrawMathNodeSequence(hdc, obj.slots[0].children, eqX, yTop, tmBase, eqColor);
            }
            else DrawPart(part1, eqX, yTop, 1);

            if (hasNestedEq2)
            {
                const COLORREF eqColor = (state.active && state.objectIndex == objIndex && state.activePart == 2) ? activeColor : normalColor;
                DrawMathNodeSequence(hdc, obj.slots[1].children, eqX, yTop + lineH, tmBase, eqColor);
            }
            else DrawPart(part2, eqX, yTop + lineH, 2);

            if (eqCount >= 3)
            {
                if (hasNestedEq3)
                {
                    const COLORREF eqColor = (state.active && state.objectIndex == objIndex && state.activePart == 3) ? activeColor : normalColor;
                    DrawMathNodeSequence(hdc, obj.slots[2].children, eqX, yTop + lineH * 2, tmBase, eqColor);
                }
                else DrawPart(part3, eqX, yTop + lineH * 2, 3);
            }

            // Draw result to the right, vertically centered on the brace
            if (!obj.resultText.empty()) {
                SetTextAlign(hdc, TA_BASELINE | TA_LEFT);
                HFONT boldFont = CreateScaledFont(baseFont, renderScale, 100);
                LOGFONTW lfB = {}; GetObjectW(boldFont, sizeof(lfB), &lfB);
                lfB.lfWeight = FW_BOLD;
                DeleteObject(boldFont);
                boldFont = CreateFontIndirectW(&lfB);
                HFONT prevF = (HFONT)SelectObject(hdc, boldFont);
                TEXTMETRICW tmR = {}; GetTextMetricsW(hdc, &tmR);
                
                // Find the rightmost extent of the equations
                SIZE eq1Sz = hasNestedEq1 ? MeasureMathNodeSequence(hdc, obj.slots[0].children, tmBase) : MeasureDisplayText(hdc, part1);
                SIZE eq2Sz = hasNestedEq2 ? MeasureMathNodeSequence(hdc, obj.slots[1].children, tmBase) : MeasureDisplayText(hdc, part2);
                SIZE eq3Sz = {};
                if (eqCount >= 3)
                    eq3Sz = hasNestedEq3 ? MeasureMathNodeSequence(hdc, obj.slots[2].children, tmBase) : MeasureDisplayText(hdc, part3);
                
                int maxEqWidth = std::max(std::max(eq1Sz.cx, eq2Sz.cx), eq3Sz.cx);
                int resultX = eqX + maxEqWidth + 10; // Add some padding
                int resultBaseline = yMid; // Vertically center with the brace
                
                SetTextColor(hdc, activeColor);
                TextOutW(hdc, resultX, resultBaseline, obj.resultText.c_str(), (int)obj.resultText.size());
                SelectObject(hdc, prevF);
                DeleteObject(boldFont);
            }
        }

        if (state.active && state.objectIndex == objIndex)
        {
            if (state.activePart == 1 || state.activePart == 2 || state.activePart == 3)
            {
                int braceW = (int)(tmBase.tmAveCharWidth * 1.2);
                if (braceW < 10) braceW = 10;
                const int slotX = ptStart.x + braceW + 6;
                const int slotBaseline = yTop + lineH * (state.activePart - 1) + tmEq.tmAscent;
                const size_t slotIndex = (size_t)(state.activePart - 1);
                if (slotIndex < obj.slots.size() && SequenceHasVisibleContent(obj.slots[slotIndex].children))
                    DrawSequenceCaret(obj.slots[slotIndex].children, slotBaseline, slotX, state.activeNodePath);
                else
                {
                    const std::wstring& slotText = obj.SlotText(state.activePart);
                    DrawActiveCaret(slotX + MeasureDisplayText(hdc, slotText).cx, slotBaseline);
                }
            }
        }
    }
    else if (obj.type == MathType::Matrix || obj.type == MathType::Determinant)
    {
        SelectObject(hdc, renderBaseFont);
        TEXTMETRICW tmRow = {};
        GetTextMetricsW(hdc, &tmRow);
        const int cellPadX = (std::max)(6, (int)tmRow.tmAveCharWidth / 2);
        const int cellPadY = (std::max)(4, (int)tmRow.tmHeight / 6);
        const int colGap = (std::max)(12, (int)tmRow.tmAveCharWidth);
        const int rowGap = (std::max)(8, (int)tmRow.tmHeight / 4);
        int bracketInset = (std::max)(8, (int)(tmBase.tmAveCharWidth * 0.8));
        int leftStrokeX = ptStart.x + bracketInset;

        const bool hasNestedCells[] = {
            !obj.slots.empty() && !obj.slots[0].children.empty(),
            obj.slots.size() > 1 && !obj.slots[1].children.empty(),
            obj.slots.size() > 2 && !obj.slots[2].children.empty(),
            obj.slots.size() > 3 && !obj.slots[3].children.empty()
        };
        const std::wstring* cellParts[] = { &part1, &part2, &part3, &part4 };
        SIZE cellSz[4] = {};
        for (int cellIndex = 0; cellIndex < 4; ++cellIndex)
        {
            if (hasNestedCells[cellIndex])
                cellSz[cellIndex] = MeasureMathNodeSequence(hdc, obj.slots[(size_t)cellIndex].children, tmBase);
            else
                GetTextExtentPoint32W(hdc,
                    cellParts[cellIndex]->empty() ? L"?" : cellParts[cellIndex]->c_str(),
                    (int)(cellParts[cellIndex]->empty() ? 1 : cellParts[cellIndex]->size()),
                    &cellSz[cellIndex]);
        }

        const int colWidths[2] = {
            (std::max)(cellSz[0].cx, cellSz[2].cx),
            (std::max)(cellSz[1].cx, cellSz[3].cx)
        };
        const int rowHeights[2] = {
            (std::max)(cellSz[0].cy, cellSz[1].cy),
            (std::max)(cellSz[2].cy, cellSz[3].cy)
        };
        int contentX = ptStart.x + bracketInset + 10;
        int contentWidth = colWidths[0] + colWidths[1] + colGap + cellPadX * 4;
        int totalH = rowHeights[0] + rowHeights[1] + rowGap + cellPadY * 4;
        int topY = yMid - totalH / 2;
        int bottomY = topY + totalH;
        int rightStrokeX = contentX + contentWidth + 8;

        HPEN hPen = CreatePen(PS_SOLID, (std::max)(1, (int)(1.4 * renderScale)), normalColor);
        HPEN oldPen = (HPEN)SelectObject(hdc, hPen);
        if (obj.type == MathType::Matrix)
        {
            MoveToEx(hdc, leftStrokeX + 5, topY, NULL);
            LineTo(hdc, leftStrokeX, topY);
            LineTo(hdc, leftStrokeX, bottomY);
            LineTo(hdc, leftStrokeX + 5, bottomY);
            MoveToEx(hdc, rightStrokeX - 5, topY, NULL);
            LineTo(hdc, rightStrokeX, topY);
            LineTo(hdc, rightStrokeX, bottomY);
            LineTo(hdc, rightStrokeX - 5, bottomY);
        }
        else
        {
            MoveToEx(hdc, leftStrokeX, topY, NULL);
            LineTo(hdc, leftStrokeX, bottomY);
            MoveToEx(hdc, rightStrokeX, topY, NULL);
            LineTo(hdc, rightStrokeX, bottomY);
        }
        SelectObject(hdc, oldPen);
        DeleteObject(hPen);

        SetTextAlign(hdc, TA_BASELINE | TA_LEFT);
        SelectObject(hdc, renderBaseFont);
        const int colLeft[2] = {
            contentX + cellPadX,
            contentX + cellPadX * 3 + colWidths[0] + colGap
        };
        const int rowTop[2] = {
            topY + cellPadY,
            topY + cellPadY * 3 + rowHeights[0] + rowGap
        };
        const int baselineY[2] = {
            rowTop[0] + tmRow.tmAscent,
            rowTop[1] + tmRow.tmAscent
        };

        auto DrawMatrixCell = [&](int partIndex, int rowIndex, int colIndex, const std::wstring& text)
        {
            const int cellIndex = partIndex - 1;
            const int drawX = colLeft[colIndex] + (colWidths[colIndex] - cellSz[cellIndex].cx) / 2;
            const int drawY = baselineY[rowIndex];
            if (hasNestedCells[cellIndex])
            {
                const COLORREF cellColor = (state.active && state.objectIndex == objIndex && state.activePart == partIndex) ? activeColor : normalColor;
                DrawMathNodeSequence(hdc, obj.slots[(size_t)cellIndex].children, drawX, drawY, tmBase, cellColor);
            }
            else
            {
                DrawPart(text, drawX, drawY, partIndex);
            }
        };

        DrawMatrixCell(1, 0, 0, part1);
        DrawMatrixCell(2, 0, 1, part2);
        DrawMatrixCell(3, 1, 0, part3);
        DrawMatrixCell(4, 1, 1, part4);

        if (!obj.resultText.empty()) {
            SetTextColor(hdc, activeColor);
            HFONT boldFont = CreateScaledFont(baseFont, renderScale, 100);
            LOGFONTW lfB = {}; GetObjectW(boldFont, sizeof(lfB), &lfB);
            lfB.lfWeight = FW_BOLD;
            DeleteObject(boldFont);
            boldFont = CreateFontIndirectW(&lfB);
            HFONT prevF = (HFONT)SelectObject(hdc, boldFont);
            TextOutW(hdc, rightStrokeX + 10, yMid, obj.resultText.c_str(), (int)obj.resultText.size());
            SelectObject(hdc, prevF);
            DeleteObject(boldFont);
        }

        if (state.active && state.objectIndex == objIndex && state.activePart >= 1 && state.activePart <= 4)
        {
            const int activeCellIndex = state.activePart - 1;
            const int rowIndex = activeCellIndex / 2;
            const int colIndex = activeCellIndex % 2;
            const int drawX = colLeft[colIndex] + (colWidths[colIndex] - cellSz[activeCellIndex].cx) / 2;
            const int drawY = baselineY[rowIndex];
            if ((size_t)activeCellIndex < obj.slots.size() && !obj.slots[(size_t)activeCellIndex].children.empty())
                DrawSequenceCaret(obj.slots[(size_t)activeCellIndex].children, drawY, drawX, state.activeNodePath);
            else
                DrawActiveCaret(drawX + MeasureDisplayText(hdc, obj.SlotText(state.activePart)).cx, drawY);
        }
    }
    else if (obj.type == MathType::SquareRoot)
    {
        // Measure the expression text to determine radical size
        SelectObject(hdc, renderBaseFont);
        TEXTMETRICW tmExpr = {};
        GetTextMetricsW(hdc, &tmExpr);

        SelectObject(hdc, limitFont);
        TEXTMETRICW tmIndex = {};
        GetTextMetricsW(hdc, &tmIndex);

        const bool hasNestedRadicand = !obj.slots.empty() && !obj.slots[0].children.empty();
        const bool hasNestedIndex = obj.slots.size() > 1 && !obj.slots[1].children.empty();

        // Get expression text size (part1 is the radicand)
        SelectObject(hdc, renderBaseFont);
        SIZE exprSz = {};
        if (hasNestedRadicand)
            exprSz = MeasureMathNodeSequence(hdc, obj.slots[0].children, tmBase);
        else {
            const wchar_t* exprText = part1.empty() ? L"?" : part1.c_str();
            int exprLen = part1.empty() ? 1 : (int)part1.size();
            GetTextExtentPoint32W(hdc, exprText, exprLen, &exprSz);
        }

        const bool showIndex = !part2.empty() || (state.active && state.objectIndex == objIndex && state.activePart == 2);
        SIZE indexSz = {};
        if (showIndex) {
            SelectObject(hdc, limitFont);
            if (hasNestedIndex)
                indexSz = MeasureMathNodeSequence(hdc, obj.slots[1].children, tmBase);
            else {
                const wchar_t* indexText = part2.empty() ? L"?" : part2.c_str();
                int indexLen = part2.empty() ? 1 : (int)part2.size();
                GetTextExtentPoint32W(hdc, indexText, indexLen, &indexSz);
            }
        }

        // Layout constants
        const int pad = (std::max)(2, (int)(tmBase.tmHeight / 8));
        const int overlineGap = (std::max)(2, (int)(tmBase.tmHeight / 10));
        const int radicalW = (int)(tmBase.tmAveCharWidth * 1.0);  // width of the radical sign part
        const int penWidth = (std::max)(1, (int)(1.2 * renderScale));
        const int indexGap = (std::max)(2, (int)(tmBase.tmAveCharWidth / 2));

        // Radical vertical extents
        int radTop = yMid - tmExpr.tmAscent - overlineGap - penWidth;  // top of the overline
        int radBot = yMid + tmExpr.tmDescent + pad;                    // bottom of radical
        int radMid = radBot - (int)((radBot - radTop) * 0.35);         // the "valley" of the checkmark

        // X positions
        int xRadStart = ptStart.x + (showIndex ? indexSz.cx + indexGap : 0);  // leftmost point (short leading stroke)
        int xValley   = xRadStart + radicalW / 3;         // bottom of V
        int xRadPeak  = xRadStart + radicalW;              // top of the upstroke, starts overline
        int xExprStart = xRadPeak + pad;                   // where expression text begins
        int xOverlineEnd = xExprStart + exprSz.cx + pad;   // right end of overline
        int xIndex = xRadStart - indexGap;
        int yIndex = radTop + tmIndex.tmAscent + overlineGap;

        // Draw the radical sign using polyline
        HPEN hPen = CreatePen(PS_SOLID, penWidth, normalColor);
        HPEN oldPen = (HPEN)SelectObject(hdc, hPen);

        // Leading short horizontal-ish stroke
        MoveToEx(hdc, xRadStart, radMid - (int)(tmBase.tmHeight * 0.05), NULL);
        LineTo(hdc, xValley, radBot);       // down to the valley
        LineTo(hdc, xRadPeak, radTop);      // up to the peak
        LineTo(hdc, xOverlineEnd, radTop);  // overline across the top

        // Small tail on the right end of overline going down slightly
        LineTo(hdc, xOverlineEnd, radTop + (int)(tmBase.tmHeight * 0.1));

        SelectObject(hdc, oldPen);
        DeleteObject(hPen);

        if (showIndex) {
            SetTextAlign(hdc, TA_BASELINE | TA_RIGHT);
            SelectObject(hdc, limitFont);
            if (hasNestedIndex)
            {
                const COLORREF indexColor = (state.active && state.objectIndex == objIndex && state.activePart == 2) ? activeColor : normalColor;
                DrawMathNodeSequence(hdc, obj.slots[1].children, xIndex - indexSz.cx, yIndex, tmBase, indexColor);
            }
            else DrawPart(part2, xIndex, yIndex, 2);
        }

        // Draw the expression (part1) under the overline
        SetTextAlign(hdc, TA_BASELINE | TA_LEFT);
        SelectObject(hdc, renderBaseFont);
        if (hasNestedRadicand)
        {
            const COLORREF nestedColor = (state.active && state.objectIndex == objIndex && state.activePart == 1) ? activeColor : normalColor;
            DrawMathNodeSequence(hdc, obj.slots[0].children, xExprStart, yMid, tmBase, nestedColor);
        }
        else
        {
            DrawPart(part1, xExprStart, yMid, 1);
        }

        // Draw result right after the radical if present
        if (!obj.resultText.empty()) {
            SetTextColor(hdc, activeColor);
            HFONT boldFont = CreateScaledFont(baseFont, renderScale, 100);
            LOGFONTW lfB = {}; GetObjectW(boldFont, sizeof(lfB), &lfB);
            lfB.lfWeight = FW_BOLD;
            DeleteObject(boldFont);
            boldFont = CreateFontIndirectW(&lfB);
            HFONT prevF = (HFONT)SelectObject(hdc, boldFont);
            TextOutW(hdc, xOverlineEnd + 6, yMid, obj.resultText.c_str(), (int)obj.resultText.size());
            SelectObject(hdc, prevF);
            DeleteObject(boldFont);
        }

        if (state.active && state.objectIndex == objIndex && (state.activePart == 1 || state.activePart == 2))
        {
            if (state.activePart == 2)
            {
                const int caretX = xIndex - indexSz.cx;
                if (hasNestedIndex)
                    DrawSequenceCaret(obj.slots[1].children, yIndex, caretX, state.activeNodePath);
                else
                    DrawActiveCaret(caretX + MeasureDisplayText(hdc, part2).cx, yIndex);
            }
            else
            {
                if (hasNestedRadicand)
                    DrawSequenceCaret(obj.slots[0].children, yMid, xExprStart, state.activeNodePath);
                else
                    DrawActiveCaret(xExprStart + MeasureDisplayText(hdc, part1).cx, yMid);
            }
        }
    }
    else if (obj.type == MathType::AbsoluteValue)
    {
        // Measure the expression text to determine bar size
        SelectObject(hdc, renderBaseFont);
        TEXTMETRICW tmExpr = {};
        GetTextMetricsW(hdc, &tmExpr);

        // Get expression text size (part1 is the expression inside bars)
        const bool hasNestedExpr = !obj.slots.empty() && !obj.slots[0].children.empty();
        SIZE exprSz = {};
        if (hasNestedExpr)
            exprSz = MeasureMathNodeSequence(hdc, obj.slots[0].children, tmBase);
        else {
            const wchar_t* exprText = part1.empty() ? L"?" : part1.c_str();
            int exprLen = part1.empty() ? 1 : (int)part1.size();
            GetTextExtentPoint32W(hdc, exprText, exprLen, &exprSz);
        }

        // Layout constants
        const int pad = (std::max)(2, (int)(tmBase.tmHeight / 6));
        const int barExtend = (std::max)(3, (int)(tmBase.tmHeight / 5));  // how far bars extend above/below text
        const int penWidth = (std::max)(2, (int)(1.8 * renderScale));

        // Vertical extents of the bars
        int barTop = yMid - tmExpr.tmAscent - barExtend;
        int barBot = yMid + tmExpr.tmDescent + barExtend;

        // X positions
        int xLeftBar = ptStart.x + pad;
        int xExprStart = xLeftBar + pad + penWidth;
        int xRightBar = xExprStart + exprSz.cx + pad;
        int xEnd = xRightBar + penWidth + pad;

        // Draw the left vertical bar
        HPEN hPen = CreatePen(PS_SOLID, penWidth, normalColor);
        HPEN oldPen = (HPEN)SelectObject(hdc, hPen);

        MoveToEx(hdc, xLeftBar, barTop, NULL);
        LineTo(hdc, xLeftBar, barBot);

        // Draw the right vertical bar
        MoveToEx(hdc, xRightBar, barTop, NULL);
        LineTo(hdc, xRightBar, barBot);

        SelectObject(hdc, oldPen);
        DeleteObject(hPen);

        // Draw the expression between the bars
        SetTextAlign(hdc, TA_BASELINE | TA_LEFT);
        SelectObject(hdc, renderBaseFont);
        if (hasNestedExpr)
        {
            const COLORREF exprColor = (state.active && state.objectIndex == objIndex && state.activePart == 1) ? activeColor : normalColor;
            DrawMathNodeSequence(hdc, obj.slots[0].children, xExprStart, yMid, tmBase, exprColor);
        }
        else DrawPart(part1, xExprStart, yMid, 1);

        // Draw result right after the right bar if present
        if (!obj.resultText.empty()) {
            SetTextColor(hdc, activeColor);
            HFONT boldFont = CreateScaledFont(baseFont, renderScale, 100);
            LOGFONTW lfB = {}; GetObjectW(boldFont, sizeof(lfB), &lfB);
            lfB.lfWeight = FW_BOLD;
            DeleteObject(boldFont);
            boldFont = CreateFontIndirectW(&lfB);
            HFONT prevF = (HFONT)SelectObject(hdc, boldFont);
            TextOutW(hdc, xEnd + 4, yMid, obj.resultText.c_str(), (int)obj.resultText.size());
            SelectObject(hdc, prevF);
            DeleteObject(boldFont);
        }

        if (state.active && state.objectIndex == objIndex && state.activePart == 1)
        {
            if (hasNestedExpr)
                DrawSequenceCaret(obj.slots[0].children, yMid, xExprStart, state.activeNodePath);
            else
                DrawActiveCaret(xExprStart + MeasureDisplayText(hdc, part1).cx, yMid);
        }
    }
    else if (obj.type == MathType::Power)
    {
        SetTextAlign(hdc, TA_BASELINE | TA_LEFT);
        SelectObject(hdc, renderBaseFont);
        const bool hasNestedBase = !obj.slots.empty() && !obj.slots[0].children.empty();
        const bool hasNestedExponent = obj.slots.size() > 1 && !obj.slots[1].children.empty();

        SIZE baseSz = {};
        if (hasNestedBase)
            baseSz = MeasureMathNodeSequence(hdc, obj.slots[0].children, tmBase);
        else {
            const wchar_t* baseText = part1.empty() ? L"?" : part1.c_str();
            int baseLen = part1.empty() ? 1 : (int)part1.size();
            GetTextExtentPoint32W(hdc, baseText, baseLen, &baseSz);
        }

        int xBase = ptStart.x + 2;
        int yBase = yMid;
        if (hasNestedBase)
        {
            const COLORREF baseColor = (state.active && state.objectIndex == objIndex && state.activePart == 1) ? activeColor : normalColor;
            DrawMathNodeSequence(hdc, obj.slots[0].children, xBase, yBase, tmBase, baseColor);
        }
        else DrawPart(part1, xBase, yBase, 1);

        SelectObject(hdc, limitFont);
        TEXTMETRICW tmExp = {};
        GetTextMetricsW(hdc, &tmExp);

        SIZE expSz = {};
        if (hasNestedExponent)
            expSz = MeasureMathNodeSequence(hdc, obj.slots[1].children, tmBase);
        else {
            const wchar_t* expText = part2.empty() ? L"?" : part2.c_str();
            int expLen = part2.empty() ? 1 : (int)part2.size();
            GetTextExtentPoint32W(hdc, expText, expLen, &expSz);
        }

        int xExp = xBase + baseSz.cx + 2;
        int yExp = yBase - (tmBase.tmAscent / 2);
        if (hasNestedExponent)
        {
            const COLORREF exponentColor = (state.active && state.objectIndex == objIndex && state.activePart == 2) ? activeColor : normalColor;
            DrawMathNodeSequence(hdc, obj.slots[1].children, xExp, yExp, tmBase, exponentColor);
        }
        else DrawPart(part2, xExp, yExp, 2);

        if (!obj.resultText.empty()) {
            SetTextColor(hdc, activeColor);
            HFONT boldFont = CreateScaledFont(baseFont, renderScale, 100);
            LOGFONTW lfB = {}; GetObjectW(boldFont, sizeof(lfB), &lfB);
            lfB.lfWeight = FW_BOLD;
            DeleteObject(boldFont);
            boldFont = CreateFontIndirectW(&lfB);
            HFONT prevF = (HFONT)SelectObject(hdc, boldFont);
            TextOutW(hdc, xExp + expSz.cx + 8, yBase, obj.resultText.c_str(), (int)obj.resultText.size());
            SelectObject(hdc, prevF);
            DeleteObject(boldFont);
        }

        if (state.active && state.objectIndex == objIndex && (state.activePart == 1 || state.activePart == 2))
        {
            if (state.activePart == 1)
            {
                if (hasNestedBase)
                    DrawSequenceCaret(obj.slots[0].children, yBase, xBase, state.activeNodePath);
                else
                    DrawActiveCaret(xBase + MeasureDisplayText(hdc, part1).cx, yBase);
            }
            else
            {
                if (hasNestedExponent)
                    DrawSequenceCaret(obj.slots[1].children, yExp, xExp, state.activeNodePath);
                else
                    DrawActiveCaret(xExp + MeasureDisplayText(hdc, part2).cx, yExp);
            }
        }
    }
    else if (obj.type == MathType::Logarithm)
    {
        SetTextAlign(hdc, TA_BASELINE | TA_LEFT);
        SelectObject(hdc, renderBaseFont);
        const bool hasNestedBase = !obj.slots.empty() && !obj.slots[0].children.empty();
        const bool hasNestedArg = obj.slots.size() > 1 && !obj.slots[1].children.empty();

        // Draw "log" text
        const wchar_t* logText = L"log";
        SIZE logSz = {};
        GetTextExtentPoint32W(hdc, logText, 3, &logSz);
        int xLog = ptStart.x + 2;
        int yLog = yMid;
        SetTextColor(hdc, normalColor);
        TextOutW(hdc, xLog, yLog, logText, 3);

        // Draw base as subscript (part1)
        SelectObject(hdc, limitFont);
        TEXTMETRICW tmSub = {};
        GetTextMetricsW(hdc, &tmSub);

        SIZE baseSz = {};
        if (hasNestedBase)
            baseSz = MeasureMathNodeSequence(hdc, obj.slots[0].children, tmBase);
        else {
            const wchar_t* baseText = part1.empty() ? L"?" : part1.c_str();
            int baseLen = part1.empty() ? 1 : (int)part1.size();
            GetTextExtentPoint32W(hdc, baseText, baseLen, &baseSz);
        }

        int xBase = xLog + logSz.cx + 1;
        int yBase = yLog + (tmBase.tmDescent / 2);  // Subscript position
        if (hasNestedBase)
        {
            const COLORREF baseColor = (state.active && state.objectIndex == objIndex && state.activePart == 1) ? activeColor : normalColor;
            DrawMathNodeSequence(hdc, obj.slots[0].children, xBase, yBase, tmBase, baseColor);
        }
        else DrawPart(part1, xBase, yBase, 1);

        // Draw argument (part2) after the base
        SelectObject(hdc, renderBaseFont);
        int xArg = xBase + baseSz.cx + 2;
        if (hasNestedArg)
        {
            const COLORREF argColor = (state.active && state.objectIndex == objIndex && state.activePart == 2) ? activeColor : normalColor;
            DrawMathNodeSequence(hdc, obj.slots[1].children, xArg, yMid, tmBase, argColor);
        }
        else DrawPart(part2, xArg, yMid, 2);

        // Draw result if present
        if (!obj.resultText.empty()) {
            SIZE argSz = {};
            if (hasNestedArg)
                argSz = MeasureMathNodeSequence(hdc, obj.slots[1].children, tmBase);
            else
                GetTextExtentPoint32W(hdc, part2.empty() ? L"?" : part2.c_str(), 
                                      part2.empty() ? 1 : (int)part2.size(), &argSz);
            SetTextColor(hdc, activeColor);
            HFONT boldFont = CreateScaledFont(baseFont, renderScale, 100);
            LOGFONTW lfB = {}; GetObjectW(boldFont, sizeof(lfB), &lfB);
            lfB.lfWeight = FW_BOLD;
            DeleteObject(boldFont);
            boldFont = CreateFontIndirectW(&lfB);
            HFONT prevF = (HFONT)SelectObject(hdc, boldFont);
            TextOutW(hdc, xArg + argSz.cx + 6, yMid, obj.resultText.c_str(), (int)obj.resultText.size());
            SelectObject(hdc, prevF);
            DeleteObject(boldFont);
        }

        if (state.active && state.objectIndex == objIndex && (state.activePart == 1 || state.activePart == 2))
        {
            if (state.activePart == 1)
            {
                if (hasNestedBase)
                    DrawSequenceCaret(obj.slots[0].children, yBase, xBase, state.activeNodePath);
                else
                    DrawActiveCaret(xBase + MeasureDisplayText(hdc, part1).cx, yBase);
            }
            else
            {
                if (hasNestedArg)
                    DrawSequenceCaret(obj.slots[1].children, yMid, xArg, state.activeNodePath);
                else
                    DrawActiveCaret(xArg + MeasureDisplayText(hdc, part2).cx, yMid);
            }
        }
    }
    RestoreDC(hdc, saved);
    DeleteObject(renderBaseFont); DeleteObject(limitFont);
}

bool MathRenderer::GetHitPart(HWND hEdit, HDC hdc, POINT ptMouse, size_t* outIndex, int* outPart, std::vector<size_t>* outNodePath)
{
    auto& objects = MathManager::Get().GetObjects();
    for (size_t i = 0; i < objects.size(); ++i)
    {
        const auto& obj = objects[i];
        const std::wstring& part1 = obj.SlotText(1);
        const std::wstring& part2 = obj.SlotText(2);
        const std::wstring& part3 = obj.SlotText(3);
        const std::wstring& part4 = obj.SlotText(4);
        POINT ptS = {}, ptE = {};
        if (!TryGetCharPos(hEdit, obj.barStart, ptS)) continue;
        if (!TryGetCharPos(hEdit, obj.barStart + std::max<LONG>(0, obj.barLen - 1), ptE)) continue;

        HFONT baseF = (HFONT)SendMessage(hEdit, WM_GETFONT, 0, 0);
        if (!baseF) baseF = (HFONT)GetStockObject(DEFAULT_GUI_FONT);

        const double scale = ComputeRenderScale(hEdit, hdc, obj, baseF);
        HFONT limitF = CreateScaledFont(baseF, scale, 70);
        HFONT baseRF = CreateScaledFont(baseF, scale, 100);

        TEXTMETRICW tmB = {};
        HFONT old = (HFONT)SelectObject(hdc, baseRF);
        GetTextMetricsW(hdc, &tmB);

        const int bW = (ptE.x - ptS.x) + (int)(tmB.tmAveCharWidth * scale);
        const int xC = ptS.x + (bW / 2);
        const int yM = ptS.y + tmB.tmAscent;
        const int gap = std::max<int>(2, (int)(tmB.tmHeight / 10));

        bool hit = false;
        std::vector<size_t> nodePath;
        auto Check = [&](RECT rc, int pIdx) {
            if (PtInRect(&rc, ptMouse)) {
                *outIndex = i;
                *outPart = pIdx;
                if (outNodePath) outNodePath->clear();
                return true;
            }
            return false;
        };

        if (obj.type == MathType::Fraction)
        {
            if (!obj.slots.empty() && !obj.slots[0].children.empty())
            {
                SIZE numeratorSz = MeasureMathNodeSequence(hdc, obj.slots[0].children, tmB);
                const int numeratorX = xC - numeratorSz.cx / 2;
                if (HitTestMathNodeSequence(hdc, obj.slots[0].children, numeratorX, yM - gap - tmB.tmDescent, tmB, ptMouse, {}, nodePath)) {
                    *outIndex = i; *outPart = 1; if (outNodePath) *outNodePath = nodePath; hit = true;
                }
            }
            if (!hit && obj.slots.size() > 1 && !obj.slots[1].children.empty())
            {
                SIZE denominatorSz = MeasureMathNodeSequence(hdc, obj.slots[1].children, tmB);
                const int denominatorX = xC - denominatorSz.cx / 2;
                if (HitTestMathNodeSequence(hdc, obj.slots[1].children, denominatorX, yM + gap + tmB.tmAscent, tmB, ptMouse, {}, nodePath)) {
                    *outIndex = i; *outPart = 2; if (outNodePath) *outNodePath = nodePath; hit = true;
                }
            }
            if (hit) {
                SelectObject(hdc, old); DeleteObject(baseRF); DeleteObject(limitF); return true;
            }
            SelectObject(hdc, limitF); SIZE sz = {};
            GetTextExtentPoint32W(hdc, part1.c_str(), (int)part1.size(), &sz);
            if (Check({ xC - sz.cx/2 - 10, yM - (int)(tmB.tmAscent*0.4) - sz.cy - 10, xC + sz.cx/2 + 10, yM - (int)(tmB.tmAscent*0.4) + 5 }, 1)) hit = true;
            if (!hit) {
                GetTextExtentPoint32W(hdc, part2.c_str(), (int)part2.size(), &sz);
                if (Check({ xC - sz.cx/2 - 10, yM + (int)(tmB.tmDescent*0.4) - 5, xC + sz.cx/2 + 10, yM + (int)(tmB.tmDescent*0.4) + sz.cy + 10 }, 2)) hit = true;
            }
        }
        else if (obj.type == MathType::Summation)
        {
            if (obj.slots.size() > 2 && !obj.slots[2].children.empty())
            {
                if (HitTestMathNodeSequence(hdc, obj.slots[2].children, ptE.x + 4, yM, tmB, ptMouse, {}, nodePath)) {
                    *outIndex = i; *outPart = 3; if (outNodePath) *outNodePath = nodePath; hit = true;
                }
            }
            if (hit) { SelectObject(hdc, old); DeleteObject(baseRF); DeleteObject(limitF); return true; }
            const bool hasNestedUpper = obj.slots.size() > 0 && SequenceHasVisibleContent(obj.slots[0].children);
            const bool hasNestedLower = obj.slots.size() > 1 && SequenceHasVisibleContent(obj.slots[1].children);
            SelectObject(hdc, limitF); SIZE sz = {};
            if (hasNestedUpper)
            {
                const NodeMetrics upper = MeasureMathSequenceMetrics(hdc, obj.slots[0].children, tmB);
                const int upperX = xC - upper.cx / 2;
                const int upperBaseline = yM - tmB.tmAscent - 2;
                if (HitTestMathNodeSequence(hdc, obj.slots[0].children, upperX, upperBaseline, tmB, ptMouse, {}, nodePath)) {
                    *outIndex = i; *outPart = 1; if (outNodePath) *outNodePath = nodePath; hit = true;
                }
            }
            if (!hit) {
                GetTextExtentPoint32W(hdc, part1.c_str(), (int)part1.size(), &sz);
                if (Check({ xC - sz.cx/2 - 10, yM - tmB.tmAscent - sz.cy - 10, xC + sz.cx/2 + 10, yM - tmB.tmAscent + 5 }, 1)) hit = true;
            }
            if (!hit) {
                if (hasNestedLower)
                {
                    const NodeMetrics lower = MeasureMathSequenceMetrics(hdc, obj.slots[1].children, tmB);
                    const int lowerX = xC - lower.cx / 2;
                    const int lowerBaseline = yM + tmB.tmDescent + tmB.tmAscent + 2;
                    if (HitTestMathNodeSequence(hdc, obj.slots[1].children, lowerX, lowerBaseline, tmB, ptMouse, {}, nodePath)) {
                        *outIndex = i; *outPart = 2; if (outNodePath) *outNodePath = nodePath; hit = true;
                    }
                }
            }
            if (!hit) {
                GetTextExtentPoint32W(hdc, part2.c_str(), (int)part2.size(), &sz);
                if (Check({ xC - sz.cx/2 - 10, yM + tmB.tmDescent - 5, xC + sz.cx/2 + 10, yM + tmB.tmDescent + sz.cy + 10 }, 2)) hit = true;
            }
            if (!hit) {
                SelectObject(hdc, baseRF);
                GetTextExtentPoint32W(hdc, part3.empty() ? L"?" : part3.c_str(), (int)std::max<size_t>(1, part3.size()), &sz);
                if (Check({ ptE.x + 2, yM - tmB.tmAscent, ptE.x + sz.cx + 20, yM + tmB.tmDescent }, 3)) hit = true;
            }
        }
        else if (obj.type == MathType::Product)
        {
            const bool hasNestedUpper = obj.slots.size() > 0 && SequenceHasVisibleContent(obj.slots[0].children);
            const bool hasNestedLower = obj.slots.size() > 1 && SequenceHasVisibleContent(obj.slots[1].children);
            if (obj.slots.size() > 2 && !obj.slots[2].children.empty())
            {
                if (HitTestMathNodeSequence(hdc, obj.slots[2].children, ptE.x + 4, yM, tmB, ptMouse, {}, nodePath)) {
                    *outIndex = i; *outPart = 3; if (outNodePath) *outNodePath = nodePath; hit = true;
                }
            }
            if (hit) { SelectObject(hdc, old); DeleteObject(baseRF); DeleteObject(limitF); return true; }
            // Hit area for product symbol (Π) similar to summation
            SelectObject(hdc, limitF); SIZE sz = {};
            if (hasNestedUpper)
            {
                const NodeMetrics upper = MeasureMathSequenceMetrics(hdc, obj.slots[0].children, tmB);
                const int upperX = xC - upper.cx / 2;
                const int upperBaseline = yM - tmB.tmAscent - 2;
                if (HitTestMathNodeSequence(hdc, obj.slots[0].children, upperX, upperBaseline, tmB, ptMouse, {}, nodePath)) {
                    *outIndex = i; *outPart = 1; if (outNodePath) *outNodePath = nodePath; hit = true;
                }
            }
            if (!hit) {
                GetTextExtentPoint32W(hdc, part1.c_str(), (int)part1.size(), &sz);
                if (Check({ xC - sz.cx/2 - 10, yM - tmB.tmAscent - sz.cy - 10, xC + sz.cx/2 + 10, yM - tmB.tmAscent + 5 }, 1)) hit = true;
            }
            if (!hit) {
                if (hasNestedLower)
                {
                    const NodeMetrics lower = MeasureMathSequenceMetrics(hdc, obj.slots[1].children, tmB);
                    const int lowerX = xC - lower.cx / 2;
                    const int lowerBaseline = yM + tmB.tmDescent + tmB.tmAscent + 2;
                    if (HitTestMathNodeSequence(hdc, obj.slots[1].children, lowerX, lowerBaseline, tmB, ptMouse, {}, nodePath)) {
                        *outIndex = i; *outPart = 2; if (outNodePath) *outNodePath = nodePath; hit = true;
                    }
                }
            }
            if (!hit) {
                GetTextExtentPoint32W(hdc, part2.c_str(), (int)part2.size(), &sz);
                if (Check({ xC - sz.cx/2 - 10, yM + tmB.tmDescent - 5, xC + sz.cx/2 + 10, yM + tmB.tmDescent + sz.cy + 10 }, 2)) hit = true;
            }
            if (!hit) {
                SelectObject(hdc, baseRF);
                GetTextExtentPoint32W(hdc, part3.empty() ? L"?" : part3.c_str(), (int)std::max<size_t>(1, part3.size()), &sz);
                if (Check({ ptE.x + 2, yM - tmB.tmAscent, ptE.x + sz.cx + 20, yM + tmB.tmDescent }, 3)) hit = true;
            }
        }
        else if (obj.type == MathType::Integral)
        {
            const bool hasNestedUpper = obj.slots.size() > 0 && SequenceHasVisibleContent(obj.slots[0].children);
            const bool hasNestedLower = obj.slots.size() > 1 && SequenceHasVisibleContent(obj.slots[1].children);
            if (obj.slots.size() > 2 && !obj.slots[2].children.empty())
            {
                if (HitTestMathNodeSequence(hdc, obj.slots[2].children, ptE.x + 6, yM, tmB, ptMouse, {}, nodePath)) {
                    *outIndex = i; *outPart = 3; if (outNodePath) *outNodePath = nodePath; hit = true;
                }
            }
            if (hit) { SelectObject(hdc, old); DeleteObject(baseRF); DeleteObject(limitF); return true; }
            SelectObject(hdc, limitF); SIZE sz = {};
            if (hasNestedUpper)
            {
                if (HitTestMathNodeSequence(hdc, obj.slots[0].children, ptE.x - 2, yM - tmB.tmAscent + (int)(tmB.tmAscent * 0.2), tmB, ptMouse, {}, nodePath)) {
                    *outIndex = i; *outPart = 1; if (outNodePath) *outNodePath = nodePath; hit = true;
                }
            }
            if (!hit) {
                GetTextExtentPoint32W(hdc, part1.c_str(), (int)part1.size(), &sz);
                if (Check({ ptE.x - 2, yM - tmB.tmAscent - 5, ptE.x + sz.cx + 10, yM - tmB.tmAscent + sz.cy + 5 }, 1)) hit = true;
            }
            if (!hit) {
                if (hasNestedLower)
                {
                    if (HitTestMathNodeSequence(hdc, obj.slots[1].children, ptE.x - 8, yM + tmB.tmDescent + 2, tmB, ptMouse, {}, nodePath)) {
                        *outIndex = i; *outPart = 2; if (outNodePath) *outNodePath = nodePath; hit = true;
                    }
                }
            }
            if (!hit) {
                GetTextExtentPoint32W(hdc, part2.c_str(), (int)part2.size(), &sz);
                if (Check({ ptE.x - 10, yM + tmB.tmDescent - 5, ptE.x + sz.cx + 5, yM + tmB.tmDescent + sz.cy + 5 }, 2)) hit = true;
            }
            if (!hit) {
                SelectObject(hdc, baseRF);
                GetTextExtentPoint32W(hdc, part3.empty() ? L"?" : part3.c_str(), (int)std::max<size_t>(1, part3.size()), &sz);
                if (Check({ ptE.x + 4, yM - tmB.tmAscent, ptE.x + sz.cx + 20, yM + tmB.tmDescent }, 3)) hit = true;
            }
        }
        else if (obj.type == MathType::SystemOfEquations)
        {
            // Compute brace width matching the renderer
            int braceW = (int)(tmB.tmAveCharWidth * 1.2);
            if (braceW < 10) braceW = 10;

            SelectObject(hdc, baseRF);
            TEXTMETRICW tmEq = {};
            GetTextMetricsW(hdc, &tmEq);
            int lineH = tmEq.tmHeight + 4;
            int eqCount = 3;
            const bool hasNestedEq1 = obj.slots.size() > 0 && SequenceHasVisibleContent(obj.slots[0].children);
            const bool hasNestedEq2 = obj.slots.size() > 1 && SequenceHasVisibleContent(obj.slots[1].children);
            const bool hasNestedEq3 = obj.slots.size() > 2 && SequenceHasVisibleContent(obj.slots[2].children);
            if (!hasNestedEq3 && part3.empty()) eqCount = 2;
            int totalH = lineH * eqCount;
            int firstBaseline = yM - totalH / 2 + tmEq.tmAscent;
            int eqX = ptS.x + braceW + 6;

            SIZE sz = {};
            if (hasNestedEq1)
            {
                if (HitTestMathNodeSequence(hdc, obj.slots[0].children, eqX, firstBaseline, tmB, ptMouse, {}, nodePath)) {
                    *outIndex = i; *outPart = 1; if (outNodePath) *outNodePath = nodePath; hit = true;
                }
            }
            if (!hit) {
                GetTextExtentPoint32W(hdc, part1.empty() ? L"?" : part1.c_str(), (int)std::max<size_t>(1, part1.size()), &sz);
                if (Check({ eqX - 5, firstBaseline - tmEq.tmAscent - 2, eqX + (std::max)(sz.cx, (LONG)40) + 10, firstBaseline + tmEq.tmDescent + 2 }, 1)) hit = true;
            }
            // Eq 2 hit area
            if (!hit) {
                if (hasNestedEq2)
                {
                    if (HitTestMathNodeSequence(hdc, obj.slots[1].children, eqX, firstBaseline + lineH, tmB, ptMouse, {}, nodePath)) {
                        *outIndex = i; *outPart = 2; if (outNodePath) *outNodePath = nodePath; hit = true;
                    }
                }
            }
            if (!hit) {
                GetTextExtentPoint32W(hdc, part2.empty() ? L"?" : part2.c_str(), (int)std::max<size_t>(1, part2.size()), &sz);
                if (Check({ eqX - 5, firstBaseline + lineH - tmEq.tmAscent - 2, eqX + (std::max)(sz.cx, (LONG)40) + 10, firstBaseline + lineH + tmEq.tmDescent + 2 }, 2)) hit = true;
            }
            // Eq 3 hit area
            if (!hit && eqCount >= 3) {
                if (hasNestedEq3)
                {
                    if (HitTestMathNodeSequence(hdc, obj.slots[2].children, eqX, firstBaseline + lineH * 2, tmB, ptMouse, {}, nodePath)) {
                        *outIndex = i; *outPart = 3; if (outNodePath) *outNodePath = nodePath; hit = true;
                    }
                }
            }
            if (!hit && eqCount >= 3) {
                GetTextExtentPoint32W(hdc, part3.empty() ? L"?" : part3.c_str(), (int)std::max<size_t>(1, part3.size()), &sz);
                if (Check({ eqX - 5, firstBaseline + lineH * 2 - tmEq.tmAscent - 2, eqX + (std::max)(sz.cx, (LONG)40) + 10, firstBaseline + lineH * 2 + tmEq.tmDescent + 2 }, 3)) hit = true;
            }
        }
        else if (obj.type == MathType::Matrix || obj.type == MathType::Determinant)
        {
            SelectObject(hdc, baseRF);
            TEXTMETRICW tmRow = {};
            GetTextMetricsW(hdc, &tmRow);
            const int cellPadX = (std::max)(6, (int)tmRow.tmAveCharWidth / 2);
            const int cellPadY = (std::max)(4, (int)tmRow.tmHeight / 6);
            const int colGap = (std::max)(12, (int)tmRow.tmAveCharWidth);
            const int rowGap = (std::max)(8, (int)tmRow.tmHeight / 4);
            int bracketInset = (std::max)(8, (int)(tmB.tmAveCharWidth * 0.8));
            const bool hasNestedCells[] = {
                !obj.slots.empty() && !obj.slots[0].children.empty(),
                obj.slots.size() > 1 && !obj.slots[1].children.empty(),
                obj.slots.size() > 2 && !obj.slots[2].children.empty(),
                obj.slots.size() > 3 && !obj.slots[3].children.empty()
            };
            const std::wstring* cellParts[] = { &part1, &part2, &part3, &part4 };
            SIZE cellSz[4] = {};
            for (int cellIndex = 0; cellIndex < 4; ++cellIndex)
            {
                if (hasNestedCells[cellIndex])
                    cellSz[cellIndex] = MeasureMathNodeSequence(hdc, obj.slots[(size_t)cellIndex].children, tmB);
                else
                    GetTextExtentPoint32W(hdc,
                        cellParts[cellIndex]->empty() ? L"?" : cellParts[cellIndex]->c_str(),
                        (int)(cellParts[cellIndex]->empty() ? 1 : cellParts[cellIndex]->size()),
                        &cellSz[cellIndex]);
            }

            const int colWidths[2] = {
                (std::max)(cellSz[0].cx, cellSz[2].cx),
                (std::max)(cellSz[1].cx, cellSz[3].cx)
            };
            const int rowHeights[2] = {
                (std::max)(cellSz[0].cy, cellSz[1].cy),
                (std::max)(cellSz[2].cy, cellSz[3].cy)
            };
            const int totalH = rowHeights[0] + rowHeights[1] + rowGap + cellPadY * 4;
            const int yTop = yM - totalH / 2;
            const int contentX = ptS.x + bracketInset + 10;
            const int colLeft[2] = {
                contentX + cellPadX,
                contentX + cellPadX * 3 + colWidths[0] + colGap
            };
            const int rowTop[2] = {
                yTop + cellPadY,
                yTop + cellPadY * 3 + rowHeights[0] + rowGap
            };

            auto CheckMatrixCell = [&](int partIndex, int rowIndex, int colIndex) {
                const int cellIndex = partIndex - 1;
                const int drawX = colLeft[colIndex] + (colWidths[colIndex] - cellSz[cellIndex].cx) / 2;
                const int drawY = rowTop[rowIndex] + tmRow.tmAscent;
                if (hasNestedCells[cellIndex])
                {
                    if (HitTestMathNodeSequence(hdc, obj.slots[(size_t)cellIndex].children, drawX, drawY, tmB, ptMouse, {}, nodePath)) {
                        *outIndex = i;
                        *outPart = partIndex;
                        if (outNodePath) *outNodePath = nodePath;
                        return true;
                    }
                }

                RECT rc = {
                    colLeft[colIndex] - cellPadX / 2,
                    rowTop[rowIndex] - cellPadY / 2,
                    colLeft[colIndex] + colWidths[colIndex] + cellPadX / 2,
                    rowTop[rowIndex] + rowHeights[rowIndex] + cellPadY / 2
                };
                return Check(rc, partIndex);
            };

            if (CheckMatrixCell(1, 0, 0)) hit = true;
            if (!hit && CheckMatrixCell(2, 0, 1)) hit = true;
            if (!hit && CheckMatrixCell(3, 1, 0)) hit = true;
            if (!hit && CheckMatrixCell(4, 1, 1)) hit = true;
        }
        else if (obj.type == MathType::SquareRoot)
        {
            if (!obj.slots.empty() && !obj.slots[0].children.empty())
            {
                SIZE exprSz = MeasureMathNodeSequence(hdc, obj.slots[0].children, tmB);
                SelectObject(hdc, limitF);
                TEXTMETRICW tmIdx = {};
                GetTextMetricsW(hdc, &tmIdx);
                SIZE indexSz = {};
                const bool showIndex = SlotHasVisibleContent(obj, 2);
                if (obj.slots.size() > 1 && !obj.slots[1].children.empty())
                    indexSz = MeasureMathNodeSequence(hdc, obj.slots[1].children, tmB);
                else if (showIndex)
                    GetTextExtentPoint32W(hdc, part2.empty() ? L"?" : part2.c_str(), part2.empty() ? 1 : (int)part2.size(), &indexSz);
                int pad = std::max<int>(2, (int)(tmB.tmHeight / 8));
                int radicalW = (int)(tmB.tmAveCharWidth * 1.0);
                int indexGap = std::max<int>(2, (int)(tmB.tmAveCharWidth / 2));
                int xRadStart = ptS.x + ((showIndex && indexSz.cx > 0) ? indexSz.cx + indexGap : 0);
                int xExprStart = xRadStart + radicalW + pad;
                int overlineGap = std::max<int>(2, (int)(tmB.tmHeight / 10));
                int radTop = yM - tmB.tmAscent - overlineGap - 2;
                if (obj.slots.size() > 1 && !obj.slots[1].children.empty() && HitTestMathNodeSequence(hdc, obj.slots[1].children, xRadStart - indexGap - indexSz.cx, radTop + tmB.tmAscent, tmB, ptMouse, {}, nodePath)) {
                    *outIndex = i; *outPart = 2; if (outNodePath) *outNodePath = nodePath; hit = true;
                }
                if (!hit && HitTestMathNodeSequence(hdc, obj.slots[0].children, xExprStart, yM, tmB, ptMouse, {}, nodePath)) {
                    *outIndex = i; *outPart = 1; if (outNodePath) *outNodePath = nodePath; hit = true;
                }
            }
            if (hit) { SelectObject(hdc, old); DeleteObject(baseRF); DeleteObject(limitF); return true; }
            SelectObject(hdc, baseRF);
            TEXTMETRICW tmExpr = {};
            GetTextMetricsW(hdc, &tmExpr);

            SIZE sz = {};
            GetTextExtentPoint32W(hdc, part1.empty() ? L"?" : part1.c_str(), (int)std::max<size_t>(1, part1.size()), &sz);

            SelectObject(hdc, limitF);
            TEXTMETRICW tmIdx = {};
            GetTextMetricsW(hdc, &tmIdx);
            SIZE indexSz = {};
            const bool showIndex = SlotHasVisibleContent(obj, 2);
            if (showIndex)
                GetTextExtentPoint32W(hdc, part2.c_str(), (int)part2.size(), &indexSz);

            int pad = std::max<int>(2, (int)(tmB.tmHeight / 8));
            int radicalW = (int)(tmB.tmAveCharWidth * 1.0);
            int indexGap = std::max<int>(2, (int)(tmB.tmAveCharWidth / 2));
            int xRadStart = ptS.x + (showIndex ? indexSz.cx + indexGap : 0);
            int xExprStart = xRadStart + radicalW + pad;
            int overlineGap = std::max<int>(2, (int)(tmB.tmHeight / 10));
            int radTop = yM - tmExpr.tmAscent - overlineGap - 2;
            int yIndex = radTop + tmIdx.tmAscent + overlineGap;

            if (showIndex) {
                RECT rcIndex = { xRadStart - indexSz.cx - indexGap - 4, yIndex - tmIdx.tmAscent - 4, xRadStart - 1, yIndex + tmIdx.tmDescent + 4 };
                if (Check(rcIndex, 2)) hit = true;
            }

            if (!hit) {
                RECT rcExpr = { xRadStart, radTop - 5, xExprStart + sz.cx + pad + 10, yM + tmExpr.tmDescent + pad + 5 };
                if (Check(rcExpr, 1)) hit = true;
            }
        }
        else if (obj.type == MathType::AbsoluteValue)
        {
            if (!obj.slots.empty() && !obj.slots[0].children.empty())
            {
                SelectObject(hdc, baseRF);
                if (HitTestMathNodeSequence(hdc, obj.slots[0].children, ptS.x + std::max<int>(2, (int)(tmB.tmHeight / 6)) * 3, yM, tmB, ptMouse, {}, nodePath)) {
                    *outIndex = i; *outPart = 1; if (outNodePath) *outNodePath = nodePath; hit = true;
                }
            }
            if (hit) { SelectObject(hdc, old); DeleteObject(baseRF); DeleteObject(limitF); return true; }
            SelectObject(hdc, baseRF);
            TEXTMETRICW tmExpr = {};
            GetTextMetricsW(hdc, &tmExpr);

            SIZE sz = {};
            GetTextExtentPoint32W(hdc, part1.empty() ? L"?" : part1.c_str(), (int)std::max<size_t>(1, part1.size()), &sz);

            int pad = std::max<int>(2, (int)(tmB.tmHeight / 6));
            int barExtend = std::max<int>(3, (int)(tmB.tmHeight / 5));
            int barTop = yM - tmExpr.tmAscent - barExtend;
            int barBot = yM + tmExpr.tmDescent + barExtend;
            int xLeftBar = ptS.x + pad;
            int xRightBar = xLeftBar + pad + 2 + sz.cx + pad;

            // Hit area covers the entire absolute value area
            RECT rcExpr = { xLeftBar - 5, barTop - 5, xRightBar + 10, barBot + 5 };
            if (Check(rcExpr, 1)) hit = true;
        }
        else if (obj.type == MathType::Power)
        {
            if (!obj.slots.empty() && !obj.slots[0].children.empty())
            {
                const int xBase = ptS.x + 2;
                if (HitTestMathNodeSequence(hdc, obj.slots[0].children, xBase, yM, tmB, ptMouse, {}, nodePath)) {
                    *outIndex = i; *outPart = 1; if (outNodePath) *outNodePath = nodePath; hit = true;
                }
            }
            if (!hit && obj.slots.size() > 1 && !obj.slots[1].children.empty())
            {
                SIZE baseSz = !obj.slots.empty() && !obj.slots[0].children.empty()
                    ? MeasureMathNodeSequence(hdc, obj.slots[0].children, tmB)
                    : SIZE{};
                if (baseSz.cx == 0)
                    GetTextExtentPoint32W(hdc, part1.empty() ? L"?" : part1.c_str(), (int)std::max<size_t>(1, part1.size()), &baseSz);
                const int xExp = ptS.x + 2 + baseSz.cx + 2;
                if (HitTestMathNodeSequence(hdc, obj.slots[1].children, xExp, yM - (tmB.tmAscent / 2), tmB, ptMouse, {}, nodePath)) {
                    *outIndex = i; *outPart = 2; if (outNodePath) *outNodePath = nodePath; hit = true;
                }
            }
            if (hit) { SelectObject(hdc, old); DeleteObject(baseRF); DeleteObject(limitF); return true; }
            SetTextAlign(hdc, TA_BASELINE | TA_LEFT);
            SelectObject(hdc, baseRF);

            SIZE baseSz = {};
            GetTextExtentPoint32W(hdc, part1.empty() ? L"?" : part1.c_str(), (int)std::max<size_t>(1, part1.size()), &baseSz);
            int xBase = ptS.x + 2;
            int yBase = yM;

            RECT rcBase = { xBase - 5, yBase - tmB.tmAscent - 5, xBase + baseSz.cx + 8, yBase + tmB.tmDescent + 5 };
            if (Check(rcBase, 1)) hit = true;

            if (!hit) {
                SelectObject(hdc, limitF);
                TEXTMETRICW tmExp = {};
                GetTextMetricsW(hdc, &tmExp);

                SIZE expSz = {};
                GetTextExtentPoint32W(hdc, part2.empty() ? L"?" : part2.c_str(), (int)std::max<size_t>(1, part2.size()), &expSz);
                int xExp = xBase + baseSz.cx + 2;
                int yExp = yBase - (tmB.tmAscent / 2);

                RECT rcExp = { xExp - 4, yExp - tmExp.tmAscent - 4, xExp + expSz.cx + 6, yExp + tmExp.tmDescent + 4 };
                if (Check(rcExp, 2)) hit = true;
            }
        }
        else if (obj.type == MathType::Logarithm)
        {
            if (!obj.slots.empty() && !obj.slots[0].children.empty())
            {
                SIZE logSz = {};
                GetTextExtentPoint32W(hdc, L"log", 3, &logSz);
                int xBase = ptS.x + 2 + logSz.cx + 1;
                if (HitTestMathNodeSequence(hdc, obj.slots[0].children, xBase, yM + (tmB.tmDescent / 2), tmB, ptMouse, {}, nodePath)) {
                    *outIndex = i; *outPart = 1; if (outNodePath) *outNodePath = nodePath; hit = true;
                }
            }
            if (!hit && obj.slots.size() > 1 && !obj.slots[1].children.empty())
            {
                SIZE logSz = {};
                GetTextExtentPoint32W(hdc, L"log", 3, &logSz);
                SIZE baseSz = !obj.slots.empty() && !obj.slots[0].children.empty()
                    ? MeasureMathNodeSequence(hdc, obj.slots[0].children, tmB)
                    : SIZE{};
                if (baseSz.cx == 0)
                    GetTextExtentPoint32W(hdc, part1.empty() ? L"?" : part1.c_str(), (int)std::max<size_t>(1, part1.size()), &baseSz);
                int xArg = ptS.x + 2 + logSz.cx + 1 + baseSz.cx + 2;
                if (HitTestMathNodeSequence(hdc, obj.slots[1].children, xArg, yM, tmB, ptMouse, {}, nodePath)) {
                    *outIndex = i; *outPart = 2; if (outNodePath) *outNodePath = nodePath; hit = true;
                }
            }
            if (hit) { SelectObject(hdc, old); DeleteObject(baseRF); DeleteObject(limitF); return true; }
            SetTextAlign(hdc, TA_BASELINE | TA_LEFT);
            SelectObject(hdc, baseRF);

            // "log" text size
            SIZE logSz = {};
            GetTextExtentPoint32W(hdc, L"log", 3, &logSz);
            int xLog = ptS.x + 2;

            // Base subscript hit area
            SelectObject(hdc, limitF);
            TEXTMETRICW tmSub = {};
            GetTextMetricsW(hdc, &tmSub);

            SIZE baseSz = {};
            GetTextExtentPoint32W(hdc, part1.empty() ? L"?" : part1.c_str(), (int)std::max<size_t>(1, part1.size()), &baseSz);
            int xBase = xLog + logSz.cx + 1;
            int yBase = yM + (tmB.tmDescent / 2);

            RECT rcBase = { xBase - 2, yBase - tmSub.tmAscent - 2, xBase + baseSz.cx + 4, yBase + tmSub.tmDescent + 2 };
            if (Check(rcBase, 1)) hit = true;

            // Argument hit area
            if (!hit) {
                SelectObject(hdc, baseRF);
                SIZE argSz = {};
                GetTextExtentPoint32W(hdc, part2.empty() ? L"?" : part2.c_str(), (int)std::max<size_t>(1, part2.size()), &argSz);
                int xArg = xBase + baseSz.cx + 2;

                RECT rcArg = { xArg - 4, yM - tmB.tmAscent - 4, xArg + argSz.cx + 8, yM + tmB.tmDescent + 4 };
                if (Check(rcArg, 2)) hit = true;
            }
        }
        SelectObject(hdc, old); DeleteObject(baseRF); DeleteObject(limitF);
        if (hit) return true;
    }
    return false;
}
