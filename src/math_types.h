#pragma once

#include <windows.h>
#include <algorithm>
#include <string>
#include <vector>

// Top-level anchored objects use slots as their editing surface.
// Invariants:
// - `slots` are the source of truth for active code paths.
// - `part1/part2/part3` are compatibility mirrors only and must be refreshed from slots.
// - Prefer slot-oriented helpers (`SlotText`, `EditableSlotText`, `EditableLeafText`) in active code.
// - Structured nested notation lives in `MathSlot::children`.
// - Structural nodes store per-slot content in `children` as `Group` nodes.
// - Text content that is directly editable must terminate in a `Text` node so typing can append safely.

enum class MathType { Fraction, Summation, Integral, SystemOfEquations, SquareRoot, AbsoluteValue, Power, Logarithm, Sum, Product, Matrix, Determinant };

enum class MathNodeKind { Text, Group, SquareRoot, Fraction, Power, AbsoluteValue, Logarithm };

struct MathNode
{
    MathNodeKind kind = MathNodeKind::Text;
    std::wstring text;
    std::vector<MathNode> children;

    static MathNode MakeText(const std::wstring& value = L"")
    {
        MathNode node;
        node.kind = MathNodeKind::Text;
        node.text = value;
        return node;
    }

    static MathNode MakeGroup()
    {
        MathNode node;
        node.kind = MathNodeKind::Group;
        node.children.push_back(MakeText());
        return node;
    }

    static MathNode MakeStructured(MathNodeKind nodeKind, size_t slotCount)
    {
        MathNode node;
        node.kind = nodeKind;
        node.children.reserve(slotCount);
        for (size_t i = 0; i < slotCount; ++i)
            node.children.push_back(MakeGroup());
        return node;
    }

    bool IsStructural() const
    {
        return kind != MathNodeKind::Text && kind != MathNodeKind::Group;
    }

    static size_t SlotCountForKind(MathNodeKind nodeKind)
    {
        switch (nodeKind)
        {
        case MathNodeKind::SquareRoot:
            return 2;
        case MathNodeKind::Fraction:
            return 2;
        case MathNodeKind::Power:
            return 2;
        case MathNodeKind::AbsoluteValue:
            return 1;
        case MathNodeKind::Logarithm:
            return 2;
        default:
            return 0;
        }
    }

    void EnsureSlotCount(size_t count)
    {
        if (!IsStructural())
            return;

        while (children.size() < count)
            children.push_back(MakeGroup());
        for (auto& child : children)
        {
            if (child.kind != MathNodeKind::Group)
            {
                MathNode group = MakeGroup();
                if (child.kind == MathNodeKind::Text)
                    group.children[0].text = child.text;
                else
                    group.children = child.children;
                child = std::move(group);
            }
            if (child.children.empty())
                child.children.push_back(MakeText());
        }
    }

    std::vector<MathNode>& SlotNodes(size_t slotIndex)
    {
        EnsureSlotCount(slotIndex + 1);
        return children[slotIndex].children;
    }

    const std::vector<MathNode>& SlotNodes(size_t slotIndex) const
    {
        static const std::vector<MathNode> kEmpty;
        if (!IsStructural() || slotIndex >= children.size() || children[slotIndex].kind != MathNodeKind::Group)
            return kEmpty;
        return children[slotIndex].children;
    }
};

struct MathSlot
{
    std::wstring text;
    std::vector<MathNode> children;
};

struct MathObject
{
    MathType type = MathType::Fraction;
    LONG barStart = 0;   // anchor character position
    LONG barLen = 0;     // anchor sequence length (5 for sum/int, variable for fraction)
    std::vector<MathSlot> slots;
    std::wstring part1;  // Numerator / Upper Limit
    std::wstring part2;  // Denominator / Lower Limit
    std::wstring part3;  // Expression / Function
    std::wstring resultText; // GDI-drawn result (e.g. "\uFF1D 302")

    static size_t SlotIndexFromPart(int partIndex)
    {
        return (size_t)(partIndex <= 1 ? 0 : partIndex - 1);
    }

    static void AppendCount(std::wstring& out, size_t value)
    {
        out += std::to_wstring(value);
    }

    static bool ParseCount(const std::wstring& input, size_t& cursor, size_t& value)
    {
        if (cursor >= input.size() || input[cursor] < L'0' || input[cursor] > L'9')
            return false;

        value = 0;
        while (cursor < input.size() && input[cursor] >= L'0' && input[cursor] <= L'9')
        {
            value = value * 10 + (size_t)(input[cursor] - L'0');
            ++cursor;
        }
        return true;
    }

    static void AppendString(std::wstring& out, const std::wstring& value)
    {
        AppendCount(out, value.size());
        out.push_back(L':');
        out += value;
    }

    static bool ParseString(const std::wstring& input, size_t& cursor, std::wstring& value)
    {
        size_t length = 0;
        if (!ParseCount(input, cursor, length))
            return false;
        if (cursor >= input.size() || input[cursor] != L':')
            return false;
        ++cursor;
        if (cursor + length > input.size())
            return false;
        value = input.substr(cursor, length);
        cursor += length;
        return true;
    }

    static wchar_t EncodeNodeKind(MathNodeKind nodeKind)
    {
        switch (nodeKind)
        {
        case MathNodeKind::Text: return L'T';
        case MathNodeKind::Group: return L'G';
        case MathNodeKind::SquareRoot: return L'R';
        case MathNodeKind::Fraction: return L'F';
        case MathNodeKind::Power: return L'P';
        case MathNodeKind::AbsoluteValue: return L'A';
        case MathNodeKind::Logarithm: return L'L';
        default: return L'?';
        }
    }

    static bool DecodeNodeKind(wchar_t encoded, MathNodeKind& nodeKind)
    {
        switch (encoded)
        {
        case L'T': nodeKind = MathNodeKind::Text; return true;
        case L'G': nodeKind = MathNodeKind::Group; return true;
        case L'R': nodeKind = MathNodeKind::SquareRoot; return true;
        case L'F': nodeKind = MathNodeKind::Fraction; return true;
        case L'P': nodeKind = MathNodeKind::Power; return true;
        case L'A': nodeKind = MathNodeKind::AbsoluteValue; return true;
        case L'L': nodeKind = MathNodeKind::Logarithm; return true;
        default: return false;
        }
    }

    static void SerializeNode(const MathNode& node, std::wstring& out)
    {
        out.push_back(EncodeNodeKind(node.kind));
        AppendString(out, node.text);
        AppendCount(out, node.children.size());
        out.push_back(L'[');
        for (const auto& child : node.children)
            SerializeNode(child, out);
        out.push_back(L']');
    }

    static bool DeserializeNode(const std::wstring& input, size_t& cursor, MathNode& node)
    {
        if (cursor >= input.size())
            return false;

        if (!DecodeNodeKind(input[cursor++], node.kind))
            return false;

        if (!ParseString(input, cursor, node.text))
            return false;

        size_t childCount = 0;
        if (!ParseCount(input, cursor, childCount))
            return false;
        if (cursor >= input.size() || input[cursor] != L'[')
            return false;
        ++cursor;

        node.children.clear();
        node.children.reserve(childCount);
        for (size_t childIndex = 0; childIndex < childCount; ++childIndex)
        {
            MathNode child;
            if (!DeserializeNode(input, cursor, child))
                return false;
            node.children.push_back(std::move(child));
        }

        if (cursor >= input.size() || input[cursor] != L']')
            return false;
        ++cursor;

        if (node.IsStructural())
        {
            const size_t expectedSlots = MathNode::SlotCountForKind(node.kind);
            if (node.children.size() != expectedSlots)
                return false;
            for (const auto& child : node.children)
            {
                if (child.kind != MathNodeKind::Group)
                    return false;
            }
        }

        return true;
    }

    static void SerializeNodeSequence(const std::vector<MathNode>& nodes, std::wstring& out)
    {
        AppendCount(out, nodes.size());
        out.push_back(L'[');
        for (const auto& node : nodes)
            SerializeNode(node, out);
        out.push_back(L']');
    }

    static bool DeserializeNodeSequence(const std::wstring& input, size_t& cursor, std::vector<MathNode>& nodes)
    {
        size_t count = 0;
        if (!ParseCount(input, cursor, count))
            return false;
        if (cursor >= input.size() || input[cursor] != L'[')
            return false;
        ++cursor;

        nodes.clear();
        nodes.reserve(count);
        for (size_t nodeIndex = 0; nodeIndex < count; ++nodeIndex)
        {
            MathNode node;
            if (!DeserializeNode(input, cursor, node))
                return false;
            nodes.push_back(std::move(node));
        }

        if (cursor >= input.size() || input[cursor] != L']')
            return false;
        ++cursor;
        return true;
    }

    std::wstring SerializeTransferPayload() const
    {
        std::wstring output = L"M1|";
        AppendCount(output, (size_t)type);
        output.push_back(L'|');
        AppendString(output, resultText);
        output.push_back(L'|');
        AppendCount(output, slots.size());
        output.push_back(L'[');
        for (const auto& slot : slots)
        {
            AppendString(output, slot.text);
            SerializeNodeSequence(slot.children, output);
        }
        output.push_back(L']');
        return output;
    }

    static bool TryDeserializeTransferPayload(const std::wstring& payload, MathObject& outObj)
    {
        size_t cursor = 0;
        if (payload.size() < 3 || payload.substr(0, 3) != L"M1|")
            return false;
        cursor = 3;

        size_t encodedType = 0;
        if (!ParseCount(payload, cursor, encodedType))
            return false;
        if (cursor >= payload.size() || payload[cursor] != L'|')
            return false;
        ++cursor;

        std::wstring resultText;
        if (!ParseString(payload, cursor, resultText))
            return false;
        if (cursor >= payload.size() || payload[cursor] != L'|')
            return false;
        ++cursor;

        size_t slotCount = 0;
        if (!ParseCount(payload, cursor, slotCount))
            return false;
        if (cursor >= payload.size() || payload[cursor] != L'[')
            return false;
        ++cursor;

        MathObject decoded;
        decoded.type = (MathType)encodedType;
        decoded.resultText = resultText;
        decoded.slots.clear();
        decoded.slots.resize(slotCount);

        for (size_t slotIndex = 0; slotIndex < slotCount; ++slotIndex)
        {
            if (!ParseString(payload, cursor, decoded.slots[slotIndex].text))
                return false;
            if (!DeserializeNodeSequence(payload, cursor, decoded.slots[slotIndex].children))
                return false;
        }

        if (cursor >= payload.size() || payload[cursor] != L']')
            return false;
        ++cursor;
        if (cursor != payload.size())
            return false;

        decoded.SyncLegacyFromSlots();
        outObj = std::move(decoded);
        return true;
    }

    std::wstring BuildPlainTextFallback() const
    {
        switch (type)
        {
        case MathType::Fraction:
            return L"(" + SlotText(1) + L")/(" + SlotText(2) + L")";
        case MathType::SquareRoot:
            if (!SlotText(2).empty() && SlotText(2) != L"2")
                return L"root(" + SlotText(2) + L", " + SlotText(1) + L")";
            return L"sqrt(" + SlotText(1) + L")";
        case MathType::AbsoluteValue:
            return L"abs(" + SlotText(1) + L")";
        case MathType::Power:
            return L"(" + SlotText(1) + L")^(" + SlotText(2) + L")";
        case MathType::Logarithm:
            if (SlotText(1).empty() || SlotText(1) == L"10")
                return L"log(" + SlotText(2) + L")";
            return L"log_[" + SlotText(1) + L"](" + SlotText(2) + L")";
        case MathType::Sum:
            return SlotText(1);
        case MathType::Summation:
            return L"sum[" + SlotText(2) + L".." + SlotText(1) + L"] " + SlotText(3);
        case MathType::Product:
            return L"prod[" + SlotText(2) + L".." + SlotText(1) + L"] " + SlotText(3);
        case MathType::Integral:
            return L"int[" + SlotText(2) + L".." + SlotText(1) + L"] " + SlotText(3);
        case MathType::SystemOfEquations:
        {
            std::wstring text = SlotText(1);
            if (!SlotText(2).empty()) text += L"\r\n" + SlotText(2);
            if (!SlotText(3).empty()) text += L"\r\n" + SlotText(3);
            return text;
        }
        case MathType::Matrix:
            return L"[[" + SlotText(1) + L", " + SlotText(2) + L"], [" + SlotText(3) + L", " + SlotText(4) + L"]]";
        case MathType::Determinant:
            return L"det([[" + SlotText(1) + L", " + SlotText(2) + L"], [" + SlotText(3) + L", " + SlotText(4) + L"]])";
        default:
            return SlotText(1);
        }
    }

    static void NormalizeNodeSequence(std::vector<MathNode>& nodes)
    {
        if (nodes.empty())
        {
            nodes.push_back(MathNode::MakeText());
            return;
        }

        std::vector<MathNode> normalized;
        normalized.reserve(nodes.size());
        for (auto& node : nodes)
        {
            if (node.kind == MathNodeKind::Group)
            {
                NormalizeNodeSequence(node.children);
            }
            else if (node.IsStructural())
            {
                node.EnsureSlotCount(MathNode::SlotCountForKind(node.kind));
                for (size_t slotIndex = 0; slotIndex < node.children.size(); ++slotIndex)
                    NormalizeNodeSequence(node.children[slotIndex].children);
            }

            if (!normalized.empty() && normalized.back().kind == MathNodeKind::Text && node.kind == MathNodeKind::Text)
            {
                normalized.back().text += node.text;
                continue;
            }
            normalized.push_back(std::move(node));
        }

        if (normalized.empty() || normalized.back().kind != MathNodeKind::Text)
            normalized.push_back(MathNode::MakeText());
        nodes = std::move(normalized);
    }

    static std::wstring FlattenNodes(const std::vector<MathNode>& nodes)
    {
        std::wstring text;
        for (const auto& node : nodes)
            text += FlattenNode(node);
        return text;
    }

    void EnsureSlotCount(size_t count = 3)
    {
        if (slots.size() < count)
            slots.resize(count);
    }

    static std::wstring FlattenNode(const MathNode& node)
    {
        if (node.kind == MathNodeKind::Group)
            return FlattenNodes(node.children);

        if (node.kind == MathNodeKind::SquareRoot)
        {
            const std::wstring radicand = FlattenNodes(node.SlotNodes(0));
            const std::wstring index = FlattenNodes(node.SlotNodes(1));
            if (!index.empty() && index != L"2")
                return L"((" + radicand + L")^(1/(" + index + L")))";
            std::wstring text = L"sqrt(" + radicand + L")";
            return text;
        }

        if (node.kind == MathNodeKind::Fraction)
            return L"((" + FlattenNodes(node.SlotNodes(0)) + L")/(" + FlattenNodes(node.SlotNodes(1)) + L"))";

        if (node.kind == MathNodeKind::Power)
            return L"((" + FlattenNodes(node.SlotNodes(0)) + L")^(" + FlattenNodes(node.SlotNodes(1)) + L"))";

        if (node.kind == MathNodeKind::AbsoluteValue)
            return L"abs(" + FlattenNodes(node.SlotNodes(0)) + L")";

        if (node.kind == MathNodeKind::Logarithm)
        {
            const std::wstring base = FlattenNodes(node.SlotNodes(0));
            const std::wstring arg = FlattenNodes(node.SlotNodes(1));
            if (base.empty())
                return L"log(" + arg + L")";
            return L"log_{" + base + L"}(" + arg + L")";
        }

        std::wstring text = node.text;
        for (const auto& child : node.children)
            text += FlattenNode(child);
        return text;
    }

    void RebuildSlotTextFromChildren(size_t slotIndex)
    {
        EnsureSlotCount(slotIndex + 1);
        auto& slot = slots[slotIndex];
        if (slot.children.empty())
            return;

        NormalizeNodeSequence(slot.children);

        slot.text = FlattenNodes(slot.children);
    }

    void EnsureStructuredEditLeaf(int partIndex)
    {
        const size_t slotIndex = SlotIndexFromPart(partIndex);
        EnsureSlotCount((std::max<size_t>)(3, slotIndex + 1));
        MathSlot& slot = slots[slotIndex];
        if (slot.children.empty())
        {
            slot.children.push_back(MathNode::MakeText(slot.text));
        }
        NormalizeNodeSequence(slot.children);
        RebuildSlotTextFromChildren(slotIndex);
        RefreshLegacyPartText(slotIndex);
    }

    std::vector<MathNode> CloneStructuredSlotNodes(int partIndex) const
    {
        const size_t slotIndex = SlotIndexFromPart(partIndex);
        if (slotIndex < slots.size())
            return slots[slotIndex].children;
        return {};
    }

    static std::vector<MathNode>* ResolveSequence(std::vector<MathNode>& rootNodes, const std::vector<size_t>& path, bool create)
    {
        std::vector<MathNode>* sequence = &rootNodes;
        if (create)
            NormalizeNodeSequence(*sequence);

        size_t cursor = 0;
        while (cursor + 1 < path.size())
        {
            const size_t nodeIndex = path[cursor++];
            const size_t slotIndex = path[cursor++];
            if (nodeIndex >= sequence->size())
                return nullptr;

            MathNode& node = (*sequence)[nodeIndex];
            if (!node.IsStructural())
                return nullptr;

            node.EnsureSlotCount(MathNode::SlotCountForKind(node.kind));
            if (slotIndex >= node.children.size())
                return nullptr;

            sequence = &node.children[slotIndex].children;
            if (create)
                NormalizeNodeSequence(*sequence);
        }

        return sequence;
    }

    static const std::vector<MathNode>* ResolveSequence(const std::vector<MathNode>& rootNodes, const std::vector<size_t>& path)
    {
        const std::vector<MathNode>* sequence = &rootNodes;
        size_t cursor = 0;
        while (cursor + 1 < path.size())
        {
            const size_t nodeIndex = path[cursor++];
            const size_t slotIndex = path[cursor++];
            if (nodeIndex >= sequence->size())
                return nullptr;

            const MathNode& node = (*sequence)[nodeIndex];
            if (!node.IsStructural() || slotIndex >= node.children.size() || node.children[slotIndex].kind != MathNodeKind::Group)
                return nullptr;

            sequence = &node.children[slotIndex].children;
        }

        return sequence;
    }

    std::wstring& EditableLeafText(int partIndex, const std::vector<size_t>* path = nullptr)
    {
        const size_t slotIndex = SlotIndexFromPart(partIndex);
        EnsureSlotCount((std::max<size_t>)(3, slotIndex + 1));
        MathSlot& slot = slots[slotIndex];
        if (slot.children.empty() && !path)
            return slot.text;

        NormalizeNodeSequence(slot.children);

        std::vector<MathNode>* sequence = &slot.children;
        if (path && !path->empty())
        {
            sequence = ResolveSequence(slot.children, *path, true);
            if (!sequence)
                return slot.text;
        }

        NormalizeNodeSequence(*sequence);
        MathNode& leaf = sequence->back();
        if (leaf.kind != MathNodeKind::Text)
        {
            sequence->push_back(MathNode::MakeText());
            return sequence->back().text;
        }
        return leaf.text;
    }

    bool InsertNestedSquareRootNode(int partIndex)
    {
        std::vector<size_t> unusedPath;
        return InsertNestedNode(partIndex, {}, L"\\sqrt", MathNodeKind::SquareRoot, unusedPath, 0);
    }

    bool InsertNestedNode(int partIndex,
                          const std::vector<size_t>& activePath,
                          const std::wstring& trigger,
                          MathNodeKind nodeKind,
                          std::vector<size_t>& outLeafPath,
                          size_t initialSlotIndex)
    {
        const size_t slotIndex = SlotIndexFromPart(partIndex);
        EnsureSlotCount((std::max<size_t>)(3, slotIndex + 1));
        MathSlot& slot = slots[slotIndex];
        EnsureStructuredEditLeaf(partIndex);
        NormalizeNodeSequence(slot.children);

        std::vector<size_t> normalizedPath = activePath;
        if (normalizedPath.empty())
            normalizedPath.push_back(slot.children.size() - 1);

        std::vector<MathNode>* sequence = ResolveSequence(slot.children, normalizedPath, true);
        if (!sequence)
            return false;

        NormalizeNodeSequence(*sequence);
        size_t leafIndex = normalizedPath.back();
        if (leafIndex >= sequence->size())
            leafIndex = sequence->size() - 1;
        if ((*sequence)[leafIndex].kind != MathNodeKind::Text)
            return false;

        std::wstring& leafText = (*sequence)[leafIndex].text;
        if (leafText.size() < trigger.size())
            return false;
        if (leafText.compare(leafText.size() - trigger.size(), trigger.size(), trigger) != 0)
            return false;

        leafText.erase(leafText.size() - trigger.size());

        MathNode nestedNode = MathNode::MakeStructured(nodeKind, MathNode::SlotCountForKind(nodeKind));
        const size_t insertedIndex = leafIndex + 1;
        sequence->insert(sequence->begin() + (ptrdiff_t)insertedIndex, std::move(nestedNode));
        sequence->insert(sequence->begin() + (ptrdiff_t)insertedIndex + 1, MathNode::MakeText());

        outLeafPath = activePath;
        if (outLeafPath.empty())
            outLeafPath.push_back(insertedIndex);
        else
            outLeafPath.back() = insertedIndex;
        outLeafPath.push_back(initialSlotIndex);
        outLeafPath.push_back(0);

        NormalizeNodeSequence(slot.children);
        RebuildSlotTextFromChildren(slotIndex);
        RefreshLegacyPartText(slotIndex);
        return true;
    }

    bool MoveToSiblingSlot(int partIndex, std::vector<size_t>& path, int direction) const
    {
        if (path.size() < 3)
            return false;

        const size_t slotIndex = SlotIndexFromPart(partIndex);
        if (slotIndex >= slots.size())
            return false;

        const std::vector<size_t> parentPath(path.begin(), path.end() - 2);
        const std::vector<MathNode>* parentSequence = ResolveSequence(slots[slotIndex].children, parentPath);
        if (!parentSequence)
            return false;

        const size_t parentNodeIndex = path[path.size() - 3];
        const size_t currentSlotIndex = path[path.size() - 2];
        if (parentNodeIndex >= parentSequence->size())
            return false;

        const MathNode& node = (*parentSequence)[parentNodeIndex];
        if (!node.IsStructural())
            return false;

        const size_t slotCount = MathNode::SlotCountForKind(node.kind);
        const ptrdiff_t nextSlot = (ptrdiff_t)currentSlotIndex + direction;
        if (nextSlot < 0 || (size_t)nextSlot >= slotCount)
            return false;

        path[path.size() - 2] = (size_t)nextSlot;
        path[path.size() - 1] = 0;
        return true;
    }

    bool EnterFirstStructuredLeaf(int partIndex, std::vector<size_t>& outPath) const
    {
        const size_t slotIndex = SlotIndexFromPart(partIndex);
        if (slotIndex >= slots.size())
            return false;

        const MathSlot& slot = slots[slotIndex];
        for (size_t nodeIndex = 0; nodeIndex < slot.children.size(); ++nodeIndex)
        {
            if (slot.children[nodeIndex].IsStructural())
            {
                outPath = { nodeIndex, 0, 0 };
                return true;
            }
        }
        return false;
    }

    void SyncLegacyFromSlots()
    {
        EnsureSlotCount();
        for (size_t slotIndex = 0; slotIndex < slots.size(); ++slotIndex)
            RebuildSlotTextFromChildren(slotIndex);
        RefreshLegacyPartText(0);
        RefreshLegacyPartText(1);
        RefreshLegacyPartText(2);
    }

    // Legacy `part1` / `part2` / `part3` mirrors remain for compatibility with older paths.

    void RefreshLegacyPartText(size_t slotIndex)
    {
        EnsureSlotCount();
        const std::wstring value = slotIndex < slots.size() ? slots[slotIndex].text : L"";
        switch (slotIndex)
        {
        case 0:
            part1 = value;
            break;
        case 1:
            part2 = value;
            break;
        case 2:
            part3 = value;
            break;
        default:
            break;
        }
    }

    const std::wstring& SlotText(int partIndex) const
    {
        static const std::wstring kEmpty;
        const size_t slotIndex = SlotIndexFromPart(partIndex);
        if (slotIndex < slots.size())
            return slots[slotIndex].text;
        return kEmpty;
    }

    const std::wstring& PartText(int partIndex) const
    {
        return SlotText(partIndex);
    }

    std::wstring& EditableSlotText(int partIndex)
    {
        const size_t slotIndex = SlotIndexFromPart(partIndex);
        EnsureSlotCount((std::max<size_t>)(3, slotIndex + 1));
        MathSlot& slot = slots[slotIndex];
        if (!slot.children.empty())
            return EditableLeafText(partIndex);
        return slot.text;
    }

    void SetPartText(int partIndex, const std::wstring& value)
    {
        const size_t slotIndex = SlotIndexFromPart(partIndex);
        EnsureSlotCount((std::max<size_t>)(3, slotIndex + 1));
        slots[slotIndex].text = value;
        RefreshLegacyPartText(slotIndex);
    }

    void SetParts(const std::wstring& value1 = L"", const std::wstring& value2 = L"", const std::wstring& value3 = L"")
    {
        EnsureSlotCount();
        slots[0].text = value1;
        slots[1].text = value2;
        slots[2].text = value3;
        RefreshLegacyPartText(0);
        RefreshLegacyPartText(1);
        RefreshLegacyPartText(2);
    }

    void SetMatrix2x2(const std::wstring& a = L"", const std::wstring& b = L"", const std::wstring& c = L"", const std::wstring& d = L"")
    {
        EnsureSlotCount(4);
        slots[0].text = a;
        slots[1].text = b;
        slots[2].text = c;
        slots[3].text = d;
        RefreshLegacyPartText(0);
        RefreshLegacyPartText(1);
        RefreshLegacyPartText(2);
    }
};

struct MathTypingState
{
    bool active = false;      
    int activePart = 0;       // 1-based slot index in MathObject
    size_t objectIndex = 0;   
    std::vector<size_t> activeNodePath;
};
