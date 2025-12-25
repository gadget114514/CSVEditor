#pragma once

#include <vector>
#include <cstdint>

struct CellPos {
    size_t row;
    size_t col;

    bool operator==(const CellPos& other) const {
        return row == other.row && col == other.col;
    }
};

enum class SelectionMode {
    None,
    Cell,
    Row,
    Column,
    All
};

struct SelectionRange {
    CellPos start;
    CellPos end; // Inclusive
    SelectionMode mode;
};

class EditorState {
public:
    EditorState();

    // Selection
    void SelectCell(size_t row, size_t col, bool multiSelect = false);
    void SelectRow(size_t row, bool multiSelect = false);
    void SelectColumn(size_t col, bool multiSelect = false);
    void SelectAll();
    void ClearSelection();
    void DragTo(size_t row, size_t col);

    const std::vector<SelectionRange>& GetSelections() const;
    bool IsSelected(size_t row, size_t col) const;
    bool IsRowSelected(size_t row) const;
    bool IsColumnSelected(size_t col) const;

    // Scroll Position
    size_t GetScrollRow() const { return m_scrollRow; }
    void SetScrollRow(size_t row) { m_scrollRow = row; }
    
    float GetScrollX() const { return m_scrollX; }
    void SetScrollX(float x) { m_scrollX = x; }

private:
    std::vector<SelectionRange> m_selections;
    SelectionMode m_currentMode = SelectionMode::None;
    CellPos m_anchor; // Start of drag

    size_t m_scrollRow = 0;
    float m_scrollX = 0.0f;
};
