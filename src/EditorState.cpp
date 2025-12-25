#include "EditorState.h"

EditorState::EditorState()
{
}

void EditorState::SelectCell(size_t row, size_t col, bool multiSelect)
{
    if (!multiSelect) m_selections.clear();
    
    SelectionRange range;
    range.start = {row, col};
    range.end = {row, col};
    range.mode = SelectionMode::Cell;
    m_selections.push_back(range);
    
    m_anchor = {row, col};
    m_currentMode = SelectionMode::Cell;
}

void EditorState::SelectRow(size_t row, bool multiSelect)
{
    if (!multiSelect) m_selections.clear();

    SelectionRange range;
    range.start = {row, 0};
    range.end = {row, 0}; // Col 0 ignored for row mode
    range.mode = SelectionMode::Row;
    m_selections.push_back(range);
    
    m_anchor = {row, 0};
    m_currentMode = SelectionMode::Row;
}

void EditorState::SelectColumn(size_t col, bool multiSelect)
{
    if (!multiSelect) m_selections.clear();

    SelectionRange range;
    range.start = {0, col};
    range.end = {0, col}; // Row 0 ignored for col mode
    range.mode = SelectionMode::Column;
    m_selections.push_back(range);
    
    m_anchor = {0, col};
    m_currentMode = SelectionMode::Column;
}

void EditorState::SelectAll()
{
    m_selections.clear();
    
    SelectionRange range;
    range.start = {0, 0};
    range.end = {0, 0}; // Meaningless for All? Or implied max? match Column/Row logic
    range.mode = SelectionMode::All;
    m_selections.push_back(range);
    
    m_currentMode = SelectionMode::All;
}

void EditorState::ClearSelection()
{
    m_selections.clear();
    m_currentMode = SelectionMode::None;
}

void EditorState::DragTo(size_t row, size_t col)
{
    if (m_selections.empty()) return;
    
    // Update last selection end
    auto& range = m_selections.back();
    
    if (m_currentMode == SelectionMode::Row) {
        // Anchor row fixed?
        // Usually anchor is fixed.
        // range.start is implicitly anchor? No. 
        // We should store Anchor properly.
        // m_anchor (from SelectRow)
        
        // Determine start/end from m_anchor and current (row, col)
        // range.start = min(anchor.row, row)
        // range.end = max(anchor.row, row)
        // But we store start/end simply.
        
        range.start.row = (std::min)(m_anchor.row, row);
        range.end.row = (std::max)(m_anchor.row, row);
    } else if (m_currentMode == SelectionMode::Column) {
        range.start.col = (std::min)(m_anchor.col, col);
        range.end.col = (std::max)(m_anchor.col, col);
    } else if (m_currentMode == SelectionMode::Cell) {
        range.start.row = (std::min)(m_anchor.row, row);
        range.start.col = (std::min)(m_anchor.col, col);
        range.end.row = (std::max)(m_anchor.row, row);
        range.end.col = (std::max)(m_anchor.col, col);
    }
}

const std::vector<SelectionRange>& EditorState::GetSelections() const
{
    return m_selections;
}

bool EditorState::IsSelected(size_t row, size_t col) const
{
    for (const auto& sel : m_selections) {
        if (sel.mode == SelectionMode::Cell) {
            if (row >= sel.start.row && row <= sel.end.row &&
                col >= sel.start.col && col <= sel.end.col) return true;
        } else if (sel.mode == SelectionMode::Row) {
            if (row >= sel.start.row && row <= sel.end.row) return true;
        } else if (sel.mode == SelectionMode::Column) {
            if (col >= sel.start.col && col <= sel.end.col) return true;
        } else if (sel.mode == SelectionMode::All) {
            return true;
        }
    }
    return false;
}

bool EditorState::IsRowSelected(size_t row) const
{
    for (const auto& sel : m_selections) {
        if (sel.mode == SelectionMode::Row) {
            if (row >= sel.start.row && row <= sel.end.row) return true;
        }
    }
    return false;
}

bool EditorState::IsColumnSelected(size_t col) const
{
    for (const auto& sel : m_selections) {
        if (sel.mode == SelectionMode::Column) {
            if (col >= sel.start.col && col <= sel.end.col) return true;
        }
    }
    return false;
}
