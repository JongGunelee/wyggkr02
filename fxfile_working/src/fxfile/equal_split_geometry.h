// GPLv3. See LICENSE. Shared geometry for automatic and explicit fitting.
#pragma once
namespace fxfile
{
struct EqualSplitGeometry
{
    int columnSizes[3];
    int rowSizes[2];
    int storedSizes[3];
    double ratios[3];
};
inline bool calculateEqualSplitGeometry(int width, int height, int rows,
                                        int columns, int gap, EqualSplitGeometry &out)
{
    if (width <= 0 || height <= 0 || rows < 1 || rows > 2 ||
        columns < 1 || columns > 3 || gap < 0) return false;
    const long long usableWidth = static_cast<long long>(width) - static_cast<long long>(gap) * (columns - 1);
    const long long usableHeight = static_cast<long long>(height) - static_cast<long long>(gap) * (rows - 1);
    if (usableWidth < columns || usableHeight < rows) return false;
    EqualSplitGeometry result = {};
    for (int column = 0; column < columns; ++column)
        result.columnSizes[column] = static_cast<int>(usableWidth * (column + 1) / columns - usableWidth * column / columns);
    for (int row = 0; row < rows; ++row)
        result.rowSizes[row] = static_cast<int>(usableHeight * (row + 1) / rows - usableHeight * row / rows);
    int index = 0;
    for (int column = 0; column < columns - 1; ++column, ++index)
    {
        result.storedSizes[index] = result.columnSizes[column];
        result.ratios[index] = static_cast<double>(result.columnSizes[column]) / width;
    }
    for (int row = 0; row < rows - 1; ++row, ++index)
    {
        result.storedSizes[index] = result.rowSizes[row];
        result.ratios[index] = static_cast<double>(result.rowSizes[row]) / height;
    }
    out = result;
    return true;
}
} // namespace fxfile
