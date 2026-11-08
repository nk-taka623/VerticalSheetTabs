using System;

namespace VerticalSheetTabs
{
    // How the rows of the vertical tabs are laid out in the pane.
    public enum RowLayout
    {
        Hanging, // the rows hang from the top: tool row, band, rows, and the "+" at the bottom (the standard)
        Stacked  // the rows pile up from the bottom: rows, band, tool row, and the "+" at the bottom
    }

    // The arithmetic of where the rows are (no screen, no Excel: the tests use it as it is).
    // "Area" is the part of the pane the rows may use: from areaTop to areaBottom.
    public static class RowPlacement
    {
        // The space above the first row when the rows pile up from the bottom and all of them fit; 0 otherwise.
        public static int StackOffset(bool stacked, int areaHeight, int rowCount, int rowPx)
        {
            if (!stacked || rowPx <= 0 || rowCount < 0) return 0;
            return Math.Max(0, areaHeight - rowCount * rowPx);
        }

        // The top of a row.
        public static int RowTop(int areaTop, int stackOffset, int index, int rowPx, int scrollY)
        {
            return areaTop + stackOffset + index * rowPx - scrollY;
        }

        // The row at the height y, or -1 (above the first row, below the last, outside the area).
        // (The division is not done on a negative number: it would round toward 0 and give the first row.)
        public static int IndexFromY(int y, int areaTop, int areaBottom, int stackOffset, int scrollY, int rowPx, int rowCount)
        {
            if (rowPx <= 0 || y < areaTop || y >= areaBottom) return -1;
            int d = y - areaTop - stackOffset + scrollY;
            if (d < 0) return -1;
            int i = d / rowPx;
            return i < rowCount ? i : -1;
        }

        // The largest scroll position (0 when the rows fit).
        public static int MaxScroll(int rowCount, int rowPx, int areaHeight)
        {
            return Math.Max(0, rowCount * rowPx - areaHeight);
        }

        // How far the list is scrolled up from its bottom end (0 = the last row is at the bottom of the area).
        public static int FromBottom(int scrollY, int rowCount, int rowPx, int areaHeight)
        {
            int max = MaxScroll(rowCount, rowPx, areaHeight);
            return Math.Max(0, Math.Min(max, max - scrollY));
        }

        // The scroll position that keeps the distance from the bottom end.
        public static int ScrollFromBottom(int fromBottom, int rowCount, int rowPx, int areaHeight)
        {
            int max = MaxScroll(rowCount, rowPx, areaHeight);
            return Math.Max(0, Math.Min(max, max - Math.Max(0, fromBottom)));
        }

        // The place of the row that is shown while a row is added or deleted: where the first row was before.
        public static int OldRowTop(int oldFirstTop, int index, int rowPx)
        {
            return oldFirstTop + index * rowPx;
        }

        // The top of the part of the area the wallpaper may cover.
        // Hanging: the rows are at the top; transparent rows -> from the top of the area, otherwise from below the last row.
        // Stacked: the rows are at the bottom; the wallpaper is above them (transparent rows: the whole area).
        public static void WallpaperSpan(bool stacked, bool transparentRows, int areaTop, int areaBottom, int firstRowTop, int rowCount, int rowPx, int scrollY, out int top, out int bottom)
        {
            top = areaTop;
            bottom = areaBottom;
            if (transparentRows) return;
            if (stacked) bottom = Math.Min(areaBottom, firstRowTop);
            else top = Math.Max(areaTop, areaTop + rowCount * rowPx - scrollY);
        }
    }
}
