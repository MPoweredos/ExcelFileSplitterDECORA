namespace ExcelFileSplitter;

public sealed record RowBlock(int First, int Last)
{
    public int Count => Last - First + 1;
}

public static class RowBlocks
{
    public static List<RowBlock> ToDelete<T>(IReadOnlyList<T> values, int firstRow, Func<T, bool> keep)
    {
        var blocks = new List<RowBlock>();
        int blockStart = -1;

        for (int i = 0; i < values.Count; i++)
        {
            if (!keep(values[i]))
            {
                if (blockStart < 0) blockStart = i;
            }
            else if (blockStart >= 0)
            {
                blocks.Add(new RowBlock(firstRow + blockStart, firstRow + i - 1));
                blockStart = -1;
            }
        }

        if (blockStart >= 0) blocks.Add(new RowBlock(firstRow + blockStart, firstRow + values.Count - 1));
        return blocks;
    }

    public static long TotalRows(IReadOnlyList<RowBlock> blocks)
    {
        long total = 0;
        foreach (RowBlock block in blocks) total += block.Count;
        return total;
    }
}
