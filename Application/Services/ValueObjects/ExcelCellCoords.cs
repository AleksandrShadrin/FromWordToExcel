namespace Application.Services.ValueObjects
{
    public record ExcelCellCoords
    {
        public int Horizontal { get; init; }
        public int Vertical { get; init; }
        public ExcelCellCoords(int horizontal, int vertical)
        {
            if (horizontal <= 0)
            {
                throw new ArgumentOutOfRangeException(nameof(horizontal));
            }
            if (vertical <= 0)
            {
                throw new ArgumentOutOfRangeException(nameof(horizontal));
            }
            Horizontal = horizontal;
            Vertical = vertical;
        }
    };
}