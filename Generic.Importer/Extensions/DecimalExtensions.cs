
namespace Generic.Importer.Extensions
{
    public static class DecimalExtensions
    {
        public static string ToCurrency(this decimal value)
        {
            return $"{value:C}";
        }
    }
}
