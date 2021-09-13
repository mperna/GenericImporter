
namespace Generic.Importer.Interfaces
{
    public interface IPartManager
    {
        int GetPartIdentifier(string partNumber);
        decimal GetPartUnitCost(string partNumber);
    }
}
