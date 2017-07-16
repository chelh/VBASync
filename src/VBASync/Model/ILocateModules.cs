namespace VBASync.Model
{
    public interface ILocateModules
    {
        string GetFrxPath(string name);
        string GetModulePath(string name, ModuleType type);
    }
}
