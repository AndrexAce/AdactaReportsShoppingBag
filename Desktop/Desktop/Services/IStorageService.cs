namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Services;

internal interface IStorageService
{
    public void CreateContainer(string name);

    public bool DoesContainerExist(string name);

    public TData? FetchData<TData>(string containerName, string key);

    public void SaveData(string containerName, string key, object data);
}
