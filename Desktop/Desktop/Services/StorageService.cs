using Windows.Storage;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Services;

internal sealed class StorageService : IStorageService
{
    public void CreateContainer(string name)
    {
        ApplicationData.Current.LocalSettings.CreateContainer(name, ApplicationDataCreateDisposition.Always);
    }

    public bool DoesContainerExist(string name)
    {
        return ApplicationData.Current.LocalSettings.Containers.ContainsKey(name);
    }

    public TData? FetchData<TData>(string containerName, string key)
    {
        if (!DoesContainerExist(containerName)) return default;

        return ApplicationData.Current.LocalSettings.Containers[containerName].Values.TryGetValue(key, out var value)
            ? (TData)value
            : default;
    }

    public void SaveData(string containerName, string key, object data)
    {
        if (!DoesContainerExist(containerName)) return;

        ApplicationData.Current.LocalSettings.Containers[containerName].Values[key] = data;
    }
}