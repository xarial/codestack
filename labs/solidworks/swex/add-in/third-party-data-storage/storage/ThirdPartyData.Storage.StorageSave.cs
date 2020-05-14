private const string STORAGE_NAME = "CodeStackStorage";
private const string STREAM1_NAME = "CodeStackStream1";
private const string STREAM2_NAME = "CodeStackStream2";
private const string SUB_STORAGE_NAME = "CodeStackSubStorage";

public class StorageStreamData
{
    public int Prp3 { get; set; }
    public bool Prp4 { get; set; }
}

private StorageStreamData m_StorageData;

private void SaveToStorageStore(IModelDoc2 model)
{
    using (var storageHandler = model.Access3rdPartyStorageStore(STORAGE_NAME, true))
    {
        using (var str = storageHandler.Storage.TryOpenStream(STREAM1_NAME, true))
        {
            var xmlSer = new XmlSerializer(typeof(StorageStreamData));

            xmlSer.Serialize(str, m_StorageData);
        }

        using (var subStorage = storageHandler.Storage.TryOpenStorage(SUB_STORAGE_NAME, true))
        {
            using (var str = subStorage.TryOpenStream(STREAM2_NAME, true))
            {
                var buffer = Encoding.UTF8.GetBytes(DateTime.Now.ToString("yyyy-MM-dd-hh-mm-ss"));
                str.Write(buffer, 0, buffer.Length);
            }
        }
    }
}