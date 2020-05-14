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

private void LoadFromStorageStore(IModelDoc2 model)
{
    using (var storageHandler = model.Access3rdPartyStorageStore(STORAGE_NAME, false))
    {
        if (storageHandler.Storage != null)
        {
            using (var str = storageHandler.Storage.TryOpenStream(STREAM1_NAME, false))
            {
                if (str != null)
                {
                    var xmlSer = new XmlSerializer(typeof(StorageStreamData));
                    m_StorageData = xmlSer.Deserialize(str) as StorageStreamData;
                }
            }

            using (var subStorage = storageHandler.Storage.TryOpenStorage(SUB_STORAGE_NAME, false))
            {
                if (subStorage != null)
                {
                    using (var str = subStorage.TryOpenStream(STREAM2_NAME, false))
                    {
                        if (str != null)
                        {
                            var buffer = new byte[str.Length];
                            str.Read(buffer, 0, buffer.Length);
                            var dateStr = Encoding.UTF8.GetString(buffer);
                            var date = DateTime.Parse(dateStr);
                        }
                    }
                }
            }
        }
    }
}