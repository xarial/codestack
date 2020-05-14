private const string STREAM_NAME = "CodeStackStream";

public class StreamData
{
    public string Prp1 { get; set; }
    public double Prp2 { get; set; }
}

private StreamData m_StreamData;

private void SaveToStream(IModelDoc2 model)
{
    using (var streamHandler = model.Access3rdPartyStream(STREAM_NAME, true))
    {
        using (var str = streamHandler.Stream)
        {
            var xmlSer = new XmlSerializer(typeof(StreamData));

            xmlSer.Serialize(str, m_StreamData);
        }
    }
}