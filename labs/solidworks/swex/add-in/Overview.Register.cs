[AutoRegister("My C# SOLIDWORKS Add-In", "Sample SOLIDWORKS add-in in C#", true)]
[ComVisible(true), Guid("736EEACF-B294-40F6-8541-CFC8E7C5AA61")]
public class SampleAddIn : SwAddInEx
{
    public override bool OnConnect()
    {
        return true;
    }
}