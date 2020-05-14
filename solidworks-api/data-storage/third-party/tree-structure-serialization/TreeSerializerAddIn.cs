using CodeStack.SwEx.AddIn;
using CodeStack.SwEx.AddIn.Attributes;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace TreeSerializer
{
    [ComVisible(true), Guid("6B8E1B39-5898-46F0-B8DE-753066A2326F")]
    [AutoRegister("Tree Serializer", "Sample Demonstrating use of 3rd party store")]
    public partial class TreeSerializer : SwAddInEx
    {
        private const string STREAM_NAME = "CodeStackSampleStream";

        [CodeStack.SwEx.Common.Attributes.Title("Tree Serializer")]
        public enum Commands_e
        {
            SaveToCurrentDoc,
            LoadFromCurrentDoc
        }

        public override bool OnConnect()
        {
            AddCommandGroup<Commands_e>(OnButtonClick);
            return true;
        }

        private async void OnButtonClick(Commands_e cmd)
        {
            switch (cmd)
            {
                case Commands_e.SaveToCurrentDoc:
                    await SaveTree();
                    break;

                case Commands_e.LoadFromCurrentDoc:
                    LoadTree();
                    break;
            }
        }

        private async Task SaveTree()
        {
            try
            {
                ElementsTree tree = null;

                try
                {
                    tree = ReadDataFromDocument(App.IActiveDoc2);
                    tree.Version = tree.Version + 1;
                }
                catch (ThirdPartyStreamNotFoundException)
                {
                    //create new tree only if stream was never created, show an error otherwise
                    tree = new ElementsTree(1,
                        new Element(1, "Root",
                            new Element(2, "Level1-A",
                                new Element(4, "Level2")),
                            new Element(5, "Level1-B")));
                }

                await SaveDataToDocument(App.IActiveDoc2, tree);
                App.SendMsgToUser2("Data saved",
                    (int)swMessageBoxIcon_e.swMbInformation,
                    (int)swMessageBoxBtn_e.swMbOk);
            }
            catch (Exception ex)
            {
                App.SendMsgToUser2(ex.Message,
                    (int)swMessageBoxIcon_e.swMbStop,
                    (int)swMessageBoxBtn_e.swMbOk);
            }
        }

        private void LoadTree()
        {
            try
            {
                var readTree = ReadDataFromDocument(App.IActiveDoc2);
                App.SendMsgToUser2($"Data Read for '{readTree.Root.Name}' ({readTree.Version})",
                    (int)swMessageBoxIcon_e.swMbInformation,
                    (int)swMessageBoxBtn_e.swMbOk);
            }
            catch (Exception ex)
            {
                App.SendMsgToUser2(ex.Message,
                    (int)swMessageBoxIcon_e.swMbStop,
                    (int)swMessageBoxBtn_e.swMbOk);
            }
        }
    }
}
