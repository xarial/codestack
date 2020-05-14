using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;
using ThirdPartyStorage;

namespace TreeSerializer
{
    public partial class TreeSerializer
    {
        public class ThirdPartyStreamNotFoundException : Exception
        {
        }

        private async Task SaveDataToDocument(IModelDoc2 model, ElementsTree data)
        {
            int err = -1;
            int warn = -1;

            model.SetSaveFlag();

            const int S_OK = 0;

            bool? result = null; ;

            var onSaveToStorageNotifyFunc = new Func<int>(() =>
            {
                try
                {
                    StoreData(model, data, STREAM_NAME);
                    result = true;
                }
                catch
                {
                    result = false;
                }
                return S_OK;
            });

            var partSaveToStorageNotify = new DPartDocEvents_SaveToStorageNotifyEventHandler(onSaveToStorageNotifyFunc);
            var assmSaveToStorageNotify = new DAssemblyDocEvents_SaveToStorageNotifyEventHandler(onSaveToStorageNotifyFunc);
            var drwSaveToStorageNotify = new DDrawingDocEvents_SaveToStorageNotifyEventHandler(onSaveToStorageNotifyFunc);

            #region Attach Event Handlers

            switch ((swDocumentTypes_e)model.GetType())
            {
                case swDocumentTypes_e.swDocPART:
                    (model as PartDoc).SaveToStorageNotify += partSaveToStorageNotify;
                    break;

                case swDocumentTypes_e.swDocASSEMBLY:
                    (model as AssemblyDoc).SaveToStorageNotify += assmSaveToStorageNotify;
                    break;

                case swDocumentTypes_e.swDocDRAWING:
                    (model as DrawingDoc).SaveToStorageNotify += drwSaveToStorageNotify;
                    break;
            }

            #endregion

            if (!model.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref err, ref warn))
            {
                throw new InvalidOperationException($"Failed to save the model: {(swFileSaveError_e)err}");
            }

            await Task.Run(() =>
            {
                while (!result.HasValue)
                {
                    System.Threading.Thread.Sleep(10);
                }
            });

            #region Detach Event Handlers

            switch ((swDocumentTypes_e)model.GetType())
            {
                case swDocumentTypes_e.swDocPART:
                    (model as PartDoc).SaveToStorageNotify -= partSaveToStorageNotify;
                    break;

                case swDocumentTypes_e.swDocASSEMBLY:
                    (model as AssemblyDoc).SaveToStorageNotify -= assmSaveToStorageNotify;
                    break;

                case swDocumentTypes_e.swDocDRAWING:
                    (model as DrawingDoc).SaveToStorageNotify -= drwSaveToStorageNotify;
                    break;
            }

            #endregion

            if (!result.Value)
            {
                throw new Exception("Failed to store the data");
            }
        }

        private ElementsTree ReadDataFromDocument(IModelDoc2 model)
        {
            return ReadData<ElementsTree>(model, STREAM_NAME);
        }

        private void StoreData<T>(IModelDoc2 model, T data, string streamName)
        {
            try
            {
                var stream = model.IGet3rdPartyStorage(streamName, true) as IStream;

                using (var comStr = new ComStream(stream, true, false))
                {
                    comStr.Seek(0, System.IO.SeekOrigin.Begin);
                    var ser = new XmlSerializer(typeof(T));
                    ser.Serialize(comStr, data);
                }
            }
            catch
            {
                throw;
            }
            finally
            {
                model.IRelease3rdPartyStorage(streamName);
            }
        }

        private T ReadData<T>(IModelDoc2 model, string streamName)
        {
            try
            {
                var stream = model.IGet3rdPartyStorage(streamName, false) as IStream;

                if (stream != null)
                {
                    using (var comStr = new ComStream(stream, false))
                    {
                        comStr.Seek(0, System.IO.SeekOrigin.Begin);
                        var ser = new XmlSerializer(typeof(T));
                        return (T)ser.Deserialize(comStr);
                    }
                }
                else
                {
                    throw new ThirdPartyStreamNotFoundException();
                }
            }
            catch
            {
                throw;
            }
            finally
            {
                model.IRelease3rdPartyStorage(streamName);
            }
        }
    }
}
