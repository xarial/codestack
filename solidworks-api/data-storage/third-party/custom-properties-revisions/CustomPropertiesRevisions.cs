using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;
using ThirdPartyStorage;

namespace CodeStack
{
    public partial class CustomPropertiesRevisions
    {
        public class ThirdPartyStoreNotFoundException : Exception
        {
        }

        private async Task SaveSnapshotToDocument(IModelDoc2 model, PropertiesSnapshot data)
        {
            int err = -1;
            int warn = -1;

            model.SetSaveFlag();

            const int S_OK = 0;

            bool? result = null; ;

            var onSaveToStorageStoreNotifyFunc = new Func<int>(() =>
            {
                try
                {
                    StoreData(model, data, STORAGE_NAME, storage =>
                    {
                        string snapshotName = "";

                        AccessStreamFromPath(storage, SNAPSHOT_INFO_STREAM_NAME, true, true, stream =>
                        {
                            var ser = new DataContractSerializer(typeof(List<SnapshotInfo>));

                            List<SnapshotInfo> snapshotInfos = null;

                            if (stream.Length > 0)
                            {
                                snapshotInfos = ser.ReadObject(stream) as List<SnapshotInfo>;
                            }
                            else
                            {
                                snapshotInfos = new List<SnapshotInfo>();
                            }

                            var info = new SnapshotInfo()
                            {
                                Revision = snapshotInfos.Count + 1,
                                TimeStamp = DateTime.Now
                            };

                            snapshotInfos.Add(info);

                            snapshotName = string.Format(SNAPSHOT_STREAM_NAME_TEMPLATE, info.Revision);

                            stream.Seek(0, System.IO.SeekOrigin.Begin);

                            ser.WriteObject(stream, snapshotInfos);
                        }, STGM.STGM_READWRITE | STGM.STGM_SHARE_EXCLUSIVE);

                        AccessStreamFromPath(storage, snapshotName, true, true, stream =>
                        {
                            var ser = new DataContractSerializer(typeof(PropertiesSnapshot));
                            ser.WriteObject(stream, data);
                        }, STGM.STGM_READWRITE | STGM.STGM_SHARE_EXCLUSIVE);

                        result = true;
                    });
                }
                catch
                {
                    result = false;
                }
                return S_OK;
            });

            var partSaveToStorageNotify = new DPartDocEvents_SaveToStorageStoreNotifyEventHandler(onSaveToStorageStoreNotifyFunc);
            var assmSaveToStorageNotify = new DAssemblyDocEvents_SaveToStorageStoreNotifyEventHandler(onSaveToStorageStoreNotifyFunc);
            var drwSaveToStorageNotify = new DDrawingDocEvents_SaveToStorageStoreNotifyEventHandler(onSaveToStorageStoreNotifyFunc);

            #region Attach Event Handlers

            switch ((swDocumentTypes_e)model.GetType())
            {
                case swDocumentTypes_e.swDocPART:
                    (model as PartDoc).SaveToStorageStoreNotify += partSaveToStorageNotify;
                    break;

                case swDocumentTypes_e.swDocASSEMBLY:
                    (model as AssemblyDoc).SaveToStorageStoreNotify += assmSaveToStorageNotify;
                    break;

                case swDocumentTypes_e.swDocDRAWING:
                    (model as DrawingDoc).SaveToStorageStoreNotify += drwSaveToStorageNotify;
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
                    (model as PartDoc).SaveToStorageStoreNotify -= partSaveToStorageNotify;
                    break;

                case swDocumentTypes_e.swDocASSEMBLY:
                    (model as AssemblyDoc).SaveToStorageStoreNotify -= assmSaveToStorageNotify;
                    break;

                case swDocumentTypes_e.swDocDRAWING:
                    (model as DrawingDoc).SaveToStorageStoreNotify -= drwSaveToStorageNotify;
                    break;
            }

            #endregion

            if (!result.Value)
            {
                throw new Exception("Failed to store the data");
            }
        }

        private PropertiesSnapshot ReadSnapshotFromDocument(IModelDoc2 model, string revName)
        {
            return ReadData<PropertiesSnapshot>(model, STORAGE_NAME, revName);
        }

        private SnapshotInfo[] GetSnapshotInfos(IModelDoc2 model)
        {
            return ReadData<SnapshotInfo[]>(model, STORAGE_NAME, SNAPSHOT_INFO_STREAM_NAME);
        }

        private void StoreData<T>(IModelDoc2 model, T data, string storageName, Action<ComStorage> action)
        {
            try
            {
                var storage = model.Extension.IGet3rdPartyStorageStore(storageName, true) as IStorage;

                using (var comStorage = new ComStorage(storage, true))
                {
                    action.Invoke(comStorage);
                }
            }
            catch
            {
                throw;
            }
            finally
            {
                model.Extension.IRelease3rdPartyStorageStore(storageName);
            }
        }

        private T ReadData<T>(IModelDoc2 model, string storageName, string streamName)
        {
            T data = default(T);

            ReadStorage(model, storageName, storage => 
            {
                AccessStreamFromPath(storage, streamName, false, false, stream=> 
                {
                    var ser = new DataContractSerializer(typeof(T));
                    data = (T)ser.ReadObject(stream);
                });
            });

            return data;
        }

        private void AccessStreamFromPath(ComStorage storage, string path, bool writable,
            bool createIfNotExist, Action<ComStream> action, STGM mode = STGM.STGM_SHARE_EXCLUSIVE)
        {
            var parentIndex = path.IndexOf('\\');

            if (parentIndex == -1)
            {
                IStream stream = null;

                try
                {
                    stream = storage.OpenStream(path, mode);
                }
                catch
                {
                    if (createIfNotExist)
                    {
                        stream = storage.CreateStream(path);
                    }
                    else
                    {
                        throw;
                    }
                }

                using (var comStream = new ComStream(stream, writable))
                {
                    action.Invoke(comStream);
                }
            }
            else
            {
                var subStorageName = path.Substring(0, parentIndex);

                IStorage subStorage;

                try
                {
                    subStorage = storage.OpenStorage(subStorageName, mode);
                }
                catch
                {
                    if (createIfNotExist)
                    {
                        subStorage = storage.CreateStorage(subStorageName);
                    }
                    else
                    {
                        throw;
                    }
                }
                
                using (var subComStorage = new ComStorage(subStorage, false))
                {
                    var nextLevelPath = path.Substring(parentIndex + 1);
                    AccessStreamFromPath(subComStorage, nextLevelPath, writable, createIfNotExist, action);
                }
            }
        }

        private void ReadStorage(IModelDoc2 model, string storageName, Action<ComStorage> action)
        {
            try
            {
                var storage = model.Extension.IGet3rdPartyStorageStore(storageName, false) as IStorage;

                if (storage != null)
                {
                    using (var comStorage = new ComStorage(storage, false))
                    {
                        action.Invoke(comStorage);
                    }
                }
                else
                {
                    throw new ThirdPartyStoreNotFoundException();
                }
            }
            catch
            {
                throw;
            }
            finally
            {
                model.Extension.IRelease3rdPartyStorageStore(storageName);
            }
        }
    }
}
