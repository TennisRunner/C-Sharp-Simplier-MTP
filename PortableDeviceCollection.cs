using System.Collections.ObjectModel;
using PortableDeviceApiLib;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System;
using PortableDevicesConstants;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;

namespace PortableDevices
{
    public class MobileDeviceManager : Collection<PortableDevice>
    {
        private readonly PortableDeviceManager _deviceManager;

        Dictionary<string, PortableDeviceObject> recentFiles;


        public MobileDeviceManager()
        {
            this._deviceManager = new PortableDeviceManager();
            this.recentFiles = new Dictionary<string, PortableDeviceObject>();
        }


        public void Refresh()
        {
            this._deviceManager.RefreshDeviceList();

            // Determine how many WPD devices are connected
            var deviceIds = new string[1];
            uint count = 1;
            this._deviceManager.GetDevices(ref deviceIds[0], ref count);

            // Retrieve the device id for each connected device
            deviceIds = new string[count];
            this._deviceManager.GetDevices(ref deviceIds[0], ref count);
            foreach (var deviceId in deviceIds)
            {
                Add(new PortableDevice(deviceId));
            }
        }


        public List<string> GetLogicalDrives()
        {
            List<string> results;


            results = new List<string>();


            this.Refresh();

            foreach (var a in this)
            {
                results.Add(a.FriendlyName);
            }

            return results;
        }

        public List<string> GetFiles(string startPath, string regexPattern, bool topLevelOnly)
        {
            List<string> results,
                queue;

            int processedIndex = 0;

            Regex pattern;

            PortableDeviceFolder folder;


            this.Refresh();

            results = new List<string>();
            queue = new List<string>();
            pattern = new Regex(regexPattern);

            folder = this.GetFolder(startPath);

            if (folder != null)
            {
                queue.Add(startPath);

                do
                {
                    folder = this.GetFolder(queue[processedIndex]);
                    folder.RefreshContent();

                    foreach (PortableDeviceObject o in folder.Files)
                    {
                        if (o is PortableDeviceFolder)
                        {
                            queue.Add(o.GetAbsolutePath());
                        }
                        else
                        {
                            if (pattern.IsMatch(o.Name) == true)
                            {
                                string path = o.GetAbsolutePath();

                                if (recentFiles.ContainsKey(path) == false)
                                {
                                    recentFiles.Add(path, o as PortableDeviceFile);
                                }
                                else
                                {
                                    recentFiles[path] = o as PortableDeviceFile;
                                }

                                results.Add(path);
                            }
                        }
                    }

                    if (++processedIndex >= queue.Count)
                        break;
                }
                while (topLevelOnly == false);

                queue.RemoveAt(0);
            }
            else
                throw new Exception("Start path does not exist");

            return results;
        }

        public List<string> GetDirectories(string startPath, string regexPattern, bool topLevelOnly)
        {
            List<string> results,
                queue;

            int processedIndex = 0;

            Regex pattern;
            
            PortableDeviceFolder folder;

            
            this.Refresh();
            
            results = new List<string>();
            queue = new List<string>();
            pattern = new Regex(regexPattern);

            folder = this.GetFolder(startPath);

            if (folder != null)
            {
                queue.Add(startPath);

                do
                {
                    folder = this.GetFolder(queue[processedIndex]);
                    folder.RefreshContent();

                    foreach (PortableDeviceObject o in folder.Files)
                    {
                        if (o is PortableDeviceFolder)
                        {
                            if (pattern.IsMatch(o.Name) == true)
                            {
                                results.Add(o.GetAbsolutePath());
                            }

                            queue.Add(o.GetAbsolutePath());
                        }
                    }

                    if (++processedIndex >= queue.Count)
                        break;
                }
                while (topLevelOnly == false);

                queue.RemoveAt(0);
            }
            else
                throw new Exception("Start path does not exist");


            return results;
        }

        public void DeleteFile(string fileName)
        {
            PortableDeviceObject file;
            
           

            if (recentFiles.ContainsKey(fileName) == true)
                file = recentFiles[fileName];
            else
                file = GetFile(fileName);

            if (file != null)
            {
                file.parentDevice.Connect();
                DeleteObject(file.content, file.Id);
                file.parentDevice.Disconnect();
                
                if (recentFiles.ContainsKey(fileName) == true)
                    recentFiles.Remove(fileName);
            }
            else
                throw new Exception("File does not exist for deletion");
        }

        public void ClearFileCache()
        {
            recentFiles.Clear();
        }

        public bool FileExists(string fileName)
        {
            PortableDeviceFile file;


            if (recentFiles.ContainsKey(fileName) == true)
                file = recentFiles[fileName] as PortableDeviceFile;
            else
                file = this.GetFile(fileName);

            if (file != null)
                return true;
            else
                return false;
        }



        public PortableDeviceFolder CreateFolder(string path)
        {
            List<string> parts;

            PortableDeviceFolder res,
                                    parentFolder;

            string targetFolderName;


            res = null;

            parts = path.Split(new char[] { '/', '\\' }, StringSplitOptions.RemoveEmptyEntries).ToList();

            if (parts.Count > 1)
            {
                targetFolderName = parts.Last();

                if (Regex.IsMatch(targetFolderName, "^[A-Za-z0-9-_ ]{1,}$") == true)
                {
                    res = this.GetFolder(path);

                    if (res == null)
                    {
                        parentFolder = this.GetFolder(string.Join("\\", parts.Take(parts.Count - 1).ToArray()));

                        if (parentFolder != null)
                        {
                            IPortableDeviceProperties deviceProperties;
                            parentFolder.content.Properties(out deviceProperties);

                            IPortableDeviceValues deviceValues = (IPortableDeviceValues)new PortableDeviceTypesLib.PortableDeviceValues();


                            deviceValues.SetStringValue(ref PortableDeviceConstants.WPD_OBJECT_PARENT_ID, parentFolder.Id);
                            deviceValues.SetStringValue(ref PortableDeviceConstants.WPD_OBJECT_NAME, parts.Last());
                            deviceValues.SetStringValue(ref PortableDeviceConstants.WPD_OBJECT_ORIGINAL_FILE_NAME, parts.Last());

                            deviceValues.SetGuidValue(ref PortableDeviceConstants.WPD_OBJECT_CONTENT_TYPE, ref PortableDeviceConstants.WPD_CONTENT_TYPE_FOLDER);


                            string objectId = string.Empty;

                            parentFolder.parentDevice.Connect();
                            parentFolder.content.CreateObjectWithPropertiesOnly(deviceValues, ref objectId);
                            parentFolder.parentDevice.Disconnect();

                            if (objectId != string.Empty)
                            {
                                res = this.GetFolder(path);
                            }
                            else
                                throw new Exception("Unable to create folder");
                        }
                        else
                            throw new Exception("Unable to find parent folder");
                    }
                    else
                        throw new Exception("Folder already exists");
                }
                else
                    throw new Exception("Folder name is not valid");
            }
            else
                throw new Exception("Path cannot be created in device folder");

            return res;
        }


        public bool CopyFileFromDevice(string sourceName, string destinationName)
        {
            PortableDeviceFile file;



            file = GetFile(sourceName);

            if (file != null)
            {
                IPortableDeviceResources resources;
                file.content.Transfer(out resources);
                file.parentDevice.Connect();
                PortableDeviceApiLib.IStream wpdStream;
                uint optimalTransferSize = 0;

                resources.GetStream(file.Id, ref PortableDeviceConstants.WPD_RESOURCE_DEFAULT, PortableDeviceConstants.STGM_READ, ref optimalTransferSize, out wpdStream);

                System.Runtime.InteropServices.ComTypes.IStream sourceStream = (System.Runtime.InteropServices.ComTypes.IStream)wpdStream;


                FileStream targetStream = new FileStream(destinationName, FileMode.Create, FileAccess.Write);

                if (targetStream != null)
                {
                    IntPtr pBytsRead = Marshal.AllocHGlobal(4);

                    byte[] buffer = new byte[1024];

                    int bytesRead;

                    do
                    {
                        sourceStream.Read(buffer, 1024, pBytsRead);

                        bytesRead = Marshal.ReadInt32(pBytsRead);
                        targetStream.Write(buffer, 0, bytesRead);
                    } while (bytesRead > 0);


                    Marshal.FreeHGlobal(pBytsRead);

                    targetStream.Close();
                    file.parentDevice.Disconnect();

                    return true;
                }
                else
                    throw new Exception("Unable to open destination file for writing");
            }
            else
                throw new Exception("File does not exist");

            return false;
        }


        private PortableDeviceObject GetCachedFile(string name)
        {
            PortableDeviceObject result;

            result = null;


            if (recentFiles.ContainsKey(name) == true)
            {
                result = recentFiles[name];
            }


            return result;
        }

        public bool CopyFileToDevice(string sourceName, string destinationName, bool deleteOriginal = true)
        {
            bool result;

            PortableDeviceFile file;

            FileStream fs;

            PortableDeviceFolder destinationFolder;
            IStream stream;



            destinationFolder = this.GetFolder(GetDirectoryNameUnsafe(destinationName));

            if (destinationFolder != null)
            {
                result = false;
                fs = new FileStream(sourceName, FileMode.Open, FileAccess.Read);

                if (fs != null)
                {
                    if (deleteOriginal == true)
                    {
                        file = GetCachedFile(destinationName) as PortableDeviceFile;

                        if (file == null)
                            file = GetFile(destinationName);

                        if (file != null)
                            DeleteFile(destinationName);
                    }

                    stream = CreateAndOpenFile(destinationName, fs.Length);

                    if (stream != null)
                    {
                        //    IPortableDeviceResources resources;
                        //    file.content.Transfer(out resources);

                        destinationFolder.parentDevice.Connect();

                        //PortableDeviceApiLib.IStream wpdStream;
                        //uint optimalTransferSize = 0;

                        //resources.GetStream(file.Id, ref PortableDeviceConstants.WPD_RESOURCE_DEFAULT, PortableDeviceConstants.STGM_WRITE, ref optimalTransferSize, out wpdStream);


                        _ULARGE_INTEGER cunt;

                        //cunt.QuadPart = (ulong)fs.Length;

                        //stream.SetSize(cunt);


                        System.Runtime.InteropServices.ComTypes.IStream sourceStream = (System.Runtime.InteropServices.ComTypes.IStream)stream;
                        //sourceStream.Commit(0);

                        IntPtr pBytesWritten = Marshal.AllocHGlobal(4);

                        byte[] buffer = new byte[1024];

                        int bytesRead,
                            bytesWritten;


                        do
                        {
                            bytesRead = fs.Read(buffer, 0, buffer.Length);

                            sourceStream.Write(buffer, bytesRead, pBytesWritten);

                            bytesWritten = Marshal.ReadInt32(pBytesWritten);

                        } while (bytesRead > 0 && bytesWritten > 0);


                        sourceStream.Commit(0);

                        Marshal.FreeHGlobal(pBytesWritten);

                        fs.Close();
                        destinationFolder.parentDevice.Disconnect();

                        result = true;
                    }
                    else
                        throw new Exception("Unable to open file or create destination file");
                }
                else
                    throw new Exception("Unable to open source file for reading");
            }
            else
                throw new Exception("Destination folder does not exist");


            return result;
        }

        private IStream CreateAndOpenFile(string fileName, long fileSize)
        {
            string targetFolder,
                   rawFileName;

            PortableDeviceFolder parentFolder;

            PortableDeviceFile result;


            IStream stream;


            stream = null;
            targetFolder = GetDirectoryNameUnsafe(fileName);
            rawFileName = fileName.Split('\\').Last();//Path.GetFileName(fileName);

            parentFolder = GetFolder(targetFolder);

            if (parentFolder != null)
            {
                //parentFolder.RefreshContent();

                //result = parentFolder.Files.Where(x => x.Name.ToLower() == rawFileName.ToLower()).FirstOrDefault() as PortableDeviceFile;

                //if (result == null)
                //{
                    IPortableDeviceProperties deviceProperties;
                    parentFolder.content.Properties(out deviceProperties);

                    IPortableDeviceValues deviceValues = (IPortableDeviceValues)new PortableDeviceTypesLib.PortableDeviceValues();


                    deviceValues.SetStringValue(ref PortableDeviceConstants.WPD_OBJECT_PARENT_ID, parentFolder.Id);
                    deviceValues.SetStringValue(ref PortableDeviceConstants.WPD_OBJECT_NAME, rawFileName);
                    deviceValues.SetStringValue(ref PortableDeviceConstants.WPD_OBJECT_ORIGINAL_FILE_NAME, rawFileName);
                    deviceValues.SetUnsignedLargeIntegerValue(ref PortableDeviceConstants.WPD_OBJECT_SIZE, (ulong)fileSize);

                    // deviceValues.SetGuidValue(ref PortableDeviceConstants.WPD_OBJECT_CONTENT_TYPE, ref PortableDeviceConstants.WPD_CONTENT_TYPE_GENERIC_FILE);


                    //string objectId = string.Empty;

                    parentFolder.parentDevice.Connect();
                    //parentFolder.content.CreateObjectWithPropertiesOnly(deviceValues, ref objectId);
                    parentFolder.content.CreateObjectWithPropertiesAndData(deviceValues, out stream, 0, null);
                    //parentFolder.parentDevice.Disconnect();

                //}
                //else
                //    throw new Exception("File already exists");
            }
            else
                throw new Exception("Unable to find folder for creating file");

            return stream;
        }

        public bool CreateFile(string fileName)
        {
            IStream stream;

            stream = CreateAndOpenFile(fileName, 0);

            if (stream != null)
            {
                stream.Commit(0);

                return true;
            }

            return false;
        }


        public void DeleteObject(IPortableDeviceContent content, string id)
        {
            //IPortableDevicePropVariantCollection objectIdCollection;


            //objectIdCollection = (IPortableDevicePropVariantCollection)(new PortableDeviceTypesLib.PortableDevicePropVariantCollectionClass());

            //objectIdCollection.Add(StringToPropVariant(id));

            //content.Delete(PortableDeviceConstants.PORTABLE_DEVICE_DELETE_NO_RECURSION, objectIdCollection, null);


            var objectIdCollection =
              (IPortableDevicePropVariantCollection)new PortableDeviceTypesLib.PortableDevicePropVariantCollectionClass();

            var propVariantValue = StringToPropVariant(id);
            objectIdCollection.Add(ref propVariantValue);

            // TODO: get the results back and handle failures correctly
            content.Delete(PortableDeviceConstants.PORTABLE_DEVICE_DELETE_NO_RECURSION, objectIdCollection, null);
        }

        public string GetDirectoryNameUnsafe(string path)
        {
            string result;

            List<string> parts;


            parts = path.Split(new char[] { '\\' }).ToList();


            if (parts.Last().Contains(".") == true)
            {
                parts.RemoveAt(parts.Count() - 1);
            }

            result = string.Join("\\", parts.ToArray());

            return result;
        }

        private PortableDeviceFile GetFile(string fileName)
        {
            string folderPath;

            string rawFileName;

            PortableDeviceFile result;


            result = null;

            folderPath = GetDirectoryNameUnsafe(fileName);

            rawFileName = fileName.Split(new char[] { '\\' }, StringSplitOptions.RemoveEmptyEntries).Last();

            PortableDeviceFolder parentFolder;


            parentFolder = this.GetFolder(folderPath);

            if (parentFolder != null)
            {
                parentFolder.RefreshContent();
                result = parentFolder.Files.Where(x => x.Name.ToLower() == rawFileName.ToLower()).FirstOrDefault() as PortableDeviceFile;
            }

            return result;
        }

        private PortableDeviceFolder GetFolder(string name)
        {
            List<string> parts;

            PortableDeviceFolder result,
                                 tempFolder;


            this.Refresh();

            result = null;

            parts = name.Split(new char[] { '\\', '/' }, System.StringSplitOptions.RemoveEmptyEntries).ToList();


            result = GetCachedFile(name) as PortableDeviceFolder;

            if (result != null)
            {
                return result;
            }

            if (parts.Count > 0)
            {
                foreach (PortableDevice d in this)
                {
                    d.Connect();

                    if (d.FriendlyName.ToLower() == parts[0].ToLower())
                    {
                        tempFolder = d.GetContents();
                        tempFolder.RefreshContent();

                        if (parts.Count == 1)
                        {
                            result = tempFolder;
                            d.Disconnect();
                            break;
                        }

                        for (int i = 1; i < parts.Count; i++)
                        {
                            bool found = false;

                            for (int k = 0; k < tempFolder.Files.Count; k++)
                            {
                                PortableDeviceFolder f = tempFolder.Files[k] as PortableDeviceFolder;

                                if (f is PortableDeviceFolder == true)
                                {
                                    if (parts[i].ToLower() == f.Name.ToLower())
                                    {
                                        found = true;
                                        tempFolder = f;
                                        tempFolder.RefreshContent();


                                        if (i == parts.Count - 1)
                                        {
                                            result = f as PortableDeviceFolder;

                                            if(recentFiles.ContainsKey(name) == false)
                                                recentFiles.Add(name, f);

                                            recentFiles[name] = f;
                                        }

                                        break;
                                    }
                                }
                            }

                            if (found == false)
                                break;
                        }

                    }

                    d.Disconnect();
                }
            }

            return result;
        }

        public static tag_inner_PROPVARIANT StringToPropVariant(string value)
        {
            // Tried using the method suggested here:
            // http://blogs.msdn.com/b/dimeby8/archive/2007/01/08/creating-wpd-propvariants-in-c-without-using-interop.aspx
            // However, the GetValue fails (Element Not Found) even though we've just added it.
            // So, I use the alternative (and I think more "correct") approach below.

            var pvSet = new PortableDeviceConstants.PropVariant
            {
                variantType = PortableDeviceConstants.VT_LPWSTR,
                pointerValue = Marshal.StringToCoTaskMemUni(value)
            };

            // Marshal our definition into a pointer
            var ptrValue = Marshal.AllocHGlobal(Marshal.SizeOf(pvSet));
            Marshal.StructureToPtr(pvSet, ptrValue, false);

            // Marshal pointer into the interop PROPVARIANT 
            return (tag_inner_PROPVARIANT)Marshal.PtrToStructure(ptrValue, typeof(tag_inner_PROPVARIANT));
        }
    }
}