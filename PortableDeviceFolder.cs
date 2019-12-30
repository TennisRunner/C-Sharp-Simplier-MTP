using System.Collections.Generic;
using System.Collections.ObjectModel;
using PortableDeviceApiLib;
using System;
using PortableDevicesConstants;

namespace PortableDevices
{
    public class PortableDeviceFolder : PortableDeviceObject
    {


        public PortableDeviceFolder(string id, string name) : base(id, name)
        {
            this.Files = new List<PortableDeviceObject>();
        }

        public void RefreshContent()
        {
            PortableDevice parentDevice = this.parentDevice;

            this.Files.Clear();


            if (this.parentDevice != null)
            {
                this.parentDevice.Connect();
            }

            EnumerateContents(this, parentDevice, ref this.content, this);


            if (this.parentDevice != null)
            {
                this.parentDevice.Disconnect();
            }
        }

        private void EnumerateContents(PortableDeviceFolder parentFolder, PortableDevice parentDevice, ref IPortableDeviceContent content, PortableDeviceFolder parent)
        {
            // Get the properties of the object
            IPortableDeviceProperties properties;
            content.Properties(out properties);

            // Enumerate the items contained by the current object
            IEnumPortableDeviceObjectIDs objectIds;
            content.EnumObjects(0, parent.Id, null, out objectIds);

            uint fetched = 0;
            do
            {
                fetched = 0;

                string objectId;

                objectIds.Next(1, out objectId, ref fetched);
                if (fetched > 0)
                {
                    var currentObject = WrapObject(properties, objectId);
                    
                    currentObject.content = content;
                    currentObject.parentDevice = parentDevice;
                    currentObject.parentObject = parentFolder;
                    parent.Files.Add(currentObject);
                }
            } while (fetched > 0);
        }

        private static PortableDeviceObject WrapObject(IPortableDeviceProperties properties, string objectId)
        {
            IPortableDeviceKeyCollection keys;
            properties.GetSupportedProperties(objectId, out keys);

            IPortableDeviceValues values;
            properties.GetValues(objectId, keys, out values);


            string name = string.Empty;
            values.GetStringValue(ref PortableDeviceConstants.WPD_OBJECT_NAME, out name);
            
            // Get the type of the object
            Guid contentType;
            values.GetGuidValue(PortableDeviceConstants.WPD_OBJECT_CONTENT_TYPE, out contentType);


            try
            {
                values.GetStringValue(PortableDeviceConstants.WPD_OBJECT_ORIGINAL_FILE_NAME, out name);
            }
            catch (Exception x)
            {
                x.ToString();
            }

            if (contentType == PortableDeviceConstants.WPD_CONTENT_TYPE_FOLDER || contentType == PortableDeviceConstants.WPD_CONTENT_TYPE_FUNCTIONAL_OBJECT)
                return new PortableDeviceFolder(objectId, name);

            return new PortableDeviceFile(objectId, name);
        }


        public IList<PortableDeviceObject> Files { get; set; }
    }
}