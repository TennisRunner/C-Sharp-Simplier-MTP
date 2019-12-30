

using PortableDeviceApiLib;
using PortableDeviceTypesLib;

namespace PortableDevices
{
    public abstract class PortableDeviceObject
    {
        protected PortableDeviceObject(string id, string name)
        {
            this.Id = id;
            this.Name = name;
        }

        public string GetAbsolutePath()
        {
            string result;

            PortableDeviceObject temp;

            result = this.Name;

            temp = this;

            do
            {
                temp = temp.parentObject;

                if (temp != null && temp.Id != "DEVICE")
                {
                    result = temp.Name + "\\" + result;
                }

            } while (temp != null);

            result = this.parentDevice.FriendlyName + "\\" + result;

            return result;
        }

        public string Id { get; private set; }

        public string Name { get; private set; }

        public IPortableDeviceContent content;

        internal PortableDevice parentDevice;

        internal PortableDeviceObject parentObject;
    }
}