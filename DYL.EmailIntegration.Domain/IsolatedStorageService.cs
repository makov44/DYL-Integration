using System;
using System.Collections.Generic;
using System.IO;
using System.IO.IsolatedStorage;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DYL.EmailIntegration.Domain
{
    public static class IsolatedStorageService
    {
        public static void Write(string fileName, byte[] data)
        {
            if (string.IsNullOrEmpty(fileName))
                return;

            var isoStore = IsolatedStorageFile.GetStore(IsolatedStorageScope.User | IsolatedStorageScope.Assembly, null, null);

            if (isoStore.FileExists(fileName))
            {
                using (var isoStream = new IsolatedStorageFileStream(fileName, FileMode.Truncate, isoStore))
                {
                    using (var writer = new BinaryWriter(isoStream))
                    {
                        writer.Write(data);
                    }
                }
            }
            else
            {
                using (var isoStream = new IsolatedStorageFileStream(fileName, FileMode.CreateNew, isoStore))
                {
                    using (var writer = new BinaryWriter(isoStream))
                    {
                        writer.Write(data);
                    }
                }
            }
        }

        public static byte[] Read(string fileName)
        {
            if (string.IsNullOrEmpty(fileName))
                return new byte[0];

            var isoStore = IsolatedStorageFile.GetStore(IsolatedStorageScope.User | IsolatedStorageScope.Assembly, null, null);

            if (isoStore.FileExists(fileName))
            {
                using (var isoStream = new IsolatedStorageFileStream(fileName, FileMode.Open, isoStore))
                {
                    using (var reader = new BinaryReader(isoStream))
                    {
                        return reader.ReadBytes(1024);
                    }
                }
            }
            else
            {
                return new byte[0]; 
            }
        }
    }
}
