# C-Sharp-Simplier-MTP
Simplifies the MTP protocol to the familiar System.IO function names 


Example of how to copy all files from an android phone


            MobileDeviceManager deviceManager = new MobileDeviceManager();

            foreach (string drive in deviceManager.GetLogicalDrives())
            {
                foreach (string dir in deviceManager.GetDirectories(drive, ".*", true))
                {
                    Console.WriteLine("Directory: " + dir);

                    foreach (string file in deviceManager.GetFiles(dir, ".*", true))
                    {
                        Console.WriteLine("File: " + file);

                        deviceManager.CopyFileFromDevice(file, "C:\\" + System.IO.Path.GetFileName(file));
                    }
                }
            }
