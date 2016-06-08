using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Rainmeter;

namespace APCPlugin
{
    internal class Measure
    {

        static API api;
        static string dllDir;
        static Type UsbDeviceType;
        static Type UsbDeviceFinderType;
        static Type UsbSetupPacketType;
        dynamic APC;

        internal Measure(IntPtr rm) {
            api = new API(rm);
            dllDir = api.ReadPath("DllFolder", "");
            AppDomain.CurrentDomain.AssemblyResolve += AssemblyDependsResolve;
            try
            {
                var dll = Assembly.LoadFile(Path.Combine(dllDir, "LibUsbDotNet.dll"));
                UsbDeviceType = dll.GetType("LibUsbDotNet.UsbDevice", true);
                UsbDeviceFinderType = dll.GetType("LibUsbDotNet.Main.UsbDeviceFinder", true);
                UsbSetupPacketType = dll.GetType("LibUsbDotNet.Main.UsbSetupPacket", true);
                API.Log(API.LogType.Notice, $"Successfully loaded DLLs");
            }
            catch (Exception ex)
            {
                API.Log(API.LogType.Error, $"Could not find DLL or its types. Exception: " + ex);
                throw;
            }
            APC = UsbDeviceType.InvokeMember(
                "OpenUsbDevice",
                BindingFlags.InvokeMethod,
                null,
                null,
                new object[] {
                    Activator.CreateInstance(
                        UsbDeviceFinderType,
                        new object[] {
                            1309,
                            2
                        }
                    )
                }
            );
        }

        internal void Reload(Rainmeter.API api, ref double maxValue)
        {
            APC = UsbDeviceType.InvokeMember(
                "OpenUsbDevice",
                BindingFlags.InvokeMethod,
                null,
                null,
                new object[] {
                    Activator.CreateInstance(
                        UsbDeviceFinderType,
                        new object[] {
                            1309,
                            2
                        }
                    )
                }
            );
        }

        internal double Update()
        {
            if (!object.ReferenceEquals(null, APC))
            {
                int transferred = 0;
                byte[] buffer = new byte[2];
                
                dynamic setup = Activator.CreateInstance(
                    UsbSetupPacketType,
                    new object[] {
                        (Byte)0xA1,
                        (Byte)0x01,
                        (Int16)0x0350,
                        (Int16)0,
                        (Int16)0x0005
                    }
                );
                
                UsbDeviceType.InvokeMember(
                    "ControlTransfer",
                    BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance,
                    null,
                    APC,
                    new object[] {
                        setup, buffer, 0x0040, (Int32) transferred
                    }
                );
                return 865 * buffer[1] / 100;
            } else {
                API.Log(API.LogType.Error, $"Unable to open connection to APC");
                return 0;
            }
            
        }

        static Assembly AssemblyDependsResolve(object sender, ResolveEventArgs e)
        {
            string dllName = e.Name.Split(',')[0] + ".dll";
            string dllPath = Path.Combine(dllDir, dllName);
            if (!File.Exists(dllPath))
            {
                return null;
            }
            return Assembly.LoadFrom(dllPath);
        }
    }

    public static class Plugin
    {
        [DllExport]
        public static void Initialize(ref IntPtr data, IntPtr rm)
        {
            data = GCHandle.ToIntPtr(GCHandle.Alloc(new Measure(rm)));
        }

        [DllExport]
        public static void Finalize(IntPtr data)
        {
            GCHandle.FromIntPtr(data).Free();
        }

        [DllExport]
        public static void Reload(IntPtr data, IntPtr rm, ref double maxValue)
        {
            Measure measure = (Measure)GCHandle.FromIntPtr(data).Target;
            measure.Reload(new Rainmeter.API(rm), ref maxValue);
        }

        [DllExport]
        public static double Update(IntPtr data)
        {
            Measure measure = (Measure)GCHandle.FromIntPtr(data).Target;
            return measure.Update();
        }

    }
}
