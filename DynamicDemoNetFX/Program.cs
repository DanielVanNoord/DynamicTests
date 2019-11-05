using System;
using System.Dynamic;
using System.Reflection;
using System.Runtime.InteropServices;

namespace DynamicDemoNetFX
{
    internal class Program
    {
        private static string DotNetFXDriver = "ASCOM.Simulator.Telescope";
        private static string DotNetCoreDriver = "ASCOM.SimulatorCore.Telescope";

        private static void Main(string[] args)
        {
            Type NetFXDriverType = Type.GetTypeFromProgID(DotNetFXDriver);
            Type NetCoreDriverType = Type.GetTypeFromProgID(DotNetCoreDriver);


            RunTest(() => TestWithInterfaces(Activator.CreateInstance(NetFXDriverType)), "Test loading a NetFX driver with Interface on ");

            RunTest(() => TestWithInterfaces(Activator.CreateInstance(NetCoreDriverType)), "Test loading a NetCore driver with Interface on ");

            RunTest(() => TestWithDynamic(Activator.CreateInstance(NetFXDriverType)), "Test loading a NetFX driver with dynamic only ");

            RunTest(() => TestWithDynamic(Activator.CreateInstance(NetCoreDriverType)), "Test loading a NetCore driver with dynamic only ");

            RunTest(() => TestWithDynamic(COMObject.CreateInstance(NetFXDriverType)), "Test loading a NetFX driver with dynamic and COM Wrapper ");

            RunTest(() => TestWithDynamic(COMObject.CreateInstance(NetCoreDriverType)), "Test loading a NetCore driver with dynamic and COM Wrapper ");
        }

        private static void RunTest(Action action, string message)
        {

            try
            {
                Console.Write(message);
                Console.WriteLine(System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription);
                action();
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Test Passed");

            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Test Failed");
                Console.WriteLine(ex.Message);

            }
            Console.ResetColor();

        }

        private static void TestWithInterfaces(object TestObject)
        {

            ASCOM.DeviceInterface.ITelescopeV3 driver = TestObject as ASCOM.DeviceInterface.ITelescopeV3;

            driver.Connected = true;

            driver.Connected = false;
        }

        private static void TestWithDynamic(object TestObject)
        {
            dynamic driver = TestObject;

            driver.Connected = true;

            driver.Connected = false;
        }
    }

    /// <summary>
    /// Com wrapper for Net Core objects from https://github.com/dotnet/coreclr/issues/24246
    /// </summary>
    internal class COMObject : DynamicObject
    {
        private readonly object instance;

        public static COMObject CreateObject(string progID)
        {
            return new COMObject(Activator.CreateInstance(Type.GetTypeFromProgID(progID, true)));
        }

        public static COMObject CreateInstance(Type type)
        {
            return new COMObject(Activator.CreateInstance(type));
        }

        public COMObject(object instance)
        {
            this.instance = instance;
        }

        public override bool TryGetMember(GetMemberBinder binder, out object result)
        {
            result = instance.GetType().InvokeMember(
                binder.Name,
                BindingFlags.GetProperty,
                Type.DefaultBinder,
                instance,
                new object[] { }
            );
            return true;
        }

        public override bool TrySetMember(SetMemberBinder binder, object value)
        {
            instance.GetType().InvokeMember(
                binder.Name,
                BindingFlags.SetProperty,
                Type.DefaultBinder,
                instance,
                new object[] { value }
            );
            return true;
        }

        public override bool TryInvokeMember(InvokeMemberBinder binder, object[] args, out object result)
        {
            result = instance.GetType().InvokeMember(
                binder.Name,
                BindingFlags.InvokeMethod,
                Type.DefaultBinder,
                instance,
                args
            );
            return true;
        }
    }
}