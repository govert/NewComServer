using System.Reflection;
using System;
using System.Runtime.InteropServices;
using System.Text;
using ExcelDna.ComInterop;
using ExcelDna.Integration;
using System.Collections;
using Microsoft.Win32;

namespace NewComServer
{
    [ComVisible(false)]
    public class ExcelAddin : IExcelAddIn
    {
        public void AutoOpen()
        {
            // ComServer.DllRegisterServer();
        }
        public void AutoClose()
        {
            // ComServer.DllUnregisterServer();
        }
    }




    public static class Functions
    {
        [ExcelFunction]
        public static object DnaComServerHello()
        {
            return "Hello from DnaComServer!";
        }

        [ExcelFunction("Dumps the list of COM server classes that were found at initialization")]
        public static object DnaDumpComServerTypes()
        {
            StringBuilder sb = new StringBuilder();

            // Get the type of the ComServer class.
            Type comServerType = Type.GetType("ExcelDna.ComInterop.ComServer, ExcelDna.Integration");

            // Get the field info of the registeredComClassTypes field.
            FieldInfo registeredComClassTypesField = comServerType.GetField("registeredComClassTypes", BindingFlags.NonPublic | BindingFlags.Static);

            // Get the value of the field. This will be the list of ExcelComClassType objects.
            var registeredComClassTypes = (IList)registeredComClassTypesField.GetValue(null);

            // Iterate through the list.
            foreach (var item in registeredComClassTypes)
            {
                Type itemType = item.GetType();
                FieldInfo[] fields = itemType.GetFields(BindingFlags.Public | BindingFlags.Instance);

                foreach (FieldInfo field in fields)
                {
                    string name = field.Name;
                    object value = field.GetValue(item);
                    sb.AppendLine($"{name}: {value}");
                }

                sb.AppendLine("-----------");
            }

            var result = sb.ToString();
            return result;
        }
        
            // Look at both the user and machine hive, in case the COM server was registered for the current user.
            // Find the ProgId, from there the ClsId, and then the registry entries for the ClsId.
            // Show the complete subtrees from ProgId and ClsId, and show where they are not present.
            // Look for both 32-bit and 64-bit entries
            // Dump the result as a nicely formatted string showing the hierarchy
        [ExcelFunction("Dumps the registry entries for a COM server class")]
        public static object DnaDumpComServerRegistry(string progId)
        {
            if (string.IsNullOrEmpty(progId))
                return ExcelError.ExcelErrorValue;

            StringBuilder result = new StringBuilder();
            RegistryHive[] rootKeys = { RegistryHive.ClassesRoot, RegistryHive.CurrentUser, RegistryHive.LocalMachine };
            RegistryView[] views = { RegistryView.Registry32, RegistryView.Registry64 };

            foreach (var root in rootKeys)
            {
                foreach (var view in views)
                {
                    using (var baseKey = RegistryKey.OpenBaseKey(root, view))
                    {
                        using (var progIdKey = baseKey.OpenSubKey(progId))
                        {
                            if (progIdKey != null)
                            {
                                result.AppendLine($"Found ProgId {progId} at {root}\\{progId}, {view}");
                                DumpRegistryKey(progIdKey, result, "  ");

                                using (var clsIdSubKey = progIdKey.OpenSubKey("CLSID"))
                                {
                                    string clsId = clsIdSubKey?.GetValue("") as string;
                                    if (clsId != null)
                                    {
                                        result.AppendLine($"Found ClsId {clsId} for ProgId {progId}, {view}");

                                        // Check ClsId in all root keys
                                        foreach (var clsRoot in rootKeys)
                                        {
                                            using (var clsBaseKey = RegistryKey.OpenBaseKey(clsRoot, view))
                                            {
                                                string clsKeyPath = clsRoot == RegistryHive.ClassesRoot ? $"CLSID\\{clsId}" : $"Software\\Classes\\CLSID\\{clsId}";
                                                using (var clsIdKey = clsBaseKey.OpenSubKey(clsKeyPath))
                                                {
                                                    if (clsIdKey != null)
                                                    {
                                                        result.AppendLine($"Found ClsId {clsId} at {clsRoot}\\{clsKeyPath}, {view}");
                                                        DumpRegistryKey(clsIdKey, result, "  ");
                                                    }
                                                    else
                                                    {
                                                        result.AppendLine($"ClsId {clsId} not found at {clsRoot}\\{clsKeyPath}, {view}");
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        result.AppendLine($"ClsId not found for ProgId {progId}, {view}");
                                    }
                                }
                            }
                            else
                            {
                                result.AppendLine($"ProgId {progId} not found at {root}\\{progId}, {view}");
                            }
                        }
                    }
                }
            }

            return result.ToString();
        }

        private static void DumpRegistryKey(RegistryKey key, StringBuilder result, string indent)
        {
            foreach (string valueName in key.GetValueNames())
            {
                object value = key.GetValue(valueName);
                string valueAsString = value as string ?? "(non-string value)";
                result.AppendLine($"{indent}{valueName} = {valueAsString}");
            }

            foreach (string subKeyName in key.GetSubKeyNames())
            {
                using (RegistryKey subKey = key.OpenSubKey(subKeyName))
                {
                    result.AppendLine($"{indent}{subKeyName}:");
                    DumpRegistryKey(subKey, result, indent + "  ");
                }
            }
        }

    }
}