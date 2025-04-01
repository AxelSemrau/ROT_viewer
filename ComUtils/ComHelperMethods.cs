using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Win32;

namespace ComUtils
{
    public static class ComHelperMethods
    {
        public static bool IsClassRegistered<T>()
        {
            return IsClassRegistered(typeof(T));
        }
        public static bool IsClassRegistered(Type type)
        {
            var origType = type;
            var typesAttributes = type.GetCustomAttributes(false);
            var coClass = typesAttributes.OfType<CoClassAttribute>().FirstOrDefault();
            // if we find a CoClass attribute here, this is not the actual type we are looking for.
            if (coClass != null)
            {
                type = coClass.CoClass;
                if (type == typeof(object))
                {
                    throw new InvalidOperationException($"Could not determine CoClass type for {origType.Name} - embedded interop types used for reference?");
                }

                typesAttributes = type.GetCustomAttributes(false);
            }
            var clsid = typesAttributes.OfType<GuidAttribute>().FirstOrDefault()?.Value;
            if (string.IsNullOrEmpty(clsid))
            {
                clsid = typesAttributes.OfType<AxHost.ClsidAttribute>().FirstOrDefault()?.Value;
            }
            return IsClsidRegistered(clsid);
        }

        private static bool IsClsidRegistered(string clsid)
        {
            if (string.IsNullOrEmpty(clsid))
            {
                return false;
            }
            var hklm = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32);
            if (!clsid.StartsWith("{"))
            {
                clsid = "{" + clsid;
            }

            if (!clsid.EndsWith("}"))
            {
                clsid = clsid + "}";
            }
            using (var regKey = hklm.OpenSubKey($@"Software\Classes\CLSID\{clsid}", false))
            {
                return regKey != null;
            }
        }

        /// <summary>
        /// Increase the reference counter of the COM object.
        /// </summary>
        /// <remarks>
        /// See CHRON-875: If this is not done, the LabSolutions software closes down as soon as Chronos releases its COM object.
        /// </remarks>
        /// <param name="someObj"></param>
        public static void IncreaseRefCount(object someObj)
        {
            if (!ObjectsWithIncreasedCounters.Contains(someObj))
            {
                var pUnknown = Marshal.GetIUnknownForObject(someObj);
                // See CHRON-875: Prevent the application from closing because of the released COM object when Chronos terminates.
                Marshal.AddRef(pUnknown);
                ObjectsWithIncreasedCounters.Add(someObj);
            }
        }

        /// <summary>
        /// If we increased the counter, we should offer a way to decrease it.
        /// </summary>
        /// <param name="someObj"></param>
        public static void DecreaseRefCount(object someObj)
        {
            // do not check if we have it in ObjectsWithIncreasedCounters 
            // it could be possible that we got it via the ROT with an increased counter from an earlier Chronos instance.
            var pUnknown = Marshal.GetIUnknownForObject(someObj);
            var tries = 0;
            while (tries++ < 100 && Marshal.Release(pUnknown) > 0)
            {
                // nothing
            }
            // object is dead, don't try to kill it again on GC collection
            GC.SuppressFinalize(someObj);
            ObjectsWithIncreasedCounters.Remove(someObj);
        }
        /// <summary>
        /// Cache the objects for which we know that the counter was increased to avoid multiple reference counter additions.
        /// </summary>
        private static readonly HashSet<object> ObjectsWithIncreasedCounters = new HashSet<object>();

        public static string GetFileFromProgID(string progID)
        {
            using (var key = RegistryKey.OpenBaseKey(RegistryHive.ClassesRoot, RegistryView.Registry32))
            {
                using (var subkey = key.OpenSubKey($"{progID}\\CLSID", false))
                {
                    if (subkey != null)
                    {
                        var clsid = subkey.GetValue("", "").ToString();
                        if (!string.IsNullOrEmpty(clsid))
                        {
                            using (var serverkey = key.OpenSubKey($"CLSID\\{clsid}\\LocalServer32", false))
                            {
                                if (serverkey != null)
                                {
                                    // quotes will do no good for further uses
                                    return serverkey.GetValue("", "").ToString().Replace("\"", "");
                                }
                            }
                        }
                    }
                }
                // sometimes the ProgID entries don't point directly t the CLSID
                using (var clsidKey = key.OpenSubKey("CLSID", false))
                {
                    if (clsidKey != null)
                    {
                        foreach (var subkeyName in clsidKey.GetSubKeyNames())
                        {
                            using (var progIdKey = clsidKey.OpenSubKey($"{subkeyName}\\ProgID", false))
                            {
                                if (progIdKey?.GetValue("", "")?.ToString() == progID)
                                {
                                    using (var serverKey = clsidKey.OpenSubKey($"{subkeyName}\\InprocServer32", false))
                                    {
                                        if (serverKey != null)
                                        {
                                            return serverKey.GetValue("", "").ToString().Replace("\"", "");
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            return "";
        }


    }
}
