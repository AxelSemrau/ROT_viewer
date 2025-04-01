using System;
using System.Security.AccessControl;
using Microsoft.Win32;

namespace ComUtils.ComHelpers
{
    public enum RunningObjectTableEntryType
    {
        Custom,
        GenericComposite,
        File,
        Anti,
        Item,
        Pointer,
        Class,
        ObjRef,
        Session,
        Elevation,
        Unknown = -1
    }

    public class RunningObjectTableEntry
    {
        public RunningObjectTableEntry(dynamic comObj, string dispName, Guid moniClassID, RunningObjectTableEntryType type)
        {
            Object = comObj;
            mDisplayName = dispName;
            MonikerClassID = moniClassID;
            Type = type;
        }

        private bool mResolvedName;
        private string mDisplayName;

        public Guid MonikerClassID { get; }

        public dynamic Object { get; }
        public string DisplayName => GetDisplayName();
        public RunningObjectTableEntryType Type { get; }

        /// <summary>
        /// Entries of type "Item" have a CLSID as the display name which we try to translate to something human-readable.
        /// </summary>
        /// <returns></returns>
        private string GetDisplayName()
        {
            if (Type == RunningObjectTableEntryType.Item && !mResolvedName)
            {
                mResolvedName = true;
                if (mDisplayName.StartsWith("!{"))
                {
                    // Looking for the CLSID in the registry
                    mDisplayName = mDisplayName.Remove(0, 1);
                    using (var reg = RegistryKey.OpenBaseKey(RegistryHive.ClassesRoot, RegistryView.Registry32))
                    {
                        using (var sk = reg.OpenSubKey($"CLSID\\{mDisplayName}", RegistryRights.QueryValues))
                        {
                            if (sk != null)
                            {
                                mDisplayName = sk.GetValue(null, $"CLSID: {mDisplayName}") as string;
                            }
                        }
                    }
                }
            }
            return mDisplayName;
        }

        public override string ToString() => $"{DisplayName} {Type}";
    }
}