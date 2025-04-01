using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace ComUtils.ComHelpers
{
    /// <summary>
    /// Simplified access to the COM Running Object Table.
    /// </summary>
    /// <remarks>
    /// Inspired and guided by https://stackoverflow.com/a/38619839/1210847
    /// </remarks>
    public class RunningObjectTable
    {
        [DllImport("ole32.dll")]
        static extern int GetRunningObjectTable(uint reserved, out IRunningObjectTable rot);

        [DllImport("ole32.dll", PreserveSig = false)]
        static extern void CreateBindCtx(uint reserved, out IBindCtx ppbc);

        private readonly IRunningObjectTable mRot;
        public RunningObjectTable()
        {
            Marshal.ThrowExceptionForHR(GetRunningObjectTable(0, out mRot));
        }

        public IEnumerable<RunningObjectTableEntry> Entries
        {
            get
            {
                mRot.EnumRunning(out var comIterator);
                if (comIterator != null)
                {
                    try
                    {
                        Marshal.AddRef(Marshal.GetIUnknownForObject(comIterator));
                        IntPtr fetched = IntPtr.Zero;
                        CreateBindCtx(0, out IBindCtx bindContext);

                        var objRefs = new IMoniker[1];

                        while (comIterator.Next(1, objRefs, fetched) == 0)
                        {
                            Marshal.ThrowExceptionForHR(mRot.GetObject(objRefs[0], out dynamic comObj));
                            var moni = objRefs[0];
                            
                            moni.GetDisplayName(bindContext, moni, out string dispName);
                            moni.GetClassID(out Guid classID);
                            RunningObjectTableEntryType moniEnumType;
                            if (moni.IsSystemMoniker(out int moniType) == 0)
                            {
                                moniEnumType = (RunningObjectTableEntryType)moniType;
                            }
                            else
                            {
                                moniEnumType = RunningObjectTableEntryType.Unknown;
                            }
                            yield return new RunningObjectTableEntry(comObj, dispName, classID, moniEnumType);
                        }
                    }
                    finally
                    {
                        Marshal.Release(Marshal.GetIUnknownForObject(comIterator));
                    }
                }
                else
                {
                    throw new InvalidOperationException("Could not get the running object table enumerator");
                }
            }
        }

    }

}
