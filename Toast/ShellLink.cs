using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Security;
using System.Text;

namespace Toast
{
    public class ShellLink : IDisposable
    {
        #region Win32 and COM

        // IShellLink Interface
        [ComImport,
         InterfaceType(ComInterfaceType.InterfaceIsIUnknown),
         Guid("000214F9-0000-0000-C000-000000000046")]
        private interface IShellLinkW
        {
            uint GetPath([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszFile,
                         int cchMaxPath, ref WIN32_FIND_DATAW pfd, uint fFlags);
            uint GetIDList(out IntPtr ppidl);
            uint SetIDList(IntPtr pidl);
            uint GetDescription([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszName,
                                int cchMaxName);
            uint SetDescription([MarshalAs(UnmanagedType.LPWStr)] string pszName);
            uint GetWorkingDirectory([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszDir,
                                     int cchMaxPath);
            uint SetWorkingDirectory([MarshalAs(UnmanagedType.LPWStr)] string pszDir);
            uint GetArguments([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszArgs,
                              int cchMaxPath);
            uint SetArguments([MarshalAs(UnmanagedType.LPWStr)] string pszArgs);
            uint GetHotKey(out ushort pwHotkey);
            uint SetHotKey(ushort wHotKey);
            uint GetShowCmd(out int piShowCmd);
            uint SetShowCmd(int iShowCmd);
            uint GetIconLocation([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszIconPath,
                                 int cchIconPath, out int piIcon);
            uint SetIconLocation([MarshalAs(UnmanagedType.LPWStr)] string pszIconPath, int iIcon);
            uint SetRelativePath([MarshalAs(UnmanagedType.LPWStr)] string pszPathRel,
                                 uint dwReserved);
            uint Resolve(IntPtr hwnd, uint fFlags);
            uint SetPath([MarshalAs(UnmanagedType.LPWStr)] string pszFile);
        }

        [SuppressUnmanagedCodeSecurity]
        [DllImport("ole32.dll")]
        public extern static int PropVariantClear(ref PropVariant pvar);

        // ShellLink CoClass (ShellLink object)
        [ComImport,
         ClassInterface(ClassInterfaceType.None),
         Guid("00021401-0000-0000-C000-000000000046")]
        private class CShellLink { }

        // WIN32_FIND_DATAW Structure
        [StructLayout(LayoutKind.Sequential, Pack = 4, CharSet = CharSet.Unicode)]
        private struct WIN32_FIND_DATAW
        {
            public uint dwFileAttributes;
            public FILETIME ftCreationTime;
            public FILETIME ftLastAccessTime;
            public FILETIME ftLastWriteTime;
            public uint nFileSizeHigh;
            public uint nFileSizeLow;
            public uint dwReserved0;
            public uint dwReserved1;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = MAX_PATH)]
            public string cFileName;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 14)]
            public string cAlternateFileName;
        }

        // IPropertyStore Interface
        [ComImport,
         InterfaceType(ComInterfaceType.InterfaceIsIUnknown),
         Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99")]
        private interface IPropertyStore
        {
            uint GetCount([Out] out uint cProps);
            uint GetAt([In] uint iProp, out PropertyKey pkey);
            uint GetValue([In] ref PropertyKey key, [Out] PropVariant pv);
            uint SetValue([In] ref PropertyKey key, [In] PropVariant pv);
            uint Commit();
        }

        // PropertyKey Structure
        // Narrowed down from PropertyKey.cs of Windows API Code Pack 1.1 
        [StructLayout(LayoutKind.Sequential, Pack = 4)]
        private struct PropertyKey
        {
            #region Fields

            private Guid formatId;    // Unique GUID for property
            private Int32 propertyId; // Property identifier (PID)

            #endregion

            #region Public Properties

            public Guid FormatId
            {
                get
                {
                    return formatId;
                }
            }

            public Int32 PropertyId
            {
                get
                {
                    return propertyId;
                }
            }

            #endregion

            #region Constructor

            public PropertyKey(Guid formatId, Int32 propertyId)
            {
                this.formatId = formatId;
                this.propertyId = propertyId;
            }

            public PropertyKey(string formatId, Int32 propertyId)
            {
                this.formatId = new Guid(formatId);
                this.propertyId = propertyId;
            }

            #endregion
        }

        [StructLayout(LayoutKind.Sequential)]
        internal struct CLIPDATA
        {
            public uint cbSize;         //ULONG
            public int ulClipFmt;       //long
            public IntPtr pClipData;    //BYTE*
        }

        // Credit: http://blogs.msdn.com/b/adamroot/archive/2008/04/11/interop-with-propvariants-in-net.aspx
        /// <summary>
        /// Represents the OLE struct PROPVARIANT.
        /// </summary>
        /// <remarks>
        /// Must call Clear when finished to avoid memory leaks. If you get the value of
        /// a VT_UNKNOWN prop, an implicit AddRef is called, thus your reference will
        /// be active even after the PropVariant struct is cleared.
        /// </remarks>
        [StructLayout(LayoutKind.Sequential)]
        public sealed class PropVariant : IDisposable
        {
            #region struct fields

            // The layout of these elements needs to be maintained.
            //
            // NOTE: We could use LayoutKind.Explicit, but we want
            //       to maintain that the IntPtr may be 8 bytes on
            //       64-bit architectures, so we'll let the CLR keep
            //       us aligned.
            //
            // NOTE: In order to allow x64 compat, we need to allow for
            //       expansion of the IntPtr. However, the BLOB struct
            //       uses a 4-byte int, followed by an IntPtr, so
            //       although the p field catches most pointer values,
            //       we need an additional 4-bytes to get the BLOB
            //       pointer. The p2 field provides this, as well as
            //       the last 4-bytes of an 8-byte value on 32-bit
            //       architectures.

            // This is actually a VarEnum value, but the VarEnum type
            // shifts the layout of the struct by 4 bytes instead of the
            // expected 2.
            ushort vt;
            ushort wReserved1;
            ushort wReserved2;
            ushort wReserved3;
            public IntPtr p;
            int p2;

            #endregion // struct fields

            #region union members

            sbyte cVal // CHAR cVal;
            {
                get { return (sbyte)GetDataBytes()[0]; }
            }

            byte bVal // UCHAR bVal;
            {
                get { return GetDataBytes()[0]; }
            }

            short iVal // SHORT iVal;
            {
                get { return BitConverter.ToInt16(GetDataBytes(), 0); }
            }

            ushort uiVal // USHORT uiVal;
            {
                get { return BitConverter.ToUInt16(GetDataBytes(), 0); }
            }

            int lVal // LONG lVal;
            {
                get { return BitConverter.ToInt32(GetDataBytes(), 0); }
            }

            uint ulVal // ULONG ulVal;
            {
                get { return BitConverter.ToUInt32(GetDataBytes(), 0); }
            }

            long hVal // LARGE_INTEGER hVal;
            {
                get { return BitConverter.ToInt64(GetDataBytes(), 0); }
            }

            ulong uhVal // ULARGE_INTEGER uhVal;
            {
                get { return BitConverter.ToUInt64(GetDataBytes(), 0); }
            }

            float fltVal // FLOAT fltVal;
            {
                get { return BitConverter.ToSingle(GetDataBytes(), 0); }
            }

            double dblVal // DOUBLE dblVal;
            {
                get { return BitConverter.ToDouble(GetDataBytes(), 0); }
            }

            bool boolVal // VARIANT_BOOL boolVal;
            {
                get { return (iVal == 0 ? false : true); }
            }

            int scode // SCODE scode;
            {
                get { return lVal; }
            }

            decimal cyVal // CY cyVal;
            {
                get { return decimal.FromOACurrency(hVal); }
            }

            DateTime date // DATE date;
            {
                get { return DateTime.FromOADate(dblVal); }
            }

            #endregion // union members

            private byte[] GetBlobData()
            {
                var blobData = new byte[lVal];
                IntPtr pBlobData;

                try
                {
                    switch (IntPtr.Size)
                    {
                        case 4:
                            pBlobData = new IntPtr(p2);
                            break;

                        case 8:
                            pBlobData = new IntPtr(BitConverter.ToInt64(GetDataBytes(), sizeof(int)));
                            break;

                        default:
                            throw new NotSupportedException();
                    }

                    Marshal.Copy(pBlobData, blobData, 0, lVal);
                }
                catch
                {
                    return null;
                }

                return blobData;
            }

            internal CLIPDATA GetCLIPDATA()
            {
                return (CLIPDATA)Marshal.PtrToStructure(p, typeof(CLIPDATA));
            }

            /// <summary>
            /// Gets a byte array containing the data bits of the struct.
            /// </summary>
            /// <returns>A byte array that is the combined size of the data bits.</returns>
            private byte[] GetDataBytes()
            {
                var ret = new byte[IntPtr.Size + sizeof(int)];

                if (IntPtr.Size == 4)
                {
                    BitConverter.GetBytes(p.ToInt32()).CopyTo(ret, 0);
                }
                else if (IntPtr.Size == 8)
                {
                    BitConverter.GetBytes(p2).CopyTo(ret, IntPtr.Size);
                }

                return ret;
            }

            /// <summary>
            /// Called to clear the PropVariant's referenced and local memory.
            /// </summary>
            /// <remarks>
            /// You must call Clear to avoid memory leaks.
            /// </remarks>
            public void Clear()
            {
                // Can't pass "this" by ref, so make a copy to call PropVariantClear with
                PropVariant var = this;
                PropVariantClear(ref var);

                // Since we couldn't pass "this" by ref, we need to clear the member fields manually
                // NOTE: PropVariantClear already freed heap data for us, so we are just setting
                //       our references to null.
                vt = (ushort)VarEnum.VT_EMPTY;
                wReserved1 = wReserved2 = wReserved3 = 0;
                p = IntPtr.Zero;
                p2 = 0;
            }

            /// <summary>
            /// Gets the variant type.
            /// </summary>
            public VarEnum Type
            {
                get { return (VarEnum)vt; }
            }

            /// <summary>
            /// Gets the variant value.
            /// </summary>
            public object Value
            {
                get
                {
                    switch ((VarEnum)vt)
                    {
                        case VarEnum.VT_I1:
                            return cVal;
                        case VarEnum.VT_UI1:
                            return bVal;
                        case VarEnum.VT_I2:
                            return iVal;
                        case VarEnum.VT_UI2:
                            return uiVal;
                        case VarEnum.VT_I4:
                        case VarEnum.VT_INT:
                            return lVal;
                        case VarEnum.VT_UI4:
                        case VarEnum.VT_UINT:
                            return ulVal;
                        case VarEnum.VT_I8:
                            return hVal;
                        case VarEnum.VT_UI8:
                            return uhVal;
                        case VarEnum.VT_R4:
                            return fltVal;
                        case VarEnum.VT_R8:
                            return dblVal;
                        case VarEnum.VT_BOOL:
                            return boolVal;
                        case VarEnum.VT_ERROR:
                            return scode;
                        case VarEnum.VT_CY:
                            return cyVal;
                        case VarEnum.VT_DATE:
                            return date;
                        case VarEnum.VT_FILETIME:
                            if (hVal > 0)
                            {
                                return DateTime.FromFileTime(hVal);
                            }
                            else
                            {
                                return null;
                            }
                        case VarEnum.VT_BSTR:
                            return Marshal.PtrToStringBSTR(p);
                        case VarEnum.VT_LPSTR:
                            return Marshal.PtrToStringAnsi(p);
                        case VarEnum.VT_LPWSTR:
                            return Marshal.PtrToStringUni(p);
                        case VarEnum.VT_UNKNOWN:
                            return Marshal.GetObjectForIUnknown(p);
                        case VarEnum.VT_DISPATCH:
                            return p;
                        case VarEnum.VT_CLSID:
                            return Marshal.PtrToStructure(p, typeof(Guid));
                            //default:
                            //    throw new NotSupportedException("The type of this variable is not support ('" + vt.ToString() + "')");
                    }

                    return null;
                }
            }

            public PropVariant(string value)
            {
                this.vt = (ushort)VarEnum.VT_LPWSTR;
                this.p = Marshal.StringToCoTaskMemUni(value);
                this.p2 = 0;
                this.wReserved1 = 0;
                this.wReserved2 = 0;
                this.wReserved3 = 0;
            }

            public PropVariant(Guid value)
            {
                this.vt = (ushort)VarEnum.VT_CLSID;
                byte[] guid = value.ToByteArray();
                this.p = Marshal.AllocCoTaskMem(guid.Length);
                Marshal.Copy(guid, 0, p, guid.Length);
                this.p2 = 0;
                this.wReserved1 = 0;
                this.wReserved2 = 0;
                this.wReserved3 = 0;
            }

            public PropVariant()
            { }

            public void Dispose()
            {
                //throw new NotImplementedException();
            }
        }

        //PropVariant Class(only for string value)
        // Narrowed down from PropVariant.cs of Windows API Code Pack 1.1
        // Originally from http://blogs.msdn.com/b/adamroot/archive/2008/04/11
        // /interop-with-propvariants-in-net.aspx

        //    [StructLayout(LayoutKind.Explicit)]
        //public sealed class PropVariant : IDisposable
        //    {
        //        #region Fields

        //        [FieldOffset(0)]
        //        ushort valueType;     // Value type 

        //        // [FieldOffset(2)]
        //        // ushort wReserved1; // Reserved field
        //        // [FieldOffset(4)]
        //        // ushort wReserved2; // Reserved field
        //        // [FieldOffset(6)]
        //        // ushort wReserved3; // Reserved field

        //        [FieldOffset(8)]
        //        IntPtr ptr;           // Value

        //        #endregion

        //        #region Public Properties

        //        // Value type (System.Runtime.InteropServices.VarEnum)
        //        public VarEnum VarType
        //        {
        //            get { return (VarEnum)valueType; }
        //            set { valueType = (ushort)value; }
        //        }

        //        // Whether value is empty or null
        //        public bool IsNullOrEmpty
        //        {
        //            get
        //            {
        //                return (valueType == (ushort)VarEnum.VT_EMPTY ||
        //                        valueType == (ushort)VarEnum.VT_NULL);
        //            }
        //        }

        //        // Value (only for string value)
        //        public string Value
        //        {
        //            get
        //            {
        //                return Marshal.PtrToStringUni(ptr);
        //            }
        //        }

        //        #endregion

        //        #region Constructor

        //        public PropVariant()
        //        { }

        //        // Construct with string value
        //        public PropVariant(string value, ushort valueType1 = (ushort)VarEnum.VT_LPWSTR)
        //        {
        //            if (value == null)
        //                throw new ArgumentException("Failed to set value.");

        //            valueType = valueType1;
        //            ptr = Marshal.StringToCoTaskMemUni(value);
        //        }

        //        #endregion

        //        #region Destructor

        //        ~PropVariant()
        //        {
        //            Dispose();
        //        }

        //        public void Dispose()
        //        {
        //            PropVariantClear(this);
        //            GC.SuppressFinalize(this);
        //        }

        //        #endregion
        //    }

        [DllImport("Ole32.dll", PreserveSig = false)]
        private extern static void PropVariantClear([In, Out] PropVariant pvar);

        #endregion

        #region Fields

        private IShellLinkW shellLinkW = null;

        // Name = System.AppUserModel.ID
        // ShellPKey = PKEY_AppUserModel_ID
        // FormatID = 9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3
        // PropID = 5
        // Type = String (VT_LPWSTR)
        private readonly PropertyKey AppUserModelIDKey =
            new PropertyKey("{9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3}", 5);

        private readonly PropertyKey CLSIDKey =
            new PropertyKey("{9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3}", 26);

        private const int MAX_PATH = 260;
        private const int INFOTIPSIZE = 1024;

        private const int STGM_READ = 0x00000000;     // STGM constants
        private const uint SLGP_UNCPRIORITY = 0x0002; // SLGP flags

        #endregion

        #region Private Properties (Interfaces)

        private IPersistFile PersistFile
        {
            get
            {
                IPersistFile PersistFile = shellLinkW as IPersistFile;

                if (PersistFile == null)
                    throw new COMException("Failed to create IPersistFile.");
                else
                    return PersistFile;
            }
        }

        private IPropertyStore PropertyStore
        {
            get
            {
                IPropertyStore PropertyStore = shellLinkW as IPropertyStore;

                if (PropertyStore == null)
                    throw new COMException("Failed to create IPropertyStore.");
                else
                    return PropertyStore;
            }
        }

        #endregion

        #region Public Properties (Minimal)

        // Path of loaded shortcut file
        public string ShortcutFile
        {
            get
            {
                string shortcutFile;

                PersistFile.GetCurFile(out shortcutFile);

                return shortcutFile;
            }
        }

        // Path of target file
        public string TargetPath
        {
            get
            {
                // No limitation to length of buffer string in the case of Unicode though.
                StringBuilder targetPath = new StringBuilder(MAX_PATH);

                WIN32_FIND_DATAW data = new WIN32_FIND_DATAW();

                VerifySucceeded(shellLinkW.GetPath(targetPath, targetPath.Capacity, ref data,
                                                   SLGP_UNCPRIORITY));

                return targetPath.ToString();
            }
            set
            {
                VerifySucceeded(shellLinkW.SetPath(value));
            }
        }

        public string Arguments
        {
            get
            {
                // No limitation to length of buffer string in the case of Unicode though.
                StringBuilder arguments = new StringBuilder(INFOTIPSIZE);

                VerifySucceeded(shellLinkW.GetArguments(arguments, arguments.Capacity));

                return arguments.ToString();
            }
            set
            {
                VerifySucceeded(shellLinkW.SetArguments(value));
            }
        }

        // AppUserModelID to be used for Windows 7 or later.
        public string AppUserModelID
        {
            get
            {
                using (PropVariant pv = new PropVariant())
                {
                    VerifySucceeded(PropertyStore.GetValue(AppUserModelIDKey, pv));

                    if (pv.Value == null)
                        return "Null";
                    else
                        return pv.Value.ToString();
                }
            }
            set
            {
                using (PropVariant pv = new PropVariant(value))
                {
                    VerifySucceeded(PropertyStore.SetValue(AppUserModelIDKey, pv));
                    VerifySucceeded(PropertyStore.Commit());
                }
            }
        }

        public Guid ToastActivatorId
        {
            get
            {
                using (PropVariant pv = new PropVariant())
                {
                    VerifySucceeded(PropertyStore.GetValue(CLSIDKey, pv));

                    if (pv.Value == null)
                        return Guid.Empty;
                    else
                        return new Guid(pv.Value.ToString());
                }
            }
            set
            {
                using (PropVariant pv = new PropVariant(value))
                {
                    VerifySucceeded(PropertyStore.SetValue(CLSIDKey, pv));
                    VerifySucceeded(PropertyStore.Commit());
                }
            }
        }

        #endregion

        #region Constructor

        public ShellLink()
            : this(null)
        { }

        // Construct with loading shortcut file.
        public ShellLink(string file)
        {
            try
            {
                shellLinkW = (IShellLinkW)new CShellLink();
            }
            catch
            {
                throw new COMException("Failed to create ShellLink object.");
            }

            if (file != null)
                Load(file);
        }

        #endregion

        #region Destructor

        ~ShellLink()
        {
            Dispose(false);
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (shellLinkW != null)
            {
                // Release all references.
                Marshal.FinalReleaseComObject(shellLinkW);
                shellLinkW = null;
            }
        }

        #endregion

        #region Methods

        // Save shortcut file.
        public void Save()
        {
            string file = ShortcutFile;

            if (file == null)
                throw new InvalidOperationException("File name is not given.");
            else
                Save(file);
        }

        public void Save(string file)
        {
            if (file == null)
                throw new ArgumentNullException("File name is required.");
            else
                PersistFile.Save(file, true);
        }

        // Load shortcut file.
        public void Load(string file)
        {
            if (!File.Exists(file))
                throw new FileNotFoundException("File is not found.", file);
            else
                PersistFile.Load(file, STGM_READ);
        }

        // Verify if operation succeeded.
        public static void VerifySucceeded(uint hresult)
        {
            if (hresult > 1)
                throw new InvalidOperationException("Failed with HRESULT: " +
                                                    hresult.ToString("X"));
        }

        #endregion
    }
}
