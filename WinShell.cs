using System;
using System.Collections.Generic;
using System.Text;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;

// Shell Property System:https://msdn.microsoft.com/en-us/library/windows/desktop/ff728898(v=vs.85).aspx
// Microsoft hasn't provided a good shell wrapper nor does the type library work: http://stackoverflow.com/questions/4450121/c-sharp-4-0-dynamic-object-and-winapi-interfaces-like-ishellitem-without-defini
// Help in managing VARIANT from managed code: https://limbioliong.wordpress.com/2011/09/04/using-variants-in-managed-code-part-1/

namespace WinShell
{

    // Wrapper Class for IPropertyStore
    class PropertyStore : IDisposable
    {
        public static PropertyStore Open(string filename, bool writeAccess = false)
        {
            NativeMethods.IPropertyStore store;
            Guid iPropertyStoreGuid = typeof(NativeMethods.IPropertyStore).GUID;
            NativeMethods.SHGetPropertyStoreFromParsingName(filename, (IntPtr)0,
                writeAccess ? GETPROPERTYSTOREFLAGS.GPS_READWRITE : GETPROPERTYSTOREFLAGS.GPS_BESTEFFORT,
                ref iPropertyStoreGuid, out store);
            return new PropertyStore(store);
        }

        NativeMethods.IPropertyStore m_IPropertyStore;

        public PropertyStore(NativeMethods.IPropertyStore propertyStore)
        {
            m_IPropertyStore = propertyStore;
        }

        public int Count
        {
            get
            {
                uint value;
                m_IPropertyStore.GetCount(out value);
                return (int)value;
            }
        }

        public PROPERTYKEY GetAt(int index)
        {
            PROPERTYKEY key;
            m_IPropertyStore.GetAt((uint)index, out key);
            return key;
        }

        public object GetValue(PROPERTYKEY key)
        {
            IntPtr pv = IntPtr.Zero;
            object value = null;
            try
            {
                pv = Marshal.AllocCoTaskMem(16);
                m_IPropertyStore.GetValue(key, pv);
                try
                {
                    value = PropVariant.ToObject(pv);
                }
                catch (Exception err)
                {
                    throw new ApplicationException("Unsupported property data type", err);
                }
            }
            finally
            {
                if (pv != (IntPtr)0)
                {
                    try
                    {
                        NativeMethods.PropVariantClear(pv);
                    }
                    catch
                    {
                        Debug.Fail("VariantClear failure");
                    }
                    Marshal.FreeCoTaskMem(pv);
                    pv = IntPtr.Zero;
                }
            }
            return value;
        }

        public void SetValue(PROPERTYKEY key, object value)
        {
            IntPtr pv = IntPtr.Zero;
            try
            {
                pv = NativeMethods.PropVariantFromObject(value);
                m_IPropertyStore.SetValue(key, pv);
            }
            finally
            {
                if (pv != IntPtr.Zero)
                {
                    NativeMethods.PropVariantClear(pv);
                    Marshal.FreeCoTaskMem(pv);
                    pv = IntPtr.Zero;
                }
            }
        }

        public void Commit()
        {
            m_IPropertyStore.Commit();
        }

        public void Dispose()
        {
            Dispose(true);
        }

        ~PropertyStore()
        {
            Dispose(false);
        }

        private void Dispose(bool disposing)
        {
            if (m_IPropertyStore != null)
            {
                if (!disposing)
                {
                    Debug.Fail("Failed to dispose PropertyStore");
                }

                Marshal.FinalReleaseComObject(m_IPropertyStore);
                m_IPropertyStore = null;
            }
            if (disposing)
            {
                GC.SuppressFinalize(this);
            }
        }

    } // PropertyStore

    public class PropertySystem : IDisposable
    {
        NativeMethods.IPropertySystem m_IPropertySystem;

        public PropertySystem()
        {
            Guid IID_IPropertySystem = typeof(NativeMethods.IPropertySystem).GUID;
            NativeMethods.PSGetPropertySystem(ref IID_IPropertySystem, out m_IPropertySystem);
        }

        public PropertyDescription GetPropertyDescription(PROPERTYKEY propKey)
        {
            Guid IID_IPropertyDescription = typeof(NativeMethods.IPropertyDescription).GUID;
            NativeMethods.IPropertyDescription iPropertyDescription;
            m_IPropertySystem.GetPropertyDescription(propKey, ref IID_IPropertyDescription, out iPropertyDescription);
            return new PropertyDescription(iPropertyDescription);
        }

        public PropertyDescription GetPropertyDescriptionByName(string canonicalName)
        {
            Guid IID_IPropertyDescription = typeof(NativeMethods.IPropertyDescription).GUID;
            NativeMethods.IPropertyDescription iPropertyDescription;
            m_IPropertySystem.GetPropertyDescriptionByName(canonicalName, ref IID_IPropertyDescription, out iPropertyDescription);
            return new PropertyDescription(iPropertyDescription);
        }

        public PROPERTYKEY GetPropertyKeyByName(string canonicalName)
        {
            using (PropertyDescription pd = GetPropertyDescriptionByName(canonicalName))
            {
                return pd.PropertyKey;
            }
        }
 
        /*
        public PropertyDescriptionList GetPropertyDescriptionListFromString(string propList)
        {
            throw new NotImplementedException();
        }

        public PropertyDescriptionList void EnumeratePropertyDescriptions(PROPDESC_ENUMFILTER)
        {
            throw new NotImplementedException();
        }

        public string FormatForDisplay(PROPERTYKEY propKey, PropVariant propvar, PROPDESC_FORMAT_FLAGS pdff)
        {
            throw new NotImplementedException();
        }

        public void RegisterPropertySchema(string path)
        {
            throw new NotImplementedException();
        }

        public void UnregisterPropertySchema(string path)
        {
            throw new NotImplementedException();
        }

        public void RefreshPropertySchema()
        {
            throw new NotImplementedException();
        }
        */

        public void Dispose()
        {
            Dispose(true);
        }

        ~PropertySystem()
        {
            Dispose(false);
        }

        private void Dispose(bool disposing)
        {
            if (m_IPropertySystem != null)
            {
                if (!disposing)
                {
                    Debug.Fail("Failed to dispose PropertySystem");
                }

                Marshal.FinalReleaseComObject(m_IPropertySystem);
                m_IPropertySystem = null;
            }
            if (disposing)
            {
                GC.SuppressFinalize(this);
            }
        }
    }

    public class PropertyDescription : IDisposable
    {

        NativeMethods.IPropertyDescription m_IPropertyDescription;

        internal PropertyDescription(NativeMethods.IPropertyDescription iPropertyDescription)
        {
            m_IPropertyDescription = iPropertyDescription;
        }

        public PROPERTYKEY PropertyKey
        {
            get
            {
                PROPERTYKEY value;
                m_IPropertyDescription.GetPropertyKey(out value);
                return value;
            }
        }

        public string CanonicalName
        {
            get
            {
                IntPtr pszName = (IntPtr)0;
                try
                {
                    m_IPropertyDescription.GetCanonicalName(out pszName);
                    return Marshal.PtrToStringUni(pszName);
                }
                finally
                {
                    if (pszName != (IntPtr)0)
                    {
                        Marshal.FreeCoTaskMem(pszName);
                        pszName = (IntPtr)0;
                    }
                }
            }
        }

        public string DisplayName
        {
            get
            {
                IntPtr pszName = (IntPtr)0;
                try
                {
                    m_IPropertyDescription.GetDisplayName(out pszName);
                    return Marshal.PtrToStringUni(pszName);
                }
                finally
                {
                    if (pszName != (IntPtr)0)
                    {
                        Marshal.FreeCoTaskMem(pszName);
                        pszName = (IntPtr)0;
                    }
                }
            }
        }

        public void Dispose()
        {
            Dispose(true);
        }

        ~PropertyDescription()
        {
            Dispose(false);
        }

        private void Dispose(bool disposing)
        {
            if (m_IPropertyDescription != null)
            {
                if (!disposing)
                {
                    Debug.Fail("Failed to dispose PropertyDescription");
                }

                Marshal.FinalReleaseComObject(m_IPropertyDescription);
                m_IPropertyDescription = null;
            }
            if (disposing)
            {
                GC.SuppressFinalize(this);
            }
        }
    }

    [StructLayout (LayoutKind.Sequential, Pack = 4)]
    public struct PROPERTYKEY
    {
        public Guid fmtid;
        public UInt32 pid;
    }

    internal static class PropVariant
    {
        public static object ToObject(IntPtr pv)
        {
            // Copy to structure
            PROPVARIANT v = (PROPVARIANT)Marshal.PtrToStructure(pv, typeof(PROPVARIANT));

            object value = null;
            switch (v.vt)
            {
                case  0: // VT_EMPTY
                case  1: // VT_NULL
                case  2: // VT_I2
                case  3: // VT_I4
                case  4: // VT_R4
                case  5: // VT_R8
                case  6: // VT_CY
                case  7: // VT_DATE
                case  8: // VT_BSTR
                case 10: // VT_ERROR
                case 11: // VT_BOOL
                case 14: // VT_DECIMAL
                case 16: // VT_I1
                case 17: // VT_UI1
                case 18: // VT_UI2
                case 19: // VT_UI4
                case 20: // VT_I8
                case 21: // VT_UI8
                case 22: // VT_INT
                case 23: // VT_UINT
                case 24: // VT_VOID
                case 25: // VT_HRESULT
                    value = Marshal.GetObjectForNativeVariant(pv);
                    break;

                case 30: // VT_LPSTR
                    value = Marshal.PtrToStringAnsi(v.dataIntPtr);
                    break;

                case 31: // VT_LPWSTR
                    value = Marshal.PtrToStringUni(v.dataIntPtr);
                    break;

                case 0x101f: // VT_VECTOR|VT_LPWSTR
                    {
                        string[] strings = new string[v.cElems];
                        for (int i=0; i<v.cElems; ++i)
                        {
                            IntPtr strPtr = Marshal.ReadIntPtr(v.pElems + i*Marshal.SizeOf(typeof(IntPtr)));
                            strings[i] = Marshal.PtrToStringUni(strPtr);
                        }
                        value = strings;
                    }
                    break;

                case 0x1005: // VT_Vector|VT_R8
                    {
                        double[] doubles = new double[v.cElems];
                        Marshal.Copy(v.pElems, doubles, 0, (int)v.cElems);
                        value = doubles;
                    }
                    break;

                case 64: // VT_FILETIME
                    value = DateTime.FromFileTime(v.dataInt64);
                    break;

                default:
                    try
                    {
                        value = Marshal.GetObjectForNativeVariant(pv);
                        if (value == null) value = "(null)";
                        value = String.Format("(Supported type 0x{0:x4}): {1}", v.vt, value.ToString());
                    }
                    catch
                    {
                        // Get the variant type
                        value = String.Format("(Unsupported type 0x{0:x4})", v.vt);
                    }
                    break;
            }

            return value;
        }

        /*
        // C++ version
        typedef struct PROPVARIANT {
            VARTYPE vt;
            WORD    wReserved1;
            WORD    wReserved2;
            WORD    wReserved3;
            union {
                // Various types of up to 8 bytes
            }
        } PROPVARIANT;
        */
        [StructLayout(LayoutKind.Explicit)]
        public struct PROPVARIANT
        {
            [FieldOffset(0)]
            public ushort vt;
            [FieldOffset(2)]
            public ushort wReserved1;
            [FieldOffset(4)]
            public ushort wReserved2;
            [FieldOffset(6)]
            public ushort wReserved3;
            [FieldOffset(8)]
            public Int32 data01;
            [FieldOffset(12)]
            public Int32 data02;

            // IntPtr (for strings and the like)
            [FieldOffset(8)]
            public IntPtr dataIntPtr;

            // For FileTime and Int64
            [FieldOffset(8)]
            public long dataInt64;

            // Vector-style arrays (for VT_VECTOR|VT_LPWSTR and such)
            [FieldOffset(8)]
            public uint cElems;
            [FieldOffset(12)]
            public IntPtr pElems;
        }
    }

    public enum GETPROPERTYSTOREFLAGS : uint
    {
        // If no flags are specified (GPS_DEFAULT), a read-only property store is returned that includes properties for the file or item.
        // In the case that the shell item is a file, the property store contains:
        //     1. properties about the file from the file system
        //     2. properties from the file itself provided by the file's property handler, unless that file is offline,
        //     see GPS_OPENSLOWITEM
        //     3. if requested by the file's property handler and supported by the file system, properties stored in the
        //     alternate property store.
        //
        // Non-file shell items should return a similar read-only store
        //
        // Specifying other GPS_ flags modifies the store that is returned
        GPS_DEFAULT = 0x00000000,
        GPS_HANDLERPROPERTIESONLY = 0x00000001,   // only include properties directly from the file's property handler
        GPS_READWRITE = 0x00000002,   // Writable stores will only include handler properties
        GPS_TEMPORARY = 0x00000004,   // A read/write store that only holds properties for the lifetime of the IShellItem object
        GPS_FASTPROPERTIESONLY = 0x00000008,   // do not include any properties from the file's property handler (because the file's property handler will hit the disk)
        GPS_OPENSLOWITEM = 0x00000010,   // include properties from a file's property handler, even if it means retrieving the file from offline storage.
        GPS_DELAYCREATION = 0x00000020,   // delay the creation of the file's property handler until those properties are read, written, or enumerated
        GPS_BESTEFFORT = 0x00000040,   // For readonly stores, succeed and return all available properties, even if one or more sources of properties fails. Not valid with GPS_READWRITE.
        GPS_NO_OPLOCK = 0x00000080,   // some data sources protect the read property store with an oplock, this disables that
        GPS_MASK_VALID = 0x000000FF,
    }

    public enum PROPDESC_ENUMFILTER : uint
    {
        PDEF_ALL	= 0,
        PDEF_SYSTEM	= 1,
        PDEF_NONSYSTEM	= 2,
        PDEF_VIEWABLE	= 3,
        PDEF_QUERYABLE	= 4,
        PDEF_INFULLTEXTQUERY	= 5,
        PDEF_COLUMN	= 6
    }

    [Flags]
    public enum PROPDESC_FORMAT_FLAGS : uint
    {
        PDFF_DEFAULT = 0,
        PDFF_PREFIXNAME = 0x1,
        PDFF_FILENAME = 0x2,
        PDFF_ALWAYSKB = 0x4,
        PDFF_RESERVED_RIGHTTOLEFT = 0x8,
        PDFF_SHORTTIME = 0x10,
        PDFF_LONGTIME = 0x20,
        PDFF_HIDETIME = 0x40,
        PDFF_SHORTDATE = 0x80,
        PDFF_LONGDATE = 0x100,
        PDFF_HIDEDATE = 0x200,
        PDFF_RELATIVEDATE = 0x400,
        PDFF_USEEDITINVITATION = 0x800,
        PDFF_READONLY = 0x1000,
        PDFF_NOAUTOREADINGORDER = 0x2000
    }

    internal static class NativeMethods
    {

        /*
        // The C++ Version
        interface IPropertyStore : IUnknown
        {
            HRESULT GetCount([out] DWORD *cProps);
            HRESULT GetAt([in] DWORD iProp, [out] PROPERTYKEY *pkey);
            HRESULT GetValue([in] REFPROPERTYKEY key, [out] PROPVARIANT *pv);
            HRESULT SetValue([in] REFPROPERTYKEY key, [in] REFPROPVARIANT propvar);
            HRESULT Commit();
        }
        */
        [ComImport, Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IPropertyStore
        {
            void GetCount([Out] out uint cProps);

            void GetAt([In] uint iProp, out PROPERTYKEY pkey);

            void GetValue([In] ref PROPERTYKEY key, [In] IntPtr pv);

            void SetValue([In] ref PROPERTYKEY key, [In] IntPtr pv);

            void Commit();
        }

        /*
        MIDL_INTERFACE("ca724e8a-c3e6-442b-88a4-6fb0db8035a3")
        IPropertySystem : public IUnknown
        {
        public:
            virtual HRESULT STDMETHODCALLTYPE GetPropertyDescription( 
                __RPC__in REFPROPERTYKEY propkey,
                __RPC__in REFIID riid,
                __RPC__deref_out_opt void **ppv) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetPropertyDescriptionByName( 
                __RPC__in_string LPCWSTR pszCanonicalName,
                __RPC__in REFIID riid,
                __RPC__deref_out_opt void **ppv) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetPropertyDescriptionListFromString( 
                __RPC__in_string LPCWSTR pszPropList,
                __RPC__in REFIID riid,
                __RPC__deref_out_opt void **ppv) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE EnumeratePropertyDescriptions( 
                PROPDESC_ENUMFILTER filterOn,
                __RPC__in REFIID riid,
                __RPC__deref_out_opt void **ppv) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE FormatForDisplay( 
                __RPC__in REFPROPERTYKEY key,
                __RPC__in REFPROPVARIANT propvar,
                PROPDESC_FORMAT_FLAGS pdff,
                __RPC__out_ecount_full_string(cchText) LPWSTR pszText,
                __RPC__in_range(0,0x8000) DWORD cchText) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE FormatForDisplayAlloc( 
                __RPC__in REFPROPERTYKEY key,
                __RPC__in REFPROPVARIANT propvar,
                PROPDESC_FORMAT_FLAGS pdff,
                __RPC__deref_out_opt_string LPWSTR *ppszDisplay) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE RegisterPropertySchema( 
                __RPC__in_string LPCWSTR pszPath) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE UnregisterPropertySchema( 
                __RPC__in_string LPCWSTR pszPath) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE RefreshPropertySchema( void) = 0;
        
        };
        */ 
        [ComImport, Guid("ca724e8a-c3e6-442b-88a4-6fb0db8035a3"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IPropertySystem
        {
            void GetPropertyDescription([In] ref PROPERTYKEY propkey, [In] ref Guid riid, [Out] out IPropertyDescription rPropertyDescription);

            void GetPropertyDescriptionByName([In][MarshalAs(UnmanagedType.LPWStr)] string pszCanonicalName, [In] ref Guid riid, [Out] out IPropertyDescription rPropertyDescription);

            void GetPropertyDescriptionListFromString([In][MarshalAs(UnmanagedType.LPWStr)] string pszPropList, [In] ref Guid riid, [Out] out IPropertyDescriptionList rPropertyDescriptionList);
        
            void EnumeratePropertyDescriptions([In] PROPDESC_ENUMFILTER filterOn, [In] ref Guid riid, [Out] out IPropertyDescriptionList rPropertyDescriptionList);

            void FormatForDisplay([In] ref PROPERTYKEY key, [In] IntPtr propvar, [In] PROPDESC_FORMAT_FLAGS pdff, [In] IntPtr pszText, ushort cchText );

            void FormatForDisplayAlloc([In] ref PROPERTYKEY key, [In] IntPtr propvar, [In] PROPDESC_FORMAT_FLAGS pdff, [Out] out IntPtr ppszText);

            void RegisterPropertySchema([In][MarshalAs(UnmanagedType.LPWStr)] string pszPath);

            void UnregisterPropertySchema([In][MarshalAs(UnmanagedType.LPWStr)] string pszPath);

            void RefreshPropertySchema();
        }

        /*
        MIDL_INTERFACE("6f79d558-3e96-4549-a1d1-7d75d2288814")
        IPropertyDescription : public IUnknown
        {
        public:
            virtual HRESULT STDMETHODCALLTYPE GetPropertyKey( 
                __RPC__out PROPERTYKEY *pkey) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetCanonicalName( 
                __RPC__deref_out_opt_string LPWSTR *ppszName) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetPropertyType( 
                __RPC__out VARTYPE *pvartype) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetDisplayName( 
                __RPC__deref_out_opt_string LPWSTR *ppszName) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetEditInvitation( 
                __RPC__deref_out_opt_string LPWSTR *ppszInvite) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetTypeFlags( 
                PROPDESC_TYPE_FLAGS mask,
                __RPC__out PROPDESC_TYPE_FLAGS *ppdtFlags) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetViewFlags( 
                __RPC__out PROPDESC_VIEW_FLAGS *ppdvFlags) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetDefaultColumnWidth( 
                __RPC__out UINT *pcxChars) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetDisplayType( 
                __RPC__out PROPDESC_DISPLAYTYPE *pdisplaytype) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetColumnState( 
                __RPC__out SHCOLSTATEF *pcsFlags) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetGroupingRange( 
                __RPC__out PROPDESC_GROUPING_RANGE *pgr) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetRelativeDescriptionType( 
                __RPC__out PROPDESC_RELATIVEDESCRIPTION_TYPE *prdt) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetRelativeDescription( 
                __RPC__in REFPROPVARIANT propvar1,
                __RPC__in REFPROPVARIANT propvar2,
                __RPC__deref_out_opt_string LPWSTR *ppszDesc1,
                __RPC__deref_out_opt_string LPWSTR *ppszDesc2) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetSortDescription( 
                __RPC__out PROPDESC_SORTDESCRIPTION *psd) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetSortDescriptionLabel( 
                BOOL fDescending,
                __RPC__deref_out_opt_string LPWSTR *ppszDescription) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetAggregationType( 
                __RPC__out PROPDESC_AGGREGATION_TYPE *paggtype) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetConditionType( 
                __RPC__out PROPDESC_CONDITION_TYPE *pcontype,
                __RPC__out CONDITION_OPERATION *popDefault) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetEnumTypeList( 
                __RPC__in REFIID riid,
                __RPC__deref_out_opt void **ppv) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE CoerceToCanonicalValue( 
                _Inout_  PROPVARIANT *ppropvar) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE FormatForDisplay( 
                __RPC__in REFPROPVARIANT propvar,
                PROPDESC_FORMAT_FLAGS pdfFlags,
                __RPC__deref_out_opt_string LPWSTR *ppszDisplay) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE IsValueCanonical( 
                __RPC__in REFPROPVARIANT propvar) = 0;
        
        };
        */
        [ComImport, Guid("6f79d558-3e96-4549-a1d1-7d75d2288814"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IPropertyDescription
        {
            void GetPropertyKey([Out] out PROPERTYKEY pkey);

            void GetCanonicalName([Out] out IntPtr ppszName);

            void GetPropertyType([Out] out ushort vartype);

            void GetDisplayName([Out] out IntPtr ppszName);
        
            // === All Other Methods Deferred Until Later! ===
        }

        public interface IPropertyDescriptionList
        {

        }

        [DllImport("shell32.dll", SetLastError = true, CallingConvention = CallingConvention.StdCall, PreserveSig=false)]
        public static extern void SHGetPropertyStoreFromParsingName(
                [In][MarshalAs(UnmanagedType.LPWStr)] string pszPath,
                [In] IntPtr zeroWorks,
                [In] GETPROPERTYSTOREFLAGS flags,
                [In] ref Guid iIdPropStore,
                [Out] out IPropertyStore propertyStore);

        [DllImport(@"ole32.dll", SetLastError = true, CallingConvention = CallingConvention.StdCall, PreserveSig=false)]
        public static extern void PropVariantInit([In] IntPtr pvarg);

        [DllImport(@"ole32.dll", SetLastError = true, CallingConvention = CallingConvention.StdCall, PreserveSig=false)]
        public static extern void PropVariantClear([In] IntPtr pvarg);

        [DllImport("propsys.dll", SetLastError=true, CallingConvention = CallingConvention.StdCall, PreserveSig=false)]
        public static extern void PSGetPropertySystem([In] ref Guid iIdPropertySystem, [Out] out IPropertySystem propertySystem);

        // Converts a string to a PropVariant with type LPWSTR instead of BSTR
        // The resulting variant must be cleared using PropVariantClear and freed using Marshal.FreeCoTaskMem
        public static IntPtr PropVariantFromString(string value)
        {
            IntPtr pstr = IntPtr.Zero;
            IntPtr pv = IntPtr.Zero;
            try
            {
                // In managed code, new automatically zeros the contents.
                PropVariant.PROPVARIANT propvariant = new PropVariant.PROPVARIANT();

                // Allocate the string
                pstr = Marshal.StringToCoTaskMemUni(value);

                // Allocate the PropVariant
                pv = Marshal.AllocCoTaskMem(16);

                // Transfer ownership of the string
                propvariant.vt = 31; // VT_LPWSTR - not documented but this is to be allocated using CoTaskMemAlloc.
                propvariant.dataIntPtr = pstr;
                Marshal.StructureToPtr(propvariant, pv, false);
                pstr = IntPtr.Zero;

                // Transfer ownership to the result
                IntPtr result = pv;
                pv = IntPtr.Zero;
                return result;
            }
            finally
            {
                if (pstr != IntPtr.Zero)
                {
                    Marshal.FreeCoTaskMem(pstr);
                    pstr = IntPtr.Zero;
                }
                if (pv != IntPtr.Zero)
                {
                    try
                    {
                        NativeMethods.PropVariantClear(pv);
                    }
                    catch
                    {
                        Debug.Fail("VariantClear failure");
                    }
                    Marshal.FreeCoTaskMem(pv);
                    pv = IntPtr.Zero;
                }
            }
        }

        // Converts an object to a PropVariant including special handling for strings
        // The resulting variant must be cleared using PropVariantClear and freed using Marshal.FreeCoTaskMem
        public static IntPtr PropVariantFromObject(object value)
        {
            string strValue = value as string;
            if (strValue != null)
            {
                return PropVariantFromString(strValue);
            }
            else
            {
                IntPtr pv = IntPtr.Zero;
                try
                {
                    pv = Marshal.AllocCoTaskMem(16);
                    Marshal.GetNativeVariantForObject(value, pv);
                    IntPtr result = pv;
                    pv = IntPtr.Zero;
                    return result;
                }
                finally
                {
                    if (pv != (IntPtr)0)
                    {
                        try
                        {
                            NativeMethods.PropVariantClear(pv);
                        }
                        catch
                        {
                            Debug.Fail("VariantClear failure");
                        }
                        Marshal.FreeCoTaskMem(pv);
                        pv = IntPtr.Zero;
                    }
                }
            }
        }

    }
}
