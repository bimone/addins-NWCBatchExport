using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
[assembly: AssemblyTitle("BatchExporter")]
[assembly: AssemblyProduct("BatchExporter")]
[assembly: AssemblyConfiguration("BIMO")]
// Version: Revit year and release number.
#if REVIT2023
[assembly: AssemblyVersion("2023.1.0.0")]
[assembly: AssemblyFileVersion("2023.1.0.0")]
#elif REVIT2024
[assembly: AssemblyVersion("2024.1.0.0")]
[assembly: AssemblyFileVersion("2024.1.0.0")]
#elif REVIT2025
[assembly: AssemblyVersion("2025.1.0.0")]
[assembly: AssemblyFileVersion("2025.1.0.0")]
#else
[assembly: AssemblyVersion("2026.1.0.0")]
[assembly: AssemblyFileVersion("2026.1.0.0")]
#endif

// Setting ComVisible to false makes the types in this assembly not visible 
// to COM components.  If you need to access a type in this assembly from 
// COM, set the ComVisible attribute to true on that type.
[assembly: ComVisible(false)]

// The following GUID is for the ID of the typelib if this project is exposed to COM
[assembly: Guid("01894EFB-E0C3-489E-9A44-7D64A4579FB0")]
