using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace QuickConvertShell
{
    [Flags]
    public enum EXPCMDFLAGS : uint
    {
        ECF_DEFAULT = 0x000,
        ECF_HASSUBCOMMANDS = 0x001,
        ECF_HASSPLITBUTTON = 0x002,
        ECF_HIDELABEL = 0x004,
        ECF_ISSEPARATOR = 0x008,
        ECF_HASLUASHIELD = 0x010,
        ECF_SEPARATORBEFORE = 0x020,
        ECF_SEPARATORAFTER = 0x040,
        ECF_ISDROPDOWN = 0x080,
        ECF_TOGGLEABLE = 0x100,
        ECF_AUTOMENUICONS = 0x200
    }

    [Flags]
    public enum EXPCMDSTATE : uint
    {
        ECS_ENABLED = 0x000,
        ECS_DISABLED = 0x001,
        ECS_HIDDEN = 0x002,
        ECS_CHECKBOX = 0x004,
        ECS_CHECKED = 0x008,
        ECS_RADIOCHECK = 0x010
    }

    public enum SIGDN : uint
    {
        SIGDN_FILESYSPATH = 0x80058000
    }

    [ComImport]
    [Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IShellItem
    {
        void BindToHandler(IntPtr pbc, [MarshalAs(UnmanagedType.LPStruct)] Guid bhid, [MarshalAs(UnmanagedType.LPStruct)] Guid riid, out IntPtr ppv);
        void GetParent(out IShellItem ppsi);
        [PreserveSig]
        int GetDisplayName(SIGDN sigdnName, [MarshalAs(UnmanagedType.LPWStr)] out string ppszName);
        void GetAttributes(uint sfgaoMask, out uint psfgaoAttribs);
        void Compare(IShellItem psi, uint hint, out int piOrder);
    }

    [ComImport]
    [Guid("b63ea76d-1f85-456f-a19c-48159efa858b")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IShellItemArray
    {
        void BindToHandler(IntPtr pbc, [MarshalAs(UnmanagedType.LPStruct)] Guid bhid, [MarshalAs(UnmanagedType.LPStruct)] Guid riid, out IntPtr ppv);
        void GetPropertyStore(int flags, [MarshalAs(UnmanagedType.LPStruct)] Guid riid, out IntPtr ppv);
        void GetPropertyDescriptionList(IntPtr keyType, [MarshalAs(UnmanagedType.LPStruct)] Guid riid, out IntPtr ppv);
        void GetAttributes(int AttribFlags, uint sfgaoMask, out uint psfgaoAttribs);
        [PreserveSig]
        int GetCount(out uint pdwNumItems);
        [PreserveSig]
        int GetItemAt(uint dwIndex, out IShellItem ppsi);
        void EnumItems(out IntPtr ppenumShellItems);
    }

    [ComImport]
    [Guid("a88826f8-186f-4987-aade-ea0cef8fbfe8")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IEnumExplorerCommand
    {
        [PreserveSig]
        int Next(uint celt, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.Interface, SizeParamIndex = 0)] out IExplorerCommand[] pUICommand, out uint pceltFetched);
        [PreserveSig]
        int Skip(uint celt);
        [PreserveSig]
        int Reset();
        [PreserveSig]
        int Clone(out IEnumExplorerCommand ppenum);
    }

    [ComImport]
    [Guid("a08ce4d0-fa25-44ab-b57c-c7b1c323e0b9")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IExplorerCommand
    {
        [PreserveSig]
        int GetTitle(IShellItemArray psiItemArray, [MarshalAs(UnmanagedType.LPWStr)] out string ppszName);
        [PreserveSig]
        int GetIcon(IShellItemArray psiItemArray, [MarshalAs(UnmanagedType.LPWStr)] out string ppszIcon);
        [PreserveSig]
        int GetToolTip(IShellItemArray psiItemArray, [MarshalAs(UnmanagedType.LPWStr)] out string ppszInfotip);
        [PreserveSig]
        int GetCanonicalName(out Guid pguidCommandName);
        [PreserveSig]
        int GetState(IShellItemArray psiItemArray, [MarshalAs(UnmanagedType.Bool)] bool fOkToBeSlow, out EXPCMDSTATE pCmdState);
        [PreserveSig]
        int Invoke(IShellItemArray psiItemArray, IntPtr pbc);
        [PreserveSig]
        int GetFlags(out EXPCMDFLAGS pFlags);
        [PreserveSig]
        int EnumSubCommands(out IEnumExplorerCommand ppEnum);
    }

    internal static class ShellHelpers
    {
        public static List<string> GetSelectedPaths(IShellItemArray psiItemArray)
        {
            var paths = new List<string>();
            if (psiItemArray == null)
            {
                return paths;
            }
            uint count = 0;
            if (psiItemArray.GetCount(out count) != 0)
            {
                return paths;
            }
            for (uint i = 0; i < count; i++)
            {
                IShellItem item;
                if (psiItemArray.GetItemAt(i, out item) == 0 && item != null)
                {
                    string path;
                    if (item.GetDisplayName(SIGDN.SIGDN_FILESYSPATH, out path) == 0 && !string.IsNullOrEmpty(path))
                    {
                        paths.Add(path);
                    }
                    Marshal.ReleaseComObject(item);
                }
            }
            return paths;
        }

        public static void RunCli(string cliArgs, List<string> paths)
        {
            string dllDir = Path.GetDirectoryName(typeof(ShellHelpers).Assembly.Location) ?? "";
            string rootDir = Path.GetFullPath(Path.Combine(dllDir, ".."));
            string pythonw = Path.Combine(rootDir, ".venv", "Scripts", "pythonw.exe");
            if (!File.Exists(pythonw))
            {
                pythonw = "pythonw.exe";
            }
            string cliPy = Path.Combine(rootDir, "cli.py");

            var sb = new StringBuilder();
            sb.Append("\"").Append(cliPy).Append("\" --notify ").Append(cliArgs);
            foreach (string p in paths)
            {
                sb.Append(" \"").Append(p).Append("\"");
            }

            var psi = new ProcessStartInfo
            {
                FileName = pythonw,
                Arguments = sb.ToString(),
                UseShellExecute = false,
                CreateNoWindow = true,
                WorkingDirectory = rootDir
            };
            Process.Start(psi);
        }
    }

    internal sealed class CommandEnum : IEnumExplorerCommand
    {
        private readonly List<IExplorerCommand> _commands;
        private int _index;

        public CommandEnum(List<IExplorerCommand> commands)
        {
            _commands = commands;
            _index = 0;
        }

        public int Next(uint celt, out IExplorerCommand[] pUICommand, out uint pceltFetched)
        {
            var result = new List<IExplorerCommand>();
            uint fetched = 0;
            while (fetched < celt && _index < _commands.Count)
            {
                result.Add(_commands[_index]);
                _index++;
                fetched++;
            }
            pUICommand = result.ToArray();
            pceltFetched = fetched;
            return fetched == celt ? 0 : 1; // S_OK : S_FALSE
        }

        public int Skip(uint celt)
        {
            _index = Math.Min(_commands.Count, _index + (int)celt);
            return _index < _commands.Count ? 0 : 1;
        }

        public int Reset()
        {
            _index = 0;
            return 0;
        }

        public int Clone(out IEnumExplorerCommand ppenum)
        {
            ppenum = new CommandEnum(_commands) { _index = _index };
            return 0;
        }
    }

    internal sealed class SubCommand : IExplorerCommand
    {
        private readonly string _title;
        private readonly string _cliArgs;
        private readonly string _icon;
        private readonly Func<List<string>, bool> _predicate;
        private readonly Guid _guid;

        public SubCommand(string title, string cliArgs, string icon, Func<List<string>, bool> predicate, string guidStr)
        {
            _title = title;
            _cliArgs = cliArgs;
            _icon = icon;
            _predicate = predicate;
            _guid = new Guid(guidStr);
        }

        public int GetTitle(IShellItemArray psiItemArray, out string ppszName)
        {
            ppszName = _title;
            return 0;
        }

        public int GetIcon(IShellItemArray psiItemArray, out string ppszIcon)
        {
            ppszIcon = _icon;
            return 0;
        }

        public int GetToolTip(IShellItemArray psiItemArray, out string ppszInfotip)
        {
            ppszInfotip = _title;
            return 0;
        }

        public int GetCanonicalName(out Guid pguidCommandName)
        {
            pguidCommandName = _guid;
            return 0;
        }

        public int GetState(IShellItemArray psiItemArray, bool fOkToBeSlow, out EXPCMDSTATE pCmdState)
        {
            var paths = ShellHelpers.GetSelectedPaths(psiItemArray);
            pCmdState = _predicate(paths) ? EXPCMDSTATE.ECS_ENABLED : EXPCMDSTATE.ECS_HIDDEN;
            return 0;
        }

        public int Invoke(IShellItemArray psiItemArray, IntPtr pbc)
        {
            var paths = ShellHelpers.GetSelectedPaths(psiItemArray);
            if (paths.Count > 0)
            {
                ShellHelpers.RunCli(_cliArgs, paths);
            }
            return 0;
        }

        public int GetFlags(out EXPCMDFLAGS pFlags)
        {
            pFlags = EXPCMDFLAGS.ECF_DEFAULT;
            return 0;
        }

        public int EnumSubCommands(out IEnumExplorerCommand ppEnum)
        {
            ppEnum = null;
            return unchecked((int)0x80004001); // E_NOTIMPL
        }
    }

    [ComVisible(true)]
    [Guid("8F3C7E21-4B9A-4D8E-9C12-6A7B8C9D0E1F")]
    [ClassInterface(ClassInterfaceType.None)]
    public sealed class QuickConvertRootCommand : IExplorerCommand
    {
        private static bool HasExt(List<string> paths, params string[] exts)
        {
            if (paths.Count == 0) return false;
            foreach (string p in paths)
            {
                string ext = Path.GetExtension(p).ToLowerInvariant();
                bool match = false;
                foreach (string e in exts)
                {
                    if (ext == e) { match = true; break; }
                }
                if (!match) return false;
            }
            return true;
        }

        private static readonly List<IExplorerCommand> SubCommands = new List<IExplorerCommand>
        {
            new SubCommand("Convert to PDF (.pdf)", "convert --to pdf", "imageres.dll,-102", p => p.Count == 1 && HasExt(p, ".docx", ".md"), "10000001-4B9A-4D8E-9C12-6A7B8C9D0E1F"),
            new SubCommand("Convert to Word (.docx)", "convert --to docx", "imageres.dll,-102", p => p.Count == 1 && HasExt(p, ".pdf", ".md"), "10000002-4B9A-4D8E-9C12-6A7B8C9D0E1F"),
            new SubCommand("Convert to Markdown (.md)", "convert --to md", "imageres.dll,-102", p => p.Count == 1 && HasExt(p, ".docx", ".pdf"), "10000003-4B9A-4D8E-9C12-6A7B8C9D0E1F"),
            new SubCommand("Extract Text (.txt)", "convert --to txt", "imageres.dll,-102", p => p.Count == 1 && HasExt(p, ".pdf"), "10000004-4B9A-4D8E-9C12-6A7B8C9D0E1F"),
            new SubCommand("Merge Selected PDFs", "pdf-merge", "imageres.dll,-109", p => p.Count >= 2 && HasExt(p, ".pdf"), "10000005-4B9A-4D8E-9C12-6A7B8C9D0E1F"),
            new SubCommand("Compress PDF (Medium)", "pdf-compress --quality medium", "imageres.dll,-165", p => p.Count == 1 && HasExt(p, ".pdf"), "10000006-4B9A-4D8E-9C12-6A7B8C9D0E1F"),
            new SubCommand("Compress PDF (Max Compression)", "pdf-compress --quality high", "imageres.dll,-165", p => p.Count == 1 && HasExt(p, ".pdf"), "10000007-4B9A-4D8E-9C12-6A7B8C9D0E1F"),
            new SubCommand("Split PDF (Burst Pages)", "pdf-split --burst", "imageres.dll,-109", p => p.Count == 1 && HasExt(p, ".pdf"), "10000008-4B9A-4D8E-9C12-6A7B8C9D0E1F"),
            new SubCommand("Extract PDF Images", "pdf-extract-images", "imageres.dll,-72", p => p.Count == 1 && HasExt(p, ".pdf"), "10000009-4B9A-4D8E-9C12-6A7B8C9D0E1F"),
            new SubCommand("Images to PDF", "images-to-pdf", "imageres.dll,-72", p => HasExt(p, ".jpg", ".jpeg", ".png", ".webp", ".bmp"), "10000010-4B9A-4D8E-9C12-6A7B8C9D0E1F"),
            new SubCommand("Compress Image", "compress-image --quality medium", "imageres.dll,-165", p => p.Count == 1 && HasExt(p, ".jpg", ".jpeg", ".png", ".webp"), "10000011-4B9A-4D8E-9C12-6A7B8C9D0E1F"),
            new SubCommand("Convert to PNG (.png)", "convert --to png", "imageres.dll,-72", p => p.Count == 1 && HasExt(p, ".jpg", ".jpeg", ".webp", ".bmp"), "10000012-4B9A-4D8E-9C12-6A7B8C9D0E1F"),
            new SubCommand("Convert to JPG (.jpg)", "convert --to jpg", "imageres.dll,-72", p => p.Count == 1 && HasExt(p, ".png", ".webp", ".bmp"), "10000013-4B9A-4D8E-9C12-6A7B8C9D0E1F"),
            new SubCommand("Convert to WebP (.webp)", "convert --to webp", "imageres.dll,-72", p => p.Count == 1 && HasExt(p, ".png", ".jpg", ".jpeg", ".bmp"), "10000014-4B9A-4D8E-9C12-6A7B8C9D0E1F"),
            new SubCommand("Convert to CSV (.csv)", "convert --to csv", "imageres.dll,-102", p => p.Count == 1 && HasExt(p, ".xlsx", ".json"), "10000015-4B9A-4D8E-9C12-6A7B8C9D0E1F"),
            new SubCommand("Convert to Excel (.xlsx)", "convert --to xlsx", "imageres.dll,-102", p => p.Count == 1 && HasExt(p, ".csv", ".json"), "10000016-4B9A-4D8E-9C12-6A7B8C9D0E1F"),
            new SubCommand("Convert to JSON (.json)", "convert --to json", "imageres.dll,-102", p => p.Count == 1 && HasExt(p, ".xlsx", ".csv"), "10000017-4B9A-4D8E-9C12-6A7B8C9D0E1F")
        };

        public int GetTitle(IShellItemArray psiItemArray, out string ppszName)
        {
            ppszName = "Quick Convert";
            return 0;
        }

        public int GetIcon(IShellItemArray psiItemArray, out string ppszIcon)
        {
            ppszIcon = "imageres.dll,-109";
            return 0;
        }

        public int GetToolTip(IShellItemArray psiItemArray, out string ppszInfotip)
        {
            ppszInfotip = "Fast local file conversion and PDF utilities";
            return 0;
        }

        public int GetCanonicalName(out Guid pguidCommandName)
        {
            pguidCommandName = new Guid("8F3C7E21-4B9A-4D8E-9C12-6A7B8C9D0E1F");
            return 0;
        }

        public int GetState(IShellItemArray psiItemArray, bool fOkToBeSlow, out EXPCMDSTATE pCmdState)
        {
            var paths = ShellHelpers.GetSelectedPaths(psiItemArray);
            bool anySupported = HasExt(
                paths,
                ".docx", ".md", ".pdf", ".xlsx", ".csv", ".json",
                ".jpg", ".jpeg", ".png", ".webp", ".bmp"
            );
            pCmdState = anySupported ? EXPCMDSTATE.ECS_ENABLED : EXPCMDSTATE.ECS_HIDDEN;
            return 0;
        }

        public int Invoke(IShellItemArray psiItemArray, IntPtr pbc)
        {
            return 0;
        }

        public int GetFlags(out EXPCMDFLAGS pFlags)
        {
            pFlags = EXPCMDFLAGS.ECF_HASSUBCOMMANDS;
            return 0;
        }

        public int EnumSubCommands(out IEnumExplorerCommand ppEnum)
        {
            ppEnum = new CommandEnum(SubCommands);
            return 0;
        }
    }

    [ComVisible(true)]
    [Guid("9E4D8F32-5C0B-4E9F-AD23-7B8C9D0E1F2A")]
    [ClassInterface(ClassInterfaceType.None)]
    public sealed class QuickConvertBgCommand : IExplorerCommand
    {
        private static readonly List<IExplorerCommand> BgSubCommands = new List<IExplorerCommand>
        {
            new SubCommand("Merge All PDFs in Folder", "folder-merge-pdfs", "imageres.dll,-109", p => p.Count >= 1, "20000001-5C0B-4E9F-AD23-7B8C9D0E1F2A"),
            new SubCommand("All Images in Folder to PDF", "folder-images-to-pdf", "imageres.dll,-72", p => p.Count >= 1, "20000002-5C0B-4E9F-AD23-7B8C9D0E1F2A")
        };

        public int GetTitle(IShellItemArray psiItemArray, out string ppszName)
        {
            ppszName = "Quick Convert";
            return 0;
        }

        public int GetIcon(IShellItemArray psiItemArray, out string ppszIcon)
        {
            ppszIcon = "imageres.dll,-109";
            return 0;
        }

        public int GetToolTip(IShellItemArray psiItemArray, out string ppszInfotip)
        {
            ppszInfotip = "Batch convert or merge files in this folder";
            return 0;
        }

        public int GetCanonicalName(out Guid pguidCommandName)
        {
            pguidCommandName = new Guid("9E4D8F32-5C0B-4E9F-AD23-7B8C9D0E1F2A");
            return 0;
        }

        public int GetState(IShellItemArray psiItemArray, bool fOkToBeSlow, out EXPCMDSTATE pCmdState)
        {
            pCmdState = EXPCMDSTATE.ECS_ENABLED;
            return 0;
        }

        public int Invoke(IShellItemArray psiItemArray, IntPtr pbc)
        {
            return 0;
        }

        public int GetFlags(out EXPCMDFLAGS pFlags)
        {
            pFlags = EXPCMDFLAGS.ECF_HASSUBCOMMANDS;
            return 0;
        }

        public int EnumSubCommands(out IEnumExplorerCommand ppEnum)
        {
            ppEnum = new CommandEnum(BgSubCommands);
            return 0;
        }
    }
}
