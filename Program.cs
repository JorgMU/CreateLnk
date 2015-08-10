using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Text.RegularExpressions;

//this exists outside the namespace to be used by the Option class
//using an enum prevents typos in the IDE
enum AllowedOptions
{
  Options,
  Comment,
  FullName,
  IcoFile,
  IcoIndex,
  Link,
  Target,
  Show,
  StartIn,
  Verbose,
  WhatIf,
}

namespace CreateLnk
{
  class Program
  {
    private static readonly string _exePath = Environment.GetCommandLineArgs()[0];
    private static readonly string _exeName = Path.GetFileName(_exePath);

    private static Options _opts;

    private static bool BeVerbose = false;
    private static bool WhatIf = false;

    static void Main(string[] args)
    {
      Console.WriteLine("CreateLnk - jorgie@missouri.edu ({0})\r\n", _exeName);

      if (args.Length < 1)
      {
        ShowUse();
        Environment.Exit(0);
      }

      _opts = new Options(args);

      if (_opts.ActiveOptions.Length < 1)
      {
        FatalExit("No options recognized.");
      }

      if (_opts.OptionExists(AllowedOptions.Verbose)) BeVerbose = true;
      if (_opts.OptionExists(AllowedOptions.WhatIf)) WhatIf = true;
      Verbose("Maximum verbosity enabled!");
      if (WhatIf) Push("* WhatIf enabled, no Link will be created.");

      Verbose(_opts.ToString());

      CreateShortcut();
    }

    private static void FatalExit(string Message)
    {
      ShowUse();
      Console.WriteLine("! " + Message);
      Environment.Exit(-1);
    }

    private static void FatalExit(string Message, params object[] Items)
    {
      FatalExit(string.Format(Message, Items));
    }

    private static void CreateShortcut()
    {
      //LinkPath
      string link = _opts[AllowedOptions.Link];
      if (!Validate(ref link, false))
        FatalExit("Cannot validate LinkPath:\r\n  [{0}]", link);
      string parent = Path.GetDirectoryName(link);
      if (!ValidateDirectoryPath(ref parent, true))
        FatalExit("Cannot finde directory for LinkPath:\r\n  [{0}]", link);
      if (File.Exists(link))
        Push("* Link exists and will be overwritten");
      Verbose("{0}: {1}", AllowedOptions.Link.ToString(), link);

      //TargetPath
      string target = _opts[AllowedOptions.Target];
      if (!Validate(ref target, true))
        FatalExit("Cannot validate TargetPath:\r\n  [{0}]", target);
      Verbose("{0}: {1}", AllowedOptions.Target.ToString(), target);

      //Options
      string optionCheck = "[^'\"A-Za-z0-9:\\\\|/\\[\\]{}<>.!?,;:!@#$%^&*()_+= -]";
      string arguments = Regex.Replace(_opts[AllowedOptions.Options], optionCheck, "");
      Verbose("{0}: {1}", AllowedOptions.Options.ToString(), arguments);


      //WindowStyle
      int iWindowStyle = 1; //default normal window
      string sWindowStyle = _opts[AllowedOptions.Show].ToUpper();
      if (sWindowStyle != "")
        switch (sWindowStyle)
        {
          case "NORMAL":
            iWindowStyle = 1;
            break;
          case "MIN":
            iWindowStyle = 7;
            break;
          case "MAX":
            iWindowStyle = 3;
            break;
          default:
            FatalExit("Cannot validate WindowStyle:\r\n  [{0}]", sWindowStyle);
            break;
        }
      Verbose("{0}: {1} [{2}]", AllowedOptions.Show.ToString(), sWindowStyle, iWindowStyle);

      //startIn
      string startIn = _opts[AllowedOptions.StartIn];
      if (startIn != "" && !ValidateDirectoryPath(ref startIn, true))
        FatalExit("Cannot validate WorkingDirectory:\r\n  [{0}]", startIn);
      Verbose("{0}: {1}", AllowedOptions.StartIn.ToString(), startIn);

      //IconFile
      string iconFile = _opts[AllowedOptions.IcoFile];
      if (iconFile != "" && !Validate(ref iconFile, true))
        FatalExit("Cannot validate IconLocation:\r\n  [{0}]", iconFile);
      Verbose("{0}: {1}", AllowedOptions.IcoFile.ToString(), iconFile);

      //IconIndex
      string sIconIndex = _opts[AllowedOptions.IcoIndex];
      int iIconIndex = 0;
      if (sIconIndex != "" && !int.TryParse(sIconIndex, out iIconIndex))
        FatalExit("Cannot validate IconIndex:\r\n  [{0}]", iIconIndex);
      Verbose("{0}: {1} [{2}]", AllowedOptions.IcoIndex.ToString(), sIconIndex, iIconIndex);

      string comment = Regex.Replace(_opts[AllowedOptions.Comment], optionCheck, "");
      Verbose("{0}: {1}", AllowedOptions.Comment.ToString(), comment);

      //Adjust
      if (!link.EndsWith(".lnk")) link += ".lnk";
      _opts[AllowedOptions.Link] = link;  //do I really need to update the opts?

      //Create shortcut
      if (!WhatIf)
      {
        try
        {
          IShellLink shortcut = (IShellLink)new ShellLink();
          shortcut.SetPath(target);
          shortcut.SetShowCmd(iWindowStyle);
          shortcut.SetWorkingDirectory(startIn);
          if (iconFile != "")
            shortcut.SetIconLocation(iconFile, iIconIndex);
          shortcut.SetDescription(comment);
          shortcut.SetArguments(arguments);

          IPersistFile file = (IPersistFile)shortcut;
          file.Save(link,false);

        }
        catch (SystemException se)
        {
          Console.WriteLine("Error creating link: " + se.Message);
          if (se.InnerException != null) Console.WriteLine("InnerException: " + se.InnerException.Message);
        }
      }

      //end CreateShortcut
    }

    public static bool Validate(ref string Path, bool MustExist)
    {
      if (Path == "") return false;
      FileInfo fi = null;

      try { fi = new FileInfo(Path); }
      catch { }
      if (fi == null) return false;

      Path = fi.FullName;

      if (!MustExist || (MustExist && fi.Exists)) return true;

      return false;
    }

    public static bool ValidateDirectoryPath(ref string Path, bool MustExist)
    {
      if (Path == "") return false;
      DirectoryInfo di = null;

      try { di = new DirectoryInfo(Path); }
      catch { }
      if (di == null) return false;

      Path = di.FullName;

      if (!MustExist || (MustExist && di.Exists)) return true;

      return false;
    }

    public static void Push() { Push(""); }
    public static void Push(string Message)
    {
      Console.WriteLine(Message);
    }

    public static void Push(string Message, params object[] Items)
    {
      Push(string.Format(Message, Items));
    }

    public static void Verbose(string Message)
    {
      if (_opts == null) throw new SystemException("Vebose method called before options were available.");

      if (BeVerbose) Push("v " + Message);
    }

    public static void Verbose(string Message, params object[] Items)
    {
      Verbose(string.Format(Message, Items));
    }

    public static void ShowUse()
    {
      StringBuilder sb = new StringBuilder("\r\n  Use:");
      sb.AppendFormat("  {0} [opt1:val1] [opt2:val2]\r\n\r\n", _exeName);

      sb.AppendFormat("  {0,10}: Full path to LNK file (required)\r\n", AllowedOptions.Link.ToString());
      sb.AppendFormat("  {0,10}: Full path to target (required)\r\n", AllowedOptions.Target.ToString());
      sb.AppendFormat("  {0,10}: Options to be passed to the target\r\n", AllowedOptions.Options.ToString());
      sb.AppendFormat("  {0,10}: Normal (default), Min, Max\r\n", AllowedOptions.Show.ToString());
      sb.AppendFormat("  {0,10}: Full path to working directory\r\n", AllowedOptions.StartIn.ToString());
      sb.AppendFormat("  {0,10}: Full path to file containing icon\r\n", AllowedOptions.IcoFile.ToString());
      sb.AppendFormat("  {0,10}: Int index into incon file, default is 0\r\n", AllowedOptions.IcoIndex.ToString());
      sb.AppendFormat("  {0,10}: Text for the link's description field\r\n", AllowedOptions.Comment.ToString());
      sb.AppendFormat("  {0,10}: Show result but do not create link\r\n", AllowedOptions.WhatIf.ToString());
      sb.AppendFormat("  {0,10}: Display extra info\r\n", AllowedOptions.Verbose.ToString());

      //sb.AppendLine("  Arguments:text - desc");
      //sb.AppendLine("  FullName:text - desc");
      //sb.AppendLine("  Hotkey:text - desc");
      //sb.AppendLine("  RelativePath:text - desc");

      Push(sb.ToString());
    }

    //end Class program
  }

  #region COM_STUFF
  [ComImport]
  [Guid("00021401-0000-0000-C000-000000000046")]
  internal class ShellLink
  {
  }

  [ComImport]
  [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
  [Guid("000214F9-0000-0000-C000-000000000046")]
  internal interface IShellLink
  {
    void GetPath([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszFile, int cchMaxPath, out IntPtr pfd, int fFlags);
    void GetIDList(out IntPtr ppidl);
    void SetIDList(IntPtr pidl);
    void GetDescription([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszName, int cchMaxName);
    void SetDescription([MarshalAs(UnmanagedType.LPWStr)] string pszName);
    void GetWorkingDirectory([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszDir, int cchMaxPath);
    void SetWorkingDirectory([MarshalAs(UnmanagedType.LPWStr)] string pszDir);
    void GetArguments([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszArgs, int cchMaxPath);
    void SetArguments([MarshalAs(UnmanagedType.LPWStr)] string pszArgs);
    void GetHotkey(out short pwHotkey);
    void SetHotkey(short wHotkey);
    void GetShowCmd(out int piShowCmd);
    void SetShowCmd(int iShowCmd);
    void GetIconLocation([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszIconPath, int cchIconPath, out int piIcon);
    void SetIconLocation([MarshalAs(UnmanagedType.LPWStr)] string pszIconPath, int iIcon);
    void SetRelativePath([MarshalAs(UnmanagedType.LPWStr)] string pszPathRel, int dwReserved);
    void Resolve(IntPtr hwnd, int fFlags);
    void SetPath([MarshalAs(UnmanagedType.LPWStr)] string pszFile);
  }

  #endregion




  //end namespace CreateLnk
}


