<#PSScriptInfo

.VERSION 1.0.0.0

.GUID ffea0459-5e1c-450b-a7dd-201e11858939

.AUTHOR Oliver Bley

.TAGS ShellFolders

.LICENSEURI https://github.com/4ourbit/ShellFolderProvider/blob/master/LICENSE.md

.PROJECTURI https://github.com/4ourbit/ShellFolderProvider

#>

<#

.DESCRIPTION
 Use ShellFolders in PowerShell

You want to dot-source this script at start of PowerShell:

powershell -Command "Invoke-Expression '. (Resolve-Path .\ShellFolderProvider.ps1)'"

#>

$t = @'
using Microsoft.Win32;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Provider;

namespace Microsoft.PowerShell.Commands {
  public class Shell {
    public string Name { get; set; }
    public object Value { get; set; }
    public object ParentFolder { get; set; }
    public object RelativePath { get; set; }
    public object Category { get; set; }
    public object Roamable { get; set; }
    public object PreCreate { get; set; }
  }

  public class ShellContentReaderWriter : IContentReader, IContentWriter, IDisposable {
    private string path;
    private ShellProvider provider;

    internal ShellContentReaderWriter(string path, ShellProvider provider) {
      this.path = path;
      this.provider = provider;
    }

    private bool contentRead;
    public IList Read(long readCount) {
      IList list = (IList) null;
      if (!this.contentRead) {
        var item = provider.FolderDescriptions[path];
        if (item != null) {
          object valueOfItem = provider.GetValueOfItem(item);
          if (valueOfItem != null) {
            list = valueOfItem as IList;
            if (list == null) list = (IList) new object[1] {valueOfItem};
          }
          this.contentRead = true;
        }
      }
      return list;
    }

    public IList Write(IList content) { return content; }
    public void Seek(long offset, SeekOrigin origin) {}
    public void Close() {}
    public void Dispose() { Close(); }
  }

  [CmdletProvider("ShellCommands", ProviderCapabilities.ShouldProcess)]
  public class ShellProvider : ContainerCmdletProvider, IContentCmdletProvider {
    private RegistryKey RegKey { set; get; }

    protected override PSDriveInfo RemoveDrive(PSDriveInfo drive) {
      if (RegKey != null) RegKey.Close();
      return drive;
    }

    protected override Collection<PSDriveInfo> InitializeDefaultDrives() {
      PSDriveInfo drive = new PSDriveInfo("shell", ProviderInfo, string.Empty, string.Empty, null);
      return new Collection<PSDriveInfo>(new[] {drive});
    }

    private IDictionary folderDescriptions;
    public IDictionary FolderDescriptions {
      get {
        if (folderDescriptions != null) return folderDescriptions;

        var fdList = new Collection<DictionaryEntry>();
        try {
          const string hklmRegKey = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\FolderDescriptions";
          RegKey = Registry.LocalMachine.OpenSubKey(hklmRegKey);
        }
        catch (Exception ex) {
            WriteError(new ErrorRecord(ex, "FolderDescriptionsUnavailable", ErrorCategory.ResourceUnavailable, FolderDescriptions));
        }
        foreach (string name in RegKey.GetSubKeyNames()) {
          var subKey = RegKey.OpenSubKey(name, RegistryKeyPermissionCheck.ReadSubTree);
          var subValue = new Shell {
            Name = subKey.GetValue("Name").ToString(),
            Value = string.Empty,
            ParentFolder = subKey.GetValue("ParentFolder"),
            RelativePath = subKey.GetValue("RelativePath"),
            Category = subKey.GetValue("Category"),
            Roamable = subKey.GetValue("Roamable"),
            PreCreate = subKey.GetValue("PreCreate")
          };
          fdList.Add(new DictionaryEntry(name, subValue));
          subKey.Close();
        }
        var fdDict = fdList.ToDictionary(entry => entry.Key, entry => entry.Value);
        fdDict.Keys.ToList().ForEach((key) => {
            Shell shell = (Shell)fdDict[key];
            if ((int)(shell.Category) >= 3) {
              for (var current = shell; current.ParentFolder != null; current = (Shell)fdDict[current.ParentFolder]) {
                if (fdDict.ContainsKey(current.ParentFolder)) {
                  var parent = (Shell)fdDict[current.ParentFolder];
                  if (parent.RelativePath != null)
                    shell.Value = "\\" + parent.RelativePath.ToString() + shell.Value;
                }
              }
              if (shell.RelativePath != null)
                shell.Value += "\\" + shell.RelativePath.ToString();
              shell.Value = Environment.GetEnvironmentVariable("USERPROFILE") + shell.Value;
            }
        });
        folderDescriptions = fdDict.ToDictionary(entry => ((Shell)entry.Value).Name, entry => entry.Value);
        return folderDescriptions;
      }
    }

    public object GetValueOfItem(object item) {
      return ((Shell)item).Value;
    }

    protected override bool ItemExists(string path) {
      if (string.IsNullOrEmpty(path)) return true;
      return FolderDescriptions.Contains(path);
    }

    protected override bool IsValidPath(string path) {
      return ItemExists(path);
    }

    protected override void GetItem(string path) {
      if (string.IsNullOrEmpty(path))
        WriteItemObject(FolderDescriptions.Values, path, true);
      else if (FolderDescriptions.Contains(path))
        WriteItemObject(FolderDescriptions[path], path, false);
    }

    protected override void InvokeDefaultAction(string path) {
      if (ItemExists(path)) Process.Start("explorer", "shell:" + path);
    }

    protected override void GetChildItems(string path, bool recurse) {
      if (string.IsNullOrEmpty(path)) {
        List<DictionaryEntry> list = new List<DictionaryEntry>(FolderDescriptions.Count + 1);
        foreach (DictionaryEntry dictionaryEntry in FolderDescriptions)
          list.Add(dictionaryEntry);
        list.Sort(((left, right) => StringComparer.CurrentCultureIgnoreCase.Compare((string) left.Key, (string) right.Key)));
        foreach (var dictionaryEntry in list)
          WriteItemObject(dictionaryEntry.Value, (string) dictionaryEntry.Key, false);
      } else if (FolderDescriptions.Contains(path))
          WriteItemObject(FolderDescriptions[path], path, false);
    }

    protected override void GetChildNames(string path, ReturnContainers returnContainers) {
      if (string.IsNullOrEmpty(path)) {
        foreach (DictionaryEntry dictionaryEntry in FolderDescriptions)
          WriteItemObject(dictionaryEntry.Key, (string) dictionaryEntry.Key, false);
      } else if (folderDescriptions.Contains(path))
          WriteItemObject(path, path, false);
    }

    protected override bool HasChildItems(string path) {
      if (string.IsNullOrEmpty(path))
        if (FolderDescriptions.Count > 0)
          return true;
      return false;
    }

    public IContentReader GetContentReader(string path) {
      return new ShellContentReaderWriter(path, this);
    }

    public IContentWriter GetContentWriter(string path) {
      return new ShellContentReaderWriter(path, this);
    }

    public object GetContentReaderDynamicParameters(string path) { return null; }
    public object GetContentWriterDynamicParameters(string path) { return null; }
    public void ClearContent(string path) {}
    public object ClearContentDynamicParameters(string path) { return null; }
  }
}
'@

$ProviderType = Add-Type -TypeDefinition $t -PassThru
Import-Module -Assembly $ProviderType.assembly
Remove-Variable t, ProviderType

Update-TypeData -TypeName Microsoft.PowerShell.Commands.Shell -DefaultDisplayPropertySet Name, Value
