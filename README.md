# RegShellWindow v2.1 - Shell Window Registration

Register your app as displaying a path for SHOpenFolderAndSelectItems etc.

![image](https://user-images.githubusercontent.com/7834493/213883166-746bb8a4-81b9-4869-8ada-0a9dbaacc12c.png)

**This is a 64-bit compatible port for twinBASIC of [the original VB6 version](https://www.vbforums.com/showthread.php?894889-VB6-Using-IShellWindows-to-register-for-SHOpenFolderAndSelectItems).**

People who've made Explorer replacements or other apps that display a window have found that they can, at best, get the folder passed to them when someone uses 'Show in folder'/'Show in Explorer' in apps like Firefox, Chrome, and torrent clients. However, if you register everything properly, any app calling SHOpenFolderAndSelectItems will be able to pass the names of selected files to a folder you've registered as displaying.

Note that for it to start your app if it's not running with the path in question registered, it will have to be the default that opens when you pass a folder to ShellExecuteEx (i.e. registered as an Explorer replacement on the system).

After Windows XP the API changed. I've completed this demo for Windows 7-10, which uses the SelectItem in IShellView. Windows XP uses IShellFolderViewDual instead.

### Update

v2.1 fixes a bug in tbShellLib where an entry is missing from IFolderView, breaking programs that query more detailed info.

### Requirements
[twinBASIC Beta 236 or newer](https://github.com/twinbasic/twinbasic/releases)

This code uses references to tbShellLib and the new Implements-compatible version, tbShellLibImpl, the successor to oleexpimp.tlb.

---

The question of how to do this has been asked many times in various forums for various languages, but apart from alluding to the need to use IShellWindows.RegisterPending, nobody definitely answered.

Special thanks to The_trick for identifying the problem with MSDN's IShellWindows documentation... RegisterPending absolutely does not accept a VT_VARIANT|VT_BYREF variant, but will accept any number of types of variants that can be converted to a pidl. This demo uses a simple BSTR (VB's String type), and The_trick notes it also accepts a CSIDLs, IPersistIDList, IPersistFolder2, IStream, and Byte Arrays (perhaps containing a full ITEMIDLIST structure rather than a pointer to one?). More work is needed to identify the best way to work with this.

The key is using both RegisterPending and Register, since RegisterPending accepts a path but not an object implementing all the required interfaces, and Register takes the object but not a path (and it never tries looking anything up). It uses APIs to convert the ThreadID to an hWnd to associate them-- you'll find that the cookie RegisterPending returns will be the same as the one RegisterReturns, even if you don't use the same variable like my code; this shows that, unlike MSDN indicates, this calls are two parts of the same whole.

In a full app, you'd call RegisterPending, create the file display, then finalize with Register.

```
    Set pSW = New ShellWindows
    Set pDisp = ucSW1.object
    pSW.RegisterPending App.ThreadID, Text1.Text, VarPtr(vre), SWC_BROWSER, lCookie
    pSW.Register pDisp, Form1.hWnd, SWC_BROWSER, lCookie
```

In the future I plan on releasing a more complete implementation of these various folder interfaces that work with my Shell Browser control, to create a full replacement for Explorer.

As a bonus, the UserControl has full prototypes for several additional interfaces that weren't used.

### Notes

-RegisterPending and Register are linked... two parts of the same whole both needed for a shell window. The way it wprks, your app can register multiple locations with RegisterPending, but each one must have a unique hWnd associated with it... this is because when Register is called to finalize it, it looks up the thread with GetWindowThreadProcessId to complete the information.

-If you do register multiple windows, you must complete them one at a time (RegisterPending, load, Register)... because each Register call will take the most recent matching thread ID because it steps backwards through the structures that have been added.

-Whenever you're closing out a folder, call IShellWindows.Revoke with the cookie (keep track of them per-window). Register will return the same cookie RegisterPending did; if these get mismatched, it would be because you called RegisterPending multiple times before Register, and represent a problem.

### UPDATE: Version 2
-Project has now been updated to provide all the details requested by the IShellWindows enumeration demo:
[[VB6] Get extended details about Explorer windows by getting their IFolderView](https://www.vbforums.com/showthread.php?818959-VB6-Get-extended-details-about-Explorer-windows-by-getting-their-IFolderView)

-It's been established that only Windows XP uses IShellFolderViewDual.SelectItem... Windows 7-10 use IShellView. I wanted to solve it for XP purely for the challenge though. Getting the pidl was very, very complicated, as it was stored inside a SAFEARRAY inside a VT_ARRAY|VT_UI4 variant. It took some doing but I worked it out:

```
Dim pidlItem As LongPtr
Dim sa1 As SAFEARRAY1D
Dim pSA As Long

CopyMemory ByVal VarPtr(pSA), ByVal (VarPtr(pvfi) + 8&), LenB(pidlItem) 'Offset 8 is where the data begins. Here, a pointer to the SAFEARRAY
CopyMemory ByVal VarPtr(sa1), ByVal pSA, LenB(sa1) 'We now have the full SAFEARRAY struct...
pidlItem = sa1.pvData '...and pvData is our pidl.
```

With that completed that functionality is now available on Windows XP, however the parts of the demo responding to the IShellWindows enumeration demo use IShellItem APIs not present on XP that you'll need to replace.
