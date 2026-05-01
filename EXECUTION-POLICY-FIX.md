# Fix: PowerShell Execution Policy Error

If you see this error when running `WinCleaner.ps1`:

```
.\WinCleaner.ps1 : File ... cannot be loaded because running scripts is disabled on this system.
```

This means PowerShell's execution policy is blocking script execution. Choose one of the fixes below.

---

## Option 1: Run Once (No Permanent Change) — Recommended

Bypass the policy for a single execution without changing any system settings:

```powershell
powershell -ExecutionPolicy Bypass -File .\WinCleaner.ps1 -SearchPaths "D:\Pictures"
```

---

## Option 2: Unblock the Downloaded File

If the script was downloaded from the internet, Windows may have blocked it. Unblock it with:

```powershell
Unblock-File -Path .\WinCleaner.ps1
```

Then run the script normally:

```powershell
.\WinCleaner.ps1 -SearchPaths "D:\Pictures"
```

---

## Option 3: Change Execution Policy for Current User (Permanent)

To allow locally written scripts to run without affecting system-wide policy:

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

- **RemoteSigned** allows local scripts to run freely and requires downloaded scripts to be signed.
- This change only applies to your user account.

To confirm the change:

```powershell
Get-ExecutionPolicy -List
```

---

## Option 4: Change Execution Policy for Current Session Only

Applies only to the current PowerShell window and reverts when closed:

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process
```

---

## Reverting Changes

To restore the default restricted policy for your user:

```powershell
Set-ExecutionPolicy -ExecutionPolicy Undefined -Scope CurrentUser
```

---

## Reference

- [Microsoft Docs — about_Execution_Policies](https://go.microsoft.com/fwlink/?LinkID=135170)
