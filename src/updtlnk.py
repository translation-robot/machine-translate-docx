import os
import sys
import platform
import pythoncom
from win32com.client import Dispatch

WINDOWS_TERMINAL_EXE = r"C:\SMTVRobot\WindowsTerminal\WindowsTerminal.exe"

def is_windows_10_or_later():
    """Return True if running on Windows 10 or newer."""
    try:
        win_ver = sys.getwindowsversion()
        # Major version 10 = Windows 10/11
        return win_ver.major >= 10
    except Exception:
        # If check fails for any reason, assume not compatible
        return False

def get_sendto_folder():
    """Return the path to the current user's SendTo folder."""
    shell = Dispatch("WScript.Shell")
    return shell.SpecialFolders("SendTo")

def modify_lnk_target(lnk_path):
    try:
        shell = Dispatch('WScript.Shell')
        shortcut = shell.CreateShortcut(lnk_path)

        original_target = shortcut.TargetPath
        args = shortcut.Arguments or ""

        combined = f"{original_target} {args}"
        if "-run " in combined:
            after_run = combined.split("-run ", 1)[1].strip()

            shortcut.TargetPath = WINDOWS_TERMINAL_EXE
            shortcut.Arguments = after_run

            if not shortcut.WorkingDirectory:
                shortcut.WorkingDirectory = os.path.dirname(WINDOWS_TERMINAL_EXE)

            shortcut.save()
            print(f"✅ Modified: {lnk_path}")
        else:
            print(f"ℹ️ No '-run ' found in: {lnk_path}")

    except Exception as e:
        print(f"⚠️ Error processing {lnk_path}: {e}")

def main():
    # Check Windows version
    if not is_windows_10_or_later():
        print("❌ This script requires Windows 10 or later. No modifications were made.")
        sys.exit(0)

    pythoncom.CoInitialize()

    # Determine folder path
    if len(sys.argv) > 1:
        folder_path = sys.argv[1]
    else:
        folder_path = get_sendto_folder()
        print(f"No folder argument given — using SendTo folder:\n{folder_path}\n")

    if not os.path.isdir(folder_path):
        print(f"Error: '{folder_path}' is not a valid folder.")
        sys.exit(1)

    modified_count = 0
    for file in os.listdir(folder_path):
        if file.lower().endswith(".lnk"):
            full_path = os.path.join(folder_path, file)
            modify_lnk_target(full_path)
            modified_count += 1

    pythoncom.CoUninitialize()
    print(f"\nDone. Processed {modified_count} shortcut(s).")

if __name__ == "__main__":
    main()
