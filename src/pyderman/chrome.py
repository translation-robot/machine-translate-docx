from __future__ import annotations

import re
import requests
import json
import os
import platform

from pyderman.util import downloader

_base_version = "https://chromedriver.storage.googleapis.com/LATEST_RELEASE"
_base_download = "https://chromedriver.storage.googleapis.com/{version}/chromedriver_{os}{os_bit}.zip"

def get_url(
    version: str = "latest", _os: str | None = None, _os_bit: str | None = None
) -> tuple[str, str, str]:

    match = re.match(r"^(\d*)[.]?(\d*)[.]?(\d*)[.]?(\d*)$", version)
    
    #print("version request: %s" % (version))
    
    if version == "latest":
        #print("version == latest")
        resolved_version = downloader.raw(_base_version)
    elif match:
        #print("Found match")
        major, minor, patch, _build = match.groups()
        
        if(int(major) > 14):
            return get_url_chrome_15_and_up(version, _os, _os_bit)
        else:
            return get_url_up_to_chrome_14(version, _os, _os_bit)


def get_url_chrome_15_and_up(
    version: str = "latest", _os: str | None = None, _os_bit: str | None = None
) -> tuple[str, str, str]:
    match = re.match(r"^(\d*)[.]?(\d*)[.]?(\d*)[.]?(\d*)$", version)
    
    #print("version request: %s" % (version))
    
    chromedriver_15_json_url='https://googlechromelabs.github.io/chrome-for-testing/known-good-versions-with-downloads.json'
    json_str = requests.get(chromedriver_15_json_url).content
    
    json_obj = json.loads(json_str)
    
    # An array of versions
    json_versions_array = json_obj["versions"]
    
    print(f"Searching for chromedriver version {version}")
    
    driver_architecture = ""
    
    platform_system = platform.system()
    platform_machine = platform.machine()
    
    if (platform_system.lower() == "windows"):
        if (platform_machine.lower() == "x86_64"):
            driver_architecture = "win64"
        else:
            driver_architecture = "win32"
    elif (platform_system.lower() == "darwin"):
        if(platform_machine == "x86_64"):
            driver_architecture = "mac-x64"
        else:
            driver_architecture = "mac-arm64"
    elif (platform_system.lower() == "linux"):
        driver_architecture = "linux64"
        
    if (driver_architecture == ""):
        print("Error, computer architecture not detected using os module.")
        return ""
    
    try:
        json_array_len = len(json_versions_array)
        #print("json_versions_array length : %s" % (json_array_len))
        if(version == "latest" and json_array_len > 0):
            version = json_versions_array[json_array_len-1]["version"]
            print ("Chromedriver latest version : %s" % (version))
    
        for json_version in json_versions_array:
            if json_version["version"] == version:
                json_drivers_array = json_version['downloads']['chromedriver']
                for json_drivers in json_drivers_array:
                    if json_drivers["platform"] == driver_architecture:
                        #print(f"Found driver url :{json_drivers['url']}")
                        download = json_drivers['url']
                        #print(json.dumps(json_drivers, indent=4))
                        #json_drivers_array = json_version['downloads']['chromedriver']
    except:            
        var = traceback.format_exc()
        print(var)
    
    resolved_version = version
    return "chromedriver", download, resolved_version



def get_url_up_to_chrome_14(
    version: str = "latest", _os: str | None = None, _os_bit: str | None = None
) -> tuple[str, str, str]:
    match = re.match(r"^(\d*)[.]?(\d*)[.]?(\d*)[.]?(\d*)$", version)
    
    print("version request: %s" % (version))
    
    if version == "latest":
        print("version == latest")
        resolved_version = downloader.raw(_base_version)
    elif match:
        print("Found match")
        major, minor, patch, _build = match.groups()
        if patch:
            patch_version = f"{major}.{minor}.{patch}"
        else:
            patch_version = major
            
        print("major: %s" % (major))
    
        resolved_version = downloader.raw(f"{_base_version}_{patch_version}")
        print(f"{_base_version}_{patch_version}")
        input("Press enter to continue again")
    else:
        resolved_version = version
    if not resolved_version:
        raise ValueError(f"Unable to locate ChromeDriver version! [{version}]")
    if _os == "mac-m1":
        _os = "mac"  # chromedriver_mac64_m1
        _os_bit = "%s_m1" % _os_bit
    download = _base_download.format(version=resolved_version, os=_os, os_bit=_os_bit)
    print(f"download={download}")
    return "chromedriver", download, resolved_version


if __name__ == "__main__":
    print([u for u in get_url("latest", "win", "32")])
