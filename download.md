
# Download binaries for the machine translation robot.

The latest version of the translation robot is 2024-10-12 (repackaged on 2024-11-03)

## Windows version
Require Chrome browser installed. No other software required. Download google chrome here :

https://www.google.com/chrome/

### Download the binaries for windows 10 and 11 (64 bits) here:

https://1drv.ms/u/c/25c35a16b8db8a90/EZKEH7LS8r5KlYNaPS7ql10BYIglv6tU8giiSACymuqHag?e=mHWiMf

### Download the binaries for windows 7, 8, 8.1 (32 bits, compatible 64 bits) here:

Not provided anymore

## Mac version

Require Chrome browser installed. No other software required. Download google chrome here :

https://www.google.com/chrome/

Download the binaries for macos here:

Mac OS X dmg installer:

https://1drv.ms/u/c/25c35a16b8db8a90/EeBrkQotaZFDt9HD4mcNDnIBPg1lykgsOkPc4Ks6Fbb5Qw?e=R1XsOn

Before installing the DMG installer, disable mac Gatekeeper:
1. From Launchpad open the Term App
2. Run this command:

    sudo spctl --master-disable

    (your mac account password will be required)

This will allow unsigned apps to run on Mac OS X.

Only after disabling Gatekeeper, double click on the DMG installer file to install the translation robot, otherwise you will have to reinstall the program again.

After opening the program "Machine Translator" in the LaunchPad, you will have to accept the program to run, at least three times, once for the graphical interface, once for the robot, and once for the selenium-manager than handles communication between the robot and Google Chrome.

After running the Machine Translator and translating one file from the Machine Translator application, disable mac Gatekeeper again:

1. From Launchpad open the Term App
2. Run this command:

    sudo spctl --master-enable

    (your mac account password will be required)
	
This will disallow new unsigned apps to run on Mac OS X.