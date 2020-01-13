# WinPrinterFix
Script that fixes issue with printer deployment to new profiles in Windows. This is especially important in environments where profiles are deleted regularly (e.g. education).

It appears that Windows simply does not create the required registry entries, this script aims to resolve that by creating them manually.

Script should be run as a logon script.
