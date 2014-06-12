PSInventory
===========

PowerShell script for computer hardware inventory.


This is my first attempt at PowerShell scripting after giving myself a crash course on the subject. By no means take these initial versions as 'finished' and ready for use on large systems - I have yet to fully assess the impact it has (testing using domain of ~500 machines, across 10 different subnets).

The intent is to create a smart script that is both meaningful for general administration, and easy to use by others, on systems big and small.

Objectives will include:
-A working script across all Windows environments, with intelligent design to bypass and deal with any differences.
-Output displayed in a readable fashion, with comparative checks against previous versions for hardware changes.
-Support for custom settings for specific user case scenarios.
-Easy for anyone to run and use on their system.
