# automating-emulator-shortcuts

this repository is a collection of notes, and an example of creating shortcuts for emulators.

# usage

1. clone this repository. it represents the inteded filestructure
2. extract your emulator to the EmuDir directory
3. if it doesnt exist, create a directory in EmuDir called "games" (and "bios" if a bios is needed)
4. put your roms/isos/whatever in the games directory. games should be named as "Game Name_disk1.example"
5. edit the createShortcut.ps1 script in EmuDir to include the game files

# createShortcut.ps1 script

this script lets the user select a directory to save shortcuts for games. it has to be edited before "Install Game Shortcut.bat" can be launched

ideallyy, this repo should be rewritten. basically we just need the bat/ps script to allow the user to just select their own emulator exe and game roms to create a shortcut for them. it would be a pain in the ass though since the user would still need to make a nice icon (might be fun to try and script though) and how to implement things like different command line arguments for different emulators

# createGameShortcut.ps1 script

this script is the rewritten version of createShortcut. it asks the user for the emulator exe and which rom to use. it can also create a thumbnail for the shortcut coz it looks nicer but it only works for ps2 games right now because its a pain the ass

i think i hate most of this code but it works so its fine