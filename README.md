BMx Sequence Editor by [UCN-Soft](http://ucn.tokonats.net/)
===========================================================

See [this page for information about BMSE](http://ucn.tokonats.net/software/bmse/).
Basically it's an app that lets you edit BMS files.
It's not actively maintained by the original author since 2008 or something and later they
decided to release the source code under zlib/libpng license.


About this Repository
---------------------

In this repository you will find the __original__ branch which contains the
original BMSE 1.3.8 code, and the __master__ branch with my own modifications added.


Changes to 1.3.8
----------------

* Make long notes [long](http://upic.me/show/20797623).
* Add some tools for my own convenience.


Download
--------

My modified version of BMSE is in [__BMSE-dttvb.exe__](https://github.com/dtinth/UCN-BMSE/blob/master/BMSE-dttvb.exe?raw=true).
You can use it in place of the original BMSE application, because only the program file is changed.

You can also get the additional themes (theDtTvB-*.ini) that I put into it in the
[BMSE/theme](https://github.com/dtinth/UCN-BMSE/tree/master/BMSE/theme) folder.


How to Run
----------

The recommended way to use my branch of BMSE is to download
[the original BMSE 1.3.8](http://ucn.tokonats.net/software/bmse/) and get it running first.
Please hunt for the missing DLLs first, I didn't remember how I got the required DLLs and OCXs, but
it should be a straightforward task.

Then download [__BMSE-dttvb.exe__](https://github.com/dtinth/UCN-BMSE/blob/master/BMSE-dttvb.exe?raw=true),
and use it in place of original BMSE.exe

You can also move BMSE-dttvb.exe inside BMSE folder and run it from there, as it should contain most required files,
but I don't recommend it, because the application files will get mixed with source files.



Development Environment
-----------------------

I run VB6 inside a Windows XP VM, running in VirtualBox,
with system locale set to Japanese (the host has Thai system locale).

Version control is done on the host (Windows 7), using Git inside Cygwin,
and the files are shared to the development VM using VirtualBox's shared folders.



License
-------

[zlib/libpng LICENSE](https://github.com/dtinth/UCN-BMSE/blob/master/BMSE/LICENSE)



