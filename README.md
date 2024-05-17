# ucAniGif
Animated GIF ActiveX Control

ucAniGif is a simple example of using twinBASIC to create an ActiveX control usable in both 64bit VBA hosts. tB is backwards compatible with VB6/VBA language, and uses the VBA7 syntax for 64bit support.

![Screenshot](https://i.imgur.com/NMQzUau.gif)

It's very simple, you can set the file at runtime or design time; if you set it at design time you can set Autoplay. There's Play/Pause/Stop commands. And a SizeToFit option to scale the image up or down for the control size.

This can be built with the free version of twinBASIC from ucAniGif.twinproj, but the 64bit OCX will then display a tB splash screen when loaded. Of course subscribing to this wonderful project is the best option, but the Releases section has binary builds without the 64bit version splash screen.

The code itself is very simple; it's just a thin wrapper over IShellImageData. But of course, using COM interfaces is painful in VBA 64bit without typelibs like my oleexp. tB is not only the only way to build 64bit ActiveX controls for 64bit Office, it also supports easily defining COM interfaces with BASIC syntax in the project, and has my WinDevLib project available, which allows development with thousands of interfaces and APIs already available.

You can browse the source online in ucAniGif.twin, or import the .tbcontrol/.twin files into a tB project to use it there, or play with the test project.

If you haven't checked out twinBASIC before, the FAQ is a great place to start: https://github.com/twinbasic/documentation/wiki/twinBASIC-Frequently-Asked-Questions-(FAQs)

