# ucAniGif
Animated GIF ActiveX Control

ucAniGif is a simple example of using twinBASIC to create an ActiveX control usable in both 32bit and 64bit VBA hosts (in addition to VB6 and .NET). tB is backwards compatible with VB6/VBA language, and uses the VBA7 syntax for 64bit support.

![Screenshot](https://i.imgur.com/NMQzUau.gif)

It's very simple, you can set the file at runtime or design time; if you set it at design time you can set Autoplay. There's Play/Pause/Stop commands. And a SizeToFit option to scale the image up or down for the control size.

This can be built with the free version of twinBASIC from ucAniGif.twinproj, but the 64bit OCX will then display a tB splash screen when loaded. Of course subscribing to this wonderful project is the best option, but the Releases section has binary builds without the 64bit version splash screen.\
NOTE: The OCX build file (ucAniGif.twinproj) requires tB to run as admin; this is because it uses the 'Register project to HKEY_LOCAL_MACHINE' option. This is neccessary for some hosts to see the control, but for VBA only you can switch the option to off and run without admin if need be, but this is not required, VBA will see it regardless.

The code itself is very simple; it's just a thin wrapper over IShellImageData. But of course, using COM interfaces is painful in VBA 64bit without typelibs like my oleexp. tB is not only the only way to build 64bit ActiveX controls for 64bit Office using the same language, it also supports easily defining COM interfaces with BASIC syntax in the project, and has my WinDevLib project available, which allows development with thousands of interfaces and APIs already available.

You can browse the source online in ucAniGif.twin, or import the .tbcontrol/.twin/mDefs.twin files into a tB project to use it there, or play with the test project.

If you haven't checked out twinBASIC before, the FAQ is a great place to start: https://github.com/twinbasic/documentation/wiki/twinBASIC-Frequently-Asked-Questions-(FAQs)

PS - Transparency is respected, but you have to set the control's BackColor manually in this initial version:

![image](https://github.com/fafalone/ucAniGif/assets/7834493/55e35ffe-dacc-493e-ae4b-04ffeb900aba)

e.g.

```vba
Private Sub UserForm_Initialize()
 ucAniGif1.BackColor = Me.BackColor
End Sub
```

### How install and select for use in VBA

#### From binaries

There's already-compiled binary builds available under the 'Releases' link on the right side column of this page, [or click here](https://github.com/fafalone/ucAniGif/releases) to go directly to the Releases page. The  zip contains a win32 folder and a win64 folder, you need to use the one that matches Office for 32bit or 64bit. Once you've extracted the one you need, if it's 32bit, drag/drop `ucAniGif.ocx` in Explorer onto the `regsvr32.exe` in the Windows\SysWOW64 folder. For 64bit, drop on the one in Windows\System32 (unless you're using 32bit version of Windows, then just that too).

> [!TIP]
> If you don't know whether you have 32bit or 64bit Office, go to File->Account then click 'About Excel/Access/etc'


#### From source

As noted above, this project was written in twinBASIC. If you don't have it already you can download it from [here](https://github.com/twinbasic/twinbasic/releases). Just download and extract, there's no installer. The Community Edition is free, and will build both 32bit and 64bit OCXs-- the only limitation being that the 64bit OCX will display a splash screen for tB when it loads. The OCX in the Releases section here will not as I've got the Pro version.\
You need the ucAniGif.twinproj file from the source code files of this repository. That contains the entire project, the other files are for browsing the source code in your browser or the ucAniGifTest.twinproj file is a demo of using the control within twinBASIC. Use the 'Browse' option to open the .twinproj file, in the bottom left of the opening dialog when you launch twinBASIC. After that, look in the toolbar for a dropdown list that says 'win32'... that means it will compile as 32bit. Select win64 from the dropdown to compile as 64bit. After that, click on File->Build. It will create and automatically register the OCX, you're all set to open Office and start using it

#### Finding it in VBA

Once you've done one of the options above, `ucAniGif` should be available in the Tools->Additional controls dialog when you're editing a UserForm in Excel VBA, or 'ActiveX Controls' in the Access form designer-- the menu that pops up from the dropdown button on the righthand side of the built-in controls box.
