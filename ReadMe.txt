-------------------------------
Semi VB Decompiler by vbgamer45
Open Source
Version: 0.03
-------------------------------
Contents
1. What's New?
2. Features
3. Questions?
4. Bugs
5. Contact
6. Credits

1. What's New?

   Version 0.03
     P-Code decoding started and image extraction.
     Numerous bug fixes.
     Event detection added.
     Dll and OCX Support added.
     External Components added to vbp file.
     Begun work on a basic antidecompiler.
     Form property editor, complete with a patch report generator.
     Procedure names are recovered.
     Api's used by the program are recovered.
     Msvbvm60.dll imports are listed in the treeview.
     Syntax coloring for Forms.
     Fixed scrolling bug.

   Version 0.02
     Rebuilds the forms
     Gets most controls and their properties.

   Intial Release version 0.01

2. Features
     Decompiling the pcode/native vb6/vb5 exe's
     Form Generation, P-Code code view
     Resource extraction wmf, ico, cur, gif, bmp, jpg, dib
     Form Editor
     P-Code Procedure Decompile View
     Shows offsets for controls
     SubMain Disassembly
     Memory Map of the exe file, so you can see what's going on.
     Advanced decompiling using COM instead of hard coding property opcodes.

3. Questions?
   Q. What about Native Code Decompiling?
   A. It is in the works. I need to get a better understanding of how VBDE works before
      I begin to work on Native Code.
      A site that is working on native vb decompiler is
      http://www.Decompiler.org
   Q. What the heck are the P-Code Tokens?
   A. P-Code tokens is the last step before turning the P-Code into readable VB Code.
      All you have to do now is link the imports of the exe with the functions in P-Code.
   Q. Why does it not show all the controls on my forms?
   A. Usally because its a property that is not detected by COM using vb6.olb.
   Q. Why doesn't it get my procedure names for Modules?
   A. VB only saves procedures names for Form's and Classes.
   Q. Why is there a ComFix file?
   A. Since Visual Basic does not support all the data types that IDL does it is needed.
      Basicly it fixes when COM returns an integer when it should really be a VB byte.
   Q. How does this decompiler work?
   A. First it gets all the main vb strutures from the exe.
      Next it gets all the controls properties via COM using vb6.olb
      I am still looking for a static pointer for the table inside msvbvm60.dll to use instead.
   Q. What files does this decompiler require?
   A. It requires the following files:
      TLBINF32.dll
      comdlg32.OCX
      RICHTX32.OCX
      MSCOMCTL.OCX
      TABCTL32.OCX
      MSFLXGRD.OCX
      Msvbvm60.dll
      And VB6.olb version 6.0.9
      All of the above files need to be registered!
   Q. Where can I learn more about Visual Basic 5/6 Decompiling?
   A. Head over to http://www.vb-decompiler.com/  tons of information on vb decompiling.

4. Bugs
     I know about most of them...
     MDI Forms and External Controls.
     Some properties aren't handled yet dataformat, and some others
     P-Code decoding may hang use the disable P-Code option
     Overflow error is caused by a property that isn't detected yet...
     Currently it does not generate user control and property pages

5. Contact
     Email=gmdecompiler@yahoo.com
     Aim=vbgamer45

6. Credits
     I would like to thank the following people for helping me with this project.
     Sarge, Mr. Unleaded, Moogman, _aLfa_, ionescu007, Warning and many others.
