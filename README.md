# Sound_MMTimer
## Use MMTimer in your exe, with PostMessage in tlb  
Project was started in Juli 2008.
in VBC for a WinAPI-function the member attribute usesgetlasterror is on by default, every WinAPI-error is handed over to the VB-Err-object.
This behaviour leads to a computer crash by calling the callbyck-function by the MMTimer-Thread, what again calls a Message (Post- or Sendmessage).
in Matthew Curlands "Type Library Editor" if you define a API function the Member-Attribute "usesgetlasterror" can be switched on, but ist off by default.
![MCPowerVBTypeLibraryEditorPostMessagetlb.png Image](Resources/MCPowerVBTypeLibraryEditorPostMessagetlb.png "MCPowerVBTypeLibraryEditorPostMessagetlb.png Image")

however, this project is obsolete now. You can find a class from Frank Schüler, that has some asm-codes to successfully circumnavigate all cliffs.
No extra UserControl nor Picturebox, nor Standard-Modul for a callback-function is needed. All happens inside the class MMTimer.
