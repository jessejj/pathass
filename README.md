# pathologyAssistant
Assisting PAs through scripts

**diffMinutesFromTimestamps.vb** can be used as a sub routine in a macro-enabled excel workbook. It calculates time beween two timestamps.

**Next Item.vbs** is a script intended to be launched using cscript.exe on 64-bit Windows. It primarily interacts with Word documents while dictation software is being used (such as Dragon), moving through the document automatically without having to say 'Next item'. 

**backupReport.vb** can be referenced as a function in Dragon Advanced scripting commands. It silently makes a chronological backup of the current text in the active Word document, without interupting or otherwise causing a distraction.

**checkReport.vb** can be referenced as a sub routine in Dragon Advanced scripting commands. It silently checks the test of the Word document for dictation errors.