# Documentation-Slayer

I want to introduce you to this VS-Code extension I have made:

✨ Documentation-Slayer ✨

Documentation Slayer is a Visual Studio Code extension that automates the creation of unit-design documentation for both Classic AUTOSAR and non-AUTOSAR projects. It lets you export your results as a structured Excel spreadsheet, a Markdown file, or a Word document—just pick the format you need.

These are mainly the most recent added features in this version:

- Supports the following 14 fields:
(Name - DESCRIPTION - Trigger - IN parameters - OUT parameters - RETURN - FUNCTION TYPE -
Inputs - Outputs - Invoked Operations - Data Types - Sync/Async - Reentrancy) ✅
- Converts to 3 Formats:
(Word docx - Excel sheet - md format) ✅
- get the shit of P2CONST(...) and P2VAR(...) out of runnables extraction.✅
- Solve the problem of duplicates between static function declarations and definitions.✅
- Exclude return from the used data type. ✅
- Detect the multiple lines for triggers. ✅
- exclude( "VStdLib_MemCmp", "memcmp", "memcpy", "memset", "sizeof", "abs", "return" ) from invoked.✅
- Detect the static and inline Runnables. ✅
- supporting all types of Rte_Read and Rte_Write in compliance with AUTOSAR RTE SWS. ✅
- Global function that isn't static needs to be parsed too, and added to a separate row in the output Excel file. ✅
- exclude MACROs from Invoked Operations with UPPERCASE ✅
- I recently added the GUI feature using Tkinter.

Please check this guide for all the info you'll need:
[Documentation Slayer ─── Installation & Usage Guide.pdf](https://github.com/user-attachments/files/20382902/Documentation.Slayer.Installation.Usage.Guide.pdf)

Created by [@abdallahissa](https://www.linkedin.com/in/abdallaissa/) - feel free to contact me!
