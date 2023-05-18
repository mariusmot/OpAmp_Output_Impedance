# OpAmp_Output_Impedance
OpAmp Output Impedance script

The source files must be in the same folder as the script files so that the script can copy necessary data!

Excel files with manual test data will be used as source files for datasheet data. All must be edited under below pattern:

Name of the file: part + _WithScores (eg. **AD8420_WithScores**). 

Script can handle file names with multiple devices (eg. **ADA4627-1_ADA4637-1_WithScores**).

Name of the sheet containing Datasheet numbers: **Datasheet**

Columns containing datasheet data must have a **row Header** that specifies the gain used (eg. **G10 freq** and **G10 values**). Data in these columns should start with second row
