# SymCliXML

The aim of this python script is to convert a Symapi_db.bin from Symmetrix (VMAX AFA or PowerMax) to an XLS File, describing the box.

Prequisites on your PC :
=> Solutions Enable 9.1+
=> Python 3.x (developped in 3.8.6)
=> openpyxl (for managing XLS file)

The reference.xlsx is the description of the XLS result file with command (%%). It is parsed and mapping from memroy structure to XLS are done this way.

Currently the following obects are covered
=> Symmetrix info
=> Disks (physical)
=> Storage Groups
=> ThinDevices (Tdev).


