# SymCliXML

The aim of this python script is to convert a Symapi_db.bin from Symmetrix (VMAX AFA or PowerMax) to an XLS File, describing the box.

Prequisites on your PC :
=> Solutions Enable 9.1+
 - symcfg command has to be in your system path
 - you have to position varialbles for offline symcli (SYMCLI_OFFLINE=1 and SYMCLI_DB_FILE to point your SYMAPI_db.bi and SYMCLI_SNAPVX_LIST_OFFLINE=enabled)
=> Python 3.x (developped in 3.8.6)
=> openpyxl (for managing XLS file) pip openpyxl install

The reference.xlsx is the description of the XLS result file with command (%%). It is parsed and mapping from memroy structure to XLS are done this way.

Currently the following obects are covered
=> Symmetrix info
=> Disks (physical)
=> Storage Groups
=> ThinDevices (Tdev).

current limits :
=> When you have mutltiple SRDF replication for 1 volume, only 1 is displayed so far.


