# SymCliXML

The aim of this python script is to convert a Symapi_db.bin from Symmetrix (VMAX AFA or PowerMax) to an XLS File, 
describing the box.

Prequisites on your PC :
=> Solutions Enable 9.1+
 - symcfg command has to be in your system path
=> Python 3.x (developped in 3.8.6)
=> git has to be installed and in you path

The reference.xlsx is the description of the XLS result file with command (%%). 
It is parsed and mapping from memroy structure to XLS are done this way.


How to install
first install dependencies :
pip install openpyxl
(or pip3 depending of you python installation)

Then download the code
git clone https://github.com/Olivier2000fr/SymCliXML

Then go to the SymCliXML Directory

python SymApiToExcel.py --help

or python SymApiToExcel.py if your symap_db_offline mode is correctory configured.