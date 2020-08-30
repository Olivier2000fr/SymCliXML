Welcome to SymCliXML
====================

Overview
--------

The aim of this python script is to convert a Symapi_db.bin from Symmetrix (VMAX AFA or PowerMax) to an XLS File, 
describing the box.

The reference.xlsx is the description of the XLS result file with command (%%). 
It is parsed and mapping from memroy structure to XLS are done this way.


+-----------------------+----------------------------+
| **Author**            | Olivier Guyot              |
+-----------------------+----------------------------+
| **Unisphere Version** | 9.1.                       |
+-----------------------+----------------------------+
| **Array Model**       | VMAX-3, VMAX AFA, PowerMax |
+-----------------------+----------------------------+
| **Platforms**         | Linux, Windows             |
+-----------------------+----------------------------+
| **Python**            | 3.8                        |
+-----------------------+----------------------------+
| **Requires**          | openpyxl                   |
+-----------------------+----------------------------+



Installation
------------

first install dependencies :
pip install openpyxl
(or pip3 depending of you python installation)

Then download the code
git clone https://github.com/Olivier2000fr/SymCliXML

Then go to the SymCliXML Directory

python SymApiToExcel.py --help

or python SymApiToExcel.py if your symap_db_offline mode is correctory configured.