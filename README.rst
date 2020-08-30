Welcome to SymCliXML
====================

Overview
--------

The aim of this python script is to convert a Symapi_db.bin from Symmetrix (VMAX AFA or PowerMax) to an XLS File, 
describing the box.

The reference.xlsx is the description of the XLS result file with command (%%). 
It is parsed and mapping from memroy structure to XLS are done this way.


+-----------------------+-------------------------------------------+
| **Author**            | Olivier Guyot (olivier.guyot2@gmail.com)  |
+-----------------------+-------------------------------------------+
| **Symcli version**    | 9.1                                       |
+-----------------------+-------------------------------------------+
| **Array Model**       | VMAX-3, VMAX AFA, PowerMax                |
+-----------------------+-------------------------------------------+
| **Platforms**         | Linux, Windows                            |
+-----------------------+-------------------------------------------+
| **Python**            | 3.8                                       |
+-----------------------+-------------------------------------------+
| **Requires**          | openpyxl                                  |
+-----------------------+-------------------------------------------+



Installation
------------

First install dependencies::

    $ pip install openpyxl

Then download the code::

    $ git clone https://github.com/Olivier2000fr/SymCliXML

Then go to the SymCliXML Directory::

    $ python SymApiToExcel.py --help

    usage: SymApiToExcel.py [-h] [-sid SID] [-symapi_dir SYMAPI_DIR]
                            [-symapi_db SYMAPI_DB] [-all] [-local]

    SympApiToExcel helps you to translate the configuration of a given Symmetrix to an XLS File. The program needs to have
    symcli 9.1 installed and in your path, as well as a openpywl in your python installation.

    optional arguments:
    -h, --help            show this help message and exit
    -sid SID              Allow you to precise a SID (needs to be fully precise as in the symapi)
    -symapi_dir SYMAPI_DIR
                        allow you to precise a directory whe the symapi_db's are located, symapi_dbs should be in the
                        form symapi*.bin
    -symapi_db SYMAPI_DB  allow you to precise a precise SYMAPI_DB.bin
    -all                  will run against all SYMIDs in the symapi_db
    -local                will run against all local SYMIDs in the symapi_db

or if your symap_db_offline mode is correctory configured.


 $ python SymApiToExcel.py

exemple d'output::

 $ C:/Users/guyoto/PycharmProjects/SymCliXML/SymApiToExcel.py -symapi_dir c:\temp\SYMAPI
    Parameter : c:\temp\SYMAPI
    0 - symapi_db.bin
    1 - symapi_db_0064.bin
    2 - symapi_db_0077.bin
    3 - symapi_db_0106_240820.bin
    4 - symapi_db_0107_250820.bin
    5 - symapi_db_0109.bin

    please enter the id of the symapi to process (0 to 5) or QUIT
    which symapi id : 0

    Selected Symapi is : c:\temp\SYMAPI\symapi_db.bin

    List of discovered systems :
    0 - 000297700XXX (Local) - VMAX950F
    1 - 000297700YYY (Local) - VMAX950F
    2 - 000297600ZZZ (Remote) - PowerMax_8000
    3 - 000297600AAA (Remote) - PowerMax_8000
    4 - 000297600BBB (Remote) - PowerMax_8000

    please enter the system to process (0 to 4) or ALL or QUIT
    which system to process :

