"""
 VplexXMLToExcel

 Transforme le XML de configuration d'un VPLEX en fichier XLS

 Author : Olivier Guyot

 Developped during off hours (vacations)
 Code is licences under GNU GPL v3

 TODO :
    TOUT

"""
import logging.config
import argparse
import xml.etree.ElementTree as ET


"""
Initialize logging


File SymApiToExcel.logging mus be present and describe all availables logging configuration
Per default in the file :
2 loggers : 1 for console; one in file. (no rotation).
File is more for debug purposes.
In case of issue, truncate the log file or upgrade the logger level to Fatal

"""
logging.config.fileConfig('SymApiToExcel.logging')
logger = logging.getLogger('root')


class mesObjets:
    """
    meObjets : mother class for all custom objects.

    you will find static methoods for code refactoring (why do it 10 times when you can write once and call many (find / findall)

    You will also find toString methods that transform an object to string listing all attributes to string.
    if an attribute is a list, then data are not printed.
    """


    def toString(self) -> str:
        """
        toString : put in a nice string all attributes of an object.

        :return: A string cotaining all attributes and values of an object
        """
        result = ""
        variables = self.__dict__.items()
        for variable, value in variables:
            if (str(variable).startswith("list_")):
                result = result + variable + " ====> LIST " + str(len(list(value))) + " Elements \n"
            else:
                result = result + variable + "  ====>  " + str(value) + "\n"
        return result

    def getValue(self, findV):
        """
        retrieve the value of a variable inside an object
        :param findV:
        :return: the value
        """
        variables = self.__dict__.items()
        for variable, value in variables:
            if variable == findV:
                return value
        return

class storageArray(mesObjets):
    VendorID = ""
    ProductID = ""
    Revision = ""
    ArrayID = ""
    FailoverMode = ""
    NumberPaths = 0
    nbVolume = 0
    nbVolumeClaim = 0


    @staticmethod
    def loadFromXML(XMLString):
        baie = storageArray()
        baie.VendorID = XMLString.find("VendorID").text.strip()
        baie.ProductID = XMLString.find("ProductID").text.strip()
        baie.Revision = XMLString.find("Revision").text.strip()
        baie.ArrayID = XMLString.find("ArrayID").text.strip()
        baie.FailoverMode = XMLString.find("FailoverMode").text.strip()
        baie.NumberPaths = int(XMLString.find("NumberPaths").text)
        todo = XMLString.find("StorageElements")
        baie.nbVolume = int(todo.find("NumberSEs").text)
        baie.nbVolumeClaim = int(todo.find("NumberClaimedSEs").text)

        return baie


class vplex(mesObjets):
    clusterTLA = ""
    clusterId = ""
    clusterNumber = 0
    directorCount = 0
    engineCount = 0
    operationalStatus = ""
    healthState = ""
    list_storageA = []


    @staticmethod
    def loadFromXML(XMLString):
        MyVPlex = vplex()
        ClusterAttributes = XMLString.find("ClusterAttributes")
        MyVPlex.clusterTLA = ClusterAttributes.find("clusterTLA").text.strip()
        MyVPlex.clusterId = ClusterAttributes.find("cluster-id").text.strip()
        MyVPlex.clusterNumber = int(ClusterAttributes.find("cluster-number").text)
        MyVPlex.directorCount = int(ClusterAttributes.find("director-count").text)
        MyVPlex.engineCount = int(MyVPlex.directorCount/2)
        MyVPlex.operationalStatus = ClusterAttributes.find("operational-status").text.strip()
        MyVPlex.healthState=ClusterAttributes.find("health-state").text.strip()

        MyVPlex.list_storageA=[]
        for elt in XMLString.findall("Storage/ArrayList/Array"):
            MyVPlex.list_storageA.append(storageArray.loadFromXML(elt))
            #print(storageArray.loadFromXML(elt).toString())




        return MyVPlex

logger.info("Start VPLEX")
FileToProcess=""

"""
Process the arguments
"""
parser = argparse.ArgumentParser(description="VplexXMLToExcel helps you to translate the configuration of a given VPLEX (local or metro) to an XLS File.\n\nThe program needs a openpyxl in your python installation.")
parser.add_argument('-file_name',help='Allow you to precise the file name (full path or not) of the XML file to parse')
args = parser.parse_args()

"""
Process the symapi_db parameter
"""
if args.file_name is not None:
    print("File Name : "+args.file_name)
    FileToProcess=args.file_name
else:
    logger.error("Missing pararemeters File_name leaving VplexXMLToExcel")
    print("Missing pararemeters File_name leaving VplexXMLToExcel")
    exit(1)

XMLStr=open(FileToProcess, 'r').read()
# clean \n
XMLStr=XMLStr.replace('\n', '')

#
# Make it XML
tableauXML = ET.fromstring(XMLStr)


#
# Is this a Metro ?
TypeVPLEX = tableauXML.find("productType").text.strip()

print(TypeVPLEX)


for nodeXML in tableauXML.findall("ClusterList/Cluster"):
    MyVPLEX = vplex.loadFromXML(nodeXML)
    print(MyVPLEX.toString())




logger.info("End VPLEX")


