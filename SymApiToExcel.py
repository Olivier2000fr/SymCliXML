"""
 SymApiToExcel

 Transforme la symapi d'une baie en fichier XML

 Author : Olivier Guyot

 Developped during off hours (vacations)
 Code is licences under GNU GPL v3
"""

import logging
import copy
import logging.config
import subprocess
import xml.etree.ElementTree as ET
import openpyxl
from openpyxl import Workbook
from shutil import copyfile

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

"""
Initialize constants for the programs.
Mainly ou will find here SymCli commands

-out xml specified output in XML that will be parsed
variable %% %% is peace of text that will be later replace byt the parameter value. (mainly sid's)

"""
SymcfgList = 'symcfg list -v -output xml'
SymcfgEfficiency = 'symcfg -sid %%sid%% -srp -efficiency list -output xml'
SymcfgDemand = 'symcfg -sid %%sid%% list  -demand -v -tb -out xml'
SymCfgListTdev = 'symcfg -sid %%sid%% list -tdev -out xml'
SymDiskList = 'symdisk list -sid %%sid%% -out xml'
SymSGList = 'symsg list -v -sid %%sid%% -out xml'
SymDevShow = 'symdev show -sid %%sid%%  %%device%% -out xml '
SymDevList = 'symdev list -sid %%sid%%  -v -out xml '
SymCfgListMemory = 'symcfg -sid %%sid%%  list -memory -out xml'
SymCfgListFa = 'symcfg -sid %%sid%%  list -fa all -v -out xml'
Supported_Platform = ['VMAX250F', 'VMAX950F', 'VMAX450F', 'VMAX850F', 'PowerMax_8000', 'PowerMax_2000']


class mesObjets:
    """
    meObjets : mother class for all custom objects.

    you will find static methoods for code refactoring (why do it 10 times when you can write once and call many (find / findall)

    You will also find toString methods that transform an object to string listing all attributes to string.
    if an attrivute is a list, then data are not printed.


    """

    def toString(self):
        result = ""
        variables = self.__dict__.items()
        for variable, value in variables:
            if (str(variable).startswith("list_")):
                result = result + variable + " ====> LIST " + str(len(list(value))) + " Elements \n"
            else:
                result = result + variable + "  ====>  " + str(value) + "\n"
        return result

    def getValue(self, findV):
        variables = self.__dict__.items()
        for variable, value in variables:
            if (variable == findV):
                return value
        return

    @staticmethod
    def runFindall(toRun, toSearch) -> list:
        logger.info("RunALL : " + toRun)
        logger.info("FilterAll : " + toSearch)
        liste = subprocess.check_output(toRun, shell=True)
        tableauXML = ET.fromstring(liste)
        return tableauXML.findall(toSearch)

    @staticmethod
    def runFind(toRun, toSearch) -> list:
        logger.info("Run : " + toRun)
        logger.info("Filter : " + toSearch)
        liste = subprocess.check_output(toRun, shell=True)
        tableauXML = ET.fromstring(liste)
        return tableauXML.find(toSearch)

    @staticmethod
    def ifNAtoInt(Value):
        result = 0
        if (Value == "N/A"):
            result = -1
        else:
            result = int(Value)

        return result

class frontEndPorts(mesObjets):
    """
    class for the FE ports (Frontend FC) FA Emulation

    Aim it to list all available Frontend Ports

    No special methods except :
    loadfromXML and loadfrom command (symdisk).
    loadfromXML do the mapping from XML to Object

    """
    dir_name = ""
    port = 0
    port_wwn = ""
    port_status = ""
    negotiated_speed = 0
    maximum_speed = 0

    @staticmethod
    def loadSymmetrixFromXML(feXML,director_name):
        newFE = frontEndPorts()
        newFE.dir_name = director_name
        port_type=feXML.find("Port_Info")
        newFE.port = int(port_type.find("port").text)
        newFE.negotiated_speed = mesObjets.ifNAtoInt(port_type.find("negotiated_speed").text)
        newFE.maximum_speed = mesObjets.ifNAtoInt(port_type.find("maximum_speed").text)
        newFE.port_wwn = port_type.find("port_wwn").text
        newFE.port_status = port_type.find("port_status").text

        return newFE

    @staticmethod
    def loadFromCommand(sid) -> list:
        toRun = SymCfgListFa.replace('%%sid%%', sid)
        listFEPorts = []
        for director in mesObjets.runFindall(toRun, 'Symmetrix/Director'):
            dir_info = director.find("Dir_Info")
            dir_name = dir_info.find("symbolic").text
            for fe in director.findall("Port"):
                listFEPorts.append(frontEndPorts.loadSymmetrixFromXML(fe,dir_name))
        return listFEPorts



class disk(mesObjets):
    """
    class for the physical spindles

    Aim it to list all disks in the symmetrix, size, vendor and revision.

    No special methods except :
    loadfromXML and loadfrom command (symdisk).
    loadfromXML do the mapping from XML to Object

    """


    ident = ""
    da_number = ""
    disk_group = 0
    disk_group_name = ""
    disk_location = ""
    technology = ""
    vendor = ""
    revision = ""
    rated_gigabytes = ""

    @staticmethod
    def loadSymmetrixFromXML(Diskette):
        newDisk = disk()
        Disk = Diskette.find('Disk_Info')
        newDisk.ident = Disk.find("ident").text
        newDisk.da_number = Disk.find("da_number").text
        newDisk.disk_group = int(Disk.find("disk_group").text)
        newDisk.disk_group_name = Disk.find("disk_group_name").text
        newDisk.disk_location = Disk.find("disk_location").text
        newDisk.technology = Disk.find("technology").text
        newDisk.vendor = Disk.find("vendor").text
        newDisk.revision = Disk.find("revision").text
        newDisk.rated_gigabytes = Disk.find("rated_gigabytes").text

        return newDisk

    @staticmethod
    def loadFromCommand(sid) -> list:
        toRun = SymDiskList.replace('%%sid%%', sid)
        listDisks = []
        for galette in mesObjets.runFindall(toRun, 'Symmetrix/Disk'):
            listDisks.append(disk.loadSymmetrixFromXML(galette))
        return listDisks



class storageGroup(mesObjets):
    """
    class for the storageGroup (list of devices altogether
    A storage groups holds the compression flag (to be or not compressed

    This object has to be intialized after Tdevs, becase StorageGroups owns Tdevices (list).

    No special methods except :
    loadfromXML and loadfrom command (symsg).
    loadfromXML do the mapping from XML to Object

    """

    name = ""
    emulation = ""
    Masking_views = ""
    SLO_name = ""
    Compression = ""
    vp_saved_percent = ""
    compression_ratio = ""
    Num_of_GKS = 0
    HostIOLimit_status = ""
    HostIOLimit_max_mb_sec = ""
    HostIOLimit_max_io_sec = ""
    list_devices = []
    nbVolumes = 0
    size_presented_in_gb = 0
    size_allocated_in_gb = 0
    volumeList = ""

    @staticmethod
    def loadSymmetrixFromXML(sgsXML, paramlist_devices):
        sg = storageGroup()
        sginfo = sgsXML.find("SG_Info")
        sg.name = sginfo.find("name").text
        sg.emulation = sginfo.find("emulation").text
        sg.Masking_views = sginfo.find("Masking_views").text
        sg.SLO_name = sginfo.find("SLO_name").text
        sg.Compression = sginfo.find("Compression").text
        sg.vp_saved_percent = sginfo.find("vp_saved_percent").text
        sg.compression_ratio = sginfo.find("compression_ratio").text
        sg.Num_of_GKS = int(sginfo.find("Num_of_GKS").text)
        sg.HostIOLimit_status = sginfo.find("HostIOLimit_status").text
        sg.HostIOLimit_max_mb_sec = sginfo.find("HostIOLimit_max_mb_sec").text
        sg.HostIOLimit_max_io_sec = sginfo.find("HostIOLimit_max_io_sec").text

        dev_lists = sgsXML.find("DEVS_List")
        if dev_lists is None:
            #
            # Storage group has no volume
            #
            sg.nbVolumes = 0
            sg.size_presented_in_gb = 0
            sg.size_allocated_in_gb = 0
            sg.volumeList = ""
        else:
            #
            # Storage group has volumes
            #
            sg.nbVolumes = 0
            sg.size_presented_in_gb = 0
            sg.size_allocated_in_gb = 0
            sg.volumeList = ""
            #
            # Fetch all volumes from the device list
            #
            for device in dev_lists.findall("Device"):
                sg.nbVolumes = sg.nbVolumes + 1
                configuration = device.find("configuration").text
                volID = device.find("dev_name").text
                sg.volumeList = sg.volumeList + volID + ","
                for Tdev in paramlist_devices:
                    if (Tdev.dev_name == volID):
                        sg.list_devices.append(Tdev)
                        sg.size_presented_in_gb = sg.size_presented_in_gb + Tdev.total_tracks_gb
                        sg.size_allocated_in_gb = sg.size_allocated_in_gb + Tdev.alloc_tracks_gb
                        Tdev.configuration = configuration

        return sg

    @staticmethod
    def loadFromCommand(sid, paramlist_devices) -> list:
        toRun = SymSGList.replace('%%sid%%', sid)
        listSgs = []
        for sg in mesObjets.runFindall(toRun, 'SG'):
            listSgs.append(storageGroup.loadSymmetrixFromXML(sg, paramlist_devices))

        return listSgs



class tdev(mesObjets):
    """
        class for the Thin Devices (tdev)


        No special methods except :
        loadfromXML and loadfrom command (symdev and symcfg).
        loadfromXML do the mapping from XML to Object

        """
    dev_name = ""
    dev_emul = ""
    total_tracks_gb = 0
    alloc_tracks_gb = 0
    compression_ratio = ""
    tdev_status = ""
    configuration = ""
    emulation = ""
    encapsulated = ""
    encapsulated_wwn = ""
    encapsulated_array_id = ""
    encapsulated_device_name = ""
    status = ""
    snapvx_source = ""
    snapvx_target = ""
    wwn = ""
    ports = ""
    pair_state = ""
    suspend_state = ""
    consistency_state = ""
    paired_with_concurrent = ""
    paired_with_cascaded = ""
    remote_dev_name = ""
    remote_symid = ""
    remote_wwn = ""
    remote_state = ""
    rdf_mode = ""

    @staticmethod
    def findDetails(ID, listeDeviceDetails):
        for device in listeDeviceDetails:
            devinfo = device.find("Dev_Info")
            dev_name = devinfo.find("dev_name").text
            if dev_name == ID:
                return device
        logger.error("Device not found : " + ID)
        return ""

    @staticmethod
    def loadSymmetrixFromXML(device, listeDeviceDetails):
        newTdev = tdev()
        newTdev.dev_name = device.find("dev_name").text
        newTdev.dev_emul = device.find("dev_emul").text
        newTdev.total_tracks_gb = float(device.find("total_tracks_gb").text)
        newTdev.alloc_tracks_gb = float(device.find("alloc_tracks_gb").text)
        newTdev.compression_ratio = device.find("compression_ratio").text
        newTdev.tdev_status = device.find("tdev_status").text

        #
        # Work on addition info
        #
        details = tdev.findDetails(newTdev.dev_name, listeDeviceDetails)

        # Device not Found
        if details == "":
            logger.debug("Skip -- Device has no details")
            return newTdev

        Dev_Info = details.find("Dev_Info")

        newTdev.encapsulated = Dev_Info.find("encapsulated").text
        newTdev.encapsulated_wwn = Dev_Info.find("encapsulated_wwn").text
        newTdev.encapsulated_array_id = Dev_Info.find("encapsulated_array_id").text
        newTdev.encapsulated_device_name = Dev_Info.find("encapsulated_device_name").text
        newTdev.status = Dev_Info.find("status").text
        newTdev.snapvx_source = Dev_Info.find("snapvx_source").text
        newTdev.snapvx_target = Dev_Info.find("snapvx_target").text

        Dev_Info = details.find("Device_External_Identity")
        newTdev.wwn = Dev_Info.find("wwn").text
        newTdev.ports = ""
        fe = Dev_Info.find("Front_End")
        if fe is not None:
            for port in fe.findall("Port"):
                newTdev.ports = newTdev.ports + port.find("director").text + "-" + port.find("port").text + ","

        rdf = details.find("RDF")
        if rdf is not None:
            rdf_info = rdf.find("RDF_Info")
            newTdev.pair_state = rdf_info.find("pair_state").text
            newTdev.suspend_state = rdf_info.find("suspend_state").text
            newTdev.consistency_state = rdf_info.find("consistency_state").text
            newTdev.paired_with_concurrent = rdf_info.find("paired_with_concurrent").text
            newTdev.paired_with_cascaded = rdf_info.find("paired_with_cascaded").text

            rdf_mode = rdf.find("Mode")
            newTdev.rdf_mode = rdf_mode.find("mode").text

            remote = rdf.find("Remote")
            newTdev.remote_dev_name = remote.find("dev_name").text
            newTdev.remote_symid = remote.find("remote_symid").text
            newTdev.remote_wwn = remote.find("wwn").text
            newTdev.remote_state = remote.find("state").text

        return newTdev

    @staticmethod
    def loadFromCommand(sid) -> list:
        toRun = SymDevList.replace('%%sid%%', sid)
        listeDetailsDevicesXML = mesObjets.runFindall(toRun, 'Symmetrix/Device')

        logger.info("NB elt : " + str(len(listeDetailsDevicesXML)))

        toRun = SymCfgListTdev.replace('%%sid%%', sid)
        listDdevices = []
        listeDevicesXML = mesObjets.runFindall(toRun, 'Symmetrix/ThinDevs/Device')
        logger.info("NB elt : " + str(len(listeDevicesXML)))

        for device in listeDevicesXML:
            listDdevices.append(tdev.loadSymmetrixFromXML(device, listeDetailsDevicesXML))

        return listDdevices


#
# Internal classes
# to move later
#
class symmetrix(mesObjets):
    symid = ""
    attachment = ""
    product_model = ""
    disks = 0
    hot_spares = 0
    patch_level = ""
    raid_level = ""
    srp_name = ""
    vp_efficiency = ""
    snapshot_efficiency = ""
    data_reduction = ""
    drr_enabled_pct = 0
    SRP_efficiency = ""
    effective_used_cap_percent = 0
    usable_capacity_tb = 0.0
    used_capacity_tb = 0.0
    free_capacity_tb = 0.0
    subscribed_capacity_tb = 0.0
    user_used_capacity_tb = 0.0
    system_used_capacity_tb = 0.0
    temp_used_capacity_tb = 0.0
    array_meta_data_used_percent = 0
    repl_meta_data_used_percent = 0
    fe_meta_data_used_percent = 0
    be_meta_data_used_percent = 0
    snapshot_capacity_tb = 0.0
    snapshot_cap_nonshared_tb = 0.0
    snapshot_cap_shared_tb = 0.0
    snapshot_cap_modified_percent = 0
    list_disks = []
    list_devices = []
    list_sgs = []
    list_fes = []
    nb_engine = 0
    nb_cache_raw_tb = 0

    @staticmethod
    def loadSymmetrixFromXML(symm):
        newSymmtrix = symmetrix()

        #
        # go thru SymmInfo for generic Info.
        #
        symmInfo = symm.find('Symm_Info')

        newSymmtrix.symid = symmInfo.find("symid").text
        newSymmtrix.attachment = symmInfo.find("attachment").text
        newSymmtrix.product_model = symmInfo.find("product_model").text
        newSymmtrix.disks = int(symmInfo.find("disks").text)
        newSymmtrix.hot_spares = int(symmInfo.find("hot_spares").text)

        #
        # go thru Enginuity for code level
        #
        symmEnginuity = symm.find('Enginuity')

        newSymmtrix.patch_level = symmEnginuity.find("patch_level").text

        if newSymmtrix.product_model not in Supported_Platform:
            logger.error("Plateform not supported : " + newSymmtrix.product_model)
            return newSymmtrix

        #
        # go thru Flags for Raid Level
        symmFlags = symm.find('Flags')

        if symmFlags.find("raid_5").text == "N/A":
            newSymmtrix.raid_level = symmFlags.find("raid_6").text
        else:
            newSymmtrix.raid_level = symmFlags.find("raid_5").text

        #
        # Get efficiency data
        #
        toRun = SymcfgEfficiency.replace('%%sid%%', newSymmtrix.symid)
        srp = mesObjets.runFind(toRun, "Symmetrix/SRP/SRP_Info")
        newSymmtrix.srp_name = srp.find("name").text

        vp_eff = srp.find('vp_efficiency')
        newSymmtrix.vp_efficiency = vp_eff.find("overall_ratio").text

        snap_eff = srp.find('snapshot_efficiency')
        newSymmtrix.snapshot_efficiency = snap_eff.find("overall_ratio").text

        drr_eff = srp.find('data_reduction')
        newSymmtrix.data_reduction = drr_eff.find("ratio").text
        newSymmtrix.drr_enabled_pct = int(drr_eff.find("enabled_percent").text)

        ovr_eff = srp.find('SRP_efficiency')
        newSymmtrix.SRP_efficiency = ovr_eff.find("overall_ratio").text

        #
        # Get The demand info
        #
        toRun = SymcfgDemand.replace('%%sid%%', newSymmtrix.symid)
        srp = mesObjets.runFind(toRun, "Symmetrix")

        newSymmtrix.effective_used_cap_percent = int(srp.find("effective_used_cap_percent").text)
        newSymmtrix.usable_capacity_tb = float(srp.find("usable_capacity_tb").text)
        newSymmtrix.user_used_capacity_tb = float(srp.find("user_used_capacity_tb").text)
        newSymmtrix.used_capacity_tb = float(srp.find("used_capacity_tb").text)
        newSymmtrix.free_capacity_tb = newSymmtrix.usable_capacity_tb - newSymmtrix.used_capacity_tb
        newSymmtrix.subscribed_capacity_tb = float(srp.find("subscribed_capacity_tb").text)
        newSymmtrix.system_used_capacity_tb = float(srp.find("system_used_capacity_tb").text)
        newSymmtrix.temp_used_capacity_tb = float(srp.find("temp_used_capacity_tb").text)
        newSymmtrix.array_meta_data_used_percent = int(srp.find("array_meta_data_used_percent").text)
        newSymmtrix.repl_meta_data_used_percent = int(srp.find("repl_meta_data_used_percent").text)
        newSymmtrix.fe_meta_data_used_percent = mesObjets.ifNAtoInt(srp.find("fe_meta_data_used_percent").text)
        newSymmtrix.be_meta_data_used_percent = mesObjets.ifNAtoInt(srp.find("be_meta_data_used_percent").text)
        newSymmtrix.snapshot_capacity_tb = float(srp.find("snapshot_capacity_tb").text)
        newSymmtrix.snapshot_cap_nonshared_tb = float(srp.find("snapshot_cap_nonshared_tb").text)
        newSymmtrix.snapshot_cap_shared_tb = float(srp.find("snapshot_cap_shared_tb").text)
        newSymmtrix.snapshot_cap_modified_percent = int(srp.find("snapshot_cap_modified_percent").text)

        #
        # Load Disks
        #
        newSymmtrix.list_disks = disk.loadFromCommand(newSymmtrix.symid)

        #
        # Load devices
        #
        newSymmtrix.list_devices = tdev.loadFromCommand(newSymmtrix.symid)

        #
        # Load SG's
        #
        newSymmtrix.list_sgs = storageGroup.loadFromCommand(newSymmtrix.symid, newSymmtrix.list_devices)

        #
        # get Engines and cache
        #
        toRun = SymCfgListMemory.replace('%%sid%%', newSymmtrix.symid)
        memory = mesObjets.runFind(toRun, "Symmetrix/Symm_Info")

        newSymmtrix.nb_engine = int(memory.find("total_mem_boards").text)//2
        newSymmtrix.nb_cache_raw_tb = ((int(memory.find("total_cap_in_mb").text)//1000000)+1)*2

        #
        # grab the FA pors (frontend FC emulation)
        #
        newSymmtrix.list_fes=frontEndPorts.loadFromCommand(newSymmtrix.symid)

        return newSymmtrix


#
# Main
#


def objectToXLS(Cell,chaine : str ,monObject : mesObjets):
    if str(cell.value).startswith(chaine):
        #
        # Manage Sym Data
        #
        Attributes = str(cell.value).replace(chaine, "")
        cell.value = monObject.getValue(Attributes)

def ListToXLS(feuille, Cell,chaine : str ,maListe : [mesObjets]):
    if str(cell.value).startswith(chaine):
        #
        # Manage Liste of disks
        #
        Attributes = str(cell.value).replace(chaine, "")
        row_x = cell.row
        column_y = cell.column
        for elt in maListe:
            #
            # on écrit de bas en haut.
            #
            celltoupd = feuille.cell(row=row_x, column=column_y)
            celltoupd.value = elt.getValue(Attributes)
            row_x = row_x + 1




# 'application' code
logger.info("Start")

for symm in mesObjets.runFindall(SymcfgList, 'Symmetrix'):
    MySymm = symmetrix.loadSymmetrixFromXML(symm)

    #

    #
    # Copy XLS
    #
    copyfile("reference.xlsx", MySymm.symid + '.xlsx')

    #
    # open the file and start to work
    #

    classeur = openpyxl.load_workbook(MySymm.symid + '.xlsx')

    for feuille_name in classeur.sheetnames:
        #
        # On parcourt les pages
        #
        feuille = classeur[feuille_name]
        for ligne in feuille.iter_rows():
            for cell in ligne:
                #
                # Analyse et travaille ici
                #
                objectToXLS(cell,"%%sym.",MySymm)
                ListToXLS(feuille,cell,"%%list.disks.",MySymm.list_disks)
                ListToXLS(feuille, cell, "%%list.tdevs.", MySymm.list_devices)
                ListToXLS(feuille, cell, "%%list.sgs.", MySymm.list_sgs)
                ListToXLS(feuille, cell, "%%list.fes.", MySymm.list_fes)

    #
    # Save File
    #
    classeur.save(MySymm.symid + '.xlsx')

    print(MySymm.toString())
#    for child in symm:
#        print(child.tag, child.attrib)


logger.info("End")
