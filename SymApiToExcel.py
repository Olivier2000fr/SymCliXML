#
# SymApiToExcel
#
# Transforme la symapi d'une baie en fichier XML
#
# Author : Olivier Guyot
#
#

import logging
import copy
import logging.config
import subprocess
import xml.etree.ElementTree as ET
import openpyxl
from openpyxl import Workbook
from shutil import copyfile

#
# Initialize logging
#
logging.config.fileConfig('SymApiToExcel.logging')
# create logger
logger = logging.getLogger('root')

#
# Constantes
#
SymcfgList = 'symcfg list -v -output xml'
SymcfgEfficiency = 'symcfg -sid %%sid%% -srp -efficiency list -output xml'
SymcfgDemand = 'symcfg -sid %%sid%% list  -demand -v -tb -out xml'
SymCfgListTdev = 'symcfg -sid %%sid%% list -tdev -out xml'
SymDiskList = 'symdisk list -sid %%sid%% -out xml'
SymSGList = 'symsg list -v -sid %%sid%% -out xml'
SymDevShow = 'symdev show -sid %%sid%%  %%device%% -out xml '
SymDevList = 'symdev list -sid %%sid%%  -v -out xml '
Supported_Platform = ['VMAX250F' , 'VMAX950F', 'VMAX450F', 'VMAX850F', 'PowerMax_8000', 'PowerMax_2000']


class mesObjets:
    def toString(self):
        result = ""
        variables = self.__dict__.items()
        for variable, value in variables:
            if (str(variable).startswith("list_")):
                result = result + variable + " ====> LIST \n"
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
        if (Value=="N/A"):
            result=-1
        else:
            result=int(Value)

        return result


#
# Class Disk
#
class disk(mesObjets):
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


#
# Class storageGroup
#
class storageGroup(mesObjets):
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

        dev_lists=sgsXML.find("DEVS_List")
        if dev_lists is None:
            sg.nbVolumes = 0
            sg.size_presented_in_gb = 0
            sg.size_allocated_in_gb = 0
            sg.volumeList=""
        else:
            sg.nbVolumes = 0
            sg.size_presented_in_gb = 0
            sg.size_allocated_in_gb = 0
            sg.volumeList=""
            for device in dev_lists.findall("Device"):
                sg.nbVolumes=sg.nbVolumes+1
                configuration=device.find("configuration").text
                volID=device.find("dev_name").text
                sg.volumeList=sg.volumeList+volID+","
                for Tdev in paramlist_devices:
                    if (Tdev.dev_name==volID):
                        sg.list_devices.append(Tdev)
                        sg.size_presented_in_gb=sg.size_presented_in_gb+Tdev.total_tracks_gb
                        sg.size_allocated_in_gb = sg.size_allocated_in_gb+Tdev.alloc_tracks_gb
                        Tdev.configuration=configuration



        return sg


    @staticmethod
    def loadFromCommand(sid, paramlist_devices) -> list:
        toRun = SymSGList.replace('%%sid%%', sid)
        listSgs = []
        for sg in mesObjets.runFindall(toRun, 'SG'):
            listSgs.append(storageGroup.loadSymmetrixFromXML(sg, paramlist_devices))

        return listSgs



#
# Class Tdev
#
class tdev(mesObjets):
    dev_name = ""
    dev_emul = ""
    total_tracks_gb = 0
    alloc_tracks_gb = 0
    compression_ratio = ""
    tdev_status = ""
    configuration=""
    emulation=""
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
    def findDetails(ID,listeDeviceDetails):
        for device in listeDeviceDetails:
            devinfo = device.find("Dev_Info")
            dev_name = devinfo.find("dev_name").text
            if dev_name==ID:
                return device
        logger.error("Device not found : "+ID)
        return ""

    @staticmethod
    def loadSymmetrixFromXML(device,listeDeviceDetails):
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
        details=tdev.findDetails(newTdev.dev_name,listeDeviceDetails)

        # Device not Found
        if details == "":
            logger.debug("Skip -- Device has no details")
            return newTdev

        Dev_Info=details.find("Dev_Info")

        newTdev.encapsulated=Dev_Info.find("encapsulated").text
        newTdev.encapsulated_wwn=Dev_Info.find("encapsulated_wwn").text
        newTdev.encapsulated_array_id=Dev_Info.find("encapsulated_array_id").text
        newTdev.encapsulated_device_name=Dev_Info.find("encapsulated_device_name").text
        newTdev.status=Dev_Info.find("status").text
        newTdev.snapvx_source=Dev_Info.find("snapvx_source").text
        newTdev.snapvx_target=Dev_Info.find("snapvx_target").text

        Dev_Info = details.find("Device_External_Identity")
        newTdev.wwn=Dev_Info.find("wwn").text
        newTdev.ports=""
        fe = Dev_Info.find("Front_End")
        if fe is not None:
            for port in fe.findall("Port"):
                newTdev.ports=newTdev.ports+port.find("director").text+"-"+port.find("port").text+","

        rdf = details.find("RDF")
        if rdf is not None:
            rdf_info=rdf.find("RDF_Info")
            newTdev.pair_state = rdf_info.find("pair_state").text
            newTdev.suspend_state = rdf_info.find("suspend_state").text
            newTdev.consistency_state = rdf_info.find("consistency_state").text
            newTdev.paired_with_concurrent = rdf_info.find("paired_with_concurrent").text
            newTdev.paired_with_cascaded = rdf_info.find("paired_with_cascaded").text

            rdf_mode=rdf.find("Mode")
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

        logger.info("NB elt : "+str(len(listeDetailsDevicesXML)))

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
            logger.error("Plateform not supported : "+newSymmtrix.product_model)
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
        srp = mesObjets.runFind(toRun,"Symmetrix/SRP/SRP_Info")
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
        newSymmtrix.list_disks=disk.loadFromCommand(newSymmtrix.symid)

        #
        # Load devices
        #
        newSymmtrix.list_devices=tdev.loadFromCommand(newSymmtrix.symid)

        #
        # Load SG's
        #
        newSymmtrix.list_sgs = storageGroup.loadFromCommand(newSymmtrix.symid,newSymmtrix.list_devices)

        #
        # Construct SG_Devices Report
        #

        return newSymmtrix



#
# Main
#


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

    for feuille_name in classeur.get_sheet_names():
        #
        # On parcourt les pages
        #
        feuille = classeur.get_sheet_by_name(feuille_name)
        for ligne in feuille.iter_rows():
            for cell in ligne:
                #
                # Analyse et travaille ici
                #
                if (str(cell.value).startswith("%%sym.")):
                    #
                    # Manage Sym Data
                    #
                    Attributes = str(cell.value).replace("%%sym.", "")
                    cell.value = MySymm.getValue(Attributes)
                if (str(cell.value).startswith("%%list.disks.")):
                    #
                    # Manage Liste of disks
                    #
                    Attributes = str(cell.value).replace("%%list.disks.", "")
                    row_x = cell.row
                    column_y = cell.column
                    for disque in MySymm.list_disks:
                        #
                        # on écrit de bas en haut.
                        #
                        celltoupd = feuille.cell(row=row_x, column=column_y)
                        celltoupd.value = disque.getValue(Attributes)
                        row_x = row_x + 1
                if (str(cell.value).startswith("%%list.tdevs.")):
                    #
                    # Manage Liste of disks
                    #
                    Attributes = str(cell.value).replace("%%list.tdevs.", "")
                    row_x = cell.row
                    column_y = cell.column
                    for device in MySymm.list_devices:
                        #
                        # on écrit de bas en haut.
                        #
                        celltoupd = feuille.cell(row=row_x, column=column_y)
                        celltoupd.value = device.getValue(Attributes)
                        row_x = row_x + 1
                if (str(cell.value).startswith("%%list.sgs.")):
                    #
                    # Manage Liste of disks
                    #
                    Attributes = str(cell.value).replace("%%list.sgs.", "")
                    row_x = cell.row
                    column_y = cell.column
                    for sg in MySymm.list_sgs:
                        #
                        # on écrit de bas en haut.
                        #
                        celltoupd = feuille.cell(row=row_x, column=column_y)
                        celltoupd.value = sg.getValue(Attributes)
                        row_x = row_x + 1

    #
    # Save File
    #
    classeur.save(MySymm.symid + '.xlsx')

    print(MySymm.toString())
#    for child in symm:
#        print(child.tag, child.attrib)


logger.info("End")
