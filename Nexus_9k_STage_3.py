from openpyxl import load_workbook
import ciscoconfparse as c
import re
import ipaddr
import itertools


#############################################
################# VARIABLES #################
#############################################

SWITCH = 'NAOSW133'
#INFRA_CH_GRP_LIST = [1,133]
SHEET = SWITCH
BASE_DIR = '/Users/aspera/Documents/Clienti/VF-2017/NMP/NA1C/' + SWITCH + '/Stage_3/'


INPUT_XLS = BASE_DIR + SWITCH + '_OUT_DB_OPT.xlsx'
#OUTPUT_XLS = BASE_DIR + SWITCH + '_OUT_DB.xlsx'
OSW_CFG_TXT = BASE_DIR + SWITCH + '.txt'

#############################################
################ FUNCTIONS ##################
#############################################

def atoi(text):
    return int(text) if text.isdigit() else text

def natural_keys(text):
    '''
    alist.sort(key=natural_keys) sorts in human order
    http://nedbatchelder.com/blog/200712/human_sorting.html
    (See Toothy's implementation in the comments)
    '''
    return [ atoi(c) for c in re.split('(\d+)', text) ]

def get_col(ws,col):
    ''' Take a worksheet, return column "col" as list '''
    return [str(ws.cell(row = r, column = col).value) for r in range(2,ws.max_row+1)]

def get_if_from_xls():
    ''' Return column col as list '''
    wb_r = load_workbook(INPUT_XLS)
    ws_r = wb_r.get_sheet_by_name(SHEET)
    SRC_IF_ROW = 1
    a = get_col(ws_r,SRC_IF_ROW)
    a.sort(key=natural_keys)
    return a

def get_if_from_cfg():
    parse = c.CiscoConfParse(OSW_CFG_TXT)
    intf_obj_list = parse.find_objects(r'^interface .*Ethernet')
    
    a = [obj.text for obj in intf_obj_list]
    a.sort(key=natural_keys)
    return a
    
    
def get_SVI_if_from_cfg():
    parse = c.CiscoConfParse(OSW_CFG_TXT)
    intf_obj_list = parse.find_objects(r'^interface Vlan')
    
    return [obj.text for obj in intf_obj_list]

def get_VLAN_from_cfg():
    parse = c.CiscoConfParse(OSW_CFG_TXT)
    vlan_obj_list = parse.find_objects(r'^vlan \d+')
    
    return [obj.text.split(' ')[1] for obj in vlan_obj_list]

def get_VLAN_from_xls():
    
    a = set()
    wb_r = load_workbook(INPUT_XLS)
    ws_r = wb_r.get_sheet_by_name(SHEET)

    lst = get_col(ws_r,4)
    for elem in lst:
        if ',' in elem:
            b = elem.split(',')
            for elem2 in b:
                a.add(elem2)
        else:
            a.add(elem)
    
    lst2 = list(a)
    lst2.sort(key=natural_keys)
    return lst2
    
def get_LIST_not_to_be_migrated(ifxls,ifcfg):
    a = set(ifxls)
    b = set(ifcfg)
    c = b-a
    d=list(c)
    if len(d) > 0:
        d.sort(key=natural_keys)
        return d 
    else:
        return []
    
def write_normalized_OSW_cfg(if_ntbm, vlan_ntbm):
    pass


#############################################
################### MAIN ####################
#############################################


if_xls = get_if_from_xls()
if_cfg = get_if_from_cfg()

print "if_xls = ", if_xls
print "if_cfg = ", if_cfg

if_not_to_be_migrated = get_LIST_not_to_be_migrated(if_xls, if_cfg)
print "if_not_to_be_migrated = " , if_not_to_be_migrated

vlan_xls = get_VLAN_from_xls()
vlan_cfg = get_VLAN_from_cfg()

vlan_not_to_be_migrated = get_LIST_not_to_be_migrated(vlan_xls, vlan_cfg)
print "vlan_not_to_be_migrated = " , vlan_not_to_be_migrated

write_normalized_OSW_cfg(if_not_to_be_migrated, vlan_not_to_be_migrated)
