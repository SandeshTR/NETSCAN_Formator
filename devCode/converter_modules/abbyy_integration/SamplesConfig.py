import sys
import os

current_dir = os.path.dirname(os.path.abspath(__file__))
#print(os.path.join(current_dir,'SWAD12410007214764846289.ABBYY.ActivationToken'))

## Return Customer Project ID for FRE
def GetCustomerProjectId():
    return "FFvrEyp5Gz8sXSwP98N9"

## Return path to the license file
def GetLicensePath():
    return os.path.join(current_dir,'SWAD12410007214764846289.ABBYY.ActivationToken')

## Return license password
def GetLicensePassword():
    return "/80HjebrjO2bzpJUiJ/DwQ=="

## Return full path to Samples directory
def GetSamplesFolder():
    return r"C:\File\NETSCAN"
