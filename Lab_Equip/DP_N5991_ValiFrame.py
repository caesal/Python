# -*- coding: utf-8 -*-
"""
Created on Wed Mar 27 14:50:31 2024

@author: CC.Cheng
"""
import platform
import string
import clr
import sys
import os
import time
import pandas as pd

################################################## Constants ##################################################

# the directory where the ValiFrame DLLs are; set to None to use the working directory
# ValiFrameDllDirectory = r'C:\Program Files\BitifEye\ValiFrameK1\PCIe\TestAutomation'
ValiFrameDllDirectory = r'C:\Program Files\BitifEye\ValiFrameK1\DisplayPort\TestAutomation'

# set to False to let the user choose an application (if there is more than one available); otherwise, put app name here
ForceApplication = False

# set to True to ask the user to change application settings before the application is configured
AskUserToChangeApplicationPropertiesBeforeConfig = False

# set to True to ask the user to change application settings after the application is configured
AskUserToChangeApplicationPropertiesAfterConfig = False

# set to True to ask the user to change procedure settings before running a procedure
AskUserToChangeProcedureProperties = False

#############################################################################
# set to True to always show the "Configure DUT" dialog, False to not show it, or None to let the user decide
ShowConfigDialogPreference = True

##########################################################
##########Pick your procedures here#######################
##########################################################
# a list of procedure IDs to run; set to None to let the user decide
##########################################################

# (4010128, Random Jitter Calibration RBR (TP1); 3/15/2024 3:26:13 PM)
# (4010090, ISI Calibration RBR (TP2); 3/15/2024 3:50:24 PM)
# (4010150, Eye Opening Calibration RBR Lane 0 (TP2); 3/15/2024 3:59:37 PM)
# (4010151, Eye Opening Calibration RBR Lane 1 (TP2); 3/15/2024 4:06:58 PM)
# (4010060, ISI Calibration RBR (TP3); 3/15/2024 4:37:03 PM)
# (4010140, Eye Opening Calibration RBR Lane 0 (TP3); 3/15/2024 5:23:36 PM)
# (4010141, Eye Opening Calibration RBR Lane 1 (TP3); 3/18/2024 10:34:55 AM)
# (4010050, Random Jitter Calibration HBR (TP1); 3/18/2024 11:46:14 AM)
# (4010100, ISI Calibration HBR (TP2); 3/18/2024 12:04:20 PM)
# (4010115, Eye Opening Calibration HBR Lane 0 (TP2); 3/18/2024 12:08:34 PM)
# (4010116, Eye Opening Calibration HBR Lane 1 (TP2); 3/18/2024 2:26:10 PM)
# (4010070, ISI Calibration HBR (TP3); 3/18/2024 12:14:49 PM)
# (4010105, Eye Opening Calibration HBR Lane 0 (TP3); 3/18/2024 12:22:44 PM)
# (4010106, Eye Opening Calibration HBR Lane 1 (TP3); 3/18/2024 2:34:11 PM)
# (4010051, Random Jitter Calibration HBR2 (TP1); 3/18/2024 11:48:06 AM)
# (4010055, Fixed Sinusoidal Jitter Calibration HBR2 (TP1); 3/18/2024 11:49:43 AM)
# (4010110, ISI Calibration HBR2 (TP2); 3/18/2024 1:33:01 PM)
# (4010134, Eye Opening Calibration HBR2 Lane 0 (TP2); 3/18/2024 1:47:50 PM)
# (4010135, Eye Opening Calibration HBR2 Lane 1 (TP2); 3/18/2024 2:49:12 PM)
# (4010080, ISI Calibration HBR2 (TP3); 3/18/2024 1:51:42 PM)
# (4010124, Eye Opening Calibration HBR2 Lane 0 (TP3); 3/18/2024 2:18:43 PM)
# (4010125, Eye Opening Calibration HBR2 Lane 1 (TP3); 3/18/2024 3:16:36 PM)
# (4010158, Aggressor Amplitude Calibration RBR; 3/18/2024 3:58:59 PM)
# (4010166, Aggressor Amplitude Calibration HBR; 3/18/2024 3:59:55 PM)
# (4010174, Aggressor Amplitude Calibration HBR2; 3/18/2024 4:01:22 PM)
# (4000480, Jitter Tolerance Test 2 MHz SJ RBR Lane 0 (TP2))
# (4000512, Jitter Tolerance Test 10 MHz SJ RBR Lane 0 (TP2))
# (4000544, Jitter Tolerance Test 20 MHz SJ RBR Lane 0 (TP2))
# (4000504, Jitter Tolerance Test 2 MHz SJ HBR Lane 0 (TP2))
# (4000536, Jitter Tolerance Test 10 MHz SJ HBR Lane 0 (TP2))
# (4000568, Jitter Tolerance Test 20 MHz SJ HBR Lane 0 (TP2))
# (4000600, Jitter Tolerance Test 100 MHz SJ HBR Lane 0 (TP2))
# (4000528, Jitter Tolerance Test 0.5 MHz SJ HBR2 Lane 0 (TP2))
# (4000560, Jitter Tolerance Test 1 MHz SJ HBR2 Lane 0 (TP2))
# (4000592, Jitter Tolerance Test 2 MHz SJ HBR2 Lane 0 (TP2))
# (4000624, Jitter Tolerance Test 10 MHz SJ HBR2 Lane 0 (TP2))
# (4000656, Jitter Tolerance Test 20 MHz SJ HBR2 Lane 0 (TP2))
# (4000688, Jitter Tolerance Test 100 MHz SJ HBR2 Lane 0 (TP2))
# (4000096, Jitter Tolerance Test 2 MHz SJ RBR Lane 0 (TP3))
# (4000320, Jitter Tolerance Test 10 MHz SJ RBR Lane 0 (TP3))
# (4000352, Jitter Tolerance Test 20 MHz SJ RBR Lane 0 (TP3))
# (4000120, Jitter Tolerance Test 2 MHz SJ HBR Lane 0 (TP3))
# (4000344, Jitter Tolerance Test 10 MHz SJ HBR Lane 0 (TP3))
# (4000376, Jitter Tolerance Test 20 MHz SJ HBR Lane 0 (TP3))
# (4000184, Jitter Tolerance Test 100 MHz SJ HBR Lane 0 (TP3))
# (4000728, Jitter Tolerance Test 0.5 MHz SJ HBR2 Lane 0 (TP3))
# (4000628, Jitter Tolerance Test 1 MHz SJ HBR2 Lane 0 (TP3))
# (4000144, Jitter Tolerance Test 2 MHz SJ HBR2 Lane 0 (TP3))
# (4000368, Jitter Tolerance Test 10 MHz SJ HBR2 Lane 0 (TP3))
# (4000400, Jitter Tolerance Test 20 MHz SJ HBR2 Lane 0 (TP3))
# (4000208, Jitter Tolerance Test 100 MHz SJ HBR2 Lane 0 (TP3))
# (4000481, Jitter Tolerance Test 2 MHz SJ RBR Lane 1 (TP2))
# (4000513, Jitter Tolerance Test 10 MHz SJ RBR Lane 1 (TP2))
# (4000545, Jitter Tolerance Test 20 MHz SJ RBR Lane 1 (TP2))
# (4000505, Jitter Tolerance Test 2 MHz SJ HBR Lane 1 (TP2))
# (4000537, Jitter Tolerance Test 10 MHz SJ HBR Lane 1 (TP2))
# (4000569, Jitter Tolerance Test 20 MHz SJ HBR Lane 1 (TP2))
# (4000601, Jitter Tolerance Test 100 MHz SJ HBR Lane 1 (TP2))
# (4000529, Jitter Tolerance Test 0.5 MHz SJ HBR2 Lane 1 (TP2))
# (4000561, Jitter Tolerance Test 1 MHz SJ HBR2 Lane 1 (TP2))
# (4000593, Jitter Tolerance Test 2 MHz SJ HBR2 Lane 1 (TP2))
# (4000625, Jitter Tolerance Test 10 MHz SJ HBR2 Lane 1 (TP2))
# (4000657, Jitter Tolerance Test 20 MHz SJ HBR2 Lane 1 (TP2))
# (4000689, Jitter Tolerance Test 100 MHz SJ HBR2 Lane 1 (TP2))
# (4000097, Jitter Tolerance Test 2 MHz SJ RBR Lane 1 (TP3))
# (4000321, Jitter Tolerance Test 10 MHz SJ RBR Lane 1 (TP3))
# (4000353, Jitter Tolerance Test 20 MHz SJ RBR Lane 1 (TP3))
# (4000121, Jitter Tolerance Test 2 MHz SJ HBR Lane 1 (TP3))
# (4000345, Jitter Tolerance Test 10 MHz SJ HBR Lane 1 (TP3))
# (4000377, Jitter Tolerance Test 20 MHz SJ HBR Lane 1 (TP3))
# (4000185, Jitter Tolerance Test 100 MHz SJ HBR Lane 1 (TP3))
# (4000729, Jitter Tolerance Test 0.5 MHz SJ HBR2 Lane 1 (TP3))
# (4000629, Jitter Tolerance Test 1 MHz SJ HBR2 Lane 1 (TP3))
# (4000145, Jitter Tolerance Test 2 MHz SJ HBR2 Lane 1 (TP3))
# (4000369, Jitter Tolerance Test 10 MHz SJ HBR2 Lane 1 (TP3))
# (4000401, Jitter Tolerance Test 20 MHz SJ HBR2 Lane 1 (TP3))
# (4000209, Jitter Tolerance Test 100 MHz SJ HBR2 Lane 1 (TP3))

##########################################################
ProcedureIdsToAutoExecute = None
# ProcedureIdsToAutoExecute = [4000600] #(4000600, Jitter Tolerance Test 100 MHz SJ HBR Lane 0 (TP2))


# set to True to confirm all dialogs automatically, set to False to ask the user each time
AutoConfirmAllDialogs = True

# set to True to always show the XML result, False to not show it, or None to let the user decide
ShowXmlPreference = True

# set to True if you want to show script events in the console
ScriptLogToConsole = True

# set to True to show ValiFrame log entries in the console
ValiFrameLogToConsole = False

# set to True to also show ValiFrame internal log messages in the console
ValiFrameLogInternalToConsole = False

# the ValiFrame log file (all messages); set to None to disable logging
ValiFrameLogFile = None

# the ValiFrame XML result file (older ones are overwritten); set to None to disable saving
# XML file autosaved address: C:\ProgramData\BitifEye\ValiFrameK1\Tmp
ValiFrameXmlResultFile = True

# set to True to automatically close the script when everything was done; otherwise, wait for some user input
AutoCloseScript = False

############################################ Import .NET Namespaces ###########################################

import clr

if ValiFrameDllDirectory != None:
    sys.path.append(ValiFrameDllDirectory)
clr.AddReference(r'ValiFrameRemote')  #
clr.AddReference(r'VFBase')
from BitifEye.ValiFrame.ValiFrameRemote import *
from BitifEye.ValiFrame.Base import *
from BitifEye.ValiFrame.Logging import *
from System import *


################################################## Functions ##################################################

def ScriptLog(line):
    if ScriptLogToConsole:
        print(line)
    return


def IsIronPython():
    return platform.python_implementation() == 'IronPython'


def StartValiFrame():
    ScriptLog('Creating ValiFrame instance...')
    valiFrame = ValiFrameRemote(ProductGroupE.ValiFrameK1)
    return valiFrame


def UserBoolQuery(trueAnswer):
    userInput = input()
    return (userInput == trueAnswer)


def LogEntryChangedHandler(logEntry):
    message = logEntry.Text
    severity = None
    isInternal = False
    if logEntry.Severity == VFLogSeverityE.Internal:
        severity = 'Internal'
        isInternal = True
    elif logEntry.Severity == VFLogSeverityE.Info:
        severity = 'Info'
    elif logEntry.Severity == VFLogSeverityE.Progress:
        severity = 'Progress'
    elif logEntry.Severity == VFLogSeverityE.Warning:
        severity = 'Warning'
    elif logEntry.Severity == VFLogSeverityE.Critical:
        severity = 'Critical'
    elif logEntry.Severity == VFLogSeverityE.Exception:
        severity = 'Exception'
    else:
        severity = 'Unknown severity'

    if ValiFrameLogToConsole:
        show = True
        if isInternal:
            show = ValiFrameLogInternalToConsole
        if show:
            print('Log (%s): %s' % (severity, message))

    if ValiFrameLogFile != None:
        with open(ValiFrameLogFile, 'a') as file:
            file.write('%s: %s\n' % (severity, message))
    return


def StatusChangedHandler(sender, description):
    print('Status changed: %s' % description)
    return


def ProcedureCompletedHandler(procedureId, xmlResult):
    print('Procedure %d is complete' % procedureId)
    ScriptLog('Saving XML result to file...')
    if ValiFrameXmlResultFile != None:
        with open(ValiFrameXmlResultFile, 'w') as file:
            file.write(xmlResult)
    showXml = ShowXmlPreference
    if showXml == None:
        print('Do you want to see the XML result (y/n)?')
        showXml = UserBoolQuery('y')
    if showXml:
        print(xmlResult)
    return


def DialogPopUpHandler(sender, dialogInformation):
    print('Popup: %s' % dialogInformation.DialogText)
    abort = False
    if not AutoConfirmAllDialogs:
        print('Continue (y/n)?')
        abort = UserBoolQuery('n')
    if abort:
        dialogInformation.Dialog.DialogResult = System.Windows.Forms.DialogResult.Cancel;
    return


def RegisterEventHandlers(valiFrame):
    valiFrame.LogEntryChanged += LogEntryChangedEventHandler(LogEntryChangedHandler)
    valiFrame.StatusChanged += StatusChangedEventHandler(StatusChangedHandler)
    valiFrame.ProcedureCompleted += ProcedureCompletedEventHandler(ProcedureCompletedHandler)
    valiFrame.DialogPopUp += DialogShowEventHandler(DialogPopUpHandler)
    return


def UnregisterEventHandlers(valiFrame):
    try:
        valiFrame.LogEntryChanged -= LogEntryChangedEventHandler(LogEntryChangedHandler)
        valiFrame.StatusChanged -= StatusChangedEventHandler(StatusChangedHandler)
        valiFrame.ProcedureCompleted -= ProcedureCompletedEventHandler(ProcedureCompletedHandler)
        valiFrame.DialogPopUp -= DialogShowEventHandler(DialogPopUpHandler)
    except:
        return  # ignore errors
    return


def SelectApplication(valiFrame):
    if ForceApplication != False:
        return ForceApplication

    ScriptLog('Getting list of available applications...')
    applicationNames = valiFrame.GetApplications()

    if len(applicationNames) < 1:
        raise RuntimeError('No applications found')
    elif len(applicationNames) == 1:
        ScriptLog('Selecting application %s (no others available)' % applicationNames[0])
        return applicationNames[0]
    else:
        while True:
            print('The following %d applications were found:' % len(applicationNames))
            for applicationName in applicationNames:
                print('- "%s"' % applicationName)
            print('Please select one (leave blank to selecte the first one): ')
            selectedApplicationName = input()
            if selectedApplicationName == '':
                return applicationNames[0]
            for applicationName in applicationNames:
                if applicationName == selectedApplicationName:
                    return applicationName
            print('Invalid selection')


def InitApplication(valiFrame, applicationName):
    ScriptLog('Initializing Application...')
    valiFrame.InitApplication(applicationName)
    return


def ConfigureApplication(valiFrame):
    showDialog = ShowConfigDialogPreference
    if showDialog == None:
        print('Do you want to show the GUI dialog to configure the application (y/n)?')
        showDialog = not UserBoolQuery('n')
    if showDialog:
        ScriptLog('Configuring application with GUI dialog...')
        valiFrame.ConfigureProduct()
    else:
        ScriptLog('Configuring application automatically...')
        valiFrame.ConfigureProductNoDialog()
    return


def GetAvailableApplicationProperties(valiFrame):
    ScriptLog('Getting a list of available application properties...')
    if IsIronPython():
        propertiesClr = valiFrame.GetApplicationPropertiesList()
        properties = dict(propertiesClr)
    else:
        propertiesClr = valiFrame.GetApplicationPropertiesList()
        properties = {}
        for prop in propertiesClr:
            print("{} : {}".format(prop.Key, prop.Value))
            #            keyValuePair = prop.split(',')
            #            key = keyValuePair[0][1:].strip()
            #            value = keyValuePair[1][0:len(keyValuePair[1])-1].strip()
            properties[prop.Key] = prop.Value
    return properties


def LetUserChangeApplicationProperties(valiFrame):
    properties = GetAvailableApplicationProperties(valiFrame)
    print('Available application properties:')
    for propertyKey in properties:
        print('- %s: %s' % (propertyKey, properties[propertyKey]))
    print('Do you want to change an application property (y/n)?')
    change = UserBoolQuery('y')
    while change:
        print('Please enter the name of the application property:')
        requestedPropertyName = input()
        okay = False
        for propertyKey in properties:
            if requestedPropertyName == propertyKey:
                print('Please enter the new value:')
                newValue = input()
                valiFrame.SetApplicationProperty(propertyKey, newValue)
                okay = True
        if okay:
            print('Do you want to change another application property (y/n)?')
            change = UserBoolQuery('y')
        else:
            print('Invalid input')
    return


def ChangePropertiesBeforeConfiguration(valiFrame):
    if AskUserToChangeApplicationPropertiesBeforeConfig:
        LetUserChangeApplicationProperties(valiFrame)
    return


def ChangePropertiesAfterConfiguration(valiFrame):
    if AskUserToChangeApplicationPropertiesAfterConfig:
        LetUserChangeApplicationProperties(valiFrame)
    return


def GetAvailableProcedures(valiFrame):
    ScriptLog('\nGetting list of available procedures...')
    if IsIronPython():
        procedureIds, procedureNames = valiFrame.GetProcedures()
        print("HERE")
    else:
        print("THERE")
        df = pd.read_excel('ProcedureIDs.xlsx', sheet_name=0)
        procedureIds = list(df['procedureIds'])
        procedureNames = list(df['procedureNames'])
        # dummyIntArray = (-1, 1)  # a dummy array of signed ints
        # dummyStrArray = ('dummy')  # a dummy array of strings
        # dummyOut, procedureIds, procedureNames = valiFrame.GetProcedures(dummyIntArray, dummyStrArray)
    return procedureIds, procedureNames


def GetAvailableProcedureProperties(valiFrame, procedureId):
    ScriptLog('\nGetting a list of available procedure properties for procedure %d...' % procedureId)
    properties = valiFrame.GetProcedureProperties(procedureId)

    # convert to generic hash
    flatProcedureProperties = {}
    for obj in properties:
        propertyName = str(obj.Name)
        propertyValue = str(obj)
        flatProcedureProperties[propertyName] = propertyValue

    return flatProcedureProperties


def GetAvailableRelatedProperties(valiFrame, procedureId):
    ScriptLog('\nGetting a list of available RELATED procedure properties for procedure %d...' % procedureId)
    if IsIronPython():
        propertiesClr = valiFrame.GetProcedureRelatedProperties(procedureId)
        properties = dict(propertiesClr)
    else:
        propertiesClr = valiFrame.GetProcedureRelatedProperties(procedureId)
        properties = {}
        for prop in propertiesClr:
            print("{} : {}".format(prop.Name, prop.Value))
            properties[prop.Name] = prop.Value
    return properties


def ChangeProcedureProperties(valiFrame, procedureId):
    if AskUserToChangeProcedureProperties:
        properties = GetAvailableProcedureProperties(valiFrame, procedureId)
        print('Available properties for procedure %d:' % procedureId)
        for propertyKey in properties:
            print('- %s: %s' % (propertyKey, properties[propertyKey]))
        print('Do you want to change a procedure property (y/n)?')
        change = UserBoolQuery('y')
        while change:
            print('Please enter the name of the procedure property:')
            requestedPropertyName = input()
            okay = False
            for propertyKey in properties:
                if requestedPropertyName == propertyKey:
                    print('Please enter the new value:')
                    newValue = input()
                    valiFrame.SetProcedureProperty(procedureId, propertyKey, newValue)
                    okay = True
            if okay:
                print('Do you want to change another procedure property (y/n)?')
                change = UserBoolQuery('y')
            else:
                print('Invalid input')
    return


def ChangeRelatedProperties(valiFrame, procedureId):
    if AskUserToChangeProcedureProperties:
        properties = GetAvailableRelatedProperties(valiFrame, procedureId)
        print('\nAvailable RELATED properties for procedure %d:' % procedureId)
        for propertyKey in properties:
            print('- %s: %s' % (propertyKey, properties[propertyKey]))
        print('Do you want to change a RELATED procedure property (y/n)?')
        change = UserBoolQuery('y')
        while change:
            print('Please enter the name of the RELATED procedure property:')
            requestedPropertyName = input()
            okay = False
            for propertyKey in properties:
                if requestedPropertyName == propertyKey:
                    print('Please enter the new value:')
                    newValue = input()
                    print("Entered Property Value: {}".format(newValue))
                    valiFrame.SetProcedureProperty(procedureId, propertyKey, newValue)
                    okay = True
            if okay:
                print('Do you want to change another RELATED procedure property (y/n)?')
                change = UserBoolQuery('y')
            else:
                print('Invalid input')
    return


def SelectProcedure(valiFrame):
    procedureIds, procedureNames = GetAvailableProcedures(valiFrame)
    if len(procedureIds) < 1:
        raise RuntimeError('No procedures found')
    elif len(procedureIds) == 1:
        ScriptLog('Selecting procedure %d (no others available)' % procedureIds[0])
        return procedureIds[0]
    else:
        while True:
            print('The following %d procedures were found:' % len(procedureIds))
            i = 0
            while i < len(procedureIds):
                print('- %d: "%s"' % (procedureIds[i], procedureNames[i]))
                i += 1
            print('Please select an ID: ')
            procedureId = int(input())
            return procedureId
            # selectedProcedureIdStr = input()
            # for procedureId in procedureIds:
            # if str(procedureId) == selectedProcedureIdStr:
            # return procedureId
            # print('Invalid selection')


def RunProcedure(valiFrame, procedureId):
    valiFrame.RunProcedure(procedureId)
    return


def RunProcedures(valiFrame):
    if ProcedureIdsToAutoExecute != None:
        print("IF")
        for procedureId in ProcedureIdsToAutoExecute:
            ChangeProcedureProperties(valiFrame, procedureId)
            RunProcedure(valiFrame, procedureId)
    else:
        continueTesting = True
        print("IFE")
        while continueTesting:
            procedureId = SelectProcedure(valiFrame)
            ChangeProcedureProperties(valiFrame, procedureId)
            ChangeRelatedProperties(valiFrame, procedureId)
            RunProcedure(valiFrame, procedureId)
            print('Run another procedure (y/n)?')
            continueTesting = UserBoolQuery('y')
    return


def RunProceduress(valiFrame, procedureId):
    print("IFS")
    for i in range(len(procedureId)):
        print('Test %d Start Now' % procedureId[i])
        RunProcedure(valiFrame, int(procedureId[i]))


def FinishScript():
    if not AutoCloseScript:
        print('Press any key to exit')
        x = input()


################################################ Main Script ################################################
"""
try:

    valiFrame = StartValiFrame()
    RegisterEventHandlers(valiFrame)
    applicationName = SelectApplication(valiFrame)
    InitApplication(valiFrame, applicationName)
    #import sys; sys.exit()
    #valiFrame.LoadProject(r'C:\Caesal\remote_test.vfp')
    ConfigureApplication(valiFrame)
    GetAvailableApplicationProperties(valiFrame)
    print ('b')
    RunProcedures(valiFrame)
    #UnregisterEventHandlers(valiFrame) 

# except Exception, e: # Original-Invalid syntax
# except Exception:
except Exception as e:
    del valiFrame
    print('EXCEPTION: %s' % str(e))

#FinishScript()
"""


################################################ Main Script ################################################
class DP_N5991_ValiFrame(object):
    def Pre():
        valiFrame = StartValiFrame()
        RegisterEventHandlers(valiFrame)
        applicationName = SelectApplication(valiFrame)
        InitApplication(valiFrame, applicationName)
        ConfigureApplication(valiFrame)
        GetAvailableApplicationProperties(valiFrame)
        return valiFrame

    def AutoRun():
        # valiFrame = StartValiFrame()
        RunProcedures(valiFrame)
        # UnregisterEventHandlers(valiFrame)

    def Set(procedureId):  # list only
        # valiFrame = StartValiFrame()
        RunProceduress(valiFrame, procedureId)
        # UnregisterEventHandlers(valiFrame)
# if __name__ == '__main__':
#    ABC()