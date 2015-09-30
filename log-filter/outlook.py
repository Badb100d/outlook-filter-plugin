#!/bin/env python
# -*- coding:utf-8 -*-
# Outlook logmail filter written by p011ux 20150630
# usage:
# python outlookAddin.py 
# python outlookAddin.py --unregister
# python outlookAddin.py --debug
# debug info in    PythonWin --> Tools --> Trace Collector Debugging Tool

from win32com import universal
from win32com.server.exception import COMException
from win32com.client import gencache, DispatchWithEvents
import winerror
import pythoncom
from win32com.client import constants
import sys

VersionInfo=u"""Log Filter Plugin v0.9 written by p011ux"""

# Forward to these email addr
forwards='''alpha@company.com;beta@company.com'''


# suspicious words,highlight line when match
susp_words       =  ur'^(alpha|beta|else)\b'
susp_words_white =  None

susp_urls       =  ur'(://)|((\s|@)\d{1,3}(\.\d{1,3}){3})'
susp_urls_white =  ur'\b((127(\.\d{1,3}){3}(:(\d{1,5})?)?\b)|(localhost(:\d{1,5})?)'

# dict of blacklists and whitelists
susp_dict={
    'words':(susp_words,susp_words_white),    'urls':(susp_urls,susp_urls_white),}

# tmp folder
BotFolder="_log_bot"
# destination folder
TargetFolder="_log"
# log sender, only handle mails from these
Log_Sender=("alarm-a","alarm-b")
# mail forward tag, make receiver easy to filter
Log_Forward_Tag='[log_tag] '

log_usr=None      # global dest obj
# Support for COM objects we use.
gencache.EnsureModule('{00062FFF-0000-0000-C000-000000000046}', 0, 9, 0, bForDemand=True) # Outlook 9
gencache.EnsureModule('{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}', 0, 2, 1, bForDemand=True) # Office 9

outlookApp=None  #outlook application object

# The TLB defiining the interfaces we implement
universal.RegisterInterfaces('{AC0714F2-3D04-11D1-AE7D-00A0C90F26F4}', 0, 1, 0, ["_IDTExtensibility2"])

# handle content of mail,return a list which contains line numbers of matched.
def suspicious(item):
    import re                             # for regex
    liMatched=[]
    body=item.Body.replace(u'\xa0',u' ')  # replace &nbsp in html format to space
    body=body.replace(u'\r\n',u'\n')      # traverse lines
    cnt=0                                 # line counter
    for eachLine in body.split(u'\n'):
        liLine=eachLine.split()           # split to words
        if(6>len(liLine)):                # get words[5]
            if(0<len(liLine)):
                print len(liLine),u'is less than 6',eachLine
            cnt=cnt+1                     # line cnter
            continue

        cmd=u''                           # recover words[5:]
        for i in liLine[5:]:
            cmd+=i
            cmd+=u' '
        cmd=cmd[:-1]                      # remove last space
        cmd=cmd.lower()

        for i in susp_dict:
            if(None != re.search(susp_dict[i][0],cmd)):                                      # match blacklist
                if((None == susp_dict[i][1]) or (None == re.search(susp_dict[i][1],cmd))):   # if no white or not match white
                    liMatched.append(cnt)                                                    # get current line num
                    break                                                                    # next line
                else:
                    print 'White list:\t',cmd                                                # print out white match
        cnt=cnt+1                         # line cnter
    #now just return
    return liMatched

# adjust font and highlight
def reform(item,linesRed):
    htmB=item.HTMLBody        # get content
    offset=htmB.find('<P>')   # find main content
    offset=offset+len('<P>')

    # adjust font size
    sizeStr='<FONT SIZE='                           # general setting: <P><FONT SIZE=2>
    offset=htmB.find(sizeStr)                       # get FONT SIZE position
    if(0 < offset):                                 # got it?
        offset = offset + len(sizeStr)              # adjust offset
        offend = htmB.find('>',offset)              # get end pos
        if(offset < offend):                        # got end?
            fontsize = int(htmB[offset:offend])     # get font size int
            htmB = htmB[:offset]+str(fontsize+1)+htmB[offend:]    # size+=1
            offset = offend + len('>')

    # highlight
    linesRed=sorted(linesRed)           # sort line nums
    cnt=0                               # cnter
    while True:
        if(cnt in linesRed):            # match
            htmB = htmB[:offset] + '<font color="red">' + htmB[offset:]
            offset = htmB.find('<BR>',offset)
            htmB = htmB[:offset] + '</font>' + htmB[offset:]

        offset = htmB.find('\n',offset) + len('\n') # next
        cnt=cnt+1

        if(cnt > linesRed[-1] or offset >= len(htmB)):
            break

    item.HTMLBody=htmB

# handle mail item
def handleItem(item):
    global log_usr, outlookApp
    try:
        sender=item.SenderName
        if(not sender in Log_Sender):
            # not alarm
            return

        linesMatched=suspicious(item)    # match
        if(0==len(linesMatched)):        # no match?
            # irrelevant log, delete
            # print 'Deleting',item.Subject
            item.Delete()
            return

        reform(item,linesMatched)        # edit mail
        if(None != log_usr):         # forward mail and move it to destination folder
            if(0<len(forwards)):
                if(None!=outlookApp):
                    sendForward=outlookApp.CreateItem(constants.olMailItem)#item.Forward()
                    sendForward.Subject=Log_Forward_Tag+item.SenderName
                    sendForward.To=forwards
                    sendForward.HTMLBody=item.HTMLBody
                    sendForward.DeleteAfterSubmit=True
                    sendForward.Send()   
                else:
                    print 'outlookApp is None'                
            item=item.Move(log_usr)
        else:
            print 'log_usr is None'
    except AttributeError:
        # when registered with python.exe %0 --debug , print to Python Trace Collector
        print "Error handling", repr(item),"in FolderEvent"

# about
class ButtonEvent:
    def OnClick(self, button, cancel):
        # print 'ButtonEvent'
        import win32ui
        win32ui.MessageBox(VersionInfo,u"ABOUT")
        return cancel

# callbacks
# handle mail when new mail arrived in specified folder, delete if not suspicious
class FolderEvent:
    def OnItemAdd(self, item):
        handleItem(item)

# remove permanently when mail with specified sender arrived in Deleted box
class DeletedEvent:
    def OnItemAdd(self, item):
        # print 'DeletedEvent'
        try:
            #print 'Deleted:',item.Subject,'From:',item.SenderName
            if(item.SenderName in Log_Sender):
                item.Delete()
        except AttributeError,e:
            print "Delete", repr(item),"from Deleted error.",e

# initialize each callback
class OutlookAddin:
    _com_interfaces_ = ['_IDTExtensibility2']
    _public_methods_ = []
    _reg_clsctx_ = pythoncom.CLSCTX_INPROC_SERVER
    _reg_clsid_ = "{0F47D9F3-598B-4d24-B7E3-92AC15ED27E2}"
    _reg_progid_ = "Python.Test.OutlookAddin"
    _reg_policy_spec_ = "win32com.server.policy.EventHandlerPolicy"
    def OnConnection(self, application, connectMode, addin, custom):
        global log_usr, outlookApp
        outlookApp=application
        print 'OnConnection'
        # ActiveExplorer may be none when started without a UI (eg, WinCE synchronisation)
        activeExplorer = application.ActiveExplorer()
        if activeExplorer is not None:
            print 'registering button'
            bars = activeExplorer.CommandBars
            toolbar = bars.Item("Standard")
            item = toolbar.Controls.Add(Type=constants.msoControlButton, Temporary=True)
            # Hook events for the item
            item = self.toolbarButton = DispatchWithEvents(item, ButtonEvent)
            item.Caption="LogFilter"
            item.TooltipText = "About"
            item.Enabled = True
            print 'registered'

        # get folder objs
        print 'Accessing folders'
        inbox        = application.Session.GetDefaultFolder(constants.olFolderInbox)
        deleted      = application.Session.GetDefaultFolder(constants.olFolderDeletedItems)
        log_bot = inbox.Folders[BotFolder]
        log_usr = inbox.Folders[TargetFolder]
        print 'Accessed'

        # filter bot folder in each start
        print 'Filtering exists'
        filterList=[]
        for iLog in log_bot.Items:
            filterList.append(iLog)
        for i in filterList:
            handleItem(i)
        print 'Filtered'

        # mesure Deleted after last filter
        print 'Cleaning Deleted'
        dellist=[]
        for iToDel in deleted.Items:
            try:
                if(iToDel.SenderName.encode('gbk') in Log_Sender):
                    dellist.append(iToDel)
            except AttributeError,e:
                print "Error removing", repr(item),"from Deleted"
        for i in dellist:
            i.Delete()
        print 'Cleaned'

        # hook up
        print 'Hooking folders'
        self.deletedItems = DispatchWithEvents(deleted.Items, DeletedEvent)
        self.LogFilter = DispatchWithEvents(log_bot.Items, FolderEvent)
        print 'Hooked'


    def OnDisconnection(self, mode, custom):
        print "OnDisconnection"
        pass
    def OnAddInsUpdate(self, custom):
        print "OnAddInsUpdate", custom
        pass
    def OnStartupComplete(self, custom):
        print "OnStartupComplete", custom
        pass
    def OnBeginShutdown(self, custom):
        print "OnBeginShutdown", custom
        pass

def RegisterAddin(klass):
    import _winreg
    key = _winreg.CreateKey(_winreg.HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins")
    subkey = _winreg.CreateKey(key, klass._reg_progid_)
    _winreg.SetValueEx(subkey, "CommandLineSafe", 0, _winreg.REG_DWORD, 0)
    _winreg.SetValueEx(subkey, "LoadBehavior", 0, _winreg.REG_DWORD, 3)
    _winreg.SetValueEx(subkey, "Description", 0, _winreg.REG_SZ, klass._reg_progid_)
    _winreg.SetValueEx(subkey, "FriendlyName", 0, _winreg.REG_SZ, klass._reg_progid_)

def UnregisterAddin(klass):
    import _winreg
    try:
        _winreg.DeleteKey(_winreg.HKEY_CURRENT_USER, "Software\Microsoft\Office\Outlook\Addins\" + klass._reg_progid_)
    except WindowsError:
        pass

if __name__ == '__main__':
    import win32com.server.register
    win32com.server.register.UseCommandLine(OutlookAddin)
    if "--unregister" in sys.argv:
        UnregisterAddin(OutlookAddin)
    else:
        RegisterAddin(OutlookAddin)
