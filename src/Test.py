# -*- coding: utf-8 -*-
from __future__ import unicode_literals

import uno, unohelper
from com.sun.star.awt import XActionListener
from com.sun.star.awt import ActionEvent
from com.sun.star.lang import EventObject
from com.sun.star.ui.dialogs.ExecutableDialogResults \
    import OK, CANCEL
import msgbox as util

_MY_BUTTON =  "Button1"
_MY_LABEL = 'Python listens..'
_DLG_PROVIDER = "com.sun.star.awt.DialogProvider"

def Main(*args):
    ui = createUnoDialog("Standard.Dialog1", embedded=True)
    ui.Title = "Python X[any]Listener"
    ctl = ui.getControl(_MY_BUTTON)
    ctl.Model.Label = _MY_LABEL
    act = ActionListener()
    ctl.addActionListener(act)
    rc = ui.execute()
    if rc == OK:
        MsgBox("The user acknowledged the dialog.")
    elif rc == CANCEL:
        MsgBox("The user canceled the dialog.")
    ui.dispose()  # ui.endExecute
    ctl.removeActionListener(act)

def createUnoDialog(libr_dlg: str, embedded=False):
    """ Create a Dialog from its location """
    smgr = XSCRIPTCONTEXT.getComponentContext().ServiceManager
    if embedded:
        model = XSCRIPTCONTEXT.getDocument()
        dp = smgr.createInstanceWithArguments(_DLG_PROVIDER, (model,))
        location = "?location=document"
    else:
        dp = smgr.createInstanceWithContext(_DLG_PROVIDER, ctx)
        location = "?location=application"
    dlg = dp.createDialog("vnd.sun.star.script:"+libr_dlg+location)
    return dlg

class ActionListener(unohelper.Base, XActionListener):
    """ Listen to & count button clicks """
    def __init__(self):
        self.count = 0

    def actionPerformed(self, evt: ActionEvent):
        self.count = self.count + 1
        #mri(evt)
        if evt.Source.Model.Name == _MY_BUTTON:
            evt.Source.Model.Label = _MY_LABEL+ str( self.count )
    # return

    def disposing(self, evt: EventObject):  # mandatory routine
        pass

def MsgBox(txt: str):
    mb = util.MsgBox(uno.getComponentContext())
    mb.addButton("Ok")
    mb.show(txt, 0, "Python")

g_exportedScripts = (Main,)
