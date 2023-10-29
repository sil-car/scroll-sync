import msgbox as util
import uno
import unohelper

from com.sun.star.awt import MessageBoxButtons as MSG_BUTTONS
from com.sun.star.awt import XActionListener
from com.sun.star.awt import ActionEvent
from com.sun.star.lang import EventObject
from com.sun.star.ui.dialogs.ExecutableDialogResults import OK, CANCEL

# References:
#   - UNO API: https://api.libreoffice.org/docs/idl/ref/
#   - OOO Writer API: https://wiki.openoffice.org/wiki/Writer/API/
#   - https://www.pitonyak.org/oo.php
#   - https://wiki.documentfoundation.org/Macros/Python_Design_Guide

# _MY_BUTTON = "CommandButton1"
# _MY_BUTTON2 = "CommandButton2"
# _MY_LABEL = 'Python listens...'
# _DLG_PROVIDER = "com.sun.star.awt.DialogProvider"

class ActionListener(unohelper.Base, XActionListener):
    def __init__(self):
        self.count = 0

    def actionPerformed(self, evt: ActionEvent):
        self.count += 1
        if evt.Source.Model.name == _MY_BUTTON:
            evt.Source.Model.Label = f"{_MY_LABEL} {self.count}"
        # return

    def disposing(self, evt: EventObject):
        pass

class ScrollGui():
    def __init__(self):
        self.ctx = uno.getComponentContext()
        self.smgr = self.ctx.ServiceManager
        self.desktop = self.smgr.createInstanceWithContext('com.sun.star.frame.Desktop', self.ctx)
        self.dlgprov = self.smgr.createInstanceWithContext('com.sun.star.awt.DialogProvider', self.ctx)
        dlgpath = 'Standard.DlgScrollSync?location=appliation'
        # self.dlg = self.dlgprov.createDialog(f"vnd.sun.star.script:{dlgpath}")
        self.active, self.inactive = self.get_docs()

    def get_docs(self):
        text_docs = [d for d in self.desktop.Components if d.supportsService('com.sun.star.text.TextDocument')]
        if len(text_docs) != 2:
            errbox("There needs to be exactly two Text documents open to run this macro.")
            return None
        # TODO: Change "compatible" to mean "same initial paragraph styles IF syncing by heading/paragraph"
        # # Verify that both docs have the same number and type of significant paragraphs.
        # if not docs_are_compatible(text_docs):
        #     msg = f"Docs are not compatible (different number and/or style of paragraphs)."
        #     errbox(msg)
        #     return
        active = ScrollDocument(self.desktop.getCurrentComponent())
        # TODO: "Inactive" could actually be a list...
        inactive_doc_index = 1 - self.get_doc_index(active.doc, text_docs)
        inactive = ScrollDocument(text_docs[inactive_doc_index])
        return (active, inactive)

    def get_doc_index(self, doc, two_docs):
        for i in [0, 1]:
            if doc == two_docs[i]:
                return i
        return None

class ScrollDocument():
    def __init__(self, doc=None):
        self.doc = doc
        self.title = self.doc.Title
        self.view_cursor = self.get_view_cursor()
        self.scrollbar = self.get_scrollbar()
        self.scroll_percent = self.get_rel_scrollbar_pos()
        self.scroll_position = self.get_abs_scrollbar_pos()
        self.cursor_percent = None
        self.cursor_position = None
        self.paragraph_styles = []

    def get_paragraphs(self):
        paragraphs = []
        text = self.doc.Text
        para_enum = text.createEnumeration()
        while para_enum.hasMoreElements():
            para = para_enum.nextElement()
            paragraphs.append(para)
        return paragraphs

    def get_scrollbar(self):
        win_ctx = self.doc.CurrentController.Frame.getComponentWindow().getAccessibleContext()
        ch_ctx = win_ctx.getAccessibleChild(0).getAccessibleContext()
        vscroll_names = [
            'Vertical scroll bar',
            'Barre de défilement verticale',
        ]
        for i in range(ch_ctx.getAccessibleChildCount()):
            c = ch_ctx.getAccessibleChild(i)
            c_ctx = c.getAccessibleContext()
            # print(c_ctx.AccessibleName) # TODO: Locale-dependent!
            if c_ctx.ImplementationName == 'com.sun.star.comp.toolkit.AccessibleScrollBar' and c_ctx.AccessibleName in vscroll_names:
                return c_ctx
        return None

    def get_view_cursor(self):
        cursor = self.doc.CurrentController.getViewCursor()
        return cursor

    def get_abs_cursor_pos(self):
        pass

    def set_abs_cursor_pos(self, value):
        self.view_cursor.gotoRange(value, False)

    def get_rel_cursor_pos(self):
        # Record current position details.
        range_cur = self.view_cursor.getStart()
        pos_cur = self.view_cursor.getPosition()
        # Move cursor to end of doc to get end position; move back.
        set_cursor_position(self.view_cursor, self.doc.Text.End)
        pos_end = self.view_cursor.getPosition()
        set_cursor_position(self.view_cursor, range_cur)
        # Calculate.
        totaly = float(pos_end.Y)
        currenty = float(pos_cur.Y)
        relativey = round(currenty / totaly, 2)
        return relativey

    def set_rel_cursor_pos(self, value):
        pass

    def get_abs_scrollbar_pos(self):
        return self.scrollbar.CurrentValue

    def set_abs_scrollbar_pos(self, value):
        self.scrollbar.setCurrentValue(int(value))

    def get_rel_scrollbar_pos(self):
        totaly = float(self.scrollbar.MaximumValue)
        currenty = float(self.scrollbar.CurrentValue)
        relativey = round(currenty / totaly, 2)
        return relativey

    def set_rel_scrollbar_pos(self, value):
        totaly = float(self.scrollbar.MaximumValue)
        absy = int(float(value) * totaly)
        self.set_abs_scrollbar_pos(absy)

def MsgBox(txt: str):
    mb = util.MsgBox(uno.getComponentContext())
    mb.addButton("OK")
    mb.show(txt, 0, "Python")

def createUnoDialog(libr_dlg: str, embedded=False):
    ctx = XSCRIPTCONTEXT.getComponentContext()
    smgr = ctx.ServiceManager
    if embedded:
        model = XSCRIPTCONTEXT.getDocument()
        dp = smgr.createInstanceWithArguments(_DLG_PROVIDER, (model,))
        # dp = SM.createInstanceWithArguments(_DLG_PROVIDER, (model,))
        location = '?location=document'
    else:
        dp = smgr.createInstanceWithContext(_DLG_PROVIDER, ctx)
        print(type(dp))
        location = '?location=application'
    dlg = dp.createDialog(f"vnd.sun.star.script:{libr_dlg}{location}")
    return dlg

def get_cursor_range_start(cursor):
    return cursor.Start

def set_cursor_position(cursor, range_obj):
    cursor.gotoRange(range_obj, False)

def get_paragraph_index(doc, paragraph):
    # TODO: This assumes each paragraph has unique content!
    paras = get_paragraphs(doc)
    for i, p in enumerate(paras):
        if paragraph.String == p.String:
            return i
    return None

def msgbox(message, title='LibreOffice', buttons=MSG_BUTTONS.BUTTONS_OK, type_msg='infobox'):
    """ Create message box
        type_msg: infobox, warningbox, errorbox, querybox, messbox
        https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1awt_1_1XMessageBoxFactory.html
    """
    toolkit = create_instance('com.sun.star.awt.Toolkit')
    parent = toolkit.getDesktopWindow()
    box = toolkit.createMessageBox(parent, type_msg, buttons, title, str(message))
    return mb.execute()

def errbox(message, title='Sync Cursors', buttons=MSG_BUTTONS.BUTTONS_OK, type_msg='errorbox'):
    toolkit = create_instance('com.sun.star.awt.Toolkit')
    parent = toolkit.getDesktopWindow()
    box = toolkit.createMessageBox(parent, type_msg, buttons, title, str(message))
    box.execute()

def create_instance(name, with_context=False):
    if with_context:
        instance = SM.createInstanceWithContext(name, CTX)
    else:
        instance = SM.createInstance(name)
    return instance

def docs_are_compatible(two_docs):
    paragraph_styles = dict()
    for d in two_docs:
        paragraph_styles[d.Title] = []
        text = d.Text
        para_enum = text.createEnumeration()
        while para_enum.hasMoreElements():
            para = para_enum.nextElement()
            paragraph_styles.get(d.Title).append(para.ParaStyleName)

    doc_titles = list(paragraph_styles.keys())
    return paragraph_styles.get(doc_titles[0]) == paragraph_styles.get(doc_titles[1])

def get_scroll_action(i_initial, i_final):
    action = None
    diff = i_initial - i_final
    if diff > 0: # need to move the view up the page
        action = 0 # line-up, 2 = screen-up
    elif diff < 0: # need to move the view down the page
        action = 1 # line-down, 3 = screen-down,
    return action

def get_relative_cursor_yposition(doc, cursor):
    # Record current position details.
    range_cur = cursor.getStart()
    pos_cur = cursor.getPosition()
    # Move cursor to end of doc to get end position; move back.
    set_cursor_position(cursor, doc.Text.End)
    pos_end = cursor.getPosition()
    set_cursor_position(cursor, range_cur)
    # Calculate.
    totaly = float(pos_end.Y)
    currenty = float(pos_cur.Y)
    relativey = round(currenty / totaly, 2)
    return relativey

def scroll_to_inactive_cursor_location(doc_active, ipara_active, doc_inactive, ipara_inactive, cur_inactive, vsb_inactive):
    # Ref: https://forum.openoffice.org/en/forum/viewtopic.php?t=102273
    action = get_scroll_action(ipara_inactive, ipara_active)
    if action is None:
        return
    # NOTE: The following locations are rounded to the nearest 1% of the maximums.
    cur_inactive_loc = get_relative_cursor_yposition(doc_inactive, cur_inactive)
    vsb_inactive_loc = get_relative_scroll_yposition(vsb_inactive)
    max_scroll_lines = 10000
    ct = 0
    while cur_inactive_loc != vsb_inactive_loc and ct < max_scroll_lines:
        if DEBUG:
            print(f"{cur_inactive_loc = }; {vsb_inactive_loc = }")
        # Move scrollbar.
        vsb_inactive.doAccessibleAction(action)
        # Recalculate.
        vsb_inactive_loc = get_relative_scroll_yposition(vsb_inactive)
        ct += 1
    if ct == max_scroll_lines:
        msgbox(f"Warning: Stopped early. The window has scrolled the maximum number of allowed lines: {max_scroll_lines}.")

def dialogTest(*args):
    # dp = SM.createInstanceWithContext(_DLG_PROVIDER, CTX)
    dp = create_instance(_DLG_PROVIDER, with_context=True)
    libr_dlg = "Standard.DlgScrollSync"
    location = '?location=application'
    ui = dp.createDialog(f"vnd.sun.star.script:{libr_dlg}{location}")
    ui.Title = "Python X[any]Listener"
    b1Ctl = ui.getControl(_MY_BUTTON)
    b2Ctl = ui.getControl(_MY_BUTTON2)
    act = ActionListener()
    b1Ctl.addActionListener(act)
    b2Ctl.addActionListener(act)
    rc = ui.execute()
    if rc == OK:
        MsgBox("User clicked 'OK'")
    elif rc == CANCEL:
        MsgBox("User clicked 'Cancel'")
    ui.dispose()
    b1Ctl.removeActionListener(act)
    b2Ctl.removeActionListener(act)

def updateInactiveDocCursorPosition():
    desktop = get_desktop() # REMOVE
    # Note: Doc indexing seems to be in reverse order of opening; i.e. last opened is i=0.
    two_docs = get_two_docs(desktop)
    if two_docs is None:
        return
    # Verify that both docs have the same number and type of significant paragraphs.
    if not docs_are_compatible(two_docs):
        # TODO: Change "compatible" to mean "same initial paragraph styles"
        msg = f"Docs are not compatible (different number and/or style of paragraphs)."
        errbox(msg)
        return
    # Determine active and inactive documents.
    idoc_active = get_active_doc_index(desktop, two_docs)
    doc_active = two_docs[idoc_active]
    doc_inactive = two_docs[1 - idoc_active]
    # Get cursor start position and paragraph index of active doc.
    cur_active = get_current_cursor(doc_active)
    cur_active_para = get_cursor_range_start(cur_active).TextParagraph
    ipara_active = get_paragraph_index(doc_active, cur_active_para)
    vsb_active = get_vscrollbar_context(doc_active)
    # Get corresponding paragraph of inactive doc.
    cur_inactive = get_current_cursor(doc_inactive)
    cur_inactive_para = get_cursor_range_start(cur_inactive).TextParagraph
    para_inactive = get_paragraphs(doc_inactive)[ipara_active]
    ipara_inactive = get_paragraph_index(doc_inactive, cur_inactive_para)
    vsb_inactive = get_vscrollbar_context(doc_inactive)
    # Place cursor at start of corresponding paragraph of inactive doc.
    set_cursor_position(cur_inactive, para_inactive.getStart())
    # Scroll window to put cursor at top.
    # scroll_to_inactive_cursor_location(doc_active, ipara_active, doc_inactive, ipara_inactive, cur_inactive, vsb_inactive)
    scroll_to_active_scrollbar_value(vsb_active, vsb_inactive)

def updateByScrollbarPercentage():
    app = ScrollGui()
    app.inactive.set_rel_scrollbar_pos(app.active.get_rel_scrollbar_pos())

def updateByScrollbarValue():
    app = ScrollGui()
    app.inactive.set_abs_scrollbar_pos(app.active.scroll_position)

def updateByHeadingPosition():
    app = ScrollGui()
    # TODO.

def updateByParagraphPosition():
    app = ScrollGui()
    # TODO.

g_exportedScripts = (
    updateByScrollbarPercentage,
    updateByScrollbarValue,
    # updateByHeadingPosition,
    # updateByParagraphPosition,
    dialogTest,
)
