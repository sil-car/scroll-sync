import logging
import platform
import re
import uno
import unohelper

from pathlib import Path

from com.sun.star.awt import MessageBoxButtons as MSG_BUTTONS
from com.sun.star.awt import XAdjustmentListener
from com.sun.star.awt import ActionEvent
from com.sun.star.lang import EventObject
from com.sun.star.task import XJobExecutor

# References:
#   - UNO API: https://api.libreoffice.org [/docs/idl/ref/]
#   - OOO Writer API: https://wiki.openoffice.org/wiki/Writer/API/
#   - https://www.pitonyak.org/oo.php
#   - https://wiki.documentfoundation.org/Macros/Python_Design_Guide
#   - https://help.libreoffice.org/latest/en-US/text/sbasic/python/python_programming.html

LOG_LEVELS = {
    'DEBUG': logging.DEBUG,
    'INFO': logging.INFO,
    'WARNING': logging.WARNING,
    'ERROR': logging.ERROR,
    'CRITICAL': logging.CRITICAL,
}

OXT_DIR = Path(__file__).parent
LO_USER_DIR = OXT_DIR.parents[4]
CFG_FILE = LO_USER_DIR / 'ScrollSync-config.txt'
LOG_FILE = LO_USER_DIR / 'ScrollSync-log.txt'
USER_CONFIG = dict()
if CFG_FILE.is_file():
    config_text = CFG_FILE.read_text()
    for l in config_text.split('\n'):
        l = l.strip()
        if l[0] == '#' or '=' not in l:
            continue
        k, v = l.split('=')
        if k.upper() == 'LOG_LEVEL':
            v = LOG_LEVELS.get(v.upper(), logging.INFO)
        USER_CONFIG[k.strip()] = v.strip()
LOG_LEVEL = USER_CONFIG.get('LOG_LEVEL', logging.INFO)
LOG_FMTR = logging.Formatter(
    fmt='%(asctime)s:%(levelname)s:%(filename)s:%(lineno)d:%(funcName)s:%(message)s',
    datefmt='%Y-%m-%d-%H:%M:%S',
)
LOG_FH = logging.FileHandler(LOG_FILE, encoding='utf8', mode='w')
LOG_FH.setFormatter(LOG_FMTR)
LOG_FH.setLevel(LOG_LEVEL)


class AdjustmentListener(XAdjustmentListener, unohelper.Base):
    # Ref: https://help.libreoffice.org/latest/en-US/text/sbasic/python/python_listener.html
    def __init__(self, sync_type, active_doc, inactive_doc):
        # Set up logging.
        self.logger = logging.getLogger(self.__class__.__name__)
        self.logger.setLevel(LOG_LEVEL)
        self.logger.addHandler(LOG_FH)

        self.sync_type = sync_type
        self.logger.debug(f"AdjustmentListener initialized.")
        self.active_doc = active_doc
        self.inactive_doc = inactive_doc

    def adjustmentValueChanged(self, evt: ActionEvent):
        self.logger.debug('Event caught.')
        if self.sync_type == 'ScrollbarPercentage':
            new_pos = self.active_doc.get_rel_scrollbar_pos()
            self.logger.debug(f"New position (P): {new_pos}")
            self.inactive_doc.set_rel_scrollbar_pos(new_pos)
        elif self.sync_type == 'ScrollbarValue':
            new_pos = self.active_doc.get_abs_scrollbar_pos()
            self.logger.debug(f"New position (I): {new_pos}")
            self.inactive_doc.set_abs_scrollbar_pos(new_pos)

    def disposing(self, evt: EventObject):
        pass

class ScrollSyncJob(unohelper.Base, XJobExecutor):
    def __init__(self, ctx):
        self.error = False
        desc = OXT_DIR / 'description.xml'
        m = re.search(r'(?<=version value=")[0-9\.]+', desc.read_text())
        if m is not None:
            self.version = m[0]
        else:
            self.version = 'unknown'

        # Set LO services.
        self.ctx = ctx
        self.smgr = self.ctx.ServiceManager

        # Get LO version.
        cfg = self.smgr.createInstance('com.sun.star.configuration.ConfigurationProvider')
        arg = uno.createUnoStruct('com.sun.star.beans.PropertyValue')
        arg.Name = "nodepath"
        arg.Value = "/org.openoffice.Setup/Product"
        node = cfg.createInstanceWithArguments("com.sun.star.configuration.ConfigurationAccess", (arg,))
        lo_ver = node.getByName("ooSetupVersion")
        del node
        del cfg
        del arg

        # Set up logging.
        self.logger = logging.getLogger(self.__class__.__name__)
        self.logger.setLevel(LOG_LEVEL)
        self.logger.addHandler(LOG_FH)

        # Log basic info.
        pf = platform.uname()
        self.logger.info(f"OS: {pf.system} / {pf.release} / {pf.version}")
        self.logger.info(f"LO version: {lo_ver}")
        self.logger.info(f"Python version: {platform.python_version()}")
        self.logger.info(f"ScrollSync version: {self.version}")
        self.logger.debug("ScrollSyncJob initialized.")

    def trigger(self, sync_type):
        self.logger.debug("ScrollSyncJob triggered.")
        self.sync_type = sync_type
        self.logger.info(f"Sync type: {sync_type}")
        self.desktop = self.smgr.createInstanceWithContext('com.sun.star.frame.Desktop', self.ctx)
        self.components = [c for c in self.desktop.Components]
        self.logger.debug(f"Found {len(self.components)} LO desktop components.")
        self.tk = self.smgr.createInstanceWithContext('com.sun.star.awt.Toolkit', self.ctx)
        self.parent = self.tk.getDesktopWindow()
        self.logger.debug(f"Parent window acquired.")

        # Get documents.
        self.active, self.inactive = self.get_docs()
        # self.active = ScrollDocument(self.desktop.getCurrentComponent(), self.ctx)
        if self.active is None:
            self.logger.warning("No active document found.")
        else:
            self.logger.info(f"Active doc name: {self.active.doc.Title}")
        if self.inactive is None:
            self.logger.warning("No inactive document found.")
        else:
            self.logger.info(f"Inactive doc name: {self.inactive.doc.Title}")

        # Configure listeners.
        # if self.active.scroll_listener is not None:
        #     print("Removing old listener.")
        #     self.active.scrollbar.removeAdjustmentListener(self.active.scroll_listener)
        # if self.inactive.scroll_listener is not None:
        #     print("Removing old listener.")
        #     self.inactive.scrollbar.removeAdjustmentListener(self.inactive.scroll_listener)
        self.active.scroll_listener = AdjustmentListener(self.sync_type, self.active, self.inactive)
        self.inactive.scroll_listener = AdjustmentListener(self.sync_type, self.inactive, self.active)
        self.active.scrollbar.addAdjustmentListener(self.active.scroll_listener)
        self.inactive.scrollbar.addAdjustmentListener(self.inactive.scroll_listener)
        self.logger.info('AdjustmentListeners added to vertical scrollbars.')

        # Do a nasty thing before exiting the python process. In case the
        # last call is a oneway call (e.g. see idl-spec of insertString),
        # it must be forced out of the remote-bridge caches before python
        # exits the process. Otherwise, the oneway call may or may not reach
        # the target object.
        # I do this here by calling a cheap synchronous call (getPropertyValue).
        # Ref: https://www.openoffice.org/udk/python/python-bridge.html
        # ctx.ServiceManager

    def get_docs(self):
        text_docs = [d for d in self.desktop.Components if d.supportsService('com.sun.star.text.TextDocument')]
        self.logger.debug(f"Found {len(text_docs)} open Text documents.")
        if len(text_docs) != 2:
            self.msgbox("ScrollSync requires exactly two open Text documents to run.", 'errorbox')
            self.error = True
            return None, None
        # TODO: Change "compatible" to mean "same initial paragraph styles IF syncing by heading/paragraph"
        # # Verify that both docs have the same number and type of significant paragraphs.
        # if not docs_are_compatible(text_docs):
        #     msg = f"Docs are not compatible (different number and/or style of paragraphs)."
        #     errbox(msg)
        #     return
        active = ScrollDocument(self.desktop.getCurrentComponent(), self.ctx)
        # TODO: "Inactive" could actually be a list...
        inactive_doc_index = 1 - self.get_doc_index(active.doc, text_docs)
        inactive = ScrollDocument(text_docs[inactive_doc_index], self.ctx)
        return (active, inactive)

    def get_doc_index(self, doc, two_docs):
        for i in [0, 1]:
            if doc == two_docs[i]:
                return i
        return None

    def msgbox(self, message, type_msg='infobox'):
        buttons = MSG_BUTTONS.BUTTONS_OK
        mb = self.tk.createMessageBox(self.parent, type_msg, buttons, "ScrollSync", str(message))
        return mb.execute()

class ScrollDocument():
    def __init__(self, doc=None, ctx=None):
        # Set up logging.
        self.logger = logging.getLogger(self.__class__.__name__)
        self.logger.setLevel(LOG_LEVEL)
        self.logger.addHandler(LOG_FH)

        self.doc = doc
        self.title = self.doc.Title
        self.logger.info(f"Found Text Document: {self.title}")
        # self.ctx = None
        # if ctx is not None:
        #     self.ctx = ctx
        # else:
        #     logger.warning(f"Component context not provided for Text Document \"{self.title}\"")
        # self.smgr = self.ctx.ServiceManager

        self.frm = self.doc.CurrentController.Frame
        self.wdw = self.frm.getComponentWindow()
        # self.view_cursor = self.get_view_cursor()
        self.scrollbar = self.get_scrollbar()
        if self.scrollbar is None:
            self.logger.error("Unable to get scrollbar control.")
            return
        self.scroll_position = self.get_abs_scrollbar_pos()
        self.scroll_percent = self.get_rel_scrollbar_pos()
        # self.cursor_percent = None
        # self.cursor_position = None
        # self.paragraph_styles = []

    # def get_paragraphs(self):
    #     paragraphs = []
    #     text = self.doc.Text
    #     para_enum = text.createEnumeration()
    #     while para_enum.hasMoreElements():
    #         para = para_enum.nextElement()
    #         paragraphs.append(para)
    #     return paragraphs

    def get_scrollbar(self):
        def find_vert_scrollbar(obj, d=0):
            found = None
            ctx = obj.AccessibleContext
            # self.logger.debug(f"d:{d}; {ctx.ImplementationName}")
            if ctx.ImplementationName == 'com.sun.star.comp.toolkit.AccessibleScrollBar':
                # self.logger.debug(f"{obj.Orientation}")
                if obj.Orientation == 1:
                    self.logger.debug(f"Found vertical scrollbar.")
                    return obj
            else:
                for i in range(ctx.getAccessibleChildCount()):
                    c = ctx.getAccessibleChild(i)
                    found = find_vert_scrollbar(c, d+1)
                    if found is not None:
                        return found

        return find_vert_scrollbar(self.wdw)


    # def get_view_cursor(self):
    #     return self.doc.CurrentController.getViewCursor()

    # def get_abs_cursor_pos(self):
    #     pass

    # def set_abs_cursor_pos(self, value):
    #     self.view_cursor.gotoRange(value, False)

    # def get_rel_cursor_pos(self):
    #     # Record current position details.
    #     range_cur = self.view_cursor.getStart()
    #     pos_cur = self.view_cursor.getPosition()
    #     # Move cursor to end of doc to get end position; move back.
    #     set_cursor_position(self.view_cursor, self.doc.Text.End)
    #     pos_end = self.view_cursor.getPosition()
    #     set_cursor_position(self.view_cursor, range_cur)
    #     # Calculate.
    #     totaly = float(pos_end.Y)
    #     currenty = float(pos_cur.Y)
    #     relativey = round(currenty / totaly, 2)
    #     return relativey

    # def set_rel_cursor_pos(self, value):
    #     pass

    def get_abs_scrollbar_pos(self):
        self.logger.debug(f"Getting absolute scrollbar position for \"{self.title}\"")
        # self.logger.debug(f"Position: {self.scrollbar.AccessibleContext.CurrentValue}")
        # pf = platform.uname()
        # if pf.system == 'Windows' and pf.release == '10':
        #     self.logger.debug(f"Special handling for Windows 10.")
        #     self.logger.debug(dir(self.scrollbar))
        #     pos = int(self.scrollbar.getCurrentValue())
        # else:
        #     pos = int(self.scrollbar.getAccessibleContext().getCurrentValue())
        # self.logger.debug(f"Position: {pos}")
        return int(self.scrollbar.AccessibleContext.CurrentValue)
        # return pos

    def set_abs_scrollbar_pos(self, value):
        self.logger.debug(f"Setting absolute scrollbar position to {value} for \"{self.title}\"")
        self.scrollbar.AccessibleContext.setCurrentValue(int(value))

    def get_rel_scrollbar_pos(self):
        self.logger.debug(f"Calculating relative scrollbar position for \"{self.title}\"")
        totaly = float(self.scrollbar.AccessibleContext.MaximumValue)
        # currenty = float(self.scrollbar.AccessibleContext.CurrentValue)
        currenty = float(self.get_abs_scrollbar_pos())
        relativey = round(currenty / totaly, 2)
        return relativey

    def set_rel_scrollbar_pos(self, value):
        self.logger.debug(f"Setting relative scrollbar position for \"{self.title}\"")
        totaly = float(self.scrollbar.AccessibleContext.MaximumValue)
        absy = int(float(value) * totaly)
        self.set_abs_scrollbar_pos(absy)


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

def disableScrollSync():
    """ Disable page synchronization. """
    # TODO: Not yet implemented.
    # Emit a signal/event that is caught and handled by the other macros?

def updateByScrollbarPercentage():
    """ Scroll the inactive document the same percent distance as the active document. """
    # TODO: Fix docstrings so that "Description" works properly.
    app = ScrollSync()
    if app.error:
        return
    app.trigger('ScrollbarPercentage')
    # app.inactive.set_rel_scrollbar_pos(app.active.get_rel_scrollbar_pos())

def updateByScrollbarValue():
    """ Scroll the inactive document to the same line number as the active document. """
    # TODO: Fix docstrings so that "Description" works properly.
    app = ScrollSync()
    if app.error:
        return
    app.trigger('ScrollbarValue')
    # app.inactive.set_abs_scrollbar_pos(app.active.scroll_position)

def updateByHeadingPosition():
    app = ScrollSync('HeadingPosition')
    # TODO: Not yet implemented.

def updateByParagraphPosition():
    app = ScrollSync('ParagraphPosition')
    # TODO: Not yet implemented.


g_exportedScripts = (
    # disableScrollSync,
    updateByScrollbarPercentage,
    updateByScrollbarValue,
    # updateByHeadingPosition,
    # updateByParagraphPosition,
)

g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
    ScrollSyncJob,
    "org.sil-car.ScrollSync.ScrollSync",
    ("com.sun.star.task.Job",),
)


# if __name__ == "__main__":
#     import os
 
#     # Start OpenOffice, listen for connections and open testing document
#     os.system( "/etc/openoffice.org-1.9/program/soffice '-accept=socket,host=localhost,port=2002;urp;' -writer ./WaveletTest.odt &" )
 
#     # Get local context info
#     localContext = uno.getComponentContext()
#     resolver = localContext.ServiceManager.createInstanceWithContext(
#         "com.sun.star.bridge.UnoUrlResolver", localContext )
 
#     ctx = None
 
#     # Wait until the OpenOffice starts and connection is established
#     while ctx == None:
#         try:
#             ctx = resolver.resolve(
#                 "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext" )
#         except:
#             pass
 
#     # Trigger our job
#     wavelet = Wavelet( ctx )
#     wavelet.trigger( () )
