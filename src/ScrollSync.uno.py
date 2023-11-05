# import logging
import uno
import unohelper

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

PYUNO_LOGLEVEL = 'ARGS'
PYUNO_LOGTARGET = "file:///tmp/LO-ScrollSync"

'''
class Logger_decorator():
    def __init__(self, function):
        LOGLEVEL = 10
        self.function = function
        self.logger = logging.getLogger(self.function.__name__)
        self.fh = logging.FileHandler('ScrollSync.log', mode='w')                                             
        self.fh.setLevel(LOGLEVEL)
        formatter = logging.Formatter(
            fmt='%(asctime)s:%(levelname)s:%(filename)s:%(lineno)d:%(message)s',
            datefmt='%Y-%m-%d-%H:%M:%S',
        )
        self.fh.setFormatter(formatter)
        self.logger.addHandler(self.fh)

    def __call__(self, *args, **kwargs):
        try:
            return self.function(*args,**kwargs)
        except Exception as ex:
            print(ex)
            self.logger.exception(ex)
'''
class AdjustmentListener(unohelper.Base, XAdjustmentListener):
    # Ref: https://help.libreoffice.org/latest/en-US/text/sbasic/python/python_listener.html
    def __init__(self, sync_type, active_doc, inactive_doc):
        self.sync_type = sync_type
        self.active_doc = active_doc
        self.inactive_doc = inactive_doc

    def adjustmentValueChanged(self, evt: ActionEvent):
        # print(evt)
        if self.sync_type == 'ScrollbarPercentage':
            new_pos = self.active_doc.get_rel_scrollbar_pos()
            self.inactive_doc.set_rel_scrollbar_pos(new_pos)
        elif self.sync_type == 'ScrollbarValue':
            new_pos = self.active_doc.get_abs_scrollbar_pos()
            self.inactive_doc.set_abs_scrollbar_pos(new_pos)

    def disposing(self, evt: EventObject):
        pass

class ScrollSyncJob(unohelper.Base, XJobExecutor):
    def __init__(self, ctx):
        self.ctx = ctx
        self.error = False

    # @Logger_decorator
    def trigger(self, sync_type):
        self.sync_type = sync_type
        self.smgr = self.ctx.ServiceManager
        self.desktop = self.smgr.createInstanceWithContext('com.sun.star.frame.Desktop', self.ctx)
        self.tk = self.smgr.createInstanceWithContext('com.sun.star.awt.Toolkit', self.ctx)
        self.parent = self.tk.getDesktopWindow()

        self.active, self.inactive = self.get_docs()
        l_active = AdjustmentListener(self.sync_type, self.active, self.inactive)
        l_inactive = AdjustmentListener(self.sync_type, self.inactive, self.active)
        self.active.scrollbar.addAdjustmentListener(l_active)
        self.inactive.scrollbar.addAdjustmentListener(l_inactive)
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
        if len(text_docs) != 2:
            self.msgbox("ScrollSync requires two open Text documents to run.", 'errorbox')
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
        self.ctx = None
        if ctx is not None:
            self.ctx = ctx
        self.smgr = self.ctx.ServiceManager

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
        for i in range(ch_ctx.getAccessibleChildCount()):
            c = ch_ctx.getAccessibleChild(i)
            c_ctx = c.AccessibleContext
            if c_ctx.ImplementationName == 'com.sun.star.comp.toolkit.AccessibleScrollBar' and c.Orientation == 1:
                return c
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
        return self.scrollbar.AccessibleContext.CurrentValue

    def set_abs_scrollbar_pos(self, value):
        self.scrollbar.AccessibleContext.setCurrentValue(int(value))

    def get_rel_scrollbar_pos(self):
        totaly = float(self.scrollbar.AccessibleContext.MaximumValue)
        currenty = float(self.scrollbar.AccessibleContext.CurrentValue)
        relativey = round(currenty / totaly, 2)
        return relativey

    def set_rel_scrollbar_pos(self, value):
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
