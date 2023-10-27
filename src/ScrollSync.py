# import sys
import uno

from com.sun.star.awt import MessageBoxButtons as MSG_BUTTONS
# from com.sun.star.uno import RuntimeException

# References:
#   - UNO API: https://api.libreoffice.org/docs/idl/ref/
#   - OOO Writer API: https://wiki.openoffice.org/wiki/Writer/API/
#   - https://www.pitonyak.org/oo.php
CTX = uno.getComponentContext()
SM = CTX.getServiceManager()
DEBUG = False


def get_current_cursor(doc):
    cursor = doc.CurrentController.getViewCursor()
    return cursor

def get_cursor_range_start(cursor):
    return cursor.Start

def set_cursor_position(cursor, range_obj):
    cursor.gotoRange(range_obj, False)

def get_desktop():
    return SM.createInstanceWithContext('com.sun.star.frame.Desktop', CTX)

def get_paragraph_index(doc, paragraph):
    # TODO: This assumes each paragraph has unique content!
    paras = get_paragraphs(doc)
    for i, p in enumerate(paras):
        if paragraph.String == p.String:
            return i
    return None

def get_paragraphs(doc):
    paragraphs = []
    text = doc.Text
    para_enum = text.createEnumeration()
    while para_enum.hasMoreElements():
        para = para_enum.nextElement()
        paragraphs.append(para)
    return paragraphs

def msgbox(message, title='LibreOffice', buttons=MSG_BUTTONS.BUTTONS_OK, type_msg='infobox'):
    """ Create message box
        type_msg: infobox, warningbox, errorbox, querybox, messbox
        https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1awt_1_1XMessageBoxFactory.html
    """
    toolkit = create_instance('com.sun.star.awt.Toolkit')
    parent = toolkit.getDesktopWindow()
    box = toolkit.createMessageBox(parent, type_msg, buttons, title, str(message))
    box.execute()
    # return mb.execute()

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

def get_two_docs(desktop):
    text_docs = [d for d in desktop.Components if d.supportsService('com.sun.star.text.TextDocument')]
    if len(text_docs) != 2:
        errbox("There needs to be exactly two Text documents open to run this macro.")
    return text_docs

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

def get_active_doc_index(desktop, two_docs):
    active_doc = desktop.getCurrentComponent()
    for i in [0, 1]:
        if active_doc == two_docs[i]:
            return i
    return None

def get_scroll_action(i_initial, i_final):
    action = None
    diff = i_initial - i_final
    if diff > 0: # need to move the view up the page
        action = 0 # line-up, 2 = screen-up
    elif diff < 0: # need to move the view down the page
        action = 1 # line-down, 3 = screen-down,
    return action

def get_vscrollbar_context(doc):
    win_ctx = doc.CurrentController.Frame.getComponentWindow().getAccessibleContext()
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

def get_relative_scroll_yposition(vscrollbar):
    totaly = float(vscrollbar.MaximumValue)
    currenty = float(vscrollbar.CurrentValue)
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

def scroll_to_active_scrollbar_value(vsb_active, vsb_inactive):
    if vsb_active.CurrentValue != vsb_inactive.CurrentValue:
        vsb_inactive.setCurrentValue(vsb_active.CurrentValue)

def update_inactive_doc_cursor_position():
    desktop = get_desktop()
    # Note: Doc indexing seems to be in reverse order of opening; i.e. last opened is i=0.
    two_docs = get_two_docs(desktop)
    # Verify that both docs have the same number and type of significant paragraphs.
    if not docs_are_compatible(two_docs):
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
