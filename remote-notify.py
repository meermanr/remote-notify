#!/usr/bin/env python
# vim: fileencoding=utf8
b"This line is a syntax error in Python versions older than v2.6."

"""
Recommendations
===============

Ubuntu 11.04's default implementation of the FreeDesktop notifications 
standard is `notify-osd`. This does not support actions or indefinite notices 
(i.e. those that must be dismissed manually). As such it is ill-suited to the 
purpose of this script.

(Also, it doesn't let an application display more than one notice at a time!)

Therefore I recommend installing the older `notificaton-daemon` and replacing 
`notify-osd` [1]_.

.. [1] http://ubuntuforums.org/showpost.php?p=8559795&postcount=8

Outlook 2007
============

Here is a Visual Basic for Applications script to make Outlook 2007 send 
notifications via this script. Note that the original script contains *literal* 
triple-double-quotation, which has to be escaped to avoid problems with 
Python::

    Sub RunAScriptRuleRoutine(MyMail As MailItem)
        Dim strID As String
        Dim olNS As Outlook.NameSpace
        Dim msg As Outlook.MailItem

        strID = MyMail.EntryID
        Set olNS = Application.GetNamespace("MAPI")
        Set msg = olNS.GetItemFromID(strID)
        ' do stuff with msg, e.g.
        'MsgBox msg.Subject

        Dim cmd As String
        cmd = ""\"C:\Python26\python.exe"" ""H:\work\remote-notify\remote-notify.py"" ""outlook"" ""\" & msg.Subject & ""\" ""\" & msg.Body & ""\""
        Shell cmd, vbHide

        Set msg = Nothing
        Set olNS = Nothing
    End Sub

"""
__docformat__ = "restructuredtext en"

import os

giPORT = 6683           # That's "note" in T9 predictive text
grHOST = "e102928-lin"  # Name of host on which to launch web-pages etc

grBaseDir = os.path.dirname( os.path.abspath( __file__ ) )

ICON_OUTLOOK    = grBaseDir + "/icons/outlook.png"
ICON_OFFICE     = grBaseDir + "/icons/office.png"
ICON_WINDOWS    = grBaseDir + "/icons/windows.png"
ICON_TI2        = grBaseDir + "/icons/ti2.png"

# -----------------------------------------------------------------------------
def server():
    """
    Display notifications sent via TCP/IP.
    """
    import sys

    try:
        import pynotify
    except ImportError:
        print >> sys.stderr, "Cannot import 'pynotify', please install " \
                             "'python-notify' package."
        exit(1)

    if not pynotify.init("Remote Notification Service"):
        print >> sys.stderr, "Could not initialise pynotify"
        exit(1)

    dPriority = {
            "low":      (pynotify.URGENCY_LOW,      3600*1000),
            "normal":   (pynotify.URGENCY_NORMAL,   8*3600*1000),
            "critical": (pynotify.URGENCY_CRITICAL, 72*3600*1000),
            }

    dSource = {
            "default":  None,
            "outlook":  ICON_OUTLOOK,
            "window":   ICON_WINDOWS,
            "office":   ICON_OFFICE,
            "ti2":      ICON_TI2,
            }

    import time
    import socket
    import subprocess as sp

    iPORT = giPORT

    if socket.gethostname() == grHOST:
        # Running on server proper, and not via some SSH tunnel etc. Allow 
        # room for SSH tunnelling by using alternative port number
        iPORT += 1
        print "Using alternative port number: %d" % iPORT

    sSock = socket.socket()
    sSock.bind( ("", iPORT) )
    sSock.listen(1)

    print "Listening on", iPORT

    while True:
        sConn, tAddr = sSock.accept()

        rMessage = ""
        while True:
            rData = sConn.recv(1024)
            if rData == "":
                break
            rMessage += rData

        sConn.shutdown( socket.SHUT_WR )     # Indicate finished

        # Spin until remote side is also finished
        while sConn.recv(1024) != "":
            pass

        sConn.close()

        rMessage = rMessage.decode("utf8", "ignore")

        rPriority, rSource, rTitle, rMessage = rMessage.split("\0")
        rTime = time.strftime("%Y-%m-%d %H:%M:%S", time.gmtime())

        if rSource == "outlook":
            # Special case - trim message length
            lLines = rMessage.splitlines()
            for i, rLine in enumerate(lLines):
                if rLine.lower().startswith("dear"):
                    continue
                if not rLine.strip():
                    continue
                rMessage = "\n".join( lLines[i:i+3] )
                break


        rMessage = "<i>%s</i>\n\n" % rTime + rMessage

        print "-" * 80
        print "    Time:", rTime
        print "Priority:", rPriority
        print "  Source:", rSource
        print "   Title:", rTitle
        print " Message:", rMessage

        rIconPath = dSource.get(rSource, dSource["default"])

        sNotification = pynotify.Notification(
            rTitle,
            rMessage,
            rIconPath,
            )

        sUrgency, iTimeout = dPriority.get(rPriority, dPriority["normal"])

        sNotification.set_urgency(sUrgency)
        sNotification.set_timeout(iTimeout)

        sNotification.show()

# -----------------------------------------------------------------------------
def client(sOptions, lArgs):
    R"""
    `lArgs` should look like so:

        1. ``source`` (e.g. "outlook", "ti2")
        2. ``title``
        3. ``message``
    """
    import sys

    rPriority = "normal"

    if sOptions.yLow:
        rPriority = "low"

    if sOptions.yNormal:
        rPriority = "normal"

    if sOptions.yCritical:
        rPriority = "critical"

    if len(sys.argv) < 4:
        # TODO: Implement fallback behaviour (start local browser?)
        return

    rMessage = "\0".join([rPriority] + lArgs)

    import socket
    import subprocess as sp

    sConn = socket.socket()
    try:
        sConn.connect( (grHOST, giPORT) )
    except socket.error:
        # Try alternative port number
        sConn.connect( (grHOST, giPORT+1) )
    sConn.sendall( rMessage )
    sConn.shutdown( socket.SHUT_WR )    # Indicate finished

    # Spin until remote side is also finished
    while sConn.recv(1024) != "":
        pass

    sConn.close()

# -----------------------------------------------------------------------------
def main():
    import platform
    from optparse import OptionParser

    sParser = OptionParser(
            usage="%prog [options] [<source> <title> <message>]",
            description=" ".join(
                x.strip() for x in
                """\
                This script acts and client and server for the transmission 
                and display (respectively) of notification messages.

                Launch with --server to display messages.
                """.splitlines()
            )
        )

    sParser.add_option(
        "--critical",
        action="store_true",
        dest="yCritical",
        default=False)

    sParser.add_option(
        "--normal",
        action="store_true",
        dest="yNormal",
        default=False)

    sParser.add_option(
        "--low",
        action="store_true",
        dest="yLow",
        default=False)

    sParser.add_option(
        "--server",
        action="store_true",
        dest="yServer",
        default=False)

    # Parse command line
    sOptions, lArgs = sParser.parse_args()

    if sOptions.yServer:
        server()
    else:
        client(sOptions, lArgs)

# -----------------------------------------------------------------------------
if __name__ == "__main__":
    main()
