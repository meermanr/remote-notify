#!/usr/bin/env python
# vim: set fileencoding=utf8 sw=4 ts=4 sts=4 et:
b"This line is a syntax error in Python versions older than v2.6."

"""
Simple client + server for passing notifications from one system to another. 
Primarily intended to capture notifications from MS Outlook in an WinXP VM and 
tunnel them to either the host Linux environment, or over an SSH session to my 
personal PC (e.g. when working from home).

Usage
=====

    1. First, modify the hostname in this script to match your primary 
       workstation (where you want to see the notifications)

    2. Run "./remote-notify.py --server" on this system

    3. Install this same script (including the modifications made in step 1!) 
       on the systems you want to transmit notifications, e.g. your Windows 
       virtual machine. Script that system to call::

            ./remote-notify.py <SourceApp> <NotificationSubject> <NotificationBody>

       For example::

            ./remote-notify.py outlook "SIGPUB" "Meet in lobby at 12h30 for trip to Robin Hood"
            ./remote-notify.py ti2 "EBM Crash" "Executor Board Manager on 10.6.120.3 crashed ..."

       Note that the body can contain multiple lines, so long as you can 
       prevent your OS from mangling the string you want to send.

       Below is a sample Visual Basic for Application (VBA) project with 
       installation instructions which allows you to trigger notifications 
       from Microsoft Outlook 2007, complete with message subject + bodies 
       from incoming mail.

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

(Be sure to update the "cmd" to match your environment!)

Create a VBA project to house the above. This is *not* a macro, but a script 
which will be invoked from one or more mail filtering rules. Create a rule 
(Tools > Rules and Alerts...) with an action "run a script", and select the 
project + method you created above.

Every time you restart Outlook, you will be prompted to enable this macro the 
fist time it is executed. To avoid this you need to sign the VGA project with 
a certificate. This is quite easy: 

    http://office.microsoft.com/en-gb/outlook-help/digitally-sign-a-macro-project-HA001231781.aspx#BM12

"""
__docformat__ = "restructuredtext en"

import os

giPORT = 6683           # That's "note" in T9 predictive text
grHOST = "Az-Pro"       # Name of host to which notifications are sent

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
    import platform

    rSystem = platform.system()
    assert rSystem in ("Linux", "Darwin"), "Unsupported platform!"

    if rSystem == "Linux":
        init_linux()

    dTimeout = {
            "low":      3600*1000,
            "normal":   8*3600*1000,
            "critical": 72*3600*1000,
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
        iTimeout = dTimeout.get(rPriority, dTimeout["normal"])
        rTime = time.strftime("%Y-%m-%d %H:%M", time.gmtime())

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

        print "-" * 80
        print "    Time:", rTime
        print "Priority:", rPriority
        print "  Source:", rSource
        print "   Title:", rTitle
        print " Message:", rMessage

        if rSystem == 'Linux':
            display_linux(rTime, rTitle, rMessage, rSource, rPriority, iTimeout)
        elif rSystem == 'Darwin':
            display_darwin(rTime, rTitle, rMessage, rSource, rPriority, iTimeout)

# -----------------------------------------------------------------------------
def init_linux():
    try:
        import pynotify
    except ImportError:
        print >> sys.stderr, "Cannot import 'pynotify', please install " \
                             "'python-notify' package."
        exit(1)

    if not pynotify.init("Remote Notification Service"):
        print >> sys.stderr, "Could not initialise pynotify"
        exit(1)

# -----------------------------------------------------------------------------
def display_linux(rTime, rTitle, rMessage, rSource, rPriority, iTimeout):
    import pynotify
    dUrgency = {
            "low":      pynotify.URGENCY_LOW,
            "normal":   pynotify.URGENCY_NORMAL,
            "critical": pynotify.URGENCY_CRITICAL,
            }

    dIcon = {
            "default":  None,
            "outlook":  ICON_OUTLOOK,
            "window":   ICON_WINDOWS,
            "office":   ICON_OFFICE,
            "ti2":      ICON_TI2,
            }


    rIconPath = dIcon.get(rSource, dIcon["default"])
    sUrgency = dUrgency.get(rPriority, dUrgency["normal"])
    rMessage = ("<i>%s</i>\n\n" % rTime) + rMessage

    sNotification = pynotify.Notification(
        rTitle,
        rMessage,
        rIconPath,
        )

    sNotification.set_urgency(sUrgency)
    sNotification.set_timeout(iTimeout)

    sNotification.show()

# -----------------------------------------------------------------------------
def display_darwin(rTime, rTitle, rMessage, rSource, rPriority, iTimeout):
    import subprocess

    rMessage = "[%s] %s" % (rTime, rMessage)
    lCMD = ["terminal-notifier",
            "-message", "\\" + rMessage,
            "-title", rTitle,
            "-group", rSource,
            ]
    print lCMD

    subprocess.check_call(lCMD)

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

    try:
        sConn = socket.create_connection( (grHOST, giPORT) )
    except socket.error:
        # Try alternative port number
        sConn = socket.create_connection( (grHOST, giPORT+1) )
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
