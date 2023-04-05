from dirsync import sync
import win32com.client as win32
from datetime import date
import os
from dirsync.syncer import Syncer

def send_email(to,subject,body,attachment=None):
    # create outlook instance
    outlook = win32.Dispatch('outlook.application')

    # create mail item
    mail = outlook.CreateItem(0)

    # configure email fields
    mail.To = to
    mail.Subject = subject
    mail.Body = body

    # attach file (optional)
    if attachment:
        mail.Attachments.Add(attachment)

    # send email
    mail.Send()


def logging(syncer, logcation):
    report = ""

    report += 'Synchronizing directory \n%s \nwith \n%s\n' %(syncer._dir2, syncer._dir1)

    tt = (str(syncer._endtime - syncer._starttime))[:4]
    report += 'Sync finished in %s seconds.\n' % (tt)
    report += '%d directories parsed, %d files copied\n' % (syncer._numdirs, syncer._numfiles)

    if syncer._numdelfiles:
        report += '%d files were purged.\n' % syncer._numdelfiles
    if syncer._numdeldirs:
        report += '%d directories were purged.\n' % syncer._numdeldirs
    if syncer._numnewdirs:
        report += '%d directories were created.\n' % syncer._numnewdirs
    if syncer._numcontupdates:
        report += '%d files were updated by content.\n' % syncer._numcontupdates
    if syncer._numtimeupdates:
        report += '%d files were updated by timestamp.\n' % syncer._numtimeupdates

    # Failure stats
    if syncer._numcopyfld:
        report += 'there were errors in copying %d files.\n' % syncer._numcopyfld
    if syncer._numdirsfld:
        report += 'there were errors in creating %d directories.\n' % syncer._numdirsfld
    if syncer._numupdsfld:
        report += 'there were errors in updating %d files.\n' % syncer._numupdsfld
    if syncer._numdeldfld:
        report += 'there were errors in purging %d directories.\n' % syncer._numdeldfld
    if syncer._numdelffld:
        report += 'there were errors in purging %d files.\n' % syncer._numdelffld

    with open(logcation,"w+") as log:
        log.write(report)
    
    return report

source = "C:/Users/ICRAVER/Desktop/Demo - AWS S3 SDK/backup_with_gateway/backup_folder"
target = "C:/Users/ICRAVER/Desktop/Demo - AWS S3 SDK/backup_with_gateway/sync_target"
email = "ricko.rinaldy@netmarks.co.id"
log_location = "log.txt"

copier = Syncer(source, target, "sync", purge=True, twoway=True, force=True, verbose=True)
copier.do_work()
report_text = logging(copier,log_location)
send_email(email, "[Notidication] Backup to AWS", report_text)
