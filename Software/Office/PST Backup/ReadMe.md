Outlook `.pst` Backup
======

A collection of scripts I wrote to managed a environment that heavily used Outlook `.pst` files and needed a way to back them up.  In addition, Folder Redirection was being enabled for users and I did not want their PST files redirected (stored) onto a server share (this is a well documented bad practice).  So I had to devise a way to move all existing `.pst` files to a location that was not being redirected.  Before enabling Folder Redirection, all users had a 'Windows Backup' task configured for them (this was a manual configuration process that I could not automate and wanted to move away from -- along with it filling up the server disk every two weeks, specifically thanks to the `.pst` files), so their `.pst` files were backed up during this process.  So after Folder Redirection, I still wanted to 'backup' all the users' `.pst` files and found an article that suggested using rsync to accomplish this task.  So this is the direction I went.

For the 'move' portion, specifically connecting to and searching Outlook for PST files, I could not find any completed scripts that accomplished this task in PowerShell (at least when I was looking into this).  There were a handful of VB Scripts that were out there, but I didn't have much luck with getting them to work, and I really wanted to use PowerShell to accomplish the full scope of my plan.  I found enough to get the Outlook portion started, but outside of that this was all written from scratch.  I wrote this some time ago, but if I can find my sources, I'll list them here.  I normally bookmark everything I use.


## Overview ##

First the `move_PSTs.ps1` script should be linked as a user login script.  This script performs various checks and searches to find `.pst` files in the users profile directory and moves the to the configured location in the script.  This essentially 'stages' the `.pst` files for the backup script.  Everything that this script does is logged, in pretty good detail I feel, so it can be reviewed.  It is also commented quite well for reviewing.

Then `backup_rsyncPSTs.ps1` script should be configured as a logoff script.  This script uses a locally copied 'instance' of rsync that I copied down through a GPO Preference item.  The 'backup' location does need to be configured to accept the backups.  Also, the 'output' of the rsync process is captured and acted upon.  If there is no errors, it is logged and if there are errors, this is logged and then an email is sent with the user and error that was captured.  (All logs are backed up as well in the rsync.)  If an error occurred, I normally waited until I received two from the same person before I worried with looking into it.

Just to note, if your users have large `.pst` files, the backup can take quite a bit of time.  I originally wanted the backup to run in the background, as opposed to at logoff, but I never got around to test it.  (To many other things became priority -- this happens as the single IT support personal.)  I'll be honest, this script was my baby for quite a while.  It took quite a bit of time, both at work and personal, to get this working how I wanted.

If I get the chance, I'll try to document the full setup for the server side rsync daemon as well.


#### move_PSTs.ps1 ####

Description:  This script first checks for PST files that are attached to Outlook, moves the attached PSTs to the specified destination, and then reattaches the previously attached PST files to Outlook.  Then searches the user's profile directories for PST files (that are not attached to Outlook) and moves them to the specified destination.  It also accounts for duplicate PST file names.


#### backup_rsyncPSTs.ps1 ####

Description:  This script uses rsync to incrementally backup local PST files. Backup destination is the specified in a variable with exact location configured on the server.