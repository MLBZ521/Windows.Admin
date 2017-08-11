XMind
======

A collection of scripts and files I used to deploy the XMind application.


#### install_XMind.ps1 ####

Description:  This script installs XMind.  It also removes the Bonjour installers so that upon user launch of app, the application does not attempt to install it.  (My users did not need this feature.)


#### install_XMind-prefs.ps1 ####

Description:  This script moves custom preference files into place for XMind users.  This script expects the files to have already been moved into place by the XMind installer script.


#### net.xmind.verify.prefs ####

Description:  This file sets the `SUPPRESS_SIGN_IN_DIALOG_ON_STARTUP` key to `true`


#### org.xmind.cathy.prefs ####

Description:  This file sets the `checkUpdatesOnStartup` key to `false`


#### xmind_install-settings.inf ####

Description:  This file is called in a cli switch when installing XMind via cli.
