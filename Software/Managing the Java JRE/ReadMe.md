Managing the Java JRE
======

These configuration files are for managing the Java Runtime Environment on client systems.  I used these files and scripts to whitelist websites and Java applets.


## Scripts ##


#### install_JavaCert.ps1 ####

Description:  This script installs a certificate into the default Java `cacerts` file.  This certificate was used to sign the `DeploymentRuleSet.jar` package to whitelist Java applets that are 'pre-approved'.  If this is not done, Firefox will fail the authentication of the `DeploymentRuleSet.jar` package and it will	only work for Internet Explorer.

Note:  (* This has to be a computer login script that runs AFTER every Java update. *)


#### clear_JavaCache.ps1 ####

Description:  This script clears the users' Java Cache.

Note:   (* This has to be a user login script. *)



## Transform Files ##

I used these Transform files in conjunction with the Java `.msi` installers that can be extracted from their `.exe`.  They are configured for a 'silent' deployment and turning off things like the EULA, Updates, etc.

#### Java8_Transform.mst ####

The Java 8 Transform was configured to install the JRE into the same directory instead of unique directories, similar to how the 1.7.x JRE's functioned.

#### Java7_Transform.mst ####



## JRE Configuration Files ##

### Source Information ###
* Java 8:  http://docs.oracle.com/javase/8/docs/technotes/guides/deploy/properties.html
* Java 7:  http://docs.oracle.com/javase/7/docs/technotes/guides/jweb/jcp/properties.html


#### deployment.config ####


#### deployment.properties ####


#### exception.sites ####



## Java Deployment Ruleset ##

### Source Information ###
* Java 8:  http://docs.oracle.com/javase/8/docs/technotes/guides/deploy/deployment_rules.html
* Java 7:  http://docs.oracle.com/javase/7/docs/technotes/guides/jweb/security/deployment_rules.html


#### ruleset.xml ####
