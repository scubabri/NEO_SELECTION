# NEO_SELECTION
this is a new project for my NEOCP and ESA priority selection script

Requirements:

ACP Expert and ACP Scheduler https://acpx.dc3.com/
Ascom 6.4SP1 https://ascom-standards.org/
Ascom developer componants https://ascom-standards.org/Downloads/PlatDevComponents.htm
Kepler Orbit Engine https://ascom-standards.org/Downloads/ScriptableComponents.htm
NOVAS-COM V2.11 https://ascom-standards.org/Downloads/ScriptableComponents.htm
Requires an MPC Observatory code https://minorplanetcenter.net/iau/info/Astrometry.html#HowObsCode

Use: 

ACP will need a user called 'neocp' 

This needs to be run from a 32bit shell on windows 10. to create a shell, the easiest way is to create a shortcut on your desktop to C:\windows\SysWOW64\cmd.exe

To run the script, in a 32bit shell use, "cscript process_neo_jobs.vbs nightly"  (this will wipe the MPCORB.dat and download a new one) this will generate an output as well as 
It can also be run on demand, appending new objects found on the NEOCP and ESA priority list, usage is like above, replacing "nightly" with "hourly"

There are several values that will need to be tailored to your capabilities in the beginning of the script. In addition there is lat/long/elevation in MPRiseTranSet.js that will need to be set for your site.

I'm still working on consolidating all user set parameters in the beginning of the script, there may be several lingering path statements that are hard coded in the body of the script

find_orb directory is one of these.  The script is expecting it to be in C:\find_o64\
You'll need to put your observatory code in for the lines that call fo.exe, you might be able to do this with a lat/long, but I havent dug into the fo.exe code to see if that would work.

Please excuse the haphazard nature of my coding, this has been a work in progress for about 4 months. There is little error checking and if it has any issues getting data from the NEOCP or ESA priorities list, the script will fail most ungracefully. 
