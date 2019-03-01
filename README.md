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

There are several values that will need to be tailored to your capabilities in the beginning of the script. In addition there is lat/long/elevation in MPRiseTranSet.js that will need to be set for your site.

I'm still working on consolidating all user set parameters in the beginning of the script, there may be several lingering path statements that are hard coded in the body of the script

find_orb directory is one of these. 
You'll need to put your observatory code in for the lines that call fo.exe, you might be able to do this with a lat/long, but I havent dug into the fo.exe code to see if that would work.
