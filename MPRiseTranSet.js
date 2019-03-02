//
// Find the next rise, transit and set times for a given minor planet,
// using the ACP configured horizon/minimum-elevation.
//
// Bob Denny for Don Pray (24-May-2006)
// Bob Denny 16-Jan-09  Add endless loop protection and tracing
// Dick Berg 24-Apr-09 - restructure iteration and convergence tests; renovate and remodel
//
var DEGRAD = Math.PI / 180.0;
var RADDEG = 180.0 / Math.PI;
var Desig;
var Elements;
var Util = new ActiveXObject("ACP.Util");
var dtj;
main()

function main()
{
	var Desig = WScript.arguments(0);
	var Elements = WScript.arguments(1);
	
    var PL = new ActiveXObject("NOVAS.Planet");                     // NOVAS for MP
    var ST = new ActiveXObject("NOVAS.Site");
	ST.Height = 1540.2;
	ST.Longitude = -111.760981;
	ST.Latitude = 40.450216;
    var KT = new ActiveXObject("Kepler.Ephemeris");                 // Kepler for MP target
    var KE = new ActiveXObject("Kepler.Ephemeris");                 // Kepler for Earth
    PL.Ephemeris = KT;                                              // Plug in target ephemeris GEN
    PL.EarthEphemeris = KE;                                         // Plug in Earth ephemeris GEN
    PL.Type = 1;                                                    // NOVAS: Minor Planet (Passed to Kepler)
    PL.Number = 1;                                                  // Must pass valid number to Kepler, but is ignored

	KT.Name = Desig;
	KT.Epoch = PackedToJulian(Elements.substring(20,25));
	KT.M = Elements.substring(26,36);                  // Mean anomaly
	KT.Peri = parseFloat(Elements.substring(37, 46));              // Arg of perihelion (J2000, deg.)
	KT.Node = parseFloat(Elements.substring(48, 57));              // Long. of asc. node (J2000, deg.)
    KT.Incl = parseFloat(Elements.substring(59, 68));              // Inclination (J2000, deg.)
	KT.e = parseFloat(Elements.substring(70, 79));                 // Orbital eccentricity
	KT.n = parseFloat(Elements.substring(80, 91));                 // Mean daily motion (deg/day)
	KT.a = parseFloat(Elements.substring(92, 103));                // Semimajor axis (AU)
	
    try {
        var nxtRise = new Date(MPNextRiseSet(PL, ST, true));
        //WScript.Echo(" " + Desig + " rises at " + nxtRise);
        //WScript.Echo(" ");
        var nxtSet = new Date(MPNextRiseSet(PL, ST, false));
        //WScript.Echo(" " + Desig + "  sets at " + nxtSet);
    } catch(ex) {
        WScript.Echo(ex);
        return;
    }
}

// NextRiseSet() - Compute next rise or set time
//
// H0 = target altitude for rise/set
// DoRise = True means calculate next rise, else next set.
//
// Since Kepler/Novas produces J2000 coordinates, we use local
// MEAN sidereal time instead of APPARENT for the hour angle.
// 
//  This algorithm works for both hemispheres
//
function MPNextRiseSet(PL, ST, DoRise)
{
    var JD, RA, Dec, LMST, HA, JDMeridian;
    var  DecRad, LatRad, H0Deg, PV, nIter, LMSTNow;
	var sinElev, sinLat, sinDec, cosLat, cosDec, cosHA, az, el, elm, cosAz;
	var riseAzE, setAzW, deltElev, obsEl;
	var azz, a1, a2, el1, el2, elataz, slope;

    LatRad = ST.Latitude * DEGRAD;	// Latitude in radians (constant)
	sinLat = Math.sin(LatRad);		//  sin(phi) and cos(phi) don't change 
	cosLat = Math.cos(LatRad);
//
// Beginning point
//
    JD = Util.SysJulianDate;      	// Start with current date/Time
//
//   Compute a JD that puts MP on the observer's meridian. This requires the computed RA of the object, which will be 
//     slightly different for the current time compared to the meridian crossing time, but we'll ignore that first. 
//
    PV = PL.GetTopocentricPosition(JD, ST, false);		// This JD is the current date & time
    RA = PV.RightAscension;  //  units  (hours)  				// For MP on meridian, LMST must become equal RA
	Dec = PV.Declination;  	 // units (degress)
	LMSTNow = Util.Julian_GMST(JD) + (ST.Longitude / 15); 	// LOCAL MEAN Sidereal Time; longitude is negative for western Hemisphere
	JDMeridian = JD + (RA - LMSTNow)/24; 					// Change the JD by an amount that puts the MP very close to the meridian
    JD = JDMeridian;
//
//  Second iteration to accommodate slight motion of the MP
    PV = PL.GetTopocentricPosition(JD, ST, false);	// This JD is very close to the meridian crossing time.
    RA = PV.RightAscension;  //  units  (hours)
	Dec = PV.Declination;    // units (degress)
	LMSTNow = Util.Julian_GMST(JD) + (ST.Longitude / 15);
	JDMeridian = JD + (RA - LMSTNow)/24; 			// Change the JD by an amount that puts the MP spot on the meridian		
	JD = JDMeridian;
//
	if (!DoRise) {
        //WScript.Echo("Transit:  RA = " + Util.Hours_HMS(RA) + ";  Dec = " + Util.Degrees_DMS(Dec));
        //WScript.Echo(" " + Desig + " transits " + Util.Julian_Date(JDMeridian));
		WScript.Echo(JDMeridian);
        //WScript.Echo(" ");
	}
// The test above ensures that the meridian crossing time will be printed out on a line between the rising and setting times.
    DecRad = Dec * DEGRAD;
	sinDec = Math.sin(DecRad);
	elm = Util.Prefs.MinimumElevation; 		// units (degrees)  (Straight-line, minimum elevation, from ACP Preferences)
//
//	Now compute the azimuth of the zero-elevation (Earth-horizon) rise and set  points of the MP
	az = Math.acos(sinDec/cosLat);   		// sine law:  cos(Az) = sin(Dec)/cos(Lat)  - units (radians)
	if (!DoRise) az = (2 * Math.PI) - az;	// set this up for the MP setting condition  -  units (radians)
	
//  the  final azimuth with respect to the observable horizon will be somewhere south (for northern hemisphere) or north (for southern hemisphere) of "az"
    H0Deg = 0;   		//   units (degrees)
	nIter = 1;
	deltElev = 0.1;  	// altitude step will be 1/10 degree.
//  
//  The idea is to start with the MP on the horizon and then step the elevation up by small increments, computing the azimuth 
//     at the new elevation from the PZS triangle, and comparing it with the observer's horizon at that new azimuth.
//  Go into elevation-stepping routine.  Increment elevation until it is above observer's horizon.  When that's true,
//		then compute the HA at that point and subtract (for East)/add (for West) it to JDMeridian, and call it quits.
//  Because of the way the visibility horizon is built in ACP, you can't get the correct azimuth any closer than 2 degrees
//  		(the azimuth increment in the table) without interpolation between adjacent azimuths.  The accuracy of
// 		the predicted rise or set will be worst closest to the meridian and best closest to the east or west points.  About 12 minutes at best.
//
//  1.   Pick an elevation (H0Deg) and compute azimuth at this elevation. PZS is constrained by the object's declination and the observer's latitude.
//  2.  Test if elevation is at or above observer's horizon at the same azimuth (max of ACP horizon or minimum horizon)
//  3.  If it is, we're done.  Compute and subtract HA from JDMeridian and post rise time
//  4.  If it's not, we're not done.  Increase elevation by a quarter of a degree. 
//  5.  Go back to do setting time.	

	while (nIter < 901)    	   //  this is just to be sure that the script keeps running.
	{
      PV = PL.GetTopocentricPosition(JD, ST, false);  //  This is at the meridian at Iteration 1
      RA = PV.RightAscension;  //  units (hours) 
	  //WScript.Echo(RA);
      Dec = PV.Declination;    //  units (degrees)
	  //WScript.Echo (Dec);
      DecRad = Dec * DEGRAD;
	  sinDec = Math.sin(DecRad);
	  cosDec = Math.cos(DecRad);
	  cosElev = Math.cos(H0Deg*DEGRAD)	
	  sinElev = Math.sin(H0Deg*DEGRAD);
	  cosAz = (sinDec - sinElev*sinLat)/(cosElev*cosLat);  // cos(Az) runs between +1 and -1 for azimuths 0 - 180.  Nice!
	  if (cosAz < -1) throw ("Object went past the South point below observer's horizon.");
	  if (cosAz >  1) throw ("Object went past the North point below observer's horizon.");
	  az = Math.acos(cosAz);    	 // units (radians) - if rising, azimuth is between 0 and 180
	  if (!DoRise) az = (2 * Math.PI) - az;	   //  if setting, azimuth is between 180 and 360 
// Interpolate the ACP table values of elevations to the computed azimuth.
	  azz = az*RADDEG;
	  a1 = 2 * parseInt(azz/2);						// This is the rounded-down-the-nearest-even-valued azimuth
	  a2 = a1 + 2;									// This is the next even-valued azimuth
	  el1 = Util.Prefs.GetHorizon(a1);
	  el2 = Util.Prefs.GetHorizon(a2);
	  slope = ((el2-el1)/(a2-a1));					//  slope  =  rise/run
	  elataz = slope*(azz-a1) + el1;					//  (y = mx + b)  Interpolate the ACP elevations to the computed azimuth. 
	  if (Math.max(elataz, elm) < H0Deg) break;   	//  This is success!  We broke over the interpolated visibility horizon.
//
//  If we get here, we need another iteration; but we're homing in; update the JD so as to be able to update the MP's  RA and Dec
      HA = Math.acos((Math.sin(H0Deg*DEGRAD) - (sinLat * sinDec)) / 
                       (cosLat * cosDec)) / DEGRAD;  // units  (degrees, not hours)
	  HA = (HA / 15) / 24;    //  units  (fractions of a day)
      if (DoRise) {
		JD = JDMeridian - HA;
	  } else {
		JD = JDMeridian + HA;
	  }
	  //
	  H0Deg += deltElev;   // we didn't rise up, so go back and try the next elevation increment.
	  nIter += 1;
	  if(nIter >= 900) { throw("Failed to converge"); }  // Convergence failed!
// This error will never happen because there are only 900 tenth-degree steps in elevation from 0 degrees to 90 degrees
    }		
//  Success if we got here.  Now compute the final Hour Angle, and add/subtract to the JDMeridian time.
    HA = Math.acos((Math.sin(H0Deg*DEGRAD) - (sinLat * sinDec)) / 
                       (cosLat * cosDec)) / DEGRAD;  // units  (degrees, not hours)
	HA = (HA / 15) / 24;    //  units  (fractions of a day)
    if (DoRise) {
		JD = JDMeridian - HA;
		//WScript.Echo("Rising RA = " + Util.Hours_HMS(RA) + ";  Dec = " + Util.Degrees_DMS(Dec) + ";  Azimuth = " + Util.Degrees_DM(az*RADDEG));
	} else {
		JD = JDMeridian + HA;
        //WScript.Echo("Setting RA = " + Util.Hours_HMS(RA) + ";  Dec = " + Util.Degrees_DMS(Dec) + ";  Azimuth = " + Util.Degrees_DM(az*RADDEG));
    }
    return(Util.Julian_Date(JD));                                   // Local time of rise or set
}

function PackedToJulian(Packed) {
	
    var yr, mo, dy, PCODE, YCODE;
    PCODE = "123456789ABCDEFGHIJKLMNOPQRSTUV";
	YCODE = "IJK";
	
	yr = parseInt(18 + YCODE.indexOf(Packed.substring(0,1))) * 100;     // Century
    yr = yr + parseInt(Packed.substring(1,3));                          // Year in century 
	mo =  parseInt(1 + PCODE.indexOf(Packed.substring(3,4)));			// Month (1-12)
	dy = parseInt(1 + PCODE.indexOf(Packed.substring(4,5)));			// Day (1-31)
	
	DateToJulian(yr, mo, dy, dtj);                   					// UTC Julian Date
	epochJulianFromPacked = dtj; 	
	return epochJulianFromPacked;
}

function DateToJulian(yr, mo, dy) {

    dtj = ((367 * yr) - Math.round((7 * (yr + Math.round((mo + 9) / 12))) / 4) + Math.round((275 * mo) / 9) + dy + 1721073.5);
	return dtj;
}
