Attribute VB_Name = "AeroToolKit"
'***************************************************************************
'Aero Toolkit Add-In for Excel
'August 2018, Lance Bays, www.veranautics.com
'Please direct questions, comments, requests, etc., to: veranautics@gmail.com
'
'ABOUT:
'Excel add-in with common aeronautical engineering functions, including standard atmosphere
'characteristics, QNH, QFE, altimetry, obstacle clearance, airspeeds, climb gradients, turns
'and common aeronautical unit conversions. The atmospheric model is based on the ICAO standard
'atmosphere (as documented in ICAO 7488). Altitude inputs up to the stratopause (51 km or
'167,000 ft) are permitted. Non-standard day characteristics are available via ISA deviations.
'
'For reference, the ICAO standard atmosphere, the ISA standard atmosphere, and the US 1962
'and 1976 standard atmospheres are identical up to 32 km (105,000 ft).
'
'For a comprehensive document substantiating these functions, and a handy "cheat sheet,"
'go to: veranautics.com/AeroToolKit
'
'These functions have been implemented in other programming languages and may be available
'for licensed use. Contact veranautics@gmail.com for more information.
'
'LICENSE:
'By downloading or using this software you acknowledge the following:
'
'Copyright (c) 2018 Lance Bays
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'copies of the Software, and to permit persons to whom the Software is
'furnished to do so, subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'SOFTWARE.
'
'(MIT License)
'
'***************************************************************************
'Program execution constants
Const constMaxAeroIter = 10 'Maximum iterations in supersonic convergence of Mach and KCAS
Const constMachEpsilon = 0.000001 'Arbitrary convergence epsilon for Mach convergence at supersonic speeds
Const constKcasEpsilon = 0.0001 'Arbitrary convergence epsilon for KCAS convergence at supersonic speeds
Const constAeroErr = -99999999.9999999 'Initialization value, an arbitrary large-magnitude, oddball, negative value
Const constMinDAlt = 100 'Altitude step for geopotential-pressure altitude conversions
'Physical constants defining the standard atmosphere
Const constGammaAir = 1.4 'From ICAO 7488, ratio of specific heats for air
Const constGo = 32.1740485564304  'From ICAO 7488 go=9.80665, standard acceleration due to gravity at latitude 45�32�33�� using Lambert�s equation of the acceleration due to gravity function latitude (PER ICAO 7488 SEE: U.S. Committee Extension Standard Atmosphere: U.S. Standard Atmosphere, 1962. U.S. Government)
Const constRAir = 3089.81137753942 'From ICAO 7488, R=287.05287 m2/(s2 K) converted with 0.3048 m/ft, gas constant, dry air, ft2/s2/K
Const constRadiusEarth = 20855531.496063 'From ICAO 7488, r=6356766 m, 0.3048 m/ft, nominal radius earth, feet
Const constBetaVisc = 0.000001458 'From ICAO 7488, Reynolds calculation constant, N-sec/(sq.m-sqrt(kelvins))
Const constSuth = 110.4 'From ICAO 7488, Sutherland�s constant, Kelvin
'Characteristics at sea level
Const constTo = 288.15 'From ICAO 7488, sea level value of absolute temperature in standard atmosphere, K (518.67 R)
Const constAo = 661.478594435162 'Derived from From ICAO 7488 values for To, R. Uses 1852 m/nm. Speed of sound at sea level, knots
Const constPo = 2116.21662367394 'From ICAO 7488 Po=101325 N/m2, converted by 0.3048 m/ft & 4.4482216152605 N/lbf
Const constRhoo = 2.37689240667515E-03 'From ICAO 7488, rhoSL=l.225 kg/m3, 4.4482216152605 (kg m)/(lbf s2) & 1 slug = 1 (s2 lbf)/ft
'Lapse rates to top of stratosphere
Const constLapseTrop = -0.0019812 'From ICAO 7488 standard lapse rate in troposhere, -6.50�C/(1000 m) & 0.3048 m/ft
'constLStrat1 would be zero (lower stratosphere is isothermal) - not needed
Const constLapseStrat2 = 0.0003048 'From ICAO 7488, lapse rate from 20 km to 32 km +1.00�C/(1000 m) & 0.3048 m/ft
Const constLapseStrat3 = 0.00085344 'From ICAO 7488, lapse rate from 32 km to 47 km +2.80�C/(1000 m) & 0.3048 m/ft
'Geopotential altitudes (in feet) at breaks in temperature profile
Const tropopause11kmInFt = 36089.2388451444 'From ICAO 7488, 11000 m / 0.3048 m/ft
Const topIsoThermLayerStrat20kmInFt = 65616.7979002625
Const toplstInverLayerStrat32kmInFt = 104986.87664042 'From ICAO 7488, 32000 m / 0.3048 m/ft
Const stratopauseStart47kmInFt = 154199.475065617 'From ICAO 7488, 47000 m / 0.3048 m/ft
Const stratopauseEnd51kmInFt = 167322.83464567 'From ICAO 7488, 51000 m / 0.3048 m/ft
'Atmospheric temperature (in Kelvin) at breaks in temperature profile
Const constOatIsoLayerStrat11to20kmInK = 216.65 'From ICAO 7488, OAT in isothermal layer of lower stratosphere, K
Const constOatStrat32kmInK = 228.65 'From ICAO 7488, OAT at transition between first and second inversion layer, K
Const constOatStratopause47to51kmInK = 270.65 'From ICAO 7488, OAT in stratopause (47 to 51 km), K
'Atmospheric pressure ratios (deltas) at breaks in temperature profile
Const constDeltaTropopause = 0.223360869430129   '0.22336086943012873
Const constDeltaStrat20km = 0.054032839124412    '0.054032839124412
Const constDeltaStrat32km = 8.5666496582306E-03  '0.008566649658230598
Const constDeltaStrat47km = 1.09455488149331E-03 '0.0010945548814933147
'Conversion Factors
Const constConvLbfPerInch2ToLbfPerFt2 = 144 'lb/in2 to lb/ft2 (exact)
Const constConvKelvinToRankine = 1.8 'Kelvin to Rankine (exact)
Const constConvFtToM = 0.3048 'feet to meters (exact)
Const constConvNmToM = 1852 'nm to meters (exact)
Const constConvHrToSec = 3600 'Seconds per hour (exact)
Const constConvStatuteMileToFt = 5280 'Statute miles to feet(exact)
Const constConvLbfToNewton = 4.4482216152605 'lb force to Newton
Const constConvFtPerSecToKts = 0.592483801295896 'ft/sec to knots, derived from 1852 m/nm, 3600 sec/hr & 0.3048 m/ft
Const constConvInHgToHectoPascal = 33.86389 'hPa to inHg, based on NIST Special Pub 811, 2008
'Temperatures for zero deg C and zero deg F on absolute temperature scales
Const constZeroDegCelsiusInKelvin = 273.15 'Temperature in Kelvin at zero Celisus, K
Const constZeroDegFahrenheitInRankine = 459.67 'Temperature in Rankine at zero Fahrenheit, R

Function AeroSpdSnd_ftPerSec_fOatKelvin(oatKelvin)
    If (oatKelvin < 0) Then
        AeroSpdSnd_ftPerSec_fOatKelvin = constAeroErr
    Else
        AeroSpdSnd_ftPerSec_fOatKelvin = (constGammaAir * constRAir * oatKelvin) ^ 0.5
    End If
End Function

Function AeroSpdSnd_kts_fOatKelvin(oatKelvin)
    AeroSpdSnd_kts_fOatKelvin = AeroSpdSnd_ftPerSec_fOatKelvin(oatKelvin) * constConvFtPerSecToKts
End Function

Function AeroSpdSnd_statuteMilesPerHr_fOatKelvin(oatKelvin)
    AeroSpdSnd_statuteMilesPerHr_fOatKelvin = AeroSpdSnd_ftPerSec_fOatKelvin(oatKelvin) * constConvHrToSec / constConvStatuteMileToFt
End Function

Function AeroSpdSnd_kmPerHr_fOatKelvin(oatKelvin)
    AeroSpdSnd_kmPerHr_fOatKelvin = AeroSpdSnd_ftPerSec_fOatKelvin(oatKelvin) * constConvHrToSec * constConvFtToM / 1000
End Function

Function AeroSpdSnd_mPerSec_fOatKelvin(oatKelvin)
    AeroSpdSnd_mPerSec_fOatKelvin = AeroSpdSnd_ftPerSec_fOatKelvin(oatKelvin) * constConvFtToM
End Function

Function AeroSpdSnd_ftPerSec_fOatCelsius(oatCelsius)
    AeroSpdSnd_ftPerSec_fOatCelsius = AeroSpdSnd_ftPerSec_fOatKelvin(AeroConvCelsiusToKelvin(oatCelsius))
End Function

Function AeroSpdSnd_kts_fOatCelsius(oatCelsius)
    AeroSpdSnd_kts_fOatCelsius = AeroSpdSnd_kts_fOatKelvin(AeroConvCelsiusToKelvin(oatCelsius))
End Function

Function AeroSpdSnd_statuteMilesPerHr_fOatCelsius(oatCelsius)
    AeroSpdSnd_statuteMilesPerHr_fOatCelsius = AeroSpdSnd_statuteMilesPerHr_fOatKelvin(AeroConvCelsiusToKelvin(oatCelsius))
End Function

Function AeroSpdSnd_kmPerHr_fOatCelsius(oatCelsius)
    AeroSpdSnd_kmPerHr_fOatCelsius = AeroSpdSnd_kmPerHr_fOatKelvin(AeroConvCelsiusToKelvin(oatCelsius))
End Function

Function AeroSpdSnd_mPerSec_fOatCelsius(oatCelsius)
    AeroSpdSnd_mPerSec_fOatCelsius = AeroSpdSnd_mPerSec_fOatKelvin(AeroConvCelsiusToKelvin(oatCelsius))
End Function

Function AeroSpdSnd_ftPerSec_fOatFahrenheit(oatFahrenheit)
    AeroSpdSnd_ftPerSec_fOatFahrenheit = AeroSpdSnd_ftPerSec_fOatKelvin(AeroConvFahrenheitToKelvin(oatFahrenheit))
End Function

Function AeroSpdSnd_kts_fOatFahrenheit(oatFahrenheit)
    AeroSpdSnd_kts_fOatFahrenheit = AeroSpdSnd_kts_fOatKelvin(AeroConvFahrenheitToKelvin(oatFahrenheit))
End Function

Function AeroSpdSnd_statuteMilesPerHr_fOatFahrenheit(oatFahrenheit)
    AeroSpdSnd_statuteMilesPerHr_fOatFahrenheit = AeroSpdSnd_statuteMilesPerHr_fOatKelvin(AeroConvFahrenheitToKelvin(oatFahrenheit))
End Function

Function AeroSpdSnd_kmPerHr_fOatFahrenheit(oatFahrenheit)
    AeroSpdSnd_kmPerHr_fOatFahrenheit = AeroSpdSnd_kmPerHr_fOatKelvin(AeroConvFahrenheitToKelvin(oatFahrenheit))
End Function

Function AeroSpdSnd_mPerSec_fOatFahrenheit(oatFahrenheit)
    AeroSpdSnd_mPerSec_fOatFahrenheit = AeroSpdSnd_mPerSec_fOatKelvin(AeroConvFahrenheitToKelvin(oatFahrenheit))
End Function

Function AeroSpdSnd_ftPerSec_fOatRankine(oatRankine)
    AeroSpdSnd_ftPerSec_fOatRankine = AeroSpdSnd_ftPerSec_fOatKelvin(AeroConvRankineToKelvin(oatRankine))
End Function

Function AeroSpdSnd_kts_fOatRankine(oatRankine)
    AeroSpdSnd_kts_fOatRankine = AeroSpdSnd_kts_fOatKelvin(AeroConvRankineToKelvin(oatRankine))
End Function

Function AeroSpdSnd_statuteMilesPerHr_fOatRankine(oatRankine)
    AeroSpdSnd_statuteMilesPerHr_fOatRankine = AeroSpdSnd_statuteMilesPerHr_fOatKelvin(AeroConvRankineToKelvin(oatRankine))
End Function

Function AeroSpdSnd_kmPerHr_fOatRankine(oatRankine)
    AeroSpdSnd_kmPerHr_fOatRankine = AeroSpdSnd_kmPerHr_fOatKelvin(AeroConvRankineToKelvin(oatRankine))
End Function

Function AeroSpdSnd_mPerSec_fOatRankine(oatRankine)
    AeroSpdSnd_mPerSec_fOatRankine = AeroSpdSnd_mPerSec_fOatKelvin(AeroConvRankineToKelvin(oatRankine))
End Function

Function AeroSpdSndStdDay_ftPerSec_fHp(hp)
    AeroSpdSndStdDay_ftPerSec_fHp = AeroSpdSnd_ftPerSec_fOatKelvin(AeroOatStdDay_Kelvin_fHp(hp))
End Function

Function AeroSpdSndStdDay_kts_fHp(hp)
    AeroSpdSndStdDay_kts_fHp = AeroSpdSndStdDay_ftPerSec_fHp(hp) * constConvFtPerSecToKts
End Function

Function AeroSpdSndStdDay_statuteMilesPerHr_fHp(hp)
    AeroSpdSndStdDay_statuteMilesPerHr_fHp = AeroSpdSndStdDay_ftPerSec_fHp(hp) * constConvHrToSec / constConvStatuteMileToFt
End Function

Function AeroSpdSndStdDay_kmPerHr_fHp(hp)
    AeroSpdSndStdDay_kmPerHr_fHp = AeroSpdSndStdDay_ftPerSec_fHp(hp) * constConvHrToSec * constConvFtToM / 1000
End Function

Function AeroSpdSndStdDay_mPerSec_fHp(hp)
    AeroSpdSndStdDay_mPerSec_fHp = AeroSpdSndStdDay_ftPerSec_fHp(hp) * constConvFtToM
End Function

Function AeroOatStdDay_Kelvin_fHp(hp)
    If (hp < tropopause11kmInFt) Then 'Troposphere
        AeroOatStdDay_Kelvin_fHp = constTo + hp * constLapseTrop
    ElseIf (hp < topIsoThermLayerStrat20kmInFt) Then 'Isothermal layer of lower stratosphere
        AeroOatStdDay_Kelvin_fHp = constOatIsoLayerStrat11to20kmInK
    ElseIf (hp < toplstInverLayerStrat32kmInFt) Then 'First inversion layer of stratosphere
        AeroOatStdDay_Kelvin_fHp = constOatIsoLayerStrat11to20kmInK + (hp - topIsoThermLayerStrat20kmInFt) * constLapseStrat2
    ElseIf (hp < stratopauseStart47kmInFt) Then 'Second inversion layer of stratosphere
        AeroOatStdDay_Kelvin_fHp = constOatStrat32kmInK + (hp - toplstInverLayerStrat32kmInFt) * constLapseStrat3
    ElseIf (hp <= stratopauseEnd51kmInFt) Then 'Stratopause (isothermal layer 47 to 51 km)
        AeroOatStdDay_Kelvin_fHp = constOatStratopause47to51kmInK
    Else 'Outside bounds of this model - return error code
        AeroOatStdDay_Kelvin_fHp = constAeroErr
    End If
End Function

Function AeroOatStdDay_Celsius_fHp(hp)
    AeroOatStdDay_Celsius_fHp = AeroConvKelvinToCelsius(AeroOatStdDay_Kelvin_fHp(hp))
End Function

Function AeroOatStdDay_Fahrenheit_fHp(hp)
    AeroOatStdDay_Fahrenheit_fHp = AeroConvKelvinToFahrenheit(AeroOatStdDay_Kelvin_fHp(hp))
End Function

Function AeroOatStdDay_Rankine_fHp(hp)
    AeroOatStdDay_Rankine_fHp = AeroConvKelvinToRankine(AeroOatStdDay_Kelvin_fHp(hp))
End Function

Function AeroTapeAlt_ft_fGeoptlAlt(hGeopotlInFt)
    AeroTapeAlt_ft_fGeoptlAlt = hGeopotlInFt * (1 + hGeopotlInFt / (constRadiusEarth - hGeopotlInFt))
End Function

Function AeroTapeAlt_m_fGeoptlAlt(hGeopotlInM)
    AeroTapeAlt_m_fGeoptlAlt = AeroTapeAlt_ft_fGeoptlAlt(hGeopotlInM / constConvFtToM) * constConvFtToM
End Function

Function AeroGeoptlAlt_ft_fTapeAlt(hTapeAltInFt)
    AeroGeoptlAlt_ft_fTapeAlt = hTapeAltInFt * constRadiusEarth / (hTapeAltInFt + constRadiusEarth)
End Function

Function AeroGeoptlAlt_m_fTapeAlt(hTapeAltInM)
    AeroGeoptlAlt_m_fTapeAlt = AeroGeoptlAlt_ft_fTapeAlt(hTapeAltInM / constConvFtToM) * constConvFtToM
End Function

Function AeroThetaStdDay_fHp(hp)
    AeroThetaStdDay_fHp = AeroOatStdDay_Kelvin_fHp(hp) / constTo
End Function

Function AeroSigmaStdDay_fHp(hp)
    AeroSigmaStdDay_fHp = AeroDelta_fHp(hp) / AeroThetaStdDay_fHp(hp)
End Function

Function AeroDelta_fHp(hp)
    If (hp < tropopause11kmInFt) Then
        AeroDelta_fHp = (1 + hp * constLapseTrop / constTo) ^ (-constGo / constLapseTrop / constRAir)
    ElseIf (hp < topIsoThermLayerStrat20kmInFt) Then 'Isothermal layer of lower stratosphere
        AeroDelta_fHp = constDeltaTropopause * Exp(constGo / constRAir / constOatIsoLayerStrat11to20kmInK * (tropopause11kmInFt - hp))
    ElseIf (hp < toplstInverLayerStrat32kmInFt) Then 'First inversion layer of stratosphere
        AeroDelta_fHp = constDeltaStrat20km * (1 + (hp - topIsoThermLayerStrat20kmInFt) * constLapseStrat2 / constOatIsoLayerStrat11to20kmInK) ^ (-constGo / constLapseStrat2 / constRAir)
    ElseIf (hp < stratopauseStart47kmInFt) Then 'Second inversion layer of stratosphere
        AeroDelta_fHp = constDeltaStrat32km * (1 + (hp - toplstInverLayerStrat32kmInFt) * constLapseStrat3 / constOatStrat32kmInK) ^ (-constGo / constLapseStrat3 / constRAir)
    ElseIf (hp < stratopauseEnd51kmInFt) Then 'Stratopause (isothermal layer 47 to 51 km)
        AeroDelta_fHp = constDeltaStrat47km * Exp(constGo / constRAir / constOatStratopause47to51kmInK * (stratopauseStart47kmInFt - hp))
    Else 'Ouside bounds of this model - return error code
        AeroDelta_fHp = constAeroErr
    End If
End Function

Function AeroPstatic_lbfPerFt2_fHp(hp)
    AeroPstatic_lbfPerFt2_fHp = constPo * AeroDelta_fHp(hp)
End Function

Function AeroPstatic_inHg_fHp(hp)
    AeroPstatic_inHg_fHp = AeroPstatic_hPa_fHp(hp) / constConvInHgToHectoPascal
End Function

Function AeroPstatic_hPa_fHp(hp)
    AeroPstatic_hPa_fHp = AeroPstatic_lbfPerFt2_fHp(hp) / constConvFtToM ^ 2 * constConvLbfToNewton / 100
End Function

Function AeroPstatic_lbfPerInch2_fHp(hp)
    AeroPstatic_lbfPerInch2_fHp = AeroPstatic_lbfPerFt2_fHp(hp) / constConvLbfPerInch2ToLbfPerFt2
End Function

Function AeroRhoStdDay_slugPerFt3_fHp(hp)
    AeroRhoStdDay_slugPerFt3_fHp = constRhoo * AeroSigmaStdDay_fHp(hp)
End Function

Function AeroRho_slugPerFt3_fHpOatCelsius(hp, oatCelsius)
    AeroRho_slugPerFt3_fHpOatCelsius = constRhoo * AeroSigma_fOatCelsiusHp(oatCelsius, hp)
End Function

Function AeroRho_slugPerFt3_fHpOatKelvin(hp, oatKelvin)
    AeroRho_slugPerFt3_fHpOatKelvin = constRhoo * AeroSigma_fOatKelvinHp(oatKelvin, hp)
End Function

Function AeroRho_slugPerFt3_fHpOatRankine(hp, oatRankine)
    AeroRho_slugPerFt3_fHpOatRankine = constRhoo * AeroSigma_fOatRankineHp(oatRankine, hp)
End Function

Function AeroRho_slugPerFt3_fHpOatFahrenheit(hp, oatFahrenheit)
    AeroRho_slugPerFt3_fHpOatFahrenheit = constRhoo * AeroSigma_fOatFahrenheitHp(oatFahrenheit, hp)
End Function

Function AeroRho_slugPerFt3_fHpIsaDevFahrenheit(hp, isaDevFahrenheit)
    AeroRho_slugPerFt3_fHpIsaDevFahrenheit = constRhoo * AeroSigma_fIsaDevFahrenheitHp(isaDevFahrenheit, hp)
End Function

Function AeroRho_slugPerFt3_fHpIsaDevCelsius(hp, isaDevCelsius)
    AeroRho_slugPerFt3_fHpIsaDevCelsius = constRhoo * AeroSigma_fIsaDevCelsiusHp(isaDevCelsius, hp)
End Function

Function AeroConvKelvinToCelsius(oatKelvin)
    AeroConvKelvinToCelsius = oatKelvin - constZeroDegCelsiusInKelvin
End Function

Function AeroConvKelvinToRankine(oatKelvin)
    AeroConvKelvinToRankine = oatKelvin * constConvKelvinToRankine
End Function

Function AeroConvKelvinToFahrenheit(oatKelvin)
    AeroConvKelvinToFahrenheit = AeroConvKelvinToRankine(oatKelvin) - constZeroDegFahrenheitInRankine
End Function

Function AeroConvCelsiusToKelvin(oatCelsius)
    AeroConvCelsiusToKelvin = oatCelsius + constZeroDegCelsiusInKelvin
End Function

Function AeroConvCelsiusToRankine(oatCelsius)
    AeroConvCelsiusToRankine = AeroConvKelvinToRankine(AeroConvCelsiusToKelvin(oatCelsius))
End Function

Function AeroConvCelsiusToFahrenheit(oatCelsius)
    AeroConvCelsiusToFahrenheit = AeroConvCelsiusToRankine(oatCelsius) - constZeroDegFahrenheitInRankine
End Function

Function AeroConvRankineToKelvin(oatRankine)
    AeroConvRankineToKelvin = oatRankine / constConvKelvinToRankine
End Function

Function AeroConvRankineToCelsius(oatRankine)
    AeroConvRankineToCelsius = AeroConvRankineToKelvin(oatRankine) - constZeroDegCelsiusInKelvin
End Function

Function AeroConvRankineToFahrenheit(oatRankine)
    AeroConvRankineToFahrenheit = oatRankine - constZeroDegFahrenheitInRankine
End Function

Function AeroConvFahrenheitToKelvin(oatFahrenheit)
    AeroConvFahrenheitToKelvin = AeroConvFahrenheitToRankine(oatFahrenheit) / constConvKelvinToRankine
End Function

Function AeroConvFahrenheitToCelsius(oatFahrenheit)
    AeroConvFahrenheitToCelsius = (oatFahrenheit - 32) / constConvKelvinToRankine
End Function

Function AeroConvFahrenheitToRankine(oatFahrenheit)
    AeroConvFahrenheitToRankine = oatFahrenheit + constZeroDegFahrenheitInRankine
End Function

Function AeroConvKtsToFtPerSec(kts)
    AeroConvKtsToFtPerSec = kts / constConvFtPerSecToKts
End Function

Function AeroConvFtPerSecToKts(ftPerSec)
    AeroConvFtPerSecToKts = ftPerSec * constConvFtPerSecToKts
End Function

Function AeroConvKtsToMPerSec(kts)
    AeroConvKtsToMPerSec = kts / constConvFtPerSecToKts * constConvFtToM
End Function

Function AeroConvMPerSecToKts(mPerSec)
    AeroConvMPerSecToKts = mPerSec * constConvFtPerSecToKts / constConvFtToM
End Function

Function AeroConvMPerSecToFtPerSec(mPerSec)
    AeroConvMPerSecToFtPerSec = mPerSec / constConvFtToM
End Function

Function AeroConvFtPerSecToMPerSec(ftPerSec)
    AeroConvFtPerSecToMPerSec = ftPerSec * constConvFtToM
End Function

Function AeroConvSlugPerFt3ToKgPerM3(density)
    AeroConvSlugPerFt3ToKgPerM3 = density * constConvLbfToNewton / constConvFtToM ^ 4
End Function

Function AeroConvKgPerM3ToSlugPerFt3(density)
    AeroConvKgPerM3ToSlugPerFt3 = density / constConvLbfToNewton * constConvFtToM ^ 4
End Function

Function AeroConvLbfPerFt2ToLbfPerInch2(lbfPerFt2)
    AeroConvLbfPerFt2ToLbfPerInch2 = lbfPerFt2 / constConvLbfPerInch2ToLbfPerFt2
End Function
Function AeroConvLbfPerFt2ToInHg(lbfPerFt2)
    AeroConvLbfPerFt2ToInHg = AeroConvLbfPerFt2ToHPa(lbfPerFt2) / constConvInHgToHectoPascal
End Function
Function AeroConvLbfPerFt2ToHPa(lbfPerFt2)
    AeroConvLbfPerFt2ToHPa = lbfPerFt2 / constConvFtToM ^ 2 * constConvLbfToNewton / 100
End Function

Function AeroConvLbfPerInch2ToLbfPerFt2(lbfPerInch2)
    AeroConvLbfPerInch2ToLbfPerFt2 = lbfPerInch2 * constConvLbfPerInch2ToLbfPerFt2
End Function

Function AeroConvLbfPerInch2ToInHg(lbfPerInch2)
    AeroConvLbfPerInch2ToInHg = AeroConvLbfPerInch2ToHPa(lbfPerInch2) / constConvInHgToHectoPascal
End Function

Function AeroConvLbfPerInch2ToHPa(lbfPerInch2)
    AeroConvLbfPerInch2ToHPa = AeroConvLbfPerFt2ToHPa(lbfPerInch2 * constConvLbfPerInch2ToLbfPerFt2)
End Function

Function AeroConvInHgToLbfPerFt2(inHg)
    AeroConvInHgToLbfPerFt2 = AeroConvHPaToLbfPerFt2(AeroConvInHgToHPa(inHg))
End Function

Function AeroConvInHgToLbfPerInch2(inHg)
    AeroConvInHgToLbfPerInch2 = AeroConvHPaToLbfPerInch2((AeroConvInHgToHPa(inHg)))
End Function

Function AeroConvInHgToHPa(inHg)
    AeroConvInHgToHPa = inHg * constConvInHgToHectoPascal
End Function

Function AeroConvHPaToLbfPerFt2(hPa)
    AeroConvHPaToLbfPerFt2 = hPa * constConvFtToM ^ 2 / constConvLbfToNewton * 100
End Function

Function AeroConvHPaToLbfPerInch2(hPa)
    AeroConvHPaToLbfPerInch2 = AeroConvHPaToLbfPerFt2(hPa) / constConvLbfPerInch2ToLbfPerFt2
End Function

Function AeroConvHPaToInHg(hPa)
    AeroConvHPaToInHg = hPa / constConvInHgToHectoPascal
End Function

Function AeroConvDegToRad(deg)
    AeroConvDegToRad = deg * WorksheetFunction.Pi / 180
End Function

Function AeroConvRadToDeg(rad)
    AeroConvRadToDeg = rad / WorksheetFunction.Pi * 180
End Function

Function AeroTheta_fOatKelvin(oatKelvin)
    AeroTheta_fOatKelvin = oatKelvin / constTo
End Function

Function AeroTheta_fOatCelsius(oatCelsius)
    AeroTheta_fOatCelsius = AeroTheta_fOatKelvin(AeroConvCelsiusToKelvin(oatCelsius))
End Function

Function AeroTheta_fOatRankine(oatRankine)
    AeroTheta_fOatRankine = AeroTheta_fOatKelvin(AeroConvRankineToKelvin(oatRankine))
End Function

Function AeroTheta_fOatFahrenheit(oatFahrenheit)
    AeroTheta_fOatFahrenheit = AeroTheta_fOatKelvin(AeroConvFahrenheitToKelvin(oatFahrenheit))
End Function

Function AeroIsaDev_Celsius_fOatKelvinHp(oatKelvin, hp)
    AeroIsaDev_Celsius_fOatKelvinHp = oatKelvin - AeroOatStdDay_Kelvin_fHp(hp)
End Function

Function AeroIsaDev_Celsius_fOatCelsiusHp(oatCelsius, hp)
    AeroIsaDev_Celsius_fOatCelsiusHp = oatCelsius - AeroOatStdDay_Celsius_fHp(hp)
End Function

Function AeroIsaDev_Celsius_fOatFahrenheitHp(oatFahrenheit, hp)
    AeroIsaDev_Celsius_fOatFahrenheitHp = AeroConvFahrenheitToKelvin(oatFahrenheit) - AeroOatStdDay_Kelvin_fHp(hp)
End Function

Function AeroIsaDev_Celsius_fOatRankineHp(oatRankine, hp)
    AeroIsaDev_Celsius_fOatRankineHp = AeroConvRankineToKelvin(oatRankine) - AeroOatStdDay_Kelvin_fHp(hp)
End Function

Function AeroIsaDev_Fahrenheit_fOatKelvinHp(oatKelvin, hp)
    AeroIsaDev_Fahrenheit_fOatKelvinHp = AeroConvKelvinToRankine(oatKelvin) - AeroOatStdDay_Rankine_fHp(hp)
End Function

Function AeroIsaDev_Fahrenheit_fOatCelsiusHp(oatCelsius, hp)
    AeroIsaDev_Fahrenheit_fOatCelsiusHp = AeroConvCelsiusToRankine(oatCelsius) - AeroOatStdDay_Rankine_fHp(hp)
End Function

Function AeroIsaDev_Fahrenheit_fOatFahrenheitHp(oatFahrenheit, hp)
    AeroIsaDev_Fahrenheit_fOatFahrenheitHp = oatFahrenheit - AeroOatStdDay_Fahrenheit_fHp(hp)
End Function

Function AeroIsaDev_Fahrenheit_fOatRankineHp(oatRankine, hp)
    AeroIsaDev_Fahrenheit_fOatRankineHp = oatRankine - AeroOatStdDay_Rankine_fHp(hp)
End Function

Function AeroOat_Celsius_fIsaDevCelsiusHp(isaDevCelsius, hp)
    AeroOat_Celsius_fIsaDevCelsiusHp = AeroOatStdDay_Celsius_fHp(hp) + isaDevCelsius
End Function

Function AeroOat_Kelvin_fIsaDevCelsiusHp(isaDevCelsius, hp)
    AeroOat_Kelvin_fIsaDevCelsiusHp = AeroOatStdDay_Kelvin_fHp(hp) + isaDevCelsius
End Function

Function AeroOat_Fahrenheit_fIsaDevCelsiusHp(isaDevCelsius, hp)
    AeroOat_Fahrenheit_fIsaDevCelsiusHp = AeroOatStdDay_Fahrenheit_fHp(hp) + isaDevCelsius * constConvKelvinToRankine
End Function

Function AeroOat_Rankine_fIsaDevCelsiusHp(isaDevCelsius, hp)
    AeroOat_Rankine_fIsaDevCelsiusHp = AeroOatStdDay_Rankine_fHp(hp) + isaDevCelsius * constConvKelvinToRankine
End Function

Function AeroSpdSnd_kts_fIsaDevCelsiusHp(isaDevCelsius, hp)
    AeroSpdSnd_kts_fIsaDevCelsiusHp = AeroSpdSnd_kts_fOatKelvin(AeroOat_Kelvin_fIsaDevCelsiusHp(isaDevCelsius, hp))
End Function

Function AeroSpdSnd_statuteMilesPerHr_fIsaDevCelsiusHp(isaDevCelsius, hp)
    AeroSpdSnd_statuteMilesPerHr_fIsaDevCelsiusHp = AeroSpdSnd_statuteMilesPerHr_fOatKelvin(AeroOat_Kelvin_fIsaDevCelsiusHp(isaDevCelsius, hp))
End Function

Function AeroSpdSnd_kmPerHr_fIsaDevCelsiusHp(isaDevCelsius, hp)
    AeroSpdSnd_kmPerHr_fIsaDevCelsiusHp = AeroSpdSnd_kmPerHr_fOatKelvin(AeroOat_Kelvin_fIsaDevCelsiusHp(isaDevCelsius, hp))
End Function

Function AeroSpdSnd_ftPerSec_fIsaDevCelsiusHp(isaDevCelsius, hp)
    AeroSpdSnd_ftPerSec_fIsaDevCelsiusHp = AeroSpdSnd_ftPerSec_fOatKelvin(AeroOat_Kelvin_fIsaDevCelsiusHp(isaDevCelsius, hp))
End Function

Function AeroSpdSnd_mPerSec_fIsaDevCelsiusHp(isaDevCelsius, hp)
    AeroSpdSnd_mPerSec_fIsaDevCelsiusHp = AeroSpdSnd_mPerSec_fOatKelvin(AeroOat_Kelvin_fIsaDevCelsiusHp(isaDevCelsius, hp))
End Function

Function AeroTheta_fIsaDevCelsiusHp(isaDevCelsius, hp)
    AeroTheta_fIsaDevCelsiusHp = AeroOat_Kelvin_fIsaDevCelsiusHp(isaDevCelsius, hp) / constTo
End Function

Function AeroSigma_fIsaDevCelsiusHp(isaDevCelsius, hp)
    AeroSigma_fIsaDevCelsiusHp = AeroDelta_fHp(hp) / AeroTheta_fIsaDevCelsiusHp(isaDevCelsius, hp)
End Function

Function AeroIsaDev_Fahrenheit_fIsaDevCelsius(isaDevCelsius)
    AeroIsaDev_Fahrenheit_fIsaDevCelsius = isaDevCelsius * constConvKelvinToRankine
End Function

Function AeroOat_Celsius_fIsaDevFahrenheitHp(isaDevFahrenheit, hp)
    AeroOat_Celsius_fIsaDevFahrenheitHp = AeroOatStdDay_Celsius_fHp(hp) + isaDevFahrenheit / constConvKelvinToRankine
End Function

Function AeroOat_Kelvin_fIsaDevFahrenheitHp(isaDevFahrenheit, hp)
    AeroOat_Kelvin_fIsaDevFahrenheitHp = AeroOatStdDay_Kelvin_fHp(hp) + isaDevFahrenheit / constConvKelvinToRankine
End Function

Function AeroOat_Fahrenheit_fIsaDevFahrenheitHp(isaDevFahrenheit, hp)
    AeroOat_Fahrenheit_fIsaDevFahrenheitHp = AeroOatStdDay_Fahrenheit_fHp(hp) + isaDevFahrenheit
End Function

Function AeroOat_Rankine_fIsaDevFahrenheitHp(isaDevFahrenheit, hp)
    AeroOat_Rankine_fIsaDevFahrenheitHp = AeroOatStdDay_Rankine_fHp(hp) + isaDevFahrenheit
End Function

Function AeroSpdSnd_kts_fIsaDevFahrenheitHp(isaDevFahrenheit, hp)
    AeroSpdSnd_kts_fIsaDevFahrenheitHp = AeroSpdSnd_kts_fOatRankine(AeroOat_Rankine_fIsaDevFahrenheitHp(isaDevFahrenheit, hp))
End Function

Function AeroSpdSnd_statuteMilesPerHr_fIsaDevFahrenheitHp(isaDevFahrenheit, hp)
    AeroSpdSnd_statuteMilesPerHr_fIsaDevFahrenheitHp = AeroSpdSnd_statuteMilesPerHr_fOatRankine(AeroOat_Rankine_fIsaDevFahrenheitHp(isaDevFahrenheit, hp))
End Function

Function AeroSpdSnd_kmPerHr_fIsaDevFahrenheitHp(isaDevFahrenheit, hp)
    AeroSpdSnd_kmPerHr_fIsaDevFahrenheitHp = AeroSpdSnd_kmPerHr_fOatRankine(AeroOat_Rankine_fIsaDevFahrenheitHp(isaDevFahrenheit, hp))
End Function

Function AeroSpdSnd_ftPerSec_fIsaDevFahrenheitHp(isaDevFahrenheit, hp)
    AeroSpdSnd_ftPerSec_fIsaDevFahrenheitHp = AeroSpdSnd_ftPerSec_fOatRankine(AeroOat_Rankine_fIsaDevFahrenheitHp(isaDevFahrenheit, hp))
End Function

Function AeroSpdSnd_mPerSec_fIsaDevFahrenheitHp(isaDevFahrenheit, hp)
    AeroSpdSnd_mPerSec_fIsaDevFahrenheitHp = AeroSpdSnd_mPerSec_fOatRankine(AeroOat_Rankine_fIsaDevFahrenheitHp(isaDevFahrenheit, hp))
End Function

Function AeroTheta_fIsaDevFahrenheitHp(isaDevFahrenheit, hp)
    AeroTheta_fIsaDevFahrenheitHp = AeroOat_Rankine_fIsaDevFahrenheitHp(isaDevFahrenheit, hp) / constTo / constConvKelvinToRankine
End Function

Function AeroSigma_fIsaDevFahrenheitHp(isaDevFahrenheit, hp)
    AeroSigma_fIsaDevFahrenheitHp = AeroDelta_fHp(hp) / AeroTheta_fIsaDevFahrenheitHp(isaDevFahrenheit, hp)
End Function

Function AeroIsaDev_Celsius_fIsaDevFahrenheitHp(isaDevFahrenheit)
    AeroIsaDev_Celsius_fIsaDevFahrenheitHp = isaDevFahrenheit / constConvKelvinToRankine
End Function

Function AeroSigma_fOatCelsiusHp(oatCelsius, hp)
    AeroSigma_fOatCelsiusHp = AeroDelta_fHp(hp) / AeroTheta_fOatCelsius(oatCelsius)
End Function

Function AeroSigma_fOatKelvinHp(oatKelvin, hp)
    AeroSigma_fOatKelvinHp = AeroDelta_fHp(hp) / AeroTheta_fOatKelvin(oatKelvin)
End Function

Function AeroSigma_fOatFahrenheitHp(oatFahrenheit, hp)
    AeroSigma_fOatFahrenheitHp = AeroDelta_fHp(hp) / AeroTheta_fOatFahrenheit(oatFahrenheit)
End Function

Function AeroSigma_fOatRankineHp(oatRankine, hp)
    AeroSigma_fOatRankineHp = AeroDelta_fHp(hp) / AeroTheta_fOatRankine(oatRankine)
End Function

Function AeroMach_fHpKcas(hp, kcas)
    deltaPressureStatic = AeroDelta_fHp(hp)
    mach = (2 / (constGammaAir - 1) * ((1 / deltaPressureStatic * ((1 + (constGammaAir - 1) / 2 * (kcas / constAo) ^ 2) _
        ^ (constGammaAir / (constGammaAir - 1)) - 1) + 1) ^ ((constGammaAir - 1) / constGammaAir) - 1)) ^ 0.5
    If mach > 1 Then
        pStatic = deltaPressureStatic * constPo
        qc = AeroQc_lbfPerFt2_fKcas(kcas)
        Do
            machLast = mach
            mach = (2 / (constGammaAir + 1) * (qc / pStatic + 1) ^ ((constGammaAir - 1) / constGammaAir) * _
            ((1 - constGammaAir + 2 * constGammaAir * machLast ^ 2) / (constGammaAir + 1)) ^ (1 / constGammaAir)) ^ 0.5
        Loop While Abs(mach - machLast) > constMachEpsilon
    End If
    AeroMach_fHpKcas = mach
End Function

Function AeroMach_fHpQc(hp, qc)
    deltaPressureStatic = AeroDelta_fHp(hp)
    pStatic = deltaPressureStatic * constPo
    mach = (2 / (constGammaAir - 1) * ((qc / pStatic + 1) ^ ((constGammaAir - 1) / constGammaAir) - 1)) ^ 0.5
    If mach > 1 Then
        Do
            machLast = mach
            mach = (2 / (constGammaAir + 1) * (qc / pStatic + 1) ^ ((constGammaAir - 1) / constGammaAir) * _
            ((1 - constGammaAir + 2 * constGammaAir * machLast ^ 2) / (constGammaAir + 1)) ^ (1 / constGammaAir)) ^ 0.5
        Loop While Abs(mach - machLast) > constMachEpsilon
    End If
    AeroMach_fHpQc = mach
End Function

Function AeroMach_fHpKeas(hp, keas)
    AeroMach_fHpKeas = AeroConvKtsToFtPerSec(keas) * (constRhoo / constGammaAir / AeroPstatic_lbfPerFt2_fHp(hp)) ^ 0.5
End Function

Function AeroMach_fHpQ(hp, q)
    AeroMach_fHpQ = (2 * q / (constGammaAir * constPo * AeroDelta_fHp(hp))) ^ 0.5
End Function

Function AeroMachStdDay_fHpKtas(hp, ktas)
    AeroMachStdDay_fHpKtas = ktas / AeroSpdSndStdDay_kts_fHp(hp)
End Function

Function AeroKtasStdDay_fHpKcas(hp, kcas)
    AeroKtasStdDay_fHpKcas = AeroSpdSndStdDay_kts_fHp(hp) * AeroMach_fHpKcas(hp, kcas)
End Function

Function AeroKtasStdDay_fHpKeas(hp, keas)
    AeroKtasStdDay_fHpKeas = keas / (AeroSigmaStdDay_fHp(hp)) ^ 0.5
End Function

Function AeroKtasStdDay_fHpMach(hp, mach)
    AeroKtasStdDay_fHpMach = AeroSpdSndStdDay_kts_fHp(hp) * mach
End Function

Function AeroKtasStdDay_fHpQ(hp, q)
    AeroKtasStdDay_fHpQ = constConvFtPerSecToKts * (2 * q / (AeroRhoStdDay_slugPerFt3_fHp(hp))) ^ 0.5
End Function

Function AeroKtas_fHpQOatCelsius(hp, q, oatCelsius)
    AeroKtas_fHpQOatCelsius = AeroMach_fHpQ(hp, q) * AeroSpdSnd_kts_fOatCelsius(oatCelsius)
End Function

Function AeroKtas_fHpQOatKelvin(hp, q, oatKelvin)
    AeroKtas_fHpQOatKelvin = AeroMach_fHpQ(hp, q) * AeroSpdSnd_kts_fOatKelvin(oatKelvin)
End Function

Function AeroKtas_fHpQOatFahrenheit(hp, q, oatFahrenheit)
    AeroKtas_fHpQOatFahrenheit = AeroMach_fHpQ(hp, q) * AeroSpdSnd_kts_fOatFahrenheit(oatFahrenheit)
End Function

Function AeroKtas_fHpQOatRankine(hp, q, oatRankine)
    AeroKtas_fHpQOatRankine = AeroMach_fHpQ(hp, q) * AeroSpdSnd_kts_fOatRankine(oatRankine)
End Function

Function AeroKtas_fHpQIsaDevCelsius(hp, q, isaDevCelsius)
    AeroKtas_fHpQIsaDevCelsius = AeroMach_fHpQ(hp, q) * AeroSpdSnd_kts_fIsaDevCelsiusHp(isaDevCelsius, hp)
End Function

Function AeroKtas_fHpQIsaDevFahrenheit(hp, q, isaDevFahrenheit)
    AeroKtas_fHpQIsaDevFahrenheit = AeroMach_fHpQ(hp, q) * AeroSpdSnd_kts_fIsaDevFahrenheitHp(isaDevFahrenheit, hp)
End Function

Function AeroKtasStdDay_fHpQc(hp, qc)
    AeroKtasStdDay_fHpQc = AeroMach_fHpQc(hp, qc) * AeroSpdSndStdDay_kts_fHp(hp)
End Function

Function AeroKtas_fHpQcOatCelsius(hp, qc, oatCelsius)
    AeroKtas_fHpQcOatCelsius = AeroMach_fHpQc(hp, qc) * AeroSpdSnd_kts_fOatCelsius(oatCelsius)
End Function

Function AeroKtas_fHpQcOatFahrenheit(hp, qc, oatFahrenheit)
    AeroKtas_fHpQcOatFahrenheit = AeroMach_fHpQc(hp, qc) * AeroSpdSnd_kts_fOatFahrenheit(oatFahrenheit)
End Function

Function AeroKtas_fHpQcOatKelvin(hp, qc, oatKelvin)
    AeroKtas_fHpQcOatKelvin = AeroMach_fHpQc(hp, qc) * AeroSpdSnd_kts_fOatKelvin(oatKelvin)
End Function

Function AeroKtas_fHpQcOatRankine(hp, qc, oatRankine)
    AeroKtas_fHpQcOatRankine = AeroMach_fHpQc(hp, qc) * AeroSpdSnd_kts_fOatRankine(oatRankine)
End Function

Function AeroKtas_fHpQcIsaDevCelsius(hp, qc, isaDevCelsius)
    AeroKtas_fHpQcIsaDevCelsius = AeroMach_fHpQc(hp, qc) * AeroSpdSnd_kts_fIsaDevCelsiusHp(isaDevCelsius, hp)
End Function

Function AeroKtas_fHpQcIsaDevFahrenheit(hp, qc, isaDevFahrenheit)
    AeroKtas_fHpQcIsaDevFahrenheit = AeroMach_fHpQc(hp, qc) * AeroSpdSnd_kts_fIsaDevFahrenheitHp(isaDevFahrenheit, hp)
End Function

Function AeroKeas_fHpKcas(hp, kcas)
    mach = AeroMach_fHpKcas(hp, kcas)
    AeroKeas_fHpKcas = AeroKeas_fHpMach(hp, mach)
End Function

Function AeroKeasStdDay_fHpKtas(hp, ktas)
    AeroKeasStdDay_fHpKtas = ktas * (AeroSigmaStdDay_fHp(hp)) ^ 0.5
End Function

Function AeroKeas_fHpMach(hp, mach)
    AeroKeas_fHpMach = AeroKtasStdDay_fHpMach(hp, mach) * AeroSigmaStdDay_fHp(hp) ^ 0.5
End Function

Function AeroKeas_fQ(q)
    AeroKeas_fQ = constConvFtPerSecToKts * (2 * q / constRhoo) ^ 0.5
End Function

Function AeroKeas_fHpQc(hp, qc)
    mach = AeroMach_fHpQc(hp, qc)
    AeroKeas_fHpQc = AeroKeas_fHpMach(hp, mach)
End Function

Function AeroQc_lbfPerFt2_fKcas(kcas)
    If kcas < constAo Then
        AeroQc_lbfPerFt2_fKcas = constPo * ((1 + (constGammaAir - 1) / 2 * (kcas / constAo) ^ 2) ^ (constGammaAir / (constGammaAir - 1)) - 1)
    Else
        AeroQc_lbfPerFt2_fKcas = constPo * (((constGammaAir + 1) / 2 * (kcas / constAo) ^ 2) ^ (constGammaAir / (constGammaAir - 1)) * _
        ((constGammaAir + 1) / (1 - constGammaAir + 2 * constGammaAir * (kcas / constAo) ^ 2)) ^ (1 / (constGammaAir - 1)) - 1)
    End If
End Function

Function AeroQc_lbfPerFt2_fHpMach(hp, mach)
    If mach < 1 Then
        AeroQc_lbfPerFt2_fHpMach = AeroPstatic_lbfPerFt2_fHp(hp) * ((1 + (constGammaAir - 1) / 2 * mach ^ 2) ^ (constGammaAir / (constGammaAir - 1)) - 1)
    Else
        AeroQc_lbfPerFt2_fHpMach = AeroPstatic_lbfPerFt2_fHp(hp) * (((constGammaAir + 1) / 2 * mach ^ 2) ^ (constGammaAir / (constGammaAir - 1)) * _
        ((constGammaAir + 1) / (1 - constGammaAir + 2 * constGammaAir * mach ^ 2)) ^ (1 / (constGammaAir - 1)) - 1)
    End If
End Function

Function AeroQc_lbfPerFt2_fHpKeas(hp, keas)
    mach = AeroMach_fHpKeas(hp, keas)
    AeroQc_lbfPerFt2_fHpKeas = AeroQc_lbfPerFt2_fHpMach(hp, mach)
End Function

Function AeroQcStdDay_lbfPerFt2_fHpKtas(hp, ktas)
    mach = ktas / AeroSpdSndStdDay_kts_fHp(hp)
    AeroQcStdDay_lbfPerFt2_fHpKtas = AeroQc_lbfPerFt2_fHpMach(hp, mach)
End Function
    
Function AeroQc_lbfPerFt2_fHpQ(hp, q)
    mach = AeroMach_fHpQ(hp, q)
    AeroQc_lbfPerFt2_fHpQ = AeroQc_lbfPerFt2_fHpMach(hp, mach)
End Function

Function AeroQ_lbfPerFt2_fHpKcas(hp, kcas)
    mach = AeroMach_fHpKcas(hp, kcas)
    AeroQ_lbfPerFt2_fHpKcas = AeroQ_lbfPerFt2_fHpMach(hp, mach)
End Function

Function AeroQ_lbfPerFt2_fKeas(keas)
    AeroQ_lbfPerFt2_fKeas = 0.5 * constRhoo * AeroConvKtsToFtPerSec(keas) ^ 2
End Function

Function AeroQStdDay_lbfPerFt2_fHpKtas(hp, ktas)
    mach = ktas / AeroSpdSndStdDay_kts_fHp(hp)
    AeroQStdDay_lbfPerFt2_fHpKtas = AeroQ_lbfPerFt2_fHpMach(hp, mach)
End Function

Function AeroQ_lbfPerFt2_fHpMach(hp, mach)
    AeroQ_lbfPerFt2_fHpMach = constGammaAir / 2 * AeroDelta_fHp(hp) * constPo * mach ^ 2
End Function

Function AeroQ_lbfPerFt2_fHpQc(hp, qc)
    mach = AeroMach_fHpQc(hp, qc)
    AeroQ_lbfPerFt2_fHpQc = AeroQ_lbfPerFt2_fHpMach(hp, mach)
End Function

Function AeroSubsonicKeasOverKcas_fHpQc(hp, qc)
    pStatic = AeroPstatic_lbfPerFt2_fHp(hp)
    'Ref Herrington, Ch. 1, eqn 2.18:
    AeroSubsonicKeasOverKcas_fHpQc = ((((qc / pStatic + 1) ^ ((constGammaAir - 1) / constGammaAir) - 1) / _
    ((qc / constPo + 1) ^ ((constGammaAir - 1) / constGammaAir) - 1)) * pStatic / constPo) ^ 0.5
End Function

Function AeroCompCorrnKcasMinusKeas_fHpKeas(hp, keas)
    AeroCompCorrnKcasMinusKeas_fHpKeas = AeroKcas_fHpKeas(hp, keas) - keas
End Function

Function AeroCompCorrnKcasMinusKeas_fHpKcas(hp, kcas)
    AeroCompCorrnKcasMinusKeas_fHpKcas = kcas - AeroKeas_fHpKcas(hp, kcas)
End Function

Function AeroKcas_fHpKeas(hp, keas)
    ktas = AeroKtasStdDay_fHpKeas(hp, keas)
    AeroKcas_fHpKeas = AeroKcasStdDay_fHpKtas(hp, ktas)
End Function

Function AeroKcasStdDay_fHpKtas(hp, ktas)
    mach = AeroMachStdDay_fHpKtas(hp, ktas)
    AeroKcasStdDay_fHpKtas = AeroKcas_fHpMach(hp, mach)
End Function

Function AeroKcas_fHpMach(hp, mach)
    qc = AeroQc_lbfPerFt2_fHpMach(hp, mach)
    AeroKcas_fHpMach = AeroKcas_fQc(qc)
End Function

Function AeroKcas_fQc(qc)
    kcas = constAo * (2 / (constGammaAir - 1) * ((qc / constPo + 1) ^ ((constGammaAir - 1) / constGammaAir) - 1)) ^ 0.5
    If (kcas > constAo) Then
        Count = 0 'Counter included as safety mechanism to prevent infinite loop in event of unconverged solution
        kcasLast = kcas
        firstConstCoeff = ((constGammaAir + 1) / 2) ^ (0.5 * constGammaAir / (1 - constGammaAir)) * ((constGammaAir + 1) / 2 / constGammaAir) ^ (0.5 / (1 - constGammaAir))
        secondConstCoeff = 2 * constGammaAir / (constGammaAir - 1)
        Do
            kcasLast = kcas
            kcas = constAo * firstConstCoeff * ((qc / constPo + 1) * (1 - 1 / (secondConstCoeff * (kcasLast / constAo) ^ 2)) ^ (1 / (constGammaAir - 1))) ^ 0.5
            If (Count > constMaxAeroIter) Then
                'Set both kcas and kcasLast equal to constAeroErr; this will terminate loop and result in return of error-flag value
                kcas = constAeroErr
                kcasLast = constAeroErr
            End If
        Loop While (Abs(kcas - kcasLast) > constKcasEpsilon)
    End If
    AeroKcas_fQc = kcas
End Function

Function AeroKcas_fHpQ(hp, q)
    mach = AeroMach_fHpQ(hp, q)
    AeroKcas_fHpQ = AeroKcas_fHpMach(hp, mach)
End Function

Function AeroViscDyn_kgPerMSec_fTheta(theta)
  t = theta * constTo
  AeroViscDyn_kgPerMSec_fTheta = constBetaVisc * (t * t * t) ^ 0.5 / (t + constSuth) 'Dynamic viscosity, mu, in kg/(m-sec)
End Function

Function AeroViscDyn_SlugPerFtSec_fTheta(theta)
  AeroViscDyn_SlugPerFtSec_fTheta = AeroViscDyn_kgPerMSec_fTheta(theta) / (constConvLbfToNewton / constConvFtToM ^ 2)
End Function

Function AeroViscKin_M2PerSec_fSigmaTheta(sigma, theta)
  AeroViscKin_M2PerSec_fSigmaTheta = AeroViscDyn_kgPerMSec_fTheta(theta) / (sigma * constRhoo * constConvLbfToNewton / constConvFtToM ^ 4)
End Function

Function AeroViscKin_ft2PerSec_fSigmaTheta(sigma, theta)
  AeroViscKin_ft2PerSec_fSigmaTheta = AeroViscDyn_SlugPerFtSec_fTheta(theta) / (sigma * constRhoo)
End Function

Function AeroRePerFtStdDay_fHpMach(hp, mach)
    v = mach * AeroSpdSndStdDay_ftPerSec_fHp(hp)
    kinVisc = AeroViscKin_ft2PerSec_fSigmaTheta(AeroSigmaStdDay_fHp(hp), AeroThetaStdDay_fHp(hp))
    AeroRePerFtStdDay_fHpMach = v / kinVisc
End Function

Function AeroRePerFt_fHpMachOatCelsius(hp, mach, oatCelsius)
    v = mach * AeroSpdSnd_ftPerSec_fOatCelsius(oatCelsius)
    kinVisc = AeroViscKin_ft2PerSec_fSigmaTheta(AeroSigma_fOatCelsiusHp(oatCelsius, hp), AeroTheta_fOatCelsius(oatCelsius))
    AeroRePerFt_fHpMachOatCelsius = v / kinVisc
End Function

Function AeroRePerFt_fHpMachOatKelvin(hp, mach, oatKelvin)
    v = mach * AeroSpdSnd_ftPerSec_fOatKelvin(oatKelvin)
    kinVisc = AeroViscKin_ft2PerSec_fSigmaTheta(AeroSigma_fOatKelvinHp(oatKelvin, hp), AeroTheta_fOatKelvin(oatKelvin))
    AeroRePerFt_fHpMachOatKelvin = v / kinVisc
End Function

Function AeroRePerFt_fHpMachOatFahrenheit(hp, mach, oatFahrenheit)
    v = mach * AeroSpdSnd_ftPerSec_fOatFahrenheit(oatFahrenheit)
    kinVisc = AeroViscKin_ft2PerSec_fSigmaTheta(AeroSigma_fOatFahrenheitHp(oatFahrenheit, hp), AeroTheta_fOatFahrenheit(oatFahrenheit))
    AeroRePerFt_fHpMachOatFahrenheit = v / kinVisc
End Function

Function AeroRePerFt_fHpMachOatRankine(hp, mach, oatRankine)
    v = mach * AeroSpdSnd_ftPerSec_fOatRankine(oatRankine)
    kinVisc = AeroViscKin_ft2PerSec_fSigmaTheta(AeroSigma_fOatRankineHp(oatRankine, hp), AeroTheta_fOatRankine(oatRankine))
    AeroRePerFt_fHpMachOatRankine = v / kinVisc
End Function

Function AeroRePerFt_fHpMachIsaDevCelsius(hp, mach, isaDevCelsius)
    v = mach * AeroSpdSnd_ftPerSec_fIsaDevCelsiusHp(isaDevCelsius, hp)
    kinVisc = AeroViscKin_ft2PerSec_fSigmaTheta(AeroSigma_fIsaDevCelsiusHp(isaDevCelsius, hp), AeroTheta_fIsaDevCelsiusHp(isaDevCelsius, hp))
    AeroRePerFt_fHpMachIsaDevCelsius = v / kinVisc
End Function

Function AeroRePerFt_fHpMachIsaDevFahrenheit(hp, mach, isaDevFahrenheit)
    v = mach * AeroSpdSnd_ftPerSec_fIsaDevFahrenheitHp(isaDevFahrenheit, hp)
    kinVisc = AeroViscKin_ft2PerSec_fSigmaTheta(AeroSigma_fIsaDevFahrenheitHp(isaDevFahrenheit, hp), AeroTheta_fIsaDevFahrenheitHp(isaDevFahrenheit, hp))
    AeroRePerFt_fHpMachIsaDevFahrenheit = v / kinVisc
End Function

Function AeroKtas_fHpKcasOatCelsius(hp, kcas, oatCelsius)
    mach = AeroMach_fHpKcas(hp, kcas)
    AeroKtas_fHpKcasOatCelsius = AeroKtas_fMachOatCelsius(mach, oatCelsius)
End Function

Function AeroKtas_fHpKcasOatKelvin(hp, kcas, oatKelvin)
    mach = AeroMach_fHpKcas(hp, kcas)
    AeroKtas_fHpKcasOatKelvin = AeroKtas_fMachOatKelvin(mach, oatKelvin)
End Function

Function AeroKtas_fHpKcasOatFahrenheit(hp, kcas, oatFahrenheit)
    mach = AeroMach_fHpKcas(hp, kcas)
    AeroKtas_fHpKcasOatFahrenheit = AeroKtas_fMachOatFahrenheit(mach, oatFahrenheit)
End Function

Function AeroKtas_fHpKcasIsaDevCelsius(hp, kcas, isaDevCelsius)
    mach = AeroMach_fHpKcas(hp, kcas)
    oatKelvin = AeroOatStdDay_Kelvin_fHp(hp) + isaDevCelsius
    AeroKtas_fHpKcasIsaDevCelsius = AeroKtas_fMachOatKelvin(mach, oatKelvin)
End Function

Function AeroKtas_fHpKcasIsaDevFahrenheit(hp, kcas, isaDevFahrenheit)
    mach = AeroMach_fHpKcas(hp, kcas)
    oatRankine = AeroOatStdDay_Rankine_fHp(hp) + isaDevFahrenheit
    AeroKtas_fHpKcasIsaDevFahrenheit = AeroKtas_fMachOatRankine(mach, oatRankine)
End Function

Function AeroKtas_fHpKcasOatRankine(hp, kcas, oatRankine)
    mach = AeroMach_fHpKcas(hp, kcas)
    AeroKtas_fHpKcasOatRankine = AeroKtas_fMachOatRankine(mach, oatRankine)
End Function

Function AeroKtas_fHpKeasOatCelsius(hp, keas, oatCelsius)
    mach = AeroMach_fHpKeas(hp, keas)
    AeroKtas_fHpKeasOatCelsius = AeroKtas_fMachOatCelsius(mach, oatCelsius)
End Function

Function AeroKtas_fHpKeasOatKelvin(hp, keas, oatKelvin)
    mach = AeroMach_fHpKeas(hp, keas)
    AeroKtas_fHpKeasOatKelvin = AeroKtas_fMachOatKelvin(mach, oatKelvin)
End Function

Function AeroKtas_fHpKeasOatFahrenheit(hp, keas, oatFahrenheit)
    mach = AeroMach_fHpKeas(hp, keas)
    AeroKtas_fHpKeasOatFahrenheit = AeroKtas_fMachOatFahrenheit(mach, oatFahrenheit)
End Function

Function AeroKtas_fHpKeasIsaDevCelsius(hp, keas, isaDevCelsius)
    mach = AeroMach_fHpKeas(hp, keas)
    oatKelvin = AeroOatStdDay_Kelvin_fHp(hp) + isaDevCelsius
    AeroKtas_fHpKeasIsaDevCelsius = AeroKtas_fMachOatKelvin(mach, oatKelvin)
End Function

Function AeroKtas_fHpKeasIsaDevFahrenheit(hp, keas, isaDevFahrenheit)
    mach = AeroMach_fHpKeas(hp, keas)
    oatRankine = AeroOatStdDay_Rankine_fHp(hp) + isaDevFahrenheit
    AeroKtas_fHpKeasIsaDevFahrenheit = AeroKtas_fMachOatRankine(mach, oatRankine)
End Function

Function AeroKtas_fHpKeasOatRankine(hp, keas, oatRankine)
    mach = AeroMach_fHpKeas(hp, keas)
    AeroKtas_fHpKeasOatRankine = AeroKtas_fMachOatRankine(mach, oatRankine)
End Function

Function AeroKtas_fMachOatCelsius(mach, oatCelsius)
    AeroKtas_fMachOatCelsius = mach * AeroSpdSnd_kts_fOatCelsius(oatCelsius)
End Function

Function AeroKtas_fMachOatKelvin(mach, oatKelvin)
    AeroKtas_fMachOatKelvin = mach * AeroSpdSnd_kts_fOatKelvin(oatKelvin)
End Function

Function AeroKtas_fMachOatFahrenheit(mach, oatFahrenheit)
    AeroKtas_fMachOatFahrenheit = mach * AeroSpdSnd_kts_fOatFahrenheit(oatFahrenheit)
End Function

Function AeroKtas_fMachOatRankine(mach, oatRankine)
    AeroKtas_fMachOatRankine = mach * AeroSpdSnd_kts_fOatRankine(oatRankine)
End Function

Function AeroKtas_fMachHpIsaDevCelsius(mach, hp, isaDevCelsius)
    AeroKtas_fMachHpIsaDevCelsius = mach * AeroSpdSnd_kts_fOatKelvin(AeroOatStdDay_Kelvin_fHp(hp) + isaDevCelsius)
End Function

Function AeroKtas_fMachHpIsaDevFahrenheit(mach, hp, isaDevFahrenheit)
    AeroKtas_fMachHpIsaDevFahrenheit = mach * AeroSpdSnd_kts_fOatRankine(AeroOatStdDay_Rankine_fHp(hp) + isaDevFahrenheit)
End Function

Function AeroMach_fKtasOatCelsius(ktas, oatCelsius)
    AeroMach_fKtasOatCelsius = ktas / AeroSpdSnd_kts_fOatCelsius(oatCelsius)
End Function

Function AeroMach_fKtasOatKelvin(ktas, oatKelvin)
    AeroMach_fKtasOatKelvin = ktas / AeroSpdSnd_kts_fOatKelvin(oatKelvin)
End Function

Function AeroMach_fKtasOatRankine(ktas, oatRankine)
    AeroMach_fKtasOatRankine = ktas / AeroSpdSnd_kts_fOatRankine(oatRankine)
End Function

Function AeroMach_fKtasOatFahrenheit(ktas, oatFahrenheit)
    AeroMach_fKtasOatFahrenheit = ktas / AeroSpdSnd_kts_fOatFahrenheit(oatFahrenheit)
End Function

Function AeroMach_fHpKtasIsaDevCelsius(hp, ktas, isaDevCelsius)
    oatCelsius = AeroOatStdDay_Celsius_fHp(hp) + isaDevCelsius
    AeroMach_fHpKtasIsaDevCelsius = ktas / AeroSpdSnd_kts_fOatCelsius(oatCelsius)
End Function

Function AeroMach_fHpKtasIsaDevFahrenheit(hp, ktas, isaDevFahrenheit)
    oatFahrenheit = AeroOatStdDay_Fahrenheit_fHp(hp) + isaDevFahrenheit
    AeroMach_fHpKtasIsaDevFahrenheit = ktas / AeroSpdSnd_kts_fOatFahrenheit(oatFahrenheit)
End Function

Function AeroKeas_fHpKtasOatCelsius(hp, ktas, oatCelsius)
    mach = AeroMach_fKtasOatCelsius(ktas, oatCelsius)
    AeroKeas_fHpKtasOatCelsius = AeroKeas_fHpMach(hp, mach)
End Function

Function AeroKeas_fHpKtasOatKelvin(hp, ktas, oatKelvin)
    mach = AeroMach_fKtasOatKelvin(ktas, oatKelvin)
    AeroKeas_fHpKtasOatKelvin = AeroKeas_fHpMach(hp, mach)
End Function

Function AeroKeas_fHpKtasOatRankine(hp, ktas, oatRankine)
    mach = AeroMach_fKtasOatRankine(ktas, oatRankine)
    AeroKeas_fHpKtasOatRankine = AeroKeas_fHpMach(hp, mach)
End Function

Function AeroKeas_fHpKtasOatFahrenheit(hp, ktas, oatFahrenheit)
    mach = AeroMach_fKtasOatFahrenheit(ktas, oatFahrenheit)
    AeroKeas_fHpKtasOatFahrenheit = AeroKeas_fHpMach(hp, mach)
End Function

Function AeroKeas_fHpKtasIsaDevCelsius(hp, ktas, isaDevCelsius)
    mach = AeroMach_fHpKtasIsaDevCelsius(hp, ktas, isaDevCelsius)
    AeroKeas_fHpKtasIsaDevCelsius = AeroKeas_fHpMach(hp, mach)
End Function

Function AeroKeas_fHpKtasIsaDevFahrenheit(hp, ktas, isaDevFahrenheit)
    mach = AeroMach_fHpKtasIsaDevFahrenheit(hp, ktas, isaDevFahrenheit)
    AeroKeas_fHpKtasIsaDevFahrenheit = AeroKeas_fHpMach(hp, mach)
End Function

Function AeroKcas_fHpKtasOatCelsius(hp, ktas, oatCelsius)
    mach = AeroMach_fKtasOatCelsius(ktas, oatCelsius)
    AeroKcas_fHpKtasOatCelsius = AeroKcas_fHpMach(hp, mach)
End Function

Function AeroKcas_fHpKtasOatKelvin(hp, ktas, oatKelvin)
    mach = AeroMach_fKtasOatKelvin(ktas, oatKelvin)
    AeroKcas_fHpKtasOatKelvin = AeroKcas_fHpMach(hp, mach)
End Function

Function AeroKcas_fHpKtasOatRankine(hp, ktas, oatRankine)
    mach = AeroMach_fKtasOatRankine(ktas, oatRankine)
    AeroKcas_fHpKtasOatRankine = AeroKcas_fHpMach(hp, mach)
End Function

Function AeroKcas_fHpKtasOatFahrenheit(hp, ktas, oatFahrenheit)
    mach = AeroMach_fKtasOatFahrenheit(ktas, oatFahrenheit)
    AeroKcas_fHpKtasOatFahrenheit = AeroKcas_fHpMach(hp, mach)
End Function

Function AeroKcas_fHpKtasIsaDevCelsius(hp, ktas, isaDevCelsius)
    mach = AeroMach_fHpKtasIsaDevCelsius(hp, ktas, isaDevCelsius)
    AeroKcas_fHpKtasIsaDevCelsius = AeroKcas_fHpMach(hp, mach)
End Function

Function AeroKcas_fHpKtasIsaDevFahrenheit(hp, ktas, isaDevFahrenheit)
    mach = AeroMach_fHpKtasIsaDevFahrenheit(hp, ktas, isaDevFahrenheit)
    AeroKcas_fHpKtasIsaDevFahrenheit = AeroKcas_fHpMach(hp, mach)
End Function

Function AeroQ_lbfPerFt2_fHpKtasOatCelsius(hp, ktas, oatCelsius)
    mach = AeroMach_fKtasOatCelsius(ktas, oatCelsius)
    AeroQ_lbfPerFt2_fHpKtasOatCelsius = AeroQ_lbfPerFt2_fHpMach(hp, mach)
End Function

Function AeroQ_lbfPerFt2_fHpKtasOatKelvin(hp, ktas, oatKelvin)
    mach = AeroMach_fKtasOatKelvin(ktas, oatKelvin)
    AeroQ_lbfPerFt2_fHpKtasOatKelvin = AeroQ_lbfPerFt2_fHpMach(hp, mach)
End Function

Function AeroQ_lbfPerFt2_fHpKtasOatRankine(hp, ktas, oatRankine)
    mach = AeroMach_fKtasOatRankine(ktas, oatRankine)
    AeroQ_lbfPerFt2_fHpKtasOatRankine = AeroQ_lbfPerFt2_fHpMach(hp, mach)
End Function

Function AeroQ_lbfPerFt2_fHpKtasOatFahrenheit(hp, ktas, oatFahrenheit)
    mach = AeroMach_fKtasOatFahrenheit(ktas, oatFahrenheit)
    AeroQ_lbfPerFt2_fHpKtasOatFahrenheit = AeroQ_lbfPerFt2_fHpMach(hp, mach)
End Function

Function AeroQ_lbfPerFt2_fHpKtasIsaDevCelsius(hp, ktas, isaDevCelsius)
    mach = AeroMach_fHpKtasIsaDevCelsius(hp, ktas, isaDevCelsius)
    AeroQ_lbfPerFt2_fHpKtasIsaDevCelsius = AeroQ_lbfPerFt2_fHpMach(hp, mach)
End Function

Function AeroQ_lbfPerFt2_fHpKtasIsaDevFahrenheit(hp, ktas, isaDevFahrenheit)
    mach = AeroMach_fHpKtasIsaDevFahrenheit(hp, ktas, isaDevFahrenheit)
    AeroQ_lbfPerFt2_fHpKtasIsaDevFahrenheit = AeroQ_lbfPerFt2_fHpMach(hp, mach)
End Function

Function AeroQc_lbfPerFt2_fHpKtasOatCelsius(hp, ktas, oatCelsius)
    mach = AeroMach_fKtasOatCelsius(ktas, oatCelsius)
    AeroQc_lbfPerFt2_fHpKtasOatCelsius = AeroQc_lbfPerFt2_fHpMach(hp, mach)
End Function

Function AeroQc_lbfPerFt2_fHpKtasOatKelvin(hp, ktas, oatKelvin)
    mach = AeroMach_fKtasOatKelvin(ktas, oatKelvin)
    AeroQc_lbfPerFt2_fHpKtasOatKelvin = AeroQc_lbfPerFt2_fHpMach(hp, mach)
End Function

Function AeroQc_lbfPerFt2_fHpKtasOatRankine(hp, ktas, oatRankine)
    mach = AeroMach_fKtasOatRankine(ktas, oatRankine)
    AeroQc_lbfPerFt2_fHpKtasOatRankine = AeroQc_lbfPerFt2_fHpMach(hp, mach)
End Function

Function AeroQc_lbfPerFt2_fHpKtasOatFahrenheit(hp, ktas, oatFahrenheit)
    mach = AeroMach_fKtasOatFahrenheit(ktas, oatFahrenheit)
    AeroQc_lbfPerFt2_fHpKtasOatFahrenheit = AeroQc_lbfPerFt2_fHpMach(hp, mach)
End Function

Function AeroQc_lbfPerFt2_fHpKtasIsaDevCelsius(hp, ktas, isaDevCelsius)
    mach = AeroMach_fHpKtasIsaDevCelsius(hp, ktas, isaDevCelsius)
    AeroQc_lbfPerFt2_fHpKtasIsaDevCelsius = AeroQc_lbfPerFt2_fHpMach(hp, mach)
End Function

Function AeroQc_lbfPerFt2_fHpKtasIsaDevFahrenheit(hp, ktas, isaDevFahrenheit)
    mach = AeroMach_fHpKtasIsaDevFahrenheit(hp, ktas, isaDevFahrenheit)
    AeroQc_lbfPerFt2_fHpKtasIsaDevFahrenheit = AeroQc_lbfPerFt2_fHpMach(hp, mach)
End Function

Function AeroHp_ft_fPstaticLbPerFt2(pstaticLbPerFt2)
    If (pstaticLbPerFt2 > constDeltaTropopause * constPo) Then 'Troposphere, Hp < 11 km
        AeroHp_ft_fPstaticLbPerFt2 = constTo / constLapseTrop * ((pstaticLbPerFt2 / constPo) ^ (-constRAir * constLapseTrop / constGo) - 1)
    ElseIf (pstaticLbPerFt2 > constDeltaStrat20km * constPo) Then 'First layer stratosphere (isothermal layer), Hp < 20 km
        AeroHp_ft_fPstaticLbPerFt2 = tropopause11kmInFt - constRAir * constOatIsoLayerStrat11to20kmInK / constGo * Log(pstaticLbPerFt2 / (constDeltaTropopause * constPo))
    ElseIf (pstaticLbPerFt2 > constDeltaStrat32km * constPo) Then 'Second layer stratosphere, Hp < 32 km
        AeroHp_ft_fPstaticLbPerFt2 = constOatIsoLayerStrat11to20kmInK / constLapseStrat2 * ((pstaticLbPerFt2 / (constDeltaStrat20km * constPo)) ^ (-constRAir * constLapseStrat2 / constGo) - 1) + topIsoThermLayerStrat20kmInFt
    ElseIf (pstaticLbPerFt2 > constDeltaStrat47km * constPo) Then 'Third layer stratosphere, Hp < 47 km
        AeroHp_ft_fPstaticLbPerFt2 = constOatStrat32kmInK / constLapseStrat3 * ((pstaticLbPerFt2 / (constDeltaStrat32km * constPo)) ^ (-constRAir * constLapseStrat3 / constGo) - 1) + toplstInverLayerStrat32kmInFt
    Else 'Assume last layer stratosphere (isothermal layer), Hp < 51 km (and assign error value if greater than 51 km)
        AeroHp_ft_fPstaticLbPerFt2 = stratopauseStart47kmInFt - constRAir * constOatStratopause47to51kmInK / constGo * Log(pstaticLbPerFt2 / (constDeltaStrat47km * constPo))
        If (AeroHp_ft_fPstaticLbPerFt2 > stratopauseEnd51kmInFt) Then
            AeroHp_ft_fPstaticLbPerFt2 = constAeroErr
        End If
    End If
End Function

Function AeroHp_ft_fPstaticLbPerInch2(pstaticLbPerInch2)
    AeroHp_ft_fPstaticLbPerInch2 = AeroHp_ft_fPstaticLbPerFt2(pstaticLbPerInch2 * constConvLbfPerInch2ToLbfPerFt2)
End Function

Function AeroHp_ft_fPstaticHPa(pstaticHPa)
    AeroHp_ft_fPstaticHPa = AeroHp_ft_fPstaticLbPerFt2(pstaticHPa * constConvFtToM ^ 2 / constConvLbfToNewton * 100)
End Function

Function AeroHp_ft_fPstaticInHg(pstaticInHg)
    AeroHp_ft_fPstaticInHg = AeroHp_ft_fPstaticHPa(pstaticInHg * constConvInHgToHectoPascal)
End Function

Function AeroHp_ft_fQfeLbPerFt2(qfeLbPerFt2)
    'Pass-thru to Pstatic function (QFE same as Pstatic)
    AeroHp_ft_fQfeLbPerFt2 = AeroHp_ft_fPstaticLbPerFt2(qfeLbPerFt2)
End Function

Function AeroHp_ft_fQfeLbPerInch2(qfeLbPerInch2)
    'Pass-thru to Pstatic function (QFE same as Pstatic)
    AeroHp_ft_fQfeLbPerInch2 = AeroHp_ft_fPstaticLbPerInch2(qfeLbPerInch2)
End Function

Function AeroHp_ft_fQfeHPa(qfeHPa)
    'Pass-thru to Pstatic function (QFE same as Pstatic)
    AeroHp_ft_fQfeHPa = AeroHp_ft_fPstaticHPa(qfeHPa)
End Function

Function AeroHp_ft_fQfeInHg(qfeInHg)
    'Pass-thru to Pstatic function (QFE same as Pstatic)
    AeroHp_ft_fQfeInHg = AeroHp_ft_fPstaticInHg(qfeInHg)
End Function

Function AeroWndHead_fAcHdgWndHdgWndSpd(acHdgDeg, wndHdgDeg, wndSpd)
    'Note: Assumes heading units degrees. Return unit not specified; returns whatever units used for input wndSpd
    If (acHdgDeg < 0 Or acHdgDeg > 360 Or wndHdgDeg < 0 Or wndHdgDeg > 360 Or wndSpd < 0) Then
        AeroWndHead_fAcHdgWndHdgWndSpd = constAeroErr
    Else
        AeroWndHead_fAcHdgWndHdgWndSpd = wndSpd * Cos((wndHdgDeg - acHdgDeg) * WorksheetFunction.Pi / 180)
    End If
End Function

Function AeroWndCross_fAcHdgWndHdgWndSpd(acHdgDeg, wndHdgDeg, wndSpd)
    'Note: Assumes heading units degrees. Return unit not specified; returns whatever units used for input wndSpd
    If (acHdgDeg < 0 Or acHdgDeg > 360 Or wndHdgDeg < 0 Or wndHdgDeg > 360 Or wndSpd < 0) Then
        AeroWndCross_fAcHdgWndHdgWndSpd = constAeroErr
    Else
        AeroWndCross_fAcHdgWndHdgWndSpd = wndSpd * Sin((wndHdgDeg - acHdgDeg) * WorksheetFunction.Pi / 180)
    End If
End Function

Function AeroDHp_fDGeoPtlAltIsaDevCBaseHp(dGeoPtlAlt, isaDevC, baseHp)
    If (Abs(dGeoPtlAlt) <= constMinDAlt) Then
        steps = 1
    Else
        steps = Int(Abs(dGeoPtlAlt / constMinDAlt) + 1)
    End If
    geoPtlIncr = dGeoPtlAlt / steps
    accumDelHp = 0
    For i = 1 To steps
        tStd = AeroOatStdDay_Kelvin_fHp(baseHp + accumDelHp + geoPtlIncr / 2)
        accumDelHp = tStd / (tStd + isaDevC) * geoPtlIncr + accumDelHp
    Next i
    AeroDHp_fDGeoPtlAltIsaDevCBaseHp = accumDelHp
End Function

Function AeroDGeoPtlAlt_fDHpIsaDevCBaseHp(dHp, isaDevC, baseHp)
    If (Abs(dHp) <= constMinDAlt) Then
        steps = 1
    Else
        steps = Int(Abs(dHp / constMinDAlt) + 1)
    End If
    hpIncr = dHp / steps
    accumGeoPtlAlt = 0
    For i = 1 To steps
        tStd = AeroOatStdDay_Kelvin_fHp(baseHp + i * hpIncr - hpIncr / 2)
        accumGeoPtlAlt = (tStd + isaDevC) / tStd * hpIncr + accumGeoPtlAlt
    Next i
    AeroDGeoPtlAlt_fDHpIsaDevCBaseHp = accumGeoPtlAlt
End Function

Function AeroQnh_lbPerFt2_fHpGeoPtlAlt(hp, geoPtlAlt)
    AeroQnh_lbPerFt2_fHpGeoPtlAlt = constPo * (1 + constLapseTrop / constTo * (hp - geoPtlAlt)) ^ (-constGo / constRAir / constLapseTrop)
End Function

Function AeroQnh_lbPerInch2_fHpGeoPtlAlt(hp, geoPtlAlt)
    AeroQnh_lbPerInch2_fHpGeoPtlAlt = AeroQnh_lbPerFt2_fHpGeoPtlAlt(hp, geoPtlAlt) / constConvLbfPerInch2ToLbfPerFt2
End Function

Function AeroQnh_inHg_fHpGeoPtlAlt(hp, geoPtlAlt)
    AeroQnh_inHg_fHpGeoPtlAlt = AeroConvLbfPerFt2ToInHg(AeroQnh_lbPerFt2_fHpGeoPtlAlt(hp, geoPtlAlt))
End Function

Function AeroQnh_hPa_fHpGeoPtlAlt(hp, geoPtlAlt)
    AeroQnh_hPa_fHpGeoPtlAlt = AeroConvLbfPerFt2ToHPa(AeroQnh_lbPerFt2_fHpGeoPtlAlt(hp, geoPtlAlt))
End Function

Function AeroHp_fQnhLbfPerFt2GeoPtlAlt(qnh, geoPtlAlt)
    AeroHp_fQnhLbfPerFt2GeoPtlAlt = constTo / constLapseTrop * ((qnh / constPo) ^ (-constRAir * constLapseTrop / constGo) - 1) + geoPtlAlt
End Function

Function AeroHp_fQnhLbfPerInch2GeoPtlAlt(qnh, geoPtlAlt)
    AeroHp_fQnhLbfPerInch2GeoPtlAlt = AeroHp_fQnhLbfPerFt2GeoPtlAlt(AeroConvLbfPerInch2ToLbfPerFt2(qnh), geoPtlAlt)
End Function

Function AeroHp_fQnhInHgGeoPtlAlt(qnh, geoPtlAlt)
    AeroHp_fQnhInHgGeoPtlAlt = AeroHp_fQnhLbfPerFt2GeoPtlAlt(AeroConvInHgToLbfPerFt2(qnh), geoPtlAlt)
End Function

Function AeroHp_fQnhHPaGeoPtlAlt(qnh, geoPtlAlt)
    AeroHp_fQnhHPaGeoPtlAlt = AeroHp_fQnhLbfPerFt2GeoPtlAlt(AeroConvHPaToLbfPerFt2(qnh), geoPtlAlt)
End Function

Function AeroConvGradFtPerNmToPct(ftPerNm)
    AeroConvGradFtPerNmToPct = AeroConvFtToNm(ftPerNm) * 100
End Function

Function AeroConvGradPctToFtPerNm(pct)
    AeroConvGradPctToFtPerNm = AeroConvNmToFt(pct / 100)
End Function

Function AeroConvGradPctToDeg(pct)
    AeroConvGradPctToDeg = AeroConvRadToDeg(Atn(pct / 100))
End Function

Function AeroConvGradDegToPct(deg)
    AeroConvGradDegToPct = Tan(AeroConvDegToRad(deg)) * 100
End Function

Function AeroConvGradFtPerNmToDeg(ftPerNm)
    AeroConvGradFtPerNmToDeg = AeroConvRadToDeg(Atn(AeroConvFtToNm(ftPerNm)))
End Function

Function AeroConvGradDegToFtPerNm(deg)
    AeroConvGradDegToFtPerNm = AeroConvNmToFt(Tan(AeroConvDegToRad(deg)))
End Function

Function AeroConvFtToM(ft)
    AeroConvFtToM = ft * constConvFtToM
End Function

Function AeroConvMToFt(m)
    AeroConvMToFt = m / constConvFtToM
End Function

Function AeroConvMToNm(m)
    AeroConvMToNm = m / constConvNmToM
End Function

Function AeroConvNmToM(nm)
    AeroConvNmToM = nm * constConvNmToM
End Function

Function AeroConvFtToNm(ft)
    AeroConvFtToNm = ft * constConvFtToM / constConvNmToM
End Function

Function AeroConvNmToFt(nm)
    AeroConvNmToFt = nm * constConvNmToM / constConvFtToM
End Function

Function AeroTurnNz_fBankDeg(bankDeg)
    AeroTurnNz_fBankDeg = 1 / Cos(AeroConvDegToRad(bankDeg))
End Function

Function AeroTurnBank_deg_fNz(nz)
    AeroTurnBank_deg_fNz = AeroConvRadToDeg(WorksheetFunction.Acos(1 / nz))
End Function

Function AeroTurnRadius_ft_fNzKtas(nz, ktas)
    AeroTurnRadius_ft_fNzKtas = (AeroConvKtsToFtPerSec(ktas)) ^ 2 / (constGo * (nz ^ 2 - 1) ^ 0.5)
End Function

Function AeroTurnRadius_ft_fBankDegKtas(bankDeg, ktas)
    AeroTurnRadius_ft_fBankDegKtas = AeroTurnRadius_ft_fNzKtas(AeroTurnNz_fBankDeg(bankDeg), ktas)
End Function

Function AeroTurnNz_fRadiusFtKtas(radius, ktas)
    AeroTurnNz_fRadiusFtKtas = ((AeroConvKtsToFtPerSec(ktas) ^ 2 / (radius * constGo)) ^ 2 + 1) ^ 0.5
End Function

Function AeroTurnBank_deg_fRadiusKtas(radius, ktas)
    AeroTurnBank_deg_fRadiusKtas = AeroTurnBank_deg_fNz(AeroTurnNz_fRadiusFtKtas(radius, ktas))
End Function

Function AeroTurnRate_degPerSec_fNzKtas(nz, ktas)
    AeroTurnRate_degPerSec_fNzKtas = AeroTurnRate_degPerSec_fRadiusFtKtas(AeroTurnRadius_ft_fNzKtas(nz, ktas), ktas)
End Function

Function AeroTurnRate_degPerSec_fBankDegKtas(bankDeg, ktas)
    AeroTurnRate_degPerSec_fBankDegKtas = AeroTurnRate_degPerSec_fRadiusFtKtas(AeroTurnRadius_ft_fBankDegKtas(bankDeg, ktas), ktas)
End Function

Function AeroTurnNz_fTurnRateDegPerSecKtas(turnRateDegPerSec, ktas)
    AeroTurnNz_fTurnRateDegPerSecKtas = AeroTurnNz_fRadiusFtKtas(AeroTurnRadius_ft_fTurnRateDegPerSecKtas(turnRateDegPerSec, ktas), ktas)
End Function

Function AeroTurnBank_deg_fTurnRateDegPerSecKtas(turnRateDegPerSec, ktas)
    AeroTurnBank_deg_fTurnRateDegPerSecKtas = AeroTurnBank_deg_fRadiusKtas(AeroTurnRadius_ft_fTurnRateDegPerSecKtas(turnRateDegPerSec, ktas), ktas)
End Function

Function AeroTurnRate_degPerSec_fRadiusFtKtas(radiusFt, ktas)
    AeroTurnRate_degPerSec_fRadiusFtKtas = ktas / 20 / (WorksheetFunction.Pi * AeroConvFtToNm(radiusFt))
End Function

Function AeroTurnRadius_ft_fTurnRateDegPerSecKtas(turnRateDegPerSec, ktas)
    AeroTurnRadius_ft_fTurnRateDegPerSecKtas = AeroConvNmToFt(ktas / (20 * WorksheetFunction.Pi * turnRateDegPerSec))
End Function

Function AeroRoc_HpFtPerMin_fHpIsaDevCGradFtPerNmKtas(hp, isaDevC, gradFtPerNm, ktas)
    oatStdC = AeroOatStdDay_Kelvin_fHp(hp)
    tRatio = (oatStdC + isaDevC) / oatStdC
    AeroRoc_HpFtPerMin_fHpIsaDevCGradFtPerNmKtas = gradFtPerNm * ktas / 60 / tRatio
End Function

Function AeroRoc_HpFtPerMin_fHpIsaDevCGradFtPerNmKcas(hp, isaDevC, gradFtPerNm, kcas)
    ktas = AeroKtas_fHpKcasIsaDevCelsius(hp, kcas, isaDevC)
    AeroRoc_HpFtPerMin_fHpIsaDevCGradFtPerNmKcas = AeroRoc_HpFtPerMin_fHpIsaDevCGradFtPerNmKtas(hp, isaDevC, gradFtPerNm, ktas)
End Function
