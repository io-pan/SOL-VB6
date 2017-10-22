Attribute VB_Name = "formules"
Private Const InclinaisonTerre = 23.45

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''               T R I G O            ''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function pi() As Double
 pi = 4 * Atn(1)
End Function
Function rad(X As Double) As Double
    rad = pi * (X) / 180
End Function
Function deg2min(X As Double) As Double
    deg2min = X * 4
End Function
Function min2deg(X As Double)
   min2deg = X / 4
End Function
Function deg(X As Double)
    deg = 180 * (X) / pi
End Function

Function CosDeg(degre As Double) As Double
    CosDeg = Cos(rad(degre))
End Function
Function SinDeg(degre As Double) As Double
    SinDeg = Sin(rad(degre))
End Function
Function TanDeg(degre As Double) As Double
     TanDeg = Tan(rad(degre))
End Function
Function AtnDeg(degre As Double) As Double
     AtnDeg = Atn(rad(degre))
End Function
Function aSin(X As Double) As Double
    Dim A As Double
    If X = 1 Then
        aSin = pi / 2
    Else
        A = -X * X
        A = 1 + A 'C uoi ce bordel:au 27.07 x=1 mais -x*x+1 =0.000000et des merdes
        aSin = Atn(X / Sqr(A))
    End If
End Function
Function aSinDeg(X As Double) As Double
    aSinDeg = aSin(rad(X))
End Function
Function aCos(X As Double) As Double
    Dim A As Double
    If X = 1 Then
        aCos = 0
    Else
        A = -X * X
        A = 1 + A 'C uoi ce bordel:au 27.07 x=1 mais -x*x+1 =0.000000et des merdes
        aCos = Atn(-X / Sqr(A)) + 2 * Atn(1)
    End If
End Function
Function aCosDeg(X As Double)
    aCosDeg = aCos(rad(X))
End Function










'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''        S O L A I R E           ''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function JourJ(jour As Date) As Double
' Rang du jour dans l'ann�e (1er janvier = 1)
    Dim d, M, A, N1, N2, k As Integer
    d = Day(jour)
    M = Month(jour)
    A = Year(jour)
    N1 = Int((M * 275) / 9)
    N2 = Int((M + 9) / 12)
    k = 1 + Int((A - 4 * Int(A / 4) + 2) / 3) 'K=2 pour une ann�e commune et K=1 pour une bissextile)
    JourJ = N1 - N2 * k + d - 30
End Function

Function eqt(jour As Double) As Double
 ' jour repr�sente le rang du jour dans l'ann�e (1er janvier = 1)
 Dim M As Double ' M est l'anomalie moyenne en degr�s
 Dim C As Double ' C est l'�quation du centre (influence de l'ellipticit� de l'orbite terrestre) en degr�s
 Dim L As Double ' L est la longitude vraie du Soleil en degr�s
 Dim R As Double ' R est la r�duction � l'�quateur (influence de l'inclinaison de l'axe terrestre) en degr�s
    
    M = 357 + 0.9856 * jour     ' 0,9856 repr�sente naturellement le d�placement quotidien
                                ' moyen de la Terre sur son orbite autour du Soleil
    C = 1.914 * Sin(rad(M)) + 0.02 * Sin(rad(2 * M))
    L = 280 + C + 0.9856 * jour
    R = -2.466 * Sin(rad(2 * L)) + 0.053 * Sin(rad(4 * L))
    eqt = (C + R)               ' � de d�calage
    eqt = eqt * 4               ' minutes de d�calage entre midiSolaire et midiHoraire
End Function
Function eqt0(jour As Double) As Double
    Dim j As Double
    j = 2 * pi * (jour - 81) / 365
    eqt0 = -9.87 * Sin(2 * j) + 7.53 * Cos(j) + 1.5 * Sin(j) ' minutes de d�calage entre midiSolaire et midiHoraire
End Function

Function Declinaison(jour As Double) As Double
 Dim sind As Double
 Dim cosd As Double
 
    sind = -0.398 * Cos(0.01721 * (jour + 9))
    cosd = Sqr(1 - sind ^ 2)
    Declinaison = deg(aSin(sind))
End Function

Function Declinaison1(jour As Double) As Double
 ' j repr�sente le rang du jour dans l'ann�e (1er janvier = 1)
 Dim M As Double ' M est l'anomalie moyenne en degr�s
 Dim C As Double ' C est l'�quation du centre (influence de l'ellipticit� de l'orbite terrestre) en degr�s
 Dim L As Double ' L est la longitude vraie du Soleil en degr�s
    
    M = rad(357) + rad(0.9856) * jour                      '0,9856 repr�sente naturellement le d�placement quotidien moyen de la Terre sur son orbite autour du Soleil
    C = rad(1.914) * Sin(M) + rad(0.02) * Sin(2 * M)
    L = rad(280) + C + rad(0.9856) * jour
    Declinaison1 = aSin(rad(0.3978) * Sin(L))           '0.3978 repr�sente le sinus de l'obliquit� de l'�cliptique
    Declinaison1 = deg(Declinaison1) * (InclinaisonTerre / 0.3978)
End Function

Function heureEte(j As Date) As Integer
'le passage � l'heure d'�t� intervient le dernier dimanche de mars � 2 heures du matin et le passage �
'L 'heure d'hiver intervient le dernier dimanche d'octobre � 3 heures du matin.
    Dim jw As Date
    Dim jHEte As Date
    Dim jHHiver As Date
    
    heureEte = 0
    
    jw = DateValue("31.03." & Year(j) & " 2:00:00")
    Do While jw >= DateValue("01.03." & Year(j)) And jHEte = DateValue("0:0:0")
        If VBA.Format(jw, "dddd") = "dimanche" Then
            jHEte = jw
        End If
        jw = DateAdd("d", -1, jw)
    Loop
    
    jw = DateValue("31.10." & Year(j) & " 3:00:00")
    Do While jw >= DateValue("1.10." & Year(j)) And jHHiver = DateValue("0:0:0")
        If VBA.Format(jw, "dddd") = "dimanche" Then
            jHHiver = jw
        End If
        jw = DateAdd("d", -1, jw)
    Loop
    
    If DateValue(j) >= DateValue(jHEte) And DateValue(j) <= DateValue(jHHiver) Then
        heureEte = 1
    End If
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' LEVER     COUCHER     MIDI
    ' Le Soleil se l�ve ou se couche quand le bord sup�rieur de son disque
    ' appara�t ou disparait � l'horizon.
    ' Du fait de la r�fraction atmosph�rique le centre du Soleil est alors �
    ' 50' sous l'horizon :
    ' 34' pour l'effet de la r�fraction et 16' pour le demi-diam�tre du Soleil.


Function LeverS(Ja As Double, lat As Double, Longitude As Double, fuseau As Integer, Hete As Integer) As Date
Dim Ho As Double
Dim dec As Double
Dim X As Double
Dim M As Long
Dim lon As Long
Dim ete As Integer
Dim h As Long
Dim d As Date

    If Hete Then
        ete = 1
    Else
        ete = 0
    End If
        
    dec = Declinaison(Ja)
    X = (-0.01454 - Sin(rad(dec))) * Sin(rad(lat)) / (Cos(rad(dec)) * Cos(rad(lat)))
    If X <= 1 And X >= -1 Then
        Ho = aCos(X)
        Ho = deg(Ho)
        Ho = Ho / 15
        M = -(Ho - Int(Ho)) * 60
        Ho = -Int(Ho)
    
        M = Round(M + eqt(Ja) - Round(Longitude * 4))
        h = fuseau + ete + Ho
    
        d = TimeValue("12:00:00")
        d = DateAdd("n", M, d)
        d = DateAdd("h", h, d)
    End If
    LeverS = d
End Function


Function Coucher(Ja As Double, lat As Double, Longitude As Double, fuseau As Integer, Hete As Integer) As Date
Dim Ho As Double
Dim dec As Double
Dim X As Double
Dim ete As Integer
Dim h As Integer
Dim M As Double
Dim d As Date
    dec = Declinaison(Ja)
    If Hete Then
        ete = 1
    Else
        ete = 0
    End If
    
    X = (-0.01454 - Sin(rad(dec))) * Sin(rad(lat)) / (Cos(rad(dec)) * Cos(rad(lat)))
    If X <= 1 And X >= -1 Then 'pas cercle polaire
        Ho = aCos(X)
        Ho = deg(Ho)
        Ho = Ho / 15
        M = (Ho - Int(Ho)) * 60
        Ho = Int(Ho)
    
        M = Round(M + eqt(Ja) - Round(Longitude * 4))
        h = fuseau + ete + Ho
    End If
    d = TimeValue("12:00:00")
    d = DateAdd("n", M, d)
    d = DateAdd("h", h, d)
    Coucher = d
End Function

Function Midi(Ja As Double, Longitude As Double, fuseau As Integer, Hete As Integer) As Date
 Dim lon
 Dim d As Date
 Dim ete As Integer
 Dim h As Integer
 Dim M As Integer

    If Hete Then
        ete = 1
    Else
        ete = 0
    End If

    'midi
    M = Round(eqt(Ja)) - Round(Longitude * 4)
    h = fuseau + ete

    d = TimeValue("12:00:00")
    d = DateAdd("n", M, d)
    d = DateAdd("h", h, d)
    Midi = d
End Function

Function Duree(Ja As Double, lat As Double, Longitude As Double) As Date
Dim Ho As Double
Dim dec As Double
Dim X As Double
Dim M As Double
Dim h As Double
Dim d As Date
    dec = Declinaison(Ja)
    X = (-0.01454 - Sin(rad(dec))) * Sin(rad(lat)) / (Cos(rad(dec)) * Cos(rad(lat)))
    If X <= 1 And X >= -1 Then
        Ho = aCos(X)
        Ho = deg(Ho)
        Ho = Ho / 15 * 2
        M = (Ho - Int(Ho)) * 60
        h = Int(Ho)

        M = M + Round(eqt(Ja)) - Round(Longitude * 4)
        d = TimeValue("00:00:00")
        d = DateAdd("n", M, d)
        d = DateAdd("h", h, d)
        Duree = d
    Else
        Duree = TimeValue("23:59:59")
        End If
End Function

Function AzimuthLever(Ja As Double, lat As Double) As Double
    Dim A As Double
    Dim dec As Double
    
    'If Lat < 66.5 and lat >-66.5 Then 'cercle pol�aire, pas de lever des fois
    dec = Declinaison(Ja)
    A = (-0.01454 * Sin(rad(lat)) - Sin(rad(dec))) / Cos(rad(lat))
    If A <= 1 And A >= -1 Then
        A = aCos(A)
        A = deg(A)
        AzimuthLever = -A
    Else
        AzimuthLever = 0
    End If
End Function
'==========================================================================================================
'Prise en compte de l'altitude de l'observateur
'
'Quand L 'altitude du lieu d'observation devient appr�ciable l'horizon visible est abaiss� et recul�, et les ph�nom�nes de lever/coucher ainsi que les cr�puscules s'en trouvent affect�s. C'est un probl�me connu en navigation astronomique : le navigateur doit estimer la d�pression de l'horizon pour passer de la hauteur apparente du soleil � sa hauteur vraie. Ici la formule donnant l'angle horaire du soleil au moments des lever/coucher devient pour un lieu de hauteur ALT:
'
'Cos(Ho) = [ -sin ( ( 1,76 �( ALT) + 50 ) / 60 ) - sin ( Dec ) * sin ( Lat ) ] / [ cos ( Dec ) * cos ( Lat ) ]
'
'si ALT est exprim�e en m�tres,
'
'-5- Cr�puscules : par d�finition la fin (le soir) ou le d�but (le matin) des cr�puscules civil, nautique et astronomique se produit quand le centre du Soleil est abaiss� de 6�, 12� et 18� sous l'horizon. La suite des calculs est la m�me que pour les lever/coucher, il suffit de remplacer -0,01454 dans la formule ci-dessus par respectivement -0,105 / -0,208 / -0,309. La remarque sur la valeur du cosinus est valable pour les cr�puscules : au dessus de 48�,5 de latitude en France (et en Bretagne en particulier) le cr�puscule astronomique dure toute la nuit autour du solstice d'�t�.
'
'Pour un observateur ce sont des moments tr�s subjectifs : pour lui, le soir par exemple, il n'y a qu'une d�croissance continue de la luminosit�. Il n'y a rien de sp�cial qui marque le moment calcul� des trois cr�puscules. Ces cr�puscules ont n�anmoins un sens r�el aux latitudes, comme les latitudes moyennes,  o� le soleil a un comportement � peu pr�s identique tout au long de l'ann�e. Dans les r�gions tropicales il s'�coule environ une heure entre les toutes premi�res lueurs de l'aube et le lever du soleil; inversement, le soir, la nuit tombe totalement une heure apr�s le coucher du soleil. Les 3 cr�puscules d�finis ci-dessus se succ�dent � des intervalles de 20 minutes et perdent un peu de leur pertinence; on a � peine le temps de les appr�cier !. Je propose alors un cr�puscule "tropical" correspondant � un abaissement de 9� sous l'horizon; le moment correspondant est � mi chemin entre la nuit compl�te et le lever ou le coucher.

Function hauteurSol(Ja As Double, heure As Date, Latitude As Double, Longitude As Double, GMT As Integer, Hete As Integer) As Double
  Dim Declin As Double
  Dim HeureSol As Double
 
    HeureSol = heureSolaireRad(Ja, heure, Longitude, GMT, Hete) ' 0=midi ; pi()=minuit ; -pi()/2=pi()/2
    
    Declin = rad(Declinaison1(Ja))
    hauteurSol = aSin(Sin(rad(Latitude)) * Sin(Declin) + Cos(rad(Latitude)) * Cos(Declin) * Cos(HeureSol))
End Function

Function hauteurSolMax(Ja As Double, Latitude As Double) As Double
 Dim Declin As Double
   
    Latitude = rad(Latitude)
    Declin = rad(Declinaison1(Ja))
    hauteurSolMax = aSin(Sin(Latitude) * Sin(Declin) + Cos(Latitude) * Cos(Declin))
End Function

Function Azimuth(Ja As Double, heure As Date, Latitude As Double, Longitude As Double, GMT As Integer, Hete As Integer) As Double
 Dim Declin As Double
 Dim HeureSol As Double
 Dim HauteurRad As Double
   
    HauteurRad = hauteurSol(Ja, heure, Latitude, Longitude, GMT, Hete)
    HeureSol = heureSolaireRad(Ja, heure, Longitude, GMT, Hete) ' 0=midi ; pi()=minuit ; -pi()/2=pi()/2
    Declin = rad(Declinaison1(Ja))
    Azimuth = aSin((Sin(rad(Latitude)) * Cos(Declin) * Cos(HeureSol) - Cos(rad(Latitude)) * Sin(Declin)) / Cos(HauteurRad)) - pi() / 2
    If HeureSol >= 0 And HeureSol <= pi() Then
        Azimuth = -Azimuth  '0�= midi solaire
    End If
End Function

Function heureSolaire(Ja As Double, heure As Date, Longitude As Double, fuseau As Integer, Hete As Integer) As Date
 ' H Solaire = H Civile -( correction longitude + 1 h (ou 2 l'�t�) + correction EQ TEMPS
 Dim lon As Long
 Dim ete As Integer
 Dim h As Integer
 Dim M As Integer
 Dim d As Date
 
    'midi solaire
    M = Round(eqt(Ja)) - Round(Longitude * 4)
    h = fuseau + Hete

    d = heure
    d = DateAdd("n", -M, d)
    d = DateAdd("h", -h, d)
    heureSolaire = d
End Function

Function heureSolaireRad(Ja As Double, heure As Date, Longitude As Double, GMT As Integer, Hete As Integer) As Double  ' 0=midi: pi()=minuit
  Dim M As Double    ' minutes
  Dim h As Integer   ' heures
  Dim X As Double    ' nombre de minutes depuis midi
  Dim Hrad As Double
  Dim corecLongi As Double
  Dim correcH As Double
  Dim correcM As Double

    M = Minute(heure)
    h = Hour(heure)
    
    ' correction longitude et Equation du temps
    correcM = eqt(Ja) - Longitude * 4
        
    ' correction heures d'�t� et GMT
    correcH = correcH + Hete + GMT

    ' convertion de l'heure en tps avec pour origine midi (tps >0 ou <0)
    h = h + 12
    If h >= 24 Then
        h = h - 24
    End If
    
    ' convertion du temps en angle
    X = (h - correcH) * 60 + (M - correcM)
    heureSolaireRad = pi() * X / 720            ' 2pi * X / 24*60
    
End Function

Function JJulien(jour As Date, Optional heure As Date) As Double
  Dim j As Integer      ' jour du mois
  Dim M As Integer      ' Mois
  Dim A As Integer      ' ann�e
  Dim HH As Integer     ' UTC sous la forme HH:MM:SS
  Dim MM As Integer
  Dim SS As Integer
    Dim C As Integer
    Dim B As Integer
    Dim t As Double
    
    j = Day(jour)
    M = Month(jour)
    A = Year(jour)
    HH = Hour(heure)
    MM = Minute(heure)
    SS = Second(heure)
    
    ' Janvier F�vrier consid�r�s comme 13� et 14� mois de l'ann�e pr�c�dente
    If M = 1 Or M = 2 Then
        A = A - 1
        M = M + 12
    End If

    ' J M A est une date du calendrier gr�gorien calculer :
    C = Int(A / 100)
    B = 2 - C + Int(C / 4)

    ' Calculer la fraction de jour correspondant � HH MM SS
    t = HH / 24 + MM / 1440 + SS / 86400

    'Le jour julien est donn� par :
    JJulien = Int(365.25 * (A + 4716)) + Int(30.6001 * (M + 1)) + j + t + B - 1524.5
End Function
