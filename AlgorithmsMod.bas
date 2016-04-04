Attribute VB_Name = "AlgorithmsMod"
Attribute VB_Description = "Copyright © 2004 by William J. McKibbin"
'****************************************
'*          ISOCAPMOD FUNCTIONS         *
'*              Version 1.0             *
'*            Copyright © 2004          *
'*         by William J. McKibbin       *
'****************************************

DefDbl A-Z

Function IsoCapCoef(magnitude1, magnitude2, magnitude3) As Double
Attribute IsoCapCoef.VB_Description = "Copyright © 2004 by William J. McKibbin"

    Dim M1, M2, M3, M123
        
    M1 = Abs(magnitude1)
    M2 = Abs(magnitude2)
    M3 = Abs(magnitude3)
    M123 = M1 + M2 + M3
    
    If M1 = 0 Or M2 = 0 Or M3 = 0 Then
        IsoCapCoef = 1
        GoTo LastLine
        
    End If
    
    Dim Pi, VG
    Pi = Application.Pi()
    VG = 1

    Dim EDF, DEF, DFE
    EDF = 180 * (M1 / M123)
    DEF = 180 * (M2 / M123)
    DFE = 180 * (M3 / M123)
    
    Dim DGV, GDV, GEV
    DGV = 90
    GDV = EDF / 2
    GEV = DEF / 2

    Dim DVG, GVE
    DVG = 180 - DGV - GDV
    GVE = 180 - DGV - GEV

    Dim ADE, AED, BDF, BFD, CEF, CFE
    ADE = (180 - EDF) / 2
    AED = (180 - DEF) / 2
    BDF = (180 - EDF) / 2
    BFD = (180 - DFE) / 2
    CEF = (180 - DEF) / 2
    CFE = (180 - DFE) / 2

    Dim DAE, ECF, DBF
    DAE = 180 - ADE - AED
    ECF = 180 - CEF - CFE
    DBF = 180 - BDF - BFD

    Dim DG, GE, DE
    DG = VG * Sin(DVG * (Pi / 180)) / Sin(GDV * (Pi / 180))
    GE = VG * Sin(GVE * (Pi / 180)) / Sin(GEV * (Pi / 180))
    DE = DG + GE

    Dim DF, EF
    DF = DE * Sin(DEF * (Pi / 180)) / Sin(DFE * (Pi / 180))
    EF = DE * Sin(EDF * (Pi / 180)) / Sin(DFE * (Pi / 180))

    Dim BD, BF
    BD = DF * Sin(BFD * (Pi / 180)) / Sin(DBF * (Pi / 180))
    BF = DF * Sin(BDF * (Pi / 180)) / Sin(DBF * (Pi / 180))

    Dim CF, CE
    CF = EF * Sin(CEF * (Pi / 180)) / Sin(ECF * (Pi / 180))
    CE = EF * Sin(CFE * (Pi / 180)) / Sin(ECF * (Pi / 180))

    Dim AD, AE
    AD = DE * Sin(AED * (Pi / 180)) / Sin(DAE * (Pi / 180))
    AE = DE * Sin(ADE * (Pi / 180)) / Sin(DAE * (Pi / 180))

    Dim SEMIADE, SEMIBDF, SEMICEF
    SEMIADE = (DE + AD + AE) / 2
    SEMIBDF = (DF + BD + BF) / 2
    SEMICEF = (EF + CF + CE) / 2

    Dim AREAADE, AREABDF, AREACEF
    AREAADE = Sqr(SEMIADE * (SEMIADE - DE) * (SEMIADE - AD) * (SEMIADE - AE))
    AREABDF = Sqr(SEMIBDF * (SEMIBDF - DF) * (SEMIBDF - BD) * (SEMIBDF - BF))
    AREACEF = Sqr(SEMICEF * (SEMICEF - EF) * (SEMICEF - CF) * (SEMICEF - CE))
    
    Dim RA, RB, RC
    RA = AREAADE / (DE / 2)
    RB = AREABDF / (DF / 2)
    RC = AREACEF / (EF / 2)
    
    Dim RABC
    RABC = RA + RB + RC
        
    Dim V2
    V2 = (RABC - VG) / 4
    
    Dim V3
    V3 = V2 / 2
            
    IsoCapCoef = ((RC / RABC) * (M123)) / M1
         
    Dim Test
    Test = VG + ((V3 * 8) - RABC)
    
    If Not Test = 0 Then Error 2015

LastLine: End Function

Function IsoCapSynergy(magnitude1, magnitude2, magnitude3) As Double
Attribute IsoCapSynergy.VB_Description = "Copyright © 2004 by William J. McKibbin"
    
    Dim M1, M2, M3, M123, S
        
    M1 = Abs(magnitude1)
    M2 = Abs(magnitude2)
    M3 = Abs(magnitude3)
    M123 = M1 + M2 + M3
    
    If M1 = 0 And M2 = 0 And M3 = 0 Then
        IsoCapSynergy = 0
        GoTo LastLine
        
    ElseIf M1 > 0 And M2 = 0 And M3 = 0 Then
        S = M1 / magnitude1
        IsoCapSynergy = 1 * S
        GoTo LastLine
        
    ElseIf M1 = 0 And M2 > 0 And M3 = 0 Then
        S = M2 / magnitude2
        IsoCapSynergy = 1 * S
        GoTo LastLine
        
    ElseIf M1 = 0 And M2 = 0 And M3 > 0 Then
        S = M3 / magnitude3
        IsoCapSynergy = 1 * S
        GoTo LastLine
        
    ElseIf M1 = 0 And M2 > 0 And M3 > 0 Then
        S = (M2 / magnitude2) * (M3 / magnitude3)
        IsoCapSynergy = 1 * S
        GoTo LastLine
       
    ElseIf M1 > 0 And M2 = 0 And M3 > 0 Then
        S = (M1 / magnitude1) * (M3 / magnitude3)
        IsoCapSynergy = 1 * S
        GoTo LastLine
        
    ElseIf M1 > 0 And M2 > 0 And M3 = 0 Then
        S = (M1 / magnitude1) * (M2 / magnitude2)
        IsoCapSynergy = 1 * S
        GoTo LastLine
        
    End If
    
    Dim Pi, VG
    Pi = Application.Pi()
    VG = 1

    Dim EDF, DEF, DFE
    EDF = 180 * (M1 / M123)
    DEF = 180 * (M2 / M123)
    DFE = 180 * (M3 / M123)
    
    Dim DGV, GDV, GEV
    DGV = 90
    GDV = EDF / 2
    GEV = DEF / 2

    Dim DVG, GVE
    DVG = 180 - DGV - GDV
    GVE = 180 - DGV - GEV

    Dim ADE, AED, BDF, BFD, CEF, CFE
    ADE = (180 - EDF) / 2
    AED = (180 - DEF) / 2
    BDF = (180 - EDF) / 2
    BFD = (180 - DFE) / 2
    CEF = (180 - DEF) / 2
    CFE = (180 - DFE) / 2

    Dim DAE, ECF, DBF
    DAE = 180 - ADE - AED
    ECF = 180 - CEF - CFE
    DBF = 180 - BDF - BFD

    Dim DG, GE, DE
    DG = VG * Sin(DVG * (Pi / 180)) / Sin(GDV * (Pi / 180))
    GE = VG * Sin(GVE * (Pi / 180)) / Sin(GEV * (Pi / 180))
    DE = DG + GE

    Dim DF, EF
    DF = DE * Sin(DEF * (Pi / 180)) / Sin(DFE * (Pi / 180))
    EF = DE * Sin(EDF * (Pi / 180)) / Sin(DFE * (Pi / 180))

    Dim BD, BF
    BD = DF * Sin(BFD * (Pi / 180)) / Sin(DBF * (Pi / 180))
    BF = DF * Sin(BDF * (Pi / 180)) / Sin(DBF * (Pi / 180))

    Dim CF, CE
    CF = EF * Sin(CEF * (Pi / 180)) / Sin(ECF * (Pi / 180))
    CE = EF * Sin(CFE * (Pi / 180)) / Sin(ECF * (Pi / 180))

    Dim AD, AE
    AD = DE * Sin(AED * (Pi / 180)) / Sin(DAE * (Pi / 180))
    AE = DE * Sin(ADE * (Pi / 180)) / Sin(DAE * (Pi / 180))

    Dim SEMIADE, SEMIBDF, SEMICEF
    SEMIADE = (DE + AD + AE) / 2
    SEMIBDF = (DF + BD + BF) / 2
    SEMICEF = (EF + CF + CE) / 2

    Dim AREAADE, AREABDF, AREACEF
    AREAADE = Sqr(SEMIADE * (SEMIADE - DE) * (SEMIADE - AD) * (SEMIADE - AE))
    AREABDF = Sqr(SEMIBDF * (SEMIBDF - DF) * (SEMIBDF - BD) * (SEMIBDF - BF))
    AREACEF = Sqr(SEMICEF * (SEMICEF - EF) * (SEMICEF - CF) * (SEMICEF - CE))
    
    Dim RA, RB, RC
    RA = AREAADE / (DE / 2)
    RB = AREABDF / (DF / 2)
    RC = AREACEF / (EF / 2)
    
    Dim RABC
    RABC = RA + RB + RC
        
    Dim V2
    V2 = (RABC - VG) / 4
    
    Dim V3
    V3 = V2 / 2
    
    S = (M1 / magnitude1) * (M2 / magnitude2) * (M3 / magnitude3)
    
    IsoCapSynergy = ((((RC / RABC) * M123) / M1) * (((RB / RABC) * M123) / M2) * (((RA / RABC) * M123) / M3)) * S
          
    Dim Test
    Test = VG + ((V3 * 8) - RABC)
    
    If Not Test = 0 Then Error 2015
          
LastLine: End Function

Function IsoCapRatio(magnitude1, magnitude2, magnitude3) As Double
Attribute IsoCapRatio.VB_Description = "Copyright © 2004 by William J. McKibbin"
      
    Dim M1, M2, M3, M123, S
        
    M1 = Abs(magnitude1)
    M2 = Abs(magnitude2)
    M3 = Abs(magnitude3)
    M123 = M1 + M2 + M3
    
    If M1 = 0 Then
        IsoCapRatio = 0
        GoTo LastLine
        
    ElseIf M2 = 0 Or M3 = 0 Then
        S = M1 / magnitude1
        IsoCapRatio = (M1 / M123) * S
        GoTo LastLine
        
    End If
    
    Dim Pi, VG
    Pi = Application.Pi()
    VG = 1

    Dim EDF, DEF, DFE
    EDF = 180 * (M1 / M123)
    DEF = 180 * (M2 / M123)
    DFE = 180 * (M3 / M123)

    Dim DGV, GDV, GEV
    DGV = 90
    GDV = EDF / 2
    GEV = DEF / 2

    Dim DVG, GVE
    DVG = 180 - DGV - GDV
    GVE = 180 - DGV - GEV

    Dim ADE, AED, BDF, BFD, CEF, CFE
    ADE = (180 - EDF) / 2
    AED = (180 - DEF) / 2
    BDF = (180 - EDF) / 2
    BFD = (180 - DFE) / 2
    CEF = (180 - DEF) / 2
    CFE = (180 - DFE) / 2

    Dim DAE, ECF, DBF
    DAE = 180 - ADE - AED
    ECF = 180 - CEF - CFE
    DBF = 180 - BDF - BFD

    Dim DG, GE, DE
    DG = VG * Sin(DVG * (Pi / 180)) / Sin(GDV * (Pi / 180))
    GE = VG * Sin(GVE * (Pi / 180)) / Sin(GEV * (Pi / 180))
    DE = DG + GE

    Dim DF, EF
    DF = DE * Sin(DEF * (Pi / 180)) / Sin(DFE * (Pi / 180))
    EF = DE * Sin(EDF * (Pi / 180)) / Sin(DFE * (Pi / 180))

    Dim BD, BF
    BD = DF * Sin(BFD * (Pi / 180)) / Sin(DBF * (Pi / 180))
    BF = DF * Sin(BDF * (Pi / 180)) / Sin(DBF * (Pi / 180))

    Dim CF, CE
    CF = EF * Sin(CEF * (Pi / 180)) / Sin(ECF * (Pi / 180))
    CE = EF * Sin(CFE * (Pi / 180)) / Sin(ECF * (Pi / 180))

    Dim AD, AE
    AD = DE * Sin(AED * (Pi / 180)) / Sin(DAE * (Pi / 180))
    AE = DE * Sin(ADE * (Pi / 180)) / Sin(DAE * (Pi / 180))

    Dim SEMIADE, SEMIBDF, SEMICEF
    SEMIADE = (DE + AD + AE) / 2
    SEMIBDF = (DF + BD + BF) / 2
    SEMICEF = (EF + CF + CE) / 2

    Dim AREAADE, AREABDF, AREACEF
    AREAADE = Sqr(SEMIADE * (SEMIADE - DE) * (SEMIADE - AD) * (SEMIADE - AE))
    AREABDF = Sqr(SEMIBDF * (SEMIBDF - DF) * (SEMIBDF - BD) * (SEMIBDF - BF))
    AREACEF = Sqr(SEMICEF * (SEMICEF - EF) * (SEMICEF - CF) * (SEMICEF - CE))
    
    Dim RA, RB, RC
    RA = AREAADE / (DE / 2)
    RB = AREABDF / (DF / 2)
    RC = AREACEF / (EF / 2)
    
    Dim RABC
    RABC = RA + RB + RC
        
    Dim V2
    V2 = (RABC - VG) / 4
    
    Dim V3
    V3 = V2 / 2
    
    S = M1 / magnitude1
    
    IsoCapRatio = (RC / RABC) * S
    
    Dim Test
    Test = VG + ((V3 * 8) - RABC)
    
    If Not Test = 0 Then Error 2015
    
LastLine: End Function

Function IsoCapChange(magnitude1, magnitude2, magnitude3) As Double
Attribute IsoCapChange.VB_Description = "Copyright © 2004 by William J. McKibbin"

    Dim M1, M2, M3, M123, S
        
    M1 = Abs(magnitude1)
    M2 = Abs(magnitude2)
    M3 = Abs(magnitude3)
    M123 = M1 + M2 + M3
    
    If M1 = 0 Or M2 = 0 Or M3 = 0 Then
        IsoCapChange = 0
        GoTo LastLine
        
    End If
    
    Dim Pi, VG
    Pi = Application.Pi()
    VG = 1

    Dim EDF, DEF, DFE
    EDF = 180 * (M1 / M123)
    DEF = 180 * (M2 / M123)
    DFE = 180 * (M3 / M123)

    Dim DGV, GDV, GEV
    DGV = 90
    GDV = EDF / 2
    GEV = DEF / 2

    Dim DVG, GVE
    DVG = 180 - DGV - GDV
    GVE = 180 - DGV - GEV

    Dim ADE, AED, BDF, BFD, CEF, CFE
    ADE = (180 - EDF) / 2
    AED = (180 - DEF) / 2
    BDF = (180 - EDF) / 2
    BFD = (180 - DFE) / 2
    CEF = (180 - DEF) / 2
    CFE = (180 - DFE) / 2

    Dim DAE, ECF, DBF
    DAE = 180 - ADE - AED
    ECF = 180 - CEF - CFE
    DBF = 180 - BDF - BFD

    Dim DG, GE, DE
    DG = VG * Sin(DVG * (Pi / 180)) / Sin(GDV * (Pi / 180))
    GE = VG * Sin(GVE * (Pi / 180)) / Sin(GEV * (Pi / 180))
    DE = DG + GE

    Dim DF, EF
    DF = DE * Sin(DEF * (Pi / 180)) / Sin(DFE * (Pi / 180))
    EF = DE * Sin(EDF * (Pi / 180)) / Sin(DFE * (Pi / 180))

    Dim BD, BF
    BD = DF * Sin(BFD * (Pi / 180)) / Sin(DBF * (Pi / 180))
    BF = DF * Sin(BDF * (Pi / 180)) / Sin(DBF * (Pi / 180))

    Dim CF, CE
    CF = EF * Sin(CEF * (Pi / 180)) / Sin(ECF * (Pi / 180))
    CE = EF * Sin(CFE * (Pi / 180)) / Sin(ECF * (Pi / 180))

    Dim AD, AE
    AD = DE * Sin(AED * (Pi / 180)) / Sin(DAE * (Pi / 180))
    AE = DE * Sin(ADE * (Pi / 180)) / Sin(DAE * (Pi / 180))

    Dim SEMIADE, SEMIBDF, SEMICEF
    SEMIADE = (DE + AD + AE) / 2
    SEMIBDF = (DF + BD + BF) / 2
    SEMICEF = (EF + CF + CE) / 2

    Dim AREAADE, AREABDF, AREACEF
    AREAADE = Sqr(SEMIADE * (SEMIADE - DE) * (SEMIADE - AD) * (SEMIADE - AE))
    AREABDF = Sqr(SEMIBDF * (SEMIBDF - DF) * (SEMIBDF - BD) * (SEMIBDF - BF))
    AREACEF = Sqr(SEMICEF * (SEMICEF - EF) * (SEMICEF - CF) * (SEMICEF - CE))
    
    Dim RA, RB, RC
    RA = AREAADE / (DE / 2)
    RB = AREABDF / (DF / 2)
    RC = AREACEF / (EF / 2)
    
    Dim RABC
    RABC = RA + RB + RC
        
    Dim V2
    V2 = (RABC - VG) / 4
    
    Dim V3
    V3 = V2 / 2
    
    S = M1 / magnitude1
    
    IsoCapChange = ((RC / RABC) - (M1 / M123)) * S
    
    Dim Test
    Test = VG + ((V3 * 8) - RABC)
    
    If Not Test = 0 Then Error 2015

LastLine: End Function

Function IsoCapRelative(magnitude1, magnitude2, magnitude3) As Double
Attribute IsoCapRelative.VB_Description = "Copyright © 2004 by William J. McKibbin"

    Dim M1, M2, M3, M123, S
        
    M1 = Abs(magnitude1)
    M2 = Abs(magnitude2)
    M3 = Abs(magnitude3)
    M123 = M1 + M2 + M3
    
    If M1 = 0 Or M2 = 0 Or M3 = 0 Then
        IsoCapRelative = 0
        GoTo LastLine
        
    End If
    
    Dim Pi, VG
    Pi = Application.Pi()
    VG = 1

    Dim EDF, DEF, DFE
    EDF = 180 * (M1 / M123)
    DEF = 180 * (M2 / M123)
    DFE = 180 * (M3 / M123)

    Dim DGV, GDV, GEV
    DGV = 90
    GDV = EDF / 2
    GEV = DEF / 2

    Dim DVG, GVE
    DVG = 180 - DGV - GDV
    GVE = 180 - DGV - GEV

    Dim ADE, AED, BDF, BFD, CEF, CFE
    ADE = (180 - EDF) / 2
    AED = (180 - DEF) / 2
    BDF = (180 - EDF) / 2
    BFD = (180 - DFE) / 2
    CEF = (180 - DEF) / 2
    CFE = (180 - DFE) / 2

    Dim DAE, ECF, DBF
    DAE = 180 - ADE - AED
    ECF = 180 - CEF - CFE
    DBF = 180 - BDF - BFD

    Dim DG, GE, DE
    DG = VG * Sin(DVG * (Pi / 180)) / Sin(GDV * (Pi / 180))
    GE = VG * Sin(GVE * (Pi / 180)) / Sin(GEV * (Pi / 180))
    DE = DG + GE

    Dim DF, EF
    DF = DE * Sin(DEF * (Pi / 180)) / Sin(DFE * (Pi / 180))
    EF = DE * Sin(EDF * (Pi / 180)) / Sin(DFE * (Pi / 180))

    Dim BD, BF
    BD = DF * Sin(BFD * (Pi / 180)) / Sin(DBF * (Pi / 180))
    BF = DF * Sin(BDF * (Pi / 180)) / Sin(DBF * (Pi / 180))

    Dim CF, CE
    CF = EF * Sin(CEF * (Pi / 180)) / Sin(ECF * (Pi / 180))
    CE = EF * Sin(CFE * (Pi / 180)) / Sin(ECF * (Pi / 180))

    Dim AD, AE
    AD = DE * Sin(AED * (Pi / 180)) / Sin(DAE * (Pi / 180))
    AE = DE * Sin(ADE * (Pi / 180)) / Sin(DAE * (Pi / 180))

    Dim SEMIADE, SEMIBDF, SEMICEF
    SEMIADE = (DE + AD + AE) / 2
    SEMIBDF = (DF + BD + BF) / 2
    SEMICEF = (EF + CF + CE) / 2

    Dim AREAADE, AREABDF, AREACEF
    AREAADE = Sqr(SEMIADE * (SEMIADE - DE) * (SEMIADE - AD) * (SEMIADE - AE))
    AREABDF = Sqr(SEMIBDF * (SEMIBDF - DF) * (SEMIBDF - BD) * (SEMIBDF - BF))
    AREACEF = Sqr(SEMICEF * (SEMICEF - EF) * (SEMICEF - CF) * (SEMICEF - CE))
    
    Dim RA, RB, RC
    RA = AREAADE / (DE / 2)
    RB = AREABDF / (DF / 2)
    RC = AREACEF / (EF / 2)
    
    Dim RABC
    RABC = RA + RB + RC
        
    Dim V2
    V2 = (RABC - VG) / 4
    
    Dim V3
    V3 = V2 / 2
    
    S = M1 / magnitude1
    
    IsoCapRelative = (((RC / RABC) - (M1 / M123)) / (M1 / M123)) * S
    
    Dim Test
    Test = VG + ((V3 * 8) - RABC)
    
    If Not Test = 0 Then Error 2015
    
LastLine: End Function

Function IsoCapValue(magnitude1, magnitude2, magnitude3) As Double
Attribute IsoCapValue.VB_Description = "Copyright © 2004 by William J. McKibbin"

    Dim M1, M2, M3, M123
        
    M1 = Abs(magnitude1)
    M2 = Abs(magnitude2)
    M3 = Abs(magnitude3)
    M123 = M1 + M2 + M3
    
    If M1 = 0 Or M2 = 0 Or M3 = 0 Then
        IsoCapValue = magnitude1
        GoTo LastLine
        
    End If
    
    Dim Pi, VG
    Pi = Application.Pi()
    VG = 1

    Dim EDF, DEF, DFE
    EDF = 180 * (M1 / M123)
    DEF = 180 * (M2 / M123)
    DFE = 180 * (M3 / M123)
    
    Dim DGV, GDV, GEV
    DGV = 90
    GDV = EDF / 2
    GEV = DEF / 2

    Dim DVG, GVE
    DVG = 180 - DGV - GDV
    GVE = 180 - DGV - GEV

    Dim ADE, AED, BDF, BFD, CEF, CFE
    ADE = (180 - EDF) / 2
    AED = (180 - DEF) / 2
    BDF = (180 - EDF) / 2
    BFD = (180 - DFE) / 2
    CEF = (180 - DEF) / 2
    CFE = (180 - DFE) / 2

    Dim DAE, ECF, DBF
    DAE = 180 - ADE - AED
    ECF = 180 - CEF - CFE
    DBF = 180 - BDF - BFD

    Dim DG, GE, DE
    DG = VG * Sin(DVG * (Pi / 180)) / Sin(GDV * (Pi / 180))
    GE = VG * Sin(GVE * (Pi / 180)) / Sin(GEV * (Pi / 180))
    DE = DG + GE

    Dim DF, EF
    DF = DE * Sin(DEF * (Pi / 180)) / Sin(DFE * (Pi / 180))
    EF = DE * Sin(EDF * (Pi / 180)) / Sin(DFE * (Pi / 180))

    Dim BD, BF
    BD = DF * Sin(BFD * (Pi / 180)) / Sin(DBF * (Pi / 180))
    BF = DF * Sin(BDF * (Pi / 180)) / Sin(DBF * (Pi / 180))

    Dim CF, CE
    CF = EF * Sin(CEF * (Pi / 180)) / Sin(ECF * (Pi / 180))
    CE = EF * Sin(CFE * (Pi / 180)) / Sin(ECF * (Pi / 180))

    Dim AD, AE
    AD = DE * Sin(AED * (Pi / 180)) / Sin(DAE * (Pi / 180))
    AE = DE * Sin(ADE * (Pi / 180)) / Sin(DAE * (Pi / 180))

    Dim SEMIADE, SEMIBDF, SEMICEF
    SEMIADE = (DE + AD + AE) / 2
    SEMIBDF = (DF + BD + BF) / 2
    SEMICEF = (EF + CF + CE) / 2

    Dim AREAADE, AREABDF, AREACEF
    AREAADE = Sqr(SEMIADE * (SEMIADE - DE) * (SEMIADE - AD) * (SEMIADE - AE))
    AREABDF = Sqr(SEMIBDF * (SEMIBDF - DF) * (SEMIBDF - BD) * (SEMIBDF - BF))
    AREACEF = Sqr(SEMICEF * (SEMICEF - EF) * (SEMICEF - CF) * (SEMICEF - CE))
    
    Dim RA, RB, RC
    RA = AREAADE / (DE / 2)
    RB = AREABDF / (DF / 2)
    RC = AREACEF / (EF / 2)
    
    Dim RABC
    RABC = RA + RB + RC
        
    Dim V2
    V2 = (RABC - VG) / 4
    
    Dim V3
    V3 = V2 / 2
            
    IsoCapValue = (((RC / RABC) * (M123)) / M1) * magnitude1
         
    Dim Test
    Test = VG + ((V3 * 8) - RABC)
    
    If Not Test = 0 Then Error 2015

LastLine: End Function

Function IsoCapHarm(magnitude1, magnitude2, magnitude3) As Double
Attribute IsoCapHarm.VB_Description = "Copyright © 2004 by William J. McKibbin"
        
    Dim M1, M2, M3, M123, S
        
    M1 = Abs(magnitude1)
    M2 = Abs(magnitude2)
    M3 = Abs(magnitude3)
    M123 = M1 + M2 + M3
    
    If M1 = 0 And M2 = 0 And M3 = 0 Then
        IsoCapHarm = 0
        GoTo LastLine
        
    ElseIf M1 > 0 And M2 = 0 And M3 = 0 Then
        S = M1 / magnitude1
        IsoCapHarm = 1 * S
        GoTo LastLine
        
    ElseIf M1 = 0 And M2 > 0 And M3 = 0 Then
        S = M2 / magnitude2
        IsoCapHarm = 1 * S
        GoTo LastLine
        
    ElseIf M1 = 0 And M2 = 0 And M3 > 0 Then
        S = M3 / magnitude3
        IsoCapHarm = 1 * S
        GoTo LastLine
        
    ElseIf M1 = 0 And M2 > 0 And M3 > 0 Then
        S = (M2 / magnitude2) * (M3 / magnitude3)
        IsoCapHarm = 1 * S
        GoTo LastLine
       
    ElseIf M1 > 0 And M2 = 0 And M3 > 0 Then
        S = (M1 / magnitude1) * (M3 / magnitude3)
        IsoCapHarm = 1 * S
        GoTo LastLine
        
    ElseIf M1 > 0 And M2 > 0 And M3 = 0 Then
        S = (M1 / magnitude1) * (M2 / magnitude2)
        IsoCapHarm = 1 * S
        GoTo LastLine
        
    End If
    
    Dim Pi, VG
    Pi = Application.Pi()
    VG = 1

    Dim EDF, DEF, DFE
    EDF = 180 * (M1 / M123)
    DEF = 180 * (M2 / M123)
    DFE = 180 * (M3 / M123)

    Dim DGV, GDV, GEV
    DGV = 90
    GDV = EDF / 2
    GEV = DEF / 2

    Dim DVG, GVE
    DVG = 180 - DGV - GDV
    GVE = 180 - DGV - GEV

    Dim ADE, AED, BDF, BFD, CEF, CFE
    ADE = (180 - EDF) / 2
    AED = (180 - DEF) / 2
    BDF = (180 - EDF) / 2
    BFD = (180 - DFE) / 2
    CEF = (180 - DEF) / 2
    CFE = (180 - DFE) / 2

    Dim DAE, ECF, DBF
    DAE = 180 - ADE - AED
    ECF = 180 - CEF - CFE
    DBF = 180 - BDF - BFD

    Dim DG, GE, DE
    DG = VG * Sin(DVG * (Pi / 180)) / Sin(GDV * (Pi / 180))
    GE = VG * Sin(GVE * (Pi / 180)) / Sin(GEV * (Pi / 180))
    DE = DG + GE

    Dim DF, EF
    DF = DE * Sin(DEF * (Pi / 180)) / Sin(DFE * (Pi / 180))
    EF = DE * Sin(EDF * (Pi / 180)) / Sin(DFE * (Pi / 180))

    Dim BD, BF
    BD = DF * Sin(BFD * (Pi / 180)) / Sin(DBF * (Pi / 180))
    BF = DF * Sin(BDF * (Pi / 180)) / Sin(DBF * (Pi / 180))

    Dim CF, CE
    CF = EF * Sin(CEF * (Pi / 180)) / Sin(ECF * (Pi / 180))
    CE = EF * Sin(CFE * (Pi / 180)) / Sin(ECF * (Pi / 180))

    Dim AD, AE
    AD = DE * Sin(AED * (Pi / 180)) / Sin(DAE * (Pi / 180))
    AE = DE * Sin(ADE * (Pi / 180)) / Sin(DAE * (Pi / 180))

    Dim SEMIADE, SEMIBDF, SEMICEF
    SEMIADE = (DE + AD + AE) / 2
    SEMIBDF = (DF + BD + BF) / 2
    SEMICEF = (EF + CF + CE) / 2

    Dim AREAADE, AREABDF, AREACEF
    AREAADE = Sqr(SEMIADE * (SEMIADE - DE) * (SEMIADE - AD) * (SEMIADE - AE))
    AREABDF = Sqr(SEMIBDF * (SEMIBDF - DF) * (SEMIBDF - BD) * (SEMIBDF - BF))
    AREACEF = Sqr(SEMICEF * (SEMICEF - EF) * (SEMICEF - CF) * (SEMICEF - CE))
    
    Dim RA, RB, RC
    RA = AREAADE / (DE / 2)
    RB = AREABDF / (DF / 2)
    RC = AREACEF / (EF / 2)
    
    Dim RABC
    RABC = RA + RB + RC
        
    Dim V2
    V2 = (RABC - VG) / 4
    
    Dim V3
    V3 = V2 / 2
    
    S = (M1 / magnitude1) * (M2 / magnitude2) * (M3 / magnitude3)
    
    IsoCapHarm = (3 / ((1 / (((RC / RABC) * M123) / M1)) + (1 / (((RB / RABC) * M123) / M2)) + (1 / (((RA / RABC) * M123) / M3)))) * S

    Dim Test
    Test = VG + ((V3 * 8) - RABC)
    
    If Not Test = 0 Then Error 2015
          
LastLine: End Function
