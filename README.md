isocapmod
=========

IsoCapMod Functions

Programming code implementing the IsoCapMod functions as Microsoft Visual Basic for Applications (Version 6.3) user-defined spreadsheet functions in Microsoft Excel (Version 11.0) appears at Appendix B.  Procedures for entering the code into Microsoft Excel are described in Excel Programming (Simon, 2002, pp. 52-53) and the Definitive Guide to Excel VBA (Kofler, 2000, pp. 30-31, 696-701).  The functions require that the user know the CapMagf, CapMags, and CapMagh rates.

The IsoCapRatio function returns the IsoCapMod ratio for magnitude1 where magnitude1, magnitude2, and magnitude3 represent CapMagf, CapMags, and CapMagh, in any order.  The syntax for entering the function into a worksheet is

=IsoCapRatio(magnitude1,magnitude2,magnitude3)

The IsoCapCoef function returns the IsoCapMod coefficient of magnitude1 where magnitude1, magnitude2, and magnitude3 represent CapMagf, CapMags, and CapMagh, in any order.  The syntax for entering the function into a worksheet is

=IsoCapCoef(magnitude1,magnitude2,magnitude3)

The IsoCapValue function returns the modulated intensity of magnitude1 where magnitude1, magnitude2, and magnitude3 represent CapMagf, CapMags, and CapMagh, in any order.  The syntax for entering the function into a worksheet is

=IsoCapValue(magnitude1,magnitude2,magnitude3

The IsoCapChange function returns the absolute degree of IsoCapMod change for magnitude1 where magnitude1, magnitude2, and magnitude3 represent CapMagf, CapMags, and CapMagh, in any order.  The syntax for entering the function into a worksheet is

=IsoCapChange(magnitude1,magnitude2,magnitude3)

The IsoCapRelative function returns the relative degree of IsoCapMod change for magnitude1 where magnitude1, magnitude2, and magnitude3 represent CapMagf, CapMags, and CapMagh, in any order.  The The syntax for entering the function into a worksheet is

=IsoCapRelative(magnitude1,magnitude2,magnitude3)

The IsoCapSynergy function returns the IsoCapMod synergy score as a conjoint of magnitude1, magnitude2, and magnitude3.  As with the other functions, magnitude1, magnitude2, and magnitude3 represent CapMagf, CapMags, and CapMagh, in any order.  The syntax for entering the function into a worksheet is

=IsoCapSynergy(magnitude1,magnitude2,magnitude3)

The IsoCapHarm function returns the harmonic mean of IsoCapCoeff, IsoCapCoefs, and IsoCapCoefh.  As with the other functions, magnitude1, magnitude2, and magnitude3 represent CapMagf, CapMags, and CapMagh, in any order.  The syntax for entering the function into a worksheet is

=IsoCapHarm(magnitude1,magnitude2,magnitude3)
