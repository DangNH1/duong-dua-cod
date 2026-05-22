// ==========================================================================
// COD RACE MASTER - CORE APPLICATION LOGIC
// Includes: Aggregated data, Chart rendering, Excel parser, and Recommender
// ==========================================================================

const appData = {
    "selectedDate":  "21/05/2026",
    "history":  {
                    "01/05/2026":  [
                                       {
                                           "TenVung":  "ÄCL",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "BC KD",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "BTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DBB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DNB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DSH",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "GXT",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "HCM",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "HNO",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "KHL",
                                           "TyLeTuNop":  0
                                       },
                                       {
                                           "TenVung":  "NTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TBB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNG",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNT",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "XBG",
                                           "TyLeTuNop":  100
                                       }
                                   ],
                    "02/05/2026":  [
                                       {
                                           "TenVung":  "ÄCL",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "BC KD",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "BTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DBB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DNB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DSH",
                                           "TyLeTuNop":  0
                                       },
                                       {
                                           "TenVung":  "GXT",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "HCM",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "HNO",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "KHL",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "NTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TBB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNG",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNT",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "XBG",
                                           "TyLeTuNop":  100
                                       }
                                   ],
                    "03/05/2026":  [
                                       {
                                           "TenVung":  "ÄCL",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "BC KD",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "BTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DBB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DNB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DSH",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "GXT",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "HCM",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "HNO",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "KHL",
                                           "TyLeTuNop":  0
                                       },
                                       {
                                           "TenVung":  "NTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TBB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNG",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNT",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "XBG",
                                           "TyLeTuNop":  100
                                       }
                                   ],
                    "04/05/2026":  [
                                       {
                                           "TenVung":  "ÄCL",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "BC KD",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "BTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DBB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DNB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DSH",
                                           "TyLeTuNop":  0
                                       },
                                       {
                                           "TenVung":  "GXT",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "HCM",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "HNO",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "KHL",
                                           "TyLeTuNop":  0
                                       },
                                       {
                                           "TenVung":  "NTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TBB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNG",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNT",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "XBG",
                                           "TyLeTuNop":  100
                                       }
                                   ],
                    "05/05/2026":  [
                                       {
                                           "TenVung":  "ÄCL",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "BC KD",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "BTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DBB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DNB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DSH",
                                           "TyLeTuNop":  0
                                       },
                                       {
                                           "TenVung":  "GXT",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "HCM",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "HNO",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "KHL",
                                           "TyLeTuNop":  0
                                       },
                                       {
                                           "TenVung":  "NTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TBB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNG",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNT",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "XBG",
                                           "TyLeTuNop":  100
                                       }
                                   ],
                    "06/05/2026":  [
                                       {
                                           "TenVung":  "ÄCL",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "BC KD",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "BTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DBB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DNB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DSH",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "GXT",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "HCM",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "HNO",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "KHL",
                                           "TyLeTuNop":  0
                                       },
                                       {
                                           "TenVung":  "NTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TBB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNG",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNT",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "XBG",
                                           "TyLeTuNop":  100
                                       }
                                   ],
                    "07/05/2026":  [
                                       {
                                           "TenVung":  "ÄCL",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "BC KD",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "BTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DBB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DNB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DSH",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "GXT",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "HCM",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "HNO",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "KHL",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "NTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TBB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNG",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNT",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "XBG",
                                           "TyLeTuNop":  100
                                       }
                                   ],
                    "08/05/2026":  [
                                       {
                                           "TenVung":  "ÄCL",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "BC KD",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "BTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DBB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DNB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DSH",
                                           "TyLeTuNop":  0
                                       },
                                       {
                                           "TenVung":  "GXT",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "HCM",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "HNO",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "KHL",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "NTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TBB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNG",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNT",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "XBG",
                                           "TyLeTuNop":  100
                                       }
                                   ],
                    "09/05/2026":  [
                                       {
                                           "TenVung":  "ÄCL",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "BC KD",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "BTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DBB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DNB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DSH",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "GXT",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "HCM",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "HNO",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "KHL",
                                           "TyLeTuNop":  0
                                       },
                                       {
                                           "TenVung":  "NTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TBB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNG",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNT",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "XBG",
                                           "TyLeTuNop":  100
                                       }
                                   ],
                    "10/05/2026":  [
                                       {
                                           "TenVung":  "ÄCL",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "BC KD",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "BTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DBB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DNB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DSH",
                                           "TyLeTuNop":  0
                                       },
                                       {
                                           "TenVung":  "GXT",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "HCM",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "HNO",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "KHL",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "NTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TBB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNG",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNT",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "XBG",
                                           "TyLeTuNop":  100
                                       }
                                   ],
                    "11/05/2026":  [
                                       {
                                           "TenVung":  "ÄCL",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "BC KD",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "BTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DBB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DNB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DSH",
                                           "TyLeTuNop":  0
                                       },
                                       {
                                           "TenVung":  "GXT",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "HCM",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "HNO",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "KHL",
                                           "TyLeTuNop":  0
                                       },
                                       {
                                           "TenVung":  "NTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TBB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNG",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNT",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "XBG",
                                           "TyLeTuNop":  100
                                       }
                                   ],
                    "12/05/2026":  [
                                       {
                                           "TenVung":  "ÄCL",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "BC KD",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "BTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DBB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DNB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DSH",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "GXT",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "HCM",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "HNO",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "KHL",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "NTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TBB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNG",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNT",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "XBG",
                                           "TyLeTuNop":  100
                                       }
                                   ],
                    "13/05/2026":  [
                                       {
                                           "TenVung":  "ÄCL",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "BC KD",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "BTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DBB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DNB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DSH",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "GXT",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "HCM",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "HNO",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "KHL",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "NTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TBB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNG",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNT",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "XBG",
                                           "TyLeTuNop":  100
                                       }
                                   ],
                    "14/05/2026":  [
                                       {
                                           "TenVung":  "ÄCL",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "BC KD",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "BTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DBB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DNB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DSH",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "GXT",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "HCM",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "HNO",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "KHL",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "NTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TBB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNG",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNT",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "XBG",
                                           "TyLeTuNop":  100
                                       }
                                   ],
                    "15/05/2026":  [
                                       {
                                           "TenVung":  "ÄCL",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "BC KD",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "BTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DBB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DNB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DSH",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "GXT",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "HCM",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "HNO",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "KHL",
                                           "TyLeTuNop":  0
                                       },
                                       {
                                           "TenVung":  "NTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TBB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNG",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNT",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "XBG",
                                           "TyLeTuNop":  100
                                       }
                                   ],
                    "16/05/2026":  [
                                       {
                                           "TenVung":  "ÄCL",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "BC KD",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "BTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DBB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DNB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DSH",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "GXT",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "HCM",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "HNO",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "KHL",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "NTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TBB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNG",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNT",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "XBG",
                                           "TyLeTuNop":  100
                                       }
                                   ],
                    "17/05/2026":  [
                                       {
                                           "TenVung":  "ÄCL",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "BC KD",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "BTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DBB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DNB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DSH",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "GXT",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "HCM",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "HNO",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "KHL",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "NTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TBB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNG",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNT",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "XBG",
                                           "TyLeTuNop":  100
                                       }
                                   ],
                    "18/05/2026":  [
                                       {
                                           "TenVung":  "ÄCL",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "BC KD",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "BTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DBB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DNB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DSH",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "GXT",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "HCM",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "HNO",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "KHL",
                                           "TyLeTuNop":  0
                                       },
                                       {
                                           "TenVung":  "NTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TBB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNG",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNT",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "XBG",
                                           "TyLeTuNop":  100
                                       }
                                   ],
                    "19/05/2026":  [
                                       {
                                           "TenVung":  "ÄCL",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "BC KD",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "BTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DBB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DNB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DSH",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "GXT",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "HCM",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "HNO",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "KHL",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "NTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TBB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNG",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNT",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "XBG",
                                           "TyLeTuNop":  100
                                       }
                                   ],
                    "20/05/2026":  [
                                       {
                                           "TenVung":  "ÄCL",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "BC KD",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "BTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DBB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DNB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DSH",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "GXT",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "HCM",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "HNO",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "KHL",
                                           "TyLeTuNop":  0
                                       },
                                       {
                                           "TenVung":  "NTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TBB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNG",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNT",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "XBG",
                                           "TyLeTuNop":  100
                                       }
                                   ],
                    "21/05/2026":  [
                                       {
                                           "TenVung":  "ÄCL",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "BC KD",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "BTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DBB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DNB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "DSH",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "GXT",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "HCM",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "HNO",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "KHL",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "NTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TBB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNG",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TNT",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "TTB",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "XBG",
                                           "TyLeTuNop":  100
                                       }
                                   ]
                },
    "regionNames":  [
                        "ÄCL",
                        "BC KD",
                        "BTB",
                        "DBB",
                        "DNB",
                        "DSH",
                        "GXT",
                        "HCM",
                        "HNO",
                        "KHL",
                        "NTB",
                        "TBB",
                        "TNB",
                        "TNG",
                        "TNT",
                        "TTB",
                        "XBG"
                    ],
    "dateRange":  [
                      "01/05/2026",
                      "02/05/2026",
                      "03/05/2026",
                      "04/05/2026",
                      "05/05/2026",
                      "06/05/2026",
                      "07/05/2026",
                      "08/05/2026",
                      "09/05/2026",
                      "10/05/2026",
                      "11/05/2026",
                      "12/05/2026",
                      "13/05/2026",
                      "14/05/2026",
                      "15/05/2026",
                      "16/05/2026",
                      "17/05/2026",
                      "18/05/2026",
                      "19/05/2026",
                      "20/05/2026",
                      "21/05/2026"
                  ],
    "regions":  [
                    {
                        "TyLeTuNop":  100,
                        "TenVung":  "ÄCL",
                        "TongCOD":  2520.78,
                        "Trend":  0,
                        "CODTuNop":  2520.78,
                        "BuuCucList":  [
                                           {
                                               "TongCOD":  42.97,
                                               "CODTuNop":  42.97,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 07 Lê Lợi-Mang Thít-Vĩnh Long"
                                           },
                                           {
                                               "TongCOD":  62.74,
                                               "CODTuNop":  62.74,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 12-14 Bà Triệu-Cao Lãnh-Đồng Tháp"
                                           },
                                           {
                                               "TongCOD":  64.99,
                                               "CODTuNop":  64.99,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 19 Đường DT848-Sa Đéc-Đồng Tháp"
                                           },
                                           {
                                               "TongCOD":  55.26,
                                               "CODTuNop":  55.26,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 206 A Đoàn Thị Nghiệp-Mỹ Tho-Tiền Giang 01"
                                           },
                                           {
                                               "TongCOD":  79.37,
                                               "CODTuNop":  79.37,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 230 Ấp Bình Hạnh-Long Bình Điền-Chợ Gạo-Tiền Giang"
                                           },
                                           {
                                               "TongCOD":  28.92,
                                               "CODTuNop":  28.92,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 286 Nguyễn Huệ-Tân Hồng-Đồng Tháp"
                                           },
                                           {
                                               "TongCOD":  47.66,
                                               "CODTuNop":  47.66,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 3/2 Thị Trấn Châu Thành-Bến Tre"
                                           },
                                           {
                                               "TongCOD":  12.82,
                                               "CODTuNop":  12.82,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 355 Võ Tấn Đức-Tam Bình-Vĩnh Long"
                                           },
                                           {
                                               "TongCOD":  64.1,
                                               "CODTuNop":  64.1,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 57/4 Phan Văn Đáng-Vĩnh Long-Vĩnh Long"
                                           },
                                           {
                                               "TongCOD":  3.28,
                                               "CODTuNop":  3.28,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 73 Phó Cơ Điều-Phường Phước Hậu-Vĩnh Long"
                                           },
                                           {
                                               "TongCOD":  30.56,
                                               "CODTuNop":  30.56,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 79 Đường Tỉnh 871-Xã Gia Thuận-Tiền Giang"
                                           },
                                           {
                                               "TongCOD":  34.95,
                                               "CODTuNop":  34.95,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 87 Lê Văn Vị-Cái Vồn-Bình Minh-Vĩnh Long"
                                           },
                                           {
                                               "TongCOD":  73.11,
                                               "CODTuNop":  73.11,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 88 Ấp Thân Hòa-Thân Cửu Nghĩa-Châu Thành-Tiền Giang"
                                           },
                                           {
                                               "TongCOD":  176.6,
                                               "CODTuNop":  176.6,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 888 QL1A-Phường 2-Cai Lậy-Tiền Giang"
                                           },
                                           {
                                               "TongCOD":  37.52,
                                               "CODTuNop":  37.52,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 992 Đường Huyện 35-Vĩnh Kim-Châu Thành-Tiền Giang"
                                           },
                                           {
                                               "TongCOD":  25.07,
                                               "CODTuNop":  25.07,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Ấp 2-Phú Nhuận-Bến Tre"
                                           },
                                           {
                                               "TongCOD":  27.75,
                                               "CODTuNop":  27.75,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu cục Ấp 5A Thị Trấn Giồng Trôm-Giồng Trôm-Bến Tre"
                                           },
                                           {
                                               "TongCOD":  20.18,
                                               "CODTuNop":  20.18,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Ấp Bình Hiệp A-Xã Lấp Vò-Đồng Tháp"
                                           },
                                           {
                                               "TongCOD":  29.96,
                                               "CODTuNop":  29.96,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Ấp Cây Hẹ-Phú Cần-Tiểu Cần-Trà Vinh"
                                           },
                                           {
                                               "TongCOD":  18.8,
                                               "CODTuNop":  18.8,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Ấp Hậu Mỹ Phú 1-Cái Bè-Tiền Giang"
                                           },
                                           {
                                               "TongCOD":  3.49,
                                               "CODTuNop":  3.49,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Ấp Phú Long-Xã Hưng Khánh Trung-Bến Tre"
                                           },
                                           {
                                               "TongCOD":  33.53,
                                               "CODTuNop":  33.53,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Ấp Tân Thuận-Bình Tân-Vĩnh Long"
                                           },
                                           {
                                               "TongCOD":  17.22,
                                               "CODTuNop":  17.22,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục DT846 Phương Trà-Xã Ba Sao-Đồng Tháp"
                                           },
                                           {
                                               "TongCOD":  116.5,
                                               "CODTuNop":  116.5,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Đường Tỉnh 861-Xã Mỹ Lợi-Tiền Giang"
                                           },
                                           {
                                               "TongCOD":  204.21,
                                               "CODTuNop":  204.21,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Duyên Hải-Trà Vinh"
                                           },
                                           {
                                               "TongCOD":  126.32,
                                               "CODTuNop":  126.32,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hồng Ngự-Đồng Tháp"
                                           },
                                           {
                                               "TongCOD":  185.5,
                                               "CODTuNop":  185.5,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Khóm 3 Trần Hưng Đạo-Xã Tháp Mười-Đồng Tháp"
                                           },
                                           {
                                               "TongCOD":  33.88,
                                               "CODTuNop":  33.88,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Khóm 3 TT Càng Long-Xã Càng Long-Trà Vinh"
                                           },
                                           {
                                               "TongCOD":  26.45,
                                               "CODTuNop":  26.45,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Khóm 4 TT Châu Thành-Xã Châu Thành-Trà Vinh"
                                           },
                                           {
                                               "TongCOD":  16.72,
                                               "CODTuNop":  16.72,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Khóm 8-Cầu Kè-Trà Vinh"
                                           },
                                           {
                                               "TongCOD":  24.78,
                                               "CODTuNop":  24.78,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Khóm Thống Nhất-Xã Cầu Ngang-Trà Vinh"
                                           },
                                           {
                                               "TongCOD":  52.03,
                                               "CODTuNop":  52.03,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Khu phố 1-Tân Phước-Tiền Giang"
                                           },
                                           {
                                               "TongCOD":  37.48,
                                               "CODTuNop":  37.48,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục KP2-Thị Trấn Ba Tri-Bến Tre"
                                           },
                                           {
                                               "TongCOD":  37.43,
                                               "CODTuNop":  37.43,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục KP3 Phường 3-Gò Công-Tiền Giang"
                                           },
                                           {
                                               "TongCOD":  75.61,
                                               "CODTuNop":  75.61,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Long Trung-Cai Lậy-Tiền Giang"
                                           },
                                           {
                                               "TongCOD":  17.02,
                                               "CODTuNop":  17.02,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Mỏ Cày Nam-Bến Tre"
                                           },
                                           {
                                               "TongCOD":  17.17,
                                               "CODTuNop":  17.17,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Mỹ Thọ-Trần Quang Khải-Cao Lãnh-Đồng Tháp"
                                           },
                                           {
                                               "TongCOD":  20.9,
                                               "CODTuNop":  20.9,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Phước Hậu-Phước Mỹ Trung-Bến Tre"
                                           },
                                           {
                                               "TongCOD":  35.97,
                                               "CODTuNop":  35.97,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Phước Long-Lộc Hòa-Long Hồ-Vĩnh Long"
                                           },
                                           {
                                               "TongCOD":  29.17,
                                               "CODTuNop":  29.17,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục QL 30-Thanh Bình-Đồng Tháp"
                                           },
                                           {
                                               "TongCOD":  60.59,
                                               "CODTuNop":  60.59,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục QL1A Hòa Khánh-Cái Bè-Tiền Giang"
                                           },
                                           {
                                               "TongCOD":  41.02,
                                               "CODTuNop":  41.02,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục QL50 Phường 9-Mỹ Tho-Tiền Giang"
                                           },
                                           {
                                               "TongCOD":  1.59,
                                               "CODTuNop":  1.59,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục QL57 KP3-Thị Trấn Chợ Lách-Bến Tre"
                                           },
                                           {
                                               "TongCOD":  10.3,
                                               "CODTuNop":  10.3,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục QL57-Bình Thạnh-Thạnh Phú-Bến Tre"
                                           },
                                           {
                                               "TongCOD":  11.57,
                                               "CODTuNop":  11.57,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục QL57-Bình Thới-Bình Đại-Bến Tre"
                                           },
                                           {
                                               "TongCOD":  34.24,
                                               "CODTuNop":  34.24,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục QL80-Xã Hòa Long-Đồng Tháp"
                                           },
                                           {
                                               "TongCOD":  67.13,
                                               "CODTuNop":  67.13,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Quốc Lộ 50-Gò Công Tây-Tiền Giang"
                                           },
                                           {
                                               "TongCOD":  17.83,
                                               "CODTuNop":  17.83,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Quốc Lộ 53-Long Hồ-Vĩnh Long"
                                           },
                                           {
                                               "TongCOD":  19.43,
                                               "CODTuNop":  19.43,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Quốc Lộ 54-Xã Lai Vung-Đồng Tháp"
                                           },
                                           {
                                               "TongCOD":  31,
                                               "CODTuNop":  31,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Sơn Thông-Phường 9-Tp.Trà Vinh-Trà Vinh"
                                           },
                                           {
                                               "TongCOD":  29.39,
                                               "CODTuNop":  29.39,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Tam Nông-Đồng Tháp"
                                           },
                                           {
                                               "TongCOD":  44.48,
                                               "CODTuNop":  44.48,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Tân Nhuận Đông-Châu Thành-Đồng Tháp"
                                           },
                                           {
                                               "TongCOD":  3.71,
                                               "CODTuNop":  3.71,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Tỉnh Lộ DT848-Xã Mỹ An Hưng-Đồng Tháp"
                                           },
                                           {
                                               "TongCOD":  23.87,
                                               "CODTuNop":  23.87,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Tổ 1 Đường 30/4-Trà Ôn-Vĩnh Long"
                                           },
                                           {
                                               "TongCOD":  47.14,
                                               "CODTuNop":  47.14,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Võ Nguyên Giáp-Phú Tân-Bến Tre"
                                           },
                                           {
                                               "TongCOD":  29.51,
                                               "CODTuNop":  29.51,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Xoài Xiêm-Ngãi Xuyên-Trà Cú-Trà Vinh"
                                           }
                                       ]
                    },
                    {
                        "TyLeTuNop":  100,
                        "TenVung":  "BC KD",
                        "TongCOD":  2.56,
                        "Trend":  0,
                        "CODTuNop":  2.56,
                        "BuuCucList":  [
                                           {
                                               "TongCOD":  1.57,
                                               "CODTuNop":  1.57,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục KD 1196 Quang Trung-Q.Gò Vấp-HCM"
                                           },
                                           {
                                               "TongCOD":  0.98,
                                               "CODTuNop":  0.98,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục KD 93D Tân Thắng-Q.Tân Phú-HCM"
                                           }
                                       ]
                    },
                    {
                        "TyLeTuNop":  100,
                        "TenVung":  "BTB",
                        "TongCOD":  4468.15,
                        "Trend":  0,
                        "CODTuNop":  4468.15,
                        "BuuCucList":  [
                                           {
                                               "TongCOD":  98.04,
                                               "CODTuNop":  98.04,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 100 Lê Lợi-Thường Xuân-Thanh Hóa"
                                           },
                                           {
                                               "TongCOD":  53.96,
                                               "CODTuNop":  53.96,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 104B đường Quế Sơn-Thọ Xuân-Thanh Hóa"
                                           },
                                           {
                                               "TongCOD":  153.22,
                                               "CODTuNop":  153.22,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 162 Đường Thành Thái-Đông Thọ-Thanh Hóa"
                                           },
                                           {
                                               "TongCOD":  67.69,
                                               "CODTuNop":  67.69,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 168 đường 1/9-Thanh Chương-Nghệ An"
                                           },
                                           {
                                               "TongCOD":  91.24,
                                               "CODTuNop":  91.24,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 18 TDP Phong Ý-Cẩm Thuỷ-Thanh Hoá"
                                           },
                                           {
                                               "TongCOD":  102.78,
                                               "CODTuNop":  102.78,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 195 Nguyễn Huệ-Phú Sơn-Bỉm Sơn-Thanh Hóa"
                                           },
                                           {
                                               "TongCOD":  20.01,
                                               "CODTuNop":  20.01,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 196 Khu 1-Vĩnh Lộc-Thanh Hóa"
                                           },
                                           {
                                               "TongCOD":  62.37,
                                               "CODTuNop":  62.37,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 242 Lê Lợi-Hương Sơn-Hà Tĩnh"
                                           },
                                           {
                                               "TongCOD":  162.86,
                                               "CODTuNop":  162.86,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 265 Tiểu Khu 2 QL45-Thiệu Hóa-Thanh Hóa"
                                           },
                                           {
                                               "TongCOD":  54.91,
                                               "CODTuNop":  54.91,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 286 Lê Lợi-Trường Sơn-Sầm Sơn-Thanh Hóa"
                                           },
                                           {
                                               "TongCOD":  64.97,
                                               "CODTuNop":  64.97,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 292A Lý Tự Trọng-Thạch Hà-Hà Tĩnh"
                                           },
                                           {
                                               "TongCOD":  70.93,
                                               "CODTuNop":  70.93,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 296 Trần Phú-Đức Thuận-Hồng Lĩnh-Hà Tĩnh"
                                           },
                                           {
                                               "TongCOD":  100.76,
                                               "CODTuNop":  100.76,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 339 Đường Lam Sơn-Nông Cống-Thanh Hóa"
                                           },
                                           {
                                               "TongCOD":  39.99,
                                               "CODTuNop":  39.99,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 357 Khối Hòa Tây-Tương Dương-Nghệ An"
                                           },
                                           {
                                               "TongCOD":  94.32,
                                               "CODTuNop":  94.32,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 415 Hà Huy Tập-Cẩm Xuyên-Hà Tĩnh"
                                           },
                                           {
                                               "TongCOD":  150.33,
                                               "CODTuNop":  150.33,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 5 Võ Thúc Đồng-Trường Vinh-Nghệ An"
                                           },
                                           {
                                               "TongCOD":  119.21,
                                               "CODTuNop":  119.21,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 509 Trần Phú-TP.Hà Tĩnh"
                                           },
                                           {
                                               "TongCOD":  32.06,
                                               "CODTuNop":  32.06,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 558 đường Tố Hữu-Quảng Xương-Thanh Hóa"
                                           },
                                           {
                                               "TongCOD":  108.2,
                                               "CODTuNop":  108.2,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 638 Phố Cống-Ngọc Lặc-Thanh Hóa"
                                           },
                                           {
                                               "TongCOD":  61.72,
                                               "CODTuNop":  61.72,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 702 Lý Thái Tổ-Kỳ Anh-Hà Tĩnh"
                                           },
                                           {
                                               "TongCOD":  38.23,
                                               "CODTuNop":  38.23,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 74 Hai Sáu Tháng Ba-Thành Sen-Hà Tĩnh"
                                           },
                                           {
                                               "TongCOD":  52.11,
                                               "CODTuNop":  52.11,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu cục 794 Quang Trung-Thanh Hóa-Thanh Hóa"
                                           },
                                           {
                                               "TongCOD":  52.15,
                                               "CODTuNop":  52.15,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Can Lộc-Hà Tĩnh"
                                           },
                                           {
                                               "TongCOD":  38.44,
                                               "CODTuNop":  38.44,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Cầu Mưng-Hưng Nguyên-Nghệ An"
                                           },
                                           {
                                               "TongCOD":  74.01,
                                               "CODTuNop":  74.01,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục DT9 Thạch Châu-Lộc Hà-Hà Tĩnh"
                                           },
                                           {
                                               "TongCOD":  38.08,
                                               "CODTuNop":  38.08,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Đường 10 tháng 6 Khu 4-Yên Định-Thanh Hóa"
                                           },
                                           {
                                               "TongCOD":  275.6,
                                               "CODTuNop":  275.6,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hưng Phúc-Vinh-Nghệ An"
                                           },
                                           {
                                               "TongCOD":  55.24,
                                               "CODTuNop":  55.24,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Khối 1-Thị trấn Con Cuông-Huyện Con Cuông-Nghệ An"
                                           },
                                           {
                                               "TongCOD":  9.05,
                                               "CODTuNop":  9.05,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Khối 2-Yên Thành-Nghệ An"
                                           },
                                           {
                                               "TongCOD":  10.72,
                                               "CODTuNop":  10.72,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Khối 3-Tân Kỳ-Nghệ An"
                                           },
                                           {
                                               "TongCOD":  71.66,
                                               "CODTuNop":  71.66,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Khối 5-Quỳnh Xuân- Tx.Hoàng Mai-Nghệ An"
                                           },
                                           {
                                               "TongCOD":  95.9,
                                               "CODTuNop":  95.9,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Khối Đông Sơn-Quế Phong-Nghệ An"
                                           },
                                           {
                                               "TongCOD":  85.24,
                                               "CODTuNop":  85.24,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Khối Tân Minh-Nghĩa Đàn-Nghệ An"
                                           },
                                           {
                                               "TongCOD":  93.16,
                                               "CODTuNop":  93.16,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Khu 1 Kim Tân-Thạch Thành-Thanh Hoá"
                                           },
                                           {
                                               "TongCOD":  22.81,
                                               "CODTuNop":  22.81,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu cục Khu 4-Quan Hoá-Thanh Hoá"
                                           },
                                           {
                                               "TongCOD":  152.83,
                                               "CODTuNop":  152.83,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Khu Phố 1-TP.Thanh Hóa-Thanh Hóa"
                                           },
                                           {
                                               "TongCOD":  33.29,
                                               "CODTuNop":  33.29,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Khu Phố Mỹ Ré-Như Xuân-Thanh Hoá"
                                           },
                                           {
                                               "TongCOD":  141.36,
                                               "CODTuNop":  141.36,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Nghĩa Đàn-Nghệ An"
                                           },
                                           {
                                               "TongCOD":  36.54,
                                               "CODTuNop":  36.54,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Nguyễn Nhữ Soạn-Rừng Thông-Đông Sơn-Thanh Hóa"
                                           },
                                           {
                                               "TongCOD":  1.48,
                                               "CODTuNop":  1.48,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Phan Đình Phùng-TT.Đức Thọ-Hà Tĩnh"
                                           },
                                           {
                                               "TongCOD":  184.59,
                                               "CODTuNop":  184.59,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Phố Trung Sơn-Xã Hoằng Hóa-Thanh Hóa"
                                           },
                                           {
                                               "TongCOD":  142.7,
                                               "CODTuNop":  142.7,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Phường Hà Huy Tập-Vinh-Nghệ An"
                                           },
                                           {
                                               "TongCOD":  38.43,
                                               "CODTuNop":  38.43,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục QL1A Đồng Tiến-Kỳ Đồng-Kỳ Anh-Hà Tĩnh"
                                           },
                                           {
                                               "TongCOD":  49.4,
                                               "CODTuNop":  49.4,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục QL1A Kỳ Thư-Lê Đại Hành-Kỳ Anh-Hà Tĩnh"
                                           },
                                           {
                                               "TongCOD":  121.39,
                                               "CODTuNop":  121.39,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục QL8B Nguyễn Nhiễm-Xã Nghi Xuân-Hà Tĩnh"
                                           },
                                           {
                                               "TongCOD":  126.58,
                                               "CODTuNop":  126.58,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Quang Trung-Vinh-Nghệ An"
                                           },
                                           {
                                               "TongCOD":  37.59,
                                               "CODTuNop":  37.59,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Quốc Lộ 12C-Xã Kỳ Lạc-Hà Tĩnh"
                                           },
                                           {
                                               "TongCOD":  124.54,
                                               "CODTuNop":  124.54,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Số 25 Khu Phú Cường-Hậu Lộc-Thanh Hóa"
                                           },
                                           {
                                               "TongCOD":  39.18,
                                               "CODTuNop":  39.18,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục TDP1-Thị Trấn Hương Khê-Hương Khê-Hà Tĩnh"
                                           },
                                           {
                                               "TongCOD":  44.36,
                                               "CODTuNop":  44.36,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thị Trấn Kim Nhan-Anh Sơn-Nghệ An"
                                           },
                                           {
                                               "TongCOD":  1.16,
                                               "CODTuNop":  1.16,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thị Trấn Quán Hành-Nghị Lộc-Nghệ An"
                                           },
                                           {
                                               "TongCOD":  46.14,
                                               "CODTuNop":  46.14,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thôn Khoa Trường-Tĩnh Gia-Thanh Hóa"
                                           },
                                           {
                                               "TongCOD":  46.46,
                                               "CODTuNop":  46.46,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thôn Nam Thạch A-Yên Trung-Yên Định-Thanh Hóa"
                                           },
                                           {
                                               "TongCOD":  43.09,
                                               "CODTuNop":  43.09,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thôn Quyết Thắng-Xã Thiết Ống-Thanh Hóa"
                                           },
                                           {
                                               "TongCOD":  54.11,
                                               "CODTuNop":  54.11,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu cục Tiểu Khu Yên Hạnh 1-Nga Sơn-Thanh Hoá"
                                           },
                                           {
                                               "TongCOD":  42.84,
                                               "CODTuNop":  42.84,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Xã Tân Phú-Tân Kỳ-Nghệ An"
                                           },
                                           {
                                               "TongCOD":  21.83,
                                               "CODTuNop":  21.83,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Xóm 15 Xã Nghi Lâm-Nghi Lộc-Nghệ An"
                                           },
                                           {
                                               "TongCOD":  7.62,
                                               "CODTuNop":  7.62,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Xóm 6 Diễn Thịnh-Diễn Châu-Nghệ An"
                                           },
                                           {
                                               "TongCOD":  74.87,
                                               "CODTuNop":  74.87,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Xóm Quỳ Lăng-Xã Giai Lạc-Nghệ An"
                                           },
                                           {
                                               "TongCOD":  75.14,
                                               "CODTuNop":  75.14,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Xóm Xuân Đình-Diễn Hồng-Diễn Châu-Nghệ An"
                                           },
                                           {
                                               "TongCOD":  8.61,
                                               "CODTuNop":  8.61,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Điểm xử lý hàng 415 Hà Huy Tập-Xã Cẩm Xuyên-Hà Tĩnh"
                                           },
                                           {
                                               "TongCOD":  19.22,
                                               "CODTuNop":  19.22,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Điểm xử lý hàng Huyện Vũ Quang-Hà Tĩnh"
                                           },
                                           {
                                               "TongCOD":  30.85,
                                               "CODTuNop":  30.85,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Điểm xử lý hàng Thị Trấn Mường Xén-Kỳ Sơn-Nghệ An"
                                           }
                                       ]
                    },
                    {
                        "TyLeTuNop":  100,
                        "TenVung":  "DBB",
                        "TongCOD":  1868,
                        "Trend":  0,
                        "CODTuNop":  1868,
                        "BuuCucList":  [
                                           {
                                               "TongCOD":  208.36,
                                               "CODTuNop":  208.36,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 243 Phan Đình Phùng-Phan Đình Phùng-Thái Nguyên"
                                           },
                                           {
                                               "TongCOD":  66.13,
                                               "CODTuNop":  66.13,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 25 Trần Hưng Đạo-Xã Yên Minh-Hà Giang"
                                           },
                                           {
                                               "TongCOD":  90.87,
                                               "CODTuNop":  90.87,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 65 Nguyễn Văn Thoát-TP.Bắc Kạn"
                                           },
                                           {
                                               "TongCOD":  5.49,
                                               "CODTuNop":  5.49,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 86 Nguyễn Thị Minh Khai-Xã Mèo Vạc-Hà Giang"
                                           },
                                           {
                                               "TongCOD":  22.63,
                                               "CODTuNop":  22.63,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Cầu Vĩnh Tuy-Xã Vĩnh Tuy-Hà Giang"
                                           },
                                           {
                                               "TongCOD":  56.76,
                                               "CODTuNop":  56.76,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Chợ TT Nam Sông Công-Xã Đại Phúc-Thái Nguyên"
                                           },
                                           {
                                               "TongCOD":  51.04,
                                               "CODTuNop":  51.04,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Đường 20/8-Phương Độ-TP.Hà Giang-Hà Giang"
                                           },
                                           {
                                               "TongCOD":  32.64,
                                               "CODTuNop":  32.64,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Đường 261-Bắc Sơn-Phổ Yên-Thái Nguyên"
                                           },
                                           {
                                               "TongCOD":  22.75,
                                               "CODTuNop":  22.75,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Đường Trần Phú-Xã Nà Hang-Tuyên Quang"
                                           },
                                           {
                                               "TongCOD":  57.36,
                                               "CODTuNop":  57.36,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hàm Yên-Tuyên Quang"
                                           },
                                           {
                                               "TongCOD":  5.69,
                                               "CODTuNop":  5.69,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hoàng Quốc Việt-Tân Hương-Phổ Yên-Thái Nguyên"
                                           },
                                           {
                                               "TongCOD":  21.07,
                                               "CODTuNop":  21.07,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục K31 Thôn Nà Sài-Xã Minh Ngọc-Hà Giang"
                                           },
                                           {
                                               "TongCOD":  60.16,
                                               "CODTuNop":  60.16,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Lê Quý Đôn-Đồng Hỷ-Thái Nguyên"
                                           },
                                           {
                                               "TongCOD":  116.32,
                                               "CODTuNop":  116.32,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Phùng Hưng-Hà Giang 2-Hà Giang"
                                           },
                                           {
                                               "TongCOD":  14.71,
                                               "CODTuNop":  14.71,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Quốc Lộ 2-Xã Yên Sơn-Tuyên Quang"
                                           },
                                           {
                                               "TongCOD":  11.19,
                                               "CODTuNop":  11.19,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Quốc Lộ 3B-Xã Na Rì-Bắc Kạn"
                                           },
                                           {
                                               "TongCOD":  182.9,
                                               "CODTuNop":  182.9,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Số 10 Phan Huy Ích-Tân Hà-Tuyên Quang"
                                           },
                                           {
                                               "TongCOD":  57.01,
                                               "CODTuNop":  57.01,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Số 8 Phan Huy Ích-Tân Hà-Tuyên Quang"
                                           },
                                           {
                                               "TongCOD":  64.04,
                                               "CODTuNop":  64.04,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Số Nhà 30-Mỏ Chè-Sông Công-Thái Nguyên"
                                           },
                                           {
                                               "TongCOD":  17.32,
                                               "CODTuNop":  17.32,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Số Nhà 882 Tổ 6-Việt Quang-Bắc Quang-Hà Giang"
                                           },
                                           {
                                               "TongCOD":  111.04,
                                               "CODTuNop":  111.04,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục TDP Lê Hồng Phong-Phú Lương-Thái Nguyên"
                                           },
                                           {
                                               "TongCOD":  87.7,
                                               "CODTuNop":  87.7,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục TDP Trung Kiên-Định Hóa-Thái Nguyên"
                                           },
                                           {
                                               "TongCOD":  67.1,
                                               "CODTuNop":  67.1,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thị trấn Vị Xuyên-Hà Giang"
                                           },
                                           {
                                               "TongCOD":  11.58,
                                               "CODTuNop":  11.58,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thị Trấn Vinh Quang-Hoàng Su Phì-Hà Giang"
                                           },
                                           {
                                               "TongCOD":  26.08,
                                               "CODTuNop":  26.08,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thôn Tân Hòa-Sơn Dương-Tuyên Quang"
                                           },
                                           {
                                               "TongCOD":  43.92,
                                               "CODTuNop":  43.92,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Tiểu Khu 2-Nà Phặc-Ngân Sơn-Bắc Kạn"
                                           },
                                           {
                                               "TongCOD":  7.82,
                                               "CODTuNop":  7.82,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Tổ 1 TT Hương Sơn-Xã Phú Bình-Thái Nguyên"
                                           },
                                           {
                                               "TongCOD":  55.93,
                                               "CODTuNop":  55.93,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Tổ 11 Đồng Tâm-Chợ Mới-Bắc Kạn"
                                           },
                                           {
                                               "TongCOD":  24.81,
                                               "CODTuNop":  24.81,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Tổ 2 Quốc Lộ 4C-Xã Quản Bạ-Hà Giang"
                                           },
                                           {
                                               "TongCOD":  7.36,
                                               "CODTuNop":  7.36,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Tổ 4-Thị trấn Nguyên Bình-Cao Bằng"
                                           },
                                           {
                                               "TongCOD":  25.84,
                                               "CODTuNop":  25.84,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Tổ 5 TT Bằng Lũng-Xã Chợ Đồn-Bắc Kạn"
                                           },
                                           {
                                               "TongCOD":  86.84,
                                               "CODTuNop":  86.84,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Tôn Đức Thắng-Phổ Yên-Thái Nguyên"
                                           },
                                           {
                                               "TongCOD":  19.43,
                                               "CODTuNop":  19.43,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Xóm 1 Đường 264-Xã Phú Thịnh-Thái Nguyên"
                                           },
                                           {
                                               "TongCOD":  5.08,
                                               "CODTuNop":  5.08,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Xóm 6-Nước Hai-Hòa An-Cao Bằng"
                                           },
                                           {
                                               "TongCOD":  29.1,
                                               "CODTuNop":  29.1,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Xóm Đoàn Kết-Nam Hòa-Đồng Hỷ-Tỉnh Thái Nguyên"
                                           },
                                           {
                                               "TongCOD":  11.08,
                                               "CODTuNop":  11.08,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Điểm xử lý hàng Huyện Xín Mần-Hà Giang"
                                           },
                                           {
                                               "TongCOD":  68.95,
                                               "CODTuNop":  68.95,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Điểm Xử Lý Hàng TT Nà Phặc-Xã Nà Phặc-Bắc Kạn"
                                           },
                                           {
                                               "TongCOD":  13.91,
                                               "CODTuNop":  13.91,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Điểm Xử Lý Hàng Xóm 6 TT Nước Hai-Xã Hòa An-Cao Bằng"
                                           }
                                       ]
                    },
                    {
                        "TyLeTuNop":  100,
                        "TenVung":  "DNB",
                        "TongCOD":  6554.63,
                        "Trend":  0,
                        "CODTuNop":  6554.63,
                        "BuuCucList":  [
                                           {
                                               "TongCOD":  14.07,
                                               "CODTuNop":  14.07,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 105C Đồng Khởi-Biên Hòa-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  25.6,
                                               "CODTuNop":  25.6,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 107 Nguyễn Chí Thanh-Tân Biên-Tây Ninh"
                                           },
                                           {
                                               "TongCOD":  76.64,
                                               "CODTuNop":  76.64,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 1070 Quốc Lộ 51-TP. Biên Hòa-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  131.38,
                                               "CODTuNop":  131.38,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 110 Nguyễn Huệ-Trảng Bom-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  103.27,
                                               "CODTuNop":  103.27,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 129 Khổng Tử-Long Khánh-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  20.31,
                                               "CODTuNop":  20.31,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 131 Trần Phú-Ninh Trung-Ninh Sơn-Tây Ninh"
                                           },
                                           {
                                               "TongCOD":  12.22,
                                               "CODTuNop":  12.22,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 1402 Hoàng Lê Kha-Châu Thành-Tây Ninh"
                                           },
                                           {
                                               "TongCOD":  190.43,
                                               "CODTuNop":  190.43,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 145/36/4 Điểu Xiển-Tân Biên-TP.Biên Hòa-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  27.67,
                                               "CODTuNop":  27.67,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 146 Nguyễn Văn Cừ-Xã Long Thành-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  30.07,
                                               "CODTuNop":  30.07,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 15 Đồng Khởi-Vĩnh Cửu-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  71.07,
                                               "CODTuNop":  71.07,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 156 Quốc Lộ 22-Gò Dầu-Tây Ninh"
                                           },
                                           {
                                               "TongCOD":  59.17,
                                               "CODTuNop":  59.17,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 219 Quốc Lộ 62-Kiến Tường-Long An"
                                           },
                                           {
                                               "TongCOD":  118.97,
                                               "CODTuNop":  118.97,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 24 Khu Phố Thanh Bình-Bù Đốp-Bình Phước"
                                           },
                                           {
                                               "TongCOD":  41.56,
                                               "CODTuNop":  41.56,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 25 Khedol -Xã Suối Đá-Dương Minh Châu-Tây Ninh"
                                           },
                                           {
                                               "TongCOD":  75.01,
                                               "CODTuNop":  75.01,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 270 Phan Văn Mãng- Bến Lức-Long An"
                                           },
                                           {
                                               "TongCOD":  95.82,
                                               "CODTuNop":  95.82,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 295 Nguyễn Văn Linh-Long Hoa-Tây Ninh"
                                           },
                                           {
                                               "TongCOD":  70.27,
                                               "CODTuNop":  70.27,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 3/3A Huỳnh Văn Nghệ-Bửu Long-Biên Hòa-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  62.65,
                                               "CODTuNop":  62.65,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 30 Nguyễn Trãi-Phường 3-Tây Ninh"
                                           },
                                           {
                                               "TongCOD":  20.12,
                                               "CODTuNop":  20.12,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 305 QL14-Bù Đăng-Bình Phước"
                                           },
                                           {
                                               "TongCOD":  112.73,
                                               "CODTuNop":  112.73,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 322 Trần Phú-Xuân Lộc-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  105.07,
                                               "CODTuNop":  105.07,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 393 Tôn Đức Thắng-Hòa Thành-Tây Ninh"
                                           },
                                           {
                                               "TongCOD":  438.42,
                                               "CODTuNop":  438.42,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 45/2b Quốc Lộ 62-Tân An-Long An 01"
                                           },
                                           {
                                               "TongCOD":  48.07,
                                               "CODTuNop":  48.07,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 478 Đường Tua Hai-Bình Minh-Tây Ninh"
                                           },
                                           {
                                               "TongCOD":  0.26,
                                               "CODTuNop":  0.26,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 4B Khu B-Xã Hậu Nghĩa-Long An"
                                           },
                                           {
                                               "TongCOD":  110.08,
                                               "CODTuNop":  110.08,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 55/3B Trần Quốc Toản-Phường Trấn Biên-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  62.46,
                                               "CODTuNop":  62.46,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 56 Dương Văn Dương-Xã Tân Thạnh-Long An"
                                           },
                                           {
                                               "TongCOD":  34.15,
                                               "CODTuNop":  34.15,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 58 Trần Văn Trà-Tân Châu-Tây Ninh"
                                           },
                                           {
                                               "TongCOD":  124.69,
                                               "CODTuNop":  124.69,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 608 Quốc Lộ 20-Định Quán-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  10.72,
                                               "CODTuNop":  10.72,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 68 Quốc Lộ 1A-Trảng Bom-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  1.88,
                                               "CODTuNop":  1.88,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 69 Bàu Cạn-Xã Long Phước-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  88.11,
                                               "CODTuNop":  88.11,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 76/540 KP8-Hố Nai-TP Biên Hòa-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  111.18,
                                               "CODTuNop":  111.18,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 893 Quốc Lộ 14-Đồng Xoài-Bình Phước"
                                           },
                                           {
                                               "TongCOD":  31.47,
                                               "CODTuNop":  31.47,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Agtex Long Bình-Biên Hoà-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  57.26,
                                               "CODTuNop":  57.26,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu cục An Lộc-Bình Long-Bình Phước"
                                           },
                                           {
                                               "TongCOD":  23.24,
                                               "CODTuNop":  23.24,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục An Lợi-Trảng Bàng-Tây Ninh"
                                           },
                                           {
                                               "TongCOD":  1.76,
                                               "CODTuNop":  1.76,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Ấp 1 Đường Tân Châu-Tân Châu-Tây Ninh"
                                           },
                                           {
                                               "TongCOD":  206.57,
                                               "CODTuNop":  206.57,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Ấp 4-Long Hoà-Cần Đước-Long An"
                                           },
                                           {
                                               "TongCOD":  139.74,
                                               "CODTuNop":  139.74,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Ấp 8-Xuân Tây-Cẩm Mỹ-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  106.61,
                                               "CODTuNop":  106.61,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Ấp Bến Đình-Xã Phú Đông-Nhơn Trạch-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  5.05,
                                               "CODTuNop":  5.05,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Ấp Cầu Xây-Xã Thủ Thừa-Long An"
                                           },
                                           {
                                               "TongCOD":  157.54,
                                               "CODTuNop":  157.54,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Ấp Đông Kim-Xã Gia Kiệm-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  8.98,
                                               "CODTuNop":  8.98,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Ấp Lủy-Phước Lại-Cần Giuộc-Long An"
                                           },
                                           {
                                               "TongCOD":  27.88,
                                               "CODTuNop":  27.88,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Ấp Ngoài-Xã Phước Lý-Long An"
                                           },
                                           {
                                               "TongCOD":  91.77,
                                               "CODTuNop":  91.77,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Ấp Rừng Sến-Xã Đức Lập-Long An"
                                           },
                                           {
                                               "TongCOD":  7.41,
                                               "CODTuNop":  7.41,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Ấp Tân Hòa-Tân Biên-Tây Ninh"
                                           },
                                           {
                                               "TongCOD":  82.56,
                                               "CODTuNop":  82.56,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Bù Nho-Phú Riềng-Bình Phước"
                                           },
                                           {
                                               "TongCOD":  123.57,
                                               "CODTuNop":  123.57,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Bùi Trọng Nghĩa-Trảng Dài-Biên Hòa-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  144.13,
                                               "CODTuNop":  144.13,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Cần Đước-Long An"
                                           },
                                           {
                                               "TongCOD":  1.33,
                                               "CODTuNop":  1.33,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Cầu Khởi-Dương Minh Châu-Tây Ninh"
                                           },
                                           {
                                               "TongCOD":  39.24,
                                               "CODTuNop":  39.24,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Chánh Hội-Tân Mỹ-Đức Hòa-Long An"
                                           },
                                           {
                                               "TongCOD":  43.83,
                                               "CODTuNop":  43.83,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Chơn Thành-Bình Phước"
                                           },
                                           {
                                               "TongCOD":  79.61,
                                               "CODTuNop":  79.61,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục DT 786 Khu Phố 4-Bến Cầu-Tây Ninh"
                                           },
                                           {
                                               "TongCOD":  182.43,
                                               "CODTuNop":  182.43,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Đường 25A3-Xã Đức Hòa Hạ-Long An"
                                           },
                                           {
                                               "TongCOD":  4.82,
                                               "CODTuNop":  4.82,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Đường Ba Mươi Tháng Tư-Trảng Bàng-Tây Ninh"
                                           },
                                           {
                                               "TongCOD":  124.32,
                                               "CODTuNop":  124.32,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Đường Hùng Vương-Xã Phước An-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  163.36,
                                               "CODTuNop":  163.36,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Đường Tỉnh 827-Tầm Vu-Châu Thành-Long An"
                                           },
                                           {
                                               "TongCOD":  75.04,
                                               "CODTuNop":  75.04,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Đường Tỉnh 830-Xã Lương Bình-Long An"
                                           },
                                           {
                                               "TongCOD":  97.66,
                                               "CODTuNop":  97.66,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Đường Võ Nguyên Giáp-Biên Hòa-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  62.23,
                                               "CODTuNop":  62.23,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hóa An-Biên Hòa-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  56.67,
                                               "CODTuNop":  56.67,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Khu 6 ấp Bàu Cá-Trung Hòa-Trảng Bom-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  136.73,
                                               "CODTuNop":  136.73,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Khu Phố 2-Thạnh Hoá-Long An"
                                           },
                                           {
                                               "TongCOD":  39.77,
                                               "CODTuNop":  39.77,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục KP1 Quốc Lộ 13-Hớn Quản-Bình Phước"
                                           },
                                           {
                                               "TongCOD":  168.36,
                                               "CODTuNop":  168.36,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục KP2 Tân Hiệp-Định Quán Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  64.79,
                                               "CODTuNop":  64.79,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Lộc Thái-Lộc Ninh-Bình Phước"
                                           },
                                           {
                                               "TongCOD":  51.13,
                                               "CODTuNop":  51.13,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Long Phước-Long Thành-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  119.97,
                                               "CODTuNop":  119.97,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Lý Thường Kiệt-Xã An Phước-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  31.89,
                                               "CODTuNop":  31.89,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Minh Hưng-TX.Chơn Thành-Bình Phước"
                                           },
                                           {
                                               "TongCOD":  61.93,
                                               "CODTuNop":  61.93,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Nguyễn Thị Minh Khai-Thị trấn Tân Phú-Huyện Đồng Phú-Bình Phước"
                                           },
                                           {
                                               "TongCOD":  146.16,
                                               "CODTuNop":  146.16,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Phạm Hùng - Phước Long - Bình Phước"
                                           },
                                           {
                                               "TongCOD":  32.25,
                                               "CODTuNop":  32.25,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Phú Lập-Tân Phú-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  134.3,
                                               "CODTuNop":  134.3,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Phước Thiền-Nhơn Trạch-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  110.38,
                                               "CODTuNop":  110.38,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Quốc Lộ 1A-Bắc Sơn-Trảng Bom-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  93.39,
                                               "CODTuNop":  93.39,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Tam Phước-Biên Hòa-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  68.62,
                                               "CODTuNop":  68.62,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Tân Đồng-Đồng Xoài-Bình Phước"
                                           },
                                           {
                                               "TongCOD":  73.55,
                                               "CODTuNop":  73.55,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Tân Phú-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  200.63,
                                               "CODTuNop":  200.63,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Tân Tiến-Biên Hòa-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  26.48,
                                               "CODTuNop":  26.48,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thống Nhất-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  221.6,
                                               "CODTuNop":  221.6,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Tổ 2 Ấp 5 Xuân Tâm-Xuân Lộc-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  16.48,
                                               "CODTuNop":  16.48,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Tổ 4 KP7 767- Vĩnh Cửu-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  11.61,
                                               "CODTuNop":  11.61,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Xã Đức Liễu-Bù Đăng-Bình Phước"
                                           },
                                           {
                                               "TongCOD":  104.45,
                                               "CODTuNop":  104.45,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Xã Long Giao-Cẩm Mỹ-Đồng Nai"
                                           }
                                       ]
                    },
                    {
                        "TyLeTuNop":  100,
                        "TenVung":  "DSH",
                        "TongCOD":  2338.28,
                        "Trend":  0,
                        "CODTuNop":  2338.28,
                        "BuuCucList":  [
                                           {
                                               "TongCOD":  135.31,
                                               "CODTuNop":  135.31,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 11 Chu Mạnh Trinh-TP.Hưng Yên"
                                           },
                                           {
                                               "TongCOD":  127.51,
                                               "CODTuNop":  127.51,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 176 Nguyễn Đức Cảnh-Hoàng Diệu-Thái Bình"
                                           },
                                           {
                                               "TongCOD":  14.88,
                                               "CODTuNop":  14.88,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 90 Nguyễn Lương Bằng-Kim Động-Hưng Yên"
                                           },
                                           {
                                               "TongCOD":  168.7,
                                               "CODTuNop":  168.7,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Cầu Khé DT377-Xã Khoái Châu-Hưng Yên"
                                           },
                                           {
                                               "TongCOD":  97.06,
                                               "CODTuNop":  97.06,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Đông Hưng-Thái Bình"
                                           },
                                           {
                                               "TongCOD":  165.56,
                                               "CODTuNop":  165.56,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục DT 376-Ngọc Long-Yên Mỹ-Hưng Yên"
                                           },
                                           {
                                               "TongCOD":  100.14,
                                               "CODTuNop":  100.14,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Đường Doãn Khuê-Phường Vũ Phúc-Thái Bình"
                                           },
                                           {
                                               "TongCOD":  127.97,
                                               "CODTuNop":  127.97,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Minh Khai-Vũ Thư-Thái Bình"
                                           },
                                           {
                                               "TongCOD":  164.2,
                                               "CODTuNop":  164.2,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Nhân Hòa-Mỹ Hào-Hưng Yên"
                                           },
                                           {
                                               "TongCOD":  167.11,
                                               "CODTuNop":  167.11,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Quốc Lộ 39B-Kiến Xương- Thái Bình"
                                           },
                                           {
                                               "TongCOD":  27.48,
                                               "CODTuNop":  27.48,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Quỳnh Hội-Quỳnh Phụ-Thái Bình"
                                           },
                                           {
                                               "TongCOD":  78.09,
                                               "CODTuNop":  78.09,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thị Trấn Ân Thi-Ân Thi-Hưng Yên"
                                           },
                                           {
                                               "TongCOD":  40.5,
                                               "CODTuNop":  40.5,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thị Trấn Như Quỳnh-Văn Lâm-Hưng Yên"
                                           },
                                           {
                                               "TongCOD":  62.88,
                                               "CODTuNop":  62.88,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thôn Cao Xá-Trần Cao-Phù Cừ-Hưng Yên"
                                           },
                                           {
                                               "TongCOD":  64.52,
                                               "CODTuNop":  64.52,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thôn Đại Từ-Văn Lâm-Hưng Yên"
                                           },
                                           {
                                               "TongCOD":  143.69,
                                               "CODTuNop":  143.69,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thôn Đông-Tiền Hải-Thái Bình"
                                           },
                                           {
                                               "TongCOD":  128.91,
                                               "CODTuNop":  128.91,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thôn Lực Điền-Yên Mỹ-Hưng Yên"
                                           },
                                           {
                                               "TongCOD":  20.7,
                                               "CODTuNop":  20.7,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thôn Vũ Hạ-Xã Phụ Dực-Thái Bình"
                                           },
                                           {
                                               "TongCOD":  189.84,
                                               "CODTuNop":  189.84,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Xã Tân Tiến-Văn Giang-Hưng Yên"
                                           },
                                           {
                                               "TongCOD":  162.6,
                                               "CODTuNop":  162.6,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Xã Thống Nhất-Hưng Hà-Thái Bình"
                                           },
                                           {
                                               "TongCOD":  150.63,
                                               "CODTuNop":  150.63,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Xóm 2 Thụy Sơn-Thái Thụy-Thái Bình"
                                           }
                                       ]
                    },
                    {
                        "TyLeTuNop":  100,
                        "TenVung":  "GXT",
                        "TongCOD":  16184.24,
                        "Trend":  0,
                        "CODTuNop":  16184.24,
                        "BuuCucList":  [
                                           {
                                               "TongCOD":  622.85,
                                               "CODTuNop":  622.85,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - An Dương - Hải Phòng"
                                           },
                                           {
                                               "TongCOD":  124.85,
                                               "CODTuNop":  124.85,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Ân Thi - Hưng Yên"
                                           },
                                           {
                                               "TongCOD":  212.36,
                                               "CODTuNop":  212.36,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Bắc Từ Liêm - Hà Nội"
                                           },
                                           {
                                               "TongCOD":  134.99,
                                               "CODTuNop":  134.99,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Bình Thủy - Cần Thơ"
                                           },
                                           {
                                               "TongCOD":  227.54,
                                               "CODTuNop":  227.54,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Bình Xuyên - Vĩnh Phúc"
                                           },
                                           {
                                               "TongCOD":  79.06,
                                               "CODTuNop":  79.06,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Buôn Hồ - Đắk Lắk"
                                           },
                                           {
                                               "TongCOD":  16.16,
                                               "CODTuNop":  16.16,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Buôn Ma Thuột - Đắk Lắk"
                                           },
                                           {
                                               "TongCOD":  326.15,
                                               "CODTuNop":  326.15,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Cái Răng - Cần Thơ"
                                           },
                                           {
                                               "TongCOD":  53.1,
                                               "CODTuNop":  53.1,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Cam Ranh - Khánh Hoà"
                                           },
                                           {
                                               "TongCOD":  259.19,
                                               "CODTuNop":  259.19,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Dĩ An - Bình Dương"
                                           },
                                           {
                                               "TongCOD":  125.29,
                                               "CODTuNop":  125.29,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Đông Anh - Hà Nội"
                                           },
                                           {
                                               "TongCOD":  81.17,
                                               "CODTuNop":  81.17,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Đông Hà - Quảng Trị"
                                           },
                                           {
                                               "TongCOD":  125.13,
                                               "CODTuNop":  125.13,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Đồng Hới - Quảng Bình"
                                           },
                                           {
                                               "TongCOD":  256.98,
                                               "CODTuNop":  256.98,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Đông Hưng - Thái Bình"
                                           },
                                           {
                                               "TongCOD":  247.25,
                                               "CODTuNop":  247.25,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Đông Thọ - Thanh Hoá"
                                           },
                                           {
                                               "TongCOD":  185.29,
                                               "CODTuNop":  185.29,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Đức Hòa - Long An"
                                           },
                                           {
                                               "TongCOD":  148.69,
                                               "CODTuNop":  148.69,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Đức Trọng - Lâm Đồng"
                                           },
                                           {
                                               "TongCOD":  55.83,
                                               "CODTuNop":  55.83,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Gò Công - Tiền Giang"
                                           },
                                           {
                                               "TongCOD":  190.45,
                                               "CODTuNop":  190.45,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Hồ Chí Minh"
                                           },
                                           {
                                               "TongCOD":  308.61,
                                               "CODTuNop":  308.61,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Hoà Xuân - Đà Nẵng"
                                           },
                                           {
                                               "TongCOD":  177.1,
                                               "CODTuNop":  177.1,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Hoài Đức - Hà Nội"
                                           },
                                           {
                                               "TongCOD":  74.71,
                                               "CODTuNop":  74.71,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Hoài Nhơn - Bình Định"
                                           },
                                           {
                                               "TongCOD":  150.01,
                                               "CODTuNop":  150.01,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Hội An - Quảng Nam"
                                           },
                                           {
                                               "TongCOD":  203.7,
                                               "CODTuNop":  203.7,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Hương Thủy - Huế"
                                           },
                                           {
                                               "TongCOD":  82.83,
                                               "CODTuNop":  82.83,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - La Gi - Bình Thuận"
                                           },
                                           {
                                               "TongCOD":  374.49,
                                               "CODTuNop":  374.49,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Liên Chiểu - Đà Nẵng"
                                           },
                                           {
                                               "TongCOD":  277.55,
                                               "CODTuNop":  277.55,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Long Biên - Hà Nội"
                                           },
                                           {
                                               "TongCOD":  248.81,
                                               "CODTuNop":  248.81,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Long Thành - Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  278.23,
                                               "CODTuNop":  278.23,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Mỹ Lộc - Nam Định"
                                           },
                                           {
                                               "TongCOD":  335.69,
                                               "CODTuNop":  335.69,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Nhà Bè - HCM"
                                           },
                                           {
                                               "TongCOD":  396.32,
                                               "CODTuNop":  396.32,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Nha Trang - Khánh Hoà"
                                           },
                                           {
                                               "TongCOD":  78.17,
                                               "CODTuNop":  78.17,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Nho Quan - Ninh Bình"
                                           },
                                           {
                                               "TongCOD":  95.27,
                                               "CODTuNop":  95.27,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Ô Môn - Cần Thơ"
                                           },
                                           {
                                               "TongCOD":  129.05,
                                               "CODTuNop":  129.05,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Phan Rang - Ninh Thuận"
                                           },
                                           {
                                               "TongCOD":  98.32,
                                               "CODTuNop":  98.32,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Phan Thiết - Bình Thuận"
                                           },
                                           {
                                               "TongCOD":  201.76,
                                               "CODTuNop":  201.76,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Pleiku - Gia Lai"
                                           },
                                           {
                                               "TongCOD":  67.24,
                                               "CODTuNop":  67.24,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Quảng Ngãi - Quảng Ngãi"
                                           },
                                           {
                                               "TongCOD":  60.6,
                                               "CODTuNop":  60.6,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Quy Nhơn - Bình Định"
                                           },
                                           {
                                               "TongCOD":  129.7,
                                               "CODTuNop":  129.7,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Tam Kỳ - Quảng Nam"
                                           },
                                           {
                                               "TongCOD":  64.21,
                                               "CODTuNop":  64.21,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Tam Nông - Đồng Tháp"
                                           },
                                           {
                                               "TongCOD":  131.43,
                                               "CODTuNop":  131.43,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Tân Bình - Hải Dương"
                                           },
                                           {
                                               "TongCOD":  530.2,
                                               "CODTuNop":  530.2,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Tân Bình - HCM"
                                           },
                                           {
                                               "TongCOD":  781.98,
                                               "CODTuNop":  781.98,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Tân Tạo - HCM"
                                           },
                                           {
                                               "TongCOD":  70.19,
                                               "CODTuNop":  70.19,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Thạch Linh - Hà Tĩnh"
                                           },
                                           {
                                               "TongCOD":  57.12,
                                               "CODTuNop":  57.12,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Thắng Lợi - Kon Tum"
                                           },
                                           {
                                               "TongCOD":  197.74,
                                               "CODTuNop":  197.74,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Thanh Oai - Hà Nội"
                                           },
                                           {
                                               "TongCOD":  82.85,
                                               "CODTuNop":  82.85,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Thanh Trì - Hà Nội"
                                           },
                                           {
                                               "TongCOD":  173.06,
                                               "CODTuNop":  173.06,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Thủ Dầu Một - Bình Dương"
                                           },
                                           {
                                               "TongCOD":  516.23,
                                               "CODTuNop":  516.23,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Thủ Đức - HCM"
                                           },
                                           {
                                               "TongCOD":  318.06,
                                               "CODTuNop":  318.06,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Thuận An - Bình Dương"
                                           },
                                           {
                                               "TongCOD":  272.07,
                                               "CODTuNop":  272.07,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Tiên Du - Bắc Ninh"
                                           },
                                           {
                                               "TongCOD":  188.01,
                                               "CODTuNop":  188.01,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Bà Rịa - BRVT"
                                           },
                                           {
                                               "TongCOD":  25.43,
                                               "CODTuNop":  25.43,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Bạc Liêu - Bạc Liêu"
                                           },
                                           {
                                               "TongCOD":  39.15,
                                               "CODTuNop":  39.15,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Bảo Lộc - Lâm Đồng"
                                           },
                                           {
                                               "TongCOD":  229.88,
                                               "CODTuNop":  229.88,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Bến Tre - Bến Tre"
                                           },
                                           {
                                               "TongCOD":  304.71,
                                               "CODTuNop":  304.71,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Biên Hoà - Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  170.9,
                                               "CODTuNop":  170.9,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Cà Mau - Cà Mau"
                                           },
                                           {
                                               "TongCOD":  297.47,
                                               "CODTuNop":  297.47,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Cao Lãnh - Đồng Tháp"
                                           },
                                           {
                                               "TongCOD":  266.11,
                                               "CODTuNop":  266.11,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Châu Đốc - An Giang"
                                           },
                                           {
                                               "TongCOD":  243.57,
                                               "CODTuNop":  243.57,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Đà Lạt - Lâm Đồng"
                                           },
                                           {
                                               "TongCOD":  115.7,
                                               "CODTuNop":  115.7,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Đồng Xoài - Bình Phước"
                                           },
                                           {
                                               "TongCOD":  356.57,
                                               "CODTuNop":  356.57,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Hạ Long - Quảng Ninh"
                                           },
                                           {
                                               "TongCOD":  46.95,
                                               "CODTuNop":  46.95,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Hà Tiên - Kiên Giang"
                                           },
                                           {
                                               "TongCOD":  43.52,
                                               "CODTuNop":  43.52,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Lạng Sơn - Lạng Sơn"
                                           },
                                           {
                                               "TongCOD":  120.21,
                                               "CODTuNop":  120.21,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Lào Cai - Lào Cai"
                                           },
                                           {
                                               "TongCOD":  312.35,
                                               "CODTuNop":  312.35,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Long Xuyên - An Giang"
                                           },
                                           {
                                               "TongCOD":  165.26,
                                               "CODTuNop":  165.26,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Mỹ Tho - Tiền Giang"
                                           },
                                           {
                                               "TongCOD":  205,
                                               "CODTuNop":  205,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Rạch Giá - Kiên Giang"
                                           },
                                           {
                                               "TongCOD":  98.11,
                                               "CODTuNop":  98.11,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Sóc Trăng - Sóc Trăng"
                                           },
                                           {
                                               "TongCOD":  58.94,
                                               "CODTuNop":  58.94,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Sơn La - Sơn La"
                                           },
                                           {
                                               "TongCOD":  145.02,
                                               "CODTuNop":  145.02,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Tân An - Long An"
                                           },
                                           {
                                               "TongCOD":  169.33,
                                               "CODTuNop":  169.33,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Tây Ninh - Tây Ninh"
                                           },
                                           {
                                               "TongCOD":  315.55,
                                               "CODTuNop":  315.55,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Thái Nguyên - Thái Nguyên"
                                           },
                                           {
                                               "TongCOD":  147.14,
                                               "CODTuNop":  147.14,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Trà Vinh - Trà Vinh"
                                           },
                                           {
                                               "TongCOD":  86.24,
                                               "CODTuNop":  86.24,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Tuyên Quang - Tuyên Quang"
                                           },
                                           {
                                               "TongCOD":  112.31,
                                               "CODTuNop":  112.31,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Vị Thanh - Hậu Giang"
                                           },
                                           {
                                               "TongCOD":  191.78,
                                               "CODTuNop":  191.78,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Vĩnh Long - Vĩnh Long"
                                           },
                                           {
                                               "TongCOD":  136.77,
                                               "CODTuNop":  136.77,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Vũng Tàu - BRVT"
                                           },
                                           {
                                               "TongCOD":  68.69,
                                               "CODTuNop":  68.69,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Yên Bái - Yên Bái"
                                           },
                                           {
                                               "TongCOD":  251.34,
                                               "CODTuNop":  251.34,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Trảng Bom - Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  123.56,
                                               "CODTuNop":  123.56,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Tuy Hoà - Phú Yên"
                                           },
                                           {
                                               "TongCOD":  63.8,
                                               "CODTuNop":  63.8,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Tuy Phong - Bình Thuận"
                                           },
                                           {
                                               "TongCOD":  301.62,
                                               "CODTuNop":  301.62,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TX Cai Lậy - Tiền Giang"
                                           },
                                           {
                                               "TongCOD":  132.69,
                                               "CODTuNop":  132.69,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Việt Trì - Phú Thọ"
                                           },
                                           {
                                               "TongCOD":  172.05,
                                               "CODTuNop":  172.05,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Việt Yên - Bắc Giang"
                                           },
                                           {
                                               "TongCOD":  42.89,
                                               "CODTuNop":  42.89,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Vinh - Nghệ An"
                                           }
                                       ]
                    },
                    {
                        "TyLeTuNop":  100,
                        "TenVung":  "HCM",
                        "TongCOD":  22517.96,
                        "Trend":  0,
                        "CODTuNop":  22517.96,
                        "BuuCucList":  [
                                           {
                                               "TongCOD":  126.56,
                                               "CODTuNop":  126.56,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 01 Đường Số 6-Long Bình-Q.9-HCM"
                                           },
                                           {
                                               "TongCOD":  160.7,
                                               "CODTuNop":  160.7,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 02 Đường Số 28 - Củ Chi - HCM"
                                           },
                                           {
                                               "TongCOD":  144.2,
                                               "CODTuNop":  144.2,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 04 Đường 13 tháng 3-Dầu Tiếng-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  101.51,
                                               "CODTuNop":  101.51,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 04 Nhất Chi Mai-Q.Tân Bình-HCM"
                                           },
                                           {
                                               "TongCOD":  192.21,
                                               "CODTuNop":  192.21,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 10 Đường số 8-Linh Xuân-Q.Thủ Đức-HCM"
                                           },
                                           {
                                               "TongCOD":  98.16,
                                               "CODTuNop":  98.16,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 104 Lương Ngọc Quyến-Q.Bình Thạnh-HCM"
                                           },
                                           {
                                               "TongCOD":  129.87,
                                               "CODTuNop":  129.87,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 106/3A Tây Lân-Hóc Môn-HCM"
                                           },
                                           {
                                               "TongCOD":  110.1,
                                               "CODTuNop":  110.1,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 107 Đường 3D-Bình Hưng Hòa B-Bình Tân-HCM"
                                           },
                                           {
                                               "TongCOD":  143.74,
                                               "CODTuNop":  143.74,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 1130 Đông Hưng Thuận 05-Q.12-HCM"
                                           },
                                           {
                                               "TongCOD":  152.75,
                                               "CODTuNop":  152.75,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 1135 Phan Văn Trị-Q.Gò Vấp-HCM"
                                           },
                                           {
                                               "TongCOD":  168.85,
                                               "CODTuNop":  168.85,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 116A Bà Hom-Phường 13-Quận 6-HCM"
                                           },
                                           {
                                               "TongCOD":  82.57,
                                               "CODTuNop":  82.57,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 117 Hùng Vương-Q.5-HCM"
                                           },
                                           {
                                               "TongCOD":  135.36,
                                               "CODTuNop":  135.36,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 119 Đường Số 1-Phường 11-Gò Vấp-HCM"
                                           },
                                           {
                                               "TongCOD":  255.84,
                                               "CODTuNop":  255.84,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 119 Nguyễn Văn Luông-Q.6-HCM"
                                           },
                                           {
                                               "TongCOD":  61.09,
                                               "CODTuNop":  61.09,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 125A Lê Lợi-Hóc Môn-HCM"
                                           },
                                           {
                                               "TongCOD":  145.56,
                                               "CODTuNop":  145.56,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 13 Đường Số 8-Hiệp Bình Phước-Q.Thủ Đức-HCM"
                                           },
                                           {
                                               "TongCOD":  81.43,
                                               "CODTuNop":  81.43,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 136 Đường Huyện 704-Minh Tân-Dầu Tiếng-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  112.31,
                                               "CODTuNop":  112.31,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 1417 Đường N4-D2 KDC Thanh Duy-TP.Tân Uyên-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  186.17,
                                               "CODTuNop":  186.17,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 1440/1 Tỉnh Lộ 10-Tân Tạo-Bình Tân-HCM"
                                           },
                                           {
                                               "TongCOD":  135.99,
                                               "CODTuNop":  135.99,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 1459 Đường 30/4-Vũng Tàu-Bà Rịa Vũng Tàu"
                                           },
                                           {
                                               "TongCOD":  81.49,
                                               "CODTuNop":  81.49,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 147E Đội Cung-Q.11-HCM"
                                           },
                                           {
                                               "TongCOD":  57.07,
                                               "CODTuNop":  57.07,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 150 Đường DT743A-TP.Dĩ An-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  117.85,
                                               "CODTuNop":  117.85,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 155 Quốc lộ 1A-Bình Chiểu-Thủ Đức-HCM"
                                           },
                                           {
                                               "TongCOD":  161.92,
                                               "CODTuNop":  161.92,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 16 Trần Nhật Duật-Q.1-HCM"
                                           },
                                           {
                                               "TongCOD":  111.58,
                                               "CODTuNop":  111.58,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 165A Nguyễn Văn Tạo-Nhà Bè-HCM"
                                           },
                                           {
                                               "TongCOD":  192.21,
                                               "CODTuNop":  192.21,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 16A Lê Hồng Phong-Q.10-HCM"
                                           },
                                           {
                                               "TongCOD":  134.58,
                                               "CODTuNop":  134.58,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 178 Đường DT741-Bến Cát-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  77.37,
                                               "CODTuNop":  77.37,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 182 Xuyên Á-TP.Dĩ An-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  102.72,
                                               "CODTuNop":  102.72,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 183 Nguyễn Oanh-Phường 17-Gò Vấp-HCM"
                                           },
                                           {
                                               "TongCOD":  113.52,
                                               "CODTuNop":  113.52,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 185 Tân Thới Nhất 8-Phường Đông Hưng Thuận-HCM"
                                           },
                                           {
                                               "TongCOD":  130.75,
                                               "CODTuNop":  130.75,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 1855 Võ Văn Kiệt-Hoà Long-Bà Rịa Vũng Tàu"
                                           },
                                           {
                                               "TongCOD":  174.67,
                                               "CODTuNop":  174.67,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 195B Lý Chính Thắng-Q.3-HCM"
                                           },
                                           {
                                               "TongCOD":  150.54,
                                               "CODTuNop":  150.54,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 197 Lê Cao Lãng-Phú Thạnh-Tân Phú-HCM"
                                           },
                                           {
                                               "TongCOD":  64.05,
                                               "CODTuNop":  64.05,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 198 Thạnh Lộc 15-Q.12-HCM"
                                           },
                                           {
                                               "TongCOD":  313.49,
                                               "CODTuNop":  313.49,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 20 Đường số 81-Q.7-HCM"
                                           },
                                           {
                                               "TongCOD":  205.79,
                                               "CODTuNop":  205.79,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 20/48A Bờ Bao Tân Thắng-Q.Tân Phú-HCM"
                                           },
                                           {
                                               "TongCOD":  147.55,
                                               "CODTuNop":  147.55,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 21 Phú Mỹ-Q.Bình Thạnh-HCM"
                                           },
                                           {
                                               "TongCOD":  128.88,
                                               "CODTuNop":  128.88,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu cục 215 Bình Thới-Q.11-HCM"
                                           },
                                           {
                                               "TongCOD":  40.62,
                                               "CODTuNop":  40.62,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 2227 Văn Tiến Dũng-Bình Chánh-HCM"
                                           },
                                           {
                                               "TongCOD":  132.91,
                                               "CODTuNop":  132.91,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 235 Hồ Học Lãm-Q.Bình Tân-HCM"
                                           },
                                           {
                                               "TongCOD":  25.74,
                                               "CODTuNop":  25.74,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 242 Tỉnh Lộ 2-Củ Chi-HCM"
                                           },
                                           {
                                               "TongCOD":  199.88,
                                               "CODTuNop":  199.88,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 248A Nơ Trang Long-Q.Bình Thạnh-HCM"
                                           },
                                           {
                                               "TongCOD":  209.83,
                                               "CODTuNop":  209.83,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 24B Nguyễn Du-Phường 1-TP.Vũng Tàu-Bà Rịa Vũng Tàu"
                                           },
                                           {
                                               "TongCOD":  184.34,
                                               "CODTuNop":  184.34,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 256/59 Đường Thạnh Xuân 25-Quận 12-HCM"
                                           },
                                           {
                                               "TongCOD":  208.63,
                                               "CODTuNop":  208.63,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 257 Bùi Ngọc Thu-Chánh Hiệp-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  106.31,
                                               "CODTuNop":  106.31,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 26 Bùi Thanh Khiết-Tân Túc-Bình Chánh"
                                           },
                                           {
                                               "TongCOD":  72.71,
                                               "CODTuNop":  72.71,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 265 Lào Cai-Chánh Nghĩa-Thủ Dầu Một-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  114.73,
                                               "CODTuNop":  114.73,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 27 Nguyễn Xí-Q.Bình Thạnh-HCM"
                                           },
                                           {
                                               "TongCOD":  27.65,
                                               "CODTuNop":  27.65,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 274 Cô Bắc-Q.1-HCM"
                                           },
                                           {
                                               "TongCOD":  261.53,
                                               "CODTuNop":  261.53,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 291A Nguyễn Duy-Q.8-HCM"
                                           },
                                           {
                                               "TongCOD":  76.43,
                                               "CODTuNop":  76.43,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 295F Lý Thường Kiệt-Quận 11-HCM"
                                           },
                                           {
                                               "TongCOD":  228.07,
                                               "CODTuNop":  228.07,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 2999 Phạm Thế Hiển-Q.8-HCM"
                                           },
                                           {
                                               "TongCOD":  238.45,
                                               "CODTuNop":  238.45,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 29A Đường Số 5-Q.2-HCM"
                                           },
                                           {
                                               "TongCOD":  185.46,
                                               "CODTuNop":  185.46,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 309 Nguyễn Thị Ngâu-Hóc Môn-HCM"
                                           },
                                           {
                                               "TongCOD":  222.36,
                                               "CODTuNop":  222.36,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 31-31A Lý Phục Man-Bình Thuận-Quận 7-HCM"
                                           },
                                           {
                                               "TongCOD":  37.42,
                                               "CODTuNop":  37.42,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 320 Nguyễn Thị Kiểu-Hiệp Thành-Quận 12-HCM"
                                           },
                                           {
                                               "TongCOD":  109.38,
                                               "CODTuNop":  109.38,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 371 Nguyễn Kiệm-Q.Gò Vấp-HCM"
                                           },
                                           {
                                               "TongCOD":  66.09,
                                               "CODTuNop":  66.09,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 374 Xô Viết Nghệ Tĩnh-Q.Bình Thạnh-HCM"
                                           },
                                           {
                                               "TongCOD":  217.75,
                                               "CODTuNop":  217.75,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 376/38 Trảng Sắn-Phú Giáo-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  165.6,
                                               "CODTuNop":  165.6,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 38 Đường Số 3-Q.Bình Tân-HCM"
                                           },
                                           {
                                               "TongCOD":  32.54,
                                               "CODTuNop":  32.54,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 39 Tôn Đức Thắng-TP.Bà Rịa"
                                           },
                                           {
                                               "TongCOD":  260.62,
                                               "CODTuNop":  260.62,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 390 Nguyễn Kim Cương-Tân Thạnh Đông-Củ Chi"
                                           },
                                           {
                                               "TongCOD":  93.47,
                                               "CODTuNop":  93.47,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 41 Nguyễn Tất Thành-Long Điền-Bà Rịa Vũng Tàu"
                                           },
                                           {
                                               "TongCOD":  194.45,
                                               "CODTuNop":  194.45,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 456 Nguyễn Kiệm-Q.Phú Nhuận-HCM"
                                           },
                                           {
                                               "TongCOD":  180.64,
                                               "CODTuNop":  180.64,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 46 Tam Bình-Hiệp Bình Chánh-Q.Thủ Đức-HCM"
                                           },
                                           {
                                               "TongCOD":  98.59,
                                               "CODTuNop":  98.59,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 47 Đường 882-Phú Hữu-Q.9-HCM"
                                           },
                                           {
                                               "TongCOD":  108.56,
                                               "CODTuNop":  108.56,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 47/2 Lê Văn Khương-Q.12-HCM"
                                           },
                                           {
                                               "TongCOD":  303.14,
                                               "CODTuNop":  303.14,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 473A Bình Giã-TP.Vũng Tàu"
                                           },
                                           {
                                               "TongCOD":  227.37,
                                               "CODTuNop":  227.37,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 480F Cao Thắng - Trường Sa-Q.3-HCM"
                                           },
                                           {
                                               "TongCOD":  159.28,
                                               "CODTuNop":  159.28,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 503/16 Lê Thị Trung-Bình Chuẩn-TP.Thuận An-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  198.5,
                                               "CODTuNop":  198.5,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 53 Đường 22-Linh Đông-Thủ Đức-HCM"
                                           },
                                           {
                                               "TongCOD":  155.79,
                                               "CODTuNop":  155.79,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 53A2 Đường 5-Linh Tây-Q.Thủ Đức-HCM"
                                           },
                                           {
                                               "TongCOD":  117.32,
                                               "CODTuNop":  117.32,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 575 Đường 22/12-Thuận An-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  63.63,
                                               "CODTuNop":  63.63,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 57A Hùng Vương-Xã Ngãi Giao-Bà Rịa Vũng Tàu"
                                           },
                                           {
                                               "TongCOD":  81.14,
                                               "CODTuNop":  81.14,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 58 Lê Thúc Hoạch-Q.Tân Phú-HCM"
                                           },
                                           {
                                               "TongCOD":  137.53,
                                               "CODTuNop":  137.53,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 58 Trương Quốc Dung-Q.Phú Nhuận-HCM"
                                           },
                                           {
                                               "TongCOD":  164.5,
                                               "CODTuNop":  164.5,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu cục 58 Vũ Tông Phan-Q.2-HCM"
                                           },
                                           {
                                               "TongCOD":  178.29,
                                               "CODTuNop":  178.29,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 59 Trương Phước Phan-Q.Bình Tân-HCM"
                                           },
                                           {
                                               "TongCOD":  122.97,
                                               "CODTuNop":  122.97,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 593A Tỉnh Lộ 15-Củ Chi-HCM"
                                           },
                                           {
                                               "TongCOD":  92.08,
                                               "CODTuNop":  92.08,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 6 Phạm Văn Hai-Q.Tân Bình-HCM"
                                           },
                                           {
                                               "TongCOD":  139.86,
                                               "CODTuNop":  139.86,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 602/26 Điện Biên Phủ-Phường 12-Bình Thạnh-HCM"
                                           },
                                           {
                                               "TongCOD":  98.58,
                                               "CODTuNop":  98.58,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 608 Phan Văn Trị-Gò Vấp-HCM"
                                           },
                                           {
                                               "TongCOD":  141.54,
                                               "CODTuNop":  141.54,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 63 Đường Số 7-Xã Nhà Bè-HCM"
                                           },
                                           {
                                               "TongCOD":  170.66,
                                               "CODTuNop":  170.66,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 63 Ký Hòa-Q.5-HCM"
                                           },
                                           {
                                               "TongCOD":  115.77,
                                               "CODTuNop":  115.77,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 63 Vĩnh Hội-Q.4-HCM"
                                           },
                                           {
                                               "TongCOD":  35.37,
                                               "CODTuNop":  35.37,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 63A Trần Đình Xu-Q.1-HCM"
                                           },
                                           {
                                               "TongCOD":  108.57,
                                               "CODTuNop":  108.57,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 692 An Phú Tây-Bình Chánh-HCM"
                                           },
                                           {
                                               "TongCOD":  233.95,
                                               "CODTuNop":  233.95,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 75 Đường Số 2 KDC Nguyên Sơn-Bình Hưng-Bình Chánh"
                                           },
                                           {
                                               "TongCOD":  163.8,
                                               "CODTuNop":  163.8,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 78 Nguyễn Minh Hoàng-Phường 12-Q.Tân Bình-HCM"
                                           },
                                           {
                                               "TongCOD":  130.83,
                                               "CODTuNop":  130.83,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 79 Nguyễn Hiền-Dĩ An-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  124.04,
                                               "CODTuNop":  124.04,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 7A/3 Thành Thái-Phường Diên Hồng-HCM"
                                           },
                                           {
                                               "TongCOD":  118.77,
                                               "CODTuNop":  118.77,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 82 Nguyễn Sỹ Sách-Q.Tân Bình-HCM"
                                           },
                                           {
                                               "TongCOD":  151.39,
                                               "CODTuNop":  151.39,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 85 Cao Văn Ngọc-Phường Tân Phú-HCM"
                                           },
                                           {
                                               "TongCOD":  330.65,
                                               "CODTuNop":  330.65,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 860/36 Huỳnh Tấn Phát-Q.7-HCM"
                                           },
                                           {
                                               "TongCOD":  191.69,
                                               "CODTuNop":  191.69,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 87 Nguyễn An Ninh-Tp.Vũng Tàu-Bà Rịa-Vũng Tàu"
                                           },
                                           {
                                               "TongCOD":  156.94,
                                               "CODTuNop":  156.94,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 870 Huỳnh Văn Lũy-Thủ Dầu Một-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  12.34,
                                               "CODTuNop":  12.34,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 88 Võ Thị Liễu-Q.12-HCM"
                                           },
                                           {
                                               "TongCOD":  227.45,
                                               "CODTuNop":  227.45,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 88/21 Phan Sào Nam-Q.Tân Bình-HCM"
                                           },
                                           {
                                               "TongCOD":  194.46,
                                               "CODTuNop":  194.46,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 8F Gò Xoài-Bình Hưng Hòa A-Bình Tân-HCM"
                                           },
                                           {
                                               "TongCOD":  96.9,
                                               "CODTuNop":  96.9,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 9 Hồng Bàng-Q.Bình Thạnh-HCM"
                                           },
                                           {
                                               "TongCOD":  253.13,
                                               "CODTuNop":  253.13,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 903B Tạ Quang Bửu-Q.8-HCM"
                                           },
                                           {
                                               "TongCOD":  92.55,
                                               "CODTuNop":  92.55,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 92 Tân Chánh Hiệp 03-Tân Chánh Hiệp-Quận 12-HCM"
                                           },
                                           {
                                               "TongCOD":  64.26,
                                               "CODTuNop":  64.26,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 95 Nguyễn Thị Tươi-Tân Bình-Dĩ An-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  259.04,
                                               "CODTuNop":  259.04,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục A4/18 Đường D10-Thuận An-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  51.09,
                                               "CODTuNop":  51.09,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục An Phước- An Ngãi - Bà Rịa Vũng Tàu"
                                           },
                                           {
                                               "TongCOD":  29.3,
                                               "CODTuNop":  29.3,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Ấp Phú Thọ-Xuyên Mộc-Bà Rịa Vũng Tàu"
                                           },
                                           {
                                               "TongCOD":  43.38,
                                               "CODTuNop":  43.38,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Ấp Phước Sơn-Phú Mỹ-Bà Rịa Vũng Tàu"
                                           },
                                           {
                                               "TongCOD":  306.9,
                                               "CODTuNop":  306.9,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Bàu Bàng-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  138.06,
                                               "CODTuNop":  138.06,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục C1/7 Lại Hùng Cường-Vĩnh Lộc B-Bình Chánh-HCM"
                                           },
                                           {
                                               "TongCOD":  97.26,
                                               "CODTuNop":  97.26,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Cần Giờ-HCM"
                                           },
                                           {
                                               "TongCOD":  0.4,
                                               "CODTuNop":  0.4,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Chuyên Lấy 01 Đường Số 6-Long Bình-Q.9-HCM"
                                           },
                                           {
                                               "TongCOD":  143.01,
                                               "CODTuNop":  143.01,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Côn Đảo-Bà Rịa Vũng Tàu"
                                           },
                                           {
                                               "TongCOD":  259.63,
                                               "CODTuNop":  259.63,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu cục D1-D39 KDC VietSing-An Phú-Thuận An-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  159.45,
                                               "CODTuNop":  159.45,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục DA 1-1 Mỹ Phước-Bến Cát-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  174.51,
                                               "CODTuNop":  174.51,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Đông Chiêu-Tân Đông Hiệp-Dĩ An-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  63,
                                               "CODTuNop":  63,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục DT 744-An Tây-Bến Cát-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  178.66,
                                               "CODTuNop":  178.66,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Đường 21-Phường 08-Gò Vấp-HCM"
                                           },
                                           {
                                               "TongCOD":  195.21,
                                               "CODTuNop":  195.21,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Đường 268 Đỗ Xuân Hợp-Phường Phước Long-HCM"
                                           },
                                           {
                                               "TongCOD":  190.53,
                                               "CODTuNop":  190.53,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Đường 743B-Thuận An-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  199.95,
                                               "CODTuNop":  199.95,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Đường NH17-Hòa Phú-Thủ Dầu Một-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  74.67,
                                               "CODTuNop":  74.67,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Đường NL12-Phường Thới Hòa-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  147.93,
                                               "CODTuNop":  147.93,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Đường Số 2 KDC Genimex-TP.Tân Uyên-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  89.55,
                                               "CODTuNop":  89.55,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục F5/9A Trần Văn Giàu-Bình Chánh-HCM"
                                           },
                                           {
                                               "TongCOD":  70.49,
                                               "CODTuNop":  70.49,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa 1130 Đông Hưng Thuận 05-Q.12-HCM"
                                           },
                                           {
                                               "TongCOD":  44.94,
                                               "CODTuNop":  44.94,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa 16A Lê Hồng Phong-Q.10-HCM"
                                           },
                                           {
                                               "TongCOD":  7.33,
                                               "CODTuNop":  7.33,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa 198C Thạnh Lộc-Q.12-HCM"
                                           },
                                           {
                                               "TongCOD":  25.27,
                                               "CODTuNop":  25.27,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa 20A Đường 12-Q.9-HCM"
                                           },
                                           {
                                               "TongCOD":  39.78,
                                               "CODTuNop":  39.78,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa 215 Bình Thới-Q.11-HCM"
                                           },
                                           {
                                               "TongCOD":  76.11,
                                               "CODTuNop":  76.11,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa 248A Nơ Trang Long-Q.Bình Thạnh-HCM"
                                           },
                                           {
                                               "TongCOD":  76.39,
                                               "CODTuNop":  76.39,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa 320 Nguyễn Thị Kiểu-Hiệp Thành-Quận 12-HCM"
                                           },
                                           {
                                               "TongCOD":  54.89,
                                               "CODTuNop":  54.89,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa 367 Lương Định Của-An Phú-Q.2-HCM"
                                           },
                                           {
                                               "TongCOD":  41.43,
                                               "CODTuNop":  41.43,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa 6 Phạm Văn Hai-Q.Tân Bình-HCM"
                                           },
                                           {
                                               "TongCOD":  60.96,
                                               "CODTuNop":  60.96,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu cục Hàng Vừa 78 Đường số 4-Q.7-HCM"
                                           },
                                           {
                                               "TongCOD":  91.31,
                                               "CODTuNop":  91.31,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa 85 Cao Văn Ngọc-Q.Tân Phú-HCM"
                                           },
                                           {
                                               "TongCOD":  108.72,
                                               "CODTuNop":  108.72,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa 971 Nguyễn Kiệm-Q.Gò Vấp-HCM"
                                           },
                                           {
                                               "TongCOD":  86.86,
                                               "CODTuNop":  86.86,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa A5 KDC Tân Thuận Nam-Phú Thuận-Quận 7-HCM"
                                           },
                                           {
                                               "TongCOD":  37.57,
                                               "CODTuNop":  37.57,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa Bình Tân Kho 01 1162/6A Tỉnh Lộ 10-HCM"
                                           },
                                           {
                                               "TongCOD":  20.32,
                                               "CODTuNop":  20.32,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa Bình Tân Kho 02 1162/6A Tỉnh Lộ 10-HCM"
                                           },
                                           {
                                               "TongCOD":  32.04,
                                               "CODTuNop":  32.04,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa Đường 268 Đỗ Xuân Hợp-Q.9-HCM"
                                           },
                                           {
                                               "TongCOD":  80.85,
                                               "CODTuNop":  80.85,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa Kho 02 367 Lương Định Của-Q.2-HCM"
                                           },
                                           {
                                               "TongCOD":  0.05,
                                               "CODTuNop":  0.05,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa Q.7 Kho 01 Số 860/36 Huỳnh Tấn Phát-Q.7-HCM"
                                           },
                                           {
                                               "TongCOD":  93.48,
                                               "CODTuNop":  93.48,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa Q.Thủ Đức Kho 01 13 Đường Số 8-HCM"
                                           },
                                           {
                                               "TongCOD":  101.01,
                                               "CODTuNop":  101.01,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa Quận 3 480F Cao Thắng-Q.10-HCM"
                                           },
                                           {
                                               "TongCOD":  7.8,
                                               "CODTuNop":  7.8,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa Vĩnh Lộc B-Bình Chánh-HCM"
                                           },
                                           {
                                               "TongCOD":  133.62,
                                               "CODTuNop":  133.62,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu cục Hội Nghĩa-Tân Uyên-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  29.64,
                                               "CODTuNop":  29.64,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hương Lộ 2-Bình Trị Đông A-Bình Tân"
                                           },
                                           {
                                               "TongCOD":  69.68,
                                               "CODTuNop":  69.68,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục KDC Đồng Chàm-Dĩ An-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  60.41,
                                               "CODTuNop":  60.41,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục KDC Giồng Ao-Cần Giờ-HCM"
                                           },
                                           {
                                               "TongCOD":  72.47,
                                               "CODTuNop":  72.47,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Khu Phố 4-Bắc Tân Uyên-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  92.22,
                                               "CODTuNop":  92.22,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Khu Phố Vạn Hạnh-Phú Mỹ-Bà Rịa Vũng Tàu"
                                           },
                                           {
                                               "TongCOD":  75.41,
                                               "CODTuNop":  75.41,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục KP Thanh Long -Đất Đỏ-Bà Rịa Vũng Tàu"
                                           },
                                           {
                                               "TongCOD":  75.82,
                                               "CODTuNop":  75.82,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục KP Trảng Lớn-Phú Mỹ-Hắc Dịch-Bà Rịa Vũng Tàu"
                                           },
                                           {
                                               "TongCOD":  158.27,
                                               "CODTuNop":  158.27,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục KP6 Vĩnh Tân-Tân Uyên-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  118.04,
                                               "CODTuNop":  118.04,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu cục Long Thạnh Mỹ-Q9-HCM"
                                           },
                                           {
                                               "TongCOD":  111.77,
                                               "CODTuNop":  111.77,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Phan Huy Ích-Gò Vấp-HCM"
                                           },
                                           {
                                               "TongCOD":  275.34,
                                               "CODTuNop":  275.34,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Phú Lợi-TP.Thủ Dầu Một-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  225.36,
                                               "CODTuNop":  225.36,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Phước Kiển-Nhà Bè-HCM"
                                           },
                                           {
                                               "TongCOD":  302.72,
                                               "CODTuNop":  302.72,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Phường Lái Thiêu-Thuận An-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  129.71,
                                               "CODTuNop":  129.71,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Phường Long Trường-Q9-HCM"
                                           },
                                           {
                                               "TongCOD":  89.19,
                                               "CODTuNop":  89.19,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Phường Tây Thạnh-Tân Phú-HCM"
                                           },
                                           {
                                               "TongCOD":  143.7,
                                               "CODTuNop":  143.7,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Số 17 Đường Hoàng Văn Hợp-An Lạc A-Bình Tân-HCM"
                                           },
                                           {
                                               "TongCOD":  262.92,
                                               "CODTuNop":  262.92,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Số 25 Đường 12-An Phú-Q.2-HCM"
                                           },
                                           {
                                               "TongCOD":  125.23,
                                               "CODTuNop":  125.23,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu cục Số 54/1 Đường Tân Phước Khánh 13-Tân Uyên-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  188.02,
                                               "CODTuNop":  188.02,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Số 8 Lê Chí Dân-Thủ Dầu Một-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  127.75,
                                               "CODTuNop":  127.75,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục T30 Nguyễn Hữu Hào-Q.4-HCM"
                                           },
                                           {
                                               "TongCOD":  168.32,
                                               "CODTuNop":  168.32,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Tân Hiệp 04-Tân Uyên-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  61.96,
                                               "CODTuNop":  61.96,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Tân Kiên-Bình Chánh-HCM"
                                           },
                                           {
                                               "TongCOD":  174.42,
                                               "CODTuNop":  174.42,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Tăng Nhơn Phú B-Q.9-HCM"
                                           },
                                           {
                                               "TongCOD":  75.5,
                                               "CODTuNop":  75.5,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thôn Thành Long-Châu Đức-Bà Rịa Vũng Tàu"
                                           },
                                           {
                                               "TongCOD":  17.45,
                                               "CODTuNop":  17.45,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Trừ Văn Thố-Bàu Bàng-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  160,
                                               "CODTuNop":  160,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Vĩnh Lộc A-Bình Chánh-HCM"
                                           },
                                           {
                                               "TongCOD":  100.8,
                                               "CODTuNop":  100.8,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu cục Xuân Thới Thượng-Hóc Môn-HCM"
                                           },
                                           {
                                               "TongCOD":  286.28,
                                               "CODTuNop":  286.28,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Xuyên Mộc-Bà Rịa Vũng Tàu"
                                           }
                                       ]
                    },
                    {
                        "TyLeTuNop":  100,
                        "TenVung":  "HNO",
                        "TongCOD":  6370.38,
                        "Trend":  0,
                        "CODTuNop":  6370.38,
                        "BuuCucList":  [
                                           {
                                               "TongCOD":  1.92,
                                               "CODTuNop":  1.92,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 10 Vĩnh Tuy-Q.Hai Bà Trưng-HN"
                                           },
                                           {
                                               "TongCOD":  2.05,
                                               "CODTuNop":  2.05,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 1057/9/26 Ngãi Cầu-Hoài Đức-HN"
                                           },
                                           {
                                               "TongCOD":  7.33,
                                               "CODTuNop":  7.33,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 108 Phố Nhổn-Q.Bắc Từ Liêm-Hà Nội"
                                           },
                                           {
                                               "TongCOD":  5.07,
                                               "CODTuNop":  5.07,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 109 E3-Q.Hai Bà Trưng-HN"
                                           },
                                           {
                                               "TongCOD":  56.92,
                                               "CODTuNop":  56.92,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 114 Yên Lãng-Q.Đống Đa-HN"
                                           },
                                           {
                                               "TongCOD":  101.8,
                                               "CODTuNop":  101.8,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu cục 126 Hoàng Ngân-Q.Cầu Giấy-HN"
                                           },
                                           {
                                               "TongCOD":  0.02,
                                               "CODTuNop":  0.02,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 131 Quan Hoa-Q.Cầu Giấy-HN"
                                           },
                                           {
                                               "TongCOD":  72.03,
                                               "CODTuNop":  72.03,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 139 Thụy Phương-Q.Bắc Từ Liêm-HN"
                                           },
                                           {
                                               "TongCOD":  3.01,
                                               "CODTuNop":  3.01,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 14 Huỳnh Thúc Kháng-Q.Đống Đa-HN"
                                           },
                                           {
                                               "TongCOD":  2.4,
                                               "CODTuNop":  2.4,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 142A Khuất Duy Tiến-Q.Thanh Xuân-HN"
                                           },
                                           {
                                               "TongCOD":  200.29,
                                               "CODTuNop":  200.29,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu cục 148 Trâu Quỳ-Gia Lâm-HN"
                                           },
                                           {
                                               "TongCOD":  122.13,
                                               "CODTuNop":  122.13,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 15 Hàng Bún-Q.Ba Đình-HN"
                                           },
                                           {
                                               "TongCOD":  55.86,
                                               "CODTuNop":  55.86,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 155 Phú Diễn-Phú Diễn-Bắc Từ Liêm-HN"
                                           },
                                           {
                                               "TongCOD":  78.09,
                                               "CODTuNop":  78.09,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 16A Hòa Mã -Q.Hai Bà Trưng-HN"
                                           },
                                           {
                                               "TongCOD":  18.75,
                                               "CODTuNop":  18.75,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 17 Phú Thượng-Q.Tây Hồ-HN"
                                           },
                                           {
                                               "TongCOD":  11.46,
                                               "CODTuNop":  11.46,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 172 Cầu Gáo-Đan Phượng-HN"
                                           },
                                           {
                                               "TongCOD":  2.35,
                                               "CODTuNop":  2.35,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 172 Võ Chí Công-Xuân La-Tây Hồ-HN"
                                           },
                                           {
                                               "TongCOD":  79.75,
                                               "CODTuNop":  79.75,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 188 Khương Đình-Q.Thanh Xuân-HN"
                                           },
                                           {
                                               "TongCOD":  149.98,
                                               "CODTuNop":  149.98,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 20 Đông Quan-Q.Cầu Giấy-HN"
                                           },
                                           {
                                               "TongCOD":  27.9,
                                               "CODTuNop":  27.9,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 203 Tỉnh Lộ 417-Phúc Thọ-HN"
                                           },
                                           {
                                               "TongCOD":  9.02,
                                               "CODTuNop":  9.02,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 224 Lê Trọng Tấn-Q.Hà Đông-HN"
                                           },
                                           {
                                               "TongCOD":  165.41,
                                               "CODTuNop":  165.41,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 229 Quan Nhân-Q.Thanh Xuân-HN"
                                           },
                                           {
                                               "TongCOD":  75.26,
                                               "CODTuNop":  75.26,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 24 ngõ 80 La Nội-Q.Hà Đông-HN"
                                           },
                                           {
                                               "TongCOD":  12.58,
                                               "CODTuNop":  12.58,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu cục 253 Phạm Văn Đồng-Q.Bắc Từ Liêm-Hà Nội"
                                           },
                                           {
                                               "TongCOD":  89.12,
                                               "CODTuNop":  89.12,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 26 Đại Mỗ-Q.Nam Từ Liêm-HN"
                                           },
                                           {
                                               "TongCOD":  2.92,
                                               "CODTuNop":  2.92,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 263 Phố Vọng-Q.Hai Bà Trưng-HN"
                                           },
                                           {
                                               "TongCOD":  59.96,
                                               "CODTuNop":  59.96,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 28 Đường Mới-Phụng Thượng-Phúc Thọ-HN"
                                           },
                                           {
                                               "TongCOD":  34.74,
                                               "CODTuNop":  34.74,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 29 Hưng Phúc-Yên Sở-Q.Hoàng Mai-HN"
                                           },
                                           {
                                               "TongCOD":  144.72,
                                               "CODTuNop":  144.72,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu cục 296 đường Ỷ Lan-Gia Lâm-HN"
                                           },
                                           {
                                               "TongCOD":  60.97,
                                               "CODTuNop":  60.97,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 298 Trần Điền-Định Công-Hoàng Mai-Hà Nội"
                                           },
                                           {
                                               "TongCOD":  9.8,
                                               "CODTuNop":  9.8,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 303 Hoàng Hoa Thám-Q.Ba Đình-HN"
                                           },
                                           {
                                               "TongCOD":  18.84,
                                               "CODTuNop":  18.84,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 325 LouisI-Q.Hoàng Mai-HN"
                                           },
                                           {
                                               "TongCOD":  71.3,
                                               "CODTuNop":  71.3,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 33 Đặng Trần Côn-Q.Đống Đa-HN"
                                           },
                                           {
                                               "TongCOD":  71.81,
                                               "CODTuNop":  71.81,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 35 Liền Kề 9 KĐT Đại Thanh-Thanh Trì-HN"
                                           },
                                           {
                                               "TongCOD":  127.28,
                                               "CODTuNop":  127.28,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 36 Lý Chiêu Hoàng-Ninh Hiệp-Gia Lâm-Hà Nội"
                                           },
                                           {
                                               "TongCOD":  23.98,
                                               "CODTuNop":  23.98,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 366 Phạm Văn Đồng-Q.Bắc Từ Liêm-HN"
                                           },
                                           {
                                               "TongCOD":  27.72,
                                               "CODTuNop":  27.72,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 39 Gò Sỏi-Sóc Sơn-HN"
                                           },
                                           {
                                               "TongCOD":  7.26,
                                               "CODTuNop":  7.26,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 39 KTĐC Tứ Hiệp-Thanh Trì-HN"
                                           },
                                           {
                                               "TongCOD":  12.37,
                                               "CODTuNop":  12.37,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 396 Kim Giang-Q.Thanh Xuân-HN"
                                           },
                                           {
                                               "TongCOD":  62.22,
                                               "CODTuNop":  62.22,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 41 Phố Nguyễn Khả Trạc-Q.Cầu Giấy-HN"
                                           },
                                           {
                                               "TongCOD":  73.49,
                                               "CODTuNop":  73.49,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu cục 41 Trần Quốc Vượng-Q.Cầu giấy-HN"
                                           },
                                           {
                                               "TongCOD":  49.85,
                                               "CODTuNop":  49.85,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 43 Trần Đăng Ninh-Ứng Hòa-HN"
                                           },
                                           {
                                               "TongCOD":  11.59,
                                               "CODTuNop":  11.59,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 433 Tân Phú-Sơn Tây-HN"
                                           },
                                           {
                                               "TongCOD":  151.66,
                                               "CODTuNop":  151.66,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu cục 442 Nguyễn Trãi-Q.Thanh Xuân-HN"
                                           },
                                           {
                                               "TongCOD":  35.06,
                                               "CODTuNop":  35.06,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 450 Bờ Tây Sông Nhuệ-Q.Bắc Từ Liêm-HN"
                                           },
                                           {
                                               "TongCOD":  80.08,
                                               "CODTuNop":  80.08,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 51A Tiến Thành-Q.Hà Đông-HN"
                                           },
                                           {
                                               "TongCOD":  26.55,
                                               "CODTuNop":  26.55,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 55C Nguyễn Khoái-Q.Hai Bà Trưng-HN"
                                           },
                                           {
                                               "TongCOD":  79.98,
                                               "CODTuNop":  79.98,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 57 Thao Chính-Phú Xuyên-HN"
                                           },
                                           {
                                               "TongCOD":  100.91,
                                               "CODTuNop":  100.91,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 60 Sài Đồng-Q.Long Biên-Hà Nội"
                                           },
                                           {
                                               "TongCOD":  20.1,
                                               "CODTuNop":  20.1,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu cục 602 Hà Huy Tập-Gia Lâm-HN"
                                           },
                                           {
                                               "TongCOD":  0.51,
                                               "CODTuNop":  0.51,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 61 Kim Mã-Q.Ba Đình-HN"
                                           },
                                           {
                                               "TongCOD":  176.2,
                                               "CODTuNop":  176.2,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 63 Đường Bờ Sông Sét-Thịnh Liệt-Hoàng Mai-HN"
                                           },
                                           {
                                               "TongCOD":  31.4,
                                               "CODTuNop":  31.4,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 709 Tam Trinh-Q.Hoàng Mai-HN"
                                           },
                                           {
                                               "TongCOD":  6.58,
                                               "CODTuNop":  6.58,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 79 Thanh Xuân Làng Nghề Liên Hà-Đan Phượng-HN"
                                           },
                                           {
                                               "TongCOD":  45.45,
                                               "CODTuNop":  45.45,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 79/9B Thanh Đàm-Q.Hoàng Mai-HN"
                                           },
                                           {
                                               "TongCOD":  0.69,
                                               "CODTuNop":  0.69,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 7C10 Nguyễn Phong Sắc-Cầu Giấy-Hà Nội"
                                           },
                                           {
                                               "TongCOD":  68.3,
                                               "CODTuNop":  68.3,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu cục 8.03-TT8 Xuân Phương-Q.Nam Từ Liêm-HN"
                                           },
                                           {
                                               "TongCOD":  81.33,
                                               "CODTuNop":  81.33,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 84 Đặng Xuân Bảng-Đại Kim-Hoàng Mai-HN"
                                           },
                                           {
                                               "TongCOD":  94.06,
                                               "CODTuNop":  94.06,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 849 Ngô Gia Tự-Đức Giang-Long Biên-HN"
                                           },
                                           {
                                               "TongCOD":  37.56,
                                               "CODTuNop":  37.56,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 88 Bạch Đằng-Hoàn Kiếm-HN"
                                           },
                                           {
                                               "TongCOD":  9.3,
                                               "CODTuNop":  9.3,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 882 Kim Giang-Thanh Trì-HN"
                                           },
                                           {
                                               "TongCOD":  105.62,
                                               "CODTuNop":  105.62,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu cục 896 Minh Khai-Q.Hai Bà Trưng-HN"
                                           },
                                           {
                                               "TongCOD":  23.86,
                                               "CODTuNop":  23.86,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 8B Nguyễn Lân-Q.Thanh Xuân-HN"
                                           },
                                           {
                                               "TongCOD":  67.95,
                                               "CODTuNop":  67.95,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 90 Trung Phụng-Q.Đống Đa-HN"
                                           },
                                           {
                                               "TongCOD":  207.35,
                                               "CODTuNop":  207.35,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 984 Đê La Thành-Ba Đình-HN"
                                           },
                                           {
                                               "TongCOD":  97.86,
                                               "CODTuNop":  97.86,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Bê Tông-Thuỷ Xuân Tiên-Chương Mỹ-HN"
                                           },
                                           {
                                               "TongCOD":  80.29,
                                               "CODTuNop":  80.29,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Chúc Lý-Ngọc Hòa-Chương Mỹ-HN"
                                           },
                                           {
                                               "TongCOD":  102.17,
                                               "CODTuNop":  102.17,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục CL9 Đường Tràng An-Bát Tràng-HN"
                                           },
                                           {
                                               "TongCOD":  81.59,
                                               "CODTuNop":  81.59,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Cổ Dương-Tiên Dương-Đông Anh-HN"
                                           },
                                           {
                                               "TongCOD":  20.59,
                                               "CODTuNop":  20.59,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Cồng Kềnh 1057/9/26 Ngãi Cầu-Hoài Đức-HN"
                                           },
                                           {
                                               "TongCOD":  4.51,
                                               "CODTuNop":  4.51,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Cồng Kềnh 17 Phố Phú Thượng-Q.Tây Hồ-HN"
                                           },
                                           {
                                               "TongCOD":  121.1,
                                               "CODTuNop":  121.1,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Cồng Kềnh 270 Nguyễn Xiển-Thanh Xuân-HN"
                                           },
                                           {
                                               "TongCOD":  36.77,
                                               "CODTuNop":  36.77,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Cồng Kềnh 51A Tiến Thành-Q.Hà Đông-HN"
                                           },
                                           {
                                               "TongCOD":  5,
                                               "CODTuNop":  5,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Cồng Kềnh 79 Thanh Xuân Làng Nghề Liên Hà-Đan Phượng-HN"
                                           },
                                           {
                                               "TongCOD":  0.35,
                                               "CODTuNop":  0.35,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Cồng Kềnh Cổ Dương-Đông Anh-HN"
                                           },
                                           {
                                               "TongCOD":  59.34,
                                               "CODTuNop":  59.34,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Cồng Kềnh Khu Công Nghiệp HNT-Đông Anh-HN"
                                           },
                                           {
                                               "TongCOD":  74.37,
                                               "CODTuNop":  74.37,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Cồng Kềnh Ngõ 134 Sài Đồng-Q.Long Biên-HN"
                                           },
                                           {
                                               "TongCOD":  6.92,
                                               "CODTuNop":  6.92,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Cồng Kềnh thôn Hải Bối-xã Vĩnh Thanh-HN"
                                           },
                                           {
                                               "TongCOD":  126.17,
                                               "CODTuNop":  126.17,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Đại Nghĩa-Mỹ Đức-HN"
                                           },
                                           {
                                               "TongCOD":  90.77,
                                               "CODTuNop":  90.77,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Đường Băng Xuân Bách-Xã Nội Bài-HN"
                                           },
                                           {
                                               "TongCOD":  30.66,
                                               "CODTuNop":  30.66,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Đường Núi Đôi-Sóc Sơn-HN"
                                           },
                                           {
                                               "TongCOD":  55.37,
                                               "CODTuNop":  55.37,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hà Trì-Hà Cầu-Hà Đông-HN"
                                           },
                                           {
                                               "TongCOD":  81.75,
                                               "CODTuNop":  81.75,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa 1267 Giải Phóng-Q.Hoàng Mai-HN"
                                           },
                                           {
                                               "TongCOD":  39.76,
                                               "CODTuNop":  39.76,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa 30 Phố Xã Đàn-Q.Đống Đa-HN"
                                           },
                                           {
                                               "TongCOD":  27.03,
                                               "CODTuNop":  27.03,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa 339 Trâu Quỳ-H.Gia Lâm-HN"
                                           },
                                           {
                                               "TongCOD":  145.36,
                                               "CODTuNop":  145.36,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa 37 TT9 KĐT Văn Phú-Q.Hà Đông-HN"
                                           },
                                           {
                                               "TongCOD":  7.27,
                                               "CODTuNop":  7.27,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa 4 KĐT Sinh Thái Xuân Phương-Q.Nam Từ Liêm-HN"
                                           },
                                           {
                                               "TongCOD":  60.02,
                                               "CODTuNop":  60.02,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa 450A Bờ Tây Sông Nhuệ-Q.Bắc Từ Liêm-HN"
                                           },
                                           {
                                               "TongCOD":  5.44,
                                               "CODTuNop":  5.44,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa 54 Ngõ 80 La Nội-Phường Dương Nội-HN"
                                           },
                                           {
                                               "TongCOD":  20.91,
                                               "CODTuNop":  20.91,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa 96 Nghĩa Tân-Q.Cầu Giấy-HN"
                                           },
                                           {
                                               "TongCOD":  69.47,
                                               "CODTuNop":  69.47,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa G4.P2 Cổ Linh-Q.Long Biên-HN"
                                           },
                                           {
                                               "TongCOD":  90.04,
                                               "CODTuNop":  90.04,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hòa Lâm-Ứng Hòa-Hà Nội"
                                           },
                                           {
                                               "TongCOD":  56.5,
                                               "CODTuNop":  56.5,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Khu tái định cư X6-Thạch Bàn-HN"
                                           },
                                           {
                                               "TongCOD":  87.91,
                                               "CODTuNop":  87.91,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Lô DM10-6 Làng nghề Vạn Phúc-Q.Hà Đông-HN"
                                           },
                                           {
                                               "TongCOD":  148.63,
                                               "CODTuNop":  148.63,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Lưu Quang Vũ-Trung Hòa-Cầu Giấy-HN"
                                           },
                                           {
                                               "TongCOD":  61.24,
                                               "CODTuNop":  61.24,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Số 1 Tân Khai-Q.Hoàng Mai-HN"
                                           },
                                           {
                                               "TongCOD":  8.21,
                                               "CODTuNop":  8.21,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu cục Thạch Hòa-Thạch Thất-Hà Nội"
                                           },
                                           {
                                               "TongCOD":  10.89,
                                               "CODTuNop":  10.89,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thọ Giáo-Tân Minh-Thường Tín-HN"
                                           },
                                           {
                                               "TongCOD":  55.52,
                                               "CODTuNop":  55.52,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu cục Thôn 6-Ba Vì-HN"
                                           },
                                           {
                                               "TongCOD":  92.38,
                                               "CODTuNop":  92.38,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thôn An Khoái-Phúc Tiến-Phú Xuyên-HN"
                                           },
                                           {
                                               "TongCOD":  59.87,
                                               "CODTuNop":  59.87,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thôn Đại Bái-Đại Thịnh-Mê Linh-HNO"
                                           },
                                           {
                                               "TongCOD":  32.02,
                                               "CODTuNop":  32.02,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu cục Thôn Đông-Ba Vì-Hà Nội"
                                           },
                                           {
                                               "TongCOD":  101.94,
                                               "CODTuNop":  101.94,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thôn Mới-Cao Dương-Thanh Oai-HN"
                                           },
                                           {
                                               "TongCOD":  44.79,
                                               "CODTuNop":  44.79,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thôn Phú Mỹ-Tự Lập-Mê Linh-HN"
                                           },
                                           {
                                               "TongCOD":  160.86,
                                               "CODTuNop":  160.86,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Tổ 5 Giang Biên-Long Biên-HN"
                                           },
                                           {
                                               "TongCOD":  210.96,
                                               "CODTuNop":  210.96,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục TT1 Khu tái định cư Ngọc Thụy-Long Biên-HN"
                                           },
                                           {
                                               "TongCOD":  26.02,
                                               "CODTuNop":  26.02,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Xã Thắng Lợi-Thường Tín-HN"
                                           }
                                       ]
                    },
                    {
                        "TyLeTuNop":  100,
                        "TenVung":  "KHL",
                        "TongCOD":  0.03,
                        "Trend":  100,
                        "CODTuNop":  0.03,
                        "BuuCucList":  {
                                           "TongCOD":  0.03,
                                           "CODTuNop":  0.03,
                                           "TyLeTuNop":  100,
                                           "TenBuuCuc":  "Key Account Warehouse Ho Chi Minh"
                                       }
                    },
                    {
                        "TyLeTuNop":  100,
                        "TenVung":  "NTB",
                        "TongCOD":  3124.21,
                        "Trend":  0,
                        "CODTuNop":  3124.21,
                        "BuuCucList":  [
                                           {
                                               "TongCOD":  129.48,
                                               "CODTuNop":  129.48,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 06 Lê Hồng Phong-TP.Nha Trang-Khánh Hòa"
                                           },
                                           {
                                               "TongCOD":  17.63,
                                               "CODTuNop":  17.63,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 1322 Hùng Vương-Di Linh-Lâm Đồng"
                                           },
                                           {
                                               "TongCOD":  30.24,
                                               "CODTuNop":  30.24,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 158 Quốc lộ 55-Tánh Linh-Bình Thuận"
                                           },
                                           {
                                               "TongCOD":  5.94,
                                               "CODTuNop":  5.94,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 16 Huỳnh Thúc Kháng-Phan Thiết-Bình Thuận"
                                           },
                                           {
                                               "TongCOD":  121.68,
                                               "CODTuNop":  121.68,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 190/11 Phù Đổng Thiên Vương-Phường 8-Đà Lạt-Lâm Đồng"
                                           },
                                           {
                                               "TongCOD":  171.56,
                                               "CODTuNop":  171.56,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 195 Đường 2/4-Nha Trang-Khánh Hòa"
                                           },
                                           {
                                               "TongCOD":  122.61,
                                               "CODTuNop":  122.61,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 21 Trần Hưng Đạo-Xã Phan Rí Cửa-Bình Thuận"
                                           },
                                           {
                                               "TongCOD":  23.85,
                                               "CODTuNop":  23.85,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 237 Nguyễn Tất Thành-Bắc Bình-Bình Thuận"
                                           },
                                           {
                                               "TongCOD":  209.5,
                                               "CODTuNop":  209.5,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 254A Đường 21 tháng 8-Phan Rang-Tháp Chàm-Ninh Thuận"
                                           },
                                           {
                                               "TongCOD":  166.91,
                                               "CODTuNop":  166.91,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 26A Lê Duẩn-Ninh Sơn-Ninh Thuận"
                                           },
                                           {
                                               "TongCOD":  74.78,
                                               "CODTuNop":  74.78,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 40 Nguyễn Văn Linh-Cư Jút-Đắk Nông"
                                           },
                                           {
                                               "TongCOD":  28.21,
                                               "CODTuNop":  28.21,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 40A Yết Kiêu-Nha Trang-Khánh Hòa"
                                           },
                                           {
                                               "TongCOD":  91.42,
                                               "CODTuNop":  91.42,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 412 Nguyễn Văn Cừ-Lộc Phát-Bảo Lộc-Lâm Đồng"
                                           },
                                           {
                                               "TongCOD":  32.4,
                                               "CODTuNop":  32.4,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 42 Nguyễn Du-Cam Lâm-Khánh Hoà"
                                           },
                                           {
                                               "TongCOD":  112.65,
                                               "CODTuNop":  112.65,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 430 Tôn Đức Thắng-Tân Phước-Thị xã La Gi-Bình Thuận"
                                           },
                                           {
                                               "TongCOD":  14.47,
                                               "CODTuNop":  14.47,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 44B Nguyễn Đình Chiểu-Đạ Teh-Lâm Đồng"
                                           },
                                           {
                                               "TongCOD":  134.17,
                                               "CODTuNop":  134.17,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 454 Võ Văn Kiệt-Đặc Khu Đảo Phú Quý-Bình Thuận"
                                           },
                                           {
                                               "TongCOD":  11.42,
                                               "CODTuNop":  11.42,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 46 Hùng Vương-Hàm Tân-Bình Thuận"
                                           },
                                           {
                                               "TongCOD":  7.35,
                                               "CODTuNop":  7.35,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 466 Đường 23/10-Nha Trang-Khánh Hòa"
                                           },
                                           {
                                               "TongCOD":  27.57,
                                               "CODTuNop":  27.57,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu cục 47 Trần Hưng Đạo-Hàm Thuận Nam-Bình Thuận"
                                           },
                                           {
                                               "TongCOD":  192.45,
                                               "CODTuNop":  192.45,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 51 Phạm Ngọc Thạch-Ninh Hải-Ninh Thuận"
                                           },
                                           {
                                               "TongCOD":  42.66,
                                               "CODTuNop":  42.66,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 571 Hùng Vương-Cam Nghĩa-TP.Cam Ranh-Khánh Hoà"
                                           },
                                           {
                                               "TongCOD":  16.45,
                                               "CODTuNop":  16.45,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 575 Quốc Lộ 28-Xã Hàm Thuận-Bình Thuận"
                                           },
                                           {
                                               "TongCOD":  4.43,
                                               "CODTuNop":  4.43,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 70 Lê Hồng Phong-Ninh Diêm-Ninh Hoà-Khánh Hoà"
                                           },
                                           {
                                               "TongCOD":  2.22,
                                               "CODTuNop":  2.22,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 73 Nguyễn Tất Thành-Krông Nô-Đắk Nông"
                                           },
                                           {
                                               "TongCOD":  103.88,
                                               "CODTuNop":  103.88,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 95 Thống Nhất-Phan Rang-Ninh Thuận"
                                           },
                                           {
                                               "TongCOD":  13.94,
                                               "CODTuNop":  13.94,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 95A Đội Cấn-Bảo Lộc-Lâm Đồng"
                                           },
                                           {
                                               "TongCOD":  20.81,
                                               "CODTuNop":  20.81,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Đạ Huoai-Lâm Đồng"
                                           },
                                           {
                                               "TongCOD":  12.33,
                                               "CODTuNop":  12.33,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Đắk Mil-Đắk Nông"
                                           },
                                           {
                                               "TongCOD":  28.23,
                                               "CODTuNop":  28.23,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu cục DT707-Hàm Thuận Nam-Bình Thuận"
                                           },
                                           {
                                               "TongCOD":  30.76,
                                               "CODTuNop":  30.76,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Đường 35 Hà Quang 1-Xã Nam Nha Trang-Khánh Hòa"
                                           },
                                           {
                                               "TongCOD":  23.87,
                                               "CODTuNop":  23.87,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Đường Trần Lê-Cát Tiên-Lâm Đồng"
                                           },
                                           {
                                               "TongCOD":  40.24,
                                               "CODTuNop":  40.24,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hàm Liêm-Hàm Thuận Bắc-Bình Thuận"
                                           },
                                           {
                                               "TongCOD":  15.19,
                                               "CODTuNop":  15.19,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hùng Vương-Nghĩa Trung-Gia Nghĩa-Đắk Nông"
                                           },
                                           {
                                               "TongCOD":  93.04,
                                               "CODTuNop":  93.04,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Langbiang-Lạc Dương-Lâm Đồng"
                                           },
                                           {
                                               "TongCOD":  91.29,
                                               "CODTuNop":  91.29,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Lê Hồng Phong-Liên Hiệp-Đức Trọng-Lâm Đồng"
                                           },
                                           {
                                               "TongCOD":  22.34,
                                               "CODTuNop":  22.34,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Nâm N\u0027Jang-Đắk Song-Đắk Nông"
                                           },
                                           {
                                               "TongCOD":  103.15,
                                               "CODTuNop":  103.15,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu cục Ngô Thỳ Sỹ - Đà Lạt - Lâm Đồng"
                                           },
                                           {
                                               "TongCOD":  59.15,
                                               "CODTuNop":  59.15,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Phú Tài-Phan Thiết-Bình Thuận"
                                           },
                                           {
                                               "TongCOD":  17.5,
                                               "CODTuNop":  17.5,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Phú Thủy-Phan Thiết-Bình Thuận"
                                           },
                                           {
                                               "TongCOD":  29.1,
                                               "CODTuNop":  29.1,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Phước Đồng-Nha Trang-Khánh Hoà"
                                           },
                                           {
                                               "TongCOD":  13.07,
                                               "CODTuNop":  13.07,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Phước Minh-Thuận Nam-Khánh Hòa"
                                           },
                                           {
                                               "TongCOD":  5.84,
                                               "CODTuNop":  5.84,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Quảng Khê-Đắk Glong-Đắk Nông"
                                           },
                                           {
                                               "TongCOD":  55.34,
                                               "CODTuNop":  55.34,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Quốc Lộ 14-Đức An-Đắk Song-Đăk Nông"
                                           },
                                           {
                                               "TongCOD":  67.92,
                                               "CODTuNop":  67.92,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Quốc Lộ 1A-Ninh Phước-Ninh Thuận"
                                           },
                                           {
                                               "TongCOD":  35.59,
                                               "CODTuNop":  35.59,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Số 7 Lê Thánh Tôn-LaGi-Bình Thuận"
                                           },
                                           {
                                               "TongCOD":  166.18,
                                               "CODTuNop":  166.18,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu cục Thị Trấn Liên Hương-Tuy Phong-Bình Thuận"
                                           },
                                           {
                                               "TongCOD":  31.78,
                                               "CODTuNop":  31.78,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thôn 2-Tuy Đức-Đắk Nông"
                                           },
                                           {
                                               "TongCOD":  49.09,
                                               "CODTuNop":  49.09,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thôn Phú An Nam 2-Xã Diên Khánh-Khánh Hòa"
                                           },
                                           {
                                               "TongCOD":  10.08,
                                               "CODTuNop":  10.08,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thôn Phúc Hưng-Xã Tân Hà Lâm Hà-Lâm Đồng"
                                           },
                                           {
                                               "TongCOD":  19.17,
                                               "CODTuNop":  19.17,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thôn Xuân Hoà 1-Ninh Hoà-Khánh Hoà"
                                           },
                                           {
                                               "TongCOD":  50.82,
                                               "CODTuNop":  50.82,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Tổ Dân Phố 2-Ninh Hoà-Khánh Hoà"
                                           },
                                           {
                                               "TongCOD":  122.46,
                                               "CODTuNop":  122.46,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Xã Nam Chính-Đức Linh-Bình Thuận"
                                           }
                                       ]
                    },
                    {
                        "TyLeTuNop":  100,
                        "TenVung":  "TBB",
                        "TongCOD":  2276.9,
                        "Trend":  0,
                        "CODTuNop":  2276.9,
                        "BuuCucList":  [
                                           {
                                               "TongCOD":  102.31,
                                               "CODTuNop":  102.31,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 02 Phó Đức Chính-Hồng Hà-Yên Bái"
                                           },
                                           {
                                               "TongCOD":  101.76,
                                               "CODTuNop":  101.76,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 052 Ngọc Uyển-Bắc Hà-Lào Cai"
                                           },
                                           {
                                               "TongCOD":  9.53,
                                               "CODTuNop":  9.53,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 1103 Lê Lợi-Xã Tân Uyên-Lai Châu"
                                           },
                                           {
                                               "TongCOD":  12.44,
                                               "CODTuNop":  12.44,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 112 Võ Nguyên Giáp-Tam Đường-Lai Châu"
                                           },
                                           {
                                               "TongCOD":  28.31,
                                               "CODTuNop":  28.31,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 136 CMT8-Sông Mã-Sơn La"
                                           },
                                           {
                                               "TongCOD":  35.04,
                                               "CODTuNop":  35.04,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 148 Hoàng Quốc Việt-TP.Lào Cai-Lào Cai"
                                           },
                                           {
                                               "TongCOD":  174.13,
                                               "CODTuNop":  174.13,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 155 Điện Biên Phủ-Sa Pa-Lào Cai"
                                           },
                                           {
                                               "TongCOD":  2.93,
                                               "CODTuNop":  2.93,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 244 Nguyễn Tất Thành-Yên Thế-Lục Yên-Yên Bái"
                                           },
                                           {
                                               "TongCOD":  85.87,
                                               "CODTuNop":  85.87,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 266 Đường 8 Tháng 5-Xã Yên Châu-Sơn La"
                                           },
                                           {
                                               "TongCOD":  35.1,
                                               "CODTuNop":  35.1,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 341 Nguyễn Tất Thành-Bảo Yên-Lào Cai"
                                           },
                                           {
                                               "TongCOD":  136.89,
                                               "CODTuNop":  136.89,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 354 Trần Đăng Ninh-Sơn La-Sơn La"
                                           },
                                           {
                                               "TongCOD":  42.61,
                                               "CODTuNop":  42.61,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 36 Đường Tô Hiệu-Xã Mường La-Sơn La"
                                           },
                                           {
                                               "TongCOD":  0.48,
                                               "CODTuNop":  0.48,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 563 Tiểu Khu 3-Xã Bắc Yên-Sơn La"
                                           },
                                           {
                                               "TongCOD":  117.39,
                                               "CODTuNop":  117.39,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 650 Đinh Tiên Hoàng-Yên Thịnh-Yên Bái"
                                           },
                                           {
                                               "TongCOD":  20.28,
                                               "CODTuNop":  20.28,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 686 Lê Duẩn-Chiềng Sinh-Sơn La"
                                           },
                                           {
                                               "TongCOD":  224.62,
                                               "CODTuNop":  224.62,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 77 Hoàng Quốc Việt-Mộc Châu-Sơn La"
                                           },
                                           {
                                               "TongCOD":  36.52,
                                               "CODTuNop":  36.52,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 83 Điện Biên Phủ-Xã Bum Tở-Lai Châu"
                                           },
                                           {
                                               "TongCOD":  5.29,
                                               "CODTuNop":  5.29,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Bản Liên Hà 3-Xã Bảo Hà-Lào Cai"
                                           },
                                           {
                                               "TongCOD":  87.9,
                                               "CODTuNop":  87.9,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Cầu Thia-Nghĩa Lộ-Yên Bái"
                                           },
                                           {
                                               "TongCOD":  34.8,
                                               "CODTuNop":  34.8,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Duyên Hải-Duyên Hải-Lào Cai"
                                           },
                                           {
                                               "TongCOD":  52.22,
                                               "CODTuNop":  52.22,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hồng Hà-Mậu A-Văn Yên-Yên Bái"
                                           },
                                           {
                                               "TongCOD":  14.75,
                                               "CODTuNop":  14.75,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Khánh Yên-Văn Bàn-Lào Cai"
                                           },
                                           {
                                               "TongCOD":  42.17,
                                               "CODTuNop":  42.17,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Ngã Ba Khe-Văn Chấn-Yên Bái"
                                           },
                                           {
                                               "TongCOD":  100.61,
                                               "CODTuNop":  100.61,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Phố Lê Thanh-TP Lào Cai-Lào Cai"
                                           },
                                           {
                                               "TongCOD":  6.84,
                                               "CODTuNop":  6.84,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục QL 6-Xã Vân Hồ-Sơn La"
                                           },
                                           {
                                               "TongCOD":  111.31,
                                               "CODTuNop":  111.31,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Tân Thanh-Điện Biên Phủ-Điện Biên"
                                           },
                                           {
                                               "TongCOD":  9.54,
                                               "CODTuNop":  9.54,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục TDP 1-Mường Chà-Điện Biên"
                                           },
                                           {
                                               "TongCOD":  65.01,
                                               "CODTuNop":  65.01,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thị Trấn Tủa Chùa-Tủa Chùa-Điện Biên"
                                           },
                                           {
                                               "TongCOD":  18.23,
                                               "CODTuNop":  18.23,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thị trấn Tuần Giáo-Điện Biên"
                                           },
                                           {
                                               "TongCOD":  89.8,
                                               "CODTuNop":  89.8,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thôn 24-xã Noong Hẹt-Điện Biên"
                                           },
                                           {
                                               "TongCOD":  6.04,
                                               "CODTuNop":  6.04,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thôn Mã Tuyển 3-Mường Khương-Lào Cai"
                                           },
                                           {
                                               "TongCOD":  98,
                                               "CODTuNop":  98,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thôn Tả Hà 3-Bảo Thắng-Lào Cai"
                                           },
                                           {
                                               "TongCOD":  3.31,
                                               "CODTuNop":  3.31,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thôn Tây Sơn-Mường So-Phong Thổ-Lai Châu"
                                           },
                                           {
                                               "TongCOD":  117.67,
                                               "CODTuNop":  117.67,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Tiểu Khu 10-Hát Lót-Mai Sơn-Sơn La"
                                           },
                                           {
                                               "TongCOD":  77.93,
                                               "CODTuNop":  77.93,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Tiểu Khu 1-Thuận Châu-Sơn La"
                                           },
                                           {
                                               "TongCOD":  65.59,
                                               "CODTuNop":  65.59,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Tiểu khu 6-Phù Yên-Sơn La"
                                           },
                                           {
                                               "TongCOD":  9.38,
                                               "CODTuNop":  9.38,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Tổ 4 Điện Biên Đông-Xã Na Son-Điện Biên"
                                           },
                                           {
                                               "TongCOD":  7.79,
                                               "CODTuNop":  7.79,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Tổ 7-Bát Xát-Lào Cai"
                                           },
                                           {
                                               "TongCOD":  14.74,
                                               "CODTuNop":  14.74,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục TT Mường Ảng-Xã Mường Ảng-Điện Biên"
                                           },
                                           {
                                               "TongCOD":  8.19,
                                               "CODTuNop":  8.19,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Vận Hành Bản Hua Mường-Xã Xốp Cộp-Sơn La"
                                           },
                                           {
                                               "TongCOD":  10.53,
                                               "CODTuNop":  10.53,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Vận Hành TDP Hoàng Thu Phố-Xã Si Ma Cai-Lào Cai"
                                           },
                                           {
                                               "TongCOD":  11.08,
                                               "CODTuNop":  11.08,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Vận Hành Tổ 1 TT Cổ Phúc-Xã Trấn Yên-Yên Bái"
                                           },
                                           {
                                               "TongCOD":  14.95,
                                               "CODTuNop":  14.95,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Xã Mường Giàng-Quỳnh Nhai-Sơn La"
                                           },
                                           {
                                               "TongCOD":  27.05,
                                               "CODTuNop":  27.05,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Điểm xử lý hàng Nậm Pồ-Điện Biên"
                                           }
                                       ]
                    },
                    {
                        "TyLeTuNop":  100,
                        "TenVung":  "TNB",
                        "TongCOD":  5891.41,
                        "Trend":  0,
                        "CODTuNop":  5891.41,
                        "BuuCucList":  [
                                           {
                                               "TongCOD":  62.13,
                                               "CODTuNop":  62.13,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 02 Lâm Văn Quang-Hà Tiên-Kiên Giang"
                                           },
                                           {
                                               "TongCOD":  12.82,
                                               "CODTuNop":  12.82,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 03 Lê Lợi-An Châu-An Giang"
                                           },
                                           {
                                               "TongCOD":  56.01,
                                               "CODTuNop":  56.01,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 06 Nguyễn Tri Phương-Vị Thuỷ-Hậu Giang"
                                           },
                                           {
                                               "TongCOD":  9.81,
                                               "CODTuNop":  9.81,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 137 Đường 30/04-Trần Văn Thời-Cà Mau"
                                           },
                                           {
                                               "TongCOD":  89.32,
                                               "CODTuNop":  89.32,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 154B Mai Thị Hồng Hạnh-Rạch Giá-Kiên Giang"
                                           },
                                           {
                                               "TongCOD":  269.12,
                                               "CODTuNop":  269.12,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 16 Đỗ Nhuận-Long Xuyên-An Giang"
                                           },
                                           {
                                               "TongCOD":  135.31,
                                               "CODTuNop":  135.31,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 162 Trần Văn Trà-Cái Răng-Cần Thơ"
                                           },
                                           {
                                               "TongCOD":  21.22,
                                               "CODTuNop":  21.22,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 20 Lộ Liên Hương-Rạch Giá-Kiên Giang"
                                           },
                                           {
                                               "TongCOD":  68.8,
                                               "CODTuNop":  68.8,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 205 Nguyễn Trãi-Thị Xã Ngã Năm-Sóc Trăng"
                                           },
                                           {
                                               "TongCOD":  4.77,
                                               "CODTuNop":  4.77,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 211 Quốc Lộ 80-Kiên Lương-Kiên Giang"
                                           },
                                           {
                                               "TongCOD":  33.68,
                                               "CODTuNop":  33.68,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 266B Hàng Gòn-Lê Bình-Cái Răng-Cần Thơ"
                                           },
                                           {
                                               "TongCOD":  49.06,
                                               "CODTuNop":  49.06,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 28 Giá Rai-Xã Long Điền-Bạc Liêu"
                                           },
                                           {
                                               "TongCOD":  48.24,
                                               "CODTuNop":  48.24,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 28 Xuân Hồng-Bình Thuỷ-Cần Thơ"
                                           },
                                           {
                                               "TongCOD":  30.95,
                                               "CODTuNop":  30.95,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 2C Châu Văn Liêm-Ngã Bảy-Hậu Giang"
                                           },
                                           {
                                               "TongCOD":  57.15,
                                               "CODTuNop":  57.15,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 320/12 Ngô Quyền-Cà Mau-Cà Mau"
                                           },
                                           {
                                               "TongCOD":  181.69,
                                               "CODTuNop":  181.69,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 32B Huyện lộ 6-TT.Kế Sách-Sóc Trăng"
                                           },
                                           {
                                               "TongCOD":  165.85,
                                               "CODTuNop":  165.85,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 33 Ấp Chợ-Xã Trần Đề-Sóc Trăng"
                                           },
                                           {
                                               "TongCOD":  42.8,
                                               "CODTuNop":  42.8,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 379 Hùng Vương-Châu Thành-Hậu Giang"
                                           },
                                           {
                                               "TongCOD":  104.18,
                                               "CODTuNop":  104.18,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 39 Tỉnh Lộ 940-Huỳnh Hữu Nghĩa-Mỹ Tú-Sóc Trăng"
                                           },
                                           {
                                               "TongCOD":  87.83,
                                               "CODTuNop":  87.83,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 535 Tôn Đức Thắng-Sóc Trăng-Sóc Trăng"
                                           },
                                           {
                                               "TongCOD":  75.9,
                                               "CODTuNop":  75.9,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 54 Hùng Vương-Năm Căn-Cà Mau"
                                           },
                                           {
                                               "TongCOD":  152.89,
                                               "CODTuNop":  152.89,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 555 Nguyễn Văn Cừ-Phú Quốc-Kiên Giang"
                                           },
                                           {
                                               "TongCOD":  187.24,
                                               "CODTuNop":  187.24,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 56 Nguyễn Văn Cừ-Cái Khế-Ninh Kiều-Cần Thơ"
                                           },
                                           {
                                               "TongCOD":  20.48,
                                               "CODTuNop":  20.48,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 56 Quốc Lộ 1A-Xã Vĩnh Lợi-Bạc Liêu"
                                           },
                                           {
                                               "TongCOD":  7.05,
                                               "CODTuNop":  7.05,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 578 QL1A-Xã Hòa Bình-Bạc Liêu"
                                           },
                                           {
                                               "TongCOD":  110.44,
                                               "CODTuNop":  110.44,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 599 Lê Hồng Phong-Phú Lợi-Sóc Trăng"
                                           },
                                           {
                                               "TongCOD":  95.5,
                                               "CODTuNop":  95.5,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 608 Trần Hưng Đạo-Bình Đức-TP.Long Xuyên-An Giang"
                                           },
                                           {
                                               "TongCOD":  103.03,
                                               "CODTuNop":  103.03,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 617 Tỉnh Lộ 942-Xã Chợ Mới-An Giang"
                                           },
                                           {
                                               "TongCOD":  30.6,
                                               "CODTuNop":  30.6,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 64 Đường 30/04 Vĩnh Châu-Sóc Trăng"
                                           },
                                           {
                                               "TongCOD":  157.06,
                                               "CODTuNop":  157.06,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 66 Nguyễn Tri Phương-Châu Phú B-TP.Châu Đốc-An Giang"
                                           },
                                           {
                                               "TongCOD":  317.93,
                                               "CODTuNop":  317.93,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 68 Nguyễn Văn Cừ-An Khánh-Ninh Kiều-Cần Thơ"
                                           },
                                           {
                                               "TongCOD":  27.05,
                                               "CODTuNop":  27.05,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 68 Quốc Lộ 1A-Xã Lương Thế Trân-Cà Mau"
                                           },
                                           {
                                               "TongCOD":  110.69,
                                               "CODTuNop":  110.69,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 729 Đông An-Xã Tân Hiệp-Kiên Giang"
                                           },
                                           {
                                               "TongCOD":  4.04,
                                               "CODTuNop":  4.04,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 79 Quốc Lộ 1A-Giá Rai-Bạc Liêu"
                                           },
                                           {
                                               "TongCOD":  10.4,
                                               "CODTuNop":  10.4,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 8 Huỳnh Mẫn Đạt-Giồng Riềng-Kiên Giang"
                                           },
                                           {
                                               "TongCOD":  18.65,
                                               "CODTuNop":  18.65,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 86 Đường 932-Phú Tân-Châu Thành-Sóc Trăng"
                                           },
                                           {
                                               "TongCOD":  122.87,
                                               "CODTuNop":  122.87,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 86 Thống nhất II-Ngan Dừa-Hồng Dân-Bạc Liêu"
                                           },
                                           {
                                               "TongCOD":  31.42,
                                               "CODTuNop":  31.42,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 88 Quốc Lộ 61C-Xã Tân Hòa-Hậu Giang"
                                           },
                                           {
                                               "TongCOD":  43.57,
                                               "CODTuNop":  43.57,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 931 Đường 30/4-Long Mỹ-Hậu Giang"
                                           },
                                           {
                                               "TongCOD":  61.92,
                                               "CODTuNop":  61.92,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Ấp 1 QL61-Châu Thành A-Hậu Giang"
                                           },
                                           {
                                               "TongCOD":  22.83,
                                               "CODTuNop":  22.83,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Ấp 3 -Thạnh Trị-Sóc Trăng"
                                           },
                                           {
                                               "TongCOD":  162.4,
                                               "CODTuNop":  162.4,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Ấp 4-Thới Bình-Thới Bình-Cà Mau"
                                           },
                                           {
                                               "TongCOD":  10.82,
                                               "CODTuNop":  10.82,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Ấp Chợ-Cù Lao Dung-Sóc Trăng"
                                           },
                                           {
                                               "TongCOD":  35.93,
                                               "CODTuNop":  35.93,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Ấp Khúc Tréo B-Tân Phong-Giá Rai-Bạc Liêu"
                                           },
                                           {
                                               "TongCOD":  2.26,
                                               "CODTuNop":  2.26,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Ấp Tân Long A-Xã Tân Tiến-Cà Mau"
                                           },
                                           {
                                               "TongCOD":  31.62,
                                               "CODTuNop":  31.62,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Ấp Thị Tứ-Thị Trấn Phong Điền-Cần Thơ"
                                           },
                                           {
                                               "TongCOD":  36.86,
                                               "CODTuNop":  36.86,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Cái Tắc-Phú Mỹ-Phú Tân-An Giang"
                                           },
                                           {
                                               "TongCOD":  11.45,
                                               "CODTuNop":  11.45,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Châu Văn Liêm-Ô Môn-Cần Thơ"
                                           },
                                           {
                                               "TongCOD":  15.81,
                                               "CODTuNop":  15.81,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục DT933 Ấp Đồng Tràm-Xã Long Thạnh-Kiên Giang"
                                           },
                                           {
                                               "TongCOD":  161.09,
                                               "CODTuNop":  161.09,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu cục Đường 23 Tháng 8-TP.Bạc Liêu-Bạc Liêu"
                                           },
                                           {
                                               "TongCOD":  132.71,
                                               "CODTuNop":  132.71,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Đường Huyện 21-Thới Lai-Cần Thơ"
                                           },
                                           {
                                               "TongCOD":  1.93,
                                               "CODTuNop":  1.93,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Đường Phan Ngọc Hiển-Xã Cái Đôi Vàm-Cà Mau"
                                           },
                                           {
                                               "TongCOD":  140.21,
                                               "CODTuNop":  140.21,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Dương Thị Cẩm Vân-Khóm 4-Đầm Dơi-Cà Mau"
                                           },
                                           {
                                               "TongCOD":  388.87,
                                               "CODTuNop":  388.87,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Đường Tuyến Tránh-Phú Quốc-Kiên Giang"
                                           },
                                           {
                                               "TongCOD":  151.44,
                                               "CODTuNop":  151.44,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hồng Khanh-Phước Long-Bạc Liêu"
                                           },
                                           {
                                               "TongCOD":  14.62,
                                               "CODTuNop":  14.62,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục KDC Định An Quốc Lộ 61-Xã Gò Quao-Kiên Giang"
                                           },
                                           {
                                               "TongCOD":  144.93,
                                               "CODTuNop":  144.93,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Lô B2 TTTM Vĩnh Thạnh-Vĩnh Thạnh-Cần Thơ"
                                           },
                                           {
                                               "TongCOD":  22.3,
                                               "CODTuNop":  22.3,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Lô C02-03 Lê Hồng Phong-An Hòa-Rạch Giá-Kiên Giang"
                                           },
                                           {
                                               "TongCOD":  128.02,
                                               "CODTuNop":  128.02,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Lý Thường Kiệt-Tân Thành-Cà Mau"
                                           },
                                           {
                                               "TongCOD":  25.82,
                                               "CODTuNop":  25.82,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Lý Văn Lâm-TP.Cà Mau-Cà Mau"
                                           },
                                           {
                                               "TongCOD":  158.87,
                                               "CODTuNop":  158.87,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Ngã 3 Bãy Vòng 973-Đặc Khu Phú Quốc-Kiên Giang"
                                           },
                                           {
                                               "TongCOD":  44.17,
                                               "CODTuNop":  44.17,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Nguyễn An Ninh-Vị Thanh-Hậu Giang"
                                           },
                                           {
                                               "TongCOD":  42.34,
                                               "CODTuNop":  42.34,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Nguyễn Trung Thành Khóm 1-U Minh-Cà Mau"
                                           },
                                           {
                                               "TongCOD":  24.47,
                                               "CODTuNop":  24.47,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Nguyễn Văn Cưng-Rạch Gốc-Ngọc Hiển-Cà Mau"
                                           },
                                           {
                                               "TongCOD":  91.44,
                                               "CODTuNop":  91.44,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Phường Trung Kiên-Thốt Nốt-Cần Thơ"
                                           },
                                           {
                                               "TongCOD":  55.55,
                                               "CODTuNop":  55.55,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục QL63 TT Thứ 3-Xã An Biên-Kiên Giang"
                                           },
                                           {
                                               "TongCOD":  89.23,
                                               "CODTuNop":  89.23,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục QL91-Châu Phú-An Giang"
                                           },
                                           {
                                               "TongCOD":  31.82,
                                               "CODTuNop":  31.82,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Quốc Lộ 80-Xã Hòn Đất-Kiên Giang"
                                           },
                                           {
                                               "TongCOD":  30.46,
                                               "CODTuNop":  30.46,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Tân Châu-An Giang"
                                           },
                                           {
                                               "TongCOD":  4.73,
                                               "CODTuNop":  4.73,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Tân Hưng Đông-Cái Nước-Cà Mau"
                                           },
                                           {
                                               "TongCOD":  18.38,
                                               "CODTuNop":  18.38,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thạnh An-An Minh-Kiên Giang"
                                           },
                                           {
                                               "TongCOD":  18.02,
                                               "CODTuNop":  18.02,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thạnh Yên-U Minh Thượng-Kiên Giang"
                                           },
                                           {
                                               "TongCOD":  33.6,
                                               "CODTuNop":  33.6,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thị Trấn Cây Dương-Phụng Hiệp-Hậu Giang"
                                           },
                                           {
                                               "TongCOD":  48.43,
                                               "CODTuNop":  48.43,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thị trấn Sông Đốc-Trần Văn Thời-Cà Mau"
                                           },
                                           {
                                               "TongCOD":  152.24,
                                               "CODTuNop":  152.24,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Tỉnh lộ 919-Thạnh Phú-Cờ Đỏ-Cần Thơ"
                                           },
                                           {
                                               "TongCOD":  25.56,
                                               "CODTuNop":  25.56,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Tỉnh Lộ 963-Xã Hòa Thuận-Kiên Giang"
                                           },
                                           {
                                               "TongCOD":  60.33,
                                               "CODTuNop":  60.33,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Tri Tôn-An Giang"
                                           },
                                           {
                                               "TongCOD":  19.09,
                                               "CODTuNop":  19.09,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Vĩnh Phước QL63-Vĩnh Thuận-Kiên Giang"
                                           },
                                           {
                                               "TongCOD":  81.26,
                                               "CODTuNop":  81.26,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Xã Cửa Cạn-Phú Quốc-Kiên Giang"
                                           }
                                       ]
                    },
                    {
                        "TyLeTuNop":  100,
                        "TenVung":  "TNG",
                        "TongCOD":  3151.74,
                        "Trend":  0,
                        "CODTuNop":  3151.74,
                        "BuuCucList":  [
                                           {
                                               "TongCOD":  43.52,
                                               "CODTuNop":  43.52,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 03 Trần Hưng Đạo-Krông Pắc-Đắk Lắk"
                                           },
                                           {
                                               "TongCOD":  42.43,
                                               "CODTuNop":  42.43,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 09 Quốc Lộ 14-Hoà Phú-Đắk Lắk"
                                           },
                                           {
                                               "TongCOD":  13.57,
                                               "CODTuNop":  13.57,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 127B Lê Văn Tám-Thị Trấn Kbang-Huyện Kbang-Gia Lai"
                                           },
                                           {
                                               "TongCOD":  29.37,
                                               "CODTuNop":  29.37,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 132 Hoàng Hoa Thám-Buôn Ma Thuột-Đắk Lắk"
                                           },
                                           {
                                               "TongCOD":  40.69,
                                               "CODTuNop":  40.69,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 138 Trần Phú- Đồng Xuân- Phú Yên"
                                           },
                                           {
                                               "TongCOD":  245.8,
                                               "CODTuNop":  245.8,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 15 Hoàng Văn Thụ-An Khê-Gia Lai"
                                           },
                                           {
                                               "TongCOD":  41.56,
                                               "CODTuNop":  41.56,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 153 Hùng Vương-Ia Grai-Gia Lai"
                                           },
                                           {
                                               "TongCOD":  32.88,
                                               "CODTuNop":  32.88,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 159 Tân Quảng-EaTóh-Krông Năng-Đắk Lắk"
                                           },
                                           {
                                               "TongCOD":  98.38,
                                               "CODTuNop":  98.38,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 172 Hùng Vương-Chư Prông-Gia Lai"
                                           },
                                           {
                                               "TongCOD":  24.87,
                                               "CODTuNop":  24.87,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 172 Quang Trung-Đức Cơ-Gia Lai"
                                           },
                                           {
                                               "TongCOD":  45.18,
                                               "CODTuNop":  45.18,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 204 Trần Hưng Đạo-Buôn Hồ-Đắk Lắk"
                                           },
                                           {
                                               "TongCOD":  77.09,
                                               "CODTuNop":  77.09,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu cục 236A Nguyễn Huệ-Đak Đoa-Gia Lai"
                                           },
                                           {
                                               "TongCOD":  162.23,
                                               "CODTuNop":  162.23,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 263 Đống Đa-Quy Nhơn-Bình Định"
                                           },
                                           {
                                               "TongCOD":  110.68,
                                               "CODTuNop":  110.68,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 269 Tây Sơn-Quy Nhơn-Bình Định"
                                           },
                                           {
                                               "TongCOD":  53.2,
                                               "CODTuNop":  53.2,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu cục 27 Nguyễn Tất Thành-Tuy Hòa-Phú Yên"
                                           },
                                           {
                                               "TongCOD":  24.19,
                                               "CODTuNop":  24.19,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 271 Nguyễn Văn Cừ-Tân Lập-Đắk Lắk"
                                           },
                                           {
                                               "TongCOD":  28.46,
                                               "CODTuNop":  28.46,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 274 Hùng Vương-Chư Pưh-Gia Lai"
                                           },
                                           {
                                               "TongCOD":  1.5,
                                               "CODTuNop":  1.5,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 274 Hùng Vương-Tuy Hòa-Phú Yên"
                                           },
                                           {
                                               "TongCOD":  86.9,
                                               "CODTuNop":  86.9,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 297 Hùng Vương-CưMgar-Đắk Lắk"
                                           },
                                           {
                                               "TongCOD":  92.47,
                                               "CODTuNop":  92.47,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu cục 327A Cách Mạng Tháng 8-Pleiku-Gia Lai"
                                           },
                                           {
                                               "TongCOD":  74.27,
                                               "CODTuNop":  74.27,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 34 Quang Trung-Eakar-Đắk Lắk"
                                           },
                                           {
                                               "TongCOD":  21.2,
                                               "CODTuNop":  21.2,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 39 Nguyễn Huệ-Ayun Pa-Gia Lai"
                                           },
                                           {
                                               "TongCOD":  33.7,
                                               "CODTuNop":  33.7,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 396 Âu Cơ-Bùi Thị Xuân-Quy Nhơn-Bình Định"
                                           },
                                           {
                                               "TongCOD":  11.06,
                                               "CODTuNop":  11.06,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 414 Trần Hưng Đạo-Mang Yang-Gia Lai"
                                           },
                                           {
                                               "TongCOD":  36.99,
                                               "CODTuNop":  36.99,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 44 Quang Trung-Tây Sơn-Bình Định"
                                           },
                                           {
                                               "TongCOD":  111.42,
                                               "CODTuNop":  111.42,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu cục 499 Phạm Văn Đồng-Pleiku-Gia Lai"
                                           },
                                           {
                                               "TongCOD":  40.58,
                                               "CODTuNop":  40.58,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 56 Đường 17 Tháng 3-Chư Sê-Gia Lai"
                                           },
                                           {
                                               "TongCOD":  54.84,
                                               "CODTuNop":  54.84,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 565 Trần Phú-Tuy Phước-Bình Định"
                                           },
                                           {
                                               "TongCOD":  120.43,
                                               "CODTuNop":  120.43,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 569 Trường Chinh-Pleiku-Gia Lai"
                                           },
                                           {
                                               "TongCOD":  9.09,
                                               "CODTuNop":  9.09,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 61 Quang Trung-Xã Vân Canh-Bình Định"
                                           },
                                           {
                                               "TongCOD":  96.59,
                                               "CODTuNop":  96.59,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 64 Hùng Vương-Thị Trấn Krông Năng-Đắk Lắk"
                                           },
                                           {
                                               "TongCOD":  90.97,
                                               "CODTuNop":  90.97,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 70 Đường 30/4-Ea Tam-Buôn Ma Thuột-Đắk Lắk"
                                           },
                                           {
                                               "TongCOD":  49.22,
                                               "CODTuNop":  49.22,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 71 Quốc Lộ 1A-Xã Tuy An Bắc-Phú Yên"
                                           },
                                           {
                                               "TongCOD":  65.71,
                                               "CODTuNop":  65.71,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu cục 73 Lê Lợi-Chư Pah-Gia Lai"
                                           },
                                           {
                                               "TongCOD":  31.94,
                                               "CODTuNop":  31.94,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 767 Hùng Vương-Quy Nhơn-Bình Định"
                                           },
                                           {
                                               "TongCOD":  67.2,
                                               "CODTuNop":  67.2,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 838 Quang Trung-Phường Hoài Nhơn Tây-Bình Định"
                                           },
                                           {
                                               "TongCOD":  82.33,
                                               "CODTuNop":  82.33,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 941 Giải Phóng-Ea Đrăng-Ea H\u0027Leo-Đắk Lắk"
                                           },
                                           {
                                               "TongCOD":  94.52,
                                               "CODTuNop":  94.52,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu cục DT644-Sông Cầu-Phú Yên"
                                           },
                                           {
                                               "TongCOD":  52.26,
                                               "CODTuNop":  52.26,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Đường 3 Tháng 2-Xã Phù Cát-Bình Định"
                                           },
                                           {
                                               "TongCOD":  28.45,
                                               "CODTuNop":  28.45,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Đường Tỉnh 643-Xã Tuy An Nam-Phú Yên"
                                           },
                                           {
                                               "TongCOD":  13.71,
                                               "CODTuNop":  13.71,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Ea Phê-Krông Pắk-Đắk Lắk"
                                           },
                                           {
                                               "TongCOD":  105.22,
                                               "CODTuNop":  105.22,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hoài Nhơn-Bình Định"
                                           },
                                           {
                                               "TongCOD":  107.96,
                                               "CODTuNop":  107.96,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Khu Vực Kim Châu-An Nhơn-Bình Định"
                                           },
                                           {
                                               "TongCOD":  57.59,
                                               "CODTuNop":  57.59,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Km 30 Tân Hòa-Buôn Đôn-Đắk Lắk"
                                           },
                                           {
                                               "TongCOD":  49.92,
                                               "CODTuNop":  49.92,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Km18-Cư Kuin-Đắk Lắk"
                                           },
                                           {
                                               "TongCOD":  51.52,
                                               "CODTuNop":  51.52,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục KM92 QL14-Xã Ea Drăng-Đắk Lắk"
                                           },
                                           {
                                               "TongCOD":  57.67,
                                               "CODTuNop":  57.67,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục KP3 TT Hai Riềng-Sông Hinh-Phú Yên"
                                           },
                                           {
                                               "TongCOD":  2.83,
                                               "CODTuNop":  2.83,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục KP4 Hòa Vinh- Đông Hòa- Phú Yên"
                                           },
                                           {
                                               "TongCOD":  36.44,
                                               "CODTuNop":  36.44,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Ngã Ba Đắk Liêng-Huyện Lắk-Đắk Lắk"
                                           },
                                           {
                                               "TongCOD":  52.92,
                                               "CODTuNop":  52.92,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Phạm Văn Đồng-Pơng Đrang-Krông Búk-Đắk Lắk"
                                           },
                                           {
                                               "TongCOD":  17.56,
                                               "CODTuNop":  17.56,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Số 35 Phạm Văn Đồng-Tăng Bạt Hổ-Hoài Ân-Bình Định"
                                           },
                                           {
                                               "TongCOD":  34.21,
                                               "CODTuNop":  34.21,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục TDP3 Thị Trấn Krông Kmart-Krông Bông-Đắk Lắk"
                                           },
                                           {
                                               "TongCOD":  30.99,
                                               "CODTuNop":  30.99,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thị Trấn Buôn Trấp-Krông A Na-Đắk Lắk"
                                           },
                                           {
                                               "TongCOD":  10.07,
                                               "CODTuNop":  10.07,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thị Trấn Phú Thứ-Tây Hoà-Phú Yên"
                                           },
                                           {
                                               "TongCOD":  28.88,
                                               "CODTuNop":  28.88,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thôn Chánh Hóa-Xã Cát Tiến-Bình Định"
                                           },
                                           {
                                               "TongCOD":  19.98,
                                               "CODTuNop":  19.98,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thôn Đông Phước-Phú Hòa-Phú Yên"
                                           },
                                           {
                                               "TongCOD":  6.54,
                                               "CODTuNop":  6.54,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Vận Hành Thôn Định Trị-Xã Vĩnh Thạnh-Bình Định"
                                           }
                                       ]
                    },
                    {
                        "TyLeTuNop":  100,
                        "TenVung":  "TNT",
                        "TongCOD":  5786.76,
                        "Trend":  0,
                        "CODTuNop":  5786.76,
                        "BuuCucList":  [
                                           {
                                               "TongCOD":  107.4,
                                               "CODTuNop":  107.4,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 104 Hòa Bình-Phường Tân Hòa-Hòa Bình"
                                           },
                                           {
                                               "TongCOD":  238.88,
                                               "CODTuNop":  238.88,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 1165 Đại Lộ Hùng Vương-Thanh Miếu-Phú Thọ"
                                           },
                                           {
                                               "TongCOD":  61.05,
                                               "CODTuNop":  61.05,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 132 Tân Long-Yên Lập-Phú Thọ"
                                           },
                                           {
                                               "TongCOD":  46.04,
                                               "CODTuNop":  46.04,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 150 Trần Hưng Đạo-Bình Lục-Hà Nam"
                                           },
                                           {
                                               "TongCOD":  42.38,
                                               "CODTuNop":  42.38,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 214 Tuệ Tĩnh-Ninh Tiến-Ninh Bình"
                                           },
                                           {
                                               "TongCOD":  79.48,
                                               "CODTuNop":  79.48,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 26 Phạm Thận Duật-Bích Đào-Ninh Bình"
                                           },
                                           {
                                               "TongCOD":  83.46,
                                               "CODTuNop":  83.46,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 362 Quốc Lộ 6-Xã Lương Sơn-Hòa Bình"
                                           },
                                           {
                                               "TongCOD":  9.79,
                                               "CODTuNop":  9.79,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 38 Lê Quý Đôn-Phường Duy Tiên-Hà Nam"
                                           },
                                           {
                                               "TongCOD":  114.95,
                                               "CODTuNop":  114.95,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu cục 57 Nguyễn Văn Nhượng-Vĩnh Tường-Vĩnh Phúc"
                                           },
                                           {
                                               "TongCOD":  62.92,
                                               "CODTuNop":  62.92,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 64 Trần Hưng Đạo-TP.Ninh Bình-Ninh Bình"
                                           },
                                           {
                                               "TongCOD":  15.33,
                                               "CODTuNop":  15.33,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 77 KĐT Yết Kiều-Phủ Lý-Hà Nam"
                                           },
                                           {
                                               "TongCOD":  181.98,
                                               "CODTuNop":  181.98,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 85 Nguyễn Tất Thành-Liên Bảo-Vĩnh Yên-Vĩnh Phúc"
                                           },
                                           {
                                               "TongCOD":  116.26,
                                               "CODTuNop":  116.26,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Biên Hòa-Kim Thanh-Hà Nam"
                                           },
                                           {
                                               "TongCOD":  13.85,
                                               "CODTuNop":  13.85,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Điềm Khê-Gia Viễn-Ninh Bình"
                                           },
                                           {
                                               "TongCOD":  8.38,
                                               "CODTuNop":  8.38,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Đinh Đức Thiện-Xã Vạn Thắng-Nam Định"
                                           },
                                           {
                                               "TongCOD":  160.85,
                                               "CODTuNop":  160.85,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Đò Quan-Cửa Nam-Nam Định"
                                           },
                                           {
                                               "TongCOD":  119.48,
                                               "CODTuNop":  119.48,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Đồng Xuân-Thanh-Ba Phú Thọ"
                                           },
                                           {
                                               "TongCOD":  44.39,
                                               "CODTuNop":  44.39,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Đường Đông Trung-Xã Tu Vũ-Phú Thọ"
                                           },
                                           {
                                               "TongCOD":  64.48,
                                               "CODTuNop":  64.48,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Đường Lạc Long Quân-Xã Hạ Hòa-Phú Thọ"
                                           },
                                           {
                                               "TongCOD":  135.68,
                                               "CODTuNop":  135.68,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Đường Trần Tự Khánh-Mỹ Lộc-Nam Định"
                                           },
                                           {
                                               "TongCOD":  80.32,
                                               "CODTuNop":  80.32,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Gia Viễn-Ninh Bình"
                                           },
                                           {
                                               "TongCOD":  118.21,
                                               "CODTuNop":  118.21,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hợp Hải-Lâm Thao-Phú Thọ"
                                           },
                                           {
                                               "TongCOD":  26.63,
                                               "CODTuNop":  26.63,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hương Nộn-Tam Nông-Phú Thọ"
                                           },
                                           {
                                               "TongCOD":  52.98,
                                               "CODTuNop":  52.98,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục KĐT Phúc Sơn-Thượng Trưng-Vĩnh Tường-Vĩnh Phúc"
                                           },
                                           {
                                               "TongCOD":  68.18,
                                               "CODTuNop":  68.18,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Khu 1 Thị trấn Chi Nê-Lạc Thủy-Hòa Bình"
                                           },
                                           {
                                               "TongCOD":  32.57,
                                               "CODTuNop":  32.57,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Khu 10 Tân Phú-Tân Sơn-Phú Thọ"
                                           },
                                           {
                                               "TongCOD":  99.65,
                                               "CODTuNop":  99.65,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Khu 10 Trường Chinh-Kim Đức-Việt Trì-Phú Thọ"
                                           },
                                           {
                                               "TongCOD":  89.15,
                                               "CODTuNop":  89.15,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Khu 1-Thanh Minh-Phú Thọ"
                                           },
                                           {
                                               "TongCOD":  160.83,
                                               "CODTuNop":  160.83,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Khu 3-Hải Hậu-Nam Định"
                                           },
                                           {
                                               "TongCOD":  145.52,
                                               "CODTuNop":  145.52,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu cục Khu đô thị Quán Chiền-Nam Dương-Nam Trực-Nam Định"
                                           },
                                           {
                                               "TongCOD":  115.11,
                                               "CODTuNop":  115.11,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Khu Đoàn Kết Đường 19/5-Xã Đoan Hùng-Phú Thọ"
                                           },
                                           {
                                               "TongCOD":  24.79,
                                               "CODTuNop":  24.79,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Khu Đồng Cổng Huyện-Tam Dương-Vĩnh Phúc"
                                           },
                                           {
                                               "TongCOD":  101,
                                               "CODTuNop":  101,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Lạc Thiện-Quang Sơn-Ninh Bình"
                                           },
                                           {
                                               "TongCOD":  22.29,
                                               "CODTuNop":  22.29,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Lê Đức Toản-Hùng Vương-Phúc Yên-Vĩnh Phúc"
                                           },
                                           {
                                               "TongCOD":  83.99,
                                               "CODTuNop":  83.99,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Lê Hồng Phong-Đồng Tâm-Vĩnh Yên-Vĩnh Phúc"
                                           },
                                           {
                                               "TongCOD":  39.78,
                                               "CODTuNop":  39.78,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Ngô Thì Nhậm-Tam Điệp-Ninh Bình"
                                           },
                                           {
                                               "TongCOD":  10.5,
                                               "CODTuNop":  10.5,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Phố Bưởi-Xã Kim Bôi-Hòa Bình"
                                           },
                                           {
                                               "TongCOD":  31.01,
                                               "CODTuNop":  31.01,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Phố Rịa-Nho Quan-Ninh Bình"
                                           },
                                           {
                                               "TongCOD":  77.62,
                                               "CODTuNop":  77.62,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục QL32 Khu Bình Dân-Xã Thanh Sơn-Phú Thọ"
                                           },
                                           {
                                               "TongCOD":  260.93,
                                               "CODTuNop":  260.93,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục QL37B-Thị trấn Liễu Đề-Nghĩa Hưng-Nam Định"
                                           },
                                           {
                                               "TongCOD":  74.11,
                                               "CODTuNop":  74.11,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục QL38B Yên Khánh-Ý Yên-Nam Định"
                                           },
                                           {
                                               "TongCOD":  85.92,
                                               "CODTuNop":  85.92,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Quốc Lộ 10-Yên Khánh-Ninh Bình"
                                           },
                                           {
                                               "TongCOD":  17.14,
                                               "CODTuNop":  17.14,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Quốc Lộ 12B-Yên Mô-Ninh Bình"
                                           },
                                           {
                                               "TongCOD":  70.91,
                                               "CODTuNop":  70.91,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Quốc Lộ 21A-Lý Thường Kiệt-Hà Nam"
                                           },
                                           {
                                               "TongCOD":  84.11,
                                               "CODTuNop":  84.11,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Quốc lộ 21B-Mỹ Hưng-Mỹ Lộc-Nam Định"
                                           },
                                           {
                                               "TongCOD":  148.38,
                                               "CODTuNop":  148.38,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Quốc Lộ 2-Xã Phù Ninh-Phú Thọ"
                                           },
                                           {
                                               "TongCOD":  80.61,
                                               "CODTuNop":  80.61,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Quốc Lộ 6-Tân Lạc-Hoà Bình"
                                           },
                                           {
                                               "TongCOD":  27.72,
                                               "CODTuNop":  27.72,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Quốc Lộ 6-Xã Cao Phong-Hòa Bình"
                                           },
                                           {
                                               "TongCOD":  93.98,
                                               "CODTuNop":  93.98,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Số 13 Nguyễn Bỉnh Khiêm-Lộc Hoà-Nam Định"
                                           },
                                           {
                                               "TongCOD":  88.68,
                                               "CODTuNop":  88.68,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Số nhà 03-Lốc Mới-Vụ Bản-Lạc Sơn-Hoà Bình"
                                           },
                                           {
                                               "TongCOD":  110.3,
                                               "CODTuNop":  110.3,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Tân Lập-Cẩm Khê-Phú Thọ"
                                           },
                                           {
                                               "TongCOD":  95.59,
                                               "CODTuNop":  95.59,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục TĐC Đường DT310-Xã Thiện Kế-Vĩnh Phúc"
                                           },
                                           {
                                               "TongCOD":  119.04,
                                               "CODTuNop":  119.04,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục TDP Nho Lâm-Xã Tam Hồng-Vĩnh Phúc"
                                           },
                                           {
                                               "TongCOD":  88.63,
                                               "CODTuNop":  88.63,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục TDP Thắng Lợi-Hương Canh-Bình Xuyên-Vĩnh Phúc"
                                           },
                                           {
                                               "TongCOD":  66.44,
                                               "CODTuNop":  66.44,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục TDP Văn Sơn-Lập Thạch-Vĩnh Phúc"
                                           },
                                           {
                                               "TongCOD":  32.29,
                                               "CODTuNop":  32.29,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thị trấn Hàng Trạm-Yên Thủy-Hòa Bình"
                                           },
                                           {
                                               "TongCOD":  99.18,
                                               "CODTuNop":  99.18,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thiện Vịnh-Cộng Hoà-Vụ Bản-Nam Định"
                                           },
                                           {
                                               "TongCOD":  104.9,
                                               "CODTuNop":  104.9,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thôn 1 Tâng-Xã Thanh Liêm-Hà Nam"
                                           },
                                           {
                                               "TongCOD":  74.03,
                                               "CODTuNop":  74.03,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thôn 1-Phù Vân-Phủ Lý-Hà Nam"
                                           },
                                           {
                                               "TongCOD":  43.84,
                                               "CODTuNop":  43.84,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu cục Thôn Gò Gai-Xã Tiên Lữ-Vĩnh Phúc"
                                           },
                                           {
                                               "TongCOD":  1.95,
                                               "CODTuNop":  1.95,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thôn Hòa Bình-Sông Lô-Vĩnh Phúc"
                                           },
                                           {
                                               "TongCOD":  78.25,
                                               "CODTuNop":  78.25,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thôn Làng Mạ-Xã Tam Dương Bắc-Vĩnh Phúc"
                                           },
                                           {
                                               "TongCOD":  7.63,
                                               "CODTuNop":  7.63,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thôn Lãng Sơn-Sông Lô-Vĩnh Phúc"
                                           },
                                           {
                                               "TongCOD":  13.09,
                                               "CODTuNop":  13.09,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thôn Xuân Hòa 1-Phúc Yên-Vĩnh Phúc"
                                           },
                                           {
                                               "TongCOD":  23.26,
                                               "CODTuNop":  23.26,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Tiểu Khu Liên Phương-Xã Đà Bắc-Hòa Bình"
                                           },
                                           {
                                               "TongCOD":  126.76,
                                               "CODTuNop":  126.76,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Trần Hưng Đạo-Quỳnh Lâm-Hòa Bình"
                                           },
                                           {
                                               "TongCOD":  114,
                                               "CODTuNop":  114,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Trực Ninh-Nam Định"
                                           },
                                           {
                                               "TongCOD":  83.09,
                                               "CODTuNop":  83.09,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Vĩnh Trụ-Lý Nhân-Hà Nam"
                                           },
                                           {
                                               "TongCOD":  68.8,
                                               "CODTuNop":  68.8,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu cục Xóm 11B-Xuân Vinh-Xuân Trường-Nam Định"
                                           },
                                           {
                                               "TongCOD":  106.24,
                                               "CODTuNop":  106.24,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Xóm 3-Xã Hải An-Nam Định"
                                           },
                                           {
                                               "TongCOD":  209.79,
                                               "CODTuNop":  209.79,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Xóm 4-Giao Yến-Giao Thuỷ-Nam Định"
                                           }
                                       ]
                    },
                    {
                        "TyLeTuNop":  100,
                        "TenVung":  "TTB",
                        "TongCOD":  9123.05,
                        "Trend":  0,
                        "CODTuNop":  9123.05,
                        "BuuCucList":  [
                                           {
                                               "TongCOD":  12.5,
                                               "CODTuNop":  12.5,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 08 Đổ Quang-Thị trấn Đông Phú-Quế Sơn-Quảng Nam"
                                           },
                                           {
                                               "TongCOD":  117.07,
                                               "CODTuNop":  117.07,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu cục 1056 Quang Trung -Chánh Lộ -Quảng Ngãi"
                                           },
                                           {
                                               "TongCOD":  41.78,
                                               "CODTuNop":  41.78,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 106 Kỳ Đồng-Q.Thanh Khê-Đà Nẵng"
                                           },
                                           {
                                               "TongCOD":  138.2,
                                               "CODTuNop":  138.2,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 106 Quang Trung-Đại Lộc-Quảng Nam"
                                           },
                                           {
                                               "TongCOD":  177.39,
                                               "CODTuNop":  177.39,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 114 Nguyễn Văn Trỗi-TP Tam Kỳ-Quảng Nam"
                                           },
                                           {
                                               "TongCOD":  22.39,
                                               "CODTuNop":  22.39,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 129 Võ Nguyên Giáp-Minh Hóa-Quảng Bình"
                                           },
                                           {
                                               "TongCOD":  105.9,
                                               "CODTuNop":  105.9,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 14 Lê Hoàn-Phường Hương Trà-Huế"
                                           },
                                           {
                                               "TongCOD":  121.07,
                                               "CODTuNop":  121.07,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 140 Tôn Đức Thắng-An Sơn-Tam Kỳ-Quảng Nam"
                                           },
                                           {
                                               "TongCOD":  99.06,
                                               "CODTuNop":  99.06,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 152 Hùng Vương-Ba Đồn-Quảng Bình"
                                           },
                                           {
                                               "TongCOD":  137.11,
                                               "CODTuNop":  137.11,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 18 Đường 3/2-Hải Lăng-Quảng Trị"
                                           },
                                           {
                                               "TongCOD":  135.46,
                                               "CODTuNop":  135.46,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 183 Trần Phú-Vĩnh Linh-Quảng Trị"
                                           },
                                           {
                                               "TongCOD":  270.87,
                                               "CODTuNop":  270.87,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 184 Nguyễn Tất Thành-Hội An-Quảng Nam"
                                           },
                                           {
                                               "TongCOD":  265.99,
                                               "CODTuNop":  265.99,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 202 Lý Thánh Tông-Phú Lộc-Huế"
                                           },
                                           {
                                               "TongCOD":  24.37,
                                               "CODTuNop":  24.37,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 215 Hùng Vương-Đắk Hà-Kon Tum"
                                           },
                                           {
                                               "TongCOD":  269.86,
                                               "CODTuNop":  269.86,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 217 Tiểu La-Hòa Cường-Đà Nẵng"
                                           },
                                           {
                                               "TongCOD":  124.02,
                                               "CODTuNop":  124.02,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 231 Võ An Ninh-Q.Cẩm Lệ-Đà Nẵng"
                                           },
                                           {
                                               "TongCOD":  368.89,
                                               "CODTuNop":  368.89,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 244 Lê Lợi-Đông Lương-Đông Hà-Quảng Trị"
                                           },
                                           {
                                               "TongCOD":  90.01,
                                               "CODTuNop":  90.01,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 25 Phạm Văn Đồng-TP.Đồng Hới-Quảng Bình"
                                           },
                                           {
                                               "TongCOD":  214.11,
                                               "CODTuNop":  214.11,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 252 Kinh Dương Vương-Thuận An-Phú Vang-Huế"
                                           },
                                           {
                                               "TongCOD":  234.66,
                                               "CODTuNop":  234.66,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 258 Lê Duẩn-Hướng Hóa-Quảng Trị"
                                           },
                                           {
                                               "TongCOD":  95.96,
                                               "CODTuNop":  95.96,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 274 Điện Biên Phủ-Duy Xuyên-Quảng Nam"
                                           },
                                           {
                                               "TongCOD":  330.78,
                                               "CODTuNop":  330.78,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 280 Lê Đại Hành-Q.Cẩm Lệ-Đà Nẵng"
                                           },
                                           {
                                               "TongCOD":  157.49,
                                               "CODTuNop":  157.49,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 303 Đào Duy Từ-Thắng Lợi-Kon Tum"
                                           },
                                           {
                                               "TongCOD":  136.91,
                                               "CODTuNop":  136.91,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 306 Hùng Vương-Bố Trạch-Quảng Bình"
                                           },
                                           {
                                               "TongCOD":  247.31,
                                               "CODTuNop":  247.31,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 345C Ngô Quyền-Sơn Trà-Đà Nẵng"
                                           },
                                           {
                                               "TongCOD":  592.87,
                                               "CODTuNop":  592.87,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 352 Lê Duẩn-Phú Thuận-Huế"
                                           },
                                           {
                                               "TongCOD":  179.3,
                                               "CODTuNop":  179.3,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 36 Lý Thái Tổ-Điện Bàn-Quảng Nam"
                                           },
                                           {
                                               "TongCOD":  176.73,
                                               "CODTuNop":  176.73,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 366 Hai Bà Trưng-Quảng Ngãi-Quảng Ngãi"
                                           },
                                           {
                                               "TongCOD":  103.38,
                                               "CODTuNop":  103.38,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 390 Phò Trạch-Phong Điền-Thừa Thiên Huế"
                                           },
                                           {
                                               "TongCOD":  290.93,
                                               "CODTuNop":  290.93,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 404 Lê Văn Hiến-Khuê Mỹ-Ngũ Hành Sơn-Đà Nẵng"
                                           },
                                           {
                                               "TongCOD":  81.67,
                                               "CODTuNop":  81.67,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 43 Nguyễn Nghiêm-Huyện Tư Nghĩa-Quảng Ngãi"
                                           },
                                           {
                                               "TongCOD":  26.92,
                                               "CODTuNop":  26.92,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 45 Trần Hưng Đạo-Thị trấn Plei Kần-Ngọc Hồi-Kon Tum"
                                           },
                                           {
                                               "TongCOD":  71.81,
                                               "CODTuNop":  71.81,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 463 Phạm Văn Đồng-Lê Lợi-TP.KonTum-KonTum"
                                           },
                                           {
                                               "TongCOD":  75.77,
                                               "CODTuNop":  75.77,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 58 Lý Thường Kiệt-Lệ Thủy-Quảng Bình"
                                           },
                                           {
                                               "TongCOD":  37.59,
                                               "CODTuNop":  37.59,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu cục 61 Tôn Đức Thắng-Tiên Phước-Quảng Nam"
                                           },
                                           {
                                               "TongCOD":  156.25,
                                               "CODTuNop":  156.25,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 62 Bắc Sơn-Q.Cẩm Lệ-Đà Nẵng"
                                           },
                                           {
                                               "TongCOD":  216.51,
                                               "CODTuNop":  216.51,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 621 Điện Biên Phủ-Q.Thanh Khê-Đà Nẵng"
                                           },
                                           {
                                               "TongCOD":  127.07,
                                               "CODTuNop":  127.07,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 63 Trần Thủ Độ-Điện Bàn-Quảng Nam"
                                           },
                                           {
                                               "TongCOD":  21.37,
                                               "CODTuNop":  21.37,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 67 Lý Thái Tổ-Xã Đồng Lê-Quảng Bình"
                                           },
                                           {
                                               "TongCOD":  51.65,
                                               "CODTuNop":  51.65,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 70 Đỗ Quang Thắng-Đức Phổ-Quảng Ngãi"
                                           },
                                           {
                                               "TongCOD":  165.73,
                                               "CODTuNop":  165.73,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 721 Võ Nguyên Giáp-Quảng Ngãi-Quảng Ngãi"
                                           },
                                           {
                                               "TongCOD":  94.13,
                                               "CODTuNop":  94.13,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 94 Thanh Tịnh-Hòa Minh-Liên Chiểu-Đà Nẵng"
                                           },
                                           {
                                               "TongCOD":  154.61,
                                               "CODTuNop":  154.61,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 955 Nguyễn Tất Thành-Hương Thủy-Thừa Thiên Huế"
                                           },
                                           {
                                               "TongCOD":  226.75,
                                               "CODTuNop":  226.75,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Bàu vàng 6-Hòa Khánh Bắc-Liên Chiểu-Đà Nẵng"
                                           },
                                           {
                                               "TongCOD":  159.71,
                                               "CODTuNop":  159.71,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Bình Phục-Thăng Bình-Quảng Nam"
                                           },
                                           {
                                               "TongCOD":  34.62,
                                               "CODTuNop":  34.62,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Châu Ổ-Bình Sơn-Quảng Ngãi"
                                           },
                                           {
                                               "TongCOD":  494.46,
                                               "CODTuNop":  494.46,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục G22 Cao Đình Độ-Phường Trường An-TP Huế"
                                           },
                                           {
                                               "TongCOD":  83.12,
                                               "CODTuNop":  83.12,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục KDC Nam Khang-Đức Thạnh-Mộ Đức-Quảng Ngãi"
                                           },
                                           {
                                               "TongCOD":  181.74,
                                               "CODTuNop":  181.74,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Nguyễn Chí Thanh-Núi Thành-Quảng Nam"
                                           },
                                           {
                                               "TongCOD":  19.53,
                                               "CODTuNop":  19.53,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Phong Nha-Bố Trạch-Quảng Bình"
                                           },
                                           {
                                               "TongCOD":  56.14,
                                               "CODTuNop":  56.14,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Phường Bắc Lý-Đồng Hới-Quảng Bình"
                                           },
                                           {
                                               "TongCOD":  278.68,
                                               "CODTuNop":  278.68,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Phường Thanh Bình-Hải Châu-Đà Nẵng"
                                           },
                                           {
                                               "TongCOD":  7.35,
                                               "CODTuNop":  7.35,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Quốc lộ 24-Kon Rẫy-Kon Tum"
                                           },
                                           {
                                               "TongCOD":  20.95,
                                               "CODTuNop":  20.95,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu cục Thị trấn Di Lăng-Sơn Hà-Quảng Ngãi"
                                           },
                                           {
                                               "TongCOD":  179.21,
                                               "CODTuNop":  179.21,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu cục Thọ Quang-Sơn Trà-Đà Nẵng"
                                           },
                                           {
                                               "TongCOD":  58.85,
                                               "CODTuNop":  58.85,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thôn Dy Lộc-Quảng Tùng-Quảng Trạch-Quảng Bình"
                                           },
                                           {
                                               "TongCOD":  36.1,
                                               "CODTuNop":  36.1,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thôn Liên Bằng-Xã Hương Thọ-Huế"
                                           },
                                           {
                                               "TongCOD":  53.15,
                                               "CODTuNop":  53.15,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thôn Phổ Lại-Thanh An-Cam Lộ-Quảng Trị"
                                           },
                                           {
                                               "TongCOD":  3.18,
                                               "CODTuNop":  3.18,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thôn Quảng Phú-A Lưới-Thừa Thiên Huế"
                                           },
                                           {
                                               "TongCOD":  250.24,
                                               "CODTuNop":  250.24,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Trường Sơn-Hòa Vang-Đà Nẵng"
                                           },
                                           {
                                               "TongCOD":  37.26,
                                               "CODTuNop":  37.26,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Vận Hành 1056 Quang-Chánh Lộ-Quảng Ngãi"
                                           },
                                           {
                                               "TongCOD":  22.54,
                                               "CODTuNop":  22.54,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Vận Hành Ba Tơ-Quảng Ngãi"
                                           },
                                           {
                                               "TongCOD":  27.38,
                                               "CODTuNop":  27.38,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Vận Hành Bắc Trà My-Quảng Nam"
                                           },
                                           {
                                               "TongCOD":  7.31,
                                               "CODTuNop":  7.31,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Vận Hành Châu Ổ-Bình Sơn-Quảng Ngãi"
                                           },
                                           {
                                               "TongCOD":  20.49,
                                               "CODTuNop":  20.49,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Vận Hành Đắk Tô-Quảng Ngãi"
                                           },
                                           {
                                               "TongCOD":  26.81,
                                               "CODTuNop":  26.81,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Vận Hành Đảo Lý Sơn-Quảng Ngãi"
                                           },
                                           {
                                               "TongCOD":  0.97,
                                               "CODTuNop":  0.97,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Vận Hành Kon Plông-Kon Tum"
                                           },
                                           {
                                               "TongCOD":  1.15,
                                               "CODTuNop":  1.15,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Vận Hành Nam Trà My-Quảng Nam"
                                           },
                                           {
                                               "TongCOD":  5.28,
                                               "CODTuNop":  5.28,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Vận Hành Sa Thầy-Quảng Ngãi"
                                           },
                                           {
                                               "TongCOD":  60.2,
                                               "CODTuNop":  60.2,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Vận Hành Thôn Tiền Phong-Xã Bắc Trạch-Quảng Bình"
                                           },
                                           {
                                               "TongCOD":  8.77,
                                               "CODTuNop":  8.77,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Vận Hành Tu Mơ Rông-Đắk Tô-Kon Tum"
                                           },
                                           {
                                               "TongCOD":  27.05,
                                               "CODTuNop":  27.05,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Vận Hành Xã Đăk Pék-Kon Tum"
                                           },
                                           {
                                               "TongCOD":  9.94,
                                               "CODTuNop":  9.94,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Vận Hành Xã Nam Giang-Đà Nẵng"
                                           },
                                           {
                                               "TongCOD":  0.13,
                                               "CODTuNop":  0.13,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Vận Hành Xã Tây Giang-Quảng Nam"
                                           },
                                           {
                                               "TongCOD":  9.62,
                                               "CODTuNop":  9.62,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Vận Hành Xã Thạnh Mỹ-Đà Nẵng"
                                           },
                                           {
                                               "TongCOD":  69.73,
                                               "CODTuNop":  69.73,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Võ Ninh-Quảng Ninh-Quảng Bình"
                                           },
                                           {
                                               "TongCOD":  43.43,
                                               "CODTuNop":  43.43,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Xóm Trảng Bông-Bình Sơn-Quảng Ngãi"
                                           },
                                           {
                                               "TongCOD":  18.68,
                                               "CODTuNop":  18.68,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Điểm Xử Lý Hàng Ia Hdrai-Quảng Ngãi"
                                           },
                                           {
                                               "TongCOD":  1.65,
                                               "CODTuNop":  1.65,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Điểm Xử Lý Hàng Mo Rai-Quảng Ngãi"
                                           },
                                           {
                                               "TongCOD":  17.43,
                                               "CODTuNop":  17.43,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Điểm xử lý hàng Nam Đông-Thừa Thiên Huế"
                                           }
                                       ]
                    },
                    {
                        "TyLeTuNop":  100,
                        "TenVung":  "XBG",
                        "TongCOD":  6046.4,
                        "Trend":  0,
                        "CODTuNop":  6046.4,
                        "BuuCucList":  [
                                           {
                                               "TongCOD":  31.35,
                                               "CODTuNop":  31.35,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 01 Khu Dân Cư-Tứ Minh-Hải Dương"
                                           },
                                           {
                                               "TongCOD":  4.75,
                                               "CODTuNop":  4.75,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 12 Nguyễn Thị Minh Khai-Bắc Sơn-Lạng Sơn"
                                           },
                                           {
                                               "TongCOD":  131.78,
                                               "CODTuNop":  131.78,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 120 An Định-Tp.Hải Dương-Hải Dương"
                                           },
                                           {
                                               "TongCOD":  61.73,
                                               "CODTuNop":  61.73,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 130 Trần Văn Lan-Q.Hải An-Hải Phòng"
                                           },
                                           {
                                               "TongCOD":  317.96,
                                               "CODTuNop":  317.96,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 132 Quang Trung-Uông Bí-Quảng Ninh"
                                           },
                                           {
                                               "TongCOD":  49.53,
                                               "CODTuNop":  49.53,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 16A An Toàn-Hoà Nghĩa-Dương Kinh-Hải Phòng"
                                           },
                                           {
                                               "TongCOD":  189.66,
                                               "CODTuNop":  189.66,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 18 Bạch Năng Thi-Hải Dương"
                                           },
                                           {
                                               "TongCOD":  59.76,
                                               "CODTuNop":  59.76,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 204 Lê Lai-Máy Chai-Ngô Quyền-Hải Phòng"
                                           },
                                           {
                                               "TongCOD":  37.54,
                                               "CODTuNop":  37.54,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 215 Đề Nắm-Yên Thế-Bắc Giang"
                                           },
                                           {
                                               "TongCOD":  310.94,
                                               "CODTuNop":  310.94,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 228 Nguyễn Bình-Quảng Yên-Quảng Ninh"
                                           },
                                           {
                                               "TongCOD":  50.47,
                                               "CODTuNop":  50.47,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 29 Ngô Gia Tự-Móng Cái-Quảng Ninh 01"
                                           },
                                           {
                                               "TongCOD":  82.72,
                                               "CODTuNop":  82.72,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 320 Đường Hùng Vương-Mai Pha-Lạng Sơn"
                                           },
                                           {
                                               "TongCOD":  132,
                                               "CODTuNop":  132,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 386 Lê Duẩn-Kiến An-Hải Phòng"
                                           },
                                           {
                                               "TongCOD":  18.59,
                                               "CODTuNop":  18.59,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 415 Nguyễn Chế Nghĩa-Gia Lộc-Hải Dương"
                                           },
                                           {
                                               "TongCOD":  50.57,
                                               "CODTuNop":  50.57,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 54 Lưu Úc-Phù Liễn-Kiến An-Hải Phòng"
                                           },
                                           {
                                               "TongCOD":  268.73,
                                               "CODTuNop":  268.73,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 61 đường Trần Đăng Tuyển-Vũ Ninh-Bắc Ninh"
                                           },
                                           {
                                               "TongCOD":  153.5,
                                               "CODTuNop":  153.5,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 610 Nguyễn Bình-Đông Triều-Quảng Ninh"
                                           },
                                           {
                                               "TongCOD":  0.04,
                                               "CODTuNop":  0.04,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 827 Trần Hưng Đạo-Kinh Môn-Hải Dương"
                                           },
                                           {
                                               "TongCOD":  161.45,
                                               "CODTuNop":  161.45,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 88 Thống Nhất-Kim Thành-Hải Dương"
                                           },
                                           {
                                               "TongCOD":  129,
                                               "CODTuNop":  129,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Chợ Ấp Đồn-Yên Trung-Yên Phong-Bắc Ninh"
                                           },
                                           {
                                               "TongCOD":  31.16,
                                               "CODTuNop":  31.16,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Chợ Vé-Ninh Giang-Hải Dương"
                                           },
                                           {
                                               "TongCOD":  71.01,
                                               "CODTuNop":  71.01,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu cục Cửa Ông-Cẩm Phả-Quảng Ninh"
                                           },
                                           {
                                               "TongCOD":  165.76,
                                               "CODTuNop":  165.76,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Dĩnh Kế- TP.Bắc Giang-Bắc Giang"
                                           },
                                           {
                                               "TongCOD":  83.8,
                                               "CODTuNop":  83.8,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu cục Đồng Đăng-Cao Lộc-Lạng Sơn"
                                           },
                                           {
                                               "TongCOD":  22.58,
                                               "CODTuNop":  22.58,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục DT359-Thủy Nguyên-Hải Phòng"
                                           },
                                           {
                                               "TongCOD":  133.09,
                                               "CODTuNop":  133.09,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Đường 25/5 Thị Trấn Thanh Hà-Hải Dương"
                                           },
                                           {
                                               "TongCOD":  67.66,
                                               "CODTuNop":  67.66,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Đường Đại Huề-Chi Lăng-Lạng Sơn"
                                           },
                                           {
                                               "TongCOD":  40.15,
                                               "CODTuNop":  40.15,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Đường Hoàng Văn Thụ-Tràng Định-Lạng Sơn"
                                           },
                                           {
                                               "TongCOD":  111.68,
                                               "CODTuNop":  111.68,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Đường Tỉnh 356 Hùng Sơn-Cát Hải-Hải Phòng"
                                           },
                                           {
                                               "TongCOD":  26.05,
                                               "CODTuNop":  26.05,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Đường Tỉnh 388-Thạch Lỗi-Cẩm Giàng-Hải Dương"
                                           },
                                           {
                                               "TongCOD":  186.15,
                                               "CODTuNop":  186.15,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục KCN Vĩnh Niệm-Q.Lê Chân-Hải Phòng"
                                           },
                                           {
                                               "TongCOD":  109.68,
                                               "CODTuNop":  109.68,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục KĐT Nam Hồng-Từ Sơn-Bắc Ninh"
                                           },
                                           {
                                               "TongCOD":  26.83,
                                               "CODTuNop":  26.83,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục KĐT Thị Trấn Chờ-Yên Phong-Bắc Ninh"
                                           },
                                           {
                                               "TongCOD":  81.19,
                                               "CODTuNop":  81.19,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Khu Sơn Đông-Phường Nam Sơn-Bắc Ninh"
                                           },
                                           {
                                               "TongCOD":  69.59,
                                               "CODTuNop":  69.59,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục KP Mới Yên Lã-Từ Sơn-Bắc Ninh"
                                           },
                                           {
                                               "TongCOD":  63.03,
                                               "CODTuNop":  63.03,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu cục Lộc Bình-Lạng Sơn"
                                           },
                                           {
                                               "TongCOD":  51.4,
                                               "CODTuNop":  51.4,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Lũng Cùng-Hoàng Việt-Văn Lãng-Lạng Sơn"
                                           },
                                           {
                                               "TongCOD":  92.04,
                                               "CODTuNop":  92.04,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Nguyễn Phi Khanh-Tam Thanh-Lạng Sơn"
                                           },
                                           {
                                               "TongCOD":  130.41,
                                               "CODTuNop":  130.41,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Phố Bình Minh-Xã Hiệp Hòa-Bắc Giang"
                                           },
                                           {
                                               "TongCOD":  175.93,
                                               "CODTuNop":  175.93,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Phố Mới-Thủy Nguyên-Hải Phòng"
                                           },
                                           {
                                               "TongCOD":  96.26,
                                               "CODTuNop":  96.26,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Sở Dầu-Hồng Bàng-Hải Phòng"
                                           },
                                           {
                                               "TongCOD":  21.03,
                                               "CODTuNop":  21.03,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục TDP Đồng Lều-Yên Thế-Bắc Giang"
                                           },
                                           {
                                               "TongCOD":  31.05,
                                               "CODTuNop":  31.05,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục TDP Đồng Tiến-Phường Nam Đồ Sơn-Hải Phòng"
                                           },
                                           {
                                               "TongCOD":  155.11,
                                               "CODTuNop":  155.11,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục TDP Minh Phượng-Nham Biền-Bắc Giang"
                                           },
                                           {
                                               "TongCOD":  190.43,
                                               "CODTuNop":  190.43,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục TDP Mới QL31-Phường Chũ-Bắc Giang"
                                           },
                                           {
                                               "TongCOD":  50.01,
                                               "CODTuNop":  50.01,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu cục TDP Vàng-Việt Yên-Bắc Giang"
                                           },
                                           {
                                               "TongCOD":  40.77,
                                               "CODTuNop":  40.77,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu cục Thị Trấn Bình Gia-Bình Gia-Lạng Sơn"
                                           },
                                           {
                                               "TongCOD":  1.91,
                                               "CODTuNop":  1.91,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thị Trấn Cao Thượng-Tân Yên-Bắc Giang"
                                           },
                                           {
                                               "TongCOD":  98.73,
                                               "CODTuNop":  98.73,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thị Trấn Đồi Ngô-Lục Nam-Bắc Giang"
                                           },
                                           {
                                               "TongCOD":  85.23,
                                               "CODTuNop":  85.23,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thị Trấn Kép-Lạng Giang-Bắc Giang"
                                           },
                                           {
                                               "TongCOD":  8.15,
                                               "CODTuNop":  8.15,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thị trấn Nhã Nam-Tân Yên-Bắc Giang"
                                           },
                                           {
                                               "TongCOD":  41.02,
                                               "CODTuNop":  41.02,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thị Trấn Vôi-Lạng Giang-Bắc Giang"
                                           },
                                           {
                                               "TongCOD":  2.56,
                                               "CODTuNop":  2.56,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thôn Bất Lự-Hoàn Sơn-Tiên Du-Bắc Ninh"
                                           },
                                           {
                                               "TongCOD":  146.36,
                                               "CODTuNop":  146.36,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thôn Chùa-Việt Yên-Bắc Giang"
                                           },
                                           {
                                               "TongCOD":  125.71,
                                               "CODTuNop":  125.71,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thôn Đương Thắng-Kiến Thụy-Hải Phòng"
                                           },
                                           {
                                               "TongCOD":  68.94,
                                               "CODTuNop":  68.94,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thôn Lán Than-Xã Sơn Động-Bắc Giang"
                                           },
                                           {
                                               "TongCOD":  131.37,
                                               "CODTuNop":  131.37,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thôn Phục Lễ-Bình Giang-Hải Dương"
                                           },
                                           {
                                               "TongCOD":  27.47,
                                               "CODTuNop":  27.47,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thôn Phúc Nghiêm-Xã Phật Tích-Bắc Ninh"
                                           },
                                           {
                                               "TongCOD":  109.1,
                                               "CODTuNop":  109.1,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Thôn Xuân Sơn-An Thắng-An Lão-Hải Phòng"
                                           },
                                           {
                                               "TongCOD":  84.4,
                                               "CODTuNop":  84.4,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Tiên Lãng-Hải Phòng"
                                           },
                                           {
                                               "TongCOD":  0.85,
                                               "CODTuNop":  0.85,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục TL 293 Xóm Rén-Xã Trường Sơn-Bắc Giang"
                                           },
                                           {
                                               "TongCOD":  146.7,
                                               "CODTuNop":  146.7,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu cục Tổ 12-Kênh Dương-Lê Chân-Hải Phòng"
                                           },
                                           {
                                               "TongCOD":  160.65,
                                               "CODTuNop":  160.65,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Vĩnh Hải-Nguyễn Đức Cảnh-Đông Triều-Quảng Ninh"
                                           },
                                           {
                                               "TongCOD":  100.56,
                                               "CODTuNop":  100.56,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Vĩnh Niệm-Lê Chân-Hải Phòng"
                                           },
                                           {
                                               "TongCOD":  39.62,
                                               "CODTuNop":  39.62,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Xã Hương Lâm-Hiệp Hòa-Bắc Giang"
                                           },
                                           {
                                               "TongCOD":  13.73,
                                               "CODTuNop":  13.73,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Xã Tân Hưng-Vĩnh Bảo-Hải Phòng"
                                           },
                                           {
                                               "TongCOD":  87.4,
                                               "CODTuNop":  87.4,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Xuân Lai-Gia Bình-Bắc Ninh"
                                           }
                                       ]
                    }
                ]
};

// Date range of historical data
let dateRange = [];
let regionNames = [];

// ChartJs instance holder
let trendChart = null;

// Initialize Dashboard Data
function initData() {
    // We try to load from localStorage first. If not present, we use the pre-compiled appData.
    let stored = localStorage.getItem('cod_race_data_v2');
    if (stored) {
        try {
            let parsed = JSON.parse(stored);
            appData.regions = parsed.regions;
            appData.history = parsed.history;
            appData.selectedDate = parsed.selectedDate;
            appData.dateRange = parsed.dateRange || Object.keys(parsed.history).sort();
            appData.regionNames = parsed.regionNames || parsed.regions.map(r => r.TenVung).sort();
        } catch(e) {
            console.error("Error loading stored data from localStorage", e);
        }
    }
    
    // Set dateRange and regionNames
    dateRange = appData.dateRange || [];
    regionNames = appData.regionNames || [];
}

function saveDataToLocalStorage() {
    localStorage.setItem('cod_race_data_v2', JSON.stringify({
        regions: appData.regions,
        history: appData.history,
        selectedDate: appData.selectedDate,
        dateRange: appData.dateRange,
        regionNames: appData.regionNames
    }));
}

// Render UI Components
function renderDashboard() {
    if (!appData.regions || appData.regions.length === 0) {
        console.warn("No region data to render");
        return;
    }
    
    // 1. Calculate and update summary stats
    let totalCOD = 0;
    let totalCODTuNop = 0;

    appData.regions.forEach(r => {
        totalCOD += r.TongCOD;
        totalCODTuNop += r.CODTuNop;
    });

    let nationalRate = totalCOD > 0 ? Math.round((totalCODTuNop / totalCOD) * 1000) / 10 : 0;
    
    // Calculate national yesterday rate for trend
    let totalYesterdayCODTuNop = 0;
    appData.regions.forEach(r => {
        let yesterdayRate = r.TyLeTuNop - r.Trend;
        totalYesterdayCODTuNop += r.TongCOD * (yesterdayRate / 100);
    });
    let nationalYesterdayRate = totalCOD > 0 ? Math.round((totalYesterdayCODTuNop / totalCOD) * 1000) / 10 : 0;
    let nationalTrend = Math.round((nationalRate - nationalYesterdayRate) * 10) / 10;

    // Sort regions by rate to find winner
    let sortedRegions = [...appData.regions].sort((a, b) => b.TyLeTuNop - a.TyLeTuNop);
    let topRegion = sortedRegions[0] || { TenVung: 'Không có', TyLeTuNop: 0 };

    // Update HTML values
    document.getElementById('national-rate').innerText = nationalRate + '%';
    document.getElementById('national-trend-val').innerText = (nationalTrend >= 0 ? '+' : '') + nationalTrend + '%';
    let nationalTrendIcon = document.getElementById('national-trend-icon');
    if (nationalTrend >= 0) {
        nationalTrendIcon.className = 'fas fa-arrow-up up';
        document.getElementById('national-trend-container').className = 'stat-sub up';
    } else {
        nationalTrendIcon.className = 'fas fa-arrow-down down';
        document.getElementById('national-trend-container').className = 'stat-sub down';
    }

    document.getElementById('total-cod').innerText = formatCurrency(totalCOD) + 'M';
    document.getElementById('top-region-name').innerText = topRegion.TenVung;
    document.getElementById('top-region-rate').innerText = topRegion.TyLeTuNop + '%';

    // 2. Render Leaderboard/Race Track
    let raceContainer = document.getElementById('race-container');
    raceContainer.innerHTML = '';

    sortedRegions.forEach((region, index) => {
        let rank = index + 1;
        let isUnderperformer = region.TyLeTuNop < 60; // Flag regions below 60%

        let item = document.createElement('div');
        item.className = 'race-item' + (isUnderperformer ? ' underperformer' : '');
        item.onclick = () => openRegionDeepDive(region.TenVung);

        let trendClass = region.Trend >= 0 ? 'up' : 'down';
        let trendIcon = region.Trend >= 0 ? 'fa-caret-up' : 'fa-caret-down';
        item.innerHTML = 
            '<div class="race-rank">' + rank + '</div>' +
            '<div class="race-name-group">' +
                '<span class="race-name">' + region.TenVung + '</span>' +
                '<span class="race-details">' + formatCurrency(region.TongCOD) + 'M COD</span>' +
            '</div>' +
            '<div class="race-track-bg">' +
                '<div class="race-track-fill" style="width: ' + region.TyLeTuNop + '%"></div>' +
            '</div>' +
            '<div class="race-value-group">' +
                '<span class="race-percent">' + region.TyLeTuNop + '%</span>' +
                '<span class="race-trend ' + trendClass + '">' +
                    '<i class="fas ' + trendIcon + '"></i>' +
                    Math.abs(region.Trend) + '%' +
                '</span>' +
            '</div>';
        raceContainer.appendChild(item);
    });

    // 3. Render Top Priority Post Offices Overall (Smart Recommendation)
    renderOverallPriorityTable();
}

// Compute Improvement Priority Index (IPI) and render Priority Table
function renderOverallPriorityTable() {
    let allOffices = [];
    
    appData.regions.forEach(region => {
        if (!region.BuuCucList) return;
        region.BuuCucList.forEach(office => {
            // Improvement Potential = 100% - Current Self-deposit Rate
            let potential = 100 - office.TyLeTuNop;
            // Un-self-deposited COD amount
            let unsubmittedCOD = office.TongCOD - office.CODTuNop;
            // Formula: IPI = Potential (%) * Unsubmitted COD
            let ipiScore = Math.round(potential * unsubmittedCOD);

            allOffices.push({
                ...office,
                TenVung: region.TenVung,
                unsubmittedCOD: unsubmittedCOD,
                ipi: ipiScore
            });
        });
    });

    // Sort by Priority Index descending
    allOffices.sort((a, b) => b.ipi - a.ipi);

    let priorityBody = document.getElementById('priority-table-body');
    priorityBody.innerHTML = '';

    // Take top 5 priority post offices for action plan
    let topPriority = allOffices.slice(0, 5);

    topPriority.forEach(office => {
        let badgeClass = 'badge-high';
        let badgeText = 'Ưu tiên Cao';
        
        if (office.ipi < 2000) {
            badgeClass = 'badge-medium';
            badgeText = 'Trung bình';
        }

        let tr = document.createElement('tr');
        tr.innerHTML = 
            '<td>' +
                '<div style="font-weight: 700;">' + office.TenBuuCuc + '</div>' +
                '<div style="font-size: 0.75rem; color: var(--text-secondary);">' + office.TenVung + '</div>' +
            '</td>' +
            '<td style="text-align: right; font-weight: 600; color: var(--text-primary);">' + office.TongCOD + 'M</td>' +
            '<td style="text-align: right; font-weight: 700; color: var(--color-cyan);">' + office.TyLeTuNop + '%</td>' +
            '<td style="text-align: right; font-weight: 600; color: var(--color-pink);">' + Math.round(office.unsubmittedCOD * 100) / 100 + 'M</td>' +
            '<td style="text-align: center;">' +
                '<span class="badge ' + badgeClass + '"><i class="fas fa-exclamation-triangle"></i> ' + badgeText + '</span>' +
            '</td>';
        priorityBody.appendChild(tr);
    });
}

// Render charts using Chart.js
function renderCharts(compareRegions = []) {
    let trendChartEl = document.getElementById('trendChart');
    if (!trendChartEl) return;
    let ctx = trendChartEl.getContext('2d');
    
    if (trendChart) {
        trendChart.destroy();
    }

    // Default regions to compare
    if (compareRegions.length === 0) {
        let available = appData.regionNames || regionNames || [];
        compareRegions = [];
        if (available.includes('HCM')) compareRegions.push('HCM');
        if (available.includes('HNO')) compareRegions.push('HNO');
        if (compareRegions.length === 0 && available.length > 0) {
            compareRegions = available.slice(0, 2);
        }
    }

    // Prepare line chart datasets
    let datasets = compareRegions.map((regionName, idx) => {
        let colors = [
            '#06b6d4', // cyan
            '#ec4899', // pink
            '#10b981', // green
            '#8b5cf6', // purple
            '#f59e0b', // amber
            '#3b82f6'  // blue
        ];
        let color = colors[idx % colors.length];
        
        let dataPoints = [];
        dateRange.forEach(date => {
            let histList = appData.history[date] || [];
            let rData = histList.find(h => h.TenVung === regionName);
            dataPoints.push(rData ? rData.TyLeTuNop : 0);
        });

        return {
            label: regionName,
            data: dataPoints,
            borderColor: color,
            backgroundColor: color + '15',
            borderWidth: 3,
            tension: 0.3,
            pointBackgroundColor: color,
            pointHoverRadius: 6,
            fill: true
        };
    });

    trendChart = new Chart(ctx, {
        type: 'line',
        data: {
            labels: dateRange.map(d => d.substring(0, 5)), // format DD/MM
            datasets: datasets
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'top',
                    labels: {
                        color: '#94a3b8',
                        font: {
                            family: 'Plus Jakarta Sans',
                            size: 11,
                            weight: 600
                        },
                        boxWidth: 12,
                        padding: 15
                    }
                },
                tooltip: {
                    mode: 'index',
                    intersect: false,
                    titleFont: { family: 'Plus Jakarta Sans', weight: 'bold' },
                    bodyFont: { family: 'Plus Jakarta Sans' },
                    callbacks: {
                        label: function(context) {
                            return context.dataset.label + ': ' + context.raw + '%';
                        }
                    }
                }
            },
            scales: {
                y: {
                    min: 0,
                    max: 100,
                    grid: {
                        color: 'rgba(255, 255, 255, 0.05)',
                        drawBorder: false
                    },
                    ticks: {
                        color: '#64748b',
                        font: { family: 'Plus Jakarta Sans', size: 10 },
                        callback: function(value) { return value + '%'; }
                    }
                },
                x: {
                    grid: {
                        display: false
                    },
                    ticks: {
                        color: '#64748b',
                        font: { family: 'Plus Jakarta Sans', size: 10 }
                    }
                }
            }
        }
    });
}

// Open Region Deep-Dive Drawer
let regionDetailChart = null;

function openRegionDeepDive(regionName) {
    let region = appData.regions.find(r => r.TenVung === regionName);
    if (!region) return;

    // Fill general text values
    document.getElementById('drawer-region-name').innerText = region.TenVung;
    document.getElementById('drawer-stat-rate').innerText = region.TyLeTuNop + '%';
    document.getElementById('drawer-stat-cod').innerText = formatCurrency(region.TongCOD) + 'M';

    // Highlight underperforming warning
    let warningBox = document.getElementById('drawer-warning-box');
    if (region.TyLeTuNop < 60) {
        warningBox.style.display = 'block';
        document.getElementById('drawer-warning-text').innerHTML = 
            '<strong>Cảnh báo:</strong> Vùng ' + region.TenVung + ' đang có tỷ lệ dưới mức tiêu chuẩn (60%). ' +
            'Cần tập trung cải thiện các bưu cục có độ ưu tiên cao phía dưới!';
    } else {
        warningBox.style.display = 'none';
    }

    // Render detailed bưu cục list inside drawer, sorted by IPI
    let buuCucList = [...region.BuuCucList];
    
    // Add calculations
    buuCucList.forEach(bc => {
        let potential = 100 - bc.TyLeTuNop;
        let unsubmittedCOD = bc.TongCOD - bc.CODTuNop;
        bc.unsubmittedCOD = unsubmittedCOD;
        bc.ipi = Math.round(potential * unsubmittedCOD);
    });

    // Sort by priority (IPI)
    buuCucList.sort((a, b) => b.ipi - a.ipi);

    let listBody = document.getElementById('drawer-buucuc-table-body');
    listBody.innerHTML = '';

    buuCucList.forEach(office => {
        let badgeClass = 'badge-low';
        let badgeText = 'Thấp';

        if (office.ipi > 2000) {
            badgeClass = 'badge-high';
            badgeText = 'Ưu Tiên Cao';
        } else if (office.ipi > 500) {
            badgeClass = 'badge-medium';
            badgeText = 'Trung BÃ¬nh';
        }

        let tr = document.createElement('tr');
        tr.innerHTML = 
            '<td>' +
                '<div style="font-weight:700;">' + office.TenBuuCuc + '</div>' +
            '</td>' +
            '<td style="text-align: right; font-weight: 600;">' + office.TongCOD + 'M</td>' +
            '<td style="text-align: right; font-weight: 700; color: var(--color-cyan);">' + office.TyLeTuNop + '%</td>' +
            '<td style="text-align: center;">' +
                '<span class="badge ' + badgeClass + '">' + badgeText + '</span>' +
            '</td>';
        listBody.appendChild(tr);
    });

    // Draw Mini Bar Chart for Post Office Rates inside drawer
    let miniCtx = document.getElementById('drawerRegionChart').getContext('2d');
    if (regionDetailChart) {
        regionDetailChart.destroy();
    }

    // Sort by rate to show performance gap
    let rateSortedOffices = [...region.BuuCucList].sort((a, b) => a.TyLeTuNop - b.TyLeTuNop);

    regionDetailChart = new Chart(miniCtx, {
        type: 'bar',
        data: {
            labels: rateSortedOffices.map(o => o.TenBuuCuc),
            datasets: [{
                label: 'Tá»· lá»‡ tá»± ná»™p (%)',
                data: rateSortedOffices.map(o => o.TyLeTuNop),
                backgroundColor: rateSortedOffices.map(o => o.TyLeTuNop < 60 ? 'rgba(239, 68, 68, 0.7)' : 'rgba(59, 130, 246, 0.7)'),
                borderColor: rateSortedOffices.map(o => o.TyLeTuNop < 60 ? '#ef4444' : '#3b82f6'),
                borderWidth: 1,
                borderRadius: 4
            }]
        },
        options: {
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
                tooltip: {
                    callbacks: { label: function(context) { return 'Tỷ lệ: ' + context.raw + '%'; } }
                }
            },
            scales: {
                x: {
                    min: 0,
                    max: 100,
                    grid: { color: 'rgba(255, 255, 255, 0.05)' },
                    ticks: { color: '#64748b', font: { family: 'Plus Jakarta Sans', size: 9 } }
                },
                y: {
                    grid: { display: false },
                    ticks: { color: '#94a3b8', font: { family: 'Plus Jakarta Sans', size: 10 } }
                }
            }
        }
    });

    // Show Drawer
    document.getElementById('region-drawer').classList.add('active');
}

function closeRegionDeepDive() {
    document.getElementById('region-drawer').classList.remove('active');
}

// Setup Event Listeners
function setupEvents() {
    // 1. Comparison Checkboxes/Selectors Setup
    let compareSelect = document.getElementById('compare-select-container');
    compareSelect.innerHTML = '';
    
    // Sort regions alphabetically to make search easy
    let sortedRegionNames = [...regionNames].sort();
    
    // Let's render checkboxes in a dropdown-like container
    sortedRegionNames.forEach(rName => {
        let label = document.createElement('label');
        label.style.display = 'flex';
        label.style.cursor = 'pointer';
        label.style.alignItems = 'center';
        label.style.gap = '0.5rem';
        label.style.fontSize = '0.8125rem';
        label.style.color = 'var(--text-secondary)';
        label.style.padding = '0.25rem 0.5rem';

        // Check first 2 by default
        let isDefaultChecked = rName === regionNames[0] || rName === regionNames[4] || (regionNames.length <= 4 && rName === regionNames[0]);

        label.innerHTML = 
            '<input type="checkbox" value="' + rName + '" ' + (isDefaultChecked ? 'checked' : '') + ' style="accent-color: var(--color-cyan);">' +
            '<span>' + rName + '</span>';
        compareSelect.appendChild(label);
        
        // Listen to checkbox change to update chart
        label.querySelector('input').addEventListener('change', () => {
            let checkedInputs = compareSelect.querySelectorAll('input:checked');
            let selected = Array.from(checkedInputs).map(input => input.value);
            renderCharts(selected);
        });
    });

    // 2. Drawer close event
    document.getElementById('close-drawer-btn').addEventListener('click', closeRegionDeepDive);
    document.getElementById('region-drawer').addEventListener('click', (e) => {
        if (e.target.id === 'region-drawer') closeRegionDeepDive();
    });

    // 3. Reset Data Action
    document.getElementById('reset-btn').addEventListener('click', () => {
        if (confirm("Báº¡n cÃ³ cháº¯c cháº¯n muá»‘n cÃ i Ä‘áº·t láº¡i dá»¯ liá»‡u gá»‘c ban Ä‘áº§u khÃ´ng?")) {
            localStorage.removeItem('cod_race_data_v2');
            location.reload();
        }
    });

    // 4. Excel Import Setup
    setupExcelImport();

    // 5. Excel Template Download Setup
    document.getElementById('download-template-btn').addEventListener('click', exportSampleExcel);
}

// Drag and drop Excel importer using SheetJS
function setupExcelImport() {
    let dropZone = document.getElementById('drop-zone');
    let fileInput = document.getElementById('excel-file-input');

    if (!dropZone || !fileInput) return;

    // Make dropZone trigger file input
    dropZone.addEventListener('click', () => fileInput.click());

    // Highlight dropZone on drag events
    ['dragenter', 'dragover'].forEach(eventName => {
        dropZone.addEventListener(eventName, (e) => {
            e.preventDefault();
            e.stopPropagation();
            dropZone.classList.add('dragover');
        }, false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, (e) => {
            e.preventDefault();
            e.stopPropagation();
            dropZone.classList.remove('dragover');
        }, false);
    });

    // Handle dropped files
    dropZone.addEventListener('drop', (e) => {
        let dt = e.dataTransfer;
        let files = dt.files;
        if (files.length) {
            handleExcelFile(files[0]);
        }
    });

    // Handle selected files
    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length) {
            handleExcelFile(e.target.files[0]);
        }
    });
}

function handleExcelFile(file) {
    if (!file.name.endsWith('.xlsx') && !file.name.endsWith('.xls') && !file.name.endsWith('.csv')) {
        alert("Vui lÃ²ng chá»‰ táº£i lÃªn tá»‡p Excel (.xlsx, .xls) hoáº·c CSV!");
        return;
    }

    let reader = new FileReader();
    reader.onload = function(e) {
        try {
            let data = new Uint8Array(e.target.result);
            let workbook = XLSX.read(data, {type: 'array'});
            
            // Assume the first sheet is the operational sheet
            let sheetName = workbook.SheetNames[0];
            let worksheet = workbook.Sheets[sheetName];
            let json = XLSX.utils.sheet_to_json(worksheet);

            if (json.length === 0) {
                alert("Tá»‡p Excel rá»—ng hoáº·c Ä‘á»‹nh dáº¡ng khÃ´ng Ä‘Ãºng!");
                return;
            }

            // Parse and format the excel JSON to appData state
            processImportedJSON(json);
            
            alert("Táº£i lÃªn dá»¯ liá»‡u má»›i thÃ nh cÃ´ng!");
            saveDataToLocalStorage();
            initData();
            renderDashboard();
            renderCharts();
            setupEvents(); // Re-render checkboxes
        } catch (err) {
            console.error(err);
            alert("ÄÃ£ xáº£y ra lá»—i khi Ä‘á»c tá»‡p Excel. Vui lÃ²ng kiá»ƒm tra láº¡i cáº¥u trÃºc file!");
        }
    };
    reader.readAsArrayBuffer(file);
}

// Helper to clean and normalize Vietnamese diacritics for comparisons
function removeVietnameseDiacritics(str) {
    if (!str) return '';
    return str.toString()
        .toLowerCase()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .replace(/\u0111/g, 'd')
        .replace(/\u0110/g, 'd')
        .trim();
}

// Check if string is a valid dd/MM/yyyy date
function isValidDateStr(str) {
    if (!str) return false;
    let parts = str.split('/');
    if (parts.length !== 3) return false;
    let d = parseInt(parts[0]);
    let m = parseInt(parts[1]);
    let y = parseInt(parts[2]);
    return d >= 1 && d <= 31 && m >= 1 && m <= 12 && y >= 2000 && y <= 2100;
}

// Convert parsed Excel rows into structured appData
function processImportedJSON(rows) {
    if (!rows || rows.length === 0) return;

    let keys = Object.keys(rows[0]);
    
    // Normalize keys to find matches robustly
    let cleanKeys = keys.map(k => {
        if (!k) return "";
        return removeVietnameseDiacritics(k).replace(/[^a-z0-9]/g, ""); // strip underscores, spaces, etc.
    });

    let idxPayment = cleanKeys.findIndex(ck => ck.includes('hinhthucthanhtoan') || ck.includes('hinhthuc') || ck.includes('payment') || ck.includes('pttt'));
    let idxDate = cleanKeys.findIndex(ck => ck.includes('ngayhethong') || ck.includes('ngayhe') || ck.includes('ngay') || ck.includes('date'));
    let idxOffice = cleanKeys.findIndex(ck => ck.includes('chinhanh') || ck.includes('buucuc') || ck.includes('office') || ck.includes('postoffice'));
    let idxMoney = cleanKeys.findIndex(ck => ck.includes('tongsotiennop') || ck.includes('tongtien') || ck.includes('tongcod') || ck.includes('amount') || ck.includes('money') || ck.includes('tiennop'));
    let idxRegion = cleanKeys.findIndex(ck => ck.includes('vung') || ck.includes('vng') || ck.includes('region'));

    // Fallbacks to indexes if not found
    let colPayment = idxPayment !== -1 ? keys[idxPayment] : keys[0];
    let colDate = idxDate !== -1 ? keys[idxDate] : keys[1];
    let colOffice = idxOffice !== -1 ? keys[idxOffice] : keys[2];
    let colMoney = idxMoney !== -1 ? keys[idxMoney] : keys[5];
    let colRegion = idxRegion !== -1 ? keys[idxRegion] : keys[6];

    // 1. Gather all unique dates and regions
    let uniqueDates = new Set();
    let uniqueRegions = new Set();
    
    rows.forEach(row => {
        let dateVal = row[colDate];
        if (dateVal) {
            let dateStr = formatDateStr(dateVal);
            if (isValidDateStr(dateStr)) {
                uniqueDates.add(dateStr);
            }
        }
        let regionVal = row[colRegion];
        if (regionVal) {
            uniqueRegions.add(regionVal.toString().trim());
        }
    });

    let sortedDates = Array.from(uniqueDates).sort((a, b) => parseDateStr(a) - parseDateStr(b));
    let sortedRegions = Array.from(uniqueRegions).sort();
    
    if (sortedDates.length === 0) {
        alert("KhÃ´ng tÃ¬m tháº¥y thÃ´ng tin ngÃ y thÃ¡ng há»£p lá»‡!");
        return;
    }
    
    let latestDate = sortedDates[sortedDates.length - 1];
    let yesterdayDate = sortedDates.length > 1 ? sortedDates[sortedDates.length - 2] : latestDate;

    // Aggregate
    let aggMap = {}; // Key: Date|Region -> { ck: 0, tm: 0, total: 0 }
    let officeAggMap = {}; // Key: Region|Office -> { ck: 0, tm: 0, total: 0 } - for the latest date only

    rows.forEach(row => {
        let dateVal = row[colDate];
        let regionVal = row[colRegion];
        let paymentVal = row[colPayment];
        let moneyVal = row[colMoney];
        let officeVal = row[colOffice];

        if (!dateVal || !regionVal) return;

        let dateStr = formatDateStr(dateVal);
        if (!isValidDateStr(dateStr)) return;

        let regionStr = regionVal.toString().trim();
        let money = parseFloat(moneyVal) || 0;
        
        let isCK = false;
        let isTM = false;
        if (paymentVal) {
            let payStr = removeVietnameseDiacritics(paymentVal).replace(/\s+/g, '');
            if (payStr === 'ck' || payStr.includes('chuyenkhoan')) {
                isCK = true;
            } else if (payStr.includes('tienmat') || payStr === 'tm') {
                isTM = true;
            }
        }

        // Skip if neither CK nor TM to calculate rate correctly based strictly on CK / (CK + TM)
        if (!isCK && !isTM) return;

        let key = dateStr + '|' + regionStr;
        if (!aggMap[key]) {
            aggMap[key] = { ck: 0, tm: 0, total: 0 };
        }
        if (isCK) {
            aggMap[key].ck += money;
        }
        if (isTM) {
            aggMap[key].tm += money;
        }
        aggMap[key].total = aggMap[key].ck + aggMap[key].tm;

        if (dateStr === latestDate && officeVal) {
            let officeStr = officeVal.toString().trim();
            if (/^\d+-/.test(officeStr)) {
                officeStr = officeStr.replace(/^\d+-/, '');
            }
            let offKey = regionStr + '|' + officeStr;
            if (!officeAggMap[offKey]) {
                officeAggMap[offKey] = { ck: 0, tm: 0, total: 0 };
            }
            if (isCK) {
                officeAggMap[offKey].ck += money;
            }
            if (isTM) {
                officeAggMap[offKey].tm += money;
            }
            officeAggMap[offKey].total = officeAggMap[offKey].ck + officeAggMap[offKey].tm;
        }
    });

    // Populate history and regions
    let history = {};
    sortedDates.forEach(dateStr => {
        let dayList = [];
        sortedRegions.forEach(regStr => {
            let key = dateStr + '|' + regStr;
            let agg = aggMap[key];
            let rate = 0;
            if (agg && agg.total > 0) {
                rate = Math.round((agg.ck / agg.total) * 1000) / 10;
            }
            dayList.push({ TenVung: regStr, TyLeTuNop: rate });
        });
        history[dateStr] = dayList;
    });

    let regionsData = [];
    sortedRegions.forEach(regStr => {
        let keyToday = latestDate + '|' + regStr;
        let keyYesterday = yesterdayDate + '|' + regStr;

        let aggToday = aggMap[keyToday];
        let aggYesterday = aggMap[keyYesterday];

        let rateToday = 0;
        let totalCOD = 0;
        let codTuNop = 0;
        if (aggToday) {
            totalCOD = aggToday.total / 1000000;
            codTuNop = aggToday.ck / 1000000;
            if (aggToday.total > 0) {
                rateToday = Math.round((aggToday.ck / aggToday.total) * 1000) / 10;
            }
        }

        let rateYesterday = 0;
        if (aggYesterday && aggYesterday.total > 0) {
            rateYesterday = Math.round((aggYesterday.ck / aggYesterday.total) * 1000) / 10;
        }

        let trend = Math.round((rateToday - rateYesterday) * 10) / 10;

        let officesList = [];
        Object.keys(officeAggMap).forEach(offKey => {
            if (offKey.startsWith(regStr + '|')) {
                let offName = offKey.substring(regStr.length + 1);
                let offAgg = officeAggMap[offKey];
                let offTotal = offAgg.total / 1000000;
                let offCK = offAgg.ck / 1000000;
                let offRate = offAgg.total > 0 ? Math.round((offAgg.ck / offAgg.total) * 1000) / 10 : 0;

                officesList.push({
                    TenBuuCuc: offName,
                    TongCOD: Math.round(offTotal * 100) / 100,
                    CODTuNop: Math.round(offCK * 100) / 100,
                    TyLeTuNop: offRate
                });
            }
        });

        officesList.sort((a, b) => a.TenBuuCuc.localeCompare(b.TenBuuCuc));

        regionsData.push({
            TenVung: regStr,
            TongCOD: Math.round(totalCOD * 100) / 100,
            CODTuNop: Math.round(codTuNop * 100) / 100,
            TyLeTuNop: rateToday,
            Trend: trend,
            BuuCucList: officesList
        });
    });

    appData.selectedDate = latestDate;
    appData.regions = regionsData;
    appData.history = history;
    appData.dateRange = sortedDates;
    appData.regionNames = sortedRegions;
}

function formatDateStr(val) {
    if (!val) return "";
    
    // Handle Excel Serial Dates
    if (!isNaN(val) && !isNaN(parseFloat(val)) && parseFloat(val) > 30000 && parseFloat(val) < 60000) {
        let serial = parseFloat(val);
        let jsDate = new Date(Math.round((serial - 25569) * 86400 * 1000));
        let d = jsDate.getDate();
        let m = jsDate.getMonth() + 1;
        let y = jsDate.getFullYear();
        return (d < 10 ? '0' + d : d) + '/' + (m < 10 ? '0' + m : m) + '/' + y;
    }
    
    if (val instanceof Date) {
        let d = val.getDate();
        let m = val.getMonth() + 1;
        let y = val.getFullYear();
        return (d < 10 ? '0' + d : d) + '/' + (m < 10 ? '0' + m : m) + '/' + y;
    }
    
    let str = val.toString().trim();
    
    // Handle yyyy-MM-dd
    if (str.includes('-')) {
        let parts = str.split(' ')[0].split('-');
        if (parts[0].length === 4) {
            let y = parts[0];
            let m = parseInt(parts[1]);
            let d = parseInt(parts[2]);
            return (d < 10 ? '0' + d : d) + '/' + (m < 10 ? '0' + m : m) + '/' + y;
        }
    }
    
    // Handle dd/MM/yyyy or MM/dd/yyyy
    if (str.includes('/')) {
        let parts = str.split(' ')[0].split('/');
        if (parts.length >= 3) {
            let p0 = parseInt(parts[0]);
            let p1 = parseInt(parts[1]);
            let y = parts[2];
            
            // Default to assuming parts[0] is day and parts[1] is month (dd/MM/yyyy)
            let d = p0;
            let m = p1;
            if (p1 > 12) {
                // If parts[1] is > 12, it must be the day, so it was MM/dd/yyyy
                d = p1;
                m = p0;
            }
            return (d < 10 ? '0' + d : d) + '/' + (m < 10 ? '0' + m : m) + '/' + y;
        }
    }
    return str;
}

function parseDateStr(str) {
    if (!str) return new Date(0);
    let parts = str.split('/');
    if (parts.length !== 3) return new Date(0);
    return new Date(parts[2], parts[1] - 1, parts[0]);
}

function exportSampleExcel() {
    let sampleRows = [];
    const sampleDates = ['19/05/2026', '20/05/2026', '21/05/2026'];
    const sampleRegions = ["HCM", "HNO", "DNB", "TNT", "TNB"];
    
    sampleDates.forEach(date => {
        sampleRegions.forEach(region => {
            const offices = region === "HCM" ? ["BC Quáº­n 1", "BC BÃ¬nh Tháº¡nh"] : ["BC Chi NhÃ¡nh A", "BC Chi NhÃ¡nh B"];
            offices.forEach(office => {
                let codCK = Math.floor(100 + Math.random() * 200) * 100000;
                let codTM = Math.floor(50 + Math.random() * 100) * 100000;
                
                sampleRows.push({
                    "hinh_thuc_thanh_toan": "CK",
                    "ngay_he_thong": date.split('/').reverse().join('-'),
                    "chi_nhanh": office,
                    "ma_chi_nhanh": "BC_" + Math.floor(1000 + Math.random() * 9000),
                    "tong_so_phieu": Math.floor(5 + Math.random() * 10),
                    "tong_so_tien_nop": codCK,
                    "VÃ¹ng": region,
                    "NgÃ y": date,
                    "AM": "Area Manager Name"
                });

                sampleRows.push({
                    "hinh_thuc_thanh_toan": "Tiá»n máº·t",
                    "ngay_he_thong": date.split('/').reverse().join('-'),
                    "chi_nhanh": office,
                    "ma_chi_nhanh": "BC_" + Math.floor(1000 + Math.random() * 9000),
                    "tong_so_phieu": Math.floor(5 + Math.random() * 10),
                    "tong_so_tien_nop": codTM,
                    "VÃ¹ng": region,
                    "NgÃ y": date,
                    "AM": "Area Manager Name"
                });
            });
        });
    });

    let wb = XLSX.utils.book_new();
    let ws = XLSX.utils.json_to_sheet(sampleRows);
    XLSX.utils.book_append_sheet(wb, ws, "Dá»¯ liá»‡u máº«u ná»™p COD");
    XLSX.writeFile(wb, "mau_bao_cao_tu_nop_cod.xlsx");
}

// Utility Helpers
function formatCurrency(val) {
    return Math.round(val).toLocaleString('vi-VN');
}

// DOM Ready initialization
window.addEventListener('DOMContentLoaded', () => {
    initData();
    setupEvents();
    renderDashboard();
    renderCharts();
});
