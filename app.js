// ==========================================================================
// COD RACE MASTER - CORE APPLICATION LOGIC
// Includes: Aggregated data, Chart rendering, Excel parser, and Recommender
// ==========================================================================

const appData = {
    "selectedDate":  "21/05/2026",
    "history":  {
                    "08/05/2026":  [
                                       {
                                           "TenVung":  "BC KD",
                                           "TyLeTuNop":  14.6
                                       },
                                       {
                                           "TenVung":  "BTB",
                                           "TyLeTuNop":  8.8
                                       },
                                       {
                                           "TenVung":  "DBB",
                                           "TyLeTuNop":  9.1
                                       },
                                       {
                                           "TenVung":  "ĐCL",
                                           "TyLeTuNop":  15.5
                                       },
                                       {
                                           "TenVung":  "DNB",
                                           "TyLeTuNop":  16.5
                                       },
                                       {
                                           "TenVung":  "DSH",
                                           "TyLeTuNop":  0
                                       },
                                       {
                                           "TenVung":  "GXT",
                                           "TyLeTuNop":  73.3
                                       },
                                       {
                                           "TenVung":  "HCM",
                                           "TyLeTuNop":  70.6
                                       },
                                       {
                                           "TenVung":  "HNO",
                                           "TyLeTuNop":  21.5
                                       },
                                       {
                                           "TenVung":  "KHL",
                                           "TyLeTuNop":  3.9
                                       },
                                       {
                                           "TenVung":  "NTB",
                                           "TyLeTuNop":  11.8
                                       },
                                       {
                                           "TenVung":  "TBB",
                                           "TyLeTuNop":  3.5
                                       },
                                       {
                                           "TenVung":  "Tech",
                                           "TyLeTuNop":  0
                                       },
                                       {
                                           "TenVung":  "TNB",
                                           "TyLeTuNop":  15.5
                                       },
                                       {
                                           "TenVung":  "TNG",
                                           "TyLeTuNop":  1
                                       },
                                       {
                                           "TenVung":  "TNT",
                                           "TyLeTuNop":  20.1
                                       },
                                       {
                                           "TenVung":  "TTB",
                                           "TyLeTuNop":  11.9
                                       },
                                       {
                                           "TenVung":  "XBG",
                                           "TyLeTuNop":  6.8
                                       }
                                   ],
                    "01/05/2026":  [
                                       {
                                           "TenVung":  "BC KD",
                                           "TyLeTuNop":  1.4
                                       },
                                       {
                                           "TenVung":  "BTB",
                                           "TyLeTuNop":  8.4
                                       },
                                       {
                                           "TenVung":  "DBB",
                                           "TyLeTuNop":  9.1
                                       },
                                       {
                                           "TenVung":  "ĐCL",
                                           "TyLeTuNop":  11.7
                                       },
                                       {
                                           "TenVung":  "DNB",
                                           "TyLeTuNop":  15.6
                                       },
                                       {
                                           "TenVung":  "DSH",
                                           "TyLeTuNop":  0.3
                                       },
                                       {
                                           "TenVung":  "GXT",
                                           "TyLeTuNop":  72.7
                                       },
                                       {
                                           "TenVung":  "HCM",
                                           "TyLeTuNop":  66.8
                                       },
                                       {
                                           "TenVung":  "HNO",
                                           "TyLeTuNop":  13.2
                                       },
                                       {
                                           "TenVung":  "KHL",
                                           "TyLeTuNop":  0
                                       },
                                       {
                                           "TenVung":  "NTB",
                                           "TyLeTuNop":  11.7
                                       },
                                       {
                                           "TenVung":  "TBB",
                                           "TyLeTuNop":  5.1
                                       },
                                       {
                                           "TenVung":  "Tech",
                                           "TyLeTuNop":  0
                                       },
                                       {
                                           "TenVung":  "TNB",
                                           "TyLeTuNop":  12.9
                                       },
                                       {
                                           "TenVung":  "TNG",
                                           "TyLeTuNop":  0.9
                                       },
                                       {
                                           "TenVung":  "TNT",
                                           "TyLeTuNop":  20.2
                                       },
                                       {
                                           "TenVung":  "TTB",
                                           "TyLeTuNop":  11.5
                                       },
                                       {
                                           "TenVung":  "XBG",
                                           "TyLeTuNop":  6.4
                                       }
                                   ],
                    "09/05/2026":  [
                                       {
                                           "TenVung":  "BC KD",
                                           "TyLeTuNop":  9.9
                                       },
                                       {
                                           "TenVung":  "BTB",
                                           "TyLeTuNop":  8.3
                                       },
                                       {
                                           "TenVung":  "DBB",
                                           "TyLeTuNop":  10.3
                                       },
                                       {
                                           "TenVung":  "ĐCL",
                                           "TyLeTuNop":  12.2
                                       },
                                       {
                                           "TenVung":  "DNB",
                                           "TyLeTuNop":  16.4
                                       },
                                       {
                                           "TenVung":  "DSH",
                                           "TyLeTuNop":  0.1
                                       },
                                       {
                                           "TenVung":  "GXT",
                                           "TyLeTuNop":  71
                                       },
                                       {
                                           "TenVung":  "HCM",
                                           "TyLeTuNop":  69.2
                                       },
                                       {
                                           "TenVung":  "HNO",
                                           "TyLeTuNop":  22.4
                                       },
                                       {
                                           "TenVung":  "KHL",
                                           "TyLeTuNop":  0
                                       },
                                       {
                                           "TenVung":  "NTB",
                                           "TyLeTuNop":  12.5
                                       },
                                       {
                                           "TenVung":  "TBB",
                                           "TyLeTuNop":  4.4
                                       },
                                       {
                                           "TenVung":  "Tech",
                                           "TyLeTuNop":  0
                                       },
                                       {
                                           "TenVung":  "TNB",
                                           "TyLeTuNop":  14.2
                                       },
                                       {
                                           "TenVung":  "TNG",
                                           "TyLeTuNop":  1
                                       },
                                       {
                                           "TenVung":  "TNT",
                                           "TyLeTuNop":  18.9
                                       },
                                       {
                                           "TenVung":  "TTB",
                                           "TyLeTuNop":  11.6
                                       },
                                       {
                                           "TenVung":  "XBG",
                                           "TyLeTuNop":  6.8
                                       }
                                   ],
                    "06/05/2026":  [
                                       {
                                           "TenVung":  "BC KD",
                                           "TyLeTuNop":  5.8
                                       },
                                       {
                                           "TenVung":  "BTB",
                                           "TyLeTuNop":  8.9
                                       },
                                       {
                                           "TenVung":  "DBB",
                                           "TyLeTuNop":  10.3
                                       },
                                       {
                                           "TenVung":  "ĐCL",
                                           "TyLeTuNop":  16.3
                                       },
                                       {
                                           "TenVung":  "DNB",
                                           "TyLeTuNop":  17.3
                                       },
                                       {
                                           "TenVung":  "DSH",
                                           "TyLeTuNop":  0.3
                                       },
                                       {
                                           "TenVung":  "GXT",
                                           "TyLeTuNop":  74.1
                                       },
                                       {
                                           "TenVung":  "HCM",
                                           "TyLeTuNop":  70
                                       },
                                       {
                                           "TenVung":  "HNO",
                                           "TyLeTuNop":  21.5
                                       },
                                       {
                                           "TenVung":  "KHL",
                                           "TyLeTuNop":  0
                                       },
                                       {
                                           "TenVung":  "NTB",
                                           "TyLeTuNop":  14.2
                                       },
                                       {
                                           "TenVung":  "TBB",
                                           "TyLeTuNop":  4.2
                                       },
                                       {
                                           "TenVung":  "Tech",
                                           "TyLeTuNop":  0
                                       },
                                       {
                                           "TenVung":  "TNB",
                                           "TyLeTuNop":  25.1
                                       },
                                       {
                                           "TenVung":  "TNG",
                                           "TyLeTuNop":  0.9
                                       },
                                       {
                                           "TenVung":  "TNT",
                                           "TyLeTuNop":  20.8
                                       },
                                       {
                                           "TenVung":  "TTB",
                                           "TyLeTuNop":  12.4
                                       },
                                       {
                                           "TenVung":  "XBG",
                                           "TyLeTuNop":  7
                                       }
                                   ],
                    "15/05/2026":  [
                                       {
                                           "TenVung":  "BC KD",
                                           "TyLeTuNop":  31.5
                                       },
                                       {
                                           "TenVung":  "BTB",
                                           "TyLeTuNop":  23.3
                                       },
                                       {
                                           "TenVung":  "DBB",
                                           "TyLeTuNop":  15.8
                                       },
                                       {
                                           "TenVung":  "ĐCL",
                                           "TyLeTuNop":  20.5
                                       },
                                       {
                                           "TenVung":  "DNB",
                                           "TyLeTuNop":  21.5
                                       },
                                       {
                                           "TenVung":  "DSH",
                                           "TyLeTuNop":  46.4
                                       },
                                       {
                                           "TenVung":  "GXT",
                                           "TyLeTuNop":  76.5
                                       },
                                       {
                                           "TenVung":  "HCM",
                                           "TyLeTuNop":  76.1
                                       },
                                       {
                                           "TenVung":  "HNO",
                                           "TyLeTuNop":  20.1
                                       },
                                       {
                                           "TenVung":  "KHL",
                                           "TyLeTuNop":  0
                                       },
                                       {
                                           "TenVung":  "NTB",
                                           "TyLeTuNop":  11.9
                                       },
                                       {
                                           "TenVung":  "TBB",
                                           "TyLeTuNop":  4.9
                                       },
                                       {
                                           "TenVung":  "Tech",
                                           "TyLeTuNop":  0
                                       },
                                       {
                                           "TenVung":  "TNB",
                                           "TyLeTuNop":  14.8
                                       },
                                       {
                                           "TenVung":  "TNG",
                                           "TyLeTuNop":  14.2
                                       },
                                       {
                                           "TenVung":  "TNT",
                                           "TyLeTuNop":  36.1
                                       },
                                       {
                                           "TenVung":  "TTB",
                                           "TyLeTuNop":  34.4
                                       },
                                       {
                                           "TenVung":  "XBG",
                                           "TyLeTuNop":  17.4
                                       }
                                   ],
                    "19/05/2026":  [
                                       {
                                           "TenVung":  "BC KD",
                                           "TyLeTuNop":  24.6
                                       },
                                       {
                                           "TenVung":  "BTB",
                                           "TyLeTuNop":  27.3
                                       },
                                       {
                                           "TenVung":  "DBB",
                                           "TyLeTuNop":  17.9
                                       },
                                       {
                                           "TenVung":  "ĐCL",
                                           "TyLeTuNop":  23.7
                                       },
                                       {
                                           "TenVung":  "DNB",
                                           "TyLeTuNop":  27.4
                                       },
                                       {
                                           "TenVung":  "DSH",
                                           "TyLeTuNop":  46.5
                                       },
                                       {
                                           "TenVung":  "GXT",
                                           "TyLeTuNop":  77.1
                                       },
                                       {
                                           "TenVung":  "HCM",
                                           "TyLeTuNop":  77.3
                                       },
                                       {
                                           "TenVung":  "HNO",
                                           "TyLeTuNop":  24.4
                                       },
                                       {
                                           "TenVung":  "KHL",
                                           "TyLeTuNop":  5.3
                                       },
                                       {
                                           "TenVung":  "NTB",
                                           "TyLeTuNop":  16.8
                                       },
                                       {
                                           "TenVung":  "TBB",
                                           "TyLeTuNop":  26.4
                                       },
                                       {
                                           "TenVung":  "Tech",
                                           "TyLeTuNop":  0
                                       },
                                       {
                                           "TenVung":  "TNB",
                                           "TyLeTuNop":  26
                                       },
                                       {
                                           "TenVung":  "TNG",
                                           "TyLeTuNop":  22.3
                                       },
                                       {
                                           "TenVung":  "TNT",
                                           "TyLeTuNop":  46.7
                                       },
                                       {
                                           "TenVung":  "TTB",
                                           "TyLeTuNop":  52.6
                                       },
                                       {
                                           "TenVung":  "XBG",
                                           "TyLeTuNop":  19.7
                                       }
                                   ],
                    "16/05/2026":  [
                                       {
                                           "TenVung":  "BC KD",
                                           "TyLeTuNop":  18.6
                                       },
                                       {
                                           "TenVung":  "BTB",
                                           "TyLeTuNop":  23.7
                                       },
                                       {
                                           "TenVung":  "DBB",
                                           "TyLeTuNop":  15.8
                                       },
                                       {
                                           "TenVung":  "ĐCL",
                                           "TyLeTuNop":  19.9
                                       },
                                       {
                                           "TenVung":  "DNB",
                                           "TyLeTuNop":  21.6
                                       },
                                       {
                                           "TenVung":  "DSH",
                                           "TyLeTuNop":  52.2
                                       },
                                       {
                                           "TenVung":  "GXT",
                                           "TyLeTuNop":  76.4
                                       },
                                       {
                                           "TenVung":  "HCM",
                                           "TyLeTuNop":  75.1
                                       },
                                       {
                                           "TenVung":  "HNO",
                                           "TyLeTuNop":  19.9
                                       },
                                       {
                                           "TenVung":  "KHL",
                                           "TyLeTuNop":  22.1
                                       },
                                       {
                                           "TenVung":  "NTB",
                                           "TyLeTuNop":  12
                                       },
                                       {
                                           "TenVung":  "TBB",
                                           "TyLeTuNop":  10
                                       },
                                       {
                                           "TenVung":  "Tech",
                                           "TyLeTuNop":  0
                                       },
                                       {
                                           "TenVung":  "TNB",
                                           "TyLeTuNop":  14.5
                                       },
                                       {
                                           "TenVung":  "TNG",
                                           "TyLeTuNop":  13.9
                                       },
                                       {
                                           "TenVung":  "TNT",
                                           "TyLeTuNop":  37.7
                                       },
                                       {
                                           "TenVung":  "TTB",
                                           "TyLeTuNop":  35.6
                                       },
                                       {
                                           "TenVung":  "XBG",
                                           "TyLeTuNop":  18.5
                                       }
                                   ],
                    "03/05/2026":  [
                                       {
                                           "TenVung":  "BC KD",
                                           "TyLeTuNop":  3.4
                                       },
                                       {
                                           "TenVung":  "BTB",
                                           "TyLeTuNop":  7.4
                                       },
                                       {
                                           "TenVung":  "DBB",
                                           "TyLeTuNop":  9.5
                                       },
                                       {
                                           "TenVung":  "ĐCL",
                                           "TyLeTuNop":  11.8
                                       },
                                       {
                                           "TenVung":  "DNB",
                                           "TyLeTuNop":  12.9
                                       },
                                       {
                                           "TenVung":  "DSH",
                                           "TyLeTuNop":  0.4
                                       },
                                       {
                                           "TenVung":  "GXT",
                                           "TyLeTuNop":  72.4
                                       },
                                       {
                                           "TenVung":  "HCM",
                                           "TyLeTuNop":  65.7
                                       },
                                       {
                                           "TenVung":  "HNO",
                                           "TyLeTuNop":  21.1
                                       },
                                       {
                                           "TenVung":  "KHL",
                                           "TyLeTuNop":  0
                                       },
                                       {
                                           "TenVung":  "NTB",
                                           "TyLeTuNop":  14.2
                                       },
                                       {
                                           "TenVung":  "TBB",
                                           "TyLeTuNop":  4.7
                                       },
                                       {
                                           "TenVung":  "Tech",
                                           "TyLeTuNop":  0
                                       },
                                       {
                                           "TenVung":  "TNB",
                                           "TyLeTuNop":  19.5
                                       },
                                       {
                                           "TenVung":  "TNG",
                                           "TyLeTuNop":  1.1
                                       },
                                       {
                                           "TenVung":  "TNT",
                                           "TyLeTuNop":  18
                                       },
                                       {
                                           "TenVung":  "TTB",
                                           "TyLeTuNop":  14.1
                                       },
                                       {
                                           "TenVung":  "XBG",
                                           "TyLeTuNop":  6.4
                                       }
                                   ],
                    "05/05/2026":  [
                                       {
                                           "TenVung":  "BC KD",
                                           "TyLeTuNop":  9.1
                                       },
                                       {
                                           "TenVung":  "BTB",
                                           "TyLeTuNop":  9.3
                                       },
                                       {
                                           "TenVung":  "DBB",
                                           "TyLeTuNop":  10.5
                                       },
                                       {
                                           "TenVung":  "ĐCL",
                                           "TyLeTuNop":  18
                                       },
                                       {
                                           "TenVung":  "DNB",
                                           "TyLeTuNop":  17.7
                                       },
                                       {
                                           "TenVung":  "DSH",
                                           "TyLeTuNop":  0
                                       },
                                       {
                                           "TenVung":  "GXT",
                                           "TyLeTuNop":  76.5
                                       },
                                       {
                                           "TenVung":  "HCM",
                                           "TyLeTuNop":  70.7
                                       },
                                       {
                                           "TenVung":  "HNO",
                                           "TyLeTuNop":  21.8
                                       },
                                       {
                                           "TenVung":  "KHL",
                                           "TyLeTuNop":  0
                                       },
                                       {
                                           "TenVung":  "NTB",
                                           "TyLeTuNop":  16
                                       },
                                       {
                                           "TenVung":  "TBB",
                                           "TyLeTuNop":  5.6
                                       },
                                       {
                                           "TenVung":  "Tech",
                                           "TyLeTuNop":  0
                                       },
                                       {
                                           "TenVung":  "TNB",
                                           "TyLeTuNop":  26.6
                                       },
                                       {
                                           "TenVung":  "TNG",
                                           "TyLeTuNop":  2
                                       },
                                       {
                                           "TenVung":  "TNT",
                                           "TyLeTuNop":  22.5
                                       },
                                       {
                                           "TenVung":  "TTB",
                                           "TyLeTuNop":  13.4
                                       },
                                       {
                                           "TenVung":  "XBG",
                                           "TyLeTuNop":  7.5
                                       }
                                   ],
                    "02/05/2026":  [
                                       {
                                           "TenVung":  "BC KD",
                                           "TyLeTuNop":  30.3
                                       },
                                       {
                                           "TenVung":  "BTB",
                                           "TyLeTuNop":  6.9
                                       },
                                       {
                                           "TenVung":  "DBB",
                                           "TyLeTuNop":  8.7
                                       },
                                       {
                                           "TenVung":  "ĐCL",
                                           "TyLeTuNop":  10.7
                                       },
                                       {
                                           "TenVung":  "DNB",
                                           "TyLeTuNop":  13.6
                                       },
                                       {
                                           "TenVung":  "DSH",
                                           "TyLeTuNop":  0
                                       },
                                       {
                                           "TenVung":  "GXT",
                                           "TyLeTuNop":  78.4
                                       },
                                       {
                                           "TenVung":  "HCM",
                                           "TyLeTuNop":  64.6
                                       },
                                       {
                                           "TenVung":  "HNO",
                                           "TyLeTuNop":  20.2
                                       },
                                       {
                                           "TenVung":  "KHL",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "NTB",
                                           "TyLeTuNop":  10.4
                                       },
                                       {
                                           "TenVung":  "TBB",
                                           "TyLeTuNop":  3.4
                                       },
                                       {
                                           "TenVung":  "Tech",
                                           "TyLeTuNop":  0
                                       },
                                       {
                                           "TenVung":  "TNB",
                                           "TyLeTuNop":  12.6
                                       },
                                       {
                                           "TenVung":  "TNG",
                                           "TyLeTuNop":  0.5
                                       },
                                       {
                                           "TenVung":  "TNT",
                                           "TyLeTuNop":  15.7
                                       },
                                       {
                                           "TenVung":  "TTB",
                                           "TyLeTuNop":  12.5
                                       },
                                       {
                                           "TenVung":  "XBG",
                                           "TyLeTuNop":  6
                                       }
                                   ],
                    "20/05/2026":  [
                                       {
                                           "TenVung":  "BC KD",
                                           "TyLeTuNop":  26.6
                                       },
                                       {
                                           "TenVung":  "BTB",
                                           "TyLeTuNop":  35.2
                                       },
                                       {
                                           "TenVung":  "DBB",
                                           "TyLeTuNop":  31.3
                                       },
                                       {
                                           "TenVung":  "ĐCL",
                                           "TyLeTuNop":  34.3
                                       },
                                       {
                                           "TenVung":  "DNB",
                                           "TyLeTuNop":  34.7
                                       },
                                       {
                                           "TenVung":  "DSH",
                                           "TyLeTuNop":  53
                                       },
                                       {
                                           "TenVung":  "GXT",
                                           "TyLeTuNop":  81.4
                                       },
                                       {
                                           "TenVung":  "HCM",
                                           "TyLeTuNop":  80.1
                                       },
                                       {
                                           "TenVung":  "HNO",
                                           "TyLeTuNop":  28.8
                                       },
                                       {
                                           "TenVung":  "KHL",
                                           "TyLeTuNop":  0
                                       },
                                       {
                                           "TenVung":  "NTB",
                                           "TyLeTuNop":  18.5
                                       },
                                       {
                                           "TenVung":  "TBB",
                                           "TyLeTuNop":  26.6
                                       },
                                       {
                                           "TenVung":  "Tech",
                                           "TyLeTuNop":  0
                                       },
                                       {
                                           "TenVung":  "TNB",
                                           "TyLeTuNop":  34.7
                                       },
                                       {
                                           "TenVung":  "TNG",
                                           "TyLeTuNop":  34
                                       },
                                       {
                                           "TenVung":  "TNT",
                                           "TyLeTuNop":  51.1
                                       },
                                       {
                                           "TenVung":  "TTB",
                                           "TyLeTuNop":  65.3
                                       },
                                       {
                                           "TenVung":  "XBG",
                                           "TyLeTuNop":  26.2
                                       }
                                   ],
                    "12/05/2026":  [
                                       {
                                           "TenVung":  "BC KD",
                                           "TyLeTuNop":  4.5
                                       },
                                       {
                                           "TenVung":  "BTB",
                                           "TyLeTuNop":  10.6
                                       },
                                       {
                                           "TenVung":  "DBB",
                                           "TyLeTuNop":  9.3
                                       },
                                       {
                                           "TenVung":  "ĐCL",
                                           "TyLeTuNop":  14.9
                                       },
                                       {
                                           "TenVung":  "DNB",
                                           "TyLeTuNop":  15.8
                                       },
                                       {
                                           "TenVung":  "DSH",
                                           "TyLeTuNop":  5
                                       },
                                       {
                                           "TenVung":  "GXT",
                                           "TyLeTuNop":  77.8
                                       },
                                       {
                                           "TenVung":  "HCM",
                                           "TyLeTuNop":  74.3
                                       },
                                       {
                                           "TenVung":  "HNO",
                                           "TyLeTuNop":  21.5
                                       },
                                       {
                                           "TenVung":  "KHL",
                                           "TyLeTuNop":  55.3
                                       },
                                       {
                                           "TenVung":  "NTB",
                                           "TyLeTuNop":  14.9
                                       },
                                       {
                                           "TenVung":  "TBB",
                                           "TyLeTuNop":  5
                                       },
                                       {
                                           "TenVung":  "Tech",
                                           "TyLeTuNop":  0
                                       },
                                       {
                                           "TenVung":  "TNB",
                                           "TyLeTuNop":  24.5
                                       },
                                       {
                                           "TenVung":  "TNG",
                                           "TyLeTuNop":  0.7
                                       },
                                       {
                                           "TenVung":  "TNT",
                                           "TyLeTuNop":  27.6
                                       },
                                       {
                                           "TenVung":  "TTB",
                                           "TyLeTuNop":  20.5
                                       },
                                       {
                                           "TenVung":  "XBG",
                                           "TyLeTuNop":  8
                                       }
                                   ],
                    "10/05/2026":  [
                                       {
                                           "TenVung":  "BC KD",
                                           "TyLeTuNop":  36.6
                                       },
                                       {
                                           "TenVung":  "BTB",
                                           "TyLeTuNop":  9.6
                                       },
                                       {
                                           "TenVung":  "DBB",
                                           "TyLeTuNop":  8.5
                                       },
                                       {
                                           "TenVung":  "ĐCL",
                                           "TyLeTuNop":  13.9
                                       },
                                       {
                                           "TenVung":  "DNB",
                                           "TyLeTuNop":  14.8
                                       },
                                       {
                                           "TenVung":  "DSH",
                                           "TyLeTuNop":  0
                                       },
                                       {
                                           "TenVung":  "GXT",
                                           "TyLeTuNop":  70.4
                                       },
                                       {
                                           "TenVung":  "HCM",
                                           "TyLeTuNop":  69.5
                                       },
                                       {
                                           "TenVung":  "HNO",
                                           "TyLeTuNop":  19.4
                                       },
                                       {
                                           "TenVung":  "KHL",
                                           "TyLeTuNop":  44.2
                                       },
                                       {
                                           "TenVung":  "NTB",
                                           "TyLeTuNop":  13.7
                                       },
                                       {
                                           "TenVung":  "TBB",
                                           "TyLeTuNop":  4.3
                                       },
                                       {
                                           "TenVung":  "Tech",
                                           "TyLeTuNop":  0
                                       },
                                       {
                                           "TenVung":  "TNB",
                                           "TyLeTuNop":  21.8
                                       },
                                       {
                                           "TenVung":  "TNG",
                                           "TyLeTuNop":  0.8
                                       },
                                       {
                                           "TenVung":  "TNT",
                                           "TyLeTuNop":  19.4
                                       },
                                       {
                                           "TenVung":  "TTB",
                                           "TyLeTuNop":  10.9
                                       },
                                       {
                                           "TenVung":  "XBG",
                                           "TyLeTuNop":  6.9
                                       }
                                   ],
                    "13/05/2026":  [
                                       {
                                           "TenVung":  "BC KD",
                                           "TyLeTuNop":  17.5
                                       },
                                       {
                                           "TenVung":  "BTB",
                                           "TyLeTuNop":  11.1
                                       },
                                       {
                                           "TenVung":  "DBB",
                                           "TyLeTuNop":  11.3
                                       },
                                       {
                                           "TenVung":  "ĐCL",
                                           "TyLeTuNop":  19.4
                                       },
                                       {
                                           "TenVung":  "DNB",
                                           "TyLeTuNop":  19.3
                                       },
                                       {
                                           "TenVung":  "DSH",
                                           "TyLeTuNop":  17.7
                                       },
                                       {
                                           "TenVung":  "GXT",
                                           "TyLeTuNop":  75.7
                                       },
                                       {
                                           "TenVung":  "HCM",
                                           "TyLeTuNop":  74.4
                                       },
                                       {
                                           "TenVung":  "HNO",
                                           "TyLeTuNop":  21.5
                                       },
                                       {
                                           "TenVung":  "KHL",
                                           "TyLeTuNop":  19.6
                                       },
                                       {
                                           "TenVung":  "NTB",
                                           "TyLeTuNop":  14.9
                                       },
                                       {
                                           "TenVung":  "TBB",
                                           "TyLeTuNop":  5.1
                                       },
                                       {
                                           "TenVung":  "Tech",
                                           "TyLeTuNop":  0
                                       },
                                       {
                                           "TenVung":  "TNB",
                                           "TyLeTuNop":  23.6
                                       },
                                       {
                                           "TenVung":  "TNG",
                                           "TyLeTuNop":  5.7
                                       },
                                       {
                                           "TenVung":  "TNT",
                                           "TyLeTuNop":  31.9
                                       },
                                       {
                                           "TenVung":  "TTB",
                                           "TyLeTuNop":  21.7
                                       },
                                       {
                                           "TenVung":  "XBG",
                                           "TyLeTuNop":  12.9
                                       }
                                   ],
                    "14/05/2026":  [
                                       {
                                           "TenVung":  "BC KD",
                                           "TyLeTuNop":  16
                                       },
                                       {
                                           "TenVung":  "BTB",
                                           "TyLeTuNop":  15.9
                                       },
                                       {
                                           "TenVung":  "DBB",
                                           "TyLeTuNop":  11.6
                                       },
                                       {
                                           "TenVung":  "ĐCL",
                                           "TyLeTuNop":  19.4
                                       },
                                       {
                                           "TenVung":  "DNB",
                                           "TyLeTuNop":  21.4
                                       },
                                       {
                                           "TenVung":  "DSH",
                                           "TyLeTuNop":  35.6
                                       },
                                       {
                                           "TenVung":  "GXT",
                                           "TyLeTuNop":  75.5
                                       },
                                       {
                                           "TenVung":  "HCM",
                                           "TyLeTuNop":  74.8
                                       },
                                       {
                                           "TenVung":  "HNO",
                                           "TyLeTuNop":  21.4
                                       },
                                       {
                                           "TenVung":  "KHL",
                                           "TyLeTuNop":  31.3
                                       },
                                       {
                                           "TenVung":  "NTB",
                                           "TyLeTuNop":  15
                                       },
                                       {
                                           "TenVung":  "TBB",
                                           "TyLeTuNop":  5.5
                                       },
                                       {
                                           "TenVung":  "Tech",
                                           "TyLeTuNop":  0
                                       },
                                       {
                                           "TenVung":  "TNB",
                                           "TyLeTuNop":  26.8
                                       },
                                       {
                                           "TenVung":  "TNG",
                                           "TyLeTuNop":  13.2
                                       },
                                       {
                                           "TenVung":  "TNT",
                                           "TyLeTuNop":  36.6
                                       },
                                       {
                                           "TenVung":  "TTB",
                                           "TyLeTuNop":  32.5
                                       },
                                       {
                                           "TenVung":  "XBG",
                                           "TyLeTuNop":  14.2
                                       }
                                   ],
                    "07/05/2026":  [
                                       {
                                           "TenVung":  "BC KD",
                                           "TyLeTuNop":  8.9
                                       },
                                       {
                                           "TenVung":  "BTB",
                                           "TyLeTuNop":  9.4
                                       },
                                       {
                                           "TenVung":  "DBB",
                                           "TyLeTuNop":  9.6
                                       },
                                       {
                                           "TenVung":  "ĐCL",
                                           "TyLeTuNop":  15.7
                                       },
                                       {
                                           "TenVung":  "DNB",
                                           "TyLeTuNop":  16
                                       },
                                       {
                                           "TenVung":  "DSH",
                                           "TyLeTuNop":  0.2
                                       },
                                       {
                                           "TenVung":  "GXT",
                                           "TyLeTuNop":  74.4
                                       },
                                       {
                                           "TenVung":  "HCM",
                                           "TyLeTuNop":  70.9
                                       },
                                       {
                                           "TenVung":  "HNO",
                                           "TyLeTuNop":  21.6
                                       },
                                       {
                                           "TenVung":  "KHL",
                                           "TyLeTuNop":  0.6
                                       },
                                       {
                                           "TenVung":  "NTB",
                                           "TyLeTuNop":  14.5
                                       },
                                       {
                                           "TenVung":  "TBB",
                                           "TyLeTuNop":  4.8
                                       },
                                       {
                                           "TenVung":  "Tech",
                                           "TyLeTuNop":  0
                                       },
                                       {
                                           "TenVung":  "TNB",
                                           "TyLeTuNop":  21.5
                                       },
                                       {
                                           "TenVung":  "TNG",
                                           "TyLeTuNop":  1
                                       },
                                       {
                                           "TenVung":  "TNT",
                                           "TyLeTuNop":  22
                                       },
                                       {
                                           "TenVung":  "TTB",
                                           "TyLeTuNop":  12.1
                                       },
                                       {
                                           "TenVung":  "XBG",
                                           "TyLeTuNop":  6.8
                                       }
                                   ],
                    "18/05/2026":  [
                                       {
                                           "TenVung":  "BC KD",
                                           "TyLeTuNop":  36.6
                                       },
                                       {
                                           "TenVung":  "BTB",
                                           "TyLeTuNop":  27.2
                                       },
                                       {
                                           "TenVung":  "DBB",
                                           "TyLeTuNop":  18.8
                                       },
                                       {
                                           "TenVung":  "ĐCL",
                                           "TyLeTuNop":  24.1
                                       },
                                       {
                                           "TenVung":  "DNB",
                                           "TyLeTuNop":  25.6
                                       },
                                       {
                                           "TenVung":  "DSH",
                                           "TyLeTuNop":  48.3
                                       },
                                       {
                                           "TenVung":  "GXT",
                                           "TyLeTuNop":  75.4
                                       },
                                       {
                                           "TenVung":  "HCM",
                                           "TyLeTuNop":  76.9
                                       },
                                       {
                                           "TenVung":  "HNO",
                                           "TyLeTuNop":  23.9
                                       },
                                       {
                                           "TenVung":  "KHL",
                                           "TyLeTuNop":  0
                                       },
                                       {
                                           "TenVung":  "NTB",
                                           "TyLeTuNop":  15.6
                                       },
                                       {
                                           "TenVung":  "TBB",
                                           "TyLeTuNop":  18.2
                                       },
                                       {
                                           "TenVung":  "Tech",
                                           "TyLeTuNop":  0
                                       },
                                       {
                                           "TenVung":  "TNB",
                                           "TyLeTuNop":  23.9
                                       },
                                       {
                                           "TenVung":  "TNG",
                                           "TyLeTuNop":  21.1
                                       },
                                       {
                                           "TenVung":  "TNT",
                                           "TyLeTuNop":  43.2
                                       },
                                       {
                                           "TenVung":  "TTB",
                                           "TyLeTuNop":  42.6
                                       },
                                       {
                                           "TenVung":  "XBG",
                                           "TyLeTuNop":  20.3
                                       }
                                   ],
                    "11/05/2026":  [
                                       {
                                           "TenVung":  "BC KD",
                                           "TyLeTuNop":  14.6
                                       },
                                       {
                                           "TenVung":  "BTB",
                                           "TyLeTuNop":  9
                                       },
                                       {
                                           "TenVung":  "DBB",
                                           "TyLeTuNop":  9.1
                                       },
                                       {
                                           "TenVung":  "ĐCL",
                                           "TyLeTuNop":  13.4
                                       },
                                       {
                                           "TenVung":  "DNB",
                                           "TyLeTuNop":  16.3
                                       },
                                       {
                                           "TenVung":  "DSH",
                                           "TyLeTuNop":  0
                                       },
                                       {
                                           "TenVung":  "GXT",
                                           "TyLeTuNop":  72.5
                                       },
                                       {
                                           "TenVung":  "HCM",
                                           "TyLeTuNop":  73.5
                                       },
                                       {
                                           "TenVung":  "HNO",
                                           "TyLeTuNop":  22.2
                                       },
                                       {
                                           "TenVung":  "KHL",
                                           "TyLeTuNop":  0
                                       },
                                       {
                                           "TenVung":  "NTB",
                                           "TyLeTuNop":  14.3
                                       },
                                       {
                                           "TenVung":  "TBB",
                                           "TyLeTuNop":  5.3
                                       },
                                       {
                                           "TenVung":  "Tech",
                                           "TyLeTuNop":  0
                                       },
                                       {
                                           "TenVung":  "TNB",
                                           "TyLeTuNop":  24
                                       },
                                       {
                                           "TenVung":  "TNG",
                                           "TyLeTuNop":  0.6
                                       },
                                       {
                                           "TenVung":  "TNT",
                                           "TyLeTuNop":  22.7
                                       },
                                       {
                                           "TenVung":  "TTB",
                                           "TyLeTuNop":  11.2
                                       },
                                       {
                                           "TenVung":  "XBG",
                                           "TyLeTuNop":  7.4
                                       }
                                   ],
                    "21/05/2026":  [
                                       {
                                           "TenVung":  "BC KD",
                                           "TyLeTuNop":  16.4
                                       },
                                       {
                                           "TenVung":  "BTB",
                                           "TyLeTuNop":  34.5
                                       },
                                       {
                                           "TenVung":  "DBB",
                                           "TyLeTuNop":  37.7
                                       },
                                       {
                                           "TenVung":  "ĐCL",
                                           "TyLeTuNop":  26.3
                                       },
                                       {
                                           "TenVung":  "DNB",
                                           "TyLeTuNop":  36.2
                                       },
                                       {
                                           "TenVung":  "DSH",
                                           "TyLeTuNop":  50
                                       },
                                       {
                                           "TenVung":  "GXT",
                                           "TyLeTuNop":  78.3
                                       },
                                       {
                                           "TenVung":  "HCM",
                                           "TyLeTuNop":  80
                                       },
                                       {
                                           "TenVung":  "HNO",
                                           "TyLeTuNop":  38.3
                                       },
                                       {
                                           "TenVung":  "KHL",
                                           "TyLeTuNop":  3.1
                                       },
                                       {
                                           "TenVung":  "NTB",
                                           "TyLeTuNop":  25.8
                                       },
                                       {
                                           "TenVung":  "TBB",
                                           "TyLeTuNop":  42
                                       },
                                       {
                                           "TenVung":  "Tech",
                                           "TyLeTuNop":  0
                                       },
                                       {
                                           "TenVung":  "TNB",
                                           "TyLeTuNop":  38.3
                                       },
                                       {
                                           "TenVung":  "TNG",
                                           "TyLeTuNop":  32.4
                                       },
                                       {
                                           "TenVung":  "TNT",
                                           "TyLeTuNop":  50.7
                                       },
                                       {
                                           "TenVung":  "TTB",
                                           "TyLeTuNop":  69
                                       },
                                       {
                                           "TenVung":  "XBG",
                                           "TyLeTuNop":  31.2
                                       }
                                   ],
                    "04/05/2026":  [
                                       {
                                           "TenVung":  "BC KD",
                                           "TyLeTuNop":  3.1
                                       },
                                       {
                                           "TenVung":  "BTB",
                                           "TyLeTuNop":  10
                                       },
                                       {
                                           "TenVung":  "DBB",
                                           "TyLeTuNop":  9.4
                                       },
                                       {
                                           "TenVung":  "ĐCL",
                                           "TyLeTuNop":  11.7
                                       },
                                       {
                                           "TenVung":  "DNB",
                                           "TyLeTuNop":  17.1
                                       },
                                       {
                                           "TenVung":  "DSH",
                                           "TyLeTuNop":  0
                                       },
                                       {
                                           "TenVung":  "GXT",
                                           "TyLeTuNop":  77.9
                                       },
                                       {
                                           "TenVung":  "HCM",
                                           "TyLeTuNop":  69.2
                                       },
                                       {
                                           "TenVung":  "HNO",
                                           "TyLeTuNop":  22.6
                                       },
                                       {
                                           "TenVung":  "KHL",
                                           "TyLeTuNop":  0
                                       },
                                       {
                                           "TenVung":  "NTB",
                                           "TyLeTuNop":  13.8
                                       },
                                       {
                                           "TenVung":  "TBB",
                                           "TyLeTuNop":  5.1
                                       },
                                       {
                                           "TenVung":  "Tech",
                                           "TyLeTuNop":  0
                                       },
                                       {
                                           "TenVung":  "TNB",
                                           "TyLeTuNop":  21.9
                                       },
                                       {
                                           "TenVung":  "TNG",
                                           "TyLeTuNop":  0.7
                                       },
                                       {
                                           "TenVung":  "TNT",
                                           "TyLeTuNop":  21.1
                                       },
                                       {
                                           "TenVung":  "TTB",
                                           "TyLeTuNop":  11.4
                                       },
                                       {
                                           "TenVung":  "XBG",
                                           "TyLeTuNop":  6.5
                                       }
                                   ],
                    "17/05/2026":  [
                                       {
                                           "TenVung":  "BC KD",
                                           "TyLeTuNop":  11.9
                                       },
                                       {
                                           "TenVung":  "BTB",
                                           "TyLeTuNop":  23.2
                                       },
                                       {
                                           "TenVung":  "DBB",
                                           "TyLeTuNop":  16.6
                                       },
                                       {
                                           "TenVung":  "ĐCL",
                                           "TyLeTuNop":  20
                                       },
                                       {
                                           "TenVung":  "DNB",
                                           "TyLeTuNop":  21.1
                                       },
                                       {
                                           "TenVung":  "DSH",
                                           "TyLeTuNop":  47
                                       },
                                       {
                                           "TenVung":  "GXT",
                                           "TyLeTuNop":  73.5
                                       },
                                       {
                                           "TenVung":  "HCM",
                                           "TyLeTuNop":  76.5
                                       },
                                       {
                                           "TenVung":  "HNO",
                                           "TyLeTuNop":  20
                                       },
                                       {
                                           "TenVung":  "KHL",
                                           "TyLeTuNop":  100
                                       },
                                       {
                                           "TenVung":  "NTB",
                                           "TyLeTuNop":  15.9
                                       },
                                       {
                                           "TenVung":  "TBB",
                                           "TyLeTuNop":  14.6
                                       },
                                       {
                                           "TenVung":  "Tech",
                                           "TyLeTuNop":  0
                                       },
                                       {
                                           "TenVung":  "TNB",
                                           "TyLeTuNop":  22.2
                                       },
                                       {
                                           "TenVung":  "TNG",
                                           "TyLeTuNop":  16.8
                                       },
                                       {
                                           "TenVung":  "TNT",
                                           "TyLeTuNop":  38.2
                                       },
                                       {
                                           "TenVung":  "TTB",
                                           "TyLeTuNop":  38.7
                                       },
                                       {
                                           "TenVung":  "XBG",
                                           "TyLeTuNop":  19.7
                                       }
                                   ]
                },
    "regionNames":  [
                        "BC KD",
                        "BTB",
                        "DBB",
                        "ĐCL",
                        "DNB",
                        "DSH",
                        "GXT",
                        "HCM",
                        "HNO",
                        "KHL",
                        "NTB",
                        "TBB",
                        "Tech",
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
                        "TyLeTuNop":  16.4,
                        "TenVung":  "BC KD",
                        "TongCOD":  15.55,
                        "Trend":  -10.2,
                        "CODTuNop":  2.56,
                        "BuuCucList":  [
                                           {
                                               "TongCOD":  4.9,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục KD C03/10 Lê Trọng Tấn-Q.Hà Đông-HN"
                                           },
                                           {
                                               "TongCOD":  0.05,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục KD 9 Phú Minh-Q.Bắc Từ Liêm-HN"
                                           },
                                           {
                                               "TongCOD":  1.69,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục KD Vũ Đức Uý-Thanh Trì-HN"
                                           },
                                           {
                                               "TongCOD":  2.79,
                                               "CODTuNop":  0.98,
                                               "TyLeTuNop":  35.2,
                                               "TenBuuCuc":  "Bưu Cục KD 93D Tân Thắng-Q.Tân Phú-HCM"
                                           },
                                           {
                                               "TongCOD":  0.7,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục KD 230 Hữu Hưng-Q.Nam Từ Liêm-HN"
                                           },
                                           {
                                               "TongCOD":  1.19,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục KD 309B Tô Ký-Hóc Môn-HCM"
                                           },
                                           {
                                               "TongCOD":  1.87,
                                               "CODTuNop":  1.57,
                                               "TyLeTuNop":  83.9,
                                               "TenBuuCuc":  "Bưu Cục KD 1196 Quang Trung-Q.Gò Vấp-HCM"
                                           },
                                           {
                                               "TongCOD":  1.53,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục KD 804 Tân Kỳ Tân Quý-Q.Bình Tân-HCM"
                                           },
                                           {
                                               "TongCOD":  0.82,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục KD Số 1 Nguyễn Cảnh Dị-Hoàng Mai-Hà Nội"
                                           }
                                       ]
                    },
                    {
                        "TyLeTuNop":  34.5,
                        "TenVung":  "BTB",
                        "TongCOD":  12958.51,
                        "Trend":  -0.7,
                        "CODTuNop":  4468.15,
                        "BuuCucList":  [
                                           {
                                               "TongCOD":  115.14,
                                               "CODTuNop":  52.15,
                                               "TyLeTuNop":  45.3,
                                               "TenBuuCuc":  "Bưu Cục Can Lộc-Hà Tĩnh"
                                           },
                                           {
                                               "TongCOD":  87.59,
                                               "CODTuNop":  33.29,
                                               "TyLeTuNop":  38,
                                               "TenBuuCuc":  "Bưu Cục Khu Phố Mỹ Ré-Như Xuân-Thanh Hoá"
                                           },
                                           {
                                               "TongCOD":  157.67,
                                               "CODTuNop":  22.81,
                                               "TyLeTuNop":  14.5,
                                               "TenBuuCuc":  "Bưu cục Khu 4-Quan Hoá-Thanh Hoá"
                                           },
                                           {
                                               "TongCOD":  189.7,
                                               "CODTuNop":  94.32,
                                               "TyLeTuNop":  49.7,
                                               "TenBuuCuc":  "Bưu Cục 415 Hà Huy Tập-Cẩm Xuyên-Hà Tĩnh"
                                           },
                                           {
                                               "TongCOD":  289.39,
                                               "CODTuNop":  100.76,
                                               "TyLeTuNop":  34.8,
                                               "TenBuuCuc":  "Bưu Cục 339 Đường Lam Sơn-Nông Cống-Thanh Hóa"
                                           },
                                           {
                                               "TongCOD":  104.05,
                                               "CODTuNop":  46.46,
                                               "TyLeTuNop":  44.7,
                                               "TenBuuCuc":  "Bưu Cục Thôn Nam Thạch A-Yên Trung-Yên Định-Thanh Hóa"
                                           },
                                           {
                                               "TongCOD":  159.08,
                                               "CODTuNop":  95.9,
                                               "TyLeTuNop":  60.3,
                                               "TenBuuCuc":  "Bưu Cục Khối Đông Sơn-Quế Phong-Nghệ An"
                                           },
                                           {
                                               "TongCOD":  176.86,
                                               "CODTuNop":  49.4,
                                               "TyLeTuNop":  27.9,
                                               "TenBuuCuc":  "Bưu Cục QL1A Kỳ Thư-Lê Đại Hành-Kỳ Anh-Hà Tĩnh"
                                           },
                                           {
                                               "TongCOD":  184.03,
                                               "CODTuNop":  54.91,
                                               "TyLeTuNop":  29.8,
                                               "TenBuuCuc":  "Bưu Cục 286 Lê Lợi-Trường Sơn-Sầm Sơn-Thanh Hóa"
                                           },
                                           {
                                               "TongCOD":  39.99,
                                               "CODTuNop":  39.99,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 357 Khối Hòa Tây-Tương Dương-Nghệ An"
                                           },
                                           {
                                               "TongCOD":  110.94,
                                               "CODTuNop":  20.01,
                                               "TyLeTuNop":  18,
                                               "TenBuuCuc":  "Bưu Cục 196 Khu 1-Vĩnh Lộc-Thanh Hóa"
                                           },
                                           {
                                               "TongCOD":  93.52,
                                               "CODTuNop":  37.59,
                                               "TyLeTuNop":  40.2,
                                               "TenBuuCuc":  "Bưu Cục Quốc Lộ 12C-Xã Kỳ Lạc-Hà Tĩnh"
                                           },
                                           {
                                               "TongCOD":  98.64,
                                               "CODTuNop":  55.24,
                                               "TyLeTuNop":  56,
                                               "TenBuuCuc":  "Bưu Cục Khối 1-Thị trấn Con Cuông-Huyện Con Cuông-Nghệ An"
                                           },
                                           {
                                               "TongCOD":  162.67,
                                               "CODTuNop":  102.78,
                                               "TyLeTuNop":  63.2,
                                               "TenBuuCuc":  "Bưu Cục 195 Nguyễn Huệ-Phú Sơn-Bỉm Sơn-Thanh Hóa"
                                           },
                                           {
                                               "TongCOD":  124.85,
                                               "CODTuNop":  61.72,
                                               "TyLeTuNop":  49.4,
                                               "TenBuuCuc":  "Bưu Cục 702 Lý Thái Tổ-Kỳ Anh-Hà Tĩnh"
                                           },
                                           {
                                               "TongCOD":  191.98,
                                               "CODTuNop":  54.11,
                                               "TyLeTuNop":  28.2,
                                               "TenBuuCuc":  "Bưu cục Tiểu Khu Yên Hạnh 1-Nga Sơn-Thanh Hoá"
                                           },
                                           {
                                               "TongCOD":  179.6,
                                               "CODTuNop":  36.54,
                                               "TyLeTuNop":  20.3,
                                               "TenBuuCuc":  "Bưu Cục Nguyễn Nhữ Soạn-Rừng Thông-Đông Sơn-Thanh Hóa"
                                           },
                                           {
                                               "TongCOD":  30.85,
                                               "CODTuNop":  30.85,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Điểm xử lý hàng Thị Trấn Mường Xén-Kỳ Sơn-Nghệ An"
                                           },
                                           {
                                               "TongCOD":  230.38,
                                               "CODTuNop":  43.09,
                                               "TyLeTuNop":  18.7,
                                               "TenBuuCuc":  "Bưu Cục Thôn Quyết Thắng-Xã Thiết Ống-Thanh Hóa"
                                           },
                                           {
                                               "TongCOD":  153.66,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục KP Tân Lập QL47-Xã Sao Vàng-Thanh Hóa"
                                           },
                                           {
                                               "TongCOD":  101.17,
                                               "CODTuNop":  1.48,
                                               "TyLeTuNop":  1.5,
                                               "TenBuuCuc":  "Bưu Cục Phan Đình Phùng-TT.Đức Thọ-Hà Tĩnh"
                                           },
                                           {
                                               "TongCOD":  94.92,
                                               "CODTuNop":  44.36,
                                               "TyLeTuNop":  46.7,
                                               "TenBuuCuc":  "Bưu Cục Thị Trấn Kim Nhan-Anh Sơn-Nghệ An"
                                           },
                                           {
                                               "TongCOD":  194.16,
                                               "CODTuNop":  74.01,
                                               "TyLeTuNop":  38.1,
                                               "TenBuuCuc":  "Bưu Cục DT9 Thạch Châu-Lộc Hà-Hà Tĩnh"
                                           },
                                           {
                                               "TongCOD":  96.17,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu cục 219 Khu Phố 1-Như Thanh-Thanh Hóa"
                                           },
                                           {
                                               "TongCOD":  192.62,
                                               "CODTuNop":  9.05,
                                               "TyLeTuNop":  4.7,
                                               "TenBuuCuc":  "Bưu Cục Khối 2-Yên Thành-Nghệ An"
                                           },
                                           {
                                               "TongCOD":  361.57,
                                               "CODTuNop":  152.83,
                                               "TyLeTuNop":  42.3,
                                               "TenBuuCuc":  "Bưu Cục Khu Phố 1-TP.Thanh Hóa-Thanh Hóa"
                                           },
                                           {
                                               "TongCOD":  154.26,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Xóm 8 Quỳnh Hậu-Quỳnh Lưu-Nghệ An"
                                           },
                                           {
                                               "TongCOD":  124.92,
                                               "CODTuNop":  62.37,
                                               "TyLeTuNop":  49.9,
                                               "TenBuuCuc":  "Bưu Cục 242 Lê Lợi-Hương Sơn-Hà Tĩnh"
                                           },
                                           {
                                               "TongCOD":  167.66,
                                               "CODTuNop":  98.04,
                                               "TyLeTuNop":  58.5,
                                               "TenBuuCuc":  "Bưu Cục 100 Lê Lợi-Thường Xuân-Thanh Hóa"
                                           },
                                           {
                                               "TongCOD":  257.11,
                                               "CODTuNop":  75.14,
                                               "TyLeTuNop":  29.2,
                                               "TenBuuCuc":  "Bưu Cục Xóm Xuân Đình-Diễn Hồng-Diễn Châu-Nghệ An"
                                           },
                                           {
                                               "TongCOD":  135.7,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Khối Quy Chính-Nam Đàn-Nghệ An"
                                           },
                                           {
                                               "TongCOD":  112.85,
                                               "CODTuNop":  70.93,
                                               "TyLeTuNop":  62.9,
                                               "TenBuuCuc":  "Bưu Cục 296 Trần Phú-Đức Thuận-Hồng Lĩnh-Hà Tĩnh"
                                           },
                                           {
                                               "TongCOD":  217.68,
                                               "CODTuNop":  124.54,
                                               "TyLeTuNop":  57.2,
                                               "TenBuuCuc":  "Bưu Cục Số 25 Khu Phú Cường-Hậu Lộc-Thanh Hóa"
                                           },
                                           {
                                               "TongCOD":  300.31,
                                               "CODTuNop":  184.59,
                                               "TyLeTuNop":  61.5,
                                               "TenBuuCuc":  "Bưu Cục Phố Trung Sơn-Xã Hoằng Hóa-Thanh Hóa"
                                           },
                                           {
                                               "TongCOD":  233.01,
                                               "CODTuNop":  71.66,
                                               "TyLeTuNop":  30.8,
                                               "TenBuuCuc":  "Bưu Cục Khối 5-Quỳnh Xuân- Tx.Hoàng Mai-Nghệ An"
                                           },
                                           {
                                               "TongCOD":  190.82,
                                               "CODTuNop":  52.11,
                                               "TyLeTuNop":  27.3,
                                               "TenBuuCuc":  "Bưu cục 794 Quang Trung-Thanh Hóa-Thanh Hóa"
                                           },
                                           {
                                               "TongCOD":  76.68,
                                               "CODTuNop":  21.83,
                                               "TyLeTuNop":  28.5,
                                               "TenBuuCuc":  "Bưu Cục Xóm 15 Xã Nghi Lâm-Nghi Lộc-Nghệ An"
                                           },
                                           {
                                               "TongCOD":  361.9,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Xóm 5-Quỳnh Giang-Quỳnh Lưu-Nghệ An"
                                           },
                                           {
                                               "TongCOD":  187.2,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Khối Hợp Xuân-Quỳ Hợp-Nghệ An"
                                           },
                                           {
                                               "TongCOD":  109.33,
                                               "CODTuNop":  10.72,
                                               "TyLeTuNop":  9.8,
                                               "TenBuuCuc":  "Bưu Cục Khối 3-Tân Kỳ-Nghệ An"
                                           },
                                           {
                                               "TongCOD":  159.7,
                                               "CODTuNop":  38.23,
                                               "TyLeTuNop":  23.9,
                                               "TenBuuCuc":  "Bưu Cục 74 Hai Sáu Tháng Ba-Thành Sen-Hà Tĩnh"
                                           },
                                           {
                                               "TongCOD":  82.1,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Đường K34-Vinh Lộc-Nghệ An"
                                           },
                                           {
                                               "TongCOD":  108.56,
                                               "CODTuNop":  64.97,
                                               "TyLeTuNop":  59.8,
                                               "TenBuuCuc":  "Bưu Cục 292A Lý Tự Trọng-Thạch Hà-Hà Tĩnh"
                                           },
                                           {
                                               "TongCOD":  68.66,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Mỹ Sơn-Đô Lương-Nghệ An"
                                           },
                                           {
                                               "TongCOD":  201.53,
                                               "CODTuNop":  141.36,
                                               "TyLeTuNop":  70.1,
                                               "TenBuuCuc":  "Bưu Cục Nghĩa Đàn-Nghệ An"
                                           },
                                           {
                                               "TongCOD":  178.37,
                                               "CODTuNop":  91.24,
                                               "TyLeTuNop":  51.2,
                                               "TenBuuCuc":  "Bưu Cục 18 TDP Phong Ý-Cẩm Thuỷ-Thanh Hoá"
                                           },
                                           {
                                               "TongCOD":  80.2,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Đường Tỉnh 542C-Xã Hưng Nguyên Nam-Nghệ An"
                                           },
                                           {
                                               "TongCOD":  138.78,
                                               "CODTuNop":  38.08,
                                               "TyLeTuNop":  27.4,
                                               "TenBuuCuc":  "Bưu Cục Đường 10 tháng 6 Khu 4-Yên Định-Thanh Hóa"
                                           },
                                           {
                                               "TongCOD":  149.65,
                                               "CODTuNop":  39.18,
                                               "TyLeTuNop":  26.2,
                                               "TenBuuCuc":  "Bưu Cục TDP1-Thị Trấn Hương Khê-Hương Khê-Hà Tĩnh"
                                           },
                                           {
                                               "TongCOD":  197.96,
                                               "CODTuNop":  121.39,
                                               "TyLeTuNop":  61.3,
                                               "TenBuuCuc":  "Bưu Cục QL8B Nguyễn Nhiễm-Xã Nghi Xuân-Hà Tĩnh"
                                           },
                                           {
                                               "TongCOD":  181.95,
                                               "CODTuNop":  53.96,
                                               "TyLeTuNop":  29.7,
                                               "TenBuuCuc":  "Bưu Cục 104B đường Quế Sơn-Thọ Xuân-Thanh Hóa"
                                           },
                                           {
                                               "TongCOD":  73.11,
                                               "CODTuNop":  42.84,
                                               "TyLeTuNop":  58.6,
                                               "TenBuuCuc":  "Bưu Cục Xã Tân Phú-Tân Kỳ-Nghệ An"
                                           },
                                           {
                                               "TongCOD":  73.18,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Xóm 3-Xã Hạnh Lâm-Nghệ An"
                                           },
                                           {
                                               "TongCOD":  16.41,
                                               "CODTuNop":  8.61,
                                               "TyLeTuNop":  52.5,
                                               "TenBuuCuc":  "Điểm xử lý hàng 415 Hà Huy Tập-Xã Cẩm Xuyên-Hà Tĩnh"
                                           },
                                           {
                                               "TongCOD":  125.1,
                                               "CODTuNop":  85.24,
                                               "TyLeTuNop":  68.1,
                                               "TenBuuCuc":  "Bưu Cục Khối Tân Minh-Nghĩa Đàn-Nghệ An"
                                           },
                                           {
                                               "TongCOD":  165.8,
                                               "CODTuNop":  46.14,
                                               "TyLeTuNop":  27.8,
                                               "TenBuuCuc":  "Bưu Cục Thôn Khoa Trường-Tĩnh Gia-Thanh Hóa"
                                           },
                                           {
                                               "TongCOD":  291.88,
                                               "CODTuNop":  32.06,
                                               "TyLeTuNop":  11,
                                               "TenBuuCuc":  "Bưu Cục 558 đường Tố Hữu-Quảng Xương-Thanh Hóa"
                                           },
                                           {
                                               "TongCOD":  239.19,
                                               "CODTuNop":  108.2,
                                               "TyLeTuNop":  45.2,
                                               "TenBuuCuc":  "Bưu Cục 638 Phố Cống-Ngọc Lặc-Thanh Hóa"
                                           },
                                           {
                                               "TongCOD":  229.26,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Tĩnh Gia-Thanh Hóa"
                                           },
                                           {
                                               "TongCOD":  133.87,
                                               "CODTuNop":  67.69,
                                               "TyLeTuNop":  50.6,
                                               "TenBuuCuc":  "Bưu Cục 168 đường 1/9-Thanh Chương-Nghệ An"
                                           },
                                           {
                                               "TongCOD":  179.87,
                                               "CODTuNop":  38.44,
                                               "TyLeTuNop":  21.4,
                                               "TenBuuCuc":  "Bưu Cục Cầu Mưng-Hưng Nguyên-Nghệ An"
                                           },
                                           {
                                               "TongCOD":  197.16,
                                               "CODTuNop":  7.62,
                                               "TyLeTuNop":  3.9,
                                               "TenBuuCuc":  "Bưu Cục Xóm 6 Diễn Thịnh-Diễn Châu-Nghệ An"
                                           },
                                           {
                                               "TongCOD":  158.81,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 37 Nguyễn Hoàng QL1A-Xã Hà Trung-Thanh Hóa"
                                           },
                                           {
                                               "TongCOD":  156.26,
                                               "CODTuNop":  1.16,
                                               "TyLeTuNop":  0.7,
                                               "TenBuuCuc":  "Bưu Cục Thị Trấn Quán Hành-Nghị Lộc-Nghệ An"
                                           },
                                           {
                                               "TongCOD":  162.86,
                                               "CODTuNop":  162.86,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 265 Tiểu Khu 2 QL45-Thiệu Hóa-Thanh Hóa"
                                           },
                                           {
                                               "TongCOD":  198.44,
                                               "CODTuNop":  93.16,
                                               "TyLeTuNop":  46.9,
                                               "TenBuuCuc":  "Bưu Cục Khu 1 Kim Tân-Thạch Thành-Thanh Hoá"
                                           },
                                           {
                                               "TongCOD":  384.91,
                                               "CODTuNop":  153.22,
                                               "TyLeTuNop":  39.8,
                                               "TenBuuCuc":  "Bưu Cục 162 Đường Thành Thái-Đông Thọ-Thanh Hóa"
                                           },
                                           {
                                               "TongCOD":  171.39,
                                               "CODTuNop":  150.33,
                                               "TyLeTuNop":  87.7,
                                               "TenBuuCuc":  "Bưu Cục 5 Võ Thúc Đồng-Trường Vinh-Nghệ An"
                                           },
                                           {
                                               "TongCOD":  210.03,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Thôn Đô Trang-Dân Lực-Triệu Sơn-Thanh Hóa"
                                           },
                                           {
                                               "TongCOD":  192.63,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Khối 9-Đô Lương-Nghệ An"
                                           },
                                           {
                                               "TongCOD":  94.51,
                                               "CODTuNop":  19.22,
                                               "TyLeTuNop":  20.3,
                                               "TenBuuCuc":  "Điểm xử lý hàng Huyện Vũ Quang-Hà Tĩnh"
                                           },
                                           {
                                               "TongCOD":  296.95,
                                               "CODTuNop":  275.6,
                                               "TyLeTuNop":  92.8,
                                               "TenBuuCuc":  "Bưu Cục Hưng Phúc-Vinh-Nghệ An"
                                           },
                                           {
                                               "TongCOD":  191.63,
                                               "CODTuNop":  119.21,
                                               "TyLeTuNop":  62.2,
                                               "TenBuuCuc":  "Bưu Cục 509 Trần Phú-TP.Hà Tĩnh"
                                           },
                                           {
                                               "TongCOD":  123.3,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 124 Sào Nam-Nghi Thu-Cửa Lò-Nghệ An"
                                           },
                                           {
                                               "TongCOD":  111.47,
                                               "CODTuNop":  38.43,
                                               "TyLeTuNop":  34.5,
                                               "TenBuuCuc":  "Bưu Cục QL1A Đồng Tiến-Kỳ Đồng-Kỳ Anh-Hà Tĩnh"
                                           },
                                           {
                                               "TongCOD":  201.48,
                                               "CODTuNop":  74.87,
                                               "TyLeTuNop":  37.2,
                                               "TenBuuCuc":  "Bưu Cục Xóm Quỳ Lăng-Xã Giai Lạc-Nghệ An"
                                           },
                                           {
                                               "TongCOD":  240.15,
                                               "CODTuNop":  126.58,
                                               "TyLeTuNop":  52.7,
                                               "TenBuuCuc":  "Bưu Cục Quang Trung-Vinh-Nghệ An"
                                           },
                                           {
                                               "TongCOD":  240.78,
                                               "CODTuNop":  142.7,
                                               "TyLeTuNop":  59.3,
                                               "TenBuuCuc":  "Bưu Cục Phường Hà Huy Tập-Vinh-Nghệ An"
                                           }
                                       ]
                    },
                    {
                        "TyLeTuNop":  37.7,
                        "TenVung":  "DBB",
                        "TongCOD":  4950.54,
                        "Trend":  6.4,
                        "CODTuNop":  1868,
                        "BuuCucList":  [
                                           {
                                               "TongCOD":  66.13,
                                               "CODTuNop":  66.13,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 25 Trần Hưng Đạo-Xã Yên Minh-Hà Giang"
                                           },
                                           {
                                               "TongCOD":  58.34,
                                               "CODTuNop":  5.08,
                                               "TyLeTuNop":  8.7,
                                               "TenBuuCuc":  "Bưu Cục Xóm 6-Nước Hai-Hòa An-Cao Bằng"
                                           },
                                           {
                                               "TongCOD":  120.68,
                                               "CODTuNop":  56.76,
                                               "TyLeTuNop":  47,
                                               "TenBuuCuc":  "Bưu Cục Chợ TT Nam Sông Công-Xã Đại Phúc-Thái Nguyên"
                                           },
                                           {
                                               "TongCOD":  118.46,
                                               "CODTuNop":  57.01,
                                               "TyLeTuNop":  48.1,
                                               "TenBuuCuc":  "Bưu Cục Số 8 Phan Huy Ích-Tân Hà-Tuyên Quang"
                                           },
                                           {
                                               "TongCOD":  49.99,
                                               "CODTuNop":  11.19,
                                               "TyLeTuNop":  22.4,
                                               "TenBuuCuc":  "Bưu Cục Quốc Lộ 3B-Xã Na Rì-Bắc Kạn"
                                           },
                                           {
                                               "TongCOD":  29.4,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 546 Xóm 2-Xã Bảo Lâm-Cao Bằng"
                                           },
                                           {
                                               "TongCOD":  47.09,
                                               "CODTuNop":  22.75,
                                               "TyLeTuNop":  48.3,
                                               "TenBuuCuc":  "Bưu Cục Đường Trần Phú-Xã Nà Hang-Tuyên Quang"
                                           },
                                           {
                                               "TongCOD":  120.57,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Lõng Khu-Sơn Nam-Sơn Dương-Tuyên Quang"
                                           },
                                           {
                                               "TongCOD":  127.44,
                                               "CODTuNop":  64.04,
                                               "TyLeTuNop":  50.3,
                                               "TenBuuCuc":  "Bưu Cục Số Nhà 30-Mỏ Chè-Sông Công-Thái Nguyên"
                                           },
                                           {
                                               "TongCOD":  138.87,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu cục 545 Lưu Nhân Trú-TP.Thái Nguyên-Thái Nguyên"
                                           },
                                           {
                                               "TongCOD":  0.8,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Điểm Xử Lý Hàng Chân Cầu Mây-Phú Bình-Thái Nguyên"
                                           },
                                           {
                                               "TongCOD":  81.66,
                                               "CODTuNop":  22.63,
                                               "TyLeTuNop":  27.7,
                                               "TenBuuCuc":  "Bưu Cục Cầu Vĩnh Tuy-Xã Vĩnh Tuy-Hà Giang"
                                           },
                                           {
                                               "TongCOD":  33.9,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Xóm 6-Xã Bảo Lạc-Cao Bằng"
                                           },
                                           {
                                               "TongCOD":  63.27,
                                               "CODTuNop":  5.49,
                                               "TyLeTuNop":  8.7,
                                               "TenBuuCuc":  "Bưu Cục 86 Nguyễn Thị Minh Khai-Xã Mèo Vạc-Hà Giang"
                                           },
                                           {
                                               "TongCOD":  81.16,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 85 QL4A-Tân Giang-Cao Bằng"
                                           },
                                           {
                                               "TongCOD":  73.88,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Ngã 4 Phố Hòa Bình-Quảng Hòa-Cao Bằng"
                                           },
                                           {
                                               "TongCOD":  146.89,
                                               "CODTuNop":  90.87,
                                               "TyLeTuNop":  61.9,
                                               "TenBuuCuc":  "Bưu Cục 65 Nguyễn Văn Thoát-TP.Bắc Kạn"
                                           },
                                           {
                                               "TongCOD":  7.89,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Điểm xử lý hàng Xín Mần 2-Hà Giang"
                                           },
                                           {
                                               "TongCOD":  125.02,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Xã Quyết Thắng-TP.Thái Nguyên-Thái Nguyên"
                                           },
                                           {
                                               "TongCOD":  87.72,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Xóm Cây Bòng-La Hiên-Võ Nhai-Thái Nguyên"
                                           },
                                           {
                                               "TongCOD":  126.35,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 786 Dương Tự Minh-Phường Quan Triều-Thái Nguyên"
                                           },
                                           {
                                               "TongCOD":  125.03,
                                               "CODTuNop":  17.32,
                                               "TyLeTuNop":  13.9,
                                               "TenBuuCuc":  "Bưu Cục Số Nhà 882 Tổ 6-Việt Quang-Bắc Quang-Hà Giang"
                                           },
                                           {
                                               "TongCOD":  208.43,
                                               "CODTuNop":  208.36,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 243 Phan Đình Phùng-Phan Đình Phùng-Thái Nguyên"
                                           },
                                           {
                                               "TongCOD":  50.76,
                                               "CODTuNop":  13.91,
                                               "TyLeTuNop":  27.4,
                                               "TenBuuCuc":  "Điểm Xử Lý Hàng Xóm 6 TT Nước Hai-Xã Hòa An-Cao Bằng"
                                           },
                                           {
                                               "TongCOD":  203.87,
                                               "CODTuNop":  86.84,
                                               "TyLeTuNop":  42.6,
                                               "TenBuuCuc":  "Bưu Cục Tôn Đức Thắng-Phổ Yên-Thái Nguyên"
                                           },
                                           {
                                               "TongCOD":  38.57,
                                               "CODTuNop":  7.36,
                                               "TyLeTuNop":  19.1,
                                               "TenBuuCuc":  "Bưu Cục Tổ 4-Thị trấn Nguyên Bình-Cao Bằng"
                                           },
                                           {
                                               "TongCOD":  51.97,
                                               "CODTuNop":  51.04,
                                               "TyLeTuNop":  98.2,
                                               "TenBuuCuc":  "Bưu Cục Đường 20/8-Phương Độ-TP.Hà Giang-Hà Giang"
                                           },
                                           {
                                               "TongCOD":  21.07,
                                               "CODTuNop":  21.07,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục K31 Thôn Nà Sài-Xã Minh Ngọc-Hà Giang"
                                           },
                                           {
                                               "TongCOD":  128.76,
                                               "CODTuNop":  87.7,
                                               "TyLeTuNop":  68.1,
                                               "TenBuuCuc":  "Bưu Cục TDP Trung Kiên-Định Hóa-Thái Nguyên"
                                           },
                                           {
                                               "TongCOD":  45.91,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Thôn Làng Chùa-Xã Lâm Bình-Tuyên Quang"
                                           },
                                           {
                                               "TongCOD":  136.86,
                                               "CODTuNop":  26.08,
                                               "TyLeTuNop":  19.1,
                                               "TenBuuCuc":  "Bưu Cục Thôn Tân Hòa-Sơn Dương-Tuyên Quang"
                                           },
                                           {
                                               "TongCOD":  149.39,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 305 Vĩnh Khang-Chiêm Hóa-Tuyên Quang"
                                           },
                                           {
                                               "TongCOD":  49.91,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Xóm Nà Ến-Xã Hạ Lang-Cao Bằng"
                                           },
                                           {
                                               "TongCOD":  110.6,
                                               "CODTuNop":  60.16,
                                               "TyLeTuNop":  54.4,
                                               "TenBuuCuc":  "Bưu Cục Lê Quý Đôn-Đồng Hỷ-Thái Nguyên"
                                           },
                                           {
                                               "TongCOD":  36.77,
                                               "CODTuNop":  11.08,
                                               "TyLeTuNop":  30.1,
                                               "TenBuuCuc":  "Điểm xử lý hàng Huyện Xín Mần-Hà Giang"
                                           },
                                           {
                                               "TongCOD":  43.48,
                                               "CODTuNop":  5.69,
                                               "TyLeTuNop":  13.1,
                                               "TenBuuCuc":  "Bưu Cục Hoàng Quốc Việt-Tân Hương-Phổ Yên-Thái Nguyên"
                                           },
                                           {
                                               "TongCOD":  168.93,
                                               "CODTuNop":  111.04,
                                               "TyLeTuNop":  65.7,
                                               "TenBuuCuc":  "Bưu Cục TDP Lê Hồng Phong-Phú Lương-Thái Nguyên"
                                           },
                                           {
                                               "TongCOD":  26.13,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Xóm Chang Khuyên-Xã Đông Khê-Cao Bằng"
                                           },
                                           {
                                               "TongCOD":  56.74,
                                               "CODTuNop":  32.64,
                                               "TyLeTuNop":  57.5,
                                               "TenBuuCuc":  "Bưu Cục Đường 261-Bắc Sơn-Phổ Yên-Thái Nguyên"
                                           },
                                           {
                                               "TongCOD":  59.52,
                                               "CODTuNop":  24.81,
                                               "TyLeTuNop":  41.7,
                                               "TenBuuCuc":  "Bưu Cục Tổ 2 Quốc Lộ 4C-Xã Quản Bạ-Hà Giang"
                                           },
                                           {
                                               "TongCOD":  43.92,
                                               "CODTuNop":  43.92,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Tiểu Khu 2-Nà Phặc-Ngân Sơn-Bắc Kạn"
                                           },
                                           {
                                               "TongCOD":  73.88,
                                               "CODTuNop":  68.95,
                                               "TyLeTuNop":  93.3,
                                               "TenBuuCuc":  "Điểm Xử Lý Hàng TT Nà Phặc-Xã Nà Phặc-Bắc Kạn"
                                           },
                                           {
                                               "TongCOD":  55.93,
                                               "CODTuNop":  55.93,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Tổ 11 Đồng Tâm-Chợ Mới-Bắc Kạn"
                                           },
                                           {
                                               "TongCOD":  84.91,
                                               "CODTuNop":  19.43,
                                               "TyLeTuNop":  22.9,
                                               "TenBuuCuc":  "Bưu Cục Xóm 1 Đường 264-Xã Phú Thịnh-Thái Nguyên"
                                           },
                                           {
                                               "TongCOD":  38.22,
                                               "CODTuNop":  11.58,
                                               "TyLeTuNop":  30.3,
                                               "TenBuuCuc":  "Bưu Cục Thị Trấn Vinh Quang-Hoàng Su Phì-Hà Giang"
                                           },
                                           {
                                               "TongCOD":  91.79,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Thôn Đoàn Kết QL4C-Xã Đồng Văn-Hà Giang"
                                           },
                                           {
                                               "TongCOD":  65.28,
                                               "CODTuNop":  29.1,
                                               "TyLeTuNop":  44.6,
                                               "TenBuuCuc":  "Bưu Cục Xóm Đoàn Kết-Nam Hòa-Đồng Hỷ-Tỉnh Thái Nguyên"
                                           },
                                           {
                                               "TongCOD":  24.94,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 92 Đường 27-Quang Bình-Hà Giang"
                                           },
                                           {
                                               "TongCOD":  56.39,
                                               "CODTuNop":  25.84,
                                               "TyLeTuNop":  45.8,
                                               "TenBuuCuc":  "Bưu Cục Tổ 5 TT Bằng Lũng-Xã Chợ Đồn-Bắc Kạn"
                                           },
                                           {
                                               "TongCOD":  77.17,
                                               "CODTuNop":  14.71,
                                               "TyLeTuNop":  19.1,
                                               "TenBuuCuc":  "Bưu Cục Quốc Lộ 2-Xã Yên Sơn-Tuyên Quang"
                                           },
                                           {
                                               "TongCOD":  67.32,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 134 Xóm 2-Xã Trùng Khánh-Cao Bằng"
                                           },
                                           {
                                               "TongCOD":  329.56,
                                               "CODTuNop":  182.9,
                                               "TyLeTuNop":  55.5,
                                               "TenBuuCuc":  "Bưu Cục Số 10 Phan Huy Ích-Tân Hà-Tuyên Quang"
                                           },
                                           {
                                               "TongCOD":  120.86,
                                               "CODTuNop":  116.32,
                                               "TyLeTuNop":  96.2,
                                               "TenBuuCuc":  "Bưu Cục Phùng Hưng-Hà Giang 2-Hà Giang"
                                           },
                                           {
                                               "TongCOD":  162,
                                               "CODTuNop":  7.82,
                                               "TyLeTuNop":  4.8,
                                               "TenBuuCuc":  "Bưu Cục Tổ 1 TT Hương Sơn-Xã Phú Bình-Thái Nguyên"
                                           },
                                           {
                                               "TongCOD":  144.08,
                                               "CODTuNop":  57.36,
                                               "TyLeTuNop":  39.8,
                                               "TenBuuCuc":  "Bưu Cục Hàm Yên-Tuyên Quang"
                                           },
                                           {
                                               "TongCOD":  96.08,
                                               "CODTuNop":  67.1,
                                               "TyLeTuNop":  69.8,
                                               "TenBuuCuc":  "Bưu Cục Thị trấn Vị Xuyên-Hà Giang"
                                           }
                                       ]
                    },
                    {
                        "TyLeTuNop":  26.3,
                        "TenVung":  "ĐCL",
                        "TongCOD":  9594.97,
                        "Trend":  -8,
                        "CODTuNop":  2520.78,
                        "BuuCucList":  [
                                           {
                                               "TongCOD":  310.17,
                                               "CODTuNop":  55.26,
                                               "TyLeTuNop":  17.8,
                                               "TenBuuCuc":  "Bưu Cục 206 A Đoàn Thị Nghiệp-Mỹ Tho-Tiền Giang 01"
                                           },
                                           {
                                               "TongCOD":  18.26,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Tỉnh Lộ 868-Phú Cường-Cai Lậy-Tiền Giang"
                                           },
                                           {
                                               "TongCOD":  157.03,
                                               "CODTuNop":  37.52,
                                               "TyLeTuNop":  23.9,
                                               "TenBuuCuc":  "Bưu Cục 992 Đường Huyện 35-Vĩnh Kim-Châu Thành-Tiền Giang"
                                           },
                                           {
                                               "TongCOD":  154.58,
                                               "CODTuNop":  47.66,
                                               "TyLeTuNop":  30.8,
                                               "TenBuuCuc":  "Bưu Cục 3/2 Thị Trấn Châu Thành-Bến Tre"
                                           },
                                           {
                                               "TongCOD":  166.24,
                                               "CODTuNop":  10.3,
                                               "TyLeTuNop":  6.2,
                                               "TenBuuCuc":  "Bưu Cục QL57-Bình Thạnh-Thạnh Phú-Bến Tre"
                                           },
                                           {
                                               "TongCOD":  104.17,
                                               "CODTuNop":  19.43,
                                               "TyLeTuNop":  18.7,
                                               "TenBuuCuc":  "Bưu Cục Quốc Lộ 54-Xã Lai Vung-Đồng Tháp"
                                           },
                                           {
                                               "TongCOD":  285.6,
                                               "CODTuNop":  37.48,
                                               "TyLeTuNop":  13.1,
                                               "TenBuuCuc":  "Bưu Cục KP2-Thị Trấn Ba Tri-Bến Tre"
                                           },
                                           {
                                               "TongCOD":  103.11,
                                               "CODTuNop":  35.97,
                                               "TyLeTuNop":  34.9,
                                               "TenBuuCuc":  "Bưu Cục Phước Long-Lộc Hòa-Long Hồ-Vĩnh Long"
                                           },
                                           {
                                               "TongCOD":  3.71,
                                               "CODTuNop":  3.71,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Tỉnh Lộ DT848-Xã Mỹ An Hưng-Đồng Tháp"
                                           },
                                           {
                                               "TongCOD":  202.65,
                                               "CODTuNop":  185.5,
                                               "TyLeTuNop":  91.5,
                                               "TenBuuCuc":  "Bưu Cục Khóm 3 Trần Hưng Đạo-Xã Tháp Mười-Đồng Tháp"
                                           },
                                           {
                                               "TongCOD":  186.47,
                                               "CODTuNop":  29.17,
                                               "TyLeTuNop":  15.6,
                                               "TenBuuCuc":  "Bưu Cục QL 30-Thanh Bình-Đồng Tháp"
                                           },
                                           {
                                               "TongCOD":  185.72,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Quốc Lộ 53-Xã Trung Thành-Vĩnh Long"
                                           },
                                           {
                                               "TongCOD":  93.75,
                                               "CODTuNop":  34.24,
                                               "TyLeTuNop":  36.5,
                                               "TenBuuCuc":  "Bưu Cục QL80-Xã Hòa Long-Đồng Tháp"
                                           },
                                           {
                                               "TongCOD":  308.32,
                                               "CODTuNop":  126.32,
                                               "TyLeTuNop":  41,
                                               "TenBuuCuc":  "Bưu Cục Hồng Ngự-Đồng Tháp"
                                           },
                                           {
                                               "TongCOD":  250.39,
                                               "CODTuNop":  79.37,
                                               "TyLeTuNop":  31.7,
                                               "TenBuuCuc":  "Bưu Cục 230 Ấp Bình Hạnh-Long Bình Điền-Chợ Gạo-Tiền Giang"
                                           },
                                           {
                                               "TongCOD":  143.66,
                                               "CODTuNop":  23.87,
                                               "TyLeTuNop":  16.6,
                                               "TenBuuCuc":  "Bưu Cục Tổ 1 Đường 30/4-Trà Ôn-Vĩnh Long"
                                           },
                                           {
                                               "TongCOD":  124.65,
                                               "CODTuNop":  24.78,
                                               "TyLeTuNop":  19.9,
                                               "TenBuuCuc":  "Bưu Cục Khóm Thống Nhất-Xã Cầu Ngang-Trà Vinh"
                                           },
                                           {
                                               "TongCOD":  188.87,
                                               "CODTuNop":  29.51,
                                               "TyLeTuNop":  15.6,
                                               "TenBuuCuc":  "Bưu Cục Xoài Xiêm-Ngãi Xuyên-Trà Cú-Trà Vinh"
                                           },
                                           {
                                               "TongCOD":  171.91,
                                               "CODTuNop":  64.1,
                                               "TyLeTuNop":  37.3,
                                               "TenBuuCuc":  "Bưu Cục 57/4 Phan Văn Đáng-Vĩnh Long-Vĩnh Long"
                                           },
                                           {
                                               "TongCOD":  89.81,
                                               "CODTuNop":  17.17,
                                               "TyLeTuNop":  19.1,
                                               "TenBuuCuc":  "Bưu Cục Mỹ Thọ-Trần Quang Khải-Cao Lãnh-Đồng Tháp"
                                           },
                                           {
                                               "TongCOD":  103.76,
                                               "CODTuNop":  3.28,
                                               "TyLeTuNop":  3.2,
                                               "TenBuuCuc":  "Bưu Cục 73 Phó Cơ Điều-Phường Phước Hậu-Vĩnh Long"
                                           },
                                           {
                                               "TongCOD":  54.83,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Quốc Lộ 57B-Xã Lộc Thuận-Bến Tre"
                                           },
                                           {
                                               "TongCOD":  117.16,
                                               "CODTuNop":  34.95,
                                               "TyLeTuNop":  29.8,
                                               "TenBuuCuc":  "Bưu Cục 87 Lê Văn Vị-Cái Vồn-Bình Minh-Vĩnh Long"
                                           },
                                           {
                                               "TongCOD":  199.73,
                                               "CODTuNop":  41.02,
                                               "TyLeTuNop":  20.5,
                                               "TenBuuCuc":  "Bưu Cục QL50 Phường 9-Mỹ Tho-Tiền Giang"
                                           },
                                           {
                                               "TongCOD":  50.59,
                                               "CODTuNop":  17.83,
                                               "TyLeTuNop":  35.2,
                                               "TenBuuCuc":  "Bưu Cục Quốc Lộ 53-Long Hồ-Vĩnh Long"
                                           },
                                           {
                                               "TongCOD":  223.83,
                                               "CODTuNop":  176.6,
                                               "TyLeTuNop":  78.9,
                                               "TenBuuCuc":  "Bưu Cục 888 QL1A-Phường 2-Cai Lậy-Tiền Giang"
                                           },
                                           {
                                               "TongCOD":  131.13,
                                               "CODTuNop":  33.53,
                                               "TyLeTuNop":  25.6,
                                               "TenBuuCuc":  "Bưu Cục Ấp Tân Thuận-Bình Tân-Vĩnh Long"
                                           },
                                           {
                                               "TongCOD":  132.77,
                                               "CODTuNop":  42.97,
                                               "TyLeTuNop":  32.4,
                                               "TenBuuCuc":  "Bưu Cục 07 Lê Lợi-Mang Thít-Vĩnh Long"
                                           },
                                           {
                                               "TongCOD":  120.14,
                                               "CODTuNop":  1.59,
                                               "TyLeTuNop":  1.3,
                                               "TenBuuCuc":  "Bưu Cục QL57 KP3-Thị Trấn Chợ Lách-Bến Tre"
                                           },
                                           {
                                               "TongCOD":  141.85,
                                               "CODTuNop":  29.96,
                                               "TyLeTuNop":  21.1,
                                               "TenBuuCuc":  "Bưu Cục Ấp Cây Hẹ-Phú Cần-Tiểu Cần-Trà Vinh"
                                           },
                                           {
                                               "TongCOD":  106.34,
                                               "CODTuNop":  20.9,
                                               "TyLeTuNop":  19.7,
                                               "TenBuuCuc":  "Bưu Cục Phước Hậu-Phước Mỹ Trung-Bến Tre"
                                           },
                                           {
                                               "TongCOD":  200.09,
                                               "CODTuNop":  37.43,
                                               "TyLeTuNop":  18.7,
                                               "TenBuuCuc":  "Bưu Cục KP3 Phường 3-Gò Công-Tiền Giang"
                                           },
                                           {
                                               "TongCOD":  177.66,
                                               "CODTuNop":  30.56,
                                               "TyLeTuNop":  17.2,
                                               "TenBuuCuc":  "Bưu Cục 79 Đường Tỉnh 871-Xã Gia Thuận-Tiền Giang"
                                           },
                                           {
                                               "TongCOD":  89.68,
                                               "CODTuNop":  25.07,
                                               "TyLeTuNop":  28,
                                               "TenBuuCuc":  "Bưu Cục Ấp 2-Phú Nhuận-Bến Tre"
                                           },
                                           {
                                               "TongCOD":  113.17,
                                               "CODTuNop":  33.88,
                                               "TyLeTuNop":  29.9,
                                               "TenBuuCuc":  "Bưu Cục Khóm 3 TT Càng Long-Xã Càng Long-Trà Vinh"
                                           },
                                           {
                                               "TongCOD":  198.13,
                                               "CODTuNop":  73.11,
                                               "TyLeTuNop":  36.9,
                                               "TenBuuCuc":  "Bưu Cục 88 Ấp Thân Hòa-Thân Cửu Nghĩa-Châu Thành-Tiền Giang"
                                           },
                                           {
                                               "TongCOD":  178.57,
                                               "CODTuNop":  26.45,
                                               "TyLeTuNop":  14.8,
                                               "TenBuuCuc":  "Bưu Cục Khóm 4 TT Châu Thành-Xã Châu Thành-Trà Vinh"
                                           },
                                           {
                                               "TongCOD":  183.96,
                                               "CODTuNop":  44.48,
                                               "TyLeTuNop":  24.2,
                                               "TenBuuCuc":  "Bưu Cục Tân Nhuận Đông-Châu Thành-Đồng Tháp"
                                           },
                                           {
                                               "TongCOD":  354.99,
                                               "CODTuNop":  31,
                                               "TyLeTuNop":  8.7,
                                               "TenBuuCuc":  "Bưu Cục Sơn Thông-Phường 9-Tp.Trà Vinh-Trà Vinh"
                                           },
                                           {
                                               "TongCOD":  134.03,
                                               "CODTuNop":  52.03,
                                               "TyLeTuNop":  38.8,
                                               "TenBuuCuc":  "Bưu Cục Khu phố 1-Tân Phước-Tiền Giang"
                                           },
                                           {
                                               "TongCOD":  104.73,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 86 Tiên Thuỷ-Châu Thành-Bến Tre"
                                           },
                                           {
                                               "TongCOD":  150.3,
                                               "CODTuNop":  29.39,
                                               "TyLeTuNop":  19.6,
                                               "TenBuuCuc":  "Bưu Cục Tam Nông-Đồng Tháp"
                                           },
                                           {
                                               "TongCOD":  204.21,
                                               "CODTuNop":  204.21,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Duyên Hải-Trà Vinh"
                                           },
                                           {
                                               "TongCOD":  67.34,
                                               "CODTuNop":  17.22,
                                               "TyLeTuNop":  25.6,
                                               "TenBuuCuc":  "Bưu Cục DT846 Phương Trà-Xã Ba Sao-Đồng Tháp"
                                           },
                                           {
                                               "TongCOD":  150.49,
                                               "CODTuNop":  12.82,
                                               "TyLeTuNop":  8.5,
                                               "TenBuuCuc":  "Bưu Cục 355 Võ Tấn Đức-Tam Bình-Vĩnh Long"
                                           },
                                           {
                                               "TongCOD":  123.46,
                                               "CODTuNop":  18.8,
                                               "TyLeTuNop":  15.2,
                                               "TenBuuCuc":  "Bưu Cục Ấp Hậu Mỹ Phú 1-Cái Bè-Tiền Giang"
                                           },
                                           {
                                               "TongCOD":  97.55,
                                               "CODTuNop":  20.18,
                                               "TyLeTuNop":  20.7,
                                               "TenBuuCuc":  "Bưu Cục Ấp Bình Hiệp A-Xã Lấp Vò-Đồng Tháp"
                                           },
                                           {
                                               "TongCOD":  193.25,
                                               "CODTuNop":  116.5,
                                               "TyLeTuNop":  60.3,
                                               "TenBuuCuc":  "Bưu Cục Đường Tỉnh 861-Xã Mỹ Lợi-Tiền Giang"
                                           },
                                           {
                                               "TongCOD":  199.68,
                                               "CODTuNop":  75.61,
                                               "TyLeTuNop":  37.9,
                                               "TenBuuCuc":  "Bưu Cục Long Trung-Cai Lậy-Tiền Giang"
                                           },
                                           {
                                               "TongCOD":  160.34,
                                               "CODTuNop":  28.92,
                                               "TyLeTuNop":  18,
                                               "TenBuuCuc":  "Bưu Cục 286 Nguyễn Huệ-Tân Hồng-Đồng Tháp"
                                           },
                                           {
                                               "TongCOD":  280.85,
                                               "CODTuNop":  11.57,
                                               "TyLeTuNop":  4.1,
                                               "TenBuuCuc":  "Bưu Cục QL57-Bình Thới-Bình Đại-Bến Tre"
                                           },
                                           {
                                               "TongCOD":  172.78,
                                               "CODTuNop":  47.14,
                                               "TyLeTuNop":  27.3,
                                               "TenBuuCuc":  "Bưu Cục Võ Nguyên Giáp-Phú Tân-Bến Tre"
                                           },
                                           {
                                               "TongCOD":  72.33,
                                               "CODTuNop":  3.49,
                                               "TyLeTuNop":  4.8,
                                               "TenBuuCuc":  "Bưu Cục Ấp Phú Long-Xã Hưng Khánh Trung-Bến Tre"
                                           },
                                           {
                                               "TongCOD":  188.55,
                                               "CODTuNop":  67.13,
                                               "TyLeTuNop":  35.6,
                                               "TenBuuCuc":  "Bưu Cục Quốc Lộ 50-Gò Công Tây-Tiền Giang"
                                           },
                                           {
                                               "TongCOD":  294.59,
                                               "CODTuNop":  62.74,
                                               "TyLeTuNop":  21.3,
                                               "TenBuuCuc":  "Bưu Cục 12-14 Bà Triệu-Cao Lãnh-Đồng Tháp"
                                           },
                                           {
                                               "TongCOD":  157.38,
                                               "CODTuNop":  27.75,
                                               "TyLeTuNop":  17.6,
                                               "TenBuuCuc":  "Bưu cục Ấp 5A Thị Trấn Giồng Trôm-Giồng Trôm-Bến Tre"
                                           },
                                           {
                                               "TongCOD":  152.04,
                                               "CODTuNop":  17.02,
                                               "TyLeTuNop":  11.2,
                                               "TenBuuCuc":  "Bưu Cục Mỏ Cày Nam-Bến Tre"
                                           },
                                           {
                                               "TongCOD":  80.07,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Quốc Lộ 57-Xã Hương Mỹ-Bến Tre"
                                           },
                                           {
                                               "TongCOD":  154.65,
                                               "CODTuNop":  60.59,
                                               "TyLeTuNop":  39.2,
                                               "TenBuuCuc":  "Bưu Cục QL1A Hòa Khánh-Cái Bè-Tiền Giang"
                                           },
                                           {
                                               "TongCOD":  104.58,
                                               "CODTuNop":  16.72,
                                               "TyLeTuNop":  16,
                                               "TenBuuCuc":  "Bưu Cục Khóm 8-Cầu Kè-Trà Vinh"
                                           },
                                           {
                                               "TongCOD":  224.3,
                                               "CODTuNop":  64.99,
                                               "TyLeTuNop":  29,
                                               "TenBuuCuc":  "Bưu Cục 19 Đường DT848-Sa Đéc-Đồng Tháp"
                                           }
                                       ]
                    },
                    {
                        "TyLeTuNop":  36.2,
                        "TenVung":  "DNB",
                        "TongCOD":  18097.47,
                        "Trend":  1.5,
                        "CODTuNop":  6554.63,
                        "BuuCucList":  [
                                           {
                                               "TongCOD":  273.02,
                                               "CODTuNop":  103.27,
                                               "TyLeTuNop":  37.8,
                                               "TenBuuCuc":  "Bưu Cục 129 Khổng Tử-Long Khánh-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  248.88,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 45 Nhật Tảo-Xã Vĩnh Hưng-Long An"
                                           },
                                           {
                                               "TongCOD":  291.46,
                                               "CODTuNop":  123.57,
                                               "TyLeTuNop":  42.4,
                                               "TenBuuCuc":  "Bưu Cục Bùi Trọng Nghĩa-Trảng Dài-Biên Hòa-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  72.15,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 79A Hưng Đạo Vương-Biên Hòa-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  281.23,
                                               "CODTuNop":  4.82,
                                               "TyLeTuNop":  1.7,
                                               "TenBuuCuc":  "Bưu Cục Đường Ba Mươi Tháng Tư-Trảng Bàng-Tây Ninh"
                                           },
                                           {
                                               "TongCOD":  91.45,
                                               "CODTuNop":  1.33,
                                               "TyLeTuNop":  1.5,
                                               "TenBuuCuc":  "Bưu Cục Cầu Khởi-Dương Minh Châu-Tây Ninh"
                                           },
                                           {
                                               "TongCOD":  211.28,
                                               "CODTuNop":  43.83,
                                               "TyLeTuNop":  20.7,
                                               "TenBuuCuc":  "Bưu Cục Chơn Thành-Bình Phước"
                                           },
                                           {
                                               "TongCOD":  260.43,
                                               "CODTuNop":  106.61,
                                               "TyLeTuNop":  40.9,
                                               "TenBuuCuc":  "Bưu Cục Ấp Bến Đình-Xã Phú Đông-Nhơn Trạch-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  113.25,
                                               "CODTuNop":  1.76,
                                               "TyLeTuNop":  1.6,
                                               "TenBuuCuc":  "Bưu Cục Ấp 1 Đường Tân Châu-Tân Châu-Tây Ninh"
                                           },
                                           {
                                               "TongCOD":  70.85,
                                               "CODTuNop":  31.89,
                                               "TyLeTuNop":  45,
                                               "TenBuuCuc":  "Bưu Cục Minh Hưng-TX.Chơn Thành-Bình Phước"
                                           },
                                           {
                                               "TongCOD":  99.17,
                                               "CODTuNop":  48.07,
                                               "TyLeTuNop":  48.5,
                                               "TenBuuCuc":  "Bưu Cục 478 Đường Tua Hai-Bình Minh-Tây Ninh"
                                           },
                                           {
                                               "TongCOD":  334.17,
                                               "CODTuNop":  110.08,
                                               "TyLeTuNop":  32.9,
                                               "TenBuuCuc":  "Bưu Cục 55/3B Trần Quốc Toản-Phường Trấn Biên-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  56.17,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Thôn Phú Hưng-Phú Riềng-Bình Phước"
                                           },
                                           {
                                               "TongCOD":  118.13,
                                               "CODTuNop":  62.23,
                                               "TyLeTuNop":  52.7,
                                               "TenBuuCuc":  "Bưu Cục Hóa An-Biên Hòa-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  120.03,
                                               "CODTuNop":  39.24,
                                               "TyLeTuNop":  32.7,
                                               "TenBuuCuc":  "Bưu Cục Chánh Hội-Tân Mỹ-Đức Hòa-Long An"
                                           },
                                           {
                                               "TongCOD":  200.66,
                                               "CODTuNop":  200.63,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Tân Tiến-Biên Hòa-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  110.28,
                                               "CODTuNop":  11.61,
                                               "TyLeTuNop":  10.5,
                                               "TenBuuCuc":  "Bưu Cục Xã Đức Liễu-Bù Đăng-Bình Phước"
                                           },
                                           {
                                               "TongCOD":  387.78,
                                               "CODTuNop":  73.55,
                                               "TyLeTuNop":  19,
                                               "TenBuuCuc":  "Bưu Cục Tân Phú-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  206.64,
                                               "CODTuNop":  206.57,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Ấp 4-Long Hoà-Cần Đước-Long An"
                                           },
                                           {
                                               "TongCOD":  162.8,
                                               "CODTuNop":  61.93,
                                               "TyLeTuNop":  38,
                                               "TenBuuCuc":  "Bưu Cục Nguyễn Thị Minh Khai-Thị trấn Tân Phú-Huyện Đồng Phú-Bình Phước"
                                           },
                                           {
                                               "TongCOD":  312.61,
                                               "CODTuNop":  134.3,
                                               "TyLeTuNop":  43,
                                               "TenBuuCuc":  "Bưu Cục Phước Thiền-Nhơn Trạch-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  159.94,
                                               "CODTuNop":  20.12,
                                               "TyLeTuNop":  12.6,
                                               "TenBuuCuc":  "Bưu Cục 305 QL14-Bù Đăng-Bình Phước"
                                           },
                                           {
                                               "TongCOD":  438.46,
                                               "CODTuNop":  438.42,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 45/2b Quốc Lộ 62-Tân An-Long An 01"
                                           },
                                           {
                                               "TongCOD":  114.82,
                                               "CODTuNop":  39.77,
                                               "TyLeTuNop":  34.6,
                                               "TenBuuCuc":  "Bưu Cục KP1 Quốc Lộ 13-Hớn Quản-Bình Phước"
                                           },
                                           {
                                               "TongCOD":  323.57,
                                               "CODTuNop":  27.88,
                                               "TyLeTuNop":  8.6,
                                               "TenBuuCuc":  "Bưu Cục Ấp Ngoài-Xã Phước Lý-Long An"
                                           },
                                           {
                                               "TongCOD":  185.4,
                                               "CODTuNop":  12.22,
                                               "TyLeTuNop":  6.6,
                                               "TenBuuCuc":  "Bưu Cục 1402 Hoàng Lê Kha-Châu Thành-Tây Ninh"
                                           },
                                           {
                                               "TongCOD":  204,
                                               "CODTuNop":  5.05,
                                               "TyLeTuNop":  2.5,
                                               "TenBuuCuc":  "Bưu Cục Ấp Cầu Xây-Xã Thủ Thừa-Long An"
                                           },
                                           {
                                               "TongCOD":  139.36,
                                               "CODTuNop":  62.46,
                                               "TyLeTuNop":  44.8,
                                               "TenBuuCuc":  "Bưu Cục 56 Dương Văn Dương-Xã Tân Thạnh-Long An"
                                           },
                                           {
                                               "TongCOD":  143.66,
                                               "CODTuNop":  16.48,
                                               "TyLeTuNop":  11.5,
                                               "TenBuuCuc":  "Bưu Cục Tổ 4 KP7 767- Vĩnh Cửu-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  191.05,
                                               "CODTuNop":  71.07,
                                               "TyLeTuNop":  37.2,
                                               "TenBuuCuc":  "Bưu Cục 156 Quốc Lộ 22-Gò Dầu-Tây Ninh"
                                           },
                                           {
                                               "TongCOD":  140.67,
                                               "CODTuNop":  70.27,
                                               "TyLeTuNop":  50,
                                               "TenBuuCuc":  "Bưu Cục 3/3A Huỳnh Văn Nghệ-Bửu Long-Biên Hòa-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  284.88,
                                               "CODTuNop":  182.43,
                                               "TyLeTuNop":  64,
                                               "TenBuuCuc":  "Bưu Cục Đường 25A3-Xã Đức Hòa Hạ-Long An"
                                           },
                                           {
                                               "TongCOD":  194.37,
                                               "CODTuNop":  0.26,
                                               "TyLeTuNop":  0.1,
                                               "TenBuuCuc":  "Bưu Cục 4B Khu B-Xã Hậu Nghĩa-Long An"
                                           },
                                           {
                                               "TongCOD":  316.16,
                                               "CODTuNop":  112.73,
                                               "TyLeTuNop":  35.7,
                                               "TenBuuCuc":  "Bưu Cục 322 Trần Phú-Xuân Lộc-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  153.87,
                                               "CODTuNop":  10.72,
                                               "TyLeTuNop":  7,
                                               "TenBuuCuc":  "Bưu Cục 68 Quốc Lộ 1A-Trảng Bom-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  149.01,
                                               "CODTuNop":  27.67,
                                               "TyLeTuNop":  18.6,
                                               "TenBuuCuc":  "Bưu Cục 146 Nguyễn Văn Cừ-Xã Long Thành-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  205.13,
                                               "CODTuNop":  95.82,
                                               "TyLeTuNop":  46.7,
                                               "TenBuuCuc":  "Bưu Cục 295 Nguyễn Văn Linh-Long Hoa-Tây Ninh"
                                           },
                                           {
                                               "TongCOD":  163.36,
                                               "CODTuNop":  163.36,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Đường Tỉnh 827-Tầm Vu-Châu Thành-Long An"
                                           },
                                           {
                                               "TongCOD":  244.29,
                                               "CODTuNop":  124.69,
                                               "TyLeTuNop":  51,
                                               "TenBuuCuc":  "Bưu Cục 608 Quốc Lộ 20-Định Quán-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  112.94,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 59A Nguyễn Văn Tạo-Tân An-Long An"
                                           },
                                           {
                                               "TongCOD":  144.17,
                                               "CODTuNop":  144.13,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Cần Đước-Long An"
                                           },
                                           {
                                               "TongCOD":  280.04,
                                               "CODTuNop":  59.17,
                                               "TyLeTuNop":  21.1,
                                               "TenBuuCuc":  "Bưu Cục 219 Quốc Lộ 62-Kiến Tường-Long An"
                                           },
                                           {
                                               "TongCOD":  237.35,
                                               "CODTuNop":  34.15,
                                               "TyLeTuNop":  14.4,
                                               "TenBuuCuc":  "Bưu Cục 58 Trần Văn Trà-Tân Châu-Tây Ninh"
                                           },
                                           {
                                               "TongCOD":  138.47,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Ấp 1 Nguyễn Văn Dương-Xã Mỹ Hạnh-Long An"
                                           },
                                           {
                                               "TongCOD":  217.14,
                                               "CODTuNop":  104.45,
                                               "TyLeTuNop":  48.1,
                                               "TenBuuCuc":  "Bưu Cục Xã Long Giao-Cẩm Mỹ-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  270.38,
                                               "CODTuNop":  131.38,
                                               "TyLeTuNop":  48.6,
                                               "TenBuuCuc":  "Bưu Cục 110 Nguyễn Huệ-Trảng Bom-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  130.73,
                                               "CODTuNop":  20.31,
                                               "TyLeTuNop":  15.5,
                                               "TenBuuCuc":  "Bưu Cục 131 Trần Phú-Ninh Trung-Ninh Sơn-Tây Ninh"
                                           },
                                           {
                                               "TongCOD":  183.79,
                                               "CODTuNop":  93.39,
                                               "TyLeTuNop":  50.8,
                                               "TenBuuCuc":  "Bưu Cục Tam Phước-Biên Hòa-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  180.37,
                                               "CODTuNop":  82.56,
                                               "TyLeTuNop":  45.8,
                                               "TenBuuCuc":  "Bưu Cục Bù Nho-Phú Riềng-Bình Phước"
                                           },
                                           {
                                               "TongCOD":  180.49,
                                               "CODTuNop":  139.74,
                                               "TyLeTuNop":  77.4,
                                               "TenBuuCuc":  "Bưu Cục Ấp 8-Xuân Tây-Cẩm Mỹ-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  142.49,
                                               "CODTuNop":  136.73,
                                               "TyLeTuNop":  96,
                                               "TenBuuCuc":  "Bưu Cục Khu Phố 2-Thạnh Hoá-Long An"
                                           },
                                           {
                                               "TongCOD":  157.04,
                                               "CODTuNop":  97.66,
                                               "TyLeTuNop":  62.2,
                                               "TenBuuCuc":  "Bưu Cục Đường Võ Nguyên Giáp-Biên Hòa-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  231.04,
                                               "CODTuNop":  30.07,
                                               "TyLeTuNop":  13,
                                               "TenBuuCuc":  "Bưu Cục 15 Đồng Khởi-Vĩnh Cửu-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  189.63,
                                               "CODTuNop":  119.97,
                                               "TyLeTuNop":  63.3,
                                               "TenBuuCuc":  "Bưu Cục Lý Thường Kiệt-Xã An Phước-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  118.97,
                                               "CODTuNop":  118.97,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 24 Khu Phố Thanh Bình-Bù Đốp-Bình Phước"
                                           },
                                           {
                                               "TongCOD":  164.15,
                                               "CODTuNop":  25.6,
                                               "TyLeTuNop":  15.6,
                                               "TenBuuCuc":  "Bưu Cục 107 Nguyễn Chí Thanh-Tân Biên-Tây Ninh"
                                           },
                                           {
                                               "TongCOD":  253.37,
                                               "CODTuNop":  124.32,
                                               "TyLeTuNop":  49.1,
                                               "TenBuuCuc":  "Bưu Cục Đường Hùng Vương-Xã Phước An-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  193.49,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 108 Khu Phố 2-Đức Huệ-Long An"
                                           },
                                           {
                                               "TongCOD":  242.68,
                                               "CODTuNop":  31.47,
                                               "TyLeTuNop":  13,
                                               "TenBuuCuc":  "Bưu Cục Agtex Long Bình-Biên Hoà-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  209.88,
                                               "CODTuNop":  8.98,
                                               "TyLeTuNop":  4.3,
                                               "TenBuuCuc":  "Bưu Cục Ấp Lủy-Phước Lại-Cần Giuộc-Long An"
                                           },
                                           {
                                               "TongCOD":  267.6,
                                               "CODTuNop":  110.38,
                                               "TyLeTuNop":  41.2,
                                               "TenBuuCuc":  "Bưu Cục Quốc Lộ 1A-Bắc Sơn-Trảng Bom-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  281.31,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Long Khánh-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  97.58,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Xã 10-Bù Đăng-Bình Phước"
                                           },
                                           {
                                               "TongCOD":  190.29,
                                               "CODTuNop":  79.61,
                                               "TyLeTuNop":  41.8,
                                               "TenBuuCuc":  "Bưu Cục DT 786 Khu Phố 4-Bến Cầu-Tây Ninh"
                                           },
                                           {
                                               "TongCOD":  151.62,
                                               "CODTuNop":  88.11,
                                               "TyLeTuNop":  58.1,
                                               "TenBuuCuc":  "Bưu Cục 76/540 KP8-Hố Nai-TP Biên Hòa-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  383.08,
                                               "CODTuNop":  168.36,
                                               "TyLeTuNop":  43.9,
                                               "TenBuuCuc":  "Bưu Cục KP2 Tân Hiệp-Định Quán Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  210.69,
                                               "CODTuNop":  62.65,
                                               "TyLeTuNop":  29.7,
                                               "TenBuuCuc":  "Bưu Cục 30 Nguyễn Trãi-Phường 3-Tây Ninh"
                                           },
                                           {
                                               "TongCOD":  281.81,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Ấp Việt Kiều QL1A-Xuân Lộc-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  124.6,
                                               "CODTuNop":  75.04,
                                               "TyLeTuNop":  60.2,
                                               "TenBuuCuc":  "Bưu Cục Đường Tỉnh 830-Xã Lương Bình-Long An"
                                           },
                                           {
                                               "TongCOD":  138.53,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 600 DT782 Bàu Đồn-Gò Dầu-Tây Ninh"
                                           },
                                           {
                                               "TongCOD":  271.24,
                                               "CODTuNop":  111.18,
                                               "TyLeTuNop":  41,
                                               "TenBuuCuc":  "Bưu Cục 893 Quốc Lộ 14-Đồng Xoài-Bình Phước"
                                           },
                                           {
                                               "TongCOD":  221.6,
                                               "CODTuNop":  221.6,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Tổ 2 Ấp 5 Xuân Tâm-Xuân Lộc-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  155.85,
                                               "CODTuNop":  23.24,
                                               "TyLeTuNop":  14.9,
                                               "TenBuuCuc":  "Bưu Cục An Lợi-Trảng Bàng-Tây Ninh"
                                           },
                                           {
                                               "TongCOD":  161.39,
                                               "CODTuNop":  26.48,
                                               "TyLeTuNop":  16.4,
                                               "TenBuuCuc":  "Bưu Cục Thống Nhất-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  123.92,
                                               "CODTuNop":  57.26,
                                               "TyLeTuNop":  46.2,
                                               "TenBuuCuc":  "Bưu cục An Lộc-Bình Long-Bình Phước"
                                           },
                                           {
                                               "TongCOD":  188.34,
                                               "CODTuNop":  56.67,
                                               "TyLeTuNop":  30.1,
                                               "TenBuuCuc":  "Bưu Cục Khu 6 ấp Bàu Cá-Trung Hòa-Trảng Bom-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  188.74,
                                               "CODTuNop":  105.07,
                                               "TyLeTuNop":  55.7,
                                               "TenBuuCuc":  "Bưu Cục 393 Tôn Đức Thắng-Hòa Thành-Tây Ninh"
                                           },
                                           {
                                               "TongCOD":  134.85,
                                               "CODTuNop":  68.62,
                                               "TyLeTuNop":  50.9,
                                               "TenBuuCuc":  "Bưu Cục Tân Đồng-Đồng Xoài-Bình Phước"
                                           },
                                           {
                                               "TongCOD":  153.88,
                                               "CODTuNop":  14.07,
                                               "TyLeTuNop":  9.1,
                                               "TenBuuCuc":  "Bưu Cục 105C Đồng Khởi-Biên Hòa-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  325.24,
                                               "CODTuNop":  157.54,
                                               "TyLeTuNop":  48.4,
                                               "TenBuuCuc":  "Bưu Cục Ấp Đông Kim-Xã Gia Kiệm-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  255.01,
                                               "CODTuNop":  64.79,
                                               "TyLeTuNop":  25.4,
                                               "TenBuuCuc":  "Bưu Cục Lộc Thái-Lộc Ninh-Bình Phước"
                                           },
                                           {
                                               "TongCOD":  250.92,
                                               "CODTuNop":  91.77,
                                               "TyLeTuNop":  36.6,
                                               "TenBuuCuc":  "Bưu Cục Ấp Rừng Sến-Xã Đức Lập-Long An"
                                           },
                                           {
                                               "TongCOD":  159.22,
                                               "CODTuNop":  51.13,
                                               "TyLeTuNop":  32.1,
                                               "TenBuuCuc":  "Bưu Cục Long Phước-Long Thành-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  70.43,
                                               "CODTuNop":  1.88,
                                               "TyLeTuNop":  2.7,
                                               "TenBuuCuc":  "Bưu Cục 69 Bàu Cạn-Xã Long Phước-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  137.24,
                                               "CODTuNop":  32.25,
                                               "TyLeTuNop":  23.5,
                                               "TenBuuCuc":  "Bưu Cục Phú Lập-Tân Phú-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  312.91,
                                               "CODTuNop":  75.01,
                                               "TyLeTuNop":  24,
                                               "TenBuuCuc":  "Bưu Cục 270 Phan Văn Mãng- Bến Lức-Long An"
                                           },
                                           {
                                               "TongCOD":  317.22,
                                               "CODTuNop":  146.16,
                                               "TyLeTuNop":  46.1,
                                               "TenBuuCuc":  "Bưu Cục Phạm Hùng - Phước Long - Bình Phước"
                                           },
                                           {
                                               "TongCOD":  214.01,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Bình Sơn-Long Thành-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  121.44,
                                               "CODTuNop":  41.56,
                                               "TyLeTuNop":  34.2,
                                               "TenBuuCuc":  "Bưu Cục 25 Khedol -Xã Suối Đá-Dương Minh Châu-Tây Ninh"
                                           },
                                           {
                                               "TongCOD":  190.43,
                                               "CODTuNop":  190.43,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 145/36/4 Điểu Xiển-Tân Biên-TP.Biên Hòa-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  172.37,
                                               "CODTuNop":  76.64,
                                               "TyLeTuNop":  44.5,
                                               "TenBuuCuc":  "Bưu Cục 1070 Quốc Lộ 51-TP. Biên Hòa-Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  91.07,
                                               "CODTuNop":  7.41,
                                               "TyLeTuNop":  8.1,
                                               "TenBuuCuc":  "Bưu Cục Ấp Tân Hòa-Tân Biên-Tây Ninh"
                                           }
                                       ]
                    },
                    {
                        "TyLeTuNop":  50,
                        "TenVung":  "DSH",
                        "TongCOD":  4675.86,
                        "Trend":  -3,
                        "CODTuNop":  2338.28,
                        "BuuCucList":  [
                                           {
                                               "TongCOD":  185.42,
                                               "CODTuNop":  127.51,
                                               "TyLeTuNop":  68.8,
                                               "TenBuuCuc":  "Bưu Cục 176 Nguyễn Đức Cảnh-Hoàng Diệu-Thái Bình"
                                           },
                                           {
                                               "TongCOD":  97.22,
                                               "CODTuNop":  62.88,
                                               "TyLeTuNop":  64.7,
                                               "TenBuuCuc":  "Bưu Cục Thôn Cao Xá-Trần Cao-Phù Cừ-Hưng Yên"
                                           },
                                           {
                                               "TongCOD":  484.79,
                                               "CODTuNop":  189.84,
                                               "TyLeTuNop":  39.2,
                                               "TenBuuCuc":  "Bưu Cục Xã Tân Tiến-Văn Giang-Hưng Yên"
                                           },
                                           {
                                               "TongCOD":  149.83,
                                               "CODTuNop":  78.09,
                                               "TyLeTuNop":  52.1,
                                               "TenBuuCuc":  "Bưu Cục Thị Trấn Ân Thi-Ân Thi-Hưng Yên"
                                           },
                                           {
                                               "TongCOD":  123.49,
                                               "CODTuNop":  64.52,
                                               "TyLeTuNop":  52.2,
                                               "TenBuuCuc":  "Bưu Cục Thôn Đại Từ-Văn Lâm-Hưng Yên"
                                           },
                                           {
                                               "TongCOD":  155.23,
                                               "CODTuNop":  14.88,
                                               "TyLeTuNop":  9.6,
                                               "TenBuuCuc":  "Bưu Cục 90 Nguyễn Lương Bằng-Kim Động-Hưng Yên"
                                           },
                                           {
                                               "TongCOD":  419.04,
                                               "CODTuNop":  135.31,
                                               "TyLeTuNop":  32.3,
                                               "TenBuuCuc":  "Bưu Cục 11 Chu Mạnh Trinh-TP.Hưng Yên"
                                           },
                                           {
                                               "TongCOD":  219.82,
                                               "CODTuNop":  162.6,
                                               "TyLeTuNop":  74,
                                               "TenBuuCuc":  "Bưu Cục Xã Thống Nhất-Hưng Hà-Thái Bình"
                                           },
                                           {
                                               "TongCOD":  161.75,
                                               "CODTuNop":  128.91,
                                               "TyLeTuNop":  79.7,
                                               "TenBuuCuc":  "Bưu Cục Thôn Lực Điền-Yên Mỹ-Hưng Yên"
                                           },
                                           {
                                               "TongCOD":  181.72,
                                               "CODTuNop":  40.5,
                                               "TyLeTuNop":  22.3,
                                               "TenBuuCuc":  "Bưu Cục Thị Trấn Như Quỳnh-Văn Lâm-Hưng Yên"
                                           },
                                           {
                                               "TongCOD":  164.3,
                                               "CODTuNop":  100.14,
                                               "TyLeTuNop":  61,
                                               "TenBuuCuc":  "Bưu Cục Đường Doãn Khuê-Phường Vũ Phúc-Thái Bình"
                                           },
                                           {
                                               "TongCOD":  236.61,
                                               "CODTuNop":  97.06,
                                               "TyLeTuNop":  41,
                                               "TenBuuCuc":  "Bưu Cục Đông Hưng-Thái Bình"
                                           },
                                           {
                                               "TongCOD":  133.86,
                                               "CODTuNop":  27.48,
                                               "TyLeTuNop":  20.5,
                                               "TenBuuCuc":  "Bưu Cục Quỳnh Hội-Quỳnh Phụ-Thái Bình"
                                           },
                                           {
                                               "TongCOD":  265.95,
                                               "CODTuNop":  167.11,
                                               "TyLeTuNop":  62.8,
                                               "TenBuuCuc":  "Bưu Cục Quốc Lộ 39B-Kiến Xương- Thái Bình"
                                           },
                                           {
                                               "TongCOD":  303.53,
                                               "CODTuNop":  168.7,
                                               "TyLeTuNop":  55.6,
                                               "TenBuuCuc":  "Bưu Cục Cầu Khé DT377-Xã Khoái Châu-Hưng Yên"
                                           },
                                           {
                                               "TongCOD":  285.45,
                                               "CODTuNop":  164.2,
                                               "TyLeTuNop":  57.5,
                                               "TenBuuCuc":  "Bưu Cục Nhân Hòa-Mỹ Hào-Hưng Yên"
                                           },
                                           {
                                               "TongCOD":  244.35,
                                               "CODTuNop":  165.56,
                                               "TyLeTuNop":  67.8,
                                               "TenBuuCuc":  "Bưu Cục DT 376-Ngọc Long-Yên Mỹ-Hưng Yên"
                                           },
                                           {
                                               "TongCOD":  301.77,
                                               "CODTuNop":  143.69,
                                               "TyLeTuNop":  47.6,
                                               "TenBuuCuc":  "Bưu Cục Thôn Đông-Tiền Hải-Thái Bình"
                                           },
                                           {
                                               "TongCOD":  258.66,
                                               "CODTuNop":  127.97,
                                               "TyLeTuNop":  49.5,
                                               "TenBuuCuc":  "Bưu Cục Minh Khai-Vũ Thư-Thái Bình"
                                           },
                                           {
                                               "TongCOD":  94.04,
                                               "CODTuNop":  20.7,
                                               "TyLeTuNop":  22,
                                               "TenBuuCuc":  "Bưu Cục Thôn Vũ Hạ-Xã Phụ Dực-Thái Bình"
                                           },
                                           {
                                               "TongCOD":  209.02,
                                               "CODTuNop":  150.63,
                                               "TyLeTuNop":  72.1,
                                               "TenBuuCuc":  "Bưu Cục Xóm 2 Thụy Sơn-Thái Thụy-Thái Bình"
                                           }
                                       ]
                    },
                    {
                        "TyLeTuNop":  78.3,
                        "TenVung":  "GXT",
                        "TongCOD":  20671.2,
                        "Trend":  -3.1,
                        "CODTuNop":  16184.24,
                        "BuuCucList":  [
                                           {
                                               "TongCOD":  256.98,
                                               "CODTuNop":  256.98,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Đông Hưng - Thái Bình"
                                           },
                                           {
                                               "TongCOD":  205,
                                               "CODTuNop":  205,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Rạch Giá - Kiên Giang"
                                           },
                                           {
                                               "TongCOD":  201.76,
                                               "CODTuNop":  201.76,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Pleiku - Gia Lai"
                                           },
                                           {
                                               "TongCOD":  125.13,
                                               "CODTuNop":  125.13,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Đồng Hới - Quảng Bình"
                                           },
                                           {
                                               "TongCOD":  82.83,
                                               "CODTuNop":  82.83,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - La Gi - Bình Thuận"
                                           },
                                           {
                                               "TongCOD":  471.48,
                                               "CODTuNop":  374.49,
                                               "TyLeTuNop":  79.4,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Liên Chiểu - Đà Nẵng"
                                           },
                                           {
                                               "TongCOD":  326.78,
                                               "CODTuNop":  251.34,
                                               "TyLeTuNop":  76.9,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Trảng Bom - Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  432.06,
                                               "CODTuNop":  304.71,
                                               "TyLeTuNop":  70.5,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Biên Hoà - Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  332.87,
                                               "CODTuNop":  266.11,
                                               "TyLeTuNop":  79.9,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Châu Đốc - An Giang"
                                           },
                                           {
                                               "TongCOD":  67.54,
                                               "CODTuNop":  53.1,
                                               "TyLeTuNop":  78.6,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Cam Ranh - Khánh Hoà"
                                           },
                                           {
                                               "TongCOD":  172.05,
                                               "CODTuNop":  172.05,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Việt Yên - Bắc Giang"
                                           },
                                           {
                                               "TongCOD":  318.06,
                                               "CODTuNop":  318.06,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Thuận An - Bình Dương"
                                           },
                                           {
                                               "TongCOD":  137.65,
                                               "CODTuNop":  78.17,
                                               "TyLeTuNop":  56.8,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Nho Quan - Ninh Bình"
                                           },
                                           {
                                               "TongCOD":  90.78,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Bình Long - Bình Phước"
                                           },
                                           {
                                               "TongCOD":  259.19,
                                               "CODTuNop":  259.19,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Dĩ An - Bình Dương"
                                           },
                                           {
                                               "TongCOD":  150.01,
                                               "CODTuNop":  150.01,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Hội An - Quảng Nam"
                                           },
                                           {
                                               "TongCOD":  125.29,
                                               "CODTuNop":  125.29,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Đông Anh - Hà Nội"
                                           },
                                           {
                                               "TongCOD":  43.52,
                                               "CODTuNop":  43.52,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Lạng Sơn - Lạng Sơn"
                                           },
                                           {
                                               "TongCOD":  301.62,
                                               "CODTuNop":  301.62,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TX Cai Lậy - Tiền Giang"
                                           },
                                           {
                                               "TongCOD":  39.31,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Kho giao hàng Ahamove Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  801.22,
                                               "CODTuNop":  781.98,
                                               "TyLeTuNop":  97.6,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Tân Tạo - HCM"
                                           },
                                           {
                                               "TongCOD":  170.9,
                                               "CODTuNop":  170.9,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Cà Mau - Cà Mau"
                                           },
                                           {
                                               "TongCOD":  118.14,
                                               "CODTuNop":  74.71,
                                               "TyLeTuNop":  63.2,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Hoài Nhơn - Bình Định"
                                           },
                                           {
                                               "TongCOD":  516.23,
                                               "CODTuNop":  516.23,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Thủ Đức - HCM"
                                           },
                                           {
                                               "TongCOD":  387.19,
                                               "CODTuNop":  308.61,
                                               "TyLeTuNop":  79.7,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Hoà Xuân - Đà Nẵng"
                                           },
                                           {
                                               "TongCOD":  227.54,
                                               "CODTuNop":  227.54,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Bình Xuyên - Vĩnh Phúc"
                                           },
                                           {
                                               "TongCOD":  291.59,
                                               "CODTuNop":  188.01,
                                               "TyLeTuNop":  64.5,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Bà Rịa - BRVT"
                                           },
                                           {
                                               "TongCOD":  278.23,
                                               "CODTuNop":  278.23,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Mỹ Lộc - Nam Định"
                                           },
                                           {
                                               "TongCOD":  293.41,
                                               "CODTuNop":  169.33,
                                               "TyLeTuNop":  57.7,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Tây Ninh - Tây Ninh"
                                           },
                                           {
                                               "TongCOD":  315.55,
                                               "CODTuNop":  315.55,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Thái Nguyên - Thái Nguyên"
                                           },
                                           {
                                               "TongCOD":  229.6,
                                               "CODTuNop":  212.36,
                                               "TyLeTuNop":  92.5,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Bắc Từ Liêm - Hà Nội"
                                           },
                                           {
                                               "TongCOD":  46.95,
                                               "CODTuNop":  46.95,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Hà Tiên - Kiên Giang"
                                           },
                                           {
                                               "TongCOD":  128.85,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Kho giao hàng Ahamove HCM"
                                           },
                                           {
                                               "TongCOD":  160.2,
                                               "CODTuNop":  98.32,
                                               "TyLeTuNop":  61.4,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Phan Thiết - Bình Thuận"
                                           },
                                           {
                                               "TongCOD":  185.29,
                                               "CODTuNop":  185.29,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Đức Hòa - Long An"
                                           },
                                           {
                                               "TongCOD":  272.07,
                                               "CODTuNop":  272.07,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Tiên Du - Bắc Ninh"
                                           },
                                           {
                                               "TongCOD":  145.02,
                                               "CODTuNop":  145.02,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Tân An - Long An"
                                           },
                                           {
                                               "TongCOD":  175.93,
                                               "CODTuNop":  131.43,
                                               "TyLeTuNop":  74.7,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Tân Bình - Hải Dương"
                                           },
                                           {
                                               "TongCOD":  55.83,
                                               "CODTuNop":  55.83,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Gò Công - Tiền Giang"
                                           },
                                           {
                                               "TongCOD":  86.24,
                                               "CODTuNop":  86.24,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Tuyên Quang - Tuyên Quang"
                                           },
                                           {
                                               "TongCOD":  58.94,
                                               "CODTuNop":  58.94,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Sơn La - Sơn La"
                                           },
                                           {
                                               "TongCOD":  203.73,
                                               "CODTuNop":  203.7,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Hương Thủy - Huế"
                                           },
                                           {
                                               "TongCOD":  247.25,
                                               "CODTuNop":  247.25,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Đông Thọ - Thanh Hoá"
                                           },
                                           {
                                               "TongCOD":  284.58,
                                               "CODTuNop":  191.78,
                                               "TyLeTuNop":  67.4,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Vĩnh Long - Vĩnh Long"
                                           },
                                           {
                                               "TongCOD":  530.2,
                                               "CODTuNop":  530.2,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Tân Bình - HCM"
                                           },
                                           {
                                               "TongCOD":  82.85,
                                               "CODTuNop":  82.85,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Thanh Trì - Hà Nội"
                                           },
                                           {
                                               "TongCOD":  120.21,
                                               "CODTuNop":  120.21,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Lào Cai - Lào Cai"
                                           },
                                           {
                                               "TongCOD":  75.62,
                                               "CODTuNop":  64.21,
                                               "TyLeTuNop":  84.9,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Tam Nông - Đồng Tháp"
                                           },
                                           {
                                               "TongCOD":  124.85,
                                               "CODTuNop":  124.85,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Ân Thi - Hưng Yên"
                                           },
                                           {
                                               "TongCOD":  165.26,
                                               "CODTuNop":  165.26,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Mỹ Tho - Tiền Giang"
                                           },
                                           {
                                               "TongCOD":  229.88,
                                               "CODTuNop":  229.88,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Bến Tre - Bến Tre"
                                           },
                                           {
                                               "TongCOD":  597.85,
                                               "CODTuNop":  396.32,
                                               "TyLeTuNop":  66.3,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Nha Trang - Khánh Hoà"
                                           },
                                           {
                                               "TongCOD":  123.56,
                                               "CODTuNop":  123.56,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Tuy Hoà - Phú Yên"
                                           },
                                           {
                                               "TongCOD":  222.27,
                                               "CODTuNop":  197.74,
                                               "TyLeTuNop":  89,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Thanh Oai - Hà Nội"
                                           },
                                           {
                                               "TongCOD":  147.14,
                                               "CODTuNop":  147.14,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Trà Vinh - Trà Vinh"
                                           },
                                           {
                                               "TongCOD":  312.35,
                                               "CODTuNop":  312.35,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Long Xuyên - An Giang"
                                           },
                                           {
                                               "TongCOD":  68.69,
                                               "CODTuNop":  68.69,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Yên Bái - Yên Bái"
                                           },
                                           {
                                               "TongCOD":  148.17,
                                               "CODTuNop":  115.7,
                                               "TyLeTuNop":  78.1,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Đồng Xoài - Bình Phước"
                                           },
                                           {
                                               "TongCOD":  129.7,
                                               "CODTuNop":  129.7,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Tam Kỳ - Quảng Nam"
                                           },
                                           {
                                               "TongCOD":  81.17,
                                               "CODTuNop":  81.17,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Đông Hà - Quảng Trị"
                                           },
                                           {
                                               "TongCOD":  243.57,
                                               "CODTuNop":  243.57,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Đà Lạt - Lâm Đồng"
                                           },
                                           {
                                               "TongCOD":  335.69,
                                               "CODTuNop":  335.69,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Nhà Bè - HCM"
                                           },
                                           {
                                               "TongCOD":  107.6,
                                               "CODTuNop":  25.43,
                                               "TyLeTuNop":  23.6,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Bạc Liêu - Bạc Liêu"
                                           },
                                           {
                                               "TongCOD":  129.05,
                                               "CODTuNop":  129.05,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Phan Rang - Ninh Thuận"
                                           },
                                           {
                                               "TongCOD":  385.16,
                                               "CODTuNop":  16.16,
                                               "TyLeTuNop":  4.2,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Buôn Ma Thuột - Đắk Lắk"
                                           },
                                           {
                                               "TongCOD":  369.8,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Bến Cát - Bình Dương"
                                           },
                                           {
                                               "TongCOD":  101.71,
                                               "CODTuNop":  67.24,
                                               "TyLeTuNop":  66.1,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Quảng Ngãi - Quảng Ngãi"
                                           },
                                           {
                                               "TongCOD":  221.82,
                                               "CODTuNop":  60.6,
                                               "TyLeTuNop":  27.3,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Quy Nhơn - Bình Định"
                                           },
                                           {
                                               "TongCOD":  297.47,
                                               "CODTuNop":  297.47,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Cao Lãnh - Đồng Tháp"
                                           },
                                           {
                                               "TongCOD":  112.31,
                                               "CODTuNop":  112.31,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Vị Thanh - Hậu Giang"
                                           },
                                           {
                                               "TongCOD":  313.2,
                                               "CODTuNop":  173.06,
                                               "TyLeTuNop":  55.3,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Thủ Dầu Một - Bình Dương"
                                           },
                                           {
                                               "TongCOD":  134.99,
                                               "CODTuNop":  134.99,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Bình Thủy - Cần Thơ"
                                           },
                                           {
                                               "TongCOD":  132.69,
                                               "CODTuNop":  132.69,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Việt Trì - Phú Thọ"
                                           },
                                           {
                                               "TongCOD":  277.55,
                                               "CODTuNop":  277.55,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Long Biên - Hà Nội"
                                           },
                                           {
                                               "TongCOD":  326.15,
                                               "CODTuNop":  326.15,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Cái Răng - Cần Thơ"
                                           },
                                           {
                                               "TongCOD":  356.57,
                                               "CODTuNop":  356.57,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Hạ Long - Quảng Ninh"
                                           },
                                           {
                                               "TongCOD":  265.16,
                                               "CODTuNop":  136.77,
                                               "TyLeTuNop":  51.6,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Vũng Tàu - BRVT"
                                           },
                                           {
                                               "TongCOD":  190.45,
                                               "CODTuNop":  190.45,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Hồ Chí Minh"
                                           },
                                           {
                                               "TongCOD":  281.68,
                                               "CODTuNop":  42.89,
                                               "TyLeTuNop":  15.2,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Vinh - Nghệ An"
                                           },
                                           {
                                               "TongCOD":  177.1,
                                               "CODTuNop":  177.1,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Hoài Đức - Hà Nội"
                                           },
                                           {
                                               "TongCOD":  206.53,
                                               "CODTuNop":  39.15,
                                               "TyLeTuNop":  19,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Bảo Lộc - Lâm Đồng"
                                           },
                                           {
                                               "TongCOD":  59.64,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Gia Nghĩa - Đắk Nông"
                                           },
                                           {
                                               "TongCOD":  605.52,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Kho giao hàng Ahamove Hà Nội"
                                           },
                                           {
                                               "TongCOD":  148.69,
                                               "CODTuNop":  148.69,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Đức Trọng - Lâm Đồng"
                                           },
                                           {
                                               "TongCOD":  152.43,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Đức Long - Hà Nội"
                                           },
                                           {
                                               "TongCOD":  86.71,
                                               "CODTuNop":  70.19,
                                               "TyLeTuNop":  81,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Thạch Linh - Hà Tĩnh"
                                           },
                                           {
                                               "TongCOD":  248.81,
                                               "CODTuNop":  248.81,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Long Thành - Đồng Nai"
                                           },
                                           {
                                               "TongCOD":  737.35,
                                               "CODTuNop":  622.85,
                                               "TyLeTuNop":  84.5,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - An Dương - Hải Phòng"
                                           },
                                           {
                                               "TongCOD":  242.95,
                                               "CODTuNop":  98.11,
                                               "TyLeTuNop":  40.4,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Sóc Trăng - Sóc Trăng"
                                           },
                                           {
                                               "TongCOD":  73.81,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TX Ngã Năm - Sóc Trăng"
                                           },
                                           {
                                               "TongCOD":  57.12,
                                               "CODTuNop":  57.12,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Thắng Lợi - Kon Tum"
                                           },
                                           {
                                               "TongCOD":  27.29,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - TP Hòa Bình - Hòa Bình"
                                           },
                                           {
                                               "TongCOD":  110.16,
                                               "CODTuNop":  63.8,
                                               "TyLeTuNop":  57.9,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Tuy Phong - Bình Thuận"
                                           },
                                           {
                                               "TongCOD":  95.27,
                                               "CODTuNop":  95.27,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Ô Môn - Cần Thơ"
                                           },
                                           {
                                               "TongCOD":  79.06,
                                               "CODTuNop":  79.06,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Kho Giao Hàng Nặng - Buôn Hồ - Đắk Lắk"
                                           }
                                       ]
                    },
                    {
                        "TyLeTuNop":  80,
                        "TenVung":  "HCM",
                        "TongCOD":  28155.47,
                        "Trend":  -0.1,
                        "CODTuNop":  22517.96,
                        "BuuCucList":  [
                                           {
                                               "TongCOD":  80.85,
                                               "CODTuNop":  80.85,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa Kho 02 367 Lương Định Của-Q.2-HCM"
                                           },
                                           {
                                               "TongCOD":  158.27,
                                               "CODTuNop":  158.27,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục KP6 Vĩnh Tân-Tân Uyên-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  193.58,
                                               "CODTuNop":  184.34,
                                               "TyLeTuNop":  95.2,
                                               "TenBuuCuc":  "Bưu Cục 256/59 Đường Thạnh Xuân 25-Quận 12-HCM"
                                           },
                                           {
                                               "TongCOD":  261.15,
                                               "CODTuNop":  160,
                                               "TyLeTuNop":  61.3,
                                               "TenBuuCuc":  "Bưu Cục Vĩnh Lộc A-Bình Chánh-HCM"
                                           },
                                           {
                                               "TongCOD":  191.7,
                                               "CODTuNop":  32.54,
                                               "TyLeTuNop":  17,
                                               "TenBuuCuc":  "Bưu Cục 39 Tôn Đức Thắng-TP.Bà Rịa"
                                           },
                                           {
                                               "TongCOD":  266.04,
                                               "CODTuNop":  185.46,
                                               "TyLeTuNop":  69.7,
                                               "TenBuuCuc":  "Bưu Cục 309 Nguyễn Thị Ngâu-Hóc Môn-HCM"
                                           },
                                           {
                                               "TongCOD":  172.43,
                                               "CODTuNop":  35.37,
                                               "TyLeTuNop":  20.5,
                                               "TenBuuCuc":  "Bưu Cục 63A Trần Đình Xu-Q.1-HCM"
                                           },
                                           {
                                               "TongCOD":  134.58,
                                               "CODTuNop":  134.58,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 178 Đường DT741-Bến Cát-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  215.03,
                                               "CODTuNop":  92.22,
                                               "TyLeTuNop":  42.9,
                                               "TenBuuCuc":  "Bưu Cục Khu Phố Vạn Hạnh-Phú Mỹ-Bà Rịa Vũng Tàu"
                                           },
                                           {
                                               "TongCOD":  87.17,
                                               "CODTuNop":  86.86,
                                               "TyLeTuNop":  99.6,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa A5 KDC Tân Thuận Nam-Phú Thuận-Quận 7-HCM"
                                           },
                                           {
                                               "TongCOD":  76.99,
                                               "CODTuNop":  76.43,
                                               "TyLeTuNop":  99.3,
                                               "TenBuuCuc":  "Bưu Cục 295F Lý Thường Kiệt-Quận 11-HCM"
                                           },
                                           {
                                               "TongCOD":  0.05,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Chuyên Lấy 1162/6A Tỉnh Lộ 10-Q.Bình Tân-HCM"
                                           },
                                           {
                                               "TongCOD":  7.45,
                                               "CODTuNop":  7.33,
                                               "TyLeTuNop":  98.3,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa 198C Thạnh Lộc-Q.12-HCM"
                                           },
                                           {
                                               "TongCOD":  145.56,
                                               "CODTuNop":  145.56,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 13 Đường Số 8-Hiệp Bình Phước-Q.Thủ Đức-HCM"
                                           },
                                           {
                                               "TongCOD":  178.69,
                                               "CODTuNop":  117.85,
                                               "TyLeTuNop":  66,
                                               "TenBuuCuc":  "Bưu Cục 155 Quốc lộ 1A-Bình Chiểu-Thủ Đức-HCM"
                                           },
                                           {
                                               "TongCOD":  174.95,
                                               "CODTuNop":  174.67,
                                               "TyLeTuNop":  99.8,
                                               "TenBuuCuc":  "Bưu Cục 195B Lý Chính Thắng-Q.3-HCM"
                                           },
                                           {
                                               "TongCOD":  144.2,
                                               "CODTuNop":  144.2,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 04 Đường 13 tháng 3-Dầu Tiếng-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  260.62,
                                               "CODTuNop":  260.62,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 390 Nguyễn Kim Cương-Tân Thạnh Đông-Củ Chi"
                                           },
                                           {
                                               "TongCOD":  257.02,
                                               "CODTuNop":  160.7,
                                               "TyLeTuNop":  62.5,
                                               "TenBuuCuc":  "Bưu Cục 02 Đường Số 28 - Củ Chi - HCM"
                                           },
                                           {
                                               "TongCOD":  214.21,
                                               "CODTuNop":  195.21,
                                               "TyLeTuNop":  91.1,
                                               "TenBuuCuc":  "Bưu Cục Đường 268 Đỗ Xuân Hợp-Phường Phước Long-HCM"
                                           },
                                           {
                                               "TongCOD":  114.65,
                                               "CODTuNop":  89.55,
                                               "TyLeTuNop":  78.1,
                                               "TenBuuCuc":  "Bưu Cục F5/9A Trần Văn Giàu-Bình Chánh-HCM"
                                           },
                                           {
                                               "TongCOD":  32.04,
                                               "CODTuNop":  32.04,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa Đường 268 Đỗ Xuân Hợp-Q.9-HCM"
                                           },
                                           {
                                               "TongCOD":  104.12,
                                               "CODTuNop":  81.43,
                                               "TyLeTuNop":  78.2,
                                               "TenBuuCuc":  "Bưu Cục 136 Đường Huyện 704-Minh Tân-Dầu Tiếng-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  126.56,
                                               "CODTuNop":  126.56,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 01 Đường Số 6-Long Bình-Q.9-HCM"
                                           },
                                           {
                                               "TongCOD":  275.46,
                                               "CODTuNop":  275.34,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Phú Lợi-TP.Thủ Dầu Một-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  199.03,
                                               "CODTuNop":  198.5,
                                               "TyLeTuNop":  99.7,
                                               "TenBuuCuc":  "Bưu Cục 53 Đường 22-Linh Đông-Thủ Đức-HCM"
                                           },
                                           {
                                               "TongCOD":  216.42,
                                               "CODTuNop":  100.8,
                                               "TyLeTuNop":  46.6,
                                               "TenBuuCuc":  "Bưu cục Xuân Thới Thượng-Hóc Môn-HCM"
                                           },
                                           {
                                               "TongCOD":  170.81,
                                               "CODTuNop":  152.75,
                                               "TyLeTuNop":  89.4,
                                               "TenBuuCuc":  "Bưu Cục 1135 Phan Văn Trị-Q.Gò Vấp-HCM"
                                           },
                                           {
                                               "TongCOD":  102.05,
                                               "CODTuNop":  89.19,
                                               "TyLeTuNop":  87.4,
                                               "TenBuuCuc":  "Bưu Cục Phường Tây Thạnh-Tân Phú-HCM"
                                           },
                                           {
                                               "TongCOD":  227.51,
                                               "CODTuNop":  227.37,
                                               "TyLeTuNop":  99.9,
                                               "TenBuuCuc":  "Bưu Cục 480F Cao Thắng - Trường Sa-Q.3-HCM"
                                           },
                                           {
                                               "TongCOD":  159.56,
                                               "CODTuNop":  159.45,
                                               "TyLeTuNop":  99.9,
                                               "TenBuuCuc":  "Bưu Cục DA 1-1 Mỹ Phước-Bến Cát-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  174.6,
                                               "CODTuNop":  174.42,
                                               "TyLeTuNop":  99.9,
                                               "TenBuuCuc":  "Bưu Cục Tăng Nhơn Phú B-Q.9-HCM"
                                           },
                                           {
                                               "TongCOD":  83.46,
                                               "CODTuNop":  66.09,
                                               "TyLeTuNop":  79.2,
                                               "TenBuuCuc":  "Bưu Cục 374 Xô Viết Nghệ Tĩnh-Q.Bình Thạnh-HCM"
                                           },
                                           {
                                               "TongCOD":  114.87,
                                               "CODTuNop":  114.73,
                                               "TyLeTuNop":  99.9,
                                               "TenBuuCuc":  "Bưu Cục 27 Nguyễn Xí-Q.Bình Thạnh-HCM"
                                           },
                                           {
                                               "TongCOD":  225.68,
                                               "CODTuNop":  72.71,
                                               "TyLeTuNop":  32.2,
                                               "TenBuuCuc":  "Bưu Cục 265 Lào Cai-Chánh Nghĩa-Thủ Dầu Một-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  70.49,
                                               "CODTuNop":  70.49,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa 1130 Đông Hưng Thuận 05-Q.12-HCM"
                                           },
                                           {
                                               "TongCOD":  134.29,
                                               "CODTuNop":  77.37,
                                               "TyLeTuNop":  57.6,
                                               "TenBuuCuc":  "Bưu Cục 182 Xuyên Á-TP.Dĩ An-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  113.68,
                                               "CODTuNop":  113.52,
                                               "TyLeTuNop":  99.9,
                                               "TenBuuCuc":  "Bưu Cục 185 Tân Thới Nhất 8-Phường Đông Hưng Thuận-HCM"
                                           },
                                           {
                                               "TongCOD":  146.24,
                                               "CODTuNop":  118.04,
                                               "TyLeTuNop":  80.7,
                                               "TenBuuCuc":  "Bưu cục Long Thạnh Mỹ-Q9-HCM"
                                           },
                                           {
                                               "TongCOD":  169.77,
                                               "CODTuNop":  143.7,
                                               "TyLeTuNop":  84.6,
                                               "TenBuuCuc":  "Bưu Cục Số 17 Đường Hoàng Văn Hợp-An Lạc A-Bình Tân-HCM"
                                           },
                                           {
                                               "TongCOD":  222.5,
                                               "CODTuNop":  188.02,
                                               "TyLeTuNop":  84.5,
                                               "TenBuuCuc":  "Bưu Cục Số 8 Lê Chí Dân-Thủ Dầu Một-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  197.69,
                                               "CODTuNop":  75.41,
                                               "TyLeTuNop":  38.1,
                                               "TenBuuCuc":  "Bưu Cục KP Thanh Long -Đất Đỏ-Bà Rịa Vũng Tàu"
                                           },
                                           {
                                               "TongCOD":  155.58,
                                               "CODTuNop":  111.77,
                                               "TyLeTuNop":  71.8,
                                               "TenBuuCuc":  "Bưu Cục Phan Huy Ích-Gò Vấp-HCM"
                                           },
                                           {
                                               "TongCOD":  76.39,
                                               "CODTuNop":  76.39,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa 320 Nguyễn Thị Kiểu-Hiệp Thành-Quận 12-HCM"
                                           },
                                           {
                                               "TongCOD":  262.92,
                                               "CODTuNop":  262.92,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Số 25 Đường 12-An Phú-Q.2-HCM"
                                           },
                                           {
                                               "TongCOD":  149.22,
                                               "CODTuNop":  101.51,
                                               "TyLeTuNop":  68,
                                               "TenBuuCuc":  "Bưu Cục 04 Nhất Chi Mai-Q.Tân Bình-HCM"
                                           },
                                           {
                                               "TongCOD":  117.26,
                                               "CODTuNop":  27.65,
                                               "TyLeTuNop":  23.6,
                                               "TenBuuCuc":  "Bưu Cục 274 Cô Bắc-Q.1-HCM"
                                           },
                                           {
                                               "TongCOD":  99.26,
                                               "CODTuNop":  98.58,
                                               "TyLeTuNop":  99.3,
                                               "TenBuuCuc":  "Bưu Cục 608 Phan Văn Trị-Gò Vấp-HCM"
                                           },
                                           {
                                               "TongCOD":  129.92,
                                               "CODTuNop":  128.88,
                                               "TyLeTuNop":  99.2,
                                               "TenBuuCuc":  "Bưu cục 215 Bình Thới-Q.11-HCM"
                                           },
                                           {
                                               "TongCOD":  137.16,
                                               "CODTuNop":  43.38,
                                               "TyLeTuNop":  31.6,
                                               "TenBuuCuc":  "Bưu Cục Ấp Phước Sơn-Phú Mỹ-Bà Rịa Vũng Tàu"
                                           },
                                           {
                                               "TongCOD":  133.2,
                                               "CODTuNop":  129.71,
                                               "TyLeTuNop":  97.4,
                                               "TenBuuCuc":  "Bưu Cục Phường Long Trường-Q9-HCM"
                                           },
                                           {
                                               "TongCOD":  152.65,
                                               "CODTuNop":  147.55,
                                               "TyLeTuNop":  96.7,
                                               "TenBuuCuc":  "Bưu Cục 21 Phú Mỹ-Q.Bình Thạnh-HCM"
                                           },
                                           {
                                               "TongCOD":  159.28,
                                               "CODTuNop":  159.28,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 503/16 Lê Thị Trung-Bình Chuẩn-TP.Thuận An-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  199.95,
                                               "CODTuNop":  199.95,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Đường NH17-Hòa Phú-Thủ Dầu Một-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  227.76,
                                               "CODTuNop":  227.45,
                                               "TyLeTuNop":  99.9,
                                               "TenBuuCuc":  "Bưu Cục 88/21 Phan Sào Nam-Q.Tân Bình-HCM"
                                           },
                                           {
                                               "TongCOD":  138.91,
                                               "CODTuNop":  82.57,
                                               "TyLeTuNop":  59.4,
                                               "TenBuuCuc":  "Bưu Cục 117 Hùng Vương-Q.5-HCM"
                                           },
                                           {
                                               "TongCOD":  188.64,
                                               "CODTuNop":  165.6,
                                               "TyLeTuNop":  87.8,
                                               "TenBuuCuc":  "Bưu Cục 38 Đường Số 3-Q.Bình Tân-HCM"
                                           },
                                           {
                                               "TongCOD":  81.79,
                                               "CODTuNop":  81.49,
                                               "TyLeTuNop":  99.6,
                                               "TenBuuCuc":  "Bưu Cục 147E Đội Cung-Q.11-HCM"
                                           },
                                           {
                                               "TongCOD":  138.89,
                                               "CODTuNop":  96.9,
                                               "TyLeTuNop":  69.8,
                                               "TenBuuCuc":  "Bưu Cục 9 Hồng Bàng-Q.Bình Thạnh-HCM"
                                           },
                                           {
                                               "TongCOD":  151.78,
                                               "CODTuNop":  108.72,
                                               "TyLeTuNop":  71.6,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa 971 Nguyễn Kiệm-Q.Gò Vấp-HCM"
                                           },
                                           {
                                               "TongCOD":  202.29,
                                               "CODTuNop":  199.88,
                                               "TyLeTuNop":  98.8,
                                               "TenBuuCuc":  "Bưu Cục 248A Nơ Trang Long-Q.Bình Thạnh-HCM"
                                           },
                                           {
                                               "TongCOD":  0.05,
                                               "CODTuNop":  0.05,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa Q.7 Kho 01 Số 860/36 Huỳnh Tấn Phát-Q.7-HCM"
                                           },
                                           {
                                               "TongCOD":  118.98,
                                               "CODTuNop":  64.26,
                                               "TyLeTuNop":  54,
                                               "TenBuuCuc":  "Bưu Cục 95 Nguyễn Thị Tươi-Tân Bình-Dĩ An-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  140.85,
                                               "CODTuNop":  40.62,
                                               "TyLeTuNop":  28.8,
                                               "TenBuuCuc":  "Bưu Cục 2227 Văn Tiến Dũng-Bình Chánh-HCM"
                                           },
                                           {
                                               "TongCOD":  228.38,
                                               "CODTuNop":  228.07,
                                               "TyLeTuNop":  99.9,
                                               "TenBuuCuc":  "Bưu Cục 2999 Phạm Thế Hiển-Q.8-HCM"
                                           },
                                           {
                                               "TongCOD":  97.26,
                                               "CODTuNop":  97.26,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Cần Giờ-HCM"
                                           },
                                           {
                                               "TongCOD":  44.94,
                                               "CODTuNop":  44.94,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa 16A Lê Hồng Phong-Q.10-HCM"
                                           },
                                           {
                                               "TongCOD":  141.8,
                                               "CODTuNop":  138.06,
                                               "TyLeTuNop":  97.4,
                                               "TenBuuCuc":  "Bưu Cục C1/7 Lại Hùng Cường-Vĩnh Lộc B-Bình Chánh-HCM"
                                           },
                                           {
                                               "TongCOD":  74.82,
                                               "CODTuNop":  41.43,
                                               "TyLeTuNop":  55.4,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa 6 Phạm Văn Hai-Q.Tân Bình-HCM"
                                           },
                                           {
                                               "TongCOD":  155.91,
                                               "CODTuNop":  155.79,
                                               "TyLeTuNop":  99.9,
                                               "TenBuuCuc":  "Bưu Cục 53A2 Đường 5-Linh Tây-Q.Thủ Đức-HCM"
                                           },
                                           {
                                               "TongCOD":  39.78,
                                               "CODTuNop":  39.78,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa 215 Bình Thới-Q.11-HCM"
                                           },
                                           {
                                               "TongCOD":  209.27,
                                               "CODTuNop":  139.86,
                                               "TyLeTuNop":  66.8,
                                               "TenBuuCuc":  "Bưu Cục 602/26 Điện Biên Phủ-Phường 12-Bình Thạnh-HCM"
                                           },
                                           {
                                               "TongCOD":  149.07,
                                               "CODTuNop":  141.54,
                                               "TyLeTuNop":  95,
                                               "TenBuuCuc":  "Bưu Cục 63 Đường Số 7-Xã Nhà Bè-HCM"
                                           },
                                           {
                                               "TongCOD":  349.11,
                                               "CODTuNop":  191.69,
                                               "TyLeTuNop":  54.9,
                                               "TenBuuCuc":  "Bưu Cục 87 Nguyễn An Ninh-Tp.Vũng Tàu-Bà Rịa-Vũng Tàu"
                                           },
                                           {
                                               "TongCOD":  165.36,
                                               "CODTuNop":  164.5,
                                               "TyLeTuNop":  99.5,
                                               "TenBuuCuc":  "Bưu cục 58 Vũ Tông Phan-Q.2-HCM"
                                           },
                                           {
                                               "TongCOD":  54.89,
                                               "CODTuNop":  54.89,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa 367 Lương Định Của-An Phú-Q.2-HCM"
                                           },
                                           {
                                               "TongCOD":  112.56,
                                               "CODTuNop":  29.3,
                                               "TyLeTuNop":  26,
                                               "TenBuuCuc":  "Bưu Cục Ấp Phú Thọ-Xuyên Mộc-Bà Rịa Vũng Tàu"
                                           },
                                           {
                                               "TongCOD":  255.99,
                                               "CODTuNop":  255.84,
                                               "TyLeTuNop":  99.9,
                                               "TenBuuCuc":  "Bưu Cục 119 Nguyễn Văn Luông-Q.6-HCM"
                                           },
                                           {
                                               "TongCOD":  162.35,
                                               "CODTuNop":  63.63,
                                               "TyLeTuNop":  39.2,
                                               "TenBuuCuc":  "Bưu Cục 57A Hùng Vương-Xã Ngãi Giao-Bà Rịa Vũng Tàu"
                                           },
                                           {
                                               "TongCOD":  269.97,
                                               "CODTuNop":  130.83,
                                               "TyLeTuNop":  48.5,
                                               "TenBuuCuc":  "Bưu Cục 79 Nguyễn Hiền-Dĩ An-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  93.48,
                                               "CODTuNop":  93.48,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa Q.Thủ Đức Kho 01 13 Đường Số 8-HCM"
                                           },
                                           {
                                               "TongCOD":  264.05,
                                               "CODTuNop":  93.47,
                                               "TyLeTuNop":  35.4,
                                               "TenBuuCuc":  "Bưu Cục 41 Nguyễn Tất Thành-Long Điền-Bà Rịa Vũng Tàu"
                                           },
                                           {
                                               "TongCOD":  157.18,
                                               "CODTuNop":  92.08,
                                               "TyLeTuNop":  58.6,
                                               "TenBuuCuc":  "Bưu Cục 6 Phạm Văn Hai-Q.Tân Bình-HCM"
                                           },
                                           {
                                               "TongCOD":  206.96,
                                               "CODTuNop":  205.79,
                                               "TyLeTuNop":  99.4,
                                               "TenBuuCuc":  "Bưu Cục 20/48A Bờ Bao Tân Thắng-Q.Tân Phú-HCM"
                                           },
                                           {
                                               "TongCOD":  101.01,
                                               "CODTuNop":  101.01,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa Quận 3 480F Cao Thắng-Q.10-HCM"
                                           },
                                           {
                                               "TongCOD":  21.34,
                                               "CODTuNop":  7.8,
                                               "TyLeTuNop":  36.5,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa Vĩnh Lộc B-Bình Chánh-HCM"
                                           },
                                           {
                                               "TongCOD":  136.64,
                                               "CODTuNop":  127.75,
                                               "TyLeTuNop":  93.5,
                                               "TenBuuCuc":  "Bưu Cục T30 Nguyễn Hữu Hào-Q.4-HCM"
                                           },
                                           {
                                               "TongCOD":  187.7,
                                               "CODTuNop":  92.55,
                                               "TyLeTuNop":  49.3,
                                               "TenBuuCuc":  "Bưu Cục 92 Tân Chánh Hiệp 03-Tân Chánh Hiệp-Quận 12-HCM"
                                           },
                                           {
                                               "TongCOD":  211.3,
                                               "CODTuNop":  110.1,
                                               "TyLeTuNop":  52.1,
                                               "TenBuuCuc":  "Bưu Cục 107 Đường 3D-Bình Hưng Hòa B-Bình Tân-HCM"
                                           },
                                           {
                                               "TongCOD":  136.46,
                                               "CODTuNop":  135.36,
                                               "TyLeTuNop":  99.2,
                                               "TenBuuCuc":  "Bưu Cục 119 Đường Số 1-Phường 11-Gò Vấp-HCM"
                                           },
                                           {
                                               "TongCOD":  359.4,
                                               "CODTuNop":  313.49,
                                               "TyLeTuNop":  87.2,
                                               "TenBuuCuc":  "Bưu Cục 20 Đường số 81-Q.7-HCM"
                                           },
                                           {
                                               "TongCOD":  125.64,
                                               "CODTuNop":  111.58,
                                               "TyLeTuNop":  88.8,
                                               "TenBuuCuc":  "Bưu Cục 165A Nguyễn Văn Tạo-Nhà Bè-HCM"
                                           },
                                           {
                                               "TongCOD":  303.21,
                                               "CODTuNop":  303.14,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 473A Bình Giã-TP.Vũng Tàu"
                                           },
                                           {
                                               "TongCOD":  181.04,
                                               "CODTuNop":  178.29,
                                               "TyLeTuNop":  98.5,
                                               "TenBuuCuc":  "Bưu Cục 59 Trương Phước Phan-Q.Bình Tân-HCM"
                                           },
                                           {
                                               "TongCOD":  240.49,
                                               "CODTuNop":  135.99,
                                               "TyLeTuNop":  56.5,
                                               "TenBuuCuc":  "Bưu Cục 1459 Đường 30/4-Vũng Tàu-Bà Rịa Vũng Tàu"
                                           },
                                           {
                                               "TongCOD":  128.48,
                                               "CODTuNop":  29.64,
                                               "TyLeTuNop":  23.1,
                                               "TenBuuCuc":  "Bưu Cục Hương Lộ 2-Bình Trị Đông A-Bình Tân"
                                           },
                                           {
                                               "TongCOD":  104.71,
                                               "CODTuNop":  60.41,
                                               "TyLeTuNop":  57.7,
                                               "TenBuuCuc":  "Bưu Cục KDC Giồng Ao-Cần Giờ-HCM"
                                           },
                                           {
                                               "TongCOD":  117.32,
                                               "CODTuNop":  117.32,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 575 Đường 22/12-Thuận An-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  133.65,
                                               "CODTuNop":  133.62,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu cục Hội Nghĩa-Tân Uyên-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  201.57,
                                               "CODTuNop":  190.53,
                                               "TyLeTuNop":  94.5,
                                               "TenBuuCuc":  "Bưu Cục Đường 743B-Thuận An-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  85.31,
                                               "CODTuNop":  12.34,
                                               "TyLeTuNop":  14.5,
                                               "TenBuuCuc":  "Bưu Cục 88 Võ Thị Liễu-Q.12-HCM"
                                           },
                                           {
                                               "TongCOD":  164.42,
                                               "CODTuNop":  124.04,
                                               "TyLeTuNop":  75.4,
                                               "TenBuuCuc":  "Bưu Cục 7A/3 Thành Thái-Phường Diên Hồng-HCM"
                                           },
                                           {
                                               "TongCOD":  197.46,
                                               "CODTuNop":  192.21,
                                               "TyLeTuNop":  97.3,
                                               "TenBuuCuc":  "Bưu Cục 10 Đường số 8-Linh Xuân-Q.Thủ Đức-HCM"
                                           },
                                           {
                                               "TongCOD":  238.69,
                                               "CODTuNop":  238.45,
                                               "TyLeTuNop":  99.9,
                                               "TenBuuCuc":  "Bưu Cục 29A Đường Số 5-Q.2-HCM"
                                           },
                                           {
                                               "TongCOD":  75.24,
                                               "CODTuNop":  64.05,
                                               "TyLeTuNop":  85.1,
                                               "TenBuuCuc":  "Bưu Cục 198 Thạnh Lộc 15-Q.12-HCM"
                                           },
                                           {
                                               "TongCOD":  98.62,
                                               "CODTuNop":  98.59,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 47 Đường 882-Phú Hữu-Q.9-HCM"
                                           },
                                           {
                                               "TongCOD":  279.45,
                                               "CODTuNop":  259.04,
                                               "TyLeTuNop":  92.7,
                                               "TenBuuCuc":  "Bưu Cục A4/18 Đường D10-Thuận An-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  155.49,
                                               "CODTuNop":  63,
                                               "TyLeTuNop":  40.5,
                                               "TenBuuCuc":  "Bưu Cục DT 744-An Tây-Bến Cát-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  226.78,
                                               "CODTuNop":  132.91,
                                               "TyLeTuNop":  58.6,
                                               "TenBuuCuc":  "Bưu Cục 235 Hồ Học Lãm-Q.Bình Tân-HCM"
                                           },
                                           {
                                               "TongCOD":  152.01,
                                               "CODTuNop":  151.39,
                                               "TyLeTuNop":  99.6,
                                               "TenBuuCuc":  "Bưu Cục 85 Cao Văn Ngọc-Phường Tân Phú-HCM"
                                           },
                                           {
                                               "TongCOD":  235.59,
                                               "CODTuNop":  75.82,
                                               "TyLeTuNop":  32.2,
                                               "TenBuuCuc":  "Bưu Cục KP Trảng Lớn-Phú Mỹ-Hắc Dịch-Bà Rịa Vũng Tàu"
                                           },
                                           {
                                               "TongCOD":  181.3,
                                               "CODTuNop":  180.64,
                                               "TyLeTuNop":  99.6,
                                               "TenBuuCuc":  "Bưu Cục 46 Tam Bình-Hiệp Bình Chánh-Q.Thủ Đức-HCM"
                                           },
                                           {
                                               "TongCOD":  170.86,
                                               "CODTuNop":  72.47,
                                               "TyLeTuNop":  42.4,
                                               "TenBuuCuc":  "Bưu Cục Khu Phố 4-Bắc Tân Uyên-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  302.75,
                                               "CODTuNop":  302.72,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Phường Lái Thiêu-Thuận An-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  260.45,
                                               "CODTuNop":  217.75,
                                               "TyLeTuNop":  83.6,
                                               "TenBuuCuc":  "Bưu Cục 376/38 Trảng Sắn-Phú Giáo-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  131.05,
                                               "CODTuNop":  75.5,
                                               "TyLeTuNop":  57.6,
                                               "TenBuuCuc":  "Bưu Cục Thôn Thành Long-Châu Đức-Bà Rịa Vũng Tàu"
                                           },
                                           {
                                               "TongCOD":  109.45,
                                               "CODTuNop":  109.38,
                                               "TyLeTuNop":  99.9,
                                               "TenBuuCuc":  "Bưu Cục 371 Nguyễn Kiệm-Q.Gò Vấp-HCM"
                                           },
                                           {
                                               "TongCOD":  171.06,
                                               "CODTuNop":  170.66,
                                               "TyLeTuNop":  99.8,
                                               "TenBuuCuc":  "Bưu Cục 63 Ký Hòa-Q.5-HCM"
                                           },
                                           {
                                               "TongCOD":  202.39,
                                               "CODTuNop":  74.67,
                                               "TyLeTuNop":  36.9,
                                               "TenBuuCuc":  "Bưu Cục Đường NL12-Phường Thới Hòa-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  132.89,
                                               "CODTuNop":  51.09,
                                               "TyLeTuNop":  38.4,
                                               "TenBuuCuc":  "Bưu Cục An Phước- An Ngãi - Bà Rịa Vũng Tàu"
                                           },
                                           {
                                               "TongCOD":  188.39,
                                               "CODTuNop":  25.74,
                                               "TyLeTuNop":  13.7,
                                               "TenBuuCuc":  "Bưu Cục 242 Tỉnh Lộ 2-Củ Chi-HCM"
                                           },
                                           {
                                               "TongCOD":  111.44,
                                               "CODTuNop":  108.57,
                                               "TyLeTuNop":  97.4,
                                               "TenBuuCuc":  "Bưu Cục 692 An Phú Tây-Bình Chánh-HCM"
                                           },
                                           {
                                               "TongCOD":  102.02,
                                               "CODTuNop":  98.16,
                                               "TyLeTuNop":  96.2,
                                               "TenBuuCuc":  "Bưu Cục 104 Lương Ngọc Quyến-Q.Bình Thạnh-HCM"
                                           },
                                           {
                                               "TongCOD":  216.69,
                                               "CODTuNop":  192.21,
                                               "TyLeTuNop":  88.7,
                                               "TenBuuCuc":  "Bưu Cục 16A Lê Hồng Phong-Q.10-HCM"
                                           },
                                           {
                                               "TongCOD":  201.23,
                                               "CODTuNop":  194.46,
                                               "TyLeTuNop":  96.6,
                                               "TenBuuCuc":  "Bưu Cục 8F Gò Xoài-Bình Hưng Hòa A-Bình Tân-HCM"
                                           },
                                           {
                                               "TongCOD":  143.76,
                                               "CODTuNop":  137.53,
                                               "TyLeTuNop":  95.7,
                                               "TenBuuCuc":  "Bưu Cục 58 Trương Quốc Dung-Q.Phú Nhuận-HCM"
                                           },
                                           {
                                               "TongCOD":  156.97,
                                               "CODTuNop":  156.94,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 870 Huỳnh Văn Lũy-Thủ Dầu Một-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  0.4,
                                               "CODTuNop":  0.4,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Chuyên Lấy 01 Đường Số 6-Long Bình-Q.9-HCM"
                                           },
                                           {
                                               "TongCOD":  189.72,
                                               "CODTuNop":  168.32,
                                               "TyLeTuNop":  88.7,
                                               "TenBuuCuc":  "Bưu Cục Tân Hiệp 04-Tân Uyên-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  168.99,
                                               "CODTuNop":  168.85,
                                               "TyLeTuNop":  99.9,
                                               "TenBuuCuc":  "Bưu Cục 116A Bà Hom-Phường 13-Quận 6-HCM"
                                           },
                                           {
                                               "TongCOD":  259.4,
                                               "CODTuNop":  186.17,
                                               "TyLeTuNop":  71.8,
                                               "TenBuuCuc":  "Bưu Cục 1440/1 Tỉnh Lộ 10-Tân Tạo-Bình Tân-HCM"
                                           },
                                           {
                                               "TongCOD":  205.74,
                                               "CODTuNop":  143.01,
                                               "TyLeTuNop":  69.5,
                                               "TenBuuCuc":  "Bưu Cục Côn Đảo-Bà Rịa Vũng Tàu"
                                           },
                                           {
                                               "TongCOD":  225.36,
                                               "CODTuNop":  225.36,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Phước Kiển-Nhà Bè-HCM"
                                           },
                                           {
                                               "TongCOD":  178.8,
                                               "CODTuNop":  178.66,
                                               "TyLeTuNop":  99.9,
                                               "TenBuuCuc":  "Bưu Cục Đường 21-Phường 08-Gò Vấp-HCM"
                                           },
                                           {
                                               "TongCOD":  157.64,
                                               "CODTuNop":  122.97,
                                               "TyLeTuNop":  78,
                                               "TenBuuCuc":  "Bưu Cục 593A Tỉnh Lộ 15-Củ Chi-HCM"
                                           },
                                           {
                                               "TongCOD":  95.89,
                                               "CODTuNop":  81.14,
                                               "TyLeTuNop":  84.6,
                                               "TenBuuCuc":  "Bưu Cục 58 Lê Thúc Hoạch-Q.Tân Phú-HCM"
                                           },
                                           {
                                               "TongCOD":  114.29,
                                               "CODTuNop":  102.72,
                                               "TyLeTuNop":  89.9,
                                               "TenBuuCuc":  "Bưu Cục 183 Nguyễn Oanh-Phường 17-Gò Vấp-HCM"
                                           },
                                           {
                                               "TongCOD":  109.17,
                                               "CODTuNop":  108.56,
                                               "TyLeTuNop":  99.4,
                                               "TenBuuCuc":  "Bưu Cục 47/2 Lê Văn Khương-Q.12-HCM"
                                           },
                                           {
                                               "TongCOD":  76.31,
                                               "CODTuNop":  17.45,
                                               "TyLeTuNop":  22.9,
                                               "TenBuuCuc":  "Bưu Cục Trừ Văn Thố-Bàu Bàng-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  222.76,
                                               "CODTuNop":  222.36,
                                               "TyLeTuNop":  99.8,
                                               "TenBuuCuc":  "Bưu Cục 31-31A Lý Phục Man-Bình Thuận-Quận 7-HCM"
                                           },
                                           {
                                               "TongCOD":  106.38,
                                               "CODTuNop":  106.31,
                                               "TyLeTuNop":  99.9,
                                               "TenBuuCuc":  "Bưu Cục 26 Bùi Thanh Khiết-Tân Túc-Bình Chánh"
                                           },
                                           {
                                               "TongCOD":  59.66,
                                               "CODTuNop":  37.42,
                                               "TyLeTuNop":  62.7,
                                               "TenBuuCuc":  "Bưu Cục 320 Nguyễn Thị Kiểu-Hiệp Thành-Quận 12-HCM"
                                           },
                                           {
                                               "TongCOD":  286.28,
                                               "CODTuNop":  286.28,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Xuyên Mộc-Bà Rịa Vũng Tàu"
                                           },
                                           {
                                               "TongCOD":  76.11,
                                               "CODTuNop":  76.11,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa 248A Nơ Trang Long-Q.Bình Thạnh-HCM"
                                           },
                                           {
                                               "TongCOD":  259.67,
                                               "CODTuNop":  259.63,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu cục D1-D39 KDC VietSing-An Phú-Thuận An-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  37.57,
                                               "CODTuNop":  37.57,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa Bình Tân Kho 01 1162/6A Tỉnh Lộ 10-HCM"
                                           },
                                           {
                                               "TongCOD":  138.16,
                                               "CODTuNop":  61.09,
                                               "TyLeTuNop":  44.2,
                                               "TenBuuCuc":  "Bưu Cục 125A Lê Lợi-Hóc Môn-HCM"
                                           },
                                           {
                                               "TongCOD":  264.5,
                                               "CODTuNop":  129.87,
                                               "TyLeTuNop":  49.1,
                                               "TenBuuCuc":  "Bưu Cục 106/3A Tây Lân-Hóc Môn-HCM"
                                           },
                                           {
                                               "TongCOD":  106.84,
                                               "CODTuNop":  57.07,
                                               "TyLeTuNop":  53.4,
                                               "TenBuuCuc":  "Bưu Cục 150 Đường DT743A-TP.Dĩ An-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  119.19,
                                               "CODTuNop":  118.77,
                                               "TyLeTuNop":  99.6,
                                               "TenBuuCuc":  "Bưu Cục 82 Nguyễn Sỹ Sách-Q.Tân Bình-HCM"
                                           },
                                           {
                                               "TongCOD":  111.07,
                                               "CODTuNop":  61.96,
                                               "TyLeTuNop":  55.8,
                                               "TenBuuCuc":  "Bưu Cục Tân Kiên-Bình Chánh-HCM"
                                           },
                                           {
                                               "TongCOD":  236.79,
                                               "CODTuNop":  174.51,
                                               "TyLeTuNop":  73.7,
                                               "TenBuuCuc":  "Bưu Cục Đông Chiêu-Tân Đông Hiệp-Dĩ An-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  20.32,
                                               "CODTuNop":  20.32,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa Bình Tân Kho 02 1162/6A Tỉnh Lộ 10-HCM"
                                           },
                                           {
                                               "TongCOD":  208.63,
                                               "CODTuNop":  208.63,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 257 Bùi Ngọc Thu-Chánh Hiệp-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  63.07,
                                               "CODTuNop":  60.96,
                                               "TyLeTuNop":  96.6,
                                               "TenBuuCuc":  "Bưu cục Hàng Vừa 78 Đường số 4-Q.7-HCM"
                                           },
                                           {
                                               "TongCOD":  206.03,
                                               "CODTuNop":  194.45,
                                               "TyLeTuNop":  94.4,
                                               "TenBuuCuc":  "Bưu Cục 456 Nguyễn Kiệm-Q.Phú Nhuận-HCM"
                                           },
                                           {
                                               "TongCOD":  164.1,
                                               "CODTuNop":  163.8,
                                               "TyLeTuNop":  99.8,
                                               "TenBuuCuc":  "Bưu Cục 78 Nguyễn Minh Hoàng-Phường 12-Q.Tân Bình-HCM"
                                           },
                                           {
                                               "TongCOD":  330.83,
                                               "CODTuNop":  330.65,
                                               "TyLeTuNop":  99.9,
                                               "TenBuuCuc":  "Bưu Cục 860/36 Huỳnh Tấn Phát-Q.7-HCM"
                                           },
                                           {
                                               "TongCOD":  173.9,
                                               "CODTuNop":  150.54,
                                               "TyLeTuNop":  86.6,
                                               "TenBuuCuc":  "Bưu Cục 197 Lê Cao Lãng-Phú Thạnh-Tân Phú-HCM"
                                           },
                                           {
                                               "TongCOD":  261.53,
                                               "CODTuNop":  261.53,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 291A Nguyễn Duy-Q.8-HCM"
                                           },
                                           {
                                               "TongCOD":  205.01,
                                               "CODTuNop":  130.75,
                                               "TyLeTuNop":  63.8,
                                               "TenBuuCuc":  "Bưu Cục 1855 Võ Văn Kiệt-Hoà Long-Bà Rịa Vũng Tàu"
                                           },
                                           {
                                               "TongCOD":  121.93,
                                               "CODTuNop":  112.31,
                                               "TyLeTuNop":  92.1,
                                               "TenBuuCuc":  "Bưu Cục 1417 Đường N4-D2 KDC Thanh Duy-TP.Tân Uyên-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  147.93,
                                               "CODTuNop":  147.93,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Đường Số 2 KDC Genimex-TP.Tân Uyên-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  253.94,
                                               "CODTuNop":  253.13,
                                               "TyLeTuNop":  99.7,
                                               "TenBuuCuc":  "Bưu Cục 903B Tạ Quang Bửu-Q.8-HCM"
                                           },
                                           {
                                               "TongCOD":  234.17,
                                               "CODTuNop":  233.95,
                                               "TyLeTuNop":  99.9,
                                               "TenBuuCuc":  "Bưu Cục 75 Đường Số 2 KDC Nguyên Sơn-Bình Hưng-Bình Chánh"
                                           },
                                           {
                                               "TongCOD":  162.34,
                                               "CODTuNop":  161.92,
                                               "TyLeTuNop":  99.7,
                                               "TenBuuCuc":  "Bưu Cục 16 Trần Nhật Duật-Q.1-HCM"
                                           },
                                           {
                                               "TongCOD":  143.99,
                                               "CODTuNop":  143.74,
                                               "TyLeTuNop":  99.8,
                                               "TenBuuCuc":  "Bưu Cục 1130 Đông Hưng Thuận 05-Q.12-HCM"
                                           },
                                           {
                                               "TongCOD":  306.9,
                                               "CODTuNop":  306.9,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Bàu Bàng-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  209.92,
                                               "CODTuNop":  209.83,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 24B Nguyễn Du-Phường 1-TP.Vũng Tàu-Bà Rịa Vũng Tàu"
                                           },
                                           {
                                               "TongCOD":  49.62,
                                               "CODTuNop":  25.27,
                                               "TyLeTuNop":  50.9,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa 20A Đường 12-Q.9-HCM"
                                           },
                                           {
                                               "TongCOD":  163.67,
                                               "CODTuNop":  69.68,
                                               "TyLeTuNop":  42.6,
                                               "TenBuuCuc":  "Bưu Cục KDC Đồng Chàm-Dĩ An-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  91.31,
                                               "CODTuNop":  91.31,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa 85 Cao Văn Ngọc-Q.Tân Phú-HCM"
                                           },
                                           {
                                               "TongCOD":  254.62,
                                               "CODTuNop":  125.23,
                                               "TyLeTuNop":  49.2,
                                               "TenBuuCuc":  "Bưu cục Số 54/1 Đường Tân Phước Khánh 13-Tân Uyên-Bình Dương"
                                           },
                                           {
                                               "TongCOD":  195.4,
                                               "CODTuNop":  115.77,
                                               "TyLeTuNop":  59.2,
                                               "TenBuuCuc":  "Bưu Cục 63 Vĩnh Hội-Q.4-HCM"
                                           }
                                       ]
                    },
                    {
                        "TyLeTuNop":  38.3,
                        "TenVung":  "HNO",
                        "TongCOD":  16643.79,
                        "Trend":  9.5,
                        "CODTuNop":  6370.38,
                        "BuuCucList":  [
                                           {
                                               "TongCOD":  220.54,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 107 Nguyễn Chí Thanh-Q.Đống Đa-HN"
                                           },
                                           {
                                               "TongCOD":  109.49,
                                               "CODTuNop":  80.08,
                                               "TyLeTuNop":  73.1,
                                               "TenBuuCuc":  "Bưu Cục 51A Tiến Thành-Q.Hà Đông-HN"
                                           },
                                           {
                                               "TongCOD":  20.59,
                                               "CODTuNop":  20.59,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Cồng Kềnh 1057/9/26 Ngãi Cầu-Hoài Đức-HN"
                                           },
                                           {
                                               "TongCOD":  174.28,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 77 Nguyễn Vĩnh Tích-Thường Tín-HN"
                                           },
                                           {
                                               "TongCOD":  138.35,
                                               "CODTuNop":  18.75,
                                               "TyLeTuNop":  13.6,
                                               "TenBuuCuc":  "Bưu Cục 17 Phú Thượng-Q.Tây Hồ-HN"
                                           },
                                           {
                                               "TongCOD":  139.02,
                                               "CODTuNop":  2.05,
                                               "TyLeTuNop":  1.5,
                                               "TenBuuCuc":  "Bưu Cục 1057/9/26 Ngãi Cầu-Hoài Đức-HN"
                                           },
                                           {
                                               "TongCOD":  186.72,
                                               "CODTuNop":  90.77,
                                               "TyLeTuNop":  48.6,
                                               "TenBuuCuc":  "Bưu Cục Đường Băng Xuân Bách-Xã Nội Bài-HN"
                                           },
                                           {
                                               "TongCOD":  38.49,
                                               "CODTuNop":  35.06,
                                               "TyLeTuNop":  91.1,
                                               "TenBuuCuc":  "Bưu Cục 450 Bờ Tây Sông Nhuệ-Q.Bắc Từ Liêm-HN"
                                           },
                                           {
                                               "TongCOD":  137.24,
                                               "CODTuNop":  81.33,
                                               "TyLeTuNop":  59.3,
                                               "TenBuuCuc":  "Bưu Cục 84 Đặng Xuân Bảng-Đại Kim-Hoàng Mai-HN"
                                           },
                                           {
                                               "TongCOD":  0.47,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 85C Nguyễn Khang-Q.Cầu Giấy-HN"
                                           },
                                           {
                                               "TongCOD":  149.98,
                                               "CODTuNop":  149.98,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 20 Đông Quan-Q.Cầu Giấy-HN"
                                           },
                                           {
                                               "TongCOD":  154.63,
                                               "CODTuNop":  55.52,
                                               "TyLeTuNop":  35.9,
                                               "TenBuuCuc":  "Bưu cục Thôn 6-Ba Vì-HN"
                                           },
                                           {
                                               "TongCOD":  230.18,
                                               "CODTuNop":  67.95,
                                               "TyLeTuNop":  29.5,
                                               "TenBuuCuc":  "Bưu Cục 90 Trung Phụng-Q.Đống Đa-HN"
                                           },
                                           {
                                               "TongCOD":  121.61,
                                               "CODTuNop":  79.98,
                                               "TyLeTuNop":  65.8,
                                               "TenBuuCuc":  "Bưu Cục 57 Thao Chính-Phú Xuyên-HN"
                                           },
                                           {
                                               "TongCOD":  145.36,
                                               "CODTuNop":  145.36,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa 37 TT9 KĐT Văn Phú-Q.Hà Đông-HN"
                                           },
                                           {
                                               "TongCOD":  150.01,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Chuyên Giao 142 Hồng Đô-Q.Nam Từ Liêm-HN"
                                           },
                                           {
                                               "TongCOD":  161.68,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 10BT4 KĐT Văn Quán-Phúc La-Hà Đông-HN"
                                           },
                                           {
                                               "TongCOD":  154.5,
                                               "CODTuNop":  2.35,
                                               "TyLeTuNop":  1.5,
                                               "TenBuuCuc":  "Bưu Cục 172 Võ Chí Công-Xuân La-Tây Hồ-HN"
                                           },
                                           {
                                               "TongCOD":  131.03,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 22 Thanh Yên-Q.Hoàn Kiếm-HN"
                                           },
                                           {
                                               "TongCOD":  213.86,
                                               "CODTuNop":  207.35,
                                               "TyLeTuNop":  97,
                                               "TenBuuCuc":  "Bưu Cục 984 Đê La Thành-Ba Đình-HN"
                                           },
                                           {
                                               "TongCOD":  61.41,
                                               "CODTuNop":  60.02,
                                               "TyLeTuNop":  97.7,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa 450A Bờ Tây Sông Nhuệ-Q.Bắc Từ Liêm-HN"
                                           },
                                           {
                                               "TongCOD":  210.8,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 416 La Thành-Sơn Tây-HN"
                                           },
                                           {
                                               "TongCOD":  66.81,
                                               "CODTuNop":  31.4,
                                               "TyLeTuNop":  47,
                                               "TenBuuCuc":  "Bưu Cục 709 Tam Trinh-Q.Hoàng Mai-HN"
                                           },
                                           {
                                               "TongCOD":  178.31,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 10/3 Thôn Ninh Dương-Thanh Oai-HN"
                                           },
                                           {
                                               "TongCOD":  15.18,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Cồng Kềnh 272 Nguyễn Xiển-Thanh Xuân-HN"
                                           },
                                           {
                                               "TongCOD":  117.28,
                                               "CODTuNop":  27.72,
                                               "TyLeTuNop":  23.6,
                                               "TenBuuCuc":  "Bưu Cục 39 Gò Sỏi-Sóc Sơn-HN"
                                           },
                                           {
                                               "TongCOD":  59.34,
                                               "CODTuNop":  59.34,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Cồng Kềnh Khu Công Nghiệp HNT-Đông Anh-HN"
                                           },
                                           {
                                               "TongCOD":  200.38,
                                               "CODTuNop":  200.29,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu cục 148 Trâu Quỳ-Gia Lâm-HN"
                                           },
                                           {
                                               "TongCOD":  152.19,
                                               "CODTuNop":  80.29,
                                               "TyLeTuNop":  52.8,
                                               "TenBuuCuc":  "Bưu Cục Chúc Lý-Ngọc Hòa-Chương Mỹ-HN"
                                           },
                                           {
                                               "TongCOD":  166.23,
                                               "CODTuNop":  160.86,
                                               "TyLeTuNop":  96.8,
                                               "TenBuuCuc":  "Bưu Cục Tổ 5 Giang Biên-Long Biên-HN"
                                           },
                                           {
                                               "TongCOD":  157.44,
                                               "CODTuNop":  68.3,
                                               "TyLeTuNop":  43.4,
                                               "TenBuuCuc":  "Bưu cục 8.03-TT8 Xuân Phương-Q.Nam Từ Liêm-HN"
                                           },
                                           {
                                               "TongCOD":  143.8,
                                               "CODTuNop":  61.24,
                                               "TyLeTuNop":  42.6,
                                               "TenBuuCuc":  "Bưu Cục Số 1 Tân Khai-Q.Hoàng Mai-HN"
                                           },
                                           {
                                               "TongCOD":  51.49,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 53 Lãng Yên-Q.Hai Bà Trưng-HN"
                                           },
                                           {
                                               "TongCOD":  122.98,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 3 Thôn Cầu Thăng Long-Đông Anh-HN"
                                           },
                                           {
                                               "TongCOD":  64.33,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu cục Ngọc Hồi-Thanh Trì-HN"
                                           },
                                           {
                                               "TongCOD":  123.27,
                                               "CODTuNop":  23.86,
                                               "TyLeTuNop":  19.4,
                                               "TenBuuCuc":  "Bưu Cục 8B Nguyễn Lân-Q.Thanh Xuân-HN"
                                           },
                                           {
                                               "TongCOD":  158.6,
                                               "CODTuNop":  105.62,
                                               "TyLeTuNop":  66.6,
                                               "TenBuuCuc":  "Bưu cục 896 Minh Khai-Q.Hai Bà Trưng-HN"
                                           },
                                           {
                                               "TongCOD":  0.78,
                                               "CODTuNop":  0.02,
                                               "TyLeTuNop":  3.1,
                                               "TenBuuCuc":  "Bưu Cục 131 Quan Hoa-Q.Cầu Giấy-HN"
                                           },
                                           {
                                               "TongCOD":  86.41,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 323 Lại Dụ An Thượng-Hoài Đức-HN"
                                           },
                                           {
                                               "TongCOD":  166.55,
                                               "CODTuNop":  92.38,
                                               "TyLeTuNop":  55.5,
                                               "TenBuuCuc":  "Bưu Cục Thôn An Khoái-Phúc Tiến-Phú Xuyên-HN"
                                           },
                                           {
                                               "TongCOD":  127.28,
                                               "CODTuNop":  127.28,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 36 Lý Chiêu Hoàng-Ninh Hiệp-Gia Lâm-Hà Nội"
                                           },
                                           {
                                               "TongCOD":  154.04,
                                               "CODTuNop":  151.66,
                                               "TyLeTuNop":  98.5,
                                               "TenBuuCuc":  "Bưu cục 442 Nguyễn Trãi-Q.Thanh Xuân-HN"
                                           },
                                           {
                                               "TongCOD":  133.13,
                                               "CODTuNop":  8.21,
                                               "TyLeTuNop":  6.2,
                                               "TenBuuCuc":  "Bưu cục Thạch Hòa-Thạch Thất-Hà Nội"
                                           },
                                           {
                                               "TongCOD":  146.56,
                                               "CODTuNop":  87.91,
                                               "TyLeTuNop":  60,
                                               "TenBuuCuc":  "Bưu Cục Lô DM10-6 Làng nghề Vạn Phúc-Q.Hà Đông-HN"
                                           },
                                           {
                                               "TongCOD":  5.49,
                                               "CODTuNop":  3.01,
                                               "TyLeTuNop":  54.8,
                                               "TenBuuCuc":  "Bưu Cục 14 Huỳnh Thúc Kháng-Q.Đống Đa-HN"
                                           },
                                           {
                                               "TongCOD":  28.62,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa Tổng Cục 5 Yên Xá-Huyện Thanh Trì-HN"
                                           },
                                           {
                                               "TongCOD":  9.93,
                                               "CODTuNop":  9.3,
                                               "TyLeTuNop":  93.6,
                                               "TenBuuCuc":  "Bưu Cục 882 Kim Giang-Thanh Trì-HN"
                                           },
                                           {
                                               "TongCOD":  0.35,
                                               "CODTuNop":  0.35,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Cồng Kềnh Cổ Dương-Đông Anh-HN"
                                           },
                                           {
                                               "TongCOD":  84.09,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Số Nhà F8-Khu Đấu Giá Phú Lương 1-Phú Lương-HN"
                                           },
                                           {
                                               "TongCOD":  46.61,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục TT03 Tây Nam Linh Đàm-Q.Hoàng Mai-HN"
                                           },
                                           {
                                               "TongCOD":  101.34,
                                               "CODTuNop":  72.03,
                                               "TyLeTuNop":  71.1,
                                               "TenBuuCuc":  "Bưu Cục 139 Thụy Phương-Q.Bắc Từ Liêm-HN"
                                           },
                                           {
                                               "TongCOD":  49.41,
                                               "CODTuNop":  34.74,
                                               "TyLeTuNop":  70.3,
                                               "TenBuuCuc":  "Bưu Cục 29 Hưng Phúc-Yên Sở-Q.Hoàng Mai-HN"
                                           },
                                           {
                                               "TongCOD":  106.14,
                                               "CODTuNop":  56.5,
                                               "TyLeTuNop":  53.2,
                                               "TenBuuCuc":  "Bưu Cục Khu tái định cư X6-Thạch Bàn-HN"
                                           },
                                           {
                                               "TongCOD":  87.86,
                                               "CODTuNop":  45.45,
                                               "TyLeTuNop":  51.7,
                                               "TenBuuCuc":  "Bưu Cục 79/9B Thanh Đàm-Q.Hoàng Mai-HN"
                                           },
                                           {
                                               "TongCOD":  218.41,
                                               "CODTuNop":  73.49,
                                               "TyLeTuNop":  33.7,
                                               "TenBuuCuc":  "Bưu cục 41 Trần Quốc Vượng-Q.Cầu giấy-HN"
                                           },
                                           {
                                               "TongCOD":  82.04,
                                               "CODTuNop":  81.75,
                                               "TyLeTuNop":  99.7,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa 1267 Giải Phóng-Q.Hoàng Mai-HN"
                                           },
                                           {
                                               "TongCOD":  188.82,
                                               "CODTuNop":  101.8,
                                               "TyLeTuNop":  53.9,
                                               "TenBuuCuc":  "Bưu cục 126 Hoàng Ngân-Q.Cầu Giấy-HN"
                                           },
                                           {
                                               "TongCOD":  189.41,
                                               "CODTuNop":  26.55,
                                               "TyLeTuNop":  14,
                                               "TenBuuCuc":  "Bưu Cục 55C Nguyễn Khoái-Q.Hai Bà Trưng-HN"
                                           },
                                           {
                                               "TongCOD":  69.49,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Đường Sông Nhuệ-Thường Tín-HN"
                                           },
                                           {
                                               "TongCOD":  94.6,
                                               "CODTuNop":  94.06,
                                               "TyLeTuNop":  99.4,
                                               "TenBuuCuc":  "Bưu Cục 849 Ngô Gia Tự-Đức Giang-Long Biên-HN"
                                           },
                                           {
                                               "TongCOD":  106.17,
                                               "CODTuNop":  7.33,
                                               "TyLeTuNop":  6.9,
                                               "TenBuuCuc":  "Bưu Cục 108 Phố Nhổn-Q.Bắc Từ Liêm-Hà Nội"
                                           },
                                           {
                                               "TongCOD":  176.2,
                                               "CODTuNop":  176.2,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 63 Đường Bờ Sông Sét-Thịnh Liệt-Hoàng Mai-HN"
                                           },
                                           {
                                               "TongCOD":  107.14,
                                               "CODTuNop":  56.92,
                                               "TyLeTuNop":  53.1,
                                               "TenBuuCuc":  "Bưu Cục 114 Yên Lãng-Q.Đống Đa-HN"
                                           },
                                           {
                                               "TongCOD":  0.81,
                                               "CODTuNop":  0.69,
                                               "TyLeTuNop":  85.4,
                                               "TenBuuCuc":  "Bưu Cục 7C10 Nguyễn Phong Sắc-Cầu Giấy-Hà Nội"
                                           },
                                           {
                                               "TongCOD":  121.08,
                                               "CODTuNop":  55.86,
                                               "TyLeTuNop":  46.1,
                                               "TenBuuCuc":  "Bưu Cục 155 Phú Diễn-Phú Diễn-Bắc Từ Liêm-HN"
                                           },
                                           {
                                               "TongCOD":  11.56,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu cục Yên Xá-Thanh Trì-HN"
                                           },
                                           {
                                               "TongCOD":  163.33,
                                               "CODTuNop":  60.97,
                                               "TyLeTuNop":  37.3,
                                               "TenBuuCuc":  "Bưu Cục 298 Trần Điền-Định Công-Hoàng Mai-Hà Nội"
                                           },
                                           {
                                               "TongCOD":  121.43,
                                               "CODTuNop":  18.84,
                                               "TyLeTuNop":  15.5,
                                               "TenBuuCuc":  "Bưu Cục 325 LouisI-Q.Hoàng Mai-HN"
                                           },
                                           {
                                               "TongCOD":  53.08,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Thôn Hải Bối-Xã Vĩnh Thanh-HN"
                                           },
                                           {
                                               "TongCOD":  36.91,
                                               "CODTuNop":  36.77,
                                               "TyLeTuNop":  99.6,
                                               "TenBuuCuc":  "Bưu Cục Cồng Kềnh 51A Tiến Thành-Q.Hà Đông-HN"
                                           },
                                           {
                                               "TongCOD":  104.58,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 09 Hoàng Cầu-Q.Đống Đa-HN"
                                           },
                                           {
                                               "TongCOD":  133.84,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 19 Đê Quai-Q.Tây Hồ-HN"
                                           },
                                           {
                                               "TongCOD":  115.92,
                                               "CODTuNop":  100.91,
                                               "TyLeTuNop":  87.1,
                                               "TenBuuCuc":  "Bưu Cục 60 Sài Đồng-Q.Long Biên-Hà Nội"
                                           },
                                           {
                                               "TongCOD":  14.88,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục C14 Bắc Hà-Q.Nam Từ Liêm-HN"
                                           },
                                           {
                                               "TongCOD":  3.63,
                                               "CODTuNop":  0.51,
                                               "TyLeTuNop":  13.9,
                                               "TenBuuCuc":  "Bưu Cục 61 Kim Mã-Q.Ba Đình-HN"
                                           },
                                           {
                                               "TongCOD":  259.19,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu cục Quốc Oai-HN"
                                           },
                                           {
                                               "TongCOD":  134.43,
                                               "CODTuNop":  37.56,
                                               "TyLeTuNop":  27.9,
                                               "TenBuuCuc":  "Bưu Cục 88 Bạch Đằng-Hoàn Kiếm-HN"
                                           },
                                           {
                                               "TongCOD":  61.45,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 126 Hồng Đô-Q.Nam Từ Liêm-HN"
                                           },
                                           {
                                               "TongCOD":  144.95,
                                               "CODTuNop":  144.72,
                                               "TyLeTuNop":  99.8,
                                               "TenBuuCuc":  "Bưu cục 296 đường Ỷ Lan-Gia Lâm-HN"
                                           },
                                           {
                                               "TongCOD":  94.58,
                                               "CODTuNop":  20.91,
                                               "TyLeTuNop":  22.1,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa 96 Nghĩa Tân-Q.Cầu Giấy-HN"
                                           },
                                           {
                                               "TongCOD":  191.04,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 219 Sơn Đồng-Hoài Đức-HN"
                                           },
                                           {
                                               "TongCOD":  242.1,
                                               "CODTuNop":  32.02,
                                               "TyLeTuNop":  13.2,
                                               "TenBuuCuc":  "Bưu cục Thôn Đông-Ba Vì-Hà Nội"
                                           },
                                           {
                                               "TongCOD":  6.92,
                                               "CODTuNop":  6.92,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Cồng Kềnh thôn Hải Bối-xã Vĩnh Thanh-HN"
                                           },
                                           {
                                               "TongCOD":  121.1,
                                               "CODTuNop":  121.1,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Cồng Kềnh 270 Nguyễn Xiển-Thanh Xuân-HN"
                                           },
                                           {
                                               "TongCOD":  185.75,
                                               "CODTuNop":  26.02,
                                               "TyLeTuNop":  14,
                                               "TenBuuCuc":  "Bưu Cục Xã Thắng Lợi-Thường Tín-HN"
                                           },
                                           {
                                               "TongCOD":  121.07,
                                               "CODTuNop":  71.3,
                                               "TyLeTuNop":  58.9,
                                               "TenBuuCuc":  "Bưu Cục 33 Đặng Trần Côn-Q.Đống Đa-HN"
                                           },
                                           {
                                               "TongCOD":  115.25,
                                               "CODTuNop":  23.98,
                                               "TyLeTuNop":  20.8,
                                               "TenBuuCuc":  "Bưu Cục 366 Phạm Văn Đồng-Q.Bắc Từ Liêm-HN"
                                           },
                                           {
                                               "TongCOD":  177.66,
                                               "CODTuNop":  89.12,
                                               "TyLeTuNop":  50.2,
                                               "TenBuuCuc":  "Bưu Cục 26 Đại Mỗ-Q.Nam Từ Liêm-HN"
                                           },
                                           {
                                               "TongCOD":  2.23,
                                               "CODTuNop":  1.92,
                                               "TyLeTuNop":  86.1,
                                               "TenBuuCuc":  "Bưu Cục 10 Vĩnh Tuy-Q.Hai Bà Trưng-HN"
                                           },
                                           {
                                               "TongCOD":  62.42,
                                               "CODTuNop":  12.37,
                                               "TyLeTuNop":  19.8,
                                               "TenBuuCuc":  "Bưu Cục 396 Kim Giang-Q.Thanh Xuân-HN"
                                           },
                                           {
                                               "TongCOD":  117.25,
                                               "CODTuNop":  49.85,
                                               "TyLeTuNop":  42.5,
                                               "TenBuuCuc":  "Bưu Cục 43 Trần Đăng Ninh-Ứng Hòa-HN"
                                           },
                                           {
                                               "TongCOD":  0.9,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 631 Phúc Diễn-Q.Nam Từ Liêm-HN"
                                           },
                                           {
                                               "TongCOD":  7.37,
                                               "CODTuNop":  7.27,
                                               "TyLeTuNop":  98.6,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa 4 KĐT Sinh Thái Xuân Phương-Q.Nam Từ Liêm-HN"
                                           },
                                           {
                                               "TongCOD":  17.6,
                                               "CODTuNop":  4.51,
                                               "TyLeTuNop":  25.7,
                                               "TenBuuCuc":  "Bưu Cục Cồng Kềnh 17 Phố Phú Thượng-Q.Tây Hồ-HN"
                                           },
                                           {
                                               "TongCOD":  147.93,
                                               "CODTuNop":  44.79,
                                               "TyLeTuNop":  30.3,
                                               "TenBuuCuc":  "Bưu Cục Thôn Phú Mỹ-Tự Lập-Mê Linh-HN"
                                           },
                                           {
                                               "TongCOD":  109.42,
                                               "CODTuNop":  59.96,
                                               "TyLeTuNop":  54.8,
                                               "TenBuuCuc":  "Bưu Cục 28 Đường Mới-Phụng Thượng-Phúc Thọ-HN"
                                           },
                                           {
                                               "TongCOD":  182.63,
                                               "CODTuNop":  165.41,
                                               "TyLeTuNop":  90.6,
                                               "TenBuuCuc":  "Bưu Cục 229 Quan Nhân-Q.Thanh Xuân-HN"
                                           },
                                           {
                                               "TongCOD":  85.24,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Cồng Kềnh 268 Trần Điền-Q.Hoàng Mai-HN"
                                           },
                                           {
                                               "TongCOD":  44.12,
                                               "CODTuNop":  10.89,
                                               "TyLeTuNop":  24.7,
                                               "TenBuuCuc":  "Bưu Cục Thọ Giáo-Tân Minh-Thường Tín-HN"
                                           },
                                           {
                                               "TongCOD":  92.57,
                                               "CODTuNop":  62.22,
                                               "TyLeTuNop":  67.2,
                                               "TenBuuCuc":  "Bưu Cục 41 Phố Nguyễn Khả Trạc-Q.Cầu Giấy-HN"
                                           },
                                           {
                                               "TongCOD":  113.28,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Mỹ Đình 2-Q.Nam Từ Liêm-HN"
                                           },
                                           {
                                               "TongCOD":  208.55,
                                               "CODTuNop":  126.17,
                                               "TyLeTuNop":  60.5,
                                               "TenBuuCuc":  "Bưu Cục Đại Nghĩa-Mỹ Đức-HN"
                                           },
                                           {
                                               "TongCOD":  109.98,
                                               "CODTuNop":  5.07,
                                               "TyLeTuNop":  4.6,
                                               "TenBuuCuc":  "Bưu Cục 109 E3-Q.Hai Bà Trưng-HN"
                                           },
                                           {
                                               "TongCOD":  3.13,
                                               "CODTuNop":  2.4,
                                               "TyLeTuNop":  76.8,
                                               "TenBuuCuc":  "Bưu Cục 142A Khuất Duy Tiến-Q.Thanh Xuân-HN"
                                           },
                                           {
                                               "TongCOD":  80.5,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Ngõ 26 Xóm Vang-Xã Đông Anh-HN"
                                           },
                                           {
                                               "TongCOD":  203.02,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu cục Trạm Trôi-Hoài Đức-Hà Nội"
                                           },
                                           {
                                               "TongCOD":  2.95,
                                               "CODTuNop":  2.92,
                                               "TyLeTuNop":  99,
                                               "TenBuuCuc":  "Bưu Cục 263 Phố Vọng-Q.Hai Bà Trưng-HN"
                                           },
                                           {
                                               "TongCOD":  198.89,
                                               "CODTuNop":  30.66,
                                               "TyLeTuNop":  15.4,
                                               "TenBuuCuc":  "Bưu Cục Đường Núi Đôi-Sóc Sơn-HN"
                                           },
                                           {
                                               "TongCOD":  205.19,
                                               "CODTuNop":  59.87,
                                               "TyLeTuNop":  29.2,
                                               "TenBuuCuc":  "Bưu Cục Thôn Đại Bái-Đại Thịnh-Mê Linh-HNO"
                                           },
                                           {
                                               "TongCOD":  138.56,
                                               "CODTuNop":  71.81,
                                               "TyLeTuNop":  51.8,
                                               "TenBuuCuc":  "Bưu Cục 35 Liền Kề 9 KĐT Đại Thanh-Thanh Trì-HN"
                                           },
                                           {
                                               "TongCOD":  100.67,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Khu Đất Dịch Vụ Hòa Bình-Yên Nghĩa-Hà Đông-HN"
                                           },
                                           {
                                               "TongCOD":  145.61,
                                               "CODTuNop":  12.58,
                                               "TyLeTuNop":  8.6,
                                               "TenBuuCuc":  "Bưu cục 253 Phạm Văn Đồng-Q.Bắc Từ Liêm-Hà Nội"
                                           },
                                           {
                                               "TongCOD":  143.47,
                                               "CODTuNop":  75.26,
                                               "TyLeTuNop":  52.5,
                                               "TenBuuCuc":  "Bưu Cục 24 ngõ 80 La Nội-Q.Hà Đông-HN"
                                           },
                                           {
                                               "TongCOD":  210.96,
                                               "CODTuNop":  210.96,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục TT1 Khu tái định cư Ngọc Thụy-Long Biên-HN"
                                           },
                                           {
                                               "TongCOD":  39.98,
                                               "CODTuNop":  39.76,
                                               "TyLeTuNop":  99.4,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa 30 Phố Xã Đàn-Q.Đống Đa-HN"
                                           },
                                           {
                                               "TongCOD":  172.25,
                                               "CODTuNop":  78.09,
                                               "TyLeTuNop":  45.3,
                                               "TenBuuCuc":  "Bưu Cục 16A Hòa Mã -Q.Hai Bà Trưng-HN"
                                           },
                                           {
                                               "TongCOD":  154.65,
                                               "CODTuNop":  11.46,
                                               "TyLeTuNop":  7.4,
                                               "TenBuuCuc":  "Bưu Cục 172 Cầu Gáo-Đan Phượng-HN"
                                           },
                                           {
                                               "TongCOD":  133.27,
                                               "CODTuNop":  55.37,
                                               "TyLeTuNop":  41.6,
                                               "TenBuuCuc":  "Bưu Cục Hà Trì-Hà Cầu-Hà Đông-HN"
                                           },
                                           {
                                               "TongCOD":  154.56,
                                               "CODTuNop":  27.9,
                                               "TyLeTuNop":  18.1,
                                               "TenBuuCuc":  "Bưu Cục 203 Tỉnh Lộ 417-Phúc Thọ-HN"
                                           },
                                           {
                                               "TongCOD":  108.15,
                                               "CODTuNop":  6.58,
                                               "TyLeTuNop":  6.1,
                                               "TenBuuCuc":  "Bưu Cục 79 Thanh Xuân Làng Nghề Liên Hà-Đan Phượng-HN"
                                           },
                                           {
                                               "TongCOD":  381.73,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu cục 171 Đồng Cam-Thạch Thất-Hà Nội"
                                           },
                                           {
                                               "TongCOD":  170.17,
                                               "CODTuNop":  148.63,
                                               "TyLeTuNop":  87.3,
                                               "TenBuuCuc":  "Bưu Cục Lưu Quang Vũ-Trung Hòa-Cầu Giấy-HN"
                                           },
                                           {
                                               "TongCOD":  27.96,
                                               "CODTuNop":  27.03,
                                               "TyLeTuNop":  96.7,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa 339 Trâu Quỳ-H.Gia Lâm-HN"
                                           },
                                           {
                                               "TongCOD":  5.44,
                                               "CODTuNop":  5.44,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa 54 Ngõ 80 La Nội-Phường Dương Nội-HN"
                                           },
                                           {
                                               "TongCOD":  135.76,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 221 Kim Ngưu-Q.Hai Bà Trưng-HN"
                                           },
                                           {
                                               "TongCOD":  199.04,
                                               "CODTuNop":  7.26,
                                               "TyLeTuNop":  3.6,
                                               "TenBuuCuc":  "Bưu Cục 39 KTĐC Tứ Hiệp-Thanh Trì-HN"
                                           },
                                           {
                                               "TongCOD":  168.68,
                                               "CODTuNop":  122.13,
                                               "TyLeTuNop":  72.4,
                                               "TenBuuCuc":  "Bưu Cục 15 Hàng Bún-Q.Ba Đình-HN"
                                           },
                                           {
                                               "TongCOD":  122.3,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 97 Thiên Hiền-Q.Nam Từ Liêm-HN"
                                           },
                                           {
                                               "TongCOD":  280.79,
                                               "CODTuNop":  81.59,
                                               "TyLeTuNop":  29.1,
                                               "TenBuuCuc":  "Bưu Cục Cổ Dương-Tiên Dương-Đông Anh-HN"
                                           },
                                           {
                                               "TongCOD":  372.06,
                                               "CODTuNop":  11.59,
                                               "TyLeTuNop":  3.1,
                                               "TenBuuCuc":  "Bưu Cục 433 Tân Phú-Sơn Tây-HN"
                                           },
                                           {
                                               "TongCOD":  126.18,
                                               "CODTuNop":  79.75,
                                               "TyLeTuNop":  63.2,
                                               "TenBuuCuc":  "Bưu Cục 188 Khương Đình-Q.Thanh Xuân-HN"
                                           },
                                           {
                                               "TongCOD":  121.6,
                                               "CODTuNop":  101.94,
                                               "TyLeTuNop":  83.8,
                                               "TenBuuCuc":  "Bưu Cục Thôn Mới-Cao Dương-Thanh Oai-HN"
                                           },
                                           {
                                               "TongCOD":  135.78,
                                               "CODTuNop":  97.86,
                                               "TyLeTuNop":  72.1,
                                               "TenBuuCuc":  "Bưu Cục Bê Tông-Thuỷ Xuân Tiên-Chương Mỹ-HN"
                                           },
                                           {
                                               "TongCOD":  155.23,
                                               "CODTuNop":  90.04,
                                               "TyLeTuNop":  58,
                                               "TenBuuCuc":  "Bưu Cục Hòa Lâm-Ứng Hòa-Hà Nội"
                                           },
                                           {
                                               "TongCOD":  1.19,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu cục 36 TT9 KĐT Văn Phú-Q.Hà Đông-HN"
                                           },
                                           {
                                               "TongCOD":  69.47,
                                               "CODTuNop":  69.47,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hàng Vừa G4.P2 Cổ Linh-Q.Long Biên-HN"
                                           },
                                           {
                                               "TongCOD":  83.56,
                                               "CODTuNop":  9.02,
                                               "TyLeTuNop":  10.8,
                                               "TenBuuCuc":  "Bưu Cục 224 Lê Trọng Tấn-Q.Hà Đông-HN"
                                           },
                                           {
                                               "TongCOD":  195.4,
                                               "CODTuNop":  9.8,
                                               "TyLeTuNop":  5,
                                               "TenBuuCuc":  "Bưu Cục 303 Hoàng Hoa Thám-Q.Ba Đình-HN"
                                           },
                                           {
                                               "TongCOD":  0.45,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Điểm Lấy Hàng 60 Tam Hiệp 3-Phúc Thọ-HN"
                                           },
                                           {
                                               "TongCOD":  74.43,
                                               "CODTuNop":  74.37,
                                               "TyLeTuNop":  99.9,
                                               "TenBuuCuc":  "Bưu Cục Cồng Kềnh Ngõ 134 Sài Đồng-Q.Long Biên-HN"
                                           },
                                           {
                                               "TongCOD":  5,
                                               "CODTuNop":  5,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Cồng Kềnh 79 Thanh Xuân Làng Nghề Liên Hà-Đan Phượng-HN"
                                           },
                                           {
                                               "TongCOD":  102.27,
                                               "CODTuNop":  102.17,
                                               "TyLeTuNop":  99.9,
                                               "TenBuuCuc":  "Bưu Cục CL9 Đường Tràng An-Bát Tràng-HN"
                                           },
                                           {
                                               "TongCOD":  80.71,
                                               "CODTuNop":  20.1,
                                               "TyLeTuNop":  24.9,
                                               "TenBuuCuc":  "Bưu cục 602 Hà Huy Tập-Gia Lâm-HN"
                                           }
                                       ]
                    },
                    {
                        "TyLeTuNop":  3.1,
                        "TenVung":  "KHL",
                        "TongCOD":  1.1,
                        "Trend":  3.1,
                        "CODTuNop":  0.03,
                        "BuuCucList":  [
                                           {
                                               "TongCOD":  0.93,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Key Account Warehouse Ha Noi"
                                           },
                                           {
                                               "TongCOD":  0.17,
                                               "CODTuNop":  0.03,
                                               "TyLeTuNop":  19.9,
                                               "TenBuuCuc":  "Key Account Warehouse Ho Chi Minh"
                                           }
                                       ]
                    },
                    {
                        "TyLeTuNop":  25.8,
                        "TenVung":  "NTB",
                        "TongCOD":  12093.9,
                        "Trend":  7.3,
                        "CODTuNop":  3124.21,
                        "BuuCucList":  [
                                           {
                                               "TongCOD":  125.21,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 39 Hùng Vương-Diên Khánh-Khánh Hòa"
                                           },
                                           {
                                               "TongCOD":  179.66,
                                               "CODTuNop":  10.08,
                                               "TyLeTuNop":  5.6,
                                               "TenBuuCuc":  "Bưu Cục Thôn Phúc Hưng-Xã Tân Hà Lâm Hà-Lâm Đồng"
                                           },
                                           {
                                               "TongCOD":  80.99,
                                               "CODTuNop":  15.19,
                                               "TyLeTuNop":  18.8,
                                               "TenBuuCuc":  "Bưu Cục Hùng Vương-Nghĩa Trung-Gia Nghĩa-Đắk Nông"
                                           },
                                           {
                                               "TongCOD":  159.95,
                                               "CODTuNop":  91.42,
                                               "TyLeTuNop":  57.2,
                                               "TenBuuCuc":  "Bưu Cục 412 Nguyễn Văn Cừ-Lộc Phát-Bảo Lộc-Lâm Đồng"
                                           },
                                           {
                                               "TongCOD":  77.8,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Thôn Dầu Sơn-Cam Lâm-Khánh Hoà"
                                           },
                                           {
                                               "TongCOD":  149.87,
                                               "CODTuNop":  35.59,
                                               "TyLeTuNop":  23.7,
                                               "TenBuuCuc":  "Bưu Cục Số 7 Lê Thánh Tôn-LaGi-Bình Thuận"
                                           },
                                           {
                                               "TongCOD":  322.19,
                                               "CODTuNop":  13.94,
                                               "TyLeTuNop":  4.3,
                                               "TenBuuCuc":  "Bưu Cục 95A Đội Cấn-Bảo Lộc-Lâm Đồng"
                                           },
                                           {
                                               "TongCOD":  300.37,
                                               "CODTuNop":  122.46,
                                               "TyLeTuNop":  40.8,
                                               "TenBuuCuc":  "Bưu Cục Xã Nam Chính-Đức Linh-Bình Thuận"
                                           },
                                           {
                                               "TongCOD":  261.03,
                                               "CODTuNop":  112.65,
                                               "TyLeTuNop":  43.2,
                                               "TenBuuCuc":  "Bưu Cục 430 Tôn Đức Thắng-Tân Phước-Thị xã La Gi-Bình Thuận"
                                           },
                                           {
                                               "TongCOD":  135.48,
                                               "CODTuNop":  23.85,
                                               "TyLeTuNop":  17.6,
                                               "TenBuuCuc":  "Bưu Cục 237 Nguyễn Tất Thành-Bắc Bình-Bình Thuận"
                                           },
                                           {
                                               "TongCOD":  197.02,
                                               "CODTuNop":  30.24,
                                               "TyLeTuNop":  15.4,
                                               "TenBuuCuc":  "Bưu Cục 158 Quốc lộ 55-Tánh Linh-Bình Thuận"
                                           },
                                           {
                                               "TongCOD":  255.29,
                                               "CODTuNop":  59.15,
                                               "TyLeTuNop":  23.2,
                                               "TenBuuCuc":  "Bưu Cục Phú Tài-Phan Thiết-Bình Thuận"
                                           },
                                           {
                                               "TongCOD":  171.14,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu cục 1479 Trần Phú-Lộc Châu-Bảo Lộc-Lâm Đồng"
                                           },
                                           {
                                               "TongCOD":  108.88,
                                               "CODTuNop":  14.47,
                                               "TyLeTuNop":  13.3,
                                               "TenBuuCuc":  "Bưu Cục 44B Nguyễn Đình Chiểu-Đạ Teh-Lâm Đồng"
                                           },
                                           {
                                               "TongCOD":  180.42,
                                               "CODTuNop":  30.76,
                                               "TyLeTuNop":  17,
                                               "TenBuuCuc":  "Bưu Cục Đường 35 Hà Quang 1-Xã Nam Nha Trang-Khánh Hòa"
                                           },
                                           {
                                               "TongCOD":  72.81,
                                               "CODTuNop":  29.1,
                                               "TyLeTuNop":  40,
                                               "TenBuuCuc":  "Bưu Cục Phước Đồng-Nha Trang-Khánh Hoà"
                                           },
                                           {
                                               "TongCOD":  118.3,
                                               "CODTuNop":  55.34,
                                               "TyLeTuNop":  46.8,
                                               "TenBuuCuc":  "Bưu Cục Quốc Lộ 14-Đức An-Đắk Song-Đăk Nông"
                                           },
                                           {
                                               "TongCOD":  195.79,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Đạ Rsal-Đam Rông-Lâm Đồng"
                                           },
                                           {
                                               "TongCOD":  94.42,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 337 Hùng Vương-Lộc Thắng-Bảo Lâm-Lâm Đồng"
                                           },
                                           {
                                               "TongCOD":  98.43,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Thôn Long Hoà-Vạn Long-Vạn Ninh-Khánh Hoà"
                                           },
                                           {
                                               "TongCOD":  122.61,
                                               "CODTuNop":  122.61,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 21 Trần Hưng Đạo-Xã Phan Rí Cửa-Bình Thuận"
                                           },
                                           {
                                               "TongCOD":  167.76,
                                               "CODTuNop":  16.45,
                                               "TyLeTuNop":  9.8,
                                               "TenBuuCuc":  "Bưu Cục 575 Quốc Lộ 28-Xã Hàm Thuận-Bình Thuận"
                                           },
                                           {
                                               "TongCOD":  115.89,
                                               "CODTuNop":  5.94,
                                               "TyLeTuNop":  5.1,
                                               "TenBuuCuc":  "Bưu Cục 16 Huỳnh Thúc Kháng-Phan Thiết-Bình Thuận"
                                           },
                                           {
                                               "TongCOD":  174.16,
                                               "CODTuNop":  74.78,
                                               "TyLeTuNop":  42.9,
                                               "TenBuuCuc":  "Bưu Cục 40 Nguyễn Văn Linh-Cư Jút-Đắk Nông"
                                           },
                                           {
                                               "TongCOD":  167.72,
                                               "CODTuNop":  17.63,
                                               "TyLeTuNop":  10.5,
                                               "TenBuuCuc":  "Bưu Cục 1322 Hùng Vương-Di Linh-Lâm Đồng"
                                           },
                                           {
                                               "TongCOD":  56.8,
                                               "CODTuNop":  5.84,
                                               "TyLeTuNop":  10.3,
                                               "TenBuuCuc":  "Bưu Cục Quảng Khê-Đắk Glong-Đắk Nông"
                                           },
                                           {
                                               "TongCOD":  209.5,
                                               "CODTuNop":  209.5,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 254A Đường 21 tháng 8-Phan Rang-Tháp Chàm-Ninh Thuận"
                                           },
                                           {
                                               "TongCOD":  109.9,
                                               "CODTuNop":  32.4,
                                               "TyLeTuNop":  29.5,
                                               "TenBuuCuc":  "Bưu Cục 42 Nguyễn Du-Cam Lâm-Khánh Hoà"
                                           },
                                           {
                                               "TongCOD":  192.45,
                                               "CODTuNop":  192.45,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 51 Phạm Ngọc Thạch-Ninh Hải-Ninh Thuận"
                                           },
                                           {
                                               "TongCOD":  124.26,
                                               "CODTuNop":  31.78,
                                               "TyLeTuNop":  25.6,
                                               "TenBuuCuc":  "Bưu Cục Thôn 2-Tuy Đức-Đắk Nông"
                                           },
                                           {
                                               "TongCOD":  66.99,
                                               "CODTuNop":  22.34,
                                               "TyLeTuNop":  33.3,
                                               "TenBuuCuc":  "Bưu Cục Nâm N\u0027Jang-Đắk Song-Đắk Nông"
                                           },
                                           {
                                               "TongCOD":  166.91,
                                               "CODTuNop":  166.91,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 26A Lê Duẩn-Ninh Sơn-Ninh Thuận"
                                           },
                                           {
                                               "TongCOD":  234.7,
                                               "CODTuNop":  129.48,
                                               "TyLeTuNop":  55.2,
                                               "TenBuuCuc":  "Bưu Cục 06 Lê Hồng Phong-TP.Nha Trang-Khánh Hòa"
                                           },
                                           {
                                               "TongCOD":  203.36,
                                               "CODTuNop":  67.92,
                                               "TyLeTuNop":  33.4,
                                               "TenBuuCuc":  "Bưu Cục Quốc Lộ 1A-Ninh Phước-Ninh Thuận"
                                           },
                                           {
                                               "TongCOD":  60.1,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Thôn 2-Xã Quảng Tín-Đắk Nông"
                                           },
                                           {
                                               "TongCOD":  105.4,
                                               "CODTuNop":  20.81,
                                               "TyLeTuNop":  19.7,
                                               "TenBuuCuc":  "Bưu Cục Đạ Huoai-Lâm Đồng"
                                           },
                                           {
                                               "TongCOD":  181.74,
                                               "CODTuNop":  103.15,
                                               "TyLeTuNop":  56.8,
                                               "TenBuuCuc":  "Bưu cục Ngô Thỳ Sỹ - Đà Lạt - Lâm Đồng"
                                           },
                                           {
                                               "TongCOD":  91.41,
                                               "CODTuNop":  27.57,
                                               "TyLeTuNop":  30.2,
                                               "TenBuuCuc":  "Bưu cục 47 Trần Hưng Đạo-Hàm Thuận Nam-Bình Thuận"
                                           },
                                           {
                                               "TongCOD":  180.18,
                                               "CODTuNop":  2.22,
                                               "TyLeTuNop":  1.2,
                                               "TenBuuCuc":  "Bưu Cục 73 Nguyễn Tất Thành-Krông Nô-Đắk Nông"
                                           },
                                           {
                                               "TongCOD":  271.33,
                                               "CODTuNop":  17.5,
                                               "TyLeTuNop":  6.4,
                                               "TenBuuCuc":  "Bưu Cục Phú Thủy-Phan Thiết-Bình Thuận"
                                           },
                                           {
                                               "TongCOD":  105.19,
                                               "CODTuNop":  23.87,
                                               "TyLeTuNop":  22.7,
                                               "TenBuuCuc":  "Bưu Cục Đường Trần Lê-Cát Tiên-Lâm Đồng"
                                           },
                                           {
                                               "TongCOD":  166.28,
                                               "CODTuNop":  166.18,
                                               "TyLeTuNop":  99.9,
                                               "TenBuuCuc":  "Bưu cục Thị Trấn Liên Hương-Tuy Phong-Bình Thuận"
                                           },
                                           {
                                               "TongCOD":  101,
                                               "CODTuNop":  91.29,
                                               "TyLeTuNop":  90.4,
                                               "TenBuuCuc":  "Bưu Cục Lê Hồng Phong-Liên Hiệp-Đức Trọng-Lâm Đồng"
                                           },
                                           {
                                               "TongCOD":  197.4,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 468 Quốc Lộ 20-Liên Nghĩa-Đức Trọng-Lâm Đồng"
                                           },
                                           {
                                               "TongCOD":  92.79,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục TDP Ba Đình-Lâm Hà-Lâm Đồng"
                                           },
                                           {
                                               "TongCOD":  76.83,
                                               "CODTuNop":  19.17,
                                               "TyLeTuNop":  25,
                                               "TenBuuCuc":  "Bưu Cục Thôn Xuân Hoà 1-Ninh Hoà-Khánh Hoà"
                                           },
                                           {
                                               "TongCOD":  134.17,
                                               "CODTuNop":  134.17,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 454 Võ Văn Kiệt-Đặc Khu Đảo Phú Quý-Bình Thuận"
                                           },
                                           {
                                               "TongCOD":  217.45,
                                               "CODTuNop":  28.21,
                                               "TyLeTuNop":  13,
                                               "TenBuuCuc":  "Bưu Cục 40A Yết Kiêu-Nha Trang-Khánh Hòa"
                                           },
                                           {
                                               "TongCOD":  181.05,
                                               "CODTuNop":  7.35,
                                               "TyLeTuNop":  4.1,
                                               "TenBuuCuc":  "Bưu Cục 466 Đường 23/10-Nha Trang-Khánh Hòa"
                                           },
                                           {
                                               "TongCOD":  238.79,
                                               "CODTuNop":  28.23,
                                               "TyLeTuNop":  11.8,
                                               "TenBuuCuc":  "Bưu cục DT707-Hàm Thuận Nam-Bình Thuận"
                                           },
                                           {
                                               "TongCOD":  125.73,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 56 Phan Đình Phùng-Cam Linh-Khánh Hòa"
                                           },
                                           {
                                               "TongCOD":  95.76,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Quảng Sơn-Đắk Glong-Đắk Nông"
                                           },
                                           {
                                               "TongCOD":  107.62,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 01 Nguyễn Tất Thành-Đắk R lấp-Đắk Nông"
                                           },
                                           {
                                               "TongCOD":  71.77,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 231 Thôn 1-Xã Hòa Ninh-Lâm Đồng"
                                           },
                                           {
                                               "TongCOD":  121.8,
                                               "CODTuNop":  121.68,
                                               "TyLeTuNop":  99.9,
                                               "TenBuuCuc":  "Bưu Cục 190/11 Phù Đổng Thiên Vương-Phường 8-Đà Lạt-Lâm Đồng"
                                           },
                                           {
                                               "TongCOD":  120.8,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 300 Tự Phước-Đà Lạt-Lâm Đồng"
                                           },
                                           {
                                               "TongCOD":  253.53,
                                               "CODTuNop":  12.33,
                                               "TyLeTuNop":  4.9,
                                               "TenBuuCuc":  "Bưu Cục Đắk Mil-Đắk Nông"
                                           },
                                           {
                                               "TongCOD":  309.68,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Cao Bá Quát-Phường 7-Đà Lạt-Lâm Đồng"
                                           },
                                           {
                                               "TongCOD":  179.61,
                                               "CODTuNop":  11.42,
                                               "TyLeTuNop":  6.4,
                                               "TenBuuCuc":  "Bưu Cục 46 Hùng Vương-Hàm Tân-Bình Thuận"
                                           },
                                           {
                                               "TongCOD":  145.36,
                                               "CODTuNop":  40.24,
                                               "TyLeTuNop":  27.7,
                                               "TenBuuCuc":  "Bưu Cục Hàm Liêm-Hàm Thuận Bắc-Bình Thuận"
                                           },
                                           {
                                               "TongCOD":  130.43,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 36 Bà Triệu-Vạn Ninh-Khánh Hoà"
                                           },
                                           {
                                               "TongCOD":  114.7,
                                               "CODTuNop":  93.04,
                                               "TyLeTuNop":  81.1,
                                               "TenBuuCuc":  "Bưu Cục Langbiang-Lạc Dương-Lâm Đồng"
                                           },
                                           {
                                               "TongCOD":  111.61,
                                               "CODTuNop":  42.66,
                                               "TyLeTuNop":  38.2,
                                               "TenBuuCuc":  "Bưu Cục 571 Hùng Vương-Cam Nghĩa-TP.Cam Ranh-Khánh Hoà"
                                           },
                                           {
                                               "TongCOD":  82.97,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 53 Tôn Đức Thắng-Nghĩa Thành-Gia Nghĩa-Đắk Nông"
                                           },
                                           {
                                               "TongCOD":  223.78,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 229 Phước Long-Nam Nha Trang-Khánh Hòa"
                                           },
                                           {
                                               "TongCOD":  192.18,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục QL20 Ninh Hòa-Đức Trọng-Lâm Đồng"
                                           },
                                           {
                                               "TongCOD":  329.36,
                                               "CODTuNop":  171.56,
                                               "TyLeTuNop":  52.1,
                                               "TenBuuCuc":  "Bưu Cục 195 Đường 2/4-Nha Trang-Khánh Hòa"
                                           },
                                           {
                                               "TongCOD":  138.24,
                                               "CODTuNop":  4.43,
                                               "TyLeTuNop":  3.2,
                                               "TenBuuCuc":  "Bưu Cục 70 Lê Hồng Phong-Ninh Diêm-Ninh Hoà-Khánh Hoà"
                                           },
                                           {
                                               "TongCOD":  170.26,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu cục Đinh Văn-Lâm Hà-Lâm Đồng"
                                           },
                                           {
                                               "TongCOD":  70.82,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 111 Lê Duẫn-Khánh Sơn-Khánh Hoà"
                                           },
                                           {
                                               "TongCOD":  242.27,
                                               "CODTuNop":  50.82,
                                               "TyLeTuNop":  21,
                                               "TenBuuCuc":  "Bưu Cục Tổ Dân Phố 2-Ninh Hoà-Khánh Hoà"
                                           },
                                           {
                                               "TongCOD":  56.42,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Đường Tỉnh Lộ 10-Xã Diên Thọ-Khánh Hòa"
                                           },
                                           {
                                               "TongCOD":  227.74,
                                               "CODTuNop":  103.88,
                                               "TyLeTuNop":  45.6,
                                               "TenBuuCuc":  "Bưu Cục 95 Thống Nhất-Phan Rang-Ninh Thuận"
                                           },
                                           {
                                               "TongCOD":  85.45,
                                               "CODTuNop":  49.09,
                                               "TyLeTuNop":  57.4,
                                               "TenBuuCuc":  "Bưu Cục Thôn Phú An Nam 2-Xã Diên Khánh-Khánh Hòa"
                                           },
                                           {
                                               "TongCOD":  65.92,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Tổ 5 Đường Ngô Gia Tự-Khánh Vĩnh-Khánh Hòa"
                                           },
                                           {
                                               "TongCOD":  95.29,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Thôn R\u0027Chai 2-Xã Đức Trọng-Lâm Đồng"
                                           },
                                           {
                                               "TongCOD":  69.04,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Điểm Xử Lý Hàng Thôn 2-Xã Nhân Cơ-Đắk Nông"
                                           },
                                           {
                                               "TongCOD":  177.16,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Số 6 Trương Văn Hoàn-Phường 9-Đà Lạt-Lâm Đồng"
                                           },
                                           {
                                               "TongCOD":  76.25,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 585 Nguyễn Tất Thành-Xã Lương Sơn-Bình Thuận"
                                           },
                                           {
                                               "TongCOD":  103.19,
                                               "CODTuNop":  13.07,
                                               "TyLeTuNop":  12.7,
                                               "TenBuuCuc":  "Bưu Cục Phước Minh-Thuận Nam-Khánh Hòa"
                                           }
                                       ]
                    },
                    {
                        "TyLeTuNop":  42,
                        "TenVung":  "TBB",
                        "TongCOD":  5416.06,
                        "Trend":  15.4,
                        "CODTuNop":  2276.9,
                        "BuuCucList":  [
                                           {
                                               "TongCOD":  34.22,
                                               "CODTuNop":  12.44,
                                               "TyLeTuNop":  36.4,
                                               "TenBuuCuc":  "Bưu Cục 112 Võ Nguyên Giáp-Tam Đường-Lai Châu"
                                           },
                                           {
                                               "TongCOD":  141.07,
                                               "CODTuNop":  2.93,
                                               "TyLeTuNop":  2.1,
                                               "TenBuuCuc":  "Bưu Cục 244 Nguyễn Tất Thành-Yên Thế-Lục Yên-Yên Bái"
                                           },
                                           {
                                               "TongCOD":  105.54,
                                               "CODTuNop":  35.1,
                                               "TyLeTuNop":  33.3,
                                               "TenBuuCuc":  "Bưu Cục 341 Nguyễn Tất Thành-Bảo Yên-Lào Cai"
                                           },
                                           {
                                               "TongCOD":  139.53,
                                               "CODTuNop":  28.31,
                                               "TyLeTuNop":  20.3,
                                               "TenBuuCuc":  "Bưu Cục 136 CMT8-Sông Mã-Sơn La"
                                           },
                                           {
                                               "TongCOD":  269.64,
                                               "CODTuNop":  111.31,
                                               "TyLeTuNop":  41.3,
                                               "TenBuuCuc":  "Bưu Cục Tân Thanh-Điện Biên Phủ-Điện Biên"
                                           },
                                           {
                                               "TongCOD":  101.89,
                                               "CODTuNop":  35.04,
                                               "TyLeTuNop":  34.4,
                                               "TenBuuCuc":  "Bưu Cục 148 Hoàng Quốc Việt-TP.Lào Cai-Lào Cai"
                                           },
                                           {
                                               "TongCOD":  24.89,
                                               "CODTuNop":  10.53,
                                               "TyLeTuNop":  42.3,
                                               "TenBuuCuc":  "Bưu Cục Vận Hành TDP Hoàng Thu Phố-Xã Si Ma Cai-Lào Cai"
                                           },
                                           {
                                               "TongCOD":  127.69,
                                               "CODTuNop":  101.76,
                                               "TyLeTuNop":  79.7,
                                               "TenBuuCuc":  "Bưu Cục 052 Ngọc Uyển-Bắc Hà-Lào Cai"
                                           },
                                           {
                                               "TongCOD":  102.19,
                                               "CODTuNop":  36.52,
                                               "TyLeTuNop":  35.7,
                                               "TenBuuCuc":  "Bưu Cục 83 Điện Biên Phủ-Xã Bum Tở-Lai Châu"
                                           },
                                           {
                                               "TongCOD":  164.62,
                                               "CODTuNop":  89.8,
                                               "TyLeTuNop":  54.5,
                                               "TenBuuCuc":  "Bưu Cục Thôn 24-xã Noong Hẹt-Điện Biên"
                                           },
                                           {
                                               "TongCOD":  181.52,
                                               "CODTuNop":  117.39,
                                               "TyLeTuNop":  64.7,
                                               "TenBuuCuc":  "Bưu Cục 650 Đinh Tiên Hoàng-Yên Thịnh-Yên Bái"
                                           },
                                           {
                                               "TongCOD":  55.7,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Số 028 Thanh Niên-Than Uyên-Lai Châu"
                                           },
                                           {
                                               "TongCOD":  77.03,
                                               "CODTuNop":  65.01,
                                               "TyLeTuNop":  84.4,
                                               "TenBuuCuc":  "Bưu Cục Thị Trấn Tủa Chùa-Tủa Chùa-Điện Biên"
                                           },
                                           {
                                               "TongCOD":  153.99,
                                               "CODTuNop":  42.17,
                                               "TyLeTuNop":  27.4,
                                               "TenBuuCuc":  "Bưu Cục Ngã Ba Khe-Văn Chấn-Yên Bái"
                                           },
                                           {
                                               "TongCOD":  145.27,
                                               "CODTuNop":  100.61,
                                               "TyLeTuNop":  69.3,
                                               "TenBuuCuc":  "Bưu Cục Phố Lê Thanh-TP Lào Cai-Lào Cai"
                                           },
                                           {
                                               "TongCOD":  41.24,
                                               "CODTuNop":  14.74,
                                               "TyLeTuNop":  35.7,
                                               "TenBuuCuc":  "Bưu Cục TT Mường Ảng-Xã Mường Ảng-Điện Biên"
                                           },
                                           {
                                               "TongCOD":  57.34,
                                               "CODTuNop":  14.95,
                                               "TyLeTuNop":  26.1,
                                               "TenBuuCuc":  "Bưu Cục Xã Mường Giàng-Quỳnh Nhai-Sơn La"
                                           },
                                           {
                                               "TongCOD":  26.32,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 61 Lê Lợi-Xã Sìn Hồ-Lai Châu"
                                           },
                                           {
                                               "TongCOD":  53.09,
                                               "CODTuNop":  6.84,
                                               "TyLeTuNop":  12.9,
                                               "TenBuuCuc":  "Bưu Cục QL 6-Xã Vân Hồ-Sơn La"
                                           },
                                           {
                                               "TongCOD":  83.2,
                                               "CODTuNop":  14.75,
                                               "TyLeTuNop":  17.7,
                                               "TenBuuCuc":  "Bưu Cục Khánh Yên-Văn Bàn-Lào Cai"
                                           },
                                           {
                                               "TongCOD":  100.53,
                                               "CODTuNop":  20.28,
                                               "TyLeTuNop":  20.2,
                                               "TenBuuCuc":  "Bưu Cục 686 Lê Duẩn-Chiềng Sinh-Sơn La"
                                           },
                                           {
                                               "TongCOD":  62.1,
                                               "CODTuNop":  0.48,
                                               "TyLeTuNop":  0.8,
                                               "TenBuuCuc":  "Bưu Cục 563 Tiểu Khu 3-Xã Bắc Yên-Sơn La"
                                           },
                                           {
                                               "TongCOD":  164.91,
                                               "CODTuNop":  77.93,
                                               "TyLeTuNop":  47.3,
                                               "TenBuuCuc":  "Bưu Cục Tiểu Khu 1-Thuận Châu-Sơn La"
                                           },
                                           {
                                               "TongCOD":  10.35,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Tổ 2 Xã Mường Nhé-Mường Nhé-Điện Biên"
                                           },
                                           {
                                               "TongCOD":  62.22,
                                               "CODTuNop":  3.31,
                                               "TyLeTuNop":  5.3,
                                               "TenBuuCuc":  "Bưu Cục Thôn Tây Sơn-Mường So-Phong Thổ-Lai Châu"
                                           },
                                           {
                                               "TongCOD":  189.58,
                                               "CODTuNop":  102.31,
                                               "TyLeTuNop":  54,
                                               "TenBuuCuc":  "Bưu Cục 02 Phó Đức Chính-Hồng Hà-Yên Bái"
                                           },
                                           {
                                               "TongCOD":  67.49,
                                               "CODTuNop":  11.08,
                                               "TyLeTuNop":  16.4,
                                               "TenBuuCuc":  "Bưu Cục Vận Hành Tổ 1 TT Cổ Phúc-Xã Trấn Yên-Yên Bái"
                                           },
                                           {
                                               "TongCOD":  61.98,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Km 24 Bảo Ái-Yên Bình-Yên Bái"
                                           },
                                           {
                                               "TongCOD":  29.59,
                                               "CODTuNop":  8.19,
                                               "TyLeTuNop":  27.7,
                                               "TenBuuCuc":  "Bưu Cục Vận Hành Bản Hua Mường-Xã Xốp Cộp-Sơn La"
                                           },
                                           {
                                               "TongCOD":  68.7,
                                               "CODTuNop":  27.05,
                                               "TyLeTuNop":  39.4,
                                               "TenBuuCuc":  "Điểm xử lý hàng Nậm Pồ-Điện Biên"
                                           },
                                           {
                                               "TongCOD":  331.53,
                                               "CODTuNop":  224.62,
                                               "TyLeTuNop":  67.8,
                                               "TenBuuCuc":  "Bưu Cục 77 Hoàng Quốc Việt-Mộc Châu-Sơn La"
                                           },
                                           {
                                               "TongCOD":  45.16,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Điểm xử lý hàng hoá cho Sìn Hồ Thấp-Lai Châu"
                                           },
                                           {
                                               "TongCOD":  38.54,
                                               "CODTuNop":  9.38,
                                               "TyLeTuNop":  24.3,
                                               "TenBuuCuc":  "Bưu Cục Tổ 4 Điện Biên Đông-Xã Na Son-Điện Biên"
                                           },
                                           {
                                               "TongCOD":  202.74,
                                               "CODTuNop":  52.22,
                                               "TyLeTuNop":  25.8,
                                               "TenBuuCuc":  "Bưu Cục Hồng Hà-Mậu A-Văn Yên-Yên Bái"
                                           },
                                           {
                                               "TongCOD":  36.46,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 11 Trường Trinh-Xã Nậm Hàng-Lai Châu"
                                           },
                                           {
                                               "TongCOD":  176.24,
                                               "CODTuNop":  87.9,
                                               "TyLeTuNop":  49.9,
                                               "TenBuuCuc":  "Bưu Cục Cầu Thia-Nghĩa Lộ-Yên Bái"
                                           },
                                           {
                                               "TongCOD":  74.73,
                                               "CODTuNop":  6.04,
                                               "TyLeTuNop":  8.1,
                                               "TenBuuCuc":  "Bưu Cục Thôn Mã Tuyển 3-Mường Khương-Lào Cai"
                                           },
                                           {
                                               "TongCOD":  64.87,
                                               "CODTuNop":  7.79,
                                               "TyLeTuNop":  12,
                                               "TenBuuCuc":  "Bưu Cục Tổ 7-Bát Xát-Lào Cai"
                                           },
                                           {
                                               "TongCOD":  174.13,
                                               "CODTuNop":  174.13,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 155 Điện Biên Phủ-Sa Pa-Lào Cai"
                                           },
                                           {
                                               "TongCOD":  221.38,
                                               "CODTuNop":  117.67,
                                               "TyLeTuNop":  53.2,
                                               "TenBuuCuc":  "Bưu Cục Tiểu Khu 10-Hát Lót-Mai Sơn-Sơn La"
                                           },
                                           {
                                               "TongCOD":  14.61,
                                               "CODTuNop":  5.29,
                                               "TyLeTuNop":  36.2,
                                               "TenBuuCuc":  "Bưu Cục Bản Liên Hà 3-Xã Bảo Hà-Lào Cai"
                                           },
                                           {
                                               "TongCOD":  91.64,
                                               "CODTuNop":  42.61,
                                               "TyLeTuNop":  46.5,
                                               "TenBuuCuc":  "Bưu Cục 36 Đường Tô Hiệu-Xã Mường La-Sơn La"
                                           },
                                           {
                                               "TongCOD":  129.01,
                                               "CODTuNop":  65.59,
                                               "TyLeTuNop":  50.8,
                                               "TenBuuCuc":  "Bưu Cục Tiểu khu 6-Phù Yên-Sơn La"
                                           },
                                           {
                                               "TongCOD":  84.34,
                                               "CODTuNop":  18.23,
                                               "TyLeTuNop":  21.6,
                                               "TenBuuCuc":  "Bưu Cục Thị trấn Tuần Giáo-Điện Biên"
                                           },
                                           {
                                               "TongCOD":  25.11,
                                               "CODTuNop":  9.54,
                                               "TyLeTuNop":  38,
                                               "TenBuuCuc":  "Bưu Cục TDP 1-Mường Chà-Điện Biên"
                                           },
                                           {
                                               "TongCOD":  117.55,
                                               "CODTuNop":  85.87,
                                               "TyLeTuNop":  73,
                                               "TenBuuCuc":  "Bưu Cục 266 Đường 8 Tháng 5-Xã Yên Châu-Sơn La"
                                           },
                                           {
                                               "TongCOD":  187.97,
                                               "CODTuNop":  98,
                                               "TyLeTuNop":  52.1,
                                               "TenBuuCuc":  "Bưu Cục Thôn Tả Hà 3-Bảo Thắng-Lào Cai"
                                           },
                                           {
                                               "TongCOD":  120.14,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 131 Lê Duẩn-Tân Phong-Lai Châu"
                                           },
                                           {
                                               "TongCOD":  47.43,
                                               "CODTuNop":  9.53,
                                               "TyLeTuNop":  20.1,
                                               "TenBuuCuc":  "Bưu Cục 1103 Lê Lợi-Xã Tân Uyên-Lai Châu"
                                           },
                                           {
                                               "TongCOD":  208.28,
                                               "CODTuNop":  136.89,
                                               "TyLeTuNop":  65.7,
                                               "TenBuuCuc":  "Bưu Cục 354 Trần Đăng Ninh-Sơn La-Sơn La"
                                           },
                                           {
                                               "TongCOD":  64.53,
                                               "CODTuNop":  34.8,
                                               "TyLeTuNop":  53.9,
                                               "TenBuuCuc":  "Bưu Cục Duyên Hải-Duyên Hải-Lào Cai"
                                           },
                                           {
                                               "TongCOD":  56.31,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 290/9B Đại Hồng-Xã Yên Bình-Yên Bái"
                                           }
                                       ]
                    },
                    {
                        "TyLeTuNop":  0,
                        "TenVung":  "Tech",
                        "TongCOD":  0,
                        "Trend":  0,
                        "CODTuNop":  0,
                        "BuuCucList":  {

                                       }
                    },
                    {
                        "TyLeTuNop":  38.3,
                        "TenVung":  "TNB",
                        "TongCOD":  15389.49,
                        "Trend":  3.6,
                        "CODTuNop":  5891.41,
                        "BuuCucList":  [
                                           {
                                               "TongCOD":  204.52,
                                               "CODTuNop":  4.77,
                                               "TyLeTuNop":  2.3,
                                               "TenBuuCuc":  "Bưu Cục 211 Quốc Lộ 80-Kiên Lương-Kiên Giang"
                                           },
                                           {
                                               "TongCOD":  120.83,
                                               "CODTuNop":  20.48,
                                               "TyLeTuNop":  17,
                                               "TenBuuCuc":  "Bưu Cục 56 Quốc Lộ 1A-Xã Vĩnh Lợi-Bạc Liêu"
                                           },
                                           {
                                               "TongCOD":  159.24,
                                               "CODTuNop":  132.71,
                                               "TyLeTuNop":  83.3,
                                               "TenBuuCuc":  "Bưu Cục Đường Huyện 21-Thới Lai-Cần Thơ"
                                           },
                                           {
                                               "TongCOD":  144.69,
                                               "CODTuNop":  48.24,
                                               "TyLeTuNop":  33.3,
                                               "TenBuuCuc":  "Bưu Cục 28 Xuân Hồng-Bình Thuỷ-Cần Thơ"
                                           },
                                           {
                                               "TongCOD":  107.48,
                                               "CODTuNop":  14.62,
                                               "TyLeTuNop":  13.6,
                                               "TenBuuCuc":  "Bưu Cục KDC Định An Quốc Lộ 61-Xã Gò Quao-Kiên Giang"
                                           },
                                           {
                                               "TongCOD":  245.03,
                                               "CODTuNop":  128.02,
                                               "TyLeTuNop":  52.2,
                                               "TenBuuCuc":  "Bưu Cục Lý Thường Kiệt-Tân Thành-Cà Mau"
                                           },
                                           {
                                               "TongCOD":  355.37,
                                               "CODTuNop":  22.3,
                                               "TyLeTuNop":  6.3,
                                               "TenBuuCuc":  "Bưu Cục Lô C02-03 Lê Hồng Phong-An Hòa-Rạch Giá-Kiên Giang"
                                           },
                                           {
                                               "TongCOD":  87.83,
                                               "CODTuNop":  87.83,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 535 Tôn Đức Thắng-Sóc Trăng-Sóc Trăng"
                                           },
                                           {
                                               "TongCOD":  138.72,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Đảo Kiên Hải-Kiên Giang"
                                           },
                                           {
                                               "TongCOD":  159.54,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Thoại Sơn-An Giang"
                                           },
                                           {
                                               "TongCOD":  295.06,
                                               "CODTuNop":  60.33,
                                               "TyLeTuNop":  20.4,
                                               "TenBuuCuc":  "Bưu Cục Tri Tôn-An Giang"
                                           },
                                           {
                                               "TongCOD":  82.49,
                                               "CODTuNop":  25.56,
                                               "TyLeTuNop":  31,
                                               "TenBuuCuc":  "Bưu Cục Tỉnh Lộ 963-Xã Hòa Thuận-Kiên Giang"
                                           },
                                           {
                                               "TongCOD":  280.07,
                                               "CODTuNop":  89.32,
                                               "TyLeTuNop":  31.9,
                                               "TenBuuCuc":  "Bưu Cục 154B Mai Thị Hồng Hạnh-Rạch Giá-Kiên Giang"
                                           },
                                           {
                                               "TongCOD":  133.39,
                                               "CODTuNop":  9.81,
                                               "TyLeTuNop":  7.4,
                                               "TenBuuCuc":  "Bưu Cục 137 Đường 30/04-Trần Văn Thời-Cà Mau"
                                           },
                                           {
                                               "TongCOD":  177.34,
                                               "CODTuNop":  75.9,
                                               "TyLeTuNop":  42.8,
                                               "TenBuuCuc":  "Bưu Cục 54 Hùng Vương-Năm Căn-Cà Mau"
                                           },
                                           {
                                               "TongCOD":  349.19,
                                               "CODTuNop":  157.06,
                                               "TyLeTuNop":  45,
                                               "TenBuuCuc":  "Bưu Cục 66 Nguyễn Tri Phương-Châu Phú B-TP.Châu Đốc-An Giang"
                                           },
                                           {
                                               "TongCOD":  132.46,
                                               "CODTuNop":  19.09,
                                               "TyLeTuNop":  14.4,
                                               "TenBuuCuc":  "Bưu Cục Vĩnh Phước QL63-Vĩnh Thuận-Kiên Giang"
                                           },
                                           {
                                               "TongCOD":  226.46,
                                               "CODTuNop":  44.17,
                                               "TyLeTuNop":  19.5,
                                               "TenBuuCuc":  "Bưu Cục Nguyễn An Ninh-Vị Thanh-Hậu Giang"
                                           },
                                           {
                                               "TongCOD":  196.49,
                                               "CODTuNop":  165.85,
                                               "TyLeTuNop":  84.4,
                                               "TenBuuCuc":  "Bưu Cục 33 Ấp Chợ-Xã Trần Đề-Sóc Trăng"
                                           },
                                           {
                                               "TongCOD":  179.34,
                                               "CODTuNop":  151.44,
                                               "TyLeTuNop":  84.4,
                                               "TenBuuCuc":  "Bưu Cục Hồng Khanh-Phước Long-Bạc Liêu"
                                           },
                                           {
                                               "TongCOD":  253.76,
                                               "CODTuNop":  110.44,
                                               "TyLeTuNop":  43.5,
                                               "TenBuuCuc":  "Bưu Cục 599 Lê Hồng Phong-Phú Lợi-Sóc Trăng"
                                           },
                                           {
                                               "TongCOD":  145.19,
                                               "CODTuNop":  56.01,
                                               "TyLeTuNop":  38.6,
                                               "TenBuuCuc":  "Bưu Cục 06 Nguyễn Tri Phương-Vị Thuỷ-Hậu Giang"
                                           },
                                           {
                                               "TongCOD":  107.27,
                                               "CODTuNop":  18.65,
                                               "TyLeTuNop":  17.4,
                                               "TenBuuCuc":  "Bưu Cục 86 Đường 932-Phú Tân-Châu Thành-Sóc Trăng"
                                           },
                                           {
                                               "TongCOD":  161.63,
                                               "CODTuNop":  42.34,
                                               "TyLeTuNop":  26.2,
                                               "TenBuuCuc":  "Bưu Cục Nguyễn Trung Thành Khóm 1-U Minh-Cà Mau"
                                           },
                                           {
                                               "TongCOD":  387.07,
                                               "CODTuNop":  103.03,
                                               "TyLeTuNop":  26.6,
                                               "TenBuuCuc":  "Bưu Cục 617 Tỉnh Lộ 942-Xã Chợ Mới-An Giang"
                                           },
                                           {
                                               "TongCOD":  184.77,
                                               "CODTuNop":  33.68,
                                               "TyLeTuNop":  18.2,
                                               "TenBuuCuc":  "Bưu Cục 266B Hàng Gòn-Lê Bình-Cái Răng-Cần Thơ"
                                           },
                                           {
                                               "TongCOD":  299.68,
                                               "CODTuNop":  152.89,
                                               "TyLeTuNop":  51,
                                               "TenBuuCuc":  "Bưu Cục 555 Nguyễn Văn Cừ-Phú Quốc-Kiên Giang"
                                           },
                                           {
                                               "TongCOD":  159.6,
                                               "CODTuNop":  49.06,
                                               "TyLeTuNop":  30.7,
                                               "TenBuuCuc":  "Bưu Cục 28 Giá Rai-Xã Long Điền-Bạc Liêu"
                                           },
                                           {
                                               "TongCOD":  122.87,
                                               "CODTuNop":  122.87,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 86 Thống nhất II-Ngan Dừa-Hồng Dân-Bạc Liêu"
                                           },
                                           {
                                               "TongCOD":  57.55,
                                               "CODTuNop":  22.83,
                                               "TyLeTuNop":  39.7,
                                               "TenBuuCuc":  "Bưu Cục Ấp 3 -Thạnh Trị-Sóc Trăng"
                                           },
                                           {
                                               "TongCOD":  240.98,
                                               "CODTuNop":  95.5,
                                               "TyLeTuNop":  39.6,
                                               "TenBuuCuc":  "Bưu Cục 608 Trần Hưng Đạo-Bình Đức-TP.Long Xuyên-An Giang"
                                           },
                                           {
                                               "TongCOD":  117.44,
                                               "CODTuNop":  42.8,
                                               "TyLeTuNop":  36.4,
                                               "TenBuuCuc":  "Bưu Cục 379 Hùng Vương-Châu Thành-Hậu Giang"
                                           },
                                           {
                                               "TongCOD":  94.49,
                                               "CODTuNop":  31.42,
                                               "TyLeTuNop":  33.3,
                                               "TenBuuCuc":  "Bưu Cục 88 Quốc Lộ 61C-Xã Tân Hòa-Hậu Giang"
                                           },
                                           {
                                               "TongCOD":  200.27,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 356 Đặng Văn Dầy-Bình Thủy-Cần Thơ"
                                           },
                                           {
                                               "TongCOD":  149.13,
                                               "CODTuNop":  140.21,
                                               "TyLeTuNop":  94,
                                               "TenBuuCuc":  "Bưu Cục Dương Thị Cẩm Vân-Khóm 4-Đầm Dơi-Cà Mau"
                                           },
                                           {
                                               "TongCOD":  212.29,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 4344 Bạch Đằng-An Phú-An Giang"
                                           },
                                           {
                                               "TongCOD":  115.32,
                                               "CODTuNop":  31.62,
                                               "TyLeTuNop":  27.4,
                                               "TenBuuCuc":  "Bưu Cục Ấp Thị Tứ-Thị Trấn Phong Điền-Cần Thơ"
                                           },
                                           {
                                               "TongCOD":  187.03,
                                               "CODTuNop":  11.45,
                                               "TyLeTuNop":  6.1,
                                               "TenBuuCuc":  "Bưu Cục Châu Văn Liêm-Ô Môn-Cần Thơ"
                                           },
                                           {
                                               "TongCOD":  141.64,
                                               "CODTuNop":  81.26,
                                               "TyLeTuNop":  57.4,
                                               "TenBuuCuc":  "Bưu Cục Xã Cửa Cạn-Phú Quốc-Kiên Giang"
                                           },
                                           {
                                               "TongCOD":  127.42,
                                               "CODTuNop":  62.13,
                                               "TyLeTuNop":  48.8,
                                               "TenBuuCuc":  "Bưu Cục 02 Lâm Văn Quang-Hà Tiên-Kiên Giang"
                                           },
                                           {
                                               "TongCOD":  141.48,
                                               "CODTuNop":  25.82,
                                               "TyLeTuNop":  18.3,
                                               "TenBuuCuc":  "Bưu Cục Lý Văn Lâm-TP.Cà Mau-Cà Mau"
                                           },
                                           {
                                               "TongCOD":  194.24,
                                               "CODTuNop":  30.6,
                                               "TyLeTuNop":  15.8,
                                               "TenBuuCuc":  "Bưu Cục 64 Đường 30/04 Vĩnh Châu-Sóc Trăng"
                                           },
                                           {
                                               "TongCOD":  329.68,
                                               "CODTuNop":  161.09,
                                               "TyLeTuNop":  48.9,
                                               "TenBuuCuc":  "Bưu cục Đường 23 Tháng 8-TP.Bạc Liêu-Bạc Liêu"
                                           },
                                           {
                                               "TongCOD":  217.85,
                                               "CODTuNop":  48.43,
                                               "TyLeTuNop":  22.2,
                                               "TenBuuCuc":  "Bưu Cục Thị trấn Sông Đốc-Trần Văn Thời-Cà Mau"
                                           },
                                           {
                                               "TongCOD":  102.01,
                                               "CODTuNop":  18.02,
                                               "TyLeTuNop":  17.7,
                                               "TenBuuCuc":  "Bưu Cục Thạnh Yên-U Minh Thượng-Kiên Giang"
                                           },
                                           {
                                               "TongCOD":  125.24,
                                               "CODTuNop":  104.18,
                                               "TyLeTuNop":  83.2,
                                               "TenBuuCuc":  "Bưu Cục 39 Tỉnh Lộ 940-Huỳnh Hữu Nghĩa-Mỹ Tú-Sóc Trăng"
                                           },
                                           {
                                               "TongCOD":  232.89,
                                               "CODTuNop":  36.86,
                                               "TyLeTuNop":  15.8,
                                               "TenBuuCuc":  "Bưu Cục Cái Tắc-Phú Mỹ-Phú Tân-An Giang"
                                           },
                                           {
                                               "TongCOD":  235.56,
                                               "CODTuNop":  31.82,
                                               "TyLeTuNop":  13.5,
                                               "TenBuuCuc":  "Bưu Cục Quốc Lộ 80-Xã Hòn Đất-Kiên Giang"
                                           },
                                           {
                                               "TongCOD":  207.43,
                                               "CODTuNop":  110.69,
                                               "TyLeTuNop":  53.4,
                                               "TenBuuCuc":  "Bưu Cục 729 Đông An-Xã Tân Hiệp-Kiên Giang"
                                           },
                                           {
                                               "TongCOD":  7.05,
                                               "CODTuNop":  7.05,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 578 QL1A-Xã Hòa Bình-Bạc Liêu"
                                           },
                                           {
                                               "TongCOD":  91.14,
                                               "CODTuNop":  4.04,
                                               "TyLeTuNop":  4.4,
                                               "TenBuuCuc":  "Bưu Cục 79 Quốc Lộ 1A-Giá Rai-Bạc Liêu"
                                           },
                                           {
                                               "TongCOD":  122.2,
                                               "CODTuNop":  24.47,
                                               "TyLeTuNop":  20,
                                               "TenBuuCuc":  "Bưu Cục Nguyễn Văn Cưng-Rạch Gốc-Ngọc Hiển-Cà Mau"
                                           },
                                           {
                                               "TongCOD":  89.5,
                                               "CODTuNop":  68.8,
                                               "TyLeTuNop":  76.9,
                                               "TenBuuCuc":  "Bưu Cục 205 Nguyễn Trãi-Thị Xã Ngã Năm-Sóc Trăng"
                                           },
                                           {
                                               "TongCOD":  220.89,
                                               "CODTuNop":  57.15,
                                               "TyLeTuNop":  25.9,
                                               "TenBuuCuc":  "Bưu Cục 320/12 Ngô Quyền-Cà Mau-Cà Mau"
                                           },
                                           {
                                               "TongCOD":  318.1,
                                               "CODTuNop":  317.93,
                                               "TyLeTuNop":  99.9,
                                               "TenBuuCuc":  "Bưu Cục 68 Nguyễn Văn Cừ-An Khánh-Ninh Kiều-Cần Thơ"
                                           },
                                           {
                                               "TongCOD":  148.79,
                                               "CODTuNop":  135.31,
                                               "TyLeTuNop":  90.9,
                                               "TenBuuCuc":  "Bưu Cục 162 Trần Văn Trà-Cái Răng-Cần Thơ"
                                           },
                                           {
                                               "TongCOD":  104.32,
                                               "CODTuNop":  10.4,
                                               "TyLeTuNop":  10,
                                               "TenBuuCuc":  "Bưu Cục 8 Huỳnh Mẫn Đạt-Giồng Riềng-Kiên Giang"
                                           },
                                           {
                                               "TongCOD":  110.9,
                                               "CODTuNop":  33.6,
                                               "TyLeTuNop":  30.3,
                                               "TenBuuCuc":  "Bưu Cục Thị Trấn Cây Dương-Phụng Hiệp-Hậu Giang"
                                           },
                                           {
                                               "TongCOD":  122.91,
                                               "CODTuNop":  61.92,
                                               "TyLeTuNop":  50.4,
                                               "TenBuuCuc":  "Bưu Cục Ấp 1 QL61-Châu Thành A-Hậu Giang"
                                           },
                                           {
                                               "TongCOD":  15.98,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Điểm xử lý hàng Thổ Châu-Phú Quốc-Kiên Giang"
                                           },
                                           {
                                               "TongCOD":  221.37,
                                               "CODTuNop":  91.44,
                                               "TyLeTuNop":  41.3,
                                               "TenBuuCuc":  "Bưu Cục Phường Trung Kiên-Thốt Nốt-Cần Thơ"
                                           },
                                           {
                                               "TongCOD":  209.45,
                                               "CODTuNop":  30.46,
                                               "TyLeTuNop":  14.5,
                                               "TenBuuCuc":  "Bưu Cục Tân Châu-An Giang"
                                           },
                                           {
                                               "TongCOD":  196.21,
                                               "CODTuNop":  18.38,
                                               "TyLeTuNop":  9.4,
                                               "TenBuuCuc":  "Bưu Cục Thạnh An-An Minh-Kiên Giang"
                                           },
                                           {
                                               "TongCOD":  246.12,
                                               "CODTuNop":  181.69,
                                               "TyLeTuNop":  73.8,
                                               "TenBuuCuc":  "Bưu Cục 32B Huyện lộ 6-TT.Kế Sách-Sóc Trăng"
                                           },
                                           {
                                               "TongCOD":  144.93,
                                               "CODTuNop":  144.93,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Lô B2 TTTM Vĩnh Thạnh-Vĩnh Thạnh-Cần Thơ"
                                           },
                                           {
                                               "TongCOD":  142.02,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 1224 Ấp Giống Có-Xã Ngọc Tố-Sóc Trăng"
                                           },
                                           {
                                               "TongCOD":  172.28,
                                               "CODTuNop":  12.82,
                                               "TyLeTuNop":  7.4,
                                               "TenBuuCuc":  "Bưu Cục 03 Lê Lợi-An Châu-An Giang"
                                           },
                                           {
                                               "TongCOD":  76.37,
                                               "CODTuNop":  15.81,
                                               "TyLeTuNop":  20.7,
                                               "TenBuuCuc":  "Bưu Cục DT933 Ấp Đồng Tràm-Xã Long Thạnh-Kiên Giang"
                                           },
                                           {
                                               "TongCOD":  323.1,
                                               "CODTuNop":  89.23,
                                               "TyLeTuNop":  27.6,
                                               "TenBuuCuc":  "Bưu Cục QL91-Châu Phú-An Giang"
                                           },
                                           {
                                               "TongCOD":  107.88,
                                               "CODTuNop":  4.73,
                                               "TyLeTuNop":  4.4,
                                               "TenBuuCuc":  "Bưu Cục Tân Hưng Đông-Cái Nước-Cà Mau"
                                           },
                                           {
                                               "TongCOD":  130.72,
                                               "CODTuNop":  10.82,
                                               "TyLeTuNop":  8.3,
                                               "TenBuuCuc":  "Bưu Cục Ấp Chợ-Cù Lao Dung-Sóc Trăng"
                                           },
                                           {
                                               "TongCOD":  269.12,
                                               "CODTuNop":  269.12,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 16 Đỗ Nhuận-Long Xuyên-An Giang"
                                           },
                                           {
                                               "TongCOD":  110.98,
                                               "CODTuNop":  2.26,
                                               "TyLeTuNop":  2,
                                               "TenBuuCuc":  "Bưu Cục Ấp Tân Long A-Xã Tân Tiến-Cà Mau"
                                           },
                                           {
                                               "TongCOD":  162.4,
                                               "CODTuNop":  162.4,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Ấp 4-Thới Bình-Thới Bình-Cà Mau"
                                           },
                                           {
                                               "TongCOD":  95.86,
                                               "CODTuNop":  27.05,
                                               "TyLeTuNop":  28.2,
                                               "TenBuuCuc":  "Bưu Cục 68 Quốc Lộ 1A-Xã Lương Thế Trân-Cà Mau"
                                           },
                                           {
                                               "TongCOD":  334.86,
                                               "CODTuNop":  158.87,
                                               "TyLeTuNop":  47.4,
                                               "TenBuuCuc":  "Bưu Cục Ngã 3 Bãy Vòng 973-Đặc Khu Phú Quốc-Kiên Giang"
                                           },
                                           {
                                               "TongCOD":  196.26,
                                               "CODTuNop":  43.57,
                                               "TyLeTuNop":  22.2,
                                               "TenBuuCuc":  "Bưu Cục 931 Đường 30/4-Long Mỹ-Hậu Giang"
                                           },
                                           {
                                               "TongCOD":  180.74,
                                               "CODTuNop":  35.93,
                                               "TyLeTuNop":  19.9,
                                               "TenBuuCuc":  "Bưu Cục Ấp Khúc Tréo B-Tân Phong-Giá Rai-Bạc Liêu"
                                           },
                                           {
                                               "TongCOD":  159.79,
                                               "CODTuNop":  1.93,
                                               "TyLeTuNop":  1.2,
                                               "TenBuuCuc":  "Bưu Cục Đường Phan Ngọc Hiển-Xã Cái Đôi Vàm-Cà Mau"
                                           },
                                           {
                                               "TongCOD":  323.46,
                                               "CODTuNop":  187.24,
                                               "TyLeTuNop":  57.9,
                                               "TenBuuCuc":  "Bưu Cục 56 Nguyễn Văn Cừ-Cái Khế-Ninh Kiều-Cần Thơ"
                                           },
                                           {
                                               "TongCOD":  152.24,
                                               "CODTuNop":  152.24,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Tỉnh lộ 919-Thạnh Phú-Cờ Đỏ-Cần Thơ"
                                           },
                                           {
                                               "TongCOD":  519.72,
                                               "CODTuNop":  388.87,
                                               "TyLeTuNop":  74.8,
                                               "TenBuuCuc":  "Bưu Cục Đường Tuyến Tránh-Phú Quốc-Kiên Giang"
                                           },
                                           {
                                               "TongCOD":  247.35,
                                               "CODTuNop":  21.22,
                                               "TyLeTuNop":  8.6,
                                               "TenBuuCuc":  "Bưu Cục 20 Lộ Liên Hương-Rạch Giá-Kiên Giang"
                                           },
                                           {
                                               "TongCOD":  121.01,
                                               "CODTuNop":  30.95,
                                               "TyLeTuNop":  25.6,
                                               "TenBuuCuc":  "Bưu Cục 2C Châu Văn Liêm-Ngã Bảy-Hậu Giang"
                                           },
                                           {
                                               "TongCOD":  166.15,
                                               "CODTuNop":  55.55,
                                               "TyLeTuNop":  33.4,
                                               "TenBuuCuc":  "Bưu Cục QL63 TT Thứ 3-Xã An Biên-Kiên Giang"
                                           }
                                       ]
                    },
                    {
                        "TyLeTuNop":  32.4,
                        "TenVung":  "TNG",
                        "TongCOD":  9714.78,
                        "Trend":  -1.6,
                        "CODTuNop":  3151.74,
                        "BuuCucList":  [
                                           {
                                               "TongCOD":  144.53,
                                               "CODTuNop":  52.26,
                                               "TyLeTuNop":  36.2,
                                               "TenBuuCuc":  "Bưu Cục Đường 3 Tháng 2-Xã Phù Cát-Bình Định"
                                           },
                                           {
                                               "TongCOD":  157.22,
                                               "CODTuNop":  74.27,
                                               "TyLeTuNop":  47.2,
                                               "TenBuuCuc":  "Bưu Cục 34 Quang Trung-Eakar-Đắk Lắk"
                                           },
                                           {
                                               "TongCOD":  145.73,
                                               "CODTuNop":  24.87,
                                               "TyLeTuNop":  17.1,
                                               "TenBuuCuc":  "Bưu Cục 172 Quang Trung-Đức Cơ-Gia Lai"
                                           },
                                           {
                                               "TongCOD":  37.54,
                                               "CODTuNop":  9.09,
                                               "TyLeTuNop":  24.2,
                                               "TenBuuCuc":  "Bưu Cục 61 Quang Trung-Xã Vân Canh-Bình Định"
                                           },
                                           {
                                               "TongCOD":  294.57,
                                               "CODTuNop":  111.42,
                                               "TyLeTuNop":  37.8,
                                               "TenBuuCuc":  "Bưu cục 499 Phạm Văn Đồng-Pleiku-Gia Lai"
                                           },
                                           {
                                               "TongCOD":  215.01,
                                               "CODTuNop":  40.58,
                                               "TyLeTuNop":  18.9,
                                               "TenBuuCuc":  "Bưu Cục 56 Đường 17 Tháng 3-Chư Sê-Gia Lai"
                                           },
                                           {
                                               "TongCOD":  332.25,
                                               "CODTuNop":  245.8,
                                               "TyLeTuNop":  74,
                                               "TenBuuCuc":  "Bưu Cục 15 Hoàng Văn Thụ-An Khê-Gia Lai"
                                           },
                                           {
                                               "TongCOD":  84.18,
                                               "CODTuNop":  28.46,
                                               "TyLeTuNop":  33.8,
                                               "TenBuuCuc":  "Bưu Cục 274 Hùng Vương-Chư Pưh-Gia Lai"
                                           },
                                           {
                                               "TongCOD":  85.99,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 01 An Bình-Ea KNốp-Ea Kar-Đắk Lắk"
                                           },
                                           {
                                               "TongCOD":  84.89,
                                               "CODTuNop":  36.44,
                                               "TyLeTuNop":  42.9,
                                               "TenBuuCuc":  "Bưu Cục Ngã Ba Đắk Liêng-Huyện Lắk-Đắk Lắk"
                                           },
                                           {
                                               "TongCOD":  246.49,
                                               "CODTuNop":  110.68,
                                               "TyLeTuNop":  44.9,
                                               "TenBuuCuc":  "Bưu Cục 269 Tây Sơn-Quy Nhơn-Bình Định"
                                           },
                                           {
                                               "TongCOD":  250.77,
                                               "CODTuNop":  98.38,
                                               "TyLeTuNop":  39.2,
                                               "TenBuuCuc":  "Bưu Cục 172 Hùng Vương-Chư Prông-Gia Lai"
                                           },
                                           {
                                               "TongCOD":  42.43,
                                               "CODTuNop":  42.43,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 09 Quốc Lộ 14-Hoà Phú-Đắk Lắk"
                                           },
                                           {
                                               "TongCOD":  105.22,
                                               "CODTuNop":  105.22,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Hoài Nhơn-Bình Định"
                                           },
                                           {
                                               "TongCOD":  173.31,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 21 Tây Sơn-Xã Phù Mỹ-Bình Định"
                                           },
                                           {
                                               "TongCOD":  260.24,
                                               "CODTuNop":  162.23,
                                               "TyLeTuNop":  62.3,
                                               "TenBuuCuc":  "Bưu Cục 263 Đống Đa-Quy Nhơn-Bình Định"
                                           },
                                           {
                                               "TongCOD":  143.28,
                                               "CODTuNop":  10.07,
                                               "TyLeTuNop":  7,
                                               "TenBuuCuc":  "Bưu Cục Thị Trấn Phú Thứ-Tây Hoà-Phú Yên"
                                           },
                                           {
                                               "TongCOD":  79.13,
                                               "CODTuNop":  33.7,
                                               "TyLeTuNop":  42.6,
                                               "TenBuuCuc":  "Bưu Cục 396 Âu Cơ-Bùi Thị Xuân-Quy Nhơn-Bình Định"
                                           },
                                           {
                                               "TongCOD":  141.14,
                                               "CODTuNop":  52.92,
                                               "TyLeTuNop":  37.5,
                                               "TenBuuCuc":  "Bưu Cục Phạm Văn Đồng-Pơng Đrang-Krông Búk-Đắk Lắk"
                                           },
                                           {
                                               "TongCOD":  197.84,
                                               "CODTuNop":  77.09,
                                               "TyLeTuNop":  39,
                                               "TenBuuCuc":  "Bưu cục 236A Nguyễn Huệ-Đak Đoa-Gia Lai"
                                           },
                                           {
                                               "TongCOD":  34.9,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Điểm xử lý hàng Thôn Xuân Phong Nam-An Lão-Bình Định"
                                           },
                                           {
                                               "TongCOD":  94.81,
                                               "CODTuNop":  67.2,
                                               "TyLeTuNop":  70.9,
                                               "TenBuuCuc":  "Bưu Cục 838 Quang Trung-Phường Hoài Nhơn Tây-Bình Định"
                                           },
                                           {
                                               "TongCOD":  86.51,
                                               "CODTuNop":  11.06,
                                               "TyLeTuNop":  12.8,
                                               "TenBuuCuc":  "Bưu Cục 414 Trần Hưng Đạo-Mang Yang-Gia Lai"
                                           },
                                           {
                                               "TongCOD":  148.84,
                                               "CODTuNop":  49.22,
                                               "TyLeTuNop":  33.1,
                                               "TenBuuCuc":  "Bưu Cục 71 Quốc Lộ 1A-Xã Tuy An Bắc-Phú Yên"
                                           },
                                           {
                                               "TongCOD":  99.15,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 245 Hùng Vương-Ea Súp-Đắk Lắk"
                                           },
                                           {
                                               "TongCOD":  160.67,
                                               "CODTuNop":  2.83,
                                               "TyLeTuNop":  1.8,
                                               "TenBuuCuc":  "Bưu Cục KP4 Hòa Vinh- Đông Hòa- Phú Yên"
                                           },
                                           {
                                               "TongCOD":  90.94,
                                               "CODTuNop":  13.57,
                                               "TyLeTuNop":  14.9,
                                               "TenBuuCuc":  "Bưu Cục 127B Lê Văn Tám-Thị Trấn Kbang-Huyện Kbang-Gia Lai"
                                           },
                                           {
                                               "TongCOD":  112.9,
                                               "CODTuNop":  57.59,
                                               "TyLeTuNop":  51,
                                               "TenBuuCuc":  "Bưu Cục Km 30 Tân Hòa-Buôn Đôn-Đắk Lắk"
                                           },
                                           {
                                               "TongCOD":  120.37,
                                               "CODTuNop":  13.71,
                                               "TyLeTuNop":  11.4,
                                               "TenBuuCuc":  "Bưu Cục Ea Phê-Krông Pắk-Đắk Lắk"
                                           },
                                           {
                                               "TongCOD":  142.88,
                                               "CODTuNop":  45.18,
                                               "TyLeTuNop":  31.6,
                                               "TenBuuCuc":  "Bưu Cục 204 Trần Hưng Đạo-Buôn Hồ-Đắk Lắk"
                                           },
                                           {
                                               "TongCOD":  144.04,
                                               "CODTuNop":  96.59,
                                               "TyLeTuNop":  67.1,
                                               "TenBuuCuc":  "Bưu Cục 64 Hùng Vương-Thị Trấn Krông Năng-Đắk Lắk"
                                           },
                                           {
                                               "TongCOD":  232.73,
                                               "CODTuNop":  92.47,
                                               "TyLeTuNop":  39.7,
                                               "TenBuuCuc":  "Bưu cục 327A Cách Mạng Tháng 8-Pleiku-Gia Lai"
                                           },
                                           {
                                               "TongCOD":  179.19,
                                               "CODTuNop":  29.37,
                                               "TyLeTuNop":  16.4,
                                               "TenBuuCuc":  "Bưu Cục 132 Hoàng Hoa Thám-Buôn Ma Thuột-Đắk Lắk"
                                           },
                                           {
                                               "TongCOD":  278.22,
                                               "CODTuNop":  43.52,
                                               "TyLeTuNop":  15.6,
                                               "TenBuuCuc":  "Bưu Cục 03 Trần Hưng Đạo-Krông Pắc-Đắk Lắk"
                                           },
                                           {
                                               "TongCOD":  113.76,
                                               "CODTuNop":  65.71,
                                               "TyLeTuNop":  57.8,
                                               "TenBuuCuc":  "Bưu cục 73 Lê Lợi-Chư Pah-Gia Lai"
                                           },
                                           {
                                               "TongCOD":  67.04,
                                               "CODTuNop":  28.88,
                                               "TyLeTuNop":  43.1,
                                               "TenBuuCuc":  "Bưu Cục Thôn Chánh Hóa-Xã Cát Tiến-Bình Định"
                                           },
                                           {
                                               "TongCOD":  146.67,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Thôn Tân Lập Quốc Lộ 26-Xã Cư M\u0027ta-M\u0027Đrắk-Đăk Lắk"
                                           },
                                           {
                                               "TongCOD":  91.49,
                                               "CODTuNop":  17.56,
                                               "TyLeTuNop":  19.2,
                                               "TenBuuCuc":  "Bưu Cục Số 35 Phạm Văn Đồng-Tăng Bạt Hổ-Hoài Ân-Bình Định"
                                           },
                                           {
                                               "TongCOD":  150.78,
                                               "CODTuNop":  21.2,
                                               "TyLeTuNop":  14.1,
                                               "TenBuuCuc":  "Bưu Cục 39 Nguyễn Huệ-Ayun Pa-Gia Lai"
                                           },
                                           {
                                               "TongCOD":  190.03,
                                               "CODTuNop":  49.92,
                                               "TyLeTuNop":  26.3,
                                               "TenBuuCuc":  "Bưu Cục Km18-Cư Kuin-Đắk Lắk"
                                           },
                                           {
                                               "TongCOD":  192.82,
                                               "CODTuNop":  24.19,
                                               "TyLeTuNop":  12.5,
                                               "TenBuuCuc":  "Bưu Cục 271 Nguyễn Văn Cừ-Tân Lập-Đắk Lắk"
                                           },
                                           {
                                               "TongCOD":  102.77,
                                               "CODTuNop":  53.2,
                                               "TyLeTuNop":  51.8,
                                               "TenBuuCuc":  "Bưu cục 27 Nguyễn Tất Thành-Tuy Hòa-Phú Yên"
                                           },
                                           {
                                               "TongCOD":  251.47,
                                               "CODTuNop":  107.96,
                                               "TyLeTuNop":  42.9,
                                               "TenBuuCuc":  "Bưu Cục Khu Vực Kim Châu-An Nhơn-Bình Định"
                                           },
                                           {
                                               "TongCOD":  58.54,
                                               "CODTuNop":  6.54,
                                               "TyLeTuNop":  11.2,
                                               "TenBuuCuc":  "Bưu Cục Vận Hành Thôn Định Trị-Xã Vĩnh Thạnh-Bình Định"
                                           },
                                           {
                                               "TongCOD":  109.78,
                                               "CODTuNop":  31.94,
                                               "TyLeTuNop":  29.1,
                                               "TenBuuCuc":  "Bưu Cục 767 Hùng Vương-Quy Nhơn-Bình Định"
                                           },
                                           {
                                               "TongCOD":  74.74,
                                               "CODTuNop":  28.45,
                                               "TyLeTuNop":  38.1,
                                               "TenBuuCuc":  "Bưu Cục Đường Tỉnh 643-Xã Tuy An Nam-Phú Yên"
                                           },
                                           {
                                               "TongCOD":  109.3,
                                               "CODTuNop":  51.52,
                                               "TyLeTuNop":  47.1,
                                               "TenBuuCuc":  "Bưu Cục KM92 QL14-Xã Ea Drăng-Đắk Lắk"
                                           },
                                           {
                                               "TongCOD":  208.99,
                                               "CODTuNop":  1.5,
                                               "TyLeTuNop":  0.7,
                                               "TenBuuCuc":  "Bưu Cục 274 Hùng Vương-Tuy Hòa-Phú Yên"
                                           },
                                           {
                                               "TongCOD":  79.75,
                                               "CODTuNop":  40.69,
                                               "TyLeTuNop":  51,
                                               "TenBuuCuc":  "Bưu Cục 138 Trần Phú- Đồng Xuân- Phú Yên"
                                           },
                                           {
                                               "TongCOD":  122.76,
                                               "CODTuNop":  90.97,
                                               "TyLeTuNop":  74.1,
                                               "TenBuuCuc":  "Bưu Cục 70 Đường 30/4-Ea Tam-Buôn Ma Thuột-Đắk Lắk"
                                           },
                                           {
                                               "TongCOD":  166.29,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 125 Lý Chính Thắng-Buôn Ma Thuột-Đắk Lắk"
                                           },
                                           {
                                               "TongCOD":  273.82,
                                               "CODTuNop":  86.9,
                                               "TyLeTuNop":  31.7,
                                               "TenBuuCuc":  "Bưu Cục 297 Hùng Vương-CưMgar-Đắk Lắk"
                                           },
                                           {
                                               "TongCOD":  159.24,
                                               "CODTuNop":  54.84,
                                               "TyLeTuNop":  34.4,
                                               "TenBuuCuc":  "Bưu Cục 565 Trần Phú-Tuy Phước-Bình Định"
                                           },
                                           {
                                               "TongCOD":  237.84,
                                               "CODTuNop":  120.43,
                                               "TyLeTuNop":  50.6,
                                               "TenBuuCuc":  "Bưu Cục 569 Trường Chinh-Pleiku-Gia Lai"
                                           },
                                           {
                                               "TongCOD":  129.18,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 138 Hùng Vương-Phú Thiện-Gia Lai"
                                           },
                                           {
                                               "TongCOD":  94.96,
                                               "CODTuNop":  19.98,
                                               "TyLeTuNop":  21,
                                               "TenBuuCuc":  "Bưu Cục Thôn Đông Phước-Phú Hòa-Phú Yên"
                                           },
                                           {
                                               "TongCOD":  114.04,
                                               "CODTuNop":  41.56,
                                               "TyLeTuNop":  36.4,
                                               "TenBuuCuc":  "Bưu Cục 153 Hùng Vương-Ia Grai-Gia Lai"
                                           },
                                           {
                                               "TongCOD":  69.64,
                                               "CODTuNop":  32.88,
                                               "TyLeTuNop":  47.2,
                                               "TenBuuCuc":  "Bưu Cục 159 Tân Quảng-EaTóh-Krông Năng-Đắk Lắk"
                                           },
                                           {
                                               "TongCOD":  196.7,
                                               "CODTuNop":  82.33,
                                               "TyLeTuNop":  41.9,
                                               "TenBuuCuc":  "Bưu Cục 941 Giải Phóng-Ea Đrăng-Ea H\u0027Leo-Đắk Lắk"
                                           },
                                           {
                                               "TongCOD":  223.61,
                                               "CODTuNop":  57.67,
                                               "TyLeTuNop":  25.8,
                                               "TenBuuCuc":  "Bưu Cục KP3 TT Hai Riềng-Sông Hinh-Phú Yên"
                                           },
                                           {
                                               "TongCOD":  134.28,
                                               "CODTuNop":  30.99,
                                               "TyLeTuNop":  23.1,
                                               "TenBuuCuc":  "Bưu Cục Thị Trấn Buôn Trấp-Krông A Na-Đắk Lắk"
                                           },
                                           {
                                               "TongCOD":  200.85,
                                               "CODTuNop":  94.52,
                                               "TyLeTuNop":  47.1,
                                               "TenBuuCuc":  "Bưu cục DT644-Sông Cầu-Phú Yên"
                                           },
                                           {
                                               "TongCOD":  106.33,
                                               "CODTuNop":  34.21,
                                               "TyLeTuNop":  32.2,
                                               "TenBuuCuc":  "Bưu Cục TDP3 Thị Trấn Krông Kmart-Krông Bông-Đắk Lắk"
                                           },
                                           {
                                               "TongCOD":  236.04,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 99 Nguyễn Khuyến-Buôn Ma Thuột-Đắk Lắk"
                                           },
                                           {
                                               "TongCOD":  153.38,
                                               "CODTuNop":  36.99,
                                               "TyLeTuNop":  24.1,
                                               "TenBuuCuc":  "Bưu Cục 44 Quang Trung-Tây Sơn-Bình Định"
                                           }
                                       ]
                    },
                    {
                        "TyLeTuNop":  50.7,
                        "TenVung":  "TNT",
                        "TongCOD":  11415.91,
                        "Trend":  -0.4,
                        "CODTuNop":  5786.76,
                        "BuuCucList":  [
                                           {
                                               "TongCOD":  270.03,
                                               "CODTuNop":  209.79,
                                               "TyLeTuNop":  77.7,
                                               "TenBuuCuc":  "Bưu Cục Xóm 4-Giao Yến-Giao Thuỷ-Nam Định"
                                           },
                                           {
                                               "TongCOD":  284.01,
                                               "CODTuNop":  62.92,
                                               "TyLeTuNop":  22.2,
                                               "TenBuuCuc":  "Bưu Cục 64 Trần Hưng Đạo-TP.Ninh Bình-Ninh Bình"
                                           },
                                           {
                                               "TongCOD":  238.88,
                                               "CODTuNop":  238.88,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 1165 Đại Lộ Hùng Vương-Thanh Miếu-Phú Thọ"
                                           },
                                           {
                                               "TongCOD":  173.34,
                                               "CODTuNop":  110.3,
                                               "TyLeTuNop":  63.6,
                                               "TenBuuCuc":  "Bưu Cục Tân Lập-Cẩm Khê-Phú Thọ"
                                           },
                                           {
                                               "TongCOD":  137.46,
                                               "CODTuNop":  78.25,
                                               "TyLeTuNop":  56.9,
                                               "TenBuuCuc":  "Bưu Cục Thôn Làng Mạ-Xã Tam Dương Bắc-Vĩnh Phúc"
                                           },
                                           {
                                               "TongCOD":  172.57,
                                               "CODTuNop":  89.15,
                                               "TyLeTuNop":  51.7,
                                               "TenBuuCuc":  "Bưu Cục Khu 1-Thanh Minh-Phú Thọ"
                                           },
                                           {
                                               "TongCOD":  153.21,
                                               "CODTuNop":  70.91,
                                               "TyLeTuNop":  46.3,
                                               "TenBuuCuc":  "Bưu Cục Quốc Lộ 21A-Lý Thường Kiệt-Hà Nam"
                                           },
                                           {
                                               "TongCOD":  148.38,
                                               "CODTuNop":  148.38,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Quốc Lộ 2-Xã Phù Ninh-Phú Thọ"
                                           },
                                           {
                                               "TongCOD":  54.52,
                                               "CODTuNop":  27.72,
                                               "TyLeTuNop":  50.8,
                                               "TenBuuCuc":  "Bưu Cục Quốc Lộ 6-Xã Cao Phong-Hòa Bình"
                                           },
                                           {
                                               "TongCOD":  214.41,
                                               "CODTuNop":  15.33,
                                               "TyLeTuNop":  7.1,
                                               "TenBuuCuc":  "Bưu Cục 77 KĐT Yết Kiều-Phủ Lý-Hà Nam"
                                           },
                                           {
                                               "TongCOD":  241.58,
                                               "CODTuNop":  93.98,
                                               "TyLeTuNop":  38.9,
                                               "TenBuuCuc":  "Bưu Cục Số 13 Nguyễn Bỉnh Khiêm-Lộc Hoà-Nam Định"
                                           },
                                           {
                                               "TongCOD":  170.3,
                                               "CODTuNop":  114,
                                               "TyLeTuNop":  66.9,
                                               "TenBuuCuc":  "Bưu Cục Trực Ninh-Nam Định"
                                           },
                                           {
                                               "TongCOD":  171.28,
                                               "CODTuNop":  160.85,
                                               "TyLeTuNop":  93.9,
                                               "TenBuuCuc":  "Bưu Cục Đò Quan-Cửa Nam-Nam Định"
                                           },
                                           {
                                               "TongCOD":  183.1,
                                               "CODTuNop":  115.11,
                                               "TyLeTuNop":  62.9,
                                               "TenBuuCuc":  "Bưu Cục Khu Đoàn Kết Đường 19/5-Xã Đoan Hùng-Phú Thọ"
                                           },
                                           {
                                               "TongCOD":  37.44,
                                               "CODTuNop":  13.85,
                                               "TyLeTuNop":  37,
                                               "TenBuuCuc":  "Bưu Cục Điềm Khê-Gia Viễn-Ninh Bình"
                                           },
                                           {
                                               "TongCOD":  226.86,
                                               "CODTuNop":  116.26,
                                               "TyLeTuNop":  51.2,
                                               "TenBuuCuc":  "Bưu Cục Biên Hòa-Kim Thanh-Hà Nam"
                                           },
                                           {
                                               "TongCOD":  204.25,
                                               "CODTuNop":  74.11,
                                               "TyLeTuNop":  36.3,
                                               "TenBuuCuc":  "Bưu Cục QL38B Yên Khánh-Ý Yên-Nam Định"
                                           },
                                           {
                                               "TongCOD":  161.67,
                                               "CODTuNop":  74.03,
                                               "TyLeTuNop":  45.8,
                                               "TenBuuCuc":  "Bưu Cục Thôn 1-Phù Vân-Phủ Lý-Hà Nam"
                                           },
                                           {
                                               "TongCOD":  96.7,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Nguyễn Văn Hoan-Nho Quan-Ninh Bình"
                                           },
                                           {
                                               "TongCOD":  151,
                                               "CODTuNop":  145.52,
                                               "TyLeTuNop":  96.4,
                                               "TenBuuCuc":  "Bưu cục Khu đô thị Quán Chiền-Nam Dương-Nam Trực-Nam Định"
                                           },
                                           {
                                               "TongCOD":  107.4,
                                               "CODTuNop":  107.4,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 104 Hòa Bình-Phường Tân Hòa-Hòa Bình"
                                           },
                                           {
                                               "TongCOD":  107,
                                               "CODTuNop":  23.26,
                                               "TyLeTuNop":  21.7,
                                               "TenBuuCuc":  "Bưu Cục Tiểu Khu Liên Phương-Xã Đà Bắc-Hòa Bình"
                                           },
                                           {
                                               "TongCOD":  25.1,
                                               "CODTuNop":  7.63,
                                               "TyLeTuNop":  30.4,
                                               "TenBuuCuc":  "Bưu Cục Thôn Lãng Sơn-Sông Lô-Vĩnh Phúc"
                                           },
                                           {
                                               "TongCOD":  152.66,
                                               "CODTuNop":  118.21,
                                               "TyLeTuNop":  77.4,
                                               "TenBuuCuc":  "Bưu Cục Hợp Hải-Lâm Thao-Phú Thọ"
                                           },
                                           {
                                               "TongCOD":  222.14,
                                               "CODTuNop":  79.48,
                                               "TyLeTuNop":  35.8,
                                               "TenBuuCuc":  "Bưu Cục 26 Phạm Thận Duật-Bích Đào-Ninh Bình"
                                           },
                                           {
                                               "TongCOD":  259.56,
                                               "CODTuNop":  126.76,
                                               "TyLeTuNop":  48.8,
                                               "TenBuuCuc":  "Bưu Cục Trần Hưng Đạo-Quỳnh Lâm-Hòa Bình"
                                           },
                                           {
                                               "TongCOD":  237.88,
                                               "CODTuNop":  181.98,
                                               "TyLeTuNop":  76.5,
                                               "TenBuuCuc":  "Bưu Cục 85 Nguyễn Tất Thành-Liên Bảo-Vĩnh Yên-Vĩnh Phúc"
                                           },
                                           {
                                               "TongCOD":  176.33,
                                               "CODTuNop":  52.98,
                                               "TyLeTuNop":  30,
                                               "TenBuuCuc":  "Bưu Cục KĐT Phúc Sơn-Thượng Trưng-Vĩnh Tường-Vĩnh Phúc"
                                           },
                                           {
                                               "TongCOD":  138.26,
                                               "CODTuNop":  39.78,
                                               "TyLeTuNop":  28.8,
                                               "TenBuuCuc":  "Bưu Cục Ngô Thì Nhậm-Tam Điệp-Ninh Bình"
                                           },
                                           {
                                               "TongCOD":  167.24,
                                               "CODTuNop":  83.46,
                                               "TyLeTuNop":  49.9,
                                               "TenBuuCuc":  "Bưu Cục 362 Quốc Lộ 6-Xã Lương Sơn-Hòa Bình"
                                           },
                                           {
                                               "TongCOD":  196.08,
                                               "CODTuNop":  13.09,
                                               "TyLeTuNop":  6.7,
                                               "TenBuuCuc":  "Bưu Cục Thôn Xuân Hòa 1-Phúc Yên-Vĩnh Phúc"
                                           },
                                           {
                                               "TongCOD":  148.39,
                                               "CODTuNop":  99.65,
                                               "TyLeTuNop":  67.2,
                                               "TenBuuCuc":  "Bưu Cục Khu 10 Trường Chinh-Kim Đức-Việt Trì-Phú Thọ"
                                           },
                                           {
                                               "TongCOD":  129.16,
                                               "CODTuNop":  43.84,
                                               "TyLeTuNop":  33.9,
                                               "TenBuuCuc":  "Bưu cục Thôn Gò Gai-Xã Tiên Lữ-Vĩnh Phúc"
                                           },
                                           {
                                               "TongCOD":  141.33,
                                               "CODTuNop":  64.48,
                                               "TyLeTuNop":  45.6,
                                               "TenBuuCuc":  "Bưu Cục Đường Lạc Long Quân-Xã Hạ Hòa-Phú Thọ"
                                           },
                                           {
                                               "TongCOD":  117.32,
                                               "CODTuNop":  10.5,
                                               "TyLeTuNop":  9,
                                               "TenBuuCuc":  "Bưu Cục Phố Bưởi-Xã Kim Bôi-Hòa Bình"
                                           },
                                           {
                                               "TongCOD":  20.45,
                                               "CODTuNop":  1.95,
                                               "TyLeTuNop":  9.5,
                                               "TenBuuCuc":  "Bưu Cục Thôn Hòa Bình-Sông Lô-Vĩnh Phúc"
                                           },
                                           {
                                               "TongCOD":  94.91,
                                               "CODTuNop":  32.29,
                                               "TyLeTuNop":  34,
                                               "TenBuuCuc":  "Bưu Cục Thị trấn Hàng Trạm-Yên Thủy-Hòa Bình"
                                           },
                                           {
                                               "TongCOD":  125.57,
                                               "CODTuNop":  66.44,
                                               "TyLeTuNop":  52.9,
                                               "TenBuuCuc":  "Bưu Cục TDP Văn Sơn-Lập Thạch-Vĩnh Phúc"
                                           },
                                           {
                                               "TongCOD":  147.04,
                                               "CODTuNop":  95.59,
                                               "TyLeTuNop":  65,
                                               "TenBuuCuc":  "Bưu Cục TĐC Đường DT310-Xã Thiện Kế-Vĩnh Phúc"
                                           },
                                           {
                                               "TongCOD":  193.59,
                                               "CODTuNop":  83.09,
                                               "TyLeTuNop":  42.9,
                                               "TenBuuCuc":  "Bưu Cục Vĩnh Trụ-Lý Nhân-Hà Nam"
                                           },
                                           {
                                               "TongCOD":  158.56,
                                               "CODTuNop":  83.99,
                                               "TyLeTuNop":  53,
                                               "TenBuuCuc":  "Bưu Cục Lê Hồng Phong-Đồng Tâm-Vĩnh Yên-Vĩnh Phúc"
                                           },
                                           {
                                               "TongCOD":  131.38,
                                               "CODTuNop":  17.14,
                                               "TyLeTuNop":  13,
                                               "TenBuuCuc":  "Bưu Cục Quốc Lộ 12B-Yên Mô-Ninh Bình"
                                           },
                                           {
                                               "TongCOD":  112.61,
                                               "CODTuNop":  61.05,
                                               "TyLeTuNop":  54.2,
                                               "TenBuuCuc":  "Bưu Cục 132 Tân Long-Yên Lập-Phú Thọ"
                                           },
                                           {
                                               "TongCOD":  164.69,
                                               "CODTuNop":  24.79,
                                               "TyLeTuNop":  15.1,
                                               "TenBuuCuc":  "Bưu Cục Khu Đồng Cổng Huyện-Tam Dương-Vĩnh Phúc"
                                           },
                                           {
                                               "TongCOD":  119.48,
                                               "CODTuNop":  119.48,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Đồng Xuân-Thanh-Ba Phú Thọ"
                                           },
                                           {
                                               "TongCOD":  122.52,
                                               "CODTuNop":  22.29,
                                               "TyLeTuNop":  18.2,
                                               "TenBuuCuc":  "Bưu Cục Lê Đức Toản-Hùng Vương-Phúc Yên-Vĩnh Phúc"
                                           },
                                           {
                                               "TongCOD":  252.11,
                                               "CODTuNop":  160.83,
                                               "TyLeTuNop":  63.8,
                                               "TenBuuCuc":  "Bưu Cục Khu 3-Hải Hậu-Nam Định"
                                           },
                                           {
                                               "TongCOD":  193.37,
                                               "CODTuNop":  77.62,
                                               "TyLeTuNop":  40.1,
                                               "TenBuuCuc":  "Bưu Cục QL32 Khu Bình Dân-Xã Thanh Sơn-Phú Thọ"
                                           },
                                           {
                                               "TongCOD":  119.19,
                                               "CODTuNop":  9.79,
                                               "TyLeTuNop":  8.2,
                                               "TenBuuCuc":  "Bưu Cục 38 Lê Quý Đôn-Phường Duy Tiên-Hà Nam"
                                           },
                                           {
                                               "TongCOD":  161.14,
                                               "CODTuNop":  46.04,
                                               "TyLeTuNop":  28.6,
                                               "TenBuuCuc":  "Bưu Cục 150 Trần Hưng Đạo-Bình Lục-Hà Nam"
                                           },
                                           {
                                               "TongCOD":  124.95,
                                               "CODTuNop":  31.01,
                                               "TyLeTuNop":  24.8,
                                               "TenBuuCuc":  "Bưu Cục Phố Rịa-Nho Quan-Ninh Bình"
                                           },
                                           {
                                               "TongCOD":  55.03,
                                               "CODTuNop":  42.38,
                                               "TyLeTuNop":  77,
                                               "TenBuuCuc":  "Bưu Cục 214 Tuệ Tĩnh-Ninh Tiến-Ninh Bình"
                                           },
                                           {
                                               "TongCOD":  123.42,
                                               "CODTuNop":  84.11,
                                               "TyLeTuNop":  68.1,
                                               "TenBuuCuc":  "Bưu Cục Quốc lộ 21B-Mỹ Hưng-Mỹ Lộc-Nam Định"
                                           },
                                           {
                                               "TongCOD":  139.17,
                                               "CODTuNop":  44.39,
                                               "TyLeTuNop":  31.9,
                                               "TenBuuCuc":  "Bưu Cục Đường Đông Trung-Xã Tu Vũ-Phú Thọ"
                                           },
                                           {
                                               "TongCOD":  176.98,
                                               "CODTuNop":  101,
                                               "TyLeTuNop":  57.1,
                                               "TenBuuCuc":  "Bưu Cục Lạc Thiện-Quang Sơn-Ninh Bình"
                                           },
                                           {
                                               "TongCOD":  260.93,
                                               "CODTuNop":  260.93,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục QL37B-Thị trấn Liễu Đề-Nghĩa Hưng-Nam Định"
                                           },
                                           {
                                               "TongCOD":  114.95,
                                               "CODTuNop":  114.95,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu cục 57 Nguyễn Văn Nhượng-Vĩnh Tường-Vĩnh Phúc"
                                           },
                                           {
                                               "TongCOD":  171.79,
                                               "CODTuNop":  80.61,
                                               "TyLeTuNop":  46.9,
                                               "TenBuuCuc":  "Bưu Cục Quốc Lộ 6-Tân Lạc-Hoà Bình"
                                           },
                                           {
                                               "TongCOD":  185.32,
                                               "CODTuNop":  99.18,
                                               "TyLeTuNop":  53.5,
                                               "TenBuuCuc":  "Bưu Cục Thiện Vịnh-Cộng Hoà-Vụ Bản-Nam Định"
                                           },
                                           {
                                               "TongCOD":  87.88,
                                               "CODTuNop":  32.57,
                                               "TyLeTuNop":  37.1,
                                               "TenBuuCuc":  "Bưu Cục Khu 10 Tân Phú-Tân Sơn-Phú Thọ"
                                           },
                                           {
                                               "TongCOD":  194.38,
                                               "CODTuNop":  68.8,
                                               "TyLeTuNop":  35.4,
                                               "TenBuuCuc":  "Bưu cục Xóm 11B-Xuân Vinh-Xuân Trường-Nam Định"
                                           },
                                           {
                                               "TongCOD":  185.92,
                                               "CODTuNop":  106.24,
                                               "TyLeTuNop":  57.1,
                                               "TenBuuCuc":  "Bưu Cục Xóm 3-Xã Hải An-Nam Định"
                                           },
                                           {
                                               "TongCOD":  141.96,
                                               "CODTuNop":  135.68,
                                               "TyLeTuNop":  95.6,
                                               "TenBuuCuc":  "Bưu Cục Đường Trần Tự Khánh-Mỹ Lộc-Nam Định"
                                           },
                                           {
                                               "TongCOD":  94.86,
                                               "CODTuNop":  26.63,
                                               "TyLeTuNop":  28.1,
                                               "TenBuuCuc":  "Bưu Cục Hương Nộn-Tam Nông-Phú Thọ"
                                           },
                                           {
                                               "TongCOD":  179.04,
                                               "CODTuNop":  88.68,
                                               "TyLeTuNop":  49.5,
                                               "TenBuuCuc":  "Bưu Cục Số nhà 03-Lốc Mới-Vụ Bản-Lạc Sơn-Hoà Bình"
                                           },
                                           {
                                               "TongCOD":  127.57,
                                               "CODTuNop":  68.18,
                                               "TyLeTuNop":  53.4,
                                               "TenBuuCuc":  "Bưu Cục Khu 1 Thị trấn Chi Nê-Lạc Thủy-Hòa Bình"
                                           },
                                           {
                                               "TongCOD":  190.29,
                                               "CODTuNop":  85.92,
                                               "TyLeTuNop":  45.1,
                                               "TenBuuCuc":  "Bưu Cục Quốc Lộ 10-Yên Khánh-Ninh Bình"
                                           },
                                           {
                                               "TongCOD":  151.72,
                                               "CODTuNop":  88.63,
                                               "TyLeTuNop":  58.4,
                                               "TenBuuCuc":  "Bưu Cục TDP Thắng Lợi-Hương Canh-Bình Xuyên-Vĩnh Phúc"
                                           },
                                           {
                                               "TongCOD":  111.53,
                                               "CODTuNop":  8.38,
                                               "TyLeTuNop":  7.5,
                                               "TenBuuCuc":  "Bưu Cục Đinh Đức Thiện-Xã Vạn Thắng-Nam Định"
                                           },
                                           {
                                               "TongCOD":  131.24,
                                               "CODTuNop":  80.32,
                                               "TyLeTuNop":  61.2,
                                               "TenBuuCuc":  "Bưu Cục Gia Viễn-Ninh Bình"
                                           },
                                           {
                                               "TongCOD":  201.61,
                                               "CODTuNop":  104.9,
                                               "TyLeTuNop":  52,
                                               "TenBuuCuc":  "Bưu Cục Thôn 1 Tâng-Xã Thanh Liêm-Hà Nam"
                                           },
                                           {
                                               "TongCOD":  331.91,
                                               "CODTuNop":  119.04,
                                               "TyLeTuNop":  35.9,
                                               "TenBuuCuc":  "Bưu Cục TDP Nho Lâm-Xã Tam Hồng-Vĩnh Phúc"
                                           }
                                       ]
                    },
                    {
                        "TyLeTuNop":  69,
                        "TenVung":  "TTB",
                        "TongCOD":  13214.74,
                        "Trend":  3.7,
                        "CODTuNop":  9123.05,
                        "BuuCucList":  [
                                           {
                                               "TongCOD":  201.74,
                                               "CODTuNop":  53.15,
                                               "TyLeTuNop":  26.3,
                                               "TenBuuCuc":  "Bưu Cục Thôn Phổ Lại-Thanh An-Cam Lộ-Quảng Trị"
                                           },
                                           {
                                               "TongCOD":  185.85,
                                               "CODTuNop":  181.74,
                                               "TyLeTuNop":  97.8,
                                               "TenBuuCuc":  "Bưu Cục Nguyễn Chí Thanh-Núi Thành-Quảng Nam"
                                           },
                                           {
                                               "TongCOD":  18.1,
                                               "CODTuNop":  1.15,
                                               "TyLeTuNop":  6.4,
                                               "TenBuuCuc":  "Bưu Cục Vận Hành Nam Trà My-Quảng Nam"
                                           },
                                           {
                                               "TongCOD":  96.85,
                                               "CODTuNop":  24.37,
                                               "TyLeTuNop":  25.2,
                                               "TenBuuCuc":  "Bưu Cục 215 Hùng Vương-Đắk Hà-Kon Tum"
                                           },
                                           {
                                               "TongCOD":  167.56,
                                               "CODTuNop":  156.25,
                                               "TyLeTuNop":  93.3,
                                               "TenBuuCuc":  "Bưu Cục 62 Bắc Sơn-Q.Cẩm Lệ-Đà Nẵng"
                                           },
                                           {
                                               "TongCOD":  40.96,
                                               "CODTuNop":  27.38,
                                               "TyLeTuNop":  66.8,
                                               "TenBuuCuc":  "Bưu Cục Vận Hành Bắc Trà My-Quảng Nam"
                                           },
                                           {
                                               "TongCOD":  243.72,
                                               "CODTuNop":  124.02,
                                               "TyLeTuNop":  50.9,
                                               "TenBuuCuc":  "Bưu Cục 231 Võ An Ninh-Q.Cẩm Lệ-Đà Nẵng"
                                           },
                                           {
                                               "TongCOD":  210.26,
                                               "CODTuNop":  90.01,
                                               "TyLeTuNop":  42.8,
                                               "TenBuuCuc":  "Bưu Cục 25 Phạm Văn Đồng-TP.Đồng Hới-Quảng Bình"
                                           },
                                           {
                                               "TongCOD":  330.57,
                                               "CODTuNop":  270.87,
                                               "TyLeTuNop":  81.9,
                                               "TenBuuCuc":  "Bưu Cục 184 Nguyễn Tất Thành-Hội An-Quảng Nam"
                                           },
                                           {
                                               "TongCOD":  134.31,
                                               "CODTuNop":  121.07,
                                               "TyLeTuNop":  90.1,
                                               "TenBuuCuc":  "Bưu Cục 140 Tôn Đức Thắng-An Sơn-Tam Kỳ-Quảng Nam"
                                           },
                                           {
                                               "TongCOD":  190.2,
                                               "CODTuNop":  95.96,
                                               "TyLeTuNop":  50.5,
                                               "TenBuuCuc":  "Bưu Cục 274 Điện Biên Phủ-Duy Xuyên-Quảng Nam"
                                           },
                                           {
                                               "TongCOD":  290.96,
                                               "CODTuNop":  290.93,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 404 Lê Văn Hiến-Khuê Mỹ-Ngũ Hành Sơn-Đà Nẵng"
                                           },
                                           {
                                               "TongCOD":  368.89,
                                               "CODTuNop":  368.89,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 244 Lê Lợi-Đông Lương-Đông Hà-Quảng Trị"
                                           },
                                           {
                                               "TongCOD":  149.51,
                                               "CODTuNop":  58.85,
                                               "TyLeTuNop":  39.4,
                                               "TenBuuCuc":  "Bưu Cục Thôn Dy Lộc-Quảng Tùng-Quảng Trạch-Quảng Bình"
                                           },
                                           {
                                               "TongCOD":  57.69,
                                               "CODTuNop":  7.31,
                                               "TyLeTuNop":  12.7,
                                               "TenBuuCuc":  "Bưu Cục Vận Hành Châu Ổ-Bình Sơn-Quảng Ngãi"
                                           },
                                           {
                                               "TongCOD":  330.78,
                                               "CODTuNop":  330.78,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 280 Lê Đại Hành-Q.Cẩm Lệ-Đà Nẵng"
                                           },
                                           {
                                               "TongCOD":  100.34,
                                               "CODTuNop":  20.95,
                                               "TyLeTuNop":  20.9,
                                               "TenBuuCuc":  "Bưu cục Thị trấn Di Lăng-Sơn Hà-Quảng Ngãi"
                                           },
                                           {
                                               "TongCOD":  234.66,
                                               "CODTuNop":  234.66,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 258 Lê Duẩn-Hướng Hóa-Quảng Trị"
                                           },
                                           {
                                               "TongCOD":  12.36,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Điểm Xử Lý Hàng Xã Đảo Tân Hiệp-Quảng Nam"
                                           },
                                           {
                                               "TongCOD":  83.66,
                                               "CODTuNop":  19.53,
                                               "TyLeTuNop":  23.3,
                                               "TenBuuCuc":  "Bưu Cục Phong Nha-Bố Trạch-Quảng Bình"
                                           },
                                           {
                                               "TongCOD":  129.61,
                                               "CODTuNop":  83.12,
                                               "TyLeTuNop":  64.1,
                                               "TenBuuCuc":  "Bưu Cục KDC Nam Khang-Đức Thạnh-Mộ Đức-Quảng Ngãi"
                                           },
                                           {
                                               "TongCOD":  179.3,
                                               "CODTuNop":  179.3,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 36 Lý Thái Tổ-Điện Bàn-Quảng Nam"
                                           },
                                           {
                                               "TongCOD":  187.25,
                                               "CODTuNop":  159.71,
                                               "TyLeTuNop":  85.3,
                                               "TenBuuCuc":  "Bưu Cục Bình Phục-Thăng Bình-Quảng Nam"
                                           },
                                           {
                                               "TongCOD":  92.94,
                                               "CODTuNop":  37.26,
                                               "TyLeTuNop":  40.1,
                                               "TenBuuCuc":  "Bưu Cục Vận Hành 1056 Quang-Chánh Lộ-Quảng Ngãi"
                                           },
                                           {
                                               "TongCOD":  60.77,
                                               "CODTuNop":  22.54,
                                               "TyLeTuNop":  37.1,
                                               "TenBuuCuc":  "Bưu Cục Vận Hành Ba Tơ-Quảng Ngãi"
                                           },
                                           {
                                               "TongCOD":  181.45,
                                               "CODTuNop":  75.77,
                                               "TyLeTuNop":  41.8,
                                               "TenBuuCuc":  "Bưu Cục 58 Lý Thường Kiệt-Lệ Thủy-Quảng Bình"
                                           },
                                           {
                                               "TongCOD":  129.22,
                                               "CODTuNop":  26.92,
                                               "TyLeTuNop":  20.8,
                                               "TenBuuCuc":  "Bưu Cục 45 Trần Hưng Đạo-Thị trấn Plei Kần-Ngọc Hồi-Kon Tum"
                                           },
                                           {
                                               "TongCOD":  246.63,
                                               "CODTuNop":  216.51,
                                               "TyLeTuNop":  87.8,
                                               "TenBuuCuc":  "Bưu Cục 621 Điện Biên Phủ-Q.Thanh Khê-Đà Nẵng"
                                           },
                                           {
                                               "TongCOD":  206.92,
                                               "CODTuNop":  94.13,
                                               "TyLeTuNop":  45.5,
                                               "TenBuuCuc":  "Bưu Cục 94 Thanh Tịnh-Hòa Minh-Liên Chiểu-Đà Nẵng"
                                           },
                                           {
                                               "TongCOD":  137.73,
                                               "CODTuNop":  137.11,
                                               "TyLeTuNop":  99.5,
                                               "TenBuuCuc":  "Bưu Cục 18 Đường 3/2-Hải Lăng-Quảng Trị"
                                           },
                                           {
                                               "TongCOD":  218.5,
                                               "CODTuNop":  103.38,
                                               "TyLeTuNop":  47.3,
                                               "TenBuuCuc":  "Bưu Cục 390 Phò Trạch-Phong Điền-Thừa Thiên Huế"
                                           },
                                           {
                                               "TongCOD":  200.21,
                                               "CODTuNop":  12.5,
                                               "TyLeTuNop":  6.2,
                                               "TenBuuCuc":  "Bưu Cục 08 Đổ Quang-Thị trấn Đông Phú-Quế Sơn-Quảng Nam"
                                           },
                                           {
                                               "TongCOD":  80.07,
                                               "CODTuNop":  5.28,
                                               "TyLeTuNop":  6.6,
                                               "TenBuuCuc":  "Bưu Cục Vận Hành Sa Thầy-Quảng Ngãi"
                                           },
                                           {
                                               "TongCOD":  82.84,
                                               "CODTuNop":  22.39,
                                               "TyLeTuNop":  27,
                                               "TenBuuCuc":  "Bưu Cục 129 Võ Nguyên Giáp-Minh Hóa-Quảng Bình"
                                           },
                                           {
                                               "TongCOD":  330.35,
                                               "CODTuNop":  214.11,
                                               "TyLeTuNop":  64.8,
                                               "TenBuuCuc":  "Bưu Cục 252 Kinh Dương Vương-Thuận An-Phú Vang-Huế"
                                           },
                                           {
                                               "TongCOD":  115.3,
                                               "CODTuNop":  56.14,
                                               "TyLeTuNop":  48.7,
                                               "TenBuuCuc":  "Bưu Cục Phường Bắc Lý-Đồng Hới-Quảng Bình"
                                           },
                                           {
                                               "TongCOD":  38.24,
                                               "CODTuNop":  7.35,
                                               "TyLeTuNop":  19.2,
                                               "TenBuuCuc":  "Bưu Cục Quốc lộ 24-Kon Rẫy-Kon Tum"
                                           },
                                           {
                                               "TongCOD":  58.71,
                                               "CODTuNop":  0.13,
                                               "TyLeTuNop":  0.2,
                                               "TenBuuCuc":  "Bưu Cục Vận Hành Xã Tây Giang-Quảng Nam"
                                           },
                                           {
                                               "TongCOD":  55.31,
                                               "CODTuNop":  17.43,
                                               "TyLeTuNop":  31.5,
                                               "TenBuuCuc":  "Điểm xử lý hàng Nam Đông-Thừa Thiên Huế"
                                           },
                                           {
                                               "TongCOD":  250.24,
                                               "CODTuNop":  250.24,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Trường Sơn-Hòa Vang-Đà Nẵng"
                                           },
                                           {
                                               "TongCOD":  22.91,
                                               "CODTuNop":  18.68,
                                               "TyLeTuNop":  81.6,
                                               "TenBuuCuc":  "Điểm Xử Lý Hàng Ia Hdrai-Quảng Ngãi"
                                           },
                                           {
                                               "TongCOD":  135.46,
                                               "CODTuNop":  135.46,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 183 Trần Phú-Vĩnh Linh-Quảng Trị"
                                           },
                                           {
                                               "TongCOD":  317.71,
                                               "CODTuNop":  278.68,
                                               "TyLeTuNop":  87.7,
                                               "TenBuuCuc":  "Bưu Cục Phường Thanh Bình-Hải Châu-Đà Nẵng"
                                           },
                                           {
                                               "TongCOD":  247.31,
                                               "CODTuNop":  247.31,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 345C Ngô Quyền-Sơn Trà-Đà Nẵng"
                                           },
                                           {
                                               "TongCOD":  177.39,
                                               "CODTuNop":  177.39,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 114 Nguyễn Văn Trỗi-TP Tam Kỳ-Quảng Nam"
                                           },
                                           {
                                               "TongCOD":  243.96,
                                               "CODTuNop":  41.78,
                                               "TyLeTuNop":  17.1,
                                               "TenBuuCuc":  "Bưu Cục 106 Kỳ Đồng-Q.Thanh Khê-Đà Nẵng"
                                           },
                                           {
                                               "TongCOD":  253.05,
                                               "CODTuNop":  165.73,
                                               "TyLeTuNop":  65.5,
                                               "TenBuuCuc":  "Bưu Cục 721 Võ Nguyên Giáp-Quảng Ngãi-Quảng Ngãi"
                                           },
                                           {
                                               "TongCOD":  89.47,
                                               "CODTuNop":  26.81,
                                               "TyLeTuNop":  30,
                                               "TenBuuCuc":  "Bưu Cục Vận Hành Đảo Lý Sơn-Quảng Ngãi"
                                           },
                                           {
                                               "TongCOD":  60.2,
                                               "CODTuNop":  60.2,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Vận Hành Thôn Tiền Phong-Xã Bắc Trạch-Quảng Bình"
                                           },
                                           {
                                               "TongCOD":  132,
                                               "CODTuNop":  43.43,
                                               "TyLeTuNop":  32.9,
                                               "TenBuuCuc":  "Bưu Cục Xóm Trảng Bông-Bình Sơn-Quảng Ngãi"
                                           },
                                           {
                                               "TongCOD":  24.69,
                                               "CODTuNop":  9.94,
                                               "TyLeTuNop":  40.3,
                                               "TenBuuCuc":  "Bưu Cục Vận Hành Xã Nam Giang-Đà Nẵng"
                                           },
                                           {
                                               "TongCOD":  70.81,
                                               "CODTuNop":  37.59,
                                               "TyLeTuNop":  53.1,
                                               "TenBuuCuc":  "Bưu cục 61 Tôn Đức Thắng-Tiên Phước-Quảng Nam"
                                           },
                                           {
                                               "TongCOD":  127.07,
                                               "CODTuNop":  127.07,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 63 Trần Thủ Độ-Điện Bàn-Quảng Nam"
                                           },
                                           {
                                               "TongCOD":  592.87,
                                               "CODTuNop":  592.87,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 352 Lê Duẩn-Phú Thuận-Huế"
                                           },
                                           {
                                               "TongCOD":  532.42,
                                               "CODTuNop":  494.46,
                                               "TyLeTuNop":  92.9,
                                               "TenBuuCuc":  "Bưu Cục G22 Cao Đình Độ-Phường Trường An-TP Huế"
                                           },
                                           {
                                               "TongCOD":  58.23,
                                               "CODTuNop":  0.97,
                                               "TyLeTuNop":  1.7,
                                               "TenBuuCuc":  "Bưu Cục Vận Hành Kon Plông-Kon Tum"
                                           },
                                           {
                                               "TongCOD":  13.93,
                                               "CODTuNop":  9.62,
                                               "TyLeTuNop":  69.1,
                                               "TenBuuCuc":  "Bưu Cục Vận Hành Xã Thạnh Mỹ-Đà Nẵng"
                                           },
                                           {
                                               "TongCOD":  179.21,
                                               "CODTuNop":  179.21,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu cục Thọ Quang-Sơn Trà-Đà Nẵng"
                                           },
                                           {
                                               "TongCOD":  136.91,
                                               "CODTuNop":  136.91,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 306 Hùng Vương-Bố Trạch-Quảng Bình"
                                           },
                                           {
                                               "TongCOD":  44.2,
                                               "CODTuNop":  27.05,
                                               "TyLeTuNop":  61.2,
                                               "TenBuuCuc":  "Bưu Cục Vận Hành Xã Đăk Pék-Kon Tum"
                                           },
                                           {
                                               "TongCOD":  330.94,
                                               "CODTuNop":  138.2,
                                               "TyLeTuNop":  41.8,
                                               "TenBuuCuc":  "Bưu Cục 106 Quang Trung-Đại Lộc-Quảng Nam"
                                           },
                                           {
                                               "TongCOD":  190.23,
                                               "CODTuNop":  99.06,
                                               "TyLeTuNop":  52.1,
                                               "TenBuuCuc":  "Bưu Cục 152 Hùng Vương-Ba Đồn-Quảng Bình"
                                           },
                                           {
                                               "TongCOD":  134.98,
                                               "CODTuNop":  34.62,
                                               "TyLeTuNop":  25.6,
                                               "TenBuuCuc":  "Bưu Cục Châu Ổ-Bình Sơn-Quảng Ngãi"
                                           },
                                           {
                                               "TongCOD":  295.41,
                                               "CODTuNop":  269.86,
                                               "TyLeTuNop":  91.4,
                                               "TenBuuCuc":  "Bưu Cục 217 Tiểu La-Hòa Cường-Đà Nẵng"
                                           },
                                           {
                                               "TongCOD":  176.73,
                                               "CODTuNop":  176.73,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 366 Hai Bà Trưng-Quảng Ngãi-Quảng Ngãi"
                                           },
                                           {
                                               "TongCOD":  70.89,
                                               "CODTuNop":  3.18,
                                               "TyLeTuNop":  4.5,
                                               "TenBuuCuc":  "Bưu Cục Thôn Quảng Phú-A Lưới-Thừa Thiên Huế"
                                           },
                                           {
                                               "TongCOD":  175.84,
                                               "CODTuNop":  105.9,
                                               "TyLeTuNop":  60.2,
                                               "TenBuuCuc":  "Bưu Cục 14 Lê Hoàn-Phường Hương Trà-Huế"
                                           },
                                           {
                                               "TongCOD":  151.74,
                                               "CODTuNop":  71.81,
                                               "TyLeTuNop":  47.3,
                                               "TenBuuCuc":  "Bưu Cục 463 Phạm Văn Đồng-Lê Lợi-TP.KonTum-KonTum"
                                           },
                                           {
                                               "TongCOD":  135.06,
                                               "CODTuNop":  51.65,
                                               "TyLeTuNop":  38.2,
                                               "TenBuuCuc":  "Bưu Cục 70 Đỗ Quang Thắng-Đức Phổ-Quảng Ngãi"
                                           },
                                           {
                                               "TongCOD":  6.81,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 26 Huỳnh Thúc Kháng-Thị Trấn P\u0027rao-Đông Giang-Quảng Nam"
                                           },
                                           {
                                               "TongCOD":  62.38,
                                               "CODTuNop":  20.49,
                                               "TyLeTuNop":  32.8,
                                               "TenBuuCuc":  "Bưu Cục Vận Hành Đắk Tô-Quảng Ngãi"
                                           },
                                           {
                                               "TongCOD":  126.83,
                                               "CODTuNop":  69.73,
                                               "TyLeTuNop":  55,
                                               "TenBuuCuc":  "Bưu Cục Võ Ninh-Quảng Ninh-Quảng Bình"
                                           },
                                           {
                                               "TongCOD":  172.6,
                                               "CODTuNop":  81.67,
                                               "TyLeTuNop":  47.3,
                                               "TenBuuCuc":  "Bưu Cục 43 Nguyễn Nghiêm-Huyện Tư Nghĩa-Quảng Ngãi"
                                           },
                                           {
                                               "TongCOD":  226.75,
                                               "CODTuNop":  226.75,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục Bàu vàng 6-Hòa Khánh Bắc-Liên Chiểu-Đà Nẵng"
                                           },
                                           {
                                               "TongCOD":  231.53,
                                               "CODTuNop":  154.61,
                                               "TyLeTuNop":  66.8,
                                               "TenBuuCuc":  "Bưu Cục 955 Nguyễn Tất Thành-Hương Thủy-Thừa Thiên Huế"
                                           },
                                           {
                                               "TongCOD":  48.78,
                                               "CODTuNop":  36.1,
                                               "TyLeTuNop":  74,
                                               "TenBuuCuc":  "Bưu Cục Thôn Liên Bằng-Xã Hương Thọ-Huế"
                                           },
                                           {
                                               "TongCOD":  6.32,
                                               "CODTuNop":  1.65,
                                               "TyLeTuNop":  26,
                                               "TenBuuCuc":  "Điểm Xử Lý Hàng Mo Rai-Quảng Ngãi"
                                           },
                                           {
                                               "TongCOD":  265.99,
                                               "CODTuNop":  265.99,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 202 Lý Thánh Tông-Phú Lộc-Huế"
                                           },
                                           {
                                               "TongCOD":  156.88,
                                               "CODTuNop":  117.07,
                                               "TyLeTuNop":  74.6,
                                               "TenBuuCuc":  "Bưu cục 1056 Quang Trung -Chánh Lộ -Quảng Ngãi"
                                           },
                                           {
                                               "TongCOD":  239.22,
                                               "CODTuNop":  157.49,
                                               "TyLeTuNop":  65.8,
                                               "TenBuuCuc":  "Bưu Cục 303 Đào Duy Từ-Thắng Lợi-Kon Tum"
                                           },
                                           {
                                               "TongCOD":  103.86,
                                               "CODTuNop":  21.37,
                                               "TyLeTuNop":  20.6,
                                               "TenBuuCuc":  "Bưu Cục 67 Lý Thái Tổ-Xã Đồng Lê-Quảng Bình"
                                           },
                                           {
                                               "TongCOD":  14.61,
                                               "CODTuNop":  8.77,
                                               "TyLeTuNop":  60,
                                               "TenBuuCuc":  "Bưu Cục Vận Hành Tu Mơ Rông-Đắk Tô-Kon Tum"
                                           }
                                       ]
                    },
                    {
                        "TyLeTuNop":  31.2,
                        "TenVung":  "XBG",
                        "TongCOD":  19371.59,
                        "Trend":  5,
                        "CODTuNop":  6046.4,
                        "BuuCucList":  [
                                           {
                                               "TongCOD":  259.57,
                                               "CODTuNop":  13.73,
                                               "TyLeTuNop":  5.3,
                                               "TenBuuCuc":  "Bưu Cục Xã Tân Hưng-Vĩnh Bảo-Hải Phòng"
                                           },
                                           {
                                               "TongCOD":  72.54,
                                               "CODTuNop":  40.77,
                                               "TyLeTuNop":  56.2,
                                               "TenBuuCuc":  "Bưu cục Thị Trấn Bình Gia-Bình Gia-Lạng Sơn"
                                           },
                                           {
                                               "TongCOD":  140.92,
                                               "CODTuNop":  1.91,
                                               "TyLeTuNop":  1.4,
                                               "TenBuuCuc":  "Bưu Cục Thị Trấn Cao Thượng-Tân Yên-Bắc Giang"
                                           },
                                           {
                                               "TongCOD":  170.76,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục An Dương-Hải Phòng"
                                           },
                                           {
                                               "TongCOD":  89.82,
                                               "CODTuNop":  37.54,
                                               "TyLeTuNop":  41.8,
                                               "TenBuuCuc":  "Bưu Cục 215 Đề Nắm-Yên Thế-Bắc Giang"
                                           },
                                           {
                                               "TongCOD":  260.31,
                                               "CODTuNop":  189.66,
                                               "TyLeTuNop":  72.9,
                                               "TenBuuCuc":  "Bưu Cục 18 Bạch Năng Thi-Hải Dương"
                                           },
                                           {
                                               "TongCOD":  202.68,
                                               "CODTuNop":  133.09,
                                               "TyLeTuNop":  65.7,
                                               "TenBuuCuc":  "Bưu Cục Đường 25/5 Thị Trấn Thanh Hà-Hải Dương"
                                           },
                                           {
                                               "TongCOD":  309.99,
                                               "CODTuNop":  186.15,
                                               "TyLeTuNop":  60.1,
                                               "TenBuuCuc":  "Bưu Cục KCN Vĩnh Niệm-Q.Lê Chân-Hải Phòng"
                                           },
                                           {
                                               "TongCOD":  194.28,
                                               "CODTuNop":  125.71,
                                               "TyLeTuNop":  64.7,
                                               "TenBuuCuc":  "Bưu Cục Thôn Đương Thắng-Kiến Thụy-Hải Phòng"
                                           },
                                           {
                                               "TongCOD":  215.51,
                                               "CODTuNop":  18.59,
                                               "TyLeTuNop":  8.6,
                                               "TenBuuCuc":  "Bưu Cục 415 Nguyễn Chế Nghĩa-Gia Lộc-Hải Dương"
                                           },
                                           {
                                               "TongCOD":  378.31,
                                               "CODTuNop":  146.36,
                                               "TyLeTuNop":  38.7,
                                               "TenBuuCuc":  "Bưu Cục Thôn Chùa-Việt Yên-Bắc Giang"
                                           },
                                           {
                                               "TongCOD":  365.19,
                                               "CODTuNop":  100.56,
                                               "TyLeTuNop":  27.5,
                                               "TenBuuCuc":  "Bưu Cục Vĩnh Niệm-Lê Chân-Hải Phòng"
                                           },
                                           {
                                               "TongCOD":  302.89,
                                               "CODTuNop":  190.43,
                                               "TyLeTuNop":  62.9,
                                               "TenBuuCuc":  "Bưu Cục TDP Mới QL31-Phường Chũ-Bắc Giang"
                                           },
                                           {
                                               "TongCOD":  49.18,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Đào Viên-Quế Võ-Bắc Ninh"
                                           },
                                           {
                                               "TongCOD":  172.79,
                                               "CODTuNop":  87.4,
                                               "TyLeTuNop":  50.6,
                                               "TenBuuCuc":  "Bưu Cục Xuân Lai-Gia Bình-Bắc Ninh"
                                           },
                                           {
                                               "TongCOD":  176.6,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục QL10 Chợ Tổng-Thủy Nguyên-Hải Phòng"
                                           },
                                           {
                                               "TongCOD":  130.49,
                                               "CODTuNop":  130.41,
                                               "TyLeTuNop":  99.9,
                                               "TenBuuCuc":  "Bưu Cục Phố Bình Minh-Xã Hiệp Hòa-Bắc Giang"
                                           },
                                           {
                                               "TongCOD":  146.93,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 140 Kiều Hạ-Đông Hải 2-Hải An-Hải Phòng"
                                           },
                                           {
                                               "TongCOD":  168.71,
                                               "CODTuNop":  59.76,
                                               "TyLeTuNop":  35.4,
                                               "TenBuuCuc":  "Bưu Cục 204 Lê Lai-Máy Chai-Ngô Quyền-Hải Phòng"
                                           },
                                           {
                                               "TongCOD":  182.61,
                                               "CODTuNop":  109.1,
                                               "TyLeTuNop":  59.7,
                                               "TenBuuCuc":  "Bưu Cục Thôn Xuân Sơn-An Thắng-An Lão-Hải Phòng"
                                           },
                                           {
                                               "TongCOD":  151.66,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Điểm xử lý hàng Huyện Đảo Cô Tô-Quảng Ninh"
                                           },
                                           {
                                               "TongCOD":  257.83,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Mao Trung-Phượng Mao-Quế Võ-Bắc Ninh"
                                           },
                                           {
                                               "TongCOD":  253.61,
                                               "CODTuNop":  155.11,
                                               "TyLeTuNop":  61.2,
                                               "TenBuuCuc":  "Bưu Cục TDP Minh Phượng-Nham Biền-Bắc Giang"
                                           },
                                           {
                                               "TongCOD":  225.17,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu cục Cẩm Thạch-Cẩm Phả-Quảng Ninh"
                                           },
                                           {
                                               "TongCOD":  338.61,
                                               "CODTuNop":  317.96,
                                               "TyLeTuNop":  93.9,
                                               "TenBuuCuc":  "Bưu Cục 132 Quang Trung-Uông Bí-Quảng Ninh"
                                           },
                                           {
                                               "TongCOD":  169.36,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Thôn 1-Hải Hà-Quảng Ninh"
                                           },
                                           {
                                               "TongCOD":  212.32,
                                               "CODTuNop":  131.37,
                                               "TyLeTuNop":  61.9,
                                               "TenBuuCuc":  "Bưu Cục Thôn Phục Lễ-Bình Giang-Hải Dương"
                                           },
                                           {
                                               "TongCOD":  151.53,
                                               "CODTuNop":  132,
                                               "TyLeTuNop":  87.1,
                                               "TenBuuCuc":  "Bưu Cục 386 Lê Duẩn-Kiến An-Hải Phòng"
                                           },
                                           {
                                               "TongCOD":  183.9,
                                               "CODTuNop":  22.58,
                                               "TyLeTuNop":  12.3,
                                               "TenBuuCuc":  "Bưu Cục DT359-Thủy Nguyên-Hải Phòng"
                                           },
                                           {
                                               "TongCOD":  102.36,
                                               "CODTuNop":  67.66,
                                               "TyLeTuNop":  66.1,
                                               "TenBuuCuc":  "Bưu Cục Đường Đại Huề-Chi Lăng-Lạng Sơn"
                                           },
                                           {
                                               "TongCOD":  351.1,
                                               "CODTuNop":  165.76,
                                               "TyLeTuNop":  47.2,
                                               "TenBuuCuc":  "Bưu Cục Dĩnh Kế- TP.Bắc Giang-Bắc Giang"
                                           },
                                           {
                                               "TongCOD":  189.95,
                                               "CODTuNop":  50.01,
                                               "TyLeTuNop":  26.3,
                                               "TenBuuCuc":  "Bưu cục TDP Vàng-Việt Yên-Bắc Giang"
                                           },
                                           {
                                               "TongCOD":  249.33,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Thôn Nam Hòa-An Dương-Hải Phòng"
                                           },
                                           {
                                               "TongCOD":  240.75,
                                               "CODTuNop":  2.56,
                                               "TyLeTuNop":  1.1,
                                               "TenBuuCuc":  "Bưu Cục Thôn Bất Lự-Hoàn Sơn-Tiên Du-Bắc Ninh"
                                           },
                                           {
                                               "TongCOD":  168.71,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Nam Sơn 2-Cẩm Phả-Quảng Ninh"
                                           },
                                           {
                                               "TongCOD":  362.07,
                                               "CODTuNop":  96.26,
                                               "TyLeTuNop":  26.6,
                                               "TenBuuCuc":  "Bưu Cục Sở Dầu-Hồng Bàng-Hải Phòng"
                                           },
                                           {
                                               "TongCOD":  206.45,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 23 Chu Văn An-Thanh Miện-Hải Dương"
                                           },
                                           {
                                               "TongCOD":  186.55,
                                               "CODTuNop":  26.83,
                                               "TyLeTuNop":  14.4,
                                               "TenBuuCuc":  "Bưu Cục KĐT Thị Trấn Chờ-Yên Phong-Bắc Ninh"
                                           },
                                           {
                                               "TongCOD":  65.2,
                                               "CODTuNop":  0.85,
                                               "TyLeTuNop":  1.3,
                                               "TenBuuCuc":  "Bưu Cục TL 293 Xóm Rén-Xã Trường Sơn-Bắc Giang"
                                           },
                                           {
                                               "TongCOD":  204.64,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 1056 Nguyễn Bỉnh Khiêm-Q.Hải An-Hải Phòng"
                                           },
                                           {
                                               "TongCOD":  145.34,
                                               "CODTuNop":  26.05,
                                               "TyLeTuNop":  17.9,
                                               "TenBuuCuc":  "Bưu Cục Đường Tỉnh 388-Thạch Lỗi-Cẩm Giàng-Hải Dương"
                                           },
                                           {
                                               "TongCOD":  185.13,
                                               "CODTuNop":  41.02,
                                               "TyLeTuNop":  22.2,
                                               "TenBuuCuc":  "Bưu Cục Thị Trấn Vôi-Lạng Giang-Bắc Giang"
                                           },
                                           {
                                               "TongCOD":  133.52,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Đảo Vân Đồn-Quảng Ninh"
                                           },
                                           {
                                               "TongCOD":  34.21,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu cục QL4B-Đình Lập-Lạng Sơn"
                                           },
                                           {
                                               "TongCOD":  194.38,
                                               "CODTuNop":  131.78,
                                               "TyLeTuNop":  67.8,
                                               "TenBuuCuc":  "Bưu Cục 120 An Định-Tp.Hải Dương-Hải Dương"
                                           },
                                           {
                                               "TongCOD":  111.47,
                                               "CODTuNop":  50.57,
                                               "TyLeTuNop":  45.4,
                                               "TenBuuCuc":  "Bưu Cục 54 Lưu Úc-Phù Liễn-Kiến An-Hải Phòng"
                                           },
                                           {
                                               "TongCOD":  219.66,
                                               "CODTuNop":  49.53,
                                               "TyLeTuNop":  22.5,
                                               "TenBuuCuc":  "Bưu Cục 16A An Toàn-Hoà Nghĩa-Dương Kinh-Hải Phòng"
                                           },
                                           {
                                               "TongCOD":  244.48,
                                               "CODTuNop":  81.19,
                                               "TyLeTuNop":  33.2,
                                               "TenBuuCuc":  "Bưu Cục Khu Sơn Đông-Phường Nam Sơn-Bắc Ninh"
                                           },
                                           {
                                               "TongCOD":  190.67,
                                               "CODTuNop":  153.5,
                                               "TyLeTuNop":  80.5,
                                               "TenBuuCuc":  "Bưu Cục 610 Nguyễn Bình-Đông Triều-Quảng Ninh"
                                           },
                                           {
                                               "TongCOD":  330.15,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Gia Đông-Thuận Thành-Bắc Ninh"
                                           },
                                           {
                                               "TongCOD":  264.21,
                                               "CODTuNop":  146.7,
                                               "TyLeTuNop":  55.5,
                                               "TenBuuCuc":  "Bưu cục Tổ 12-Kênh Dương-Lê Chân-Hải Phòng"
                                           },
                                           {
                                               "TongCOD":  70.08,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 144 Trần Phú-Xã Đầm Hà-Quảng Ninh"
                                           },
                                           {
                                               "TongCOD":  268.73,
                                               "CODTuNop":  268.73,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 61 đường Trần Đăng Tuyển-Vũ Ninh-Bắc Ninh"
                                           },
                                           {
                                               "TongCOD":  347.75,
                                               "CODTuNop":  0.04,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 827 Trần Hưng Đạo-Kinh Môn-Hải Dương"
                                           },
                                           {
                                               "TongCOD":  74.94,
                                               "CODTuNop":  4.75,
                                               "TyLeTuNop":  6.3,
                                               "TenBuuCuc":  "Bưu Cục 12 Nguyễn Thị Minh Khai-Bắc Sơn-Lạng Sơn"
                                           },
                                           {
                                               "TongCOD":  141.04,
                                               "CODTuNop":  111.68,
                                               "TyLeTuNop":  79.2,
                                               "TenBuuCuc":  "Bưu Cục Đường Tỉnh 356 Hùng Sơn-Cát Hải-Hải Phòng"
                                           },
                                           {
                                               "TongCOD":  215.75,
                                               "CODTuNop":  160.65,
                                               "TyLeTuNop":  74.5,
                                               "TenBuuCuc":  "Bưu Cục Vĩnh Hải-Nguyễn Đức Cảnh-Đông Triều-Quảng Ninh"
                                           },
                                           {
                                               "TongCOD":  140.18,
                                               "CODTuNop":  8.15,
                                               "TyLeTuNop":  5.8,
                                               "TenBuuCuc":  "Bưu Cục Thị trấn Nhã Nam-Tân Yên-Bắc Giang"
                                           },
                                           {
                                               "TongCOD":  106.08,
                                               "CODTuNop":  68.94,
                                               "TyLeTuNop":  65,
                                               "TenBuuCuc":  "Bưu Cục Thôn Lán Than-Xã Sơn Động-Bắc Giang"
                                           },
                                           {
                                               "TongCOD":  90.01,
                                               "CODTuNop":  39.62,
                                               "TyLeTuNop":  44,
                                               "TenBuuCuc":  "Bưu Cục Xã Hương Lâm-Hiệp Hòa-Bắc Giang"
                                           },
                                           {
                                               "TongCOD":  139.26,
                                               "CODTuNop":  71.01,
                                               "TyLeTuNop":  51,
                                               "TenBuuCuc":  "Bưu cục Cửa Ông-Cẩm Phả-Quảng Ninh"
                                           },
                                           {
                                               "TongCOD":  177.11,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Đào Xá-Lương Tài-Bắc Ninh"
                                           },
                                           {
                                               "TongCOD":  561.73,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 4B Hùng Thắng-Hạ Long-Quảng Ninh"
                                           },
                                           {
                                               "TongCOD":  335.62,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Đặng Châu Tuệ-Cao Xanh-Hạ Long-Quảng Ninh"
                                           },
                                           {
                                               "TongCOD":  140.33,
                                               "CODTuNop":  27.47,
                                               "TyLeTuNop":  19.6,
                                               "TenBuuCuc":  "Bưu Cục Thôn Phúc Nghiêm-Xã Phật Tích-Bắc Ninh"
                                           },
                                           {
                                               "TongCOD":  250.78,
                                               "CODTuNop":  175.93,
                                               "TyLeTuNop":  70.2,
                                               "TenBuuCuc":  "Bưu Cục Phố Mới-Thủy Nguyên-Hải Phòng"
                                           },
                                           {
                                               "TongCOD":  58.59,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục TT Ba Chẽ-Xã Ba Chẽ-Quảng Ninh"
                                           },
                                           {
                                               "TongCOD":  153.01,
                                               "CODTuNop":  31.05,
                                               "TyLeTuNop":  20.3,
                                               "TenBuuCuc":  "Bưu Cục TDP Đồng Tiến-Phường Nam Đồ Sơn-Hải Phòng"
                                           },
                                           {
                                               "TongCOD":  183.67,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục QL1A-Hữu Lũng-Lạng Sơn"
                                           },
                                           {
                                               "TongCOD":  141.73,
                                               "CODTuNop":  31.35,
                                               "TyLeTuNop":  22.1,
                                               "TenBuuCuc":  "Bưu Cục 01 Khu Dân Cư-Tứ Minh-Hải Dương"
                                           },
                                               "TyLeTuNop":  57.4,
                                               "TenBuuCuc":  "Bưu Cục Chợ Ấp Đồn-Yên Trung-Yên Phong-Bắc Ninh"
                                           },
                                           {
                                               "TongCOD":  128.74,
                                               "CODTuNop":  83.8,
                                               "TyLeTuNop":  65.1,
                                               "TenBuuCuc":  "Bưu cục Đồng Đăng-Cao Lộc-Lạng Sơn"
                                           },
                                           {
                                               "TongCOD":  310.94,
                                               "CODTuNop":  310.94,
                                               "TyLeTuNop":  100,
                                               "TenBuuCuc":  "Bưu Cục 228 Nguyễn Bình-Quảng Yên-Quảng Ninh"
                                           },
                                           {
                                               "TongCOD":  298.78,
                                               "CODTuNop":  98.73,
                                               "TyLeTuNop":  33,
                                               "TenBuuCuc":  "Bưu Cục Thị Trấn Đồi Ngô-Lục Nam-Bắc Giang"
                                           },
                                           {
                                               "TongCOD":  427.29,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Lý Quốc Sư-Võ Cường-Bắc Ninh"
                                           },
                                           {
                                               "TongCOD":  132.2,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục An Dương 2 - Hải Phòng"
                                           },
                                           {
                                               "TongCOD":  126.26,
                                               "CODTuNop":  63.03,
                                               "TyLeTuNop":  49.9,
                                               "TenBuuCuc":  "Bưu cục Lộc Bình-Lạng Sơn"
                                           },
                                           {
                                               "TongCOD":  191.01,
                                               "CODTuNop":  69.59,
                                               "TyLeTuNop":  36.4,
                                               "TenBuuCuc":  "Bưu Cục KP Mới Yên Lã-Từ Sơn-Bắc Ninh"
                                           },
                                           {
                                               "TongCOD":  182.76,
                                               "CODTuNop":  31.16,
                                               "TyLeTuNop":  17.1,
                                               "TenBuuCuc":  "Bưu Cục Chợ Vé-Ninh Giang-Hải Dương"
                                           },
                                           {
                                               "TongCOD":  191.32,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 102 Nguyễn Đăng Lành-Nam Sách-Hải Dương"
                                           },
                                           {
                                               "TongCOD":  63.51,
                                               "CODTuNop":  21.03,
                                               "TyLeTuNop":  33.1,
                                               "TenBuuCuc":  "Bưu Cục TDP Đồng Lều-Yên Thế-Bắc Giang"
                                           },
                                           {
                                               "TongCOD":  180.47,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 281 Nguyễn Trãi-Hạ Long-Quảng Ninh"
                                           },
                                           {
                                               "TongCOD":  65.2,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Khu Bình Quân-Xã Bình Liêu-Quảng Ninh"
                                           },
                                           {
                                               "TongCOD":  124.98,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Thôn Thác Bưởi-Xã Tiên Yên-Quảng Ninh"
                                           },
                                           {
                                               "TongCOD":  229.48,
                                               "CODTuNop":  109.68,
                                               "TyLeTuNop":  47.8,
                                               "TenBuuCuc":  "Bưu Cục KĐT Nam Hồng-Từ Sơn-Bắc Ninh"
                                           },
                                           {
                                               "TongCOD":  126.68,
                                               "CODTuNop":  61.73,
                                               "TyLeTuNop":  48.7,
                                               "TenBuuCuc":  "Bưu Cục 130 Trần Văn Lan-Q.Hải An-Hải Phòng"
                                           },
                                           {
                                               "TongCOD":  177.36,
                                               "CODTuNop":  82.72,
                                               "TyLeTuNop":  46.6,
                                               "TenBuuCuc":  "Bưu Cục 320 Đường Hùng Vương-Mai Pha-Lạng Sơn"
                                           },
                                           {
                                               "TongCOD":  190.66,
                                               "CODTuNop":  85.23,
                                               "TyLeTuNop":  44.7,
                                               "TenBuuCuc":  "Bưu Cục Thị Trấn Kép-Lạng Giang-Bắc Giang"
                                           },
                                           {
                                               "TongCOD":  106.85,
                                               "CODTuNop":  40.15,
                                               "TyLeTuNop":  37.6,
                                               "TenBuuCuc":  "Bưu Cục Đường Hoàng Văn Thụ-Tràng Định-Lạng Sơn"
                                           },
                                           {
                                               "TongCOD":  282.34,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục 108 Tôn Đức Thắng-Chu Văn An-Hải Dương"
                                           },
                                           {
                                               "TongCOD":  65.87,
                                               "CODTuNop":  51.4,
                                               "TyLeTuNop":  78,
                                               "TenBuuCuc":  "Bưu Cục Lũng Cùng-Hoàng Việt-Văn Lãng-Lạng Sơn"
                                           },
                                           {
                                               "TongCOD":  209.38,
                                               "CODTuNop":  92.04,
                                               "TyLeTuNop":  44,
                                               "TenBuuCuc":  "Bưu Cục Nguyễn Phi Khanh-Tam Thanh-Lạng Sơn"
                                           },
                                           {
                                               "TongCOD":  238.02,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Cột 5-Hạ Long-Quảng Ninh"
                                           },
                                           {
                                               "TongCOD":  445.99,
                                               "CODTuNop":  50.47,
                                               "TyLeTuNop":  11.3,
                                               "TenBuuCuc":  "Bưu Cục 29 Ngô Gia Tự-Móng Cái-Quảng Ninh 01"
                                           },
                                           {
                                               "TongCOD":  268.52,
                                               "CODTuNop":  84.4,
                                               "TyLeTuNop":  31.4,
                                               "TenBuuCuc":  "Bưu Cục Tiên Lãng-Hải Phòng"
                                           },
                                           {
                                               "TongCOD":  230.13,
                                               "CODTuNop":  0,
                                               "TyLeTuNop":  0,
                                               "TenBuuCuc":  "Bưu Cục Thị Trấn Tứ Kỳ-Tứ Kỳ-Hải Dương"
                                           },
                                           {
                                               "TongCOD":  241.15,
                                               "CODTuNop":  161.45,
                                               "TyLeTuNop":  66.9,
                                               "TenBuuCuc":  "Bưu Cục 88 Thống Nhất-Kim Thành-Hải Dương"
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

// Convert parsed Excel rows into structured appData
function processImportedJSON(rows) {
    if (!rows || rows.length === 0) return;

    let keys = Object.keys(rows[0]);
    let colPayment = keys.find(k => k.toLowerCase().replace(/_/g, '').includes('hinhthucthanhtoan') || k.toLowerCase().includes('hinh_thuc') || k.toLowerCase().includes('hÃ¬nh thá»©c'));
    let colDate = keys.find(k => k.toLowerCase().replace(/_/g, '').includes('ngayhethong') || k.toLowerCase().includes('ngay_he_thong') || k.toLowerCase().includes('ngÃ y') || k.toLowerCase().includes('ngay'));
    let colOffice = keys.find(k => k.toLowerCase().includes('chi_nhanh') || k.toLowerCase().includes('buucuc') || k.toLowerCase().includes('bÆ°u cá»¥c'));
    let colMoney = keys.find(k => k.toLowerCase().replace(/_/g, '').includes('tongsotiennop') || k.toLowerCase().includes('tongcod') || k.toLowerCase().includes('tá»•ng cod') || k.toLowerCase().includes('tien_nop'));
    let colRegion = keys.find(k => k.toLowerCase().includes('vung') || k.toLowerCase().includes('vng') || k.toLowerCase().includes('vÃ¹ng'));

    // Fallbacks if not found
    colPayment = colPayment || keys[0];
    colDate = colDate || keys[1];
    colOffice = colOffice || keys[2];
    colMoney = colMoney || keys[5];
    colRegion = colRegion || keys[6];

    // 1. Gather all unique dates and regions
    let uniqueDates = new Set();
    let uniqueRegions = new Set();
    
    rows.forEach(row => {
        let dateVal = row[colDate];
        if (dateVal) {
            let dateStr = formatDateStr(dateVal);
            uniqueDates.add(dateStr);
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
    let aggMap = {}; // Key: Date|Region -> { total: X, ck: Y }
    let officeAggMap = {}; // Key: Region|Office -> { total: X, ck: Y } - for the latest date only

    rows.forEach(row => {
        let dateVal = row[colDate];
        let regionVal = row[colRegion];
        let paymentVal = row[colPayment];
        let moneyVal = row[colMoney];
        let officeVal = row[colOffice];

        if (!dateVal || !regionVal) return;

        let dateStr = formatDateStr(dateVal);
        let regionStr = regionVal.toString().trim();
        let money = parseFloat(moneyVal) || 0;
        
        let isCK = false;
        if (paymentVal) {
            let payStr = paymentVal.toString().trim().toUpperCase();
            if (payStr === 'CK' || payStr.includes('CHUYá»‚N KHOáº¢N') || payStr.includes('CHUYEN KHOAN')) {
                isCK = true;
            }
        }
        let key = dateStr + '|' + regionStr;
        if (!aggMap[key]) {
            aggMap[key] = { total: 0, ck: 0 };
        }
        aggMap[key].total += money;
        if (isCK) {
            aggMap[key].ck += money;
        }

        if (dateStr === latestDate && officeVal) {
            let officeStr = officeVal.toString().trim();
            if (/^\d+-/.test(officeStr)) {
                officeStr = officeStr.replace(/^\d+-/, '');
            }
            let offKey = regionStr + '|' + officeStr;
            if (!officeAggMap[offKey]) {
                officeAggMap[offKey] = { total: 0, ck: 0 };
            }
            officeAggMap[offKey].total += money;
            if (isCK) {
                officeAggMap[offKey].ck += money;
            }
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
    if (val instanceof Date) {
        let d = val.getDate();
        let m = val.getMonth() + 1;
        let y = val.getFullYear();
        return (d < 10 ? '0' + d : d) + '/' + (m < 10 ? '0' + m : m) + '/' + y;
    }
    let str = val.toString().trim();
    if (str.includes('-')) {
        let parts = str.split(' ')[0].split('-');
        if (parts[0].length === 4) { // yyyy-MM-dd
            return parts[2] + '/' + parts[1] + '/' + parts[0];
        }
    }
    if (str.includes('/')) {
        let parts = str.split(' ')[0].split('/');
        if (parts[2] && parts[2].length === 4) {
            let d = parseInt(parts[1]) || 1;
            let m = parseInt(parts[0]) || 1;
            let y = parts[2];
            if (m > 12) {
                let temp = d;
                d = m;
                m = temp;
            }
            return (d < 10 ? '0' + d : d) + '/' + (m < 10 ? '0' + m : m) + '/' + y;
        }
    }
    return str;
}

function parseDateStr(str) {
    let parts = str.split('/');
    return new Date(parts[2], parts[1] - 1, parts[0]);
}

// Generate and trigger download of a sample CSV/Excel layout
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