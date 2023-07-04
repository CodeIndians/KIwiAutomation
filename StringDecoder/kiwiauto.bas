Sub chromeAuto()
    
    Dim obj
    Set obj = CreateObject("Selenium.WebDriver")
    
    obj.Start "chrome", ""
    obj.Get "https://kiwi.cygni-systems.com/"
       
    ' Login Page
    obj.FindElementByName("USERNAME").SendKeys ("Bruce-Canary")
    obj.FindElementByName("PASSWORD").SendKeys ("Welcome")
    obj.FindElementByClass("mt-2").Click
    
    ' Add new cultlist
    obj.FindElementByClass("dropdown-toggle").Click
    obj.FindElementByCss(".new-cutlist-link.dropdown-item").Click
    
    ' add a delay to load the elements
    Delay 5000
    
    ' Hardcoding light gauge descriptions
    Dim structuralArray
    structuralArray = Array("#1 BLANK LINE (STRUCTURAL)", "1"" X 4"" X 3 X 14GA Concrete Stop", "1"" X 4-1/2"" X 3 X 14GA Concrete Stop", "1"" x 47"" FLAT PANEL x 22GA", "1"" x 47"" FLAT PANEL x 26GA", "1"" X 5"" X 3"" X 14GA Concrete Stop", "1"" X 5-1/2"" X 2-1/2"" X 14GA Concrete Stop", "1-5/8 X 1-1/2 X 16GA CEE", "1-5/8 X 1-1/2 X 18GA CEE HW3110", "1-5/8 X 1-1/2 X 20GA CEE", "10 X 2-1/2 X 12GA CEE", "10 X 2-1/2 X 12GA ZEE", "10 X 2-1/2 X 14GA CEE", "10 X 2-1/2 X 14GA ZEE", "10 X 2-1/2 X 16GA CEE" _
, "10 X 2-1/2 X 16GA ZEE", "10 X 3-1/2 X 12GA CEE", "10 X 3-1/2 X 12GA ZEE", "10 X 3-1/2 X 14GA CEE", "10 X 3-1/2 X 14GA ZEE", "12 X 2-1/2 X 12GA CEE", "12 X 2-1/2 X 12GA ZEE", "12 X 2-1/2 X 14GA CEE", "12 X 2-1/2 X 14GA ZEE", "12 X 3 X 12GA CEE", "12 X 3-1/2 X 12GA CEE", "12 X 3-1/2 X 12GA ZEE", "14"" FLAT STOCK X 12GA", "3"" Flat Stock x 12ga", "3"" Flat Stock x 16ga", "3-5/8 X 1-1/2 X 16GA CEE", "4 X 1-1/2 X 16GA CEE", "4 X 2 X 12GA CEE", "4 X 2 X 14GA CEE", "4 X 2 X 16GA CEE" _
, "4 X 2-1/2 X 12GA CEE", "4 X 2-1/2 X 12GA ZEE", "4 X 2-1/2 X 14GA CEE", "4 X 2-1/2 X 14GA ZEE", "4 X 2-1/2 X 16GA CEE", "4 X 2-1/2 X 16GA ZEE", "4 X 3-1/2 X 12GA CEE", "4 X 3-1/2 X 12GA ZEE", "4 X 3-1/2 X 14GA CEE", "4 X 3-1/2 X 14GA ZEE", "4 X 3-1/2 X 16GA CEE", "4 X 3-1/2 X 16GA ZEE", "4"" Flat Stock X 12GA ", "4"" Flat Stock X 16GA ", "4"" x 8"" x 4"" x 1"" x 16ga. Special Break Ledger", "5"" x 4"" x 4"" x 1"" x 16ga. Special Break Ledger ", "6 X 2 X 12GA CEE", "6 X 2 X 14GA CEE" _
, "6 X 2 X 16GA CEE", "6 X 2-1/2 X 12GA CEE", "6 X 2-1/2 X 12GA ZEE", "6 X 2-1/2 X 14GA CEE", "6 X 2-1/2 X 14GA ZEE", "6 X 2-1/2 X 16GA CEE", "6 X 2-1/2 X 16GA ZEE", "6 X 3 X 12GA CEE", "6 X 3 X 14GA CEE", "6 X 3 X 16GA CEE", "6 X 3-1/2 X 12GA CEE", "6 X 3-1/2 X 12GA ZEE", "6 X 3-1/2 X 14GA CEE", "6 X 3-1/2 X 14GA ZEE", "6 X 3-1/2 X 16GA CEE", "6 X 3-1/2 X 16GA ZEE", "6 X 4 X 16GA CEE", "8 X 2-1/2 X 12GA CEE", "8 X 2-1/2 X 12GA ZEE", "8 X 2-1/2 X 14GA CEE", "8 X 2-1/2 X 14GA ZEE" _
, "8 X 2-1/2 X 16GA CEE", "8 X 2-1/2 X 16GA ZEE", "8 X 3 X 12GA CEE", "8 X 3 X 12GA ZEE", "8 X 3 X 14GA CEE", "8 X 3 X 14GA ZEE", "8 X 3 X 16GA CEE", "8 X 3 X 16GA ZEE", "8 X 3-1/2 X 12GA CEE ", "8 X 3-1/2 X 12GA ZEE ", "9"" FLAT STOCK X 16GA", "HW3120 1"" Sub-Girt 18GA", "HW3130 1/2"" Sub-Girt 18GA", "J 1-1/2 X 1-1/4 X 20GA", "J 1-5/8 X 1-1/2 X 16GA", "J 1-5/8 X 1-1/2 X 7/8 18GA", "L 1-1/2 X 2-1/2 X 14GA ", "L 1-1/2 X 2-1/2 X 16GA ", "L 1-1/2 X 2-1/2 X 18GA ", "L 2 X 2 X 14GA " _
, "L 2 X 2 X 16GA ", "L 2 X 2 X 18GA", "L 2 X 4 X 14GA", "L 2 X 4 X 16GA", "L 2 X 6 X 16GA ", "L 2 X 8 X 18GA", "L 3 X 1 X 16GA ", "L 3 X 3 X 12GA", "L 3 X 3 X 14GA", "L 3 X 3 X 16GA", "L 3 X 3 X 18GA", "L 3 X 4 X 14GA", "L 3 X 8 X 12GA", "L 3 X 8 X 14GA", "L 3 X 8 X 16GA", "L 3-5/8 X 3-5/8 X 16GA", "L 3/4 X 5-1/4 X 16GA", "L 3/4 X 6-1/2 X 16GA", "L 4 X 2 X 16GA", "L 4 X 2-1/2 X 16GA", "L 4 X 4 X 12GA ", "L 4 X 4 X 14GA ", "L 4 X 4 X 16GA ", "L 5 X 5 X 12GA", "L 5 X 5 X 14GA" _
, "L 5 X 5 X 16GA", "L 6 X 4 X 16GA ", "L 8 X 4 X 12GA ", "U 3-5/8 X 1-7/8 X 16GA CHANNEL", "U 4-1/8 X 2 X 12GA CHANNEL", "U 4-1/8 X 2 X 14GA CHANNEL", "U 4-1/8 X 2 X 16GA CHANNEL", "U 4-1/8 X 2-1/2 X 16GA CHANNEL", "U 4-1/8 X 3 X 12GA CHANNEL", "U 4-1/8 X 3 X 14GA CHANNEL", "U 4-1/8 X 3 X 16GA CHANNEL", "U 4-1/8 X 3-1/2 X 16GA CHANNEL", "U 4-1/8 X 4 X 12GA CHANNEL", "U 4-1/8 X 4 X 14GA CHANNEL", "U 4-1/8 X 4 X 16GA CHANNEL", "U 6-1/8 X 2 X 12GA CHANNEL", "U 6-1/8 X 2 X 14GA CHANNEL" _
, "U 6-1/8 X 2 X 16GA CHANNEL", "U 6-1/8 X 2-1/2 X 16GA CHANNEL", "U 6-1/8 X 3 X 12GA CHANNEL", "U 6-1/8 X 3 X 14GA CHANNEL", "U 6-1/8 X 3 X 16GA CHANNEL", "U 6-1/8 X 4 X 12GA CHANNEL", "U 6-1/8 X 4 X 14GA CHANNEL", "U 6-1/8 X 4 X 16GA CHANNEL", "U 8-1/8 X 2 X 16GA CHANNEL", "U 8-1/8 X 3 X 12GA CHANNEL", "U 8-1/8 X 3 X 14GA CHANNEL", "U 8-1/8 X 3 X 16GA CHANNEL")

    Dim compositeArray
    compositeArray = Array("#1 BLANK LINE (COMPOSITE_DECK)", "2.0CF X 18GA", "2W Foam Grommets ", "2WX18GA Comp Deck", "3 X 20GA LOK FLOOR", "3.0CF X 20GA", "3VLI X 20GA", "3W Foam Grommets ", "3WX18GA Comp Deck ", "PLW2 ", "PLW2-36 X 18GA FORMLOK")
    
    Dim roofDeckArray
    roofDeckArray = Array("#1 BLANK LINE (ROOF_DECK)", "3N X 18GA", "3N X 22GA", "DOUBLE-LOK 18"" Wide x 24ga ", "DOUBLE-LOK 24"" Wide x 24ga ", "F313 ROLLFOAM DOWNSPOUT X 26GA", "F320 ROLLFOAM DOWNSPOUT X 26GA", "F322 ROLLFOAM DOWNSPOUT X 26GA", "F797 DOWNSPOUT STRAP X 24GA", "F797 DOWNSPOUT STRAP X 26GA", "FASTENER #1", "FASTENER #12A", "FASTENER #14", "FASTENER #1E", "FASTENER #3", "FASTENER #4", "FASTENER #45", "FASTENER #5", "FL-246 GUTTER STRAP", "FL-290 PARAPET RAKE CLEAT (1""X3/4"")", "FL115 RAKE SLIDE X 24GA" _
, "FL115 RAKE SLIDE X 26GA", "FL27 OUTSIDE ANGLE X 24GA", "FL29 OUTSIDE ANGLE TRIM X 24GA", "HSB36 SS DECK X 16GA", "HSB36 SS DECK X 18GA ", "HSB36 SS DECK x 20GA ", "HSB36 SS DECK x 22GA ", "HW-200 3-3/8"" FIXED CLIP", "HW-204 4-3/8"" FIXED CLIP", "HW-208 3"" UTILITY CLIP", "HW-2102 3-3/8"" SLIDING CLIP", "HW-2104 SLIDING CLIP", "HW-426 METAL INSIDE CLOSURE", "HW-430 24"" OUTSIDE CLOSURE", "HW-432 18"" OUTSIDE CLOSURE", "HW-4601 CLIP (FW-120)", "HW-502 TRIPLE BEAD", "HW-504 TRI-BEAD (TACKY TAPE)" _
, "HW-512 MINOR RIB SEALER", "HW-540", "HW-541 SEALANT ", "HW-7710 3-3/8"" LOW RAKE", "HW-7720 4-3/8"" HIGH RAKE", "HW-7722 3"" UTILITY RAKE ", "HW-7760 24"" BACK-UP PLATE (UD/DL)", "HW-7762 18"" BACK-UP PLATE (LS)", "HW-7764 12"" BACK-UP PLATE (BL/SL/LS)", "HW-7766 16"" BACK-UP PLATE (BL/SL/LS)", "HW-7769 18"" BACK-UP PLATE (UD/DL)", "HW6200 CLIP (LOKSEAM)", "HW7616 EAVE PLATE HIGH X 14GA", "HW7617 EAVE PLATE LOW 2"" FLOATING ", "HW7618 EAVE PLATE HIGH 1"" FLOATING ", "LOKSEAM 12"" WIDE X 24GA" _
, "LOKSEAM 16"" WIDE X 24GA", "LOKSEAM 18"" WIDE X 24GA", "PLN24", "SP22", "SP24 ", "SP26", "SUPERLOK 16"" WIDE X 24GA", "T5014 Z CLOSURE X 24GA", "T5219 PARAPET RAKE CLEAT X 24GA", "T5526 ML SUPPORT ZEE X 24GA", "T5530 ML PANEL CLEAT X 24GA", "TS 324 QUAD-LOK X 22GA", "TS 324 QUAD-LOK X 24GA", "ULTRA-DEK 18"" Wide x 24ga ", "ULTRA-DEK 24"" Wide x 24ga")
    
    Dim partitionPanelArray
    partitionPanelArray = Array("#1 BLANK LINE (PARTITION_PANEL)", "1 X 12 X 26GA FILLER PANEL", "1 X 12 X 29GA FILLER PANEL", "1 X 6-1/4 X 26GA FILLER PANEL", "1 X 6-1/4 X 29GA FILLER PANEL", "1 X 9-1/4 X 22GA FILLER PANEL", "1 X 9-1/4 X 24GA FILLER PANEL", "1 X 9-1/4 X 26GA FILLER PANEL", "1 X 9-1/4 X 29GA FILLER PANEL", "AEP U-Panel FULL Width X 22GA", "AEP U-Panel FULL Width X 24GA", "AEP U-Panel FULL Width X 26GA", "AEP U-Panel FULL Width X 29GA", "HW459 U-PANEL INSIDE CLOSURE ", "HW460 U-PANEL OUTSIDE CLOSURE", "HW7600 EAVE PLATE LOW X 14GA", "U-Panel FULL Width X 22GA", "U-Panel FULL Width x 24GA", "U-Panel FULL Width X 26GA", "U-Panel FULL Width X 29GA", "U-Panel HALF Width X 22GA", "U-Panel HALF Width X 26GA", "U-Panel HALF Width X 29GA")
    
    Dim linearPanelArray
    linearPanelArray = Array("#1 BLANK LINE (LINER_PANEL)", "AEP U-Panel FULL Width X 22GA LP", "AEP U-Panel FULL Width X 24GA LP", "AEP U-Panel FULL Width X 26GA LP", "AEP U-Panel FULL Width X 29GA LP", "U-Panel FULL Width x 22GA LP", "U-Panel FULL Width x 24GA LP", "U-Panel FULL Width X 26GA LP", "U-Panel FULL Width X 29GA LP", "U-Panel HALF Width x 22GA LP", "U-Panel HALF Width x 24GA LP", "U-Panel HALF Width X 26GA LP", "U-Panel HALF Width X 29GA LP")
    
    Dim sidingPanelArray
    sidingPanelArray = Array("#1 BLANK LINE (SIDING_PANEL)", "#10-16 X 1"" PANCAKE HEAD DRILLER", "#12 X 3/4"" PANCAKE HEAD DRILLER ", "1/4-14 X 7/8 LL TAP TEK SEALER ", "1/8 POP RIVET ", "12-14 X 1-1/4 LL DRILLER WW PAINTED ", "14 X 1-3/8"" LL MASONARY FASTENER ", "14 X 1-3/8"" MASONARY FASTENER ", "28.8"" 7.2 PANEL X 26GA", "48"" Flat Panel x 24GA", "48"" FLAT PANEL X 26GA", "7.2 PANEL X 24GA", "7.2 PANEL X 26GA", "APS500 SMOOTH ADVANCED POLYMER SEALANT", "ARTISAN L12 X 24GA", "DESIGNER FLUTED 16"" WIDE X 24GA" _
, "FASTENER #12", "FW-120 X 22GA", "FW-120 X 24GA", "FW-120-1 X 22GA", "FW-120-1 X 24GA", "FW-120-2 X 22GA", "FW-120-2 X 24GA", "HW-433 16"" OUTSIDE CLOSURE FOR MASTERLINE", "HW455 R PANEL INSIDE CLOSURE BARE", "HW461 7.2 PANEL CLOSURE BARE", "HW462 PBC PANEL CLOSURE BARE ", "HW463 PBD PANEL CLOSURE BARE ", "IC72 X 24GA", "MASTERLINE 16 X 24GA", "PBC PANEL X 24GA", "PBC PANEL X 26GA", "PBD PANEL X 24GA", "PBD PANEL X 26GA", "PBR PANEL X 22GA", "PBR PANEL X 24GA", "PBR PANEL X 26GA" _
, "SDS #12-14 x 3/4""", "SHADOWRIB X 24GA", "SP24 TRIM", "SP26 TRIM", "STK2000", "T6023 SR OUTSIDE CORNER X 24GA", "T6131 SR FURRING ZEE X 24GA", "T6132 SR FURRING ZEE X 24GA", "TLC-2 X 24GA", "U-Panel HALF Width x 24GA")
    
    Dim insulationArray
    insulationArray = Array("#1 BLANK LINE (INSULATION)", "1 GALLON OF GLUE ", "1"" Rigid Board (4'x8')", "1-1/2"" Rigid Board (4'x8')", "2"" Rigid Board (4'x8')", "3-1/2"" Stick Pins ", "3/4"" White Banding x 24ga", "5-1/2"" Stick Pins", "6-1/2"" Stick Pins", "Double-Sided Tape 150' Roll", "Patch Tape 150' Roll", "R11 White Vinyl-Backed 2'-6"" Wide", "R11 White Vinyl-Backed 5'-0"" Wide", "R11 WMP-VR-R 4'-10"" Wide w/Tape Tabs", "R11 WMP-VR-R 5'-0"" Wide w/Tape Tabs", "R11 WMP-VR-R 5'-4"" Wide w/Tape Tabs" _
, "R13 White Vinyl-Backed 1'-0"" Wide", "R13 White Vinyl-Backed 1'-10"" Wide", "R13 White Vinyl-Backed 1'-6"" Wide", "R13 White Vinyl-Backed 2'-0"" Wide", "R13 White Vinyl-Backed 2'-6"" Wide", "R13 White Vinyl-Backed 3'-0"" Wide", "R13 White Vinyl-Backed 3'-6"" Wide", "R13 White Vinyl-Backed 4'-0"" Wide", "R13 White Vinyl-Backed 5'-0"" Wide", "R13 WMP-VR-R 4'-10"" Wide w/Tape Tabs", "R13 WMP-VR-R 5'-0"" Wide w/Tape Tabs", "R13 WMP-VR-R 5'-4"" Wide w/Tape Tabs" _
, "R19 White Vinyl-Backed 1'-10"" Wide ", "R19 White Vinyl-Backed 1'-4"" Wide ", "R19 White Vinyl-Backed 2'-0"" Wide ", "R19 White Vinyl-Backed 2'-6"" Wide ", "R19 White Vinyl-Backed 3'-0"" Wide ", "R19 White Vinyl-Backed 5'-0"" Wide ", "R19 WMP-VR-R 4'-10"" Wide w/Tape Tabs", "R19 WMP-VR-R 5'-0"" Wide w/Tape Tabs", "R19 WMP-VR-R 5'-4"" Wide w/Tape Tabs", "R25 White Vinyl-Backed 1'-4"" Wide", "R25 White Vinyl-Backed 2'-0"" Wide", "R25 White Vinyl-Backed 2'-6"" Wide" _
, "R25 White Vinyl-Backed 3'-0"" Wide", "R25 White Vinyl-Backed 4'-0"" Wide", "R25 White Vinyl-Backed 5'-0"" Wide", "R25 WMP-VR-R 4'-10"" Wide w/Tape Tabs", "R25 WMP-VR-R 5'-0"" Wide w/Tape Tabs", "R25 WMP-VR-R 5'-4"" Wide w/Tape Tabs", "R30 WMP-VR-R 1'-0"" Wide w/Tape Tabs", "R30 WMP-VR-R 2'-0"" Wide w/Tape Tabs", "R30 WMP-VR-R 5'-0"" Wide w/Tape Tabs", "R6 WMP-VR-R 6'-0"" Wide Perforated ")
    
    Dim anchorsArray
    anchorsArray = Array("#1 BLANK LINE (ANCHORS)", "1/2"" Dia. X 1-1/2"" M.B. With (2) nuts and (2) Washers ", "1/2"" dia. x 2"" Dewalt Screw Bolt+", "1/2"" Dia. x 3 3/4"" Dewalt Power Stud + SD2", "1/2"" Dia. x 3"" Simpson Titen Anchor Bolt", "1/2"" dia. x 4"" Dewalt Screw Bolt+", "1/2"" Dia. x 4"" Simpson Titen Anchor Bolt", "1/2"" Dia. x 5"" Dewalt Power Stud + SD2", "1/2"" Dia. x 5"" Simpson Titen Anchor Bolt", "1/2"" dia. x 6"" Dewalt Screw Bolt+", "1/2"" Dia. x 6"" Simpson Titen HD ", "1/2"" Dia. X 8"" All Thread" _
, "1/2"" dia. x 8"" Dewalt Screw Bolt+", "1/4"" Dia. x 2-1/4"" Hex Head Concrete Screw Anchor ", "1/4"" Dia. x 2-3/4"" Simpson Titen Anchor Bolt ", "1/4"" dia. x 3"" Dewalt Screw Bolt+", "3/4"" dia. x 4"" Dewalt Screw Bolt+", "3/4"" Dia. x 6"" Dewalt Power Stud + SD2", "3/4"" dia. x 6"" Dewalt Screw Bolt+", "3/4"" dia. x 8"" Dewalt Screw Bolt+", "3/4"" Dia. X 8"" Threaded Rod ", "3/4"" x 7"" Simpson Titen HD ", "3/8"" Dia. x 1-3/4"" Simpson Titen Anchor Bolt", "3/8"" Dia. x 2-1/2"" Simpson Titen Anchor Bolt " _
, "3/8"" Dia. x 3 3/4"" Dewalt Power Stud + SD2", "3/8"" Dia. x 3"" Simpson Titen Anchor Bolt ", "3/8"" Dia. x 4"" Simpson Titen Anchor Bolt ", "3/8"" Dia. x 5"" Simpson Titen Anchor Bolt", "3/8"" Dia. x 6"" Dewalt Power Stud + SD2", "5/8"" Dia. x 3 3/4"" Dewalt Power Stud + SD2", "5/8"" dia. x 3"" Dewalt Screw Bolt+", "5/8"" dia. x 4"" Dewalt Screw Bolt+", "5/8"" Dia. x 4"" Simpson Titen Anchor Bolt ", "5/8"" Dia. x 5"" Dewalt Power Stud + SD2", "5/8"" dia. x 5"" Dewalt Screw Bolt+" _
, "5/8"" Dia. x 5"" Simpson Titen Anchor Bolt ", "5/8"" dia. x 6"" Dewalt Screw Bolt+", "5/8"" Dia. x 6"" Simpson Titen Anchor Bolt ", "5/8"" Dia. X 8"" All Thread ", "5/8"" dia. x 8"" Dewalt Screw Bolt+", "5/8"" Dia. x 8"" Simpson Titen Anchor Bolt ", "5/8"" x 10 All Thread", "5/8"" x 12"" All Thread", "KB1 Stud anchor 1/2"" Dia. x 2 3/4""", "KB1 Stud anchor 1/2"" Dia. x 3 3/4"" LT", "KB1 Stud anchor 1/2"" Dia. x 4 1/2"" LT", "KB1 Stud anchor 1/2"" Dia. x 5 1/2"" LT ", "KB1 Stud anchor 1/2"" Dia. x 7"" LT" _
, "KB1 Stud anchor 1/4"" Dia. x 1 3/4""", "KB1 Stud anchor 1/4"" Dia. x 2 1/4"" ", "KB1 Stud anchor 3/4"" Dia. x 10"" LT ", "KB1 Stud anchor 3/4"" Dia. x 4 3/4"" LT", "KB1 Stud anchor 3/4"" Dia. x 5 1/2"" LT ", "KB1 Stud anchor 3/4"" Dia. x 7"" LT", "KB1 Stud anchor 3/4"" Dia. x 8"" LT ", "KB1 Stud anchor 3/8"" Dia. x 3-3/4"" LT", "KB1 Stud anchor 3/8"" Dia. x 5"" LT", "KB1 Stud anchor 5/8"" Dia. x 3 3/4""", "KB1 Stud anchor 5/8"" Dia. x 4 3/4"" LT ", "KB1 Stud anchor 5/8"" Dia. x 6"" LT" _
, "KB1 Stud anchor 5/8"" Dia. x 7"" LT ", "KB1 Stud anchor 5/8"" Dia. x 8 1/2"" LT", "KBTZ-2 Stud anchor 1/2"" Dia. x 3 3/4""", "KBTZ-2 Stud anchor 1/2"" Dia. x 4 1/2"" ", "KBTZ-2 Stud anchor 1/2"" Dia. x 5 1/2""", "KBTZ-2 Stud anchor 1/2"" Dia. x 7""", "KBTZ-2 Stud anchor 3/4"" Dia. x 5 1/2"" ", "KBTZ-2 Stud anchor 3/4"" Dia. x 7""", "KBTZ-2 Stud anchor 3/4"" Dia. x 8"" ", "KBTZ-2 Stud anchor 3/8"" Dia. x 3 3/4"" ", "KBTZ-2 Stud anchor 3/8"" Dia. x 3""", "KBTZ-2 Stud anchor 5/8"" Dia. x 4 3/4""" _
, "KBTZ-2 Stud anchor 5/8"" Dia. x 6"" ", "KBTZ-2 Stud anchor 5/8"" Dia. x 8 1/2""", "KH-EZ 5/8"" Dia. x 6"" ", "KH-EZ Screw Anchor 1/2"" Dia. x 2 1/2"" ", "KH-EZ Screw Anchor 1/2"" Dia. x 3 1/2"" ", "KH-EZ Screw Anchor 1/2"" Dia. x 3""", "KH-EZ Screw Anchor 1/2"" Dia. x 4 1/2"" ", "KH-EZ Screw Anchor 1/2"" Dia. x 4"" ", "KH-EZ Screw Anchor 1/2"" Dia. x 5""", "KH-EZ Screw Anchor 1/2"" Dia. x 6""", "KH-EZ Screw Anchor 3/8"" Dia. x 2 1/2"" ", "KH-EZ Screw Anchor 5/8"" Dia. x 5 1/2"" " _
, "Metal anchor ZAMAC 1/4"" Dia. x 1 1/4"" ", "Setting tool HS-SC 150 ")
    
    Dim fastenersArray
    fastenersArray = Array("#1 BLANK LINE (FASTENERS)", "1 USS Flat Washer Zinc ", "1-8 Hex Nut Zinc ", "1/2"" Dia. A307 x 2"" w/ (2) nuts and washers", "1/2"" Flat Washer F436 ", "1/2-13 Hex Nut Zinc ", "1/2-13 x 1 1/2 Hex Bolt ", "1/2-13 x 1-1/2 Hex Bolt A325 ", "1/2-13 x 2 Hex Bolt A325 ", "1/2-13 x 2-1/2 Hex Bolt A325 ", "1/4 x 1-1/4 Hex Head Concrete Screw Tapcon ", "1/4 x 1-1/4 Nail in Anchor Zammet", "1/4 x 1-3/4 Hex Head Concrete Screw Tapcon ", "1/4-14 x 1 1/4 Shoulder Tek 2 (5)", "1/4-14 x 4 Hex Washer Head Tek Zinc" _
, "1/4-14 x 6 Hex Washer Head Zinc", "1/4-14 x 7/8 Hex Washer Head Stitch Tek - No Washer ", "1/4-14 x 8 Hex Washer Head Tek Zinc ", "1/4-14X1 Sealer", "1/4-14X1-1/4 LL Driller WW Painted", "1/4-14X1-1/4 LL Sealer", "1/4-14X1-1/4 Sealer", "1/4-14X7/8 LAP TEK Sealer (4A)", "1/4-14X7/8 LL TAP TEK Sealer (4)", "1/8 Pop Rivet (14)", "10-16 x 1 Hex Washer Head Tek Zinc", "10-16 x 3/4 Buildex Hex Washer Head Tek ", "10-16 x 3/4 Hex Washer Head Tek Zinc ", "12 Neo Sealing Washer " _
, "12 x 1 Pancake Combo Phil/Square Drive Tek Coated ", "12 x 1"" Pancake Head (12A)", "12 x 2 1/4 Hextras", "12-14 X 1 Hex Washer Head Tek ", "12-14 X 1 LL Sealer (27)", "12-14 x 1-1/2 Buildex Hex Washer Head Tek ", "12-14 X 1-1/2 Hex Washer Head Tek ", "12-14 x 1-1/4 Buildex Hex Washer Head Tek 5 ", "12-14 x 1-1/4 Elco Hex Washer Tek 5 w/sealer ", "12-14 x 1-1/4 Elco HWH Tek 5 - ECC720", "12-14 x 1-1/4 LL Sealer (3)", "12-14 X 2 Hex Washer Head Tek ", "12-14 x 2 Hex Washer Head Tek Zinc w/sealer " _
, "12-14 x 2-1/2 Buildex Hex Washer Head Tek ", "12-14 X 2-1/2 Hex Washer Head Tek ", "12-14 x 3/4 Buildex Hex Washer Head Tek", "12-14 x 3/4 Pancake Head Phil Tek Zinc ", "12-14X1-1/4 Driller WW Painted Sealer (1D)", "12-14X1-1/4 LL Driller WW Painted (1E)", "12-14X2 LL Driller WW Painted (1)", "12-14X2-1/2 LL Driller WW Painted", "12X3 LL Driller WW Painted", "14 x 7/8 Hex Washer Head Stitch Tek w/Seal and washer ", "14X1-3/8 LL Masonary Sealer (45)", "17-14 x 1"" LL AB Sealer (2A)" _
, "3 x 3 x 1/4 Square Washer 3/4 ", "3 x 3 x 1/4 Square Washer 5/8 ", "3 x 3 x 1/4 Thick 1/2"" Hole Zinc ", "3/4 USS Flat Washer Zinc", "3/4-10 Hex Nut Zinc ", "3/4-10 x 1 1/2 Hex Bolt Zinc", "5/8 USS Flat Washer Zinc ", "5/8-11 Hex Nut Zinc", "5/8-11 x 1-1/2 Hex Tap Bolt Zinc", "5/8-11 x 1-3/4 Coupling Nut Zinc", "5/8-11 x 10 Hex Bolt Zinc ", "7/8 Galv USS Flat Washer ", "7/8 USS Flat Washer", "7/8-9 Hex Nut Zinc ", "7/8-9 x 1-1/2 Hex Tap Bolt Zinc", "7/8-9 x 2 Hex Bolt Zinc " _
, "A307 1/2"" Dia. X 12"" w/ (2) nuts and washers ", "ATS 5/8 x 13"" Zinc ", "ATS 7/8 x 13"" Zinc ", "Driver bit S-B PH2 50/2"" S (25) ", "DX Cartridge 6 8/11 M red ", "DX Cartridge 6 8/11 M red (sold in BOX of 1000)", "DX Cartridge 6 8/11 M yellow", "DX Cartridge 6 8/11 M yellow (sold in BOX of 1000)", "KIT X-C 20 MX+6.8/11 M BULK Yellow", "KIT X-C 27 MX+6.8/11 M BULK Yellow ", "KIT X-ENP19 MXR+6.8/18 M C-T RED", "KIT X-ENP19+6.8/18 M40 CT (2K) RED ", "KIT X-HSN 24+6.8/11 M40 (2K) RED " _
, "KIT X-HSN 24+6.8/11M RED ", "MOCAP .187-1/2 Gray - Fits #12 &amp; #14", "Nut Setter S-NS 5/16"" M 50/2 ", "SDS  #12-14 x 3/4""", "SDS #1/4-14 x 1 1/2""", "SDS #10-16 x 3/4"" ", "SDS #12-14 x 1 1/2"" ", "SDS #12-14 x 1 1/4"" Tek 5 ", "SDS #12-14 x 1 3/4"" ", "SDS #12-14 x 1"" ", "SDS #12-14 x 2""", "SDS #12-24 x 3"" HWH5 KC ", "Strong-Drive® PPSD SHEATHING-TO-CFS Screw", "Universal nail X-U 19 MX ", "Universal nail X-U 27 MX ", "X-EDN19 THQ12 SHOT/PIN RED")
    
    Dim miscArray
    miscArray = Array("#1 BLANK LINE (MISC)", "#1 Decktite ", "#1 Retro-Fit Decktite ", "#2 Decktite", "#2 Retro-fit Decktite 6"" PIPE", "#3 Decktite ", "#3 Retro-fit Pipe Flashing - SQUARE BASE", "1/2-13 x 7"" All Threaded Stud Zinc", "1/2-13 x 8"" All Threaded Stud Zinc", "12"" Fix-A-Flash Decktite", "2OZ Touch Up Paint", "3/4-10 x 10"" All Threaded Rod Zinc", "3/4-10 x 12"" All Threaded Rod Zinc", "3/4-10 x 14"" All Threaded Rod Zinc", "5/8-11 x 10"" All Threaded Stud Zinc", "5/8-11 x 12"" All Threaded Stud Zinc" _
, "5/8-11 x 13"" All Threaded Stud Zinc", "5/8-11 x 6"" All Threaded Stud Zinc", "5/8-11 x 8"" All Threaded Stud Zinc", "7/8-9 x 12"" All Threaded Rod Zinc", "7/8-9 x 13"" All Threaded Rod Zinc", "7/8-9 x 15"" All Threaded Rod Zinc", "7/8-9 x 8"" All Threaded Rod Zinc", "Anchor rod HAS-E-55 1"" Dia. x 12""", "Anchor rod HAS-E-55 1/2"" Dia. x 10""", "Anchor rod HAS-E-55 1/2"" Dia. x 4 1/2""", "Anchor rod HAS-E-55 1/2"" Dia. x 6 1/2""", "Anchor rod HAS-E-55 1/2"" Dia. x 8""" _
, "Anchor rod HAS-E-55 3/4"" Dia. x 12"" ", "Anchor rod HAS-E-55 3/4"" Dia. x 8""", "Anchor rod HAS-E-55 5/8"" Dia. x 12""", "Anchor rod HAS-E-55 5/8"" Dia. x 8""", "Anchor rod HAS-E-55 5/8"" Dia. x 9""", "Anchor rod HAS-E-55 7/8"" Dia. x 13"" ", "Anchor rod HAS-V-36 1"" Dia. x 12""", "Anchor rod HAS-V-36 1/2"" Dia. x 4 1/2"" ", "Anchor rod HAS-V-36 1/2"" Dia. x 8"" ", "Backup Plate UD/DL 18"" HW7769", "Backup Plate UD/DL 24"" HW7760", "Cartridge holder HIT-CB 500", "Cartridge holder HIT-CR 500 " _
, "Chuck TE-C/TE 30 ", "Closure UD 18"" HW432", "Closure UD 24"" HW430", "Closure UD HW426", "Epoxy Nozzle With Nut Simpson", "Hammer drill bit TE-CX 1/2""-6"" MP32", "Hammer drill bit TE-CX 26/27 1""-10"" ", "Hammer drill bit TE-CX 26/48 1""-18""", "Hammer drill bit TE-YX 5/8""-14"" ", "Handle HIT-RBH ", "HDM 500 Manual Dispenser with HIT-CB ", "HDM 500 Manual Dispenser with HIT-CR ", "Hexagon nut 5/8"" zinced ", "HIT-HY 270", "HTS30 ", "HY200 Epoxy", "Injectable Mortar HIT-HY 200-R 300/1/WH " _
, "Injectable Mortar HIT-HY 270 330/1 ", "Injectable Mortar HIT-RE 500 V3/330/1 ", "Marino-Ware RCC 358", "Marino-Ware RCC 600", "Mesh sleeve HIT-SC 18x85", "Powder-actuated tool DX 460 MX 72", "Powder-actuated tool DX 5 F8", "Powder-actuated tool DX 5 MX ", "RE 500 V3 11.1oz/330ml MC ", "RE500 EPOXY", "Retro-Fit Decktite 2-7 1/4", "Retro-Fit Decktite 3/4-2 3/4", "Roof Clip UD 3 3/8"" LOW-Fixed HW 200", "Roof Clip UD 3"" Utility HW208", "Roof Clip UD 4 3/8"" HIGH-Fixed HW204" _
, "Roof UD Rake Support 3"" HW 7722", "Roof UD Rake Support 4 3/8"" HIGH", "Roof UD Rake Support LOW HW 7710", "Round steel brush HIT-RB 1"" ", "Round steel brush HIT-RB 3/4""", "Round steel brush HIT-RB 5/8"" ", "Round steel brush HIT-RB 7/8""", "Screw anchor Kwik-Con II+ 1/4"" Dia x 1 1/4"" THH", "Simpson #FCB49.5 R-25 ", "Simpson A21 Hanger ", "Simpson CMSTC16 ", "Simpson D/HD15B Holdown", "Simpson Epoxy Cartridge ATXP-30-AC-RILIC CURE 30oz. ", "Simpson Epoxy Cartridge Set 22 22oz" _
, "Simpson Epoxy Cartridge Set 22-XP", "Simpson HDU5 Holddown", "Simpson HDU6 Holddown ", "Simpson Heavy Twist Strap 16"" ", "Simpson LSTA09 Strap", "Simpson LSTA18 Strap", "Simpson LSTA30 Strap ", "Simpson MST126 Strap ", "Simpson MST37 Strap ", "Simpson MSTC28 Strap", "Simpson MSTC40 Strap", "Simpson S/DTT2Z", "Simpson S/H2.5 Clips ", "Simpson S/HD105 Holddown ", "Simpson S/HD10B Holdown ", "Simpson S/HD8S Holdown", "Simpson S/HDU04 Holdown ", "Simpson S/HDU11 Holdown ", "Simpson S/HDU9 Holdown " _
, "Simpson S/HJCT Hanger ", "Simpson S/HJCT Kit", "Simpson S/HTT14", "Simpson S/HTT5", "Simpson S/LBV 2.06 x H = 8 ", "Simpson SHTT-4 Holddown", "Simpson SHTT14 Holdown", "Simpson SLTT-20 Holdown ", "Simpson ST6236 Strap", "Simpson STHD-14 Holdown ", "Tape Seal MINOR RIB 144", "Tape Seal Tri-Bead 6", "Tape Seal Tri-Bead 8", "TE 30 Performance Package fixed ", "TE 30-C-AVR Performance Package ", "TE 30-C-AVR Trade Pack ", "Tube Sealant Urethane Almond", "Tube Sealant Urethane Bronze" _
, "Tube Sealant Urethane Gray", "Tube Sealant Urethane White", "UA-143-12 CLIP", "UD Gutter Strap FL426", "VertiClip SL600", "Vulkem Vu116 Sealant Gray ", "Washer 3/4"" zinced ", "Washer 5/8"" zinced ")
    
    
    Dim currenWorkSheet As Worksheet
    
    ' call light gauge
    Call GetWorksheet("LIGHT GAUGE", currenWorkSheet)
    Call FillData(currenWorkSheet, obj, structuralArray, "STRUCTURALdatalist", "colorSTRUCTURAL", "noteSTRUCTURAL", "unitSTRUCTURAL", "ftSTRUCTURAL", "inchSTRUCTURAL", "fractionDropSTRUCTURAL", "pmSTRUCTURAL", 1)
    
    ' call composite deck
    Call GetWorksheet("COMPOSITE DECK", currenWorkSheet)
    Call FillData(currenWorkSheet, obj, compositeArray, "COMPOSITE_DECKdatalist", "colorCOMPOSITE_DECK", "noteCOMPOSITE_DECK", "unitCOMPOSITE_DECK", "ftCOMPOSITE_DECK", "inchCOMPOSITE_DECK", "fractionDropCOMPOSITE_DECK", "pmCOMPOSITE_DECK", 2)
    
    ' call roof deck
    Call GetWorksheet("ROOF DECK", currenWorkSheet)
    Call FillData(currenWorkSheet, obj, roofDeckArray, "ROOF_DECKdatalist", "", "noteROOF_DECK", "unitROOF_DECK", "ftROOF_DECK", "inchROOF_DECK", "fractionDropROOF_DECK", "pmROOF_DECK", 3)
    
    ' call partition panel
    Call GetWorksheet("PARTITION PANEL", currenWorkSheet)
    Call FillData(currenWorkSheet, obj, partitionPanelArray, "PARTITION_PANELdatalist", "colorPARTITION_PANEL", "notePARTITION_PANEL", "unitPARTITION_PANEL", "ftPARTITION_PANEL", "inchPARTITION_PANEL", "fractionDropPARTITION_PANEL", "pmPARTITION_PANEL", 4)
    
    ' call liner panel
    Call GetWorksheet("LINER PANEL", currenWorkSheet)
    Call FillData(currenWorkSheet, obj, linearPanelArray, "LINER_PANELdatalist", "colorLINER_PANEL", "noteLINER_PANEL", "unitLINER_PANEL", "ftLINER_PANEL", "inchLINER_PANEL", "fractionDropLINER_PANEL", "pmLINER_PANEL", 5)
       
    ' call siding panel
    Call GetWorksheet("SIDING PANEL", currenWorkSheet)
    Call FillData(currenWorkSheet, obj, sidingPanelArray, "SIDING_PANELdatalist", "colorSIDING_PANEL", "noteSIDING_PANEL", "unitSIDING_PANEL", "ftSIDING_PANEL", "inchSIDING_PANEL", "", "pmSIDING_PANEL", 6)
    
    ' call insulation
    Call GetWorksheet("INSULATION", currenWorkSheet)
    Call FillData(currenWorkSheet, obj, insulationArray, "INSULATIONdatalist", "colorINSULATION", "noteINSULATION", "unitINSULATION", "ftINSULATION", "inchINSULATION", "fractionDropINSULATION", "pmINSULATION", 7)
    
    ' call anchors
    Call GetWorksheet("ANCHORS", currenWorkSheet)
    Call FillData(currenWorkSheet, obj, anchorsArray, "ANCHORSdatalist", "colorANCHORS", "noteANCHORS", "unitANCHORS", "ftANCHORS", "inchANCHORS", "fractionDropANCHORS", "pmANCHORS", 8)
    
    ' call fasteners
    Call GetWorksheet("FASTENERS", currenWorkSheet)
    Call FillData(currenWorkSheet, obj, fastenersArray, "FASTENERSdatalist", "colorFASTENERS", "noteFASTENERS", "unitFASTENERS", "ftFASTENERS", "inchFASTENERS", "fractionDropFASTENERS", "pmFASTENERS", 9)
    
    ' call misc
    Call GetWorksheet("MISC", currenWorkSheet)
     Call FillData(currenWorkSheet, obj, miscArray, "MISCdatalist", "colorMISC", "noteMISC", "unitMISC", "ftMISC", "inchMISC", "fractionDropMISC", "pmMISC", 10)
    
    MsgBox "Script execution completed. Please verify the data, Save cutlist and press OK to close the browser", vbInformation, "Script Completed"
                       
End Sub

' Function to introduce a delay
Sub Delay(ms)
    ' Calculate end time
    endTime = Timer + (ms / 1000)
    
    ' Loop until end time is reached
    Do While Timer < endTime
        ' Continue the loop
        ' This loop is necessary to introduce the delay
    Loop
End Sub

Sub FillData(Worksheet, obj, ByRef myArray, description, colord, note, unit, ft, inch, fractionDrop, pm, index)
    
    ' Iterate over non-empty values starting from cell C3
    Dim row, column
    row = 3 ' Start from row 3
    column = 1 ' Start from column 1

     ' variables to collecy the information from the excel sheet
    Dim LGdescription
    Dim PMValue
    Dim color
    Dim Notes
    Dim Units
    Dim Dimentions, parts, var1, var2, var3
    Dim Fab
    
    ' obj.FindElementById("STRUCTURALdatalist").SendKeys (LGdescription)
    Do While Not IsEmpty(Worksheet.Cells(row, column).Value) Or Not IsEmpty(Worksheet.Cells(row + 1, column).Value)
    
        If IsEmpty(Worksheet.Cells(row, 3).Value) Then
            ' Click on Add Blank element
                Set parentBlankElement = obj.FindElementsByClass("CutlistTable").Item(index)
                Set childBlankElement = parentBlankElement.FindElementByClass("btn-outline-dark")
                childBlankElement.SendKeys ("a")
                childBlankElement.Click

                row = row + 1
        End If
        
        ' Capture the value from the worksheet
        LGdescription = Worksheet.Cells(row, 3).Value
        PMValue = Worksheet.Cells(row, 2).Value
        color = Worksheet.Cells(row, 8).Value
        Notes = Worksheet.Cells(row, 10).Value
        Units = Worksheet.Cells(row, 11).Value
        Dimentions = Worksheet.Cells(row, 5).Value
        Fab = Worksheet.Cells(row, 9).Value
        
        ' Flag to indicate if the target string is found
        Dim found
        found = False
      
        ' Flag to indicate if the description is same as above row
        ' Used for custom items
        Dim isDescriptionSameAsAbove
        If Worksheet.Cells(row - 1, 3).Value = Worksheet.Cells(row, 3).Value Then
            isDescriptionSameAsAbove = True
        Else
            isDescriptionSameAsAbove = False
        End If
            
        
        ' Loop through the array and check each element
        For i = 0 To UBound(myArray)
            Dim temp1
            temp1 = myArray(i)
            temp1 = Trim(temp1)
            
            Dim temp2
            temp2 = LGdescription
            temp2 = Trim(temp2)
            
            If temp1 = temp2 Then
                ' Target string found in the array
                LGdescription = myArray(i)
                found = True
                Exit For
            End If
        Next
        
        ' Split the dimension into three different variables
        Call SplitDimensions(Dimentions, var1, var2, var3)
         
        If found Then
        
            ' Access the HTML element and set its value
            obj.FindElementById(description).SendKeys (LGdescription)
            
            If Len(colord) > 0 Then
                obj.FindElementById(colord).SendKeys (color)
            End If
            
            obj.FindElementById(note).SendKeys (Notes)
            obj.FindElementById(unit).SendKeys (Units)
            
            ' Update dimensions
            obj.FindElementById(ft).SendKeys (var1)
            obj.FindElementById(inch).SendKeys (var2)
            If Len(fractionDrop) > 0 Then
                obj.FindElementById(fractionDrop).SendKeys (var3)
            End If

            If Len(Fab) > 0 Then
                obj.FindElementById("punchSTRUCTURAL").SendKeys (Fab)
            End If

            ' Click on Add Row element
            obj.FindElementById(pm).SendKeys (PMValue)
            Set parentElement = obj.FindElementsByClass("CutlistTable").Item(index)
            Set childElement = parentElement.FindElementByClass("btn-outline-success")
            childElement.Click
        Else
        
            If Not isDescriptionSameAsAbove Then
                ' Click on Add Blank element
                Set parentBlankElement = obj.FindElementsByClass("CutlistTable").Item(11)
                Set childBlankElement = parentBlankElement.FindElementByClass("btn-outline-dark")
                childBlankElement.SendKeys ("a")
                childBlankElement.Click
            End If
        
            ' Access the HTML element and set its value
            obj.FindElementById("CUSTOMDescription").SendKeys (LGdescription)
               
            obj.FindElementById("colorCUSTOM").SendKeys (color)

            If Len(Fab) > 0 Then
                obj.FindElementById("punchCUSTOM").SendKeys (Fab)
            End If
            
            obj.FindElementById("noteCUSTOM").SendKeys (Notes)
            obj.FindElementById("unitCUSTOM").SendKeys (Units)
            
            ' Update dimensions
            obj.FindElementById("ftCUSTOM").SendKeys (var1)
            obj.FindElementById("inchCUSTOM").SendKeys (var2)
            obj.FindElementById("fractionDropCUSTOM").SendKeys (var3)
                              
            ' Add PM
            obj.FindElementById("pmCUSTOM").SendKeys (PMValue)
            
            ' Extra Elements on CUSTOM
            Dim customga
            customga = Worksheet.Cells(row, 4).Value
            
            If customga = "" Then
                customga = "0"
            End If
            
            obj.FindElementById("CUSTOMga").SendKeys (customga)
            obj.FindElementById("CUSTOMcoilWidth").SendKeys ("0")
            
            ' Click on Add Row element
            Set parentElement = obj.FindElementsByClass("CutlistTable").Item(11)
            Set childElement = parentElement.FindElementByClass("btn-outline-success")
            childElement.Click
        End If
        row = row + 1
    Loop
        
End Sub

Sub SplitDimensions(Dimentions, ByRef var1, ByRef var2, ByRef var3)
    ' Remove double quotes if they exist
        Dimentions = Replace(Dimentions, """", "")
        
        ' Split the string into parts
        parts = Split(Dimentions, " - ")
        
        If UBound(parts) > 0 Then
            ' Assign the parts to separate variables
            var1 = Split(parts(0), "'")(0)
            var2 = Split(parts(1), " ")(0)
            If UBound(Split(parts(1), " ")) > 0 Then
                var3 = Split(parts(1), " ")(1)
            Else
                var3 = ""
            End If
        Else
            var1 = ""
            var2 = ""
            var3 = ""
        End If
End Sub

Sub GetWorksheet(name As String, ByRef workbk As Worksheet)
    Dim wb As Workbook
    ' Loop through all open workbooks
    For Each wb In Workbooks
        ' Check if the workbook contains the desired worksheet
        If WorksheetExists(wb, name) Then
            ' Fetch the worksheet named "LIGHT GAUGE" from the current workbook
            Set workbk = wb.Worksheets(name)
                  
            ' Exit the loop once the worksheet is found
            Exit For
        End If
    Next wb
End Sub

' Function to check if a worksheet exists in a workbook
Function WorksheetExists(wb As Workbook, sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    WorksheetExists = Not ws Is Nothing
End Function














