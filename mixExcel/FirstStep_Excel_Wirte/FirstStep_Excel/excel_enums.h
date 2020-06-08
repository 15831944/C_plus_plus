#pragma once

    typedef enum {
        xlAll = 0xffffeff8,
        xlAutomatic = 0xffffeff7,
        xlBoth = 1,
        xlCenter = 0xffffeff4,
        xlChecker = 9,
        xlCircle = 8,
        xlCorner = 2,
        xlCrissCross = 16,
        xlCross = 4,
        xlDiamond = 2,
        xlDistributed = 0xffffefeb,
        xlDoubleAccounting = 5,
        xlFixedValue = 1,
        xlFormats = 0xffffefe6,
        xlGray16 = 17,
        xlGray8 = 18,
        xlGrid = 15,
        xlHigh = 0xffffefe1,
        xlInside = 2,
        xlJustify = 0xffffefde,
        xlLightDown = 13,
        xlLightHorizontal = 11,
        xlLightUp = 14,
        xlLightVertical = 12,
        xlLow = 0xffffefda,
        xlManual = 0xffffefd9,
        xlMinusValues = 3,
        xlModule = 0xffffefd3,
        xlNextToAxis = 4,
        xlNone = 0xffffefd2,
        xlNotes = 0xffffefd0,
        xlOff = 0xffffefce,
        xlOn = 1,
        xlPercent = 2,
        xlPlus = 9,
        xlPlusValues = 2,
        xlSemiGray75 = 10,
        xlShowLabel = 4,
        xlShowLabelAndPercent = 5,
        xlShowPercent = 3,
        xlShowValue = 2,
        xlSimple = 0xffffefc6,
        xlSingle = 2,
        xlSingleAccounting = 4,
        xlSolid = 1,
        xlSquare = 1,
        xlStar = 5,
        xlStError = 4,
        xlToolbarButton = 2,
        xlTriangle = 3,
        xlGray25 = 0xffffefe4,
        xlGray50 = 0xffffefe3,
        xlGray75 = 0xffffefe2,
        xlBottom = 0xffffeff5,
        xlLeft = 0xffffefdd,
        xlRight = 0xffffefc8,
        xlTop = 0xffffefc0,
        xl3DBar = 0xffffeffd,
        xl3DSurface = 0xffffeff9,
        xlBar = 2,
        xlColumn = 3,
        xlCombination = 0xffffeff1,
        xlCustom = 0xffffefee,
        xlDefaultAutoFormat = 0xffffffff,
        xlMaximum = 2,
        xlMinimum = 4,
        xlOpaque = 3,
        xlTransparent = 2,
        xlBidi = 0xffffec78,
        xlLatin = 0xffffec77,
        xlContext = 0xffffec76,
        xlLTR = 0xffffec75,
        xlRTL = 0xffffec74,
        xlFullScript = 1,
        xlPartialScript = 2,
        xlMixedScript = 3,
        xlMixedAuthorizedScript = 4,
        xlVisualCursor = 2,
        xlLogicalCursor = 1,
        xlSystem = 1,
        xlPartial = 3,
        xlHindiNumerals = 3,
        xlBidiCalendar = 3,
        xlGregorian = 2,
        xlComplete = 4,
        xlScale = 3,
        xlClosed = 3,
        xlColor1 = 7,
        xlColor2 = 8,
        xlColor3 = 9,
        xlConstants = 2,
        xlContents = 2,
        xlBelow = 1,
        xlCascade = 7,
        xlCenterAcrossSelection = 7,
        xlChart4 = 2,
        xlChartSeries = 17,
        xlChartShort = 6,
        xlChartTitles = 18,
        xlClassic1 = 1,
        xlClassic2 = 2,
        xlClassic3 = 3,
        xl3DEffects1 = 13,
        xl3DEffects2 = 14,
        xlAbove = 0,
        xlAccounting1 = 4,
        xlAccounting2 = 5,
        xlAccounting3 = 6,
        xlAccounting4 = 17,
        xlAdd = 2,
        xlDebugCodePane = 13,
        xlDesktop = 9,
        xlDirect = 1,
        xlDivide = 5,
        xlDoubleClosed = 5,
        xlDoubleOpen = 4,
        xlDoubleQuote = 1,
        xlEntireChart = 20,
        xlExcelMenus = 1,
        xlExtended = 3,
        xlFill = 5,
        xlFirst = 0,
        xlFloating = 5,
        xlFormula = 5,
        xlGeneral = 1,
        xlGridline = 22,
        xlIcons = 1,
        xlImmediatePane = 12,
        xlInteger = 2,
        xlLast = 1,
        xlLastCell = 11,
        xlList1 = 10,
        xlList2 = 11,
        xlList3 = 12,
        xlLocalFormat1 = 15,
        xlLocalFormat2 = 16,
        xlLong = 3,
        xlLotusHelp = 2,
        xlMacrosheetCell = 7,
        xlMixed = 2,
        xlMultiply = 4,
        xlNarrow = 1,
        xlNoDocuments = 3,
        xlOpen = 2,
        xlOutside = 3,
        xlReference = 4,
        xlSemiautomatic = 2,
        xlShort = 1,
        xlSingleQuote = 2,
        xlStrict = 2,
        xlSubtract = 3,
        xlTextBox = 16,
        xlTiled = 1,
        xlTitleBar = 8,
        xlToolbar = 1,
        xlVisible = 12,
        xlWatchPane = 11,
        xlWide = 3,
        xlWorkbookTab = 6,
        xlWorksheet4 = 1,
        xlWorksheetCell = 3,
        xlWorksheetShort = 5,
        xlAllExceptBorders = 7,
        xlLeftToRight = 2,
        xlTopToBottom = 1,
        xlVeryHidden = 2,
        xlDrawingObject = 14
    } Constants;

    typedef enum {
        xlCreatorCode = 0x5843454c
    } XlCreator;

    typedef enum {
        xlBuiltIn = 21,
        xlUserDefined = 22,
        xlAnyGallery = 23
    } XlChartGallery;

    typedef enum {
        xlColorIndexAutomatic = 0xffffeff7,
        xlColorIndexNone = 0xffffefd2
    } XlColorIndex;

    typedef enum {
        xlCap = 1,
        xlNoCap = 2
    } XlEndStyleCap;

    typedef enum {
        xlColumns = 2,
        xlRows = 1
    } XlRowCol;

    typedef enum {
        xlScaleLinear = 0xffffefdc,
        xlScaleLogarithmic = 0xffffefdb
    } XlScaleType;

    typedef enum {
        xlAutoFill = 4,
        xlChronological = 3,
        xlGrowth = 2,
        xlDataSeriesLinear = 0xffffefdc
    } XlDataSeriesType;

    typedef enum {
        xlAxisCrossesAutomatic = 0xffffeff7,
        xlAxisCrossesCustom = 0xffffefee,
        xlAxisCrossesMaximum = 2,
        xlAxisCrossesMinimum = 4
    } XlAxisCrosses;

    typedef enum {
        xlPrimary = 1,
        xlSecondary = 2
    } XlAxisGroup;

    typedef enum {
        xlBackgroundAutomatic = 0xffffeff7,
        xlBackgroundOpaque = 3,
        xlBackgroundTransparent = 2
    } XlBackground;

    typedef enum {
        xlMaximized = 0xffffefd7,
        xlMinimized = 0xffffefd4,
        xlNormal = 0xffffefd1
    } XlWindowState;

    typedef enum {
        xlCategory = 1,
        xlSeriesAxis = 3,
        xlValue = 2
    } XlAxisType;

    typedef enum {
        xlArrowHeadLengthLong = 3,
        xlArrowHeadLengthMedium = 0xffffefd6,
        xlArrowHeadLengthShort = 1
    } XlArrowHeadLength;

    typedef enum {
        xlVAlignBottom = 0xffffeff5,
        xlVAlignCenter = 0xffffeff4,
        xlVAlignDistributed = 0xffffefeb,
        xlVAlignJustify = 0xffffefde,
        xlVAlignTop = 0xffffefc0
    } XlVAlign;

    typedef enum {
        xlTickMarkCross = 4,
        xlTickMarkInside = 2,
        xlTickMarkNone = 0xffffefd2,
        xlTickMarkOutside = 3
    } XlTickMark;

    typedef enum {
        xlX = 0xffffefb8,
        xlY = 1
    } XlErrorBarDirection;

    typedef enum {
        xlErrorBarIncludeBoth = 1,
        xlErrorBarIncludeMinusValues = 3,
        xlErrorBarIncludeNone = 0xffffefd2,
        xlErrorBarIncludePlusValues = 2
    } XlErrorBarInclude;

    typedef enum {
        xlInterpolated = 3,
        xlNotPlotted = 1,
        xlZero = 2
    } XlDisplayBlanksAs;

    typedef enum {
        xlArrowHeadStyleClosed = 3,
        xlArrowHeadStyleDoubleClosed = 5,
        xlArrowHeadStyleDoubleOpen = 4,
        xlArrowHeadStyleNone = 0xffffefd2,
        xlArrowHeadStyleOpen = 2
    } XlArrowHeadStyle;

    typedef enum {
        xlArrowHeadWidthMedium = 0xffffefd6,
        xlArrowHeadWidthNarrow = 1,
        xlArrowHeadWidthWide = 3
    } XlArrowHeadWidth;

    typedef enum {
        xlHAlignCenter = 0xffffeff4,
        xlHAlignCenterAcrossSelection = 7,
        xlHAlignDistributed = 0xffffefeb,
        xlHAlignFill = 5,
        xlHAlignGeneral = 1,
        xlHAlignJustify = 0xffffefde,
        xlHAlignLeft = 0xffffefdd,
        xlHAlignRight = 0xffffefc8
    } XlHAlign;

    typedef enum {
        xlTickLabelPositionHigh = 0xffffefe1,
        xlTickLabelPositionLow = 0xffffefda,
        xlTickLabelPositionNextToAxis = 4,
        xlTickLabelPositionNone = 0xffffefd2
    } XlTickLabelPosition;

    typedef enum {
        xlLegendPositionBottom = 0xffffeff5,
        xlLegendPositionCorner = 2,
        xlLegendPositionLeft = 0xffffefdd,
        xlLegendPositionRight = 0xffffefc8,
        xlLegendPositionTop = 0xffffefc0
    } XlLegendPosition;

    typedef enum {
        xlStackScale = 3,
        xlStack = 2,
        xlStretch = 1
    } XlChartPictureType;

    typedef enum {
        xlSides = 1,
        xlEnd = 2,
        xlEndSides = 3,
        xlFront = 4,
        xlFrontSides = 5,
        xlFrontEnd = 6,
        xlAllFaces = 7
    } XlChartPicturePlacement;

    typedef enum {
        xlDownward = 0xffffefb6,
        xlHorizontal = 0xffffefe0,
        xlUpward = 0xffffefb5,
        xlVertical = 0xffffefba
    } XlOrientation;

    typedef enum {
        xlTickLabelOrientationAutomatic = 0xffffeff7,
        xlTickLabelOrientationDownward = 0xffffefb6,
        xlTickLabelOrientationHorizontal = 0xffffefe0,
        xlTickLabelOrientationUpward = 0xffffefb5,
        xlTickLabelOrientationVertical = 0xffffefba
    } XlTickLabelOrientation;

    typedef enum {
        xlHairline = 1,
        xlMedium = 0xffffefd6,
        xlThick = 4,
        xlThin = 2
    } XlBorderWeight;

    typedef enum {
        xlDay = 1,
        xlMonth = 3,
        xlWeekday = 2,
        xlYear = 4
    } XlDataSeriesDate;

    typedef enum {
        xlUnderlineStyleDouble = 0xffffefe9,
        xlUnderlineStyleDoubleAccounting = 5,
        xlUnderlineStyleNone = 0xffffefd2,
        xlUnderlineStyleSingle = 2,
        xlUnderlineStyleSingleAccounting = 4
    } XlUnderlineStyle;

    typedef enum {
        xlErrorBarTypeCustom = 0xffffefee,
        xlErrorBarTypeFixedValue = 1,
        xlErrorBarTypePercent = 2,
        xlErrorBarTypeStDev = 0xffffefc5,
        xlErrorBarTypeStError = 4
    } XlErrorBarType;

    typedef enum {
        xlExponential = 5,
        xlLinear = 0xffffefdc,
        xlLogarithmic = 0xffffefdb,
        xlMovingAvg = 6,
        xlPolynomial = 3,
        xlPower = 4
    } XlTrendlineType;

    typedef enum {
        xlContinuous = 1,
        xlDash = 0xffffefed,
        xlDashDot = 4,
        xlDashDotDot = 5,
        xlDot = 0xffffefea,
        xlDouble = 0xffffefe9,
        xlSlantDashDot = 13,
        xlLineStyleNone = 0xffffefd2
    } XlLineStyle;

    typedef enum {
        xlDataLabelsShowNone = 0xffffefd2,
        xlDataLabelsShowValue = 2,
        xlDataLabelsShowPercent = 3,
        xlDataLabelsShowLabel = 4,
        xlDataLabelsShowLabelAndPercent = 5,
        xlDataLabelsShowBubbleSizes = 6
    } XlDataLabelsType;

    typedef enum {
        xlMarkerStyleAutomatic = 0xffffeff7,
        xlMarkerStyleCircle = 8,
        xlMarkerStyleDash = 0xffffefed,
        xlMarkerStyleDiamond = 2,
        xlMarkerStyleDot = 0xffffefea,
        xlMarkerStyleNone = 0xffffefd2,
        xlMarkerStylePicture = 0xffffefcd,
        xlMarkerStylePlus = 9,
        xlMarkerStyleSquare = 1,
        xlMarkerStyleStar = 5,
        xlMarkerStyleTriangle = 3,
        xlMarkerStyleX = 0xffffefb8
    } XlMarkerStyle;

    typedef enum {
        xlBMP = 1,
        xlCGM = 7,
        xlDRW = 4,
        xlDXF = 5,
        xlEPS = 8,
        xlHGL = 6,
        xlPCT = 13,
        xlPCX = 10,
        xlPIC = 11,
        xlPLT = 12,
        xlTIF = 9,
        xlWMF = 2,
        xlWPG = 3
    } XlPictureConvertorType;

    typedef enum {
        xlPatternAutomatic = 0xffffeff7,
        xlPatternChecker = 9,
        xlPatternCrissCross = 16,
        xlPatternDown = 0xffffefe7,
        xlPatternGray16 = 17,
        xlPatternGray25 = 0xffffefe4,
        xlPatternGray50 = 0xffffefe3,
        xlPatternGray75 = 0xffffefe2,
        xlPatternGray8 = 18,
        xlPatternGrid = 15,
        xlPatternHorizontal = 0xffffefe0,
        xlPatternLightDown = 13,
        xlPatternLightHorizontal = 11,
        xlPatternLightUp = 14,
        xlPatternLightVertical = 12,
        xlPatternNone = 0xffffefd2,
        xlPatternSemiGray75 = 10,
        xlPatternSolid = 1,
        xlPatternUp = 0xffffefbe,
        xlPatternVertical = 0xffffefba
    } XlPattern;

    typedef enum {
        xlSplitByPosition = 1,
        xlSplitByPercentValue = 3,
        xlSplitByCustomSplit = 4,
        xlSplitByValue = 2
    } XlChartSplitType;

    typedef enum {
        xlHundreds = 0xfffffffe,
        xlThousands = 0xfffffffd,
        xlTenThousands = 0xfffffffc,
        xlHundredThousands = 0xfffffffb,
        xlMillions = 0xfffffffa,
        xlTenMillions = 0xfffffff9,
        xlHundredMillions = 0xfffffff8,
        xlThousandMillions = 0xfffffff7,
        xlMillionMillions = 0xfffffff6
    } XlDisplayUnit;

    typedef enum {
        xlLabelPositionCenter = 0xffffeff4,
        xlLabelPositionAbove = 0,
        xlLabelPositionBelow = 1,
        xlLabelPositionLeft = 0xffffefdd,
        xlLabelPositionRight = 0xffffefc8,
        xlLabelPositionOutsideEnd = 2,
        xlLabelPositionInsideEnd = 3,
        xlLabelPositionInsideBase = 4,
        xlLabelPositionBestFit = 5,
        xlLabelPositionMixed = 6,
        xlLabelPositionCustom = 7
    } XlDataLabelPosition;

    typedef enum {
        xlDays = 0,
        xlMonths = 1,
        xlYears = 2
    } XlTimeUnit;

    typedef enum {
        xlCategoryScale = 2,
        xlTimeScale = 3,
        xlAutomaticScale = 0xffffeff7
    } XlCategoryType;

    typedef enum {
        xlBox = 0,
        xlPyramidToPoint = 1,
        xlPyramidToMax = 2,
        xlCylinder = 3,
        xlConeToPoint = 4,
        xlConeToMax = 5
    } XlBarShape;

    typedef enum {
        xlColumnClustered = 51,
        xlColumnStacked = 52,
        xlColumnStacked100 = 53,
        xl3DColumnClustered = 54,
        xl3DColumnStacked = 55,
        xl3DColumnStacked100 = 56,
        xlBarClustered = 57,
        xlBarStacked = 58,
        xlBarStacked100 = 59,
        xl3DBarClustered = 60,
        xl3DBarStacked = 61,
        xl3DBarStacked100 = 62,
        xlLineStacked = 63,
        xlLineStacked100 = 64,
        xlLineMarkers = 65,
        xlLineMarkersStacked = 66,
        xlLineMarkersStacked100 = 67,
        xlPieOfPie = 68,
        xlPieExploded = 69,
        xl3DPieExploded = 70,
        xlBarOfPie = 71,
        xlXYScatterSmooth = 72,
        xlXYScatterSmoothNoMarkers = 73,
        xlXYScatterLines = 74,
        xlXYScatterLinesNoMarkers = 75,
        xlAreaStacked = 76,
        xlAreaStacked100 = 77,
        xl3DAreaStacked = 78,
        xl3DAreaStacked100 = 79,
        xlDoughnutExploded = 80,
        xlRadarMarkers = 81,
        xlRadarFilled = 82,
        xlSurface = 83,
        xlSurfaceWireframe = 84,
        xlSurfaceTopView = 85,
        xlSurfaceTopViewWireframe = 86,
        xlBubble = 15,
        xlBubble3DEffect = 87,
        xlStockHLC = 88,
        xlStockOHLC = 89,
        xlStockVHLC = 90,
        xlStockVOHLC = 91,
        xlCylinderColClustered = 92,
        xlCylinderColStacked = 93,
        xlCylinderColStacked100 = 94,
        xlCylinderBarClustered = 95,
        xlCylinderBarStacked = 96,
        xlCylinderBarStacked100 = 97,
        xlCylinderCol = 98,
        xlConeColClustered = 99,
        xlConeColStacked = 100,
        xlConeColStacked100 = 101,
        xlConeBarClustered = 102,
        xlConeBarStacked = 103,
        xlConeBarStacked100 = 104,
        xlConeCol = 105,
        xlPyramidColClustered = 106,
        xlPyramidColStacked = 107,
        xlPyramidColStacked100 = 108,
        xlPyramidBarClustered = 109,
        xlPyramidBarStacked = 110,
        xlPyramidBarStacked100 = 111,
        xlPyramidCol = 112,
        xl3DColumn = 0xffffeffc,
        xlLine = 4,
        xl3DLine = 0xffffeffb,
        xl3DPie = 0xffffeffa,
        xlPie = 5,
        xlXYScatter = 0xffffefb7,
        xl3DArea = 0xffffeffe,
        xlArea = 1,
        xlDoughnut = 0xffffefe8,
        xlRadar = 0xffffefc9
    } XlChartType;

    typedef enum {
        xlDataLabel = 0,
        xlChartArea = 2,
        xlSeries = 3,
        xlChartTitle = 4,
        xlWalls = 5,
        xlCorners = 6,
        xlDataTable = 7,
        xlTrendline = 8,
        xlErrorBars = 9,
        xlXErrorBars = 10,
        xlYErrorBars = 11,
        xlLegendEntry = 12,
        xlLegendKey = 13,
        xlShape = 14,
        xlMajorGridlines = 15,
        xlMinorGridlines = 16,
        xlAxisTitle = 17,
        xlUpBars = 18,
        xlPlotArea = 19,
        xlDownBars = 20,
        xlAxis = 21,
        xlSeriesLines = 22,
        xlFloor = 23,
        xlLegend = 24,
        xlHiLoLines = 25,
        xlDropLines = 26,
        xlRadarAxisLabels = 27,
        xlNothing = 28,
        xlLeaderLines = 29,
        xlDisplayUnitLabel = 30,
        xlPivotChartFieldButton = 31,
        xlPivotChartDropZone = 32
    } XlChartItem;

    typedef enum {
        xlSizeIsWidth = 2,
        xlSizeIsArea = 1
    } XlSizeRepresents;

    typedef enum {
        xlShiftDown = 0xffffefe7,
        xlShiftToRight = 0xffffefbf
    } XlInsertShiftDirection;

    typedef enum {
        xlShiftToLeft = 0xffffefc1,
        xlShiftUp = 0xffffefbe
    } XlDeleteShiftDirection;

    typedef enum {
        xlDown = 0xffffefe7,
        xlToLeft = 0xffffefc1,
        xlToRight = 0xffffefbf,
        xlUp = 0xffffefbe
    } XlDirection;

    typedef enum {
        xlAverage = 0xffffeff6,
        xlCount = 0xffffeff0,
        xlCountNums = 0xffffefef,
        xlMax = 0xffffefd8,
        xlMin = 0xffffefd5,
        xlProduct = 0xffffefcb,
        xlStDev = 0xffffefc5,
        xlStDevP = 0xffffefc4,
        xlSum = 0xffffefc3,
        xlVar = 0xffffefbc,
        xlVarP = 0xffffefbb,
        xlUnknown = 1000
    } XlConsolidationFunction;

    typedef enum {
        xlChart = 0xffffeff3,
        xlDialogSheet = 0xffffefec,
        xlExcel4IntlMacroSheet = 4,
        xlExcel4MacroSheet = 3,
        xlWorksheet = 0xffffefb9
    } XlSheetType;

    typedef enum {
        xlColumnHeader = 0xffffeff2,
        xlColumnItem = 5,
        xlDataHeader = 3,
        xlDataItem = 7,
        xlPageHeader = 2,
        xlPageItem = 6,
        xlRowHeader = 0xffffefc7,
        xlRowItem = 4,
        xlTableBody = 8
    } XlLocationInTable;

    typedef enum {
        xlFormulas = 0xffffefe5,
        xlComments = 0xffffefd0,
        xlValues = 0xffffefbd
    } XlFindLookIn;

    typedef enum {
        xlChartAsWindow = 5,
        xlChartInPlace = 4,
        xlClipboard = 3,
        xlInfo = 0xffffefdf,
        xlWorkbook = 1
    } XlWindowType;

    typedef enum {
        xlDate = 2,
        xlNumber = 0xffffefcf,
        xlText = 0xffffefc2
    } XlPivotFieldDataType;

    typedef enum {
        xlBitmap = 2,
        xlPicture = 0xffffefcd
    } XlCopyPictureFormat;

    typedef enum {
        xlScenario = 4,
        xlConsolidation = 3,
        xlDatabase = 1,
        xlExternal = 2,
        xlPivotTable = 0xffffefcc
    } XlPivotTableSourceType;

    typedef enum {
        xlA1 = 1,
        xlR1C1 = 0xffffefca
    } XlReferenceStyle;

    typedef enum {
        xlMicrosoftAccess = 4,
        xlMicrosoftFoxPro = 5,
        xlMicrosoftMail = 3,
        xlMicrosoftPowerPoint = 2,
        xlMicrosoftProject = 6,
        xlMicrosoftSchedulePlus = 7,
        xlMicrosoftWord = 1
    } XlMSApplication;

    typedef enum {
        xlNoButton = 0,
        xlPrimaryButton = 1,
        xlSecondaryButton = 2
    } XlMouseButton;

    typedef enum {
        xlCopy = 1,
        xlCut = 2
    } XlCutCopyMode;

    typedef enum {
        xlFillWithAll = 0xffffeff8,
        xlFillWithContents = 2,
        xlFillWithFormats = 0xffffefe6
    } XlFillWith;

    typedef enum {
        xlFilterCopy = 2,
        xlFilterInPlace = 1
    } XlFilterAction;

    typedef enum {
        xlDownThenOver = 1,
        xlOverThenDown = 2
    } XlOrder;

    typedef enum {
        xlLinkTypeExcelLinks = 1,
        xlLinkTypeOLELinks = 2
    } XlLinkType;

    typedef enum {
        xlColumnThenRow = 2,
        xlRowThenColumn = 1
    } XlApplyNamesOrder;

    typedef enum {
        xlDisabled = 0,
        xlErrorHandler = 2,
        xlInterrupt = 1
    } XlEnableCancelKey;

    typedef enum {
        xlPageBreakAutomatic = 0xffffeff7,
        xlPageBreakManual = 0xffffefd9,
        xlPageBreakNone = 0xffffefd2
    } XlPageBreak;

    typedef enum {
        xlOLEControl = 2,
        xlOLEEmbed = 1,
        xlOLELink = 0
    } XlOLEType;

    typedef enum {
        xlLandscape = 2,
        xlPortrait = 1
    } XlPageOrientation;

    typedef enum {
        xlEditionDate = 2,
        xlUpdateState = 1,
        xlLinkInfoStatus = 3
    } XlLinkInfo;

    typedef enum {
        xlCommandUnderlinesAutomatic = 0xffffeff7,
        xlCommandUnderlinesOff = 0xffffefce,
        xlCommandUnderlinesOn = 1
    } XlCommandUnderlines;

    typedef enum {
        xlVerbOpen = 2,
        xlVerbPrimary = 1
    } XlOLEVerb;

    typedef enum {
        xlCalculationAutomatic = 0xffffeff7,
        xlCalculationManual = 0xffffefd9,
        xlCalculationSemiautomatic = 2
    } XlCalculation;

    typedef enum {
        xlReadOnly = 3,
        xlReadWrite = 2
    } XlFileAccess;

    typedef enum {
        xlPublisher = 1,
        xlSubscriber = 2
    } XlEditionType;

    typedef enum {
        xlFitToPage = 2,
        xlFullPage = 3,
        xlScreenSize = 1
    } XlObjectSize;

    typedef enum {
        xlPart = 2,
        xlWhole = 1
    } XlLookAt;

    typedef enum {
        xlMAPI = 1,
        xlNoMailSystem = 0,
        xlPowerTalk = 2
    } XlMailSystem;

    typedef enum {
        xlLinkInfoOLELinks = 2,
        xlLinkInfoPublishers = 5,
        xlLinkInfoSubscribers = 6
    } XlLinkInfoType;

    typedef enum {
        xlErrDiv0 = 2007,
        xlErrNA = 2042,
        xlErrName = 2029,
        xlErrNull = 2000,
        xlErrNum = 2036,
        xlErrRef = 2023,
        xlErrValue = 2015
    } XlCVError;

    typedef enum {
        xlBIFF = 2,
        xlPICT = 1,
        xlRTF = 4,
        xlVALU = 8
    } XlEditionFormat;

    typedef enum {
        xlExcelLinks = 1,
        xlOLELinks = 2,
        xlPublishers = 5,
        xlSubscribers = 6
    } XlLink;

    typedef enum {
        xlCellTypeBlanks = 4,
        xlCellTypeConstants = 2,
        xlCellTypeFormulas = 0xffffefe5,
        xlCellTypeLastCell = 11,
        xlCellTypeComments = 0xffffefd0,
        xlCellTypeVisible = 12,
        xlCellTypeAllFormatConditions = 0xffffefb4,
        xlCellTypeSameFormatConditions = 0xffffefb3,
        xlCellTypeAllValidation = 0xffffefb2,
        xlCellTypeSameValidation = 0xffffefb1
    } XlCellType;

    typedef enum {
        xlArrangeStyleCascade = 7,
        xlArrangeStyleHorizontal = 0xffffefe0,
        xlArrangeStyleTiled = 1,
        xlArrangeStyleVertical = 0xffffefba
    } XlArrangeStyle;

    typedef enum {
        xlIBeam = 3,
        xlDefault = 0xffffefd1,
        xlNorthwestArrow = 1,
        xlWait = 2
    } XlMousePointer;

    typedef enum {
        xlAutomaticUpdate = 4,
        xlCancel = 1,
        xlChangeAttributes = 6,
        xlManualUpdate = 5,
        xlOpenSource = 3,
        xlSelect = 3,
        xlSendPublisher = 2,
        xlUpdateSubscriber = 2
    } XlEditionOptionsOption;

    typedef enum {
        xlFillCopy = 1,
        xlFillDays = 5,
        xlFillDefault = 0,
        xlFillFormats = 3,
        xlFillMonths = 7,
        xlFillSeries = 2,
        xlFillValues = 4,
        xlFillWeekdays = 6,
        xlFillYears = 8,
        xlGrowthTrend = 10,
        xlLinearTrend = 9
    } XlAutoFillType;

    typedef enum {
        xlAnd = 1,
        xlBottom10Items = 4,
        xlBottom10Percent = 6,
        xlOr = 2,
        xlTop10Items = 3,
        xlTop10Percent = 5
    } XlAutoFilterOperator;

    typedef enum {
        xlClipboardFormatBIFF = 8,
        xlClipboardFormatBIFF2 = 18,
        xlClipboardFormatBIFF3 = 20,
        xlClipboardFormatBIFF4 = 30,
        xlClipboardFormatBinary = 15,
        xlClipboardFormatBitmap = 9,
        xlClipboardFormatCGM = 13,
        xlClipboardFormatCSV = 5,
        xlClipboardFormatDIF = 4,
        xlClipboardFormatDspText = 12,
        xlClipboardFormatEmbeddedObject = 21,
        xlClipboardFormatEmbedSource = 22,
        xlClipboardFormatLink = 11,
        xlClipboardFormatLinkSource = 23,
        xlClipboardFormatLinkSourceDesc = 32,
        xlClipboardFormatMovie = 24,
        xlClipboardFormatNative = 14,
        xlClipboardFormatObjectDesc = 31,
        xlClipboardFormatObjectLink = 19,
        xlClipboardFormatOwnerLink = 17,
        xlClipboardFormatPICT = 2,
        xlClipboardFormatPrintPICT = 3,
        xlClipboardFormatRTF = 7,
        xlClipboardFormatScreenPICT = 29,
        xlClipboardFormatStandardFont = 28,
        xlClipboardFormatStandardScale = 27,
        xlClipboardFormatSYLK = 6,
        xlClipboardFormatTable = 16,
        xlClipboardFormatText = 0,
        xlClipboardFormatToolFace = 25,
        xlClipboardFormatToolFacePICT = 26,
        xlClipboardFormatVALU = 1,
        xlClipboardFormatWK1 = 10
    } XlClipboardFormat;

    typedef enum {
        xlAddIn = 18,
        xlCSV = 6,
        xlCSVMac = 22,
        xlCSVMSDOS = 24,
        xlCSVWindows = 23,
        xlDBF2 = 7,
        xlDBF3 = 8,
        xlDBF4 = 11,
        xlDIF = 9,
        xlExcel2 = 16,
        xlExcel2FarEast = 27,
        xlExcel3 = 29,
        xlExcel4 = 33,
        xlExcel5 = 39,
        xlExcel7 = 39,
        xlExcel9795 = 43,
        xlExcel4Workbook = 35,
        xlIntlAddIn = 26,
        xlIntlMacro = 25,
        xlWorkbookNormal = 0xffffefd1,
        xlSYLK = 2,
        xlTemplate = 17,
        xlCurrentPlatformText = 0xffffefc2,
        xlTextMac = 19,
        xlTextMSDOS = 21,
        xlTextPrinter = 36,
        xlTextWindows = 20,
        xlWJ2WD1 = 14,
        xlWK1 = 5,
        xlWK1ALL = 31,
        xlWK1FMT = 30,
        xlWK3 = 15,
        xlWK4 = 38,
        xlWK3FM3 = 32,
        xlWKS = 4,
        xlWorks2FarEast = 28,
        xlWQ1 = 34,
        xlWJ3 = 40,
        xlWJ3FJ3 = 41,
        xlUnicodeText = 42,
        xlHtml = 44,
        xlWebArchive = 45,
        xlXMLSpreadsheet = 46
    } XlFileFormat;

    typedef enum {
        xl24HourClock = 33,
        xl4DigitYears = 43,
        xlAlternateArraySeparator = 16,
        xlColumnSeparator = 14,
        xlCountryCode = 1,
        xlCountrySetting = 2,
        xlCurrencyBefore = 37,
        xlCurrencyCode = 25,
        xlCurrencyDigits = 27,
        xlCurrencyLeadingZeros = 40,
        xlCurrencyMinusSign = 38,
        xlCurrencyNegative = 28,
        xlCurrencySpaceBefore = 36,
        xlCurrencyTrailingZeros = 39,
        xlDateOrder = 32,
        xlDateSeparator = 17,
        xlDayCode = 21,
        xlDayLeadingZero = 42,
        xlDecimalSeparator = 3,
        xlGeneralFormatName = 26,
        xlHourCode = 22,
        xlLeftBrace = 12,
        xlLeftBracket = 10,
        xlListSeparator = 5,
        xlLowerCaseColumnLetter = 9,
        xlLowerCaseRowLetter = 8,
        xlMDY = 44,
        xlMetric = 35,
        xlMinuteCode = 23,
        xlMonthCode = 20,
        xlMonthLeadingZero = 41,
        xlMonthNameChars = 30,
        xlNoncurrencyDigits = 29,
        xlNonEnglishFunctions = 34,
        xlRightBrace = 13,
        xlRightBracket = 11,
        xlRowSeparator = 15,
        xlSecondCode = 24,
        xlThousandsSeparator = 4,
        xlTimeLeadingZero = 45,
        xlTimeSeparator = 18,
        xlUpperCaseColumnLetter = 7,
        xlUpperCaseRowLetter = 6,
        xlWeekdayNameChars = 31,
        xlYearCode = 19
    } XlApplicationInternational;

    typedef enum {
        xlPageBreakFull = 1,
        xlPageBreakPartial = 2
    } XlPageBreakExtent;

    typedef enum {
        xlOverwriteCells = 0,
        xlInsertDeleteCells = 1,
        xlInsertEntireRows = 2
    } XlCellInsertionMode;

    typedef enum {
        xlNoLabels = 0xffffefd2,
        xlRowLabels = 1,
        xlColumnLabels = 2,
        xlMixedLabels = 3
    } XlFormulaLabel;

    typedef enum {
        xlSinceMyLastSave = 1,
        xlAllChanges = 2,
        xlNotYetReviewed = 3
    } XlHighlightChangesTime;

    typedef enum {
        xlNoIndicator = 0,
        xlCommentIndicatorOnly = 0xffffffff,
        xlCommentAndIndicator = 1
    } XlCommentDisplayMode;

    typedef enum {
        xlCellValue = 1,
        xlExpression = 2
    } XlFormatConditionType;

    typedef enum {
        xlBetween = 1,
        xlNotBetween = 2,
        xlEqual = 3,
        xlNotEqual = 4,
        xlGreater = 5,
        xlLess = 6,
        xlGreaterEqual = 7,
        xlLessEqual = 8
    } XlFormatConditionOperator;

    typedef enum {
        xlNoRestrictions = 0,
        xlUnlockedCells = 1,
        xlNoSelection = 0xffffefd2
    } XlEnableSelection;

    typedef enum {
        xlValidateInputOnly = 0,
        xlValidateWholeNumber = 1,
        xlValidateDecimal = 2,
        xlValidateList = 3,
        xlValidateDate = 4,
        xlValidateTime = 5,
        xlValidateTextLength = 6,
        xlValidateCustom = 7
    } XlDVType;

    typedef enum {
        xlIMEModeNoControl = 0,
        xlIMEModeOn = 1,
        xlIMEModeOff = 2,
        xlIMEModeDisable = 3,
        xlIMEModeHiragana = 4,
        xlIMEModeKatakana = 5,
        xlIMEModeKatakanaHalf = 6,
        xlIMEModeAlphaFull = 7,
        xlIMEModeAlpha = 8,
        xlIMEModeHangulFull = 9,
        xlIMEModeHangul = 10
    } XlIMEMode;

    typedef enum {
        xlValidAlertStop = 1,
        xlValidAlertWarning = 2,
        xlValidAlertInformation = 3
    } XlDVAlertStyle;

    typedef enum {
        xlLocationAsNewSheet = 1,
        xlLocationAsObject = 2,
        xlLocationAutomatic = 3
    } XlChartLocation;

    typedef enum {
        xlPaper10x14 = 16,
        xlPaper11x17 = 17,
        xlPaperA3 = 8,
        xlPaperA4 = 9,
        xlPaperA4Small = 10,
        xlPaperA5 = 11,
        xlPaperB4 = 12,
        xlPaperB5 = 13,
        xlPaperCsheet = 24,
        xlPaperDsheet = 25,
        xlPaperEnvelope10 = 20,
        xlPaperEnvelope11 = 21,
        xlPaperEnvelope12 = 22,
        xlPaperEnvelope14 = 23,
        xlPaperEnvelope9 = 19,
        xlPaperEnvelopeB4 = 33,
        xlPaperEnvelopeB5 = 34,
        xlPaperEnvelopeB6 = 35,
        xlPaperEnvelopeC3 = 29,
        xlPaperEnvelopeC4 = 30,
        xlPaperEnvelopeC5 = 28,
        xlPaperEnvelopeC6 = 31,
        xlPaperEnvelopeC65 = 32,
        xlPaperEnvelopeDL = 27,
        xlPaperEnvelopeItaly = 36,
        xlPaperEnvelopeMonarch = 37,
        xlPaperEnvelopePersonal = 38,
        xlPaperEsheet = 26,
        xlPaperExecutive = 7,
        xlPaperFanfoldLegalGerman = 41,
        xlPaperFanfoldStdGerman = 40,
        xlPaperFanfoldUS = 39,
        xlPaperFolio = 14,
        xlPaperLedger = 4,
        xlPaperLegal = 5,
        xlPaperLetter = 1,
        xlPaperLetterSmall = 2,
        xlPaperNote = 18,
        xlPaperQuarto = 15,
        xlPaperStatement = 6,
        xlPaperTabloid = 3,
        xlPaperUser = 256
    } XlPaperSize;

    typedef enum {
        xlPasteSpecialOperationAdd = 2,
        xlPasteSpecialOperationDivide = 5,
        xlPasteSpecialOperationMultiply = 4,
        xlPasteSpecialOperationNone = 0xffffefd2,
        xlPasteSpecialOperationSubtract = 3
    } XlPasteSpecialOperation;

    typedef enum {
        xlPasteAll = 0xffffeff8,
        xlPasteAllExceptBorders = 7,
        xlPasteFormats = 0xffffefe6,
        xlPasteFormulas = 0xffffefe5,
        xlPasteComments = 0xffffefd0,
        xlPasteValues = 0xffffefbd,
        xlPasteColumnWidths = 8,
        xlPasteValidation = 6,
        xlPasteFormulasAndNumberFormats = 11,
        xlPasteValuesAndNumberFormats = 12
    } XlPasteType;

    typedef enum {
        xlKatakanaHalf = 0,
        xlKatakana = 1,
        xlHiragana = 2,
        xlNoConversion = 3
    } XlPhoneticCharacterType;

    typedef enum {
        xlPhoneticAlignNoControl = 0,
        xlPhoneticAlignLeft = 1,
        xlPhoneticAlignCenter = 2,
        xlPhoneticAlignDistributed = 3
    } XlPhoneticAlignment;

    typedef enum {
        xlPrinter = 2,
        xlScreen = 1
    } XlPictureAppearance;

    typedef enum {
        xlColumnField = 2,
        xlDataField = 4,
        xlHidden = 0,
        xlPageField = 3,
        xlRowField = 1
    } XlPivotFieldOrientation;

    typedef enum {
        xlDifferenceFrom = 2,
        xlIndex = 9,
        xlNoAdditionalCalculation = 0xffffefd1,
        xlPercentDifferenceFrom = 4,
        xlPercentOf = 3,
        xlPercentOfColumn = 7,
        xlPercentOfRow = 6,
        xlPercentOfTotal = 8,
        xlRunningTotal = 5
    } XlPivotFieldCalculation;

    typedef enum {
        xlFreeFloating = 3,
        xlMove = 2,
        xlMoveAndSize = 1
    } XlPlacement;

    typedef enum {
        xlMacintosh = 1,
        xlMSDOS = 3,
        xlWindows = 2
    } XlPlatform;

    typedef enum {
        xlPrintSheetEnd = 1,
        xlPrintInPlace = 16,
        xlPrintNoComments = 0xffffefd2
    } XlPrintLocation;

    typedef enum {
        xlPriorityHigh = 0xffffefe1,
        xlPriorityLow = 0xffffefda,
        xlPriorityNormal = 0xffffefd1
    } XlPriority;

    typedef enum {
        xlLabelOnly = 1,
        xlDataAndLabel = 0,
        xlDataOnly = 2,
        xlOrigin = 3,
        xlButton = 15,
        xlBlanks = 4,
        xlFirstRow = 256
    } XlPTSelectionMode;

    typedef enum {
        xlRangeAutoFormat3DEffects1 = 13,
        xlRangeAutoFormat3DEffects2 = 14,
        xlRangeAutoFormatAccounting1 = 4,
        xlRangeAutoFormatAccounting2 = 5,
        xlRangeAutoFormatAccounting3 = 6,
        xlRangeAutoFormatAccounting4 = 17,
        xlRangeAutoFormatClassic1 = 1,
        xlRangeAutoFormatClassic2 = 2,
        xlRangeAutoFormatClassic3 = 3,
        xlRangeAutoFormatColor1 = 7,
        xlRangeAutoFormatColor2 = 8,
        xlRangeAutoFormatColor3 = 9,
        xlRangeAutoFormatList1 = 10,
        xlRangeAutoFormatList2 = 11,
        xlRangeAutoFormatList3 = 12,
        xlRangeAutoFormatLocalFormat1 = 15,
        xlRangeAutoFormatLocalFormat2 = 16,
        xlRangeAutoFormatLocalFormat3 = 19,
        xlRangeAutoFormatLocalFormat4 = 20,
        xlRangeAutoFormatReport1 = 21,
        xlRangeAutoFormatReport2 = 22,
        xlRangeAutoFormatReport3 = 23,
        xlRangeAutoFormatReport4 = 24,
        xlRangeAutoFormatReport5 = 25,
        xlRangeAutoFormatReport6 = 26,
        xlRangeAutoFormatReport7 = 27,
        xlRangeAutoFormatReport8 = 28,
        xlRangeAutoFormatReport9 = 29,
        xlRangeAutoFormatReport10 = 30,
        xlRangeAutoFormatClassicPivotTable = 31,
        xlRangeAutoFormatTable1 = 32,
        xlRangeAutoFormatTable2 = 33,
        xlRangeAutoFormatTable3 = 34,
        xlRangeAutoFormatTable4 = 35,
        xlRangeAutoFormatTable5 = 36,
        xlRangeAutoFormatTable6 = 37,
        xlRangeAutoFormatTable7 = 38,
        xlRangeAutoFormatTable8 = 39,
        xlRangeAutoFormatTable9 = 40,
        xlRangeAutoFormatTable10 = 41,
        xlRangeAutoFormatPTNone = 42,
        xlRangeAutoFormatNone = 0xffffefd2,
        xlRangeAutoFormatSimple = 0xffffefc6
    } XlRangeAutoFormat;

    typedef enum {
        xlAbsolute = 1,
        xlAbsRowRelColumn = 2,
        xlRelative = 4,
        xlRelRowAbsColumn = 3
    } XlReferenceType;

    typedef enum {
        xlTabular = 0,
        xlOutline = 1
    } XlLayoutFormType;

    typedef enum {
        xlAllAtOnce = 2,
        xlOneAfterAnother = 1
    } XlRoutingSlipDelivery;

    typedef enum {
        xlNotYetRouted = 0,
        xlRoutingComplete = 2,
        xlRoutingInProgress = 1
    } XlRoutingSlipStatus;

    typedef enum {
        xlAutoActivate = 3,
        xlAutoClose = 2,
        xlAutoDeactivate = 4,
        xlAutoOpen = 1
    } XlRunAutoMacro;

    typedef enum {
        xlDoNotSaveChanges = 2,
        xlSaveChanges = 1
    } XlSaveAction;

    typedef enum {
        xlExclusive = 3,
        xlNoChange = 1,
        xlShared = 2
    } XlSaveAsAccessMode;

    typedef enum {
        xlLocalSessionChanges = 2,
        xlOtherSessionChanges = 3,
        xlUserResolution = 1
    } XlSaveConflictResolution;

    typedef enum {
        xlNext = 1,
        xlPrevious = 2
    } XlSearchDirection;

    typedef enum {
        xlByColumns = 2,
        xlByRows = 1
    } XlSearchOrder;

    typedef enum {
        xlSheetVisible = 0xffffffff,
        xlSheetHidden = 0,
        xlSheetVeryHidden = 2
    } XlSheetVisibility;

    typedef enum {
        xlPinYin = 1,
        xlStroke = 2
    } XlSortMethod;

    typedef enum {
        xlCodePage = 2,
        xlSyllabary = 1
    } XlSortMethodOld;

    typedef enum {
        xlAscending = 1,
        xlDescending = 2
    } XlSortOrder;

    typedef enum {
        xlSortRows = 2,
        xlSortColumns = 1
    } XlSortOrientation;

    typedef enum {
        xlSortLabels = 2,
        xlSortValues = 1
    } XlSortType;

    typedef enum {
        xlErrors = 16,
        xlLogical = 4,
        xlNumbers = 1,
        xlTextValues = 2
    } XlSpecialCellsValue;

    typedef enum {
        xlSubscribeToPicture = 0xffffefcd,
        xlSubscribeToText = 0xffffefc2
    } XlSubscribeToFormat;

    typedef enum {
        xlSummaryAbove = 0,
        xlSummaryBelow = 1
    } XlSummaryRow;

    typedef enum {
        xlSummaryOnLeft = 0xffffefdd,
        xlSummaryOnRight = 0xffffefc8
    } XlSummaryColumn;

    typedef enum {
        xlSummaryPivotTable = 0xffffefcc,
        xlStandardSummary = 1
    } XlSummaryReportType;

    typedef enum {
        xlTabPositionFirst = 0,
        xlTabPositionLast = 1
    } XlTabPosition;

    typedef enum {
        xlDelimited = 1,
        xlFixedWidth = 2
    } XlTextParsingType;

    typedef enum {
        xlTextQualifierDoubleQuote = 1,
        xlTextQualifierNone = 0xffffefd2,
        xlTextQualifierSingleQuote = 2
    } XlTextQualifier;

    typedef enum {
        xlWBATChart = 0xffffeff3,
        xlWBATExcel4IntlMacroSheet = 4,
        xlWBATExcel4MacroSheet = 3,
        xlWBATWorksheet = 0xffffefb9
    } XlWBATemplate;

    typedef enum {
        xlNormalView = 1,
        xlPageBreakPreview = 2
    } XlWindowView;

    typedef enum {
        xlCommand = 2,
        xlFunction = 1,
        xlNotXLM = 3
    } XlXLMMacroType;

    typedef enum {
        xlGuess = 0,
        xlNo = 2,
        xlYes = 1
    } XlYesNoGuess;

    typedef enum {
        xlInsideHorizontal = 12,
        xlInsideVertical = 11,
        xlDiagonalDown = 5,
        xlDiagonalUp = 6,
        xlEdgeBottom = 9,
        xlEdgeLeft = 7,
        xlEdgeRight = 10,
        xlEdgeTop = 8
    } XlBordersIndex;

    typedef enum {
        xlNoButtonChanges = 1,
        xlNoChanges = 4,
        xlNoDockingChanges = 3,
        xlToolbarProtectionNone = 0xffffefd1,
        xlNoShapeChanges = 2
    } XlToolbarProtection;

    typedef enum {
        xlDialogOpen = 1,
        xlDialogOpenLinks = 2,
        xlDialogSaveAs = 5,
        xlDialogFileDelete = 6,
        xlDialogPageSetup = 7,
        xlDialogPrint = 8,
        xlDialogPrinterSetup = 9,
        xlDialogArrangeAll = 12,
        xlDialogWindowSize = 13,
        xlDialogWindowMove = 14,
        xlDialogRun = 17,
        xlDialogSetPrintTitles = 23,
        xlDialogFont = 26,
        xlDialogDisplay = 27,
        xlDialogProtectDocument = 28,
        xlDialogCalculation = 32,
        xlDialogExtract = 35,
        xlDialogDataDelete = 36,
        xlDialogSort = 39,
        xlDialogDataSeries = 40,
        xlDialogTable = 41,
        xlDialogFormatNumber = 42,
        xlDialogAlignment = 43,
        xlDialogStyle = 44,
        xlDialogBorder = 45,
        xlDialogCellProtection = 46,
        xlDialogColumnWidth = 47,
        xlDialogClear = 52,
        xlDialogPasteSpecial = 53,
        xlDialogEditDelete = 54,
        xlDialogInsert = 55,
        xlDialogPasteNames = 58,
        xlDialogDefineName = 61,
        xlDialogCreateNames = 62,
        xlDialogFormulaGoto = 63,
        xlDialogFormulaFind = 64,
        xlDialogGalleryArea = 67,
        xlDialogGalleryBar = 68,
        xlDialogGalleryColumn = 69,
        xlDialogGalleryLine = 70,
        xlDialogGalleryPie = 71,
        xlDialogGalleryScatter = 72,
        xlDialogCombination = 73,
        xlDialogGridlines = 76,
        xlDialogAxes = 78,
        xlDialogAttachText = 80,
        xlDialogPatterns = 84,
        xlDialogMainChart = 85,
        xlDialogOverlay = 86,
        xlDialogScale = 87,
        xlDialogFormatLegend = 88,
        xlDialogFormatText = 89,
        xlDialogParse = 91,
        xlDialogUnhide = 94,
        xlDialogWorkspace = 95,
        xlDialogActivate = 103,
        xlDialogCopyPicture = 108,
        xlDialogDeleteName = 110,
        xlDialogDeleteFormat = 111,
        xlDialogNew = 119,
        xlDialogRowHeight = 127,
        xlDialogFormatMove = 128,
        xlDialogFormatSize = 129,
        xlDialogFormulaReplace = 130,
        xlDialogSelectSpecial = 132,
        xlDialogApplyNames = 133,
        xlDialogReplaceFont = 134,
        xlDialogSplit = 137,
        xlDialogOutline = 142,
        xlDialogSaveWorkbook = 145,
        xlDialogCopyChart = 147,
        xlDialogFormatFont = 150,
        xlDialogNote = 154,
        xlDialogSetUpdateStatus = 159,
        xlDialogColorPalette = 161,
        xlDialogChangeLink = 166,
        xlDialogAppMove = 170,
        xlDialogAppSize = 171,
        xlDialogMainChartType = 185,
        xlDialogOverlayChartType = 186,
        xlDialogOpenMail = 188,
        xlDialogSendMail = 189,
        xlDialogStandardFont = 190,
        xlDialogConsolidate = 191,
        xlDialogSortSpecial = 192,
        xlDialogGallery3dArea = 193,
        xlDialogGallery3dColumn = 194,
        xlDialogGallery3dLine = 195,
        xlDialogGallery3dPie = 196,
        xlDialogView3d = 197,
        xlDialogGoalSeek = 198,
        xlDialogWorkgroup = 199,
        xlDialogFillGroup = 200,
        xlDialogUpdateLink = 201,
        xlDialogPromote = 202,
        xlDialogDemote = 203,
        xlDialogShowDetail = 204,
        xlDialogObjectProperties = 207,
        xlDialogSaveNewObject = 208,
        xlDialogApplyStyle = 212,
        xlDialogAssignToObject = 213,
        xlDialogObjectProtection = 214,
        xlDialogCreatePublisher = 217,
        xlDialogSubscribeTo = 218,
        xlDialogShowToolbar = 220,
        xlDialogPrintPreview = 222,
        xlDialogEditColor = 223,
        xlDialogFormatMain = 225,
        xlDialogFormatOverlay = 226,
        xlDialogEditSeries = 228,
        xlDialogDefineStyle = 229,
        xlDialogGalleryRadar = 249,
        xlDialogEditionOptions = 251,
        xlDialogZoom = 256,
        xlDialogInsertObject = 259,
        xlDialogSize = 261,
        xlDialogMove = 262,
        xlDialogFormatAuto = 269,
        xlDialogGallery3dBar = 272,
        xlDialogGallery3dSurface = 273,
        xlDialogCustomizeToolbar = 276,
        xlDialogWorkbookAdd = 281,
        xlDialogWorkbookMove = 282,
        xlDialogWorkbookCopy = 283,
        xlDialogWorkbookOptions = 284,
        xlDialogSaveWorkspace = 285,
        xlDialogChartWizard = 288,
        xlDialogAssignToTool = 293,
        xlDialogPlacement = 300,
        xlDialogFillWorkgroup = 301,
        xlDialogWorkbookNew = 302,
        xlDialogScenarioCells = 305,
        xlDialogScenarioAdd = 307,
        xlDialogScenarioEdit = 308,
        xlDialogScenarioSummary = 311,
        xlDialogPivotTableWizard = 312,
        xlDialogPivotFieldProperties = 313,
        xlDialogOptionsCalculation = 318,
        xlDialogOptionsEdit = 319,
        xlDialogOptionsView = 320,
        xlDialogAddinManager = 321,
        xlDialogMenuEditor = 322,
        xlDialogAttachToolbars = 323,
        xlDialogOptionsChart = 325,
        xlDialogVbaInsertFile = 328,
        xlDialogVbaProcedureDefinition = 330,
        xlDialogRoutingSlip = 336,
        xlDialogMailLogon = 339,
        xlDialogInsertPicture = 342,
        xlDialogGalleryDoughnut = 344,
        xlDialogChartTrend = 350,
        xlDialogWorkbookInsert = 354,
        xlDialogOptionsTransition = 355,
        xlDialogOptionsGeneral = 356,
        xlDialogFilterAdvanced = 370,
        xlDialogMailNextLetter = 378,
        xlDialogDataLabel = 379,
        xlDialogInsertTitle = 380,
        xlDialogFontProperties = 381,
        xlDialogMacroOptions = 382,
        xlDialogWorkbookUnhide = 384,
        xlDialogWorkbookName = 386,
        xlDialogGalleryCustom = 388,
        xlDialogAddChartAutoformat = 390,
        xlDialogChartAddData = 392,
        xlDialogTabOrder = 394,
        xlDialogSubtotalCreate = 398,
        xlDialogWorkbookTabSplit = 415,
        xlDialogWorkbookProtect = 417,
        xlDialogScrollbarProperties = 420,
        xlDialogPivotShowPages = 421,
        xlDialogTextToColumns = 422,
        xlDialogFormatCharttype = 423,
        xlDialogPivotFieldGroup = 433,
        xlDialogPivotFieldUngroup = 434,
        xlDialogCheckboxProperties = 435,
        xlDialogLabelProperties = 436,
        xlDialogListboxProperties = 437,
        xlDialogEditboxProperties = 438,
        xlDialogOpenText = 441,
        xlDialogPushbuttonProperties = 445,
        xlDialogFilter = 447,
        xlDialogFunctionWizard = 450,
        xlDialogSaveCopyAs = 456,
        xlDialogOptionsListsAdd = 458,
        xlDialogSeriesAxes = 460,
        xlDialogSeriesX = 461,
        xlDialogSeriesY = 462,
        xlDialogErrorbarX = 463,
        xlDialogErrorbarY = 464,
        xlDialogFormatChart = 465,
        xlDialogSeriesOrder = 466,
        xlDialogMailEditMailer = 470,
        xlDialogStandardWidth = 472,
        xlDialogScenarioMerge = 473,
        xlDialogProperties = 474,
        xlDialogSummaryInfo = 474,
        xlDialogFindFile = 475,
        xlDialogActiveCellFont = 476,
        xlDialogVbaMakeAddin = 478,
        xlDialogFileSharing = 481,
        xlDialogAutoCorrect = 485,
        xlDialogCustomViews = 493,
        xlDialogInsertNameLabel = 496,
        xlDialogSeriesShape = 504,
        xlDialogChartOptionsDataLabels = 505,
        xlDialogChartOptionsDataTable = 506,
        xlDialogSetBackgroundPicture = 509,
        xlDialogDataValidation = 525,
        xlDialogChartType = 526,
        xlDialogChartLocation = 527,
        _xlDialogPhonetic = 538,
        xlDialogChartSourceData = 540,
        _xlDialogChartSourceData = 541,
        xlDialogSeriesOptions = 557,
        xlDialogPivotTableOptions = 567,
        xlDialogPivotSolveOrder = 568,
        xlDialogPivotCalculatedField = 570,
        xlDialogPivotCalculatedItem = 572,
        xlDialogConditionalFormatting = 583,
        xlDialogInsertHyperlink = 596,
        xlDialogProtectSharing = 620,
        xlDialogOptionsME = 647,
        xlDialogPublishAsWebPage = 653,
        xlDialogPhonetic = 656,
        xlDialogNewWebQuery = 667,
        xlDialogImportTextFile = 666,
        xlDialogExternalDataProperties = 530,
        xlDialogWebOptionsGeneral = 683,
        xlDialogWebOptionsFiles = 684,
        xlDialogWebOptionsPictures = 685,
        xlDialogWebOptionsEncoding = 686,
        xlDialogWebOptionsFonts = 687,
        xlDialogPivotClientServerSet = 689,
        xlDialogPropertyFields = 754,
        xlDialogSearch = 731,
        xlDialogEvaluateFormula = 709,
        xlDialogDataLabelMultiple = 723,
        xlDialogChartOptionsDataLabelMultiple = 724,
        xlDialogErrorChecking = 732,
        xlDialogWebOptionsBrowsers = 773
    } XlBuiltInDialog;

    typedef enum {
        xlPrompt = 0,
        xlConstant = 1,
        xlRange = 2
    } XlParameterType;

    typedef enum {
        xlParamTypeUnknown = 0,
        xlParamTypeChar = 1,
        xlParamTypeNumeric = 2,
        xlParamTypeDecimal = 3,
        xlParamTypeInteger = 4,
        xlParamTypeSmallInt = 5,
        xlParamTypeFloat = 6,
        xlParamTypeReal = 7,
        xlParamTypeDouble = 8,
        xlParamTypeVarChar = 12,
        xlParamTypeDate = 9,
        xlParamTypeTime = 10,
        xlParamTypeTimestamp = 11,
        xlParamTypeLongVarChar = 0xffffffff,
        xlParamTypeBinary = 0xfffffffe,
        xlParamTypeVarBinary = 0xfffffffd,
        xlParamTypeLongVarBinary = 0xfffffffc,
        xlParamTypeBigInt = 0xfffffffb,
        xlParamTypeTinyInt = 0xfffffffa,
        xlParamTypeBit = 0xfffffff9,
        xlParamTypeWChar = 0xfffffff8
    } XlParameterDataType;

    typedef enum {
        xlButtonControl = 0,
        xlCheckBox = 1,
        xlDropDown = 2,
        xlEditBox = 3,
        xlGroupBox = 4,
        xlLabel = 5,
        xlListBox = 6,
        xlOptionButton = 7,
        xlScrollBar = 8,
        xlSpinner = 9
    } XlFormControl;

    typedef enum {
        xlSourceWorkbook = 0,
        xlSourceSheet = 1,
        xlSourcePrintArea = 2,
        xlSourceAutoFilter = 3,
        xlSourceRange = 4,
        xlSourceChart = 5,
        xlSourcePivotTable = 6,
        xlSourceQuery = 7
    } XlSourceType;

    typedef enum {
        xlHtmlStatic = 0,
        xlHtmlCalc = 1,
        xlHtmlList = 2,
        xlHtmlChart = 3
    } XlHtmlType;

    typedef enum {
        xlReport1 = 0,
        xlReport2 = 1,
        xlReport3 = 2,
        xlReport4 = 3,
        xlReport5 = 4,
        xlReport6 = 5,
        xlReport7 = 6,
        xlReport8 = 7,
        xlReport9 = 8,
        xlReport10 = 9,
        xlTable1 = 10,
        xlTable2 = 11,
        xlTable3 = 12,
        xlTable4 = 13,
        xlTable5 = 14,
        xlTable6 = 15,
        xlTable7 = 16,
        xlTable8 = 17,
        xlTable9 = 18,
        xlTable10 = 19,
        xlPTClassic = 20,
        xlPTNone = 21
    } XlPivotFormatType;

    typedef enum {
        xlCmdCube = 1,
        xlCmdSql = 2,
        xlCmdTable = 3,
        xlCmdDefault = 4
    } XlCmdType;

    typedef enum {
        xlGeneralFormat = 1,
        xlTextFormat = 2,
        xlMDYFormat = 3,
        xlDMYFormat = 4,
        xlYMDFormat = 5,
        xlMYDFormat = 6,
        xlDYMFormat = 7,
        xlYDMFormat = 8,
        xlSkipColumn = 9,
        xlEMDFormat = 10
    } XlColumnDataType;

    typedef enum {
        xlODBCQuery = 1,
        xlDAORecordset = 2,
        xlWebQuery = 4,
        xlOLEDBQuery = 5,
        xlTextImport = 6,
        xlADORecordset = 7
    } XlQueryType;

    typedef enum {
        xlEntirePage = 1,
        xlAllTables = 2,
        xlSpecifiedTables = 3
    } XlWebSelectionType;

    typedef enum {
        xlHierarchy = 1,
        xlMeasure = 2,
        xlSet = 3
    } XlCubeFieldType;

    typedef enum {
        xlWebFormattingAll = 1,
        xlWebFormattingRTF = 2,
        xlWebFormattingNone = 3
    } XlWebFormatting;

    typedef enum {
        xlDisplayShapes = 0xffffeff8,
        xlHide = 3,
        xlPlaceholders = 2
    } XlDisplayDrawingObjects;

    typedef enum {
        xlAtTop = 1,
        xlAtBottom = 2
    } XlSubtototalLocationType;

    typedef enum {
        xlPivotTableVersion2000 = 0,
        xlPivotTableVersion10 = 1,
        xlPivotTableVersionCurrent = 0xffffffff
    } XlPivotTableVersionList;

    typedef enum {
        xlPrintErrorsDisplayed = 0,
        xlPrintErrorsBlank = 1,
        xlPrintErrorsDash = 2,
        xlPrintErrorsNA = 3
    } XlPrintErrors;

    typedef enum {
        xlPivotCellValue = 0,
        xlPivotCellPivotItem = 1,
        xlPivotCellSubtotal = 2,
        xlPivotCellGrandTotal = 3,
        xlPivotCellDataField = 4,
        xlPivotCellPivotField = 5,
        xlPivotCellPageFieldItem = 6,
        xlPivotCellCustomSubtotal = 7,
        xlPivotCellDataPivotField = 8,
        xlPivotCellBlankCell = 9
    } XlPivotCellType;

    typedef enum {
        xlMissingItemsDefault = 0xffffffff,
        xlMissingItemsNone = 0,
        xlMissingItemsMax = 32500
    } XlPivotTableMissingItems;

    typedef enum {
        xlDone = 0,
        xlCalculating = 1,
        xlPending = 2
    } XlCalculationState;

    typedef enum {
        xlNoKey = 0,
        xlEscKey = 1,
        xlAnyKey = 2
    } XlCalculationInterruptKey;

    typedef enum {
        xlSortNormal = 0,
        xlSortTextAsNumbers = 1
    } XlSortDataOption;

    typedef enum {
        xlUpdateLinksUserSetting = 1,
        xlUpdateLinksNever = 2,
        xlUpdateLinksAlways = 3
    } XlUpdateLinks;

    typedef enum {
        xlLinkStatusOK = 0,
        xlLinkStatusMissingFile = 1,
        xlLinkStatusMissingSheet = 2,
        xlLinkStatusOld = 3,
        xlLinkStatusSourceNotCalculated = 4,
        xlLinkStatusIndeterminate = 5,
        xlLinkStatusNotStarted = 6,
        xlLinkStatusInvalidName = 7,
        xlLinkStatusSourceNotOpen = 8,
        xlLinkStatusSourceOpen = 9,
        xlLinkStatusCopiedValues = 10
    } XlLinkStatus;

    typedef enum {
        xlWithinSheet = 1,
        xlWithinWorkbook = 2
    } XlSearchWithin;

    typedef enum {
        xlNormalLoad = 0,
        xlRepairFile = 1,
        xlExtractData = 2
    } XlCorruptLoad;

    typedef enum {
        xlAsRequired = 0,
        xlAlways = 1,
        xlNever = 2
    } XlRobustConnect;

    typedef enum {
        xlEvaluateToError = 1,
        xlTextDate = 2,
        xlNumberAsText = 3,
        xlInconsistentFormula = 4,
        xlOmittedCells = 5,
        xlUnlockedFormulaCells = 6,
        xlEmptyCellReferences = 7
    } XlErrorChecks;

    typedef enum {
        xlDataLabelSeparatorDefault = 1
    } XlDataLabelSeparator;

    typedef enum {
        xlIndicatorAndButton = 0,
        xlDisplayNone = 1,
        xlButtonOnly = 2
    } XlSmartTagDisplayMode;

    typedef enum {
        xlRangeValueDefault = 10,
        xlRangeValueXMLSpreadsheet = 11,
        xlRangeValueMSPersistXML = 12
    } XlRangeValueDataType;

    typedef enum {
        xlSpeakByRows = 0,
        xlSpeakByColumns = 1
    } XlSpeakDirection;

    typedef enum {
        xlFormatFromLeftOrAbove = 0,
        xlFormatFromRightOrBelow = 1
    } XlInsertFormatOrigin;

    typedef enum {
        xlArabicNone = 0,
        xlArabicStrictAlefHamza = 1,
        xlArabicStrictFinalYaa = 2,
        xlArabicBothStrict = 3
    } XlArabicModes;

    typedef enum {
        xlQueryTable = 0,
        xlPivotTableReport = 1
    } XlImportDataAs;

    typedef enum {
        xlCalculatedMember = 0,
        xlCalculatedSet = 1
    } XlCalculatedMemberType;

    typedef enum {
        xlHebrewFullScript = 0,
        xlHebrewPartialScript = 1,
        xlHebrewMixedScript = 2,
        xlHebrewMixedAuthorizedScript = 3
    } XlHebrewModes;
