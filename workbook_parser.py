#-------------------------------------------------------------------------------
# Name: xls_macro4_deobfuscator
# Purpose: To deobfuscate xls macro4 commands used by malware
#
# Author:      sganachari
#
# Created:     20/04/2020
# Copyright:   (c) sganachari 2020
# Licence:     <your licence>
#-------------------------------------------------------------------------------

import os
import sys
import struct
import string
import re
import olefile

Debug = False
XLS_columns = string.ascii_uppercase
XLS_Record = {"Formula":6,"EOF":10,"CalcCount":12,"CalcMode":13,"CalcPrecision":14,"CalcRefMode":15,"CalcDelta":16,"CalcIter":17,"Protect":18,"Password":19,"Header":20,"Footer":21,"ExternSheet":23,"Lbl":24,"WinProtect":25,"VerticalPageBreaks":26,"HorizontalPageBreaks":27,"Note":28,"Selection":29,"Date1904":34,"ExternName":35,"LeftMargin":38,"RightMargin":39,"TopMargin":40,"BottomMargin":41,"PrintRowCol":42,"PrintGrid":43,"FilePass":47,"Font":49,"PrintSize":51,"Continue":60,"Window1":61,"Backup":64,"Pane":65,"CodePage":66,"Pls":77,"DCon":80,"DConRef":81,"DConName":82,"DefColWidth":85,"XCT":89,"CRN":90,"FileSharing":91,"WriteAccess":92,"Obj":93,"Uncalced":94,"CalcSaveRecalc":95,"Template":96,"Intl":97,"ObjProtect":99,"ColInfo":125,"Guts":128,"WsBool":129,"GridSet":130,"HCenter":131,"VCenter":132,"BoundSheet8":133,"WriteProtect":134,"Country":140,"HideObj":141,"Sort":144,"Palette":146,"Sync":151,"LPr":152,"DxGCol":153,"FnGroupName":154,"FilterMode":155,"BuiltInFnGroupCount":156,"AutoFilterInfo":157,"AutoFilter":158,"Scl":160,"Setup":161,"ScenMan":174,"SCENARIO":175,"SxView":176,"Sxvd":177,"SXVI":178,"SxIvd":180,"SXLI":181,"SXPI":182,"DocRoute":184,"RecipName":185,"MulRk":189,"MulBlank":190,"Mms":193,"SXDI":197,"SXDB":198,"SXFDB":199,"SXDBB":200,"SXNum":201,"SxBool":202,"SxErr":203,"SXInt":204,"SXString":205,"SXDtr":206,"SxNil":207,"SXTbl":208,"SXTBRGIITM":209,"SxTbpg":210,"ObProj":211,"SXStreamID":213,"DBCell":215,"SXRng":216,"SxIsxoper":217,"BookBool":218,"DbOrParamQry":220,"ScenarioProtect":221,"OleObjectSize":222,"XF":224,"InterfaceHdr":225,"InterfaceEnd":226,"SXVS":227,"MergeCells":229,"BkHim":233,"MsoDrawingGroup":235,"MsoDrawing":236,"MsoDrawingSelection":237,"PhoneticInfo":239,"SxRule":240,"SXEx":241,"SxFilt":242,"SxDXF":244,"SxItm":245,"SxName":246,"SxSelect":247,"SXPair":248,"SxFmla":249,"SxFormat":251,"SST":252,"LabelSst":253,"ExtSST":255,"SXVDEx":256,"SXFormula":259,"SXDBEx":290,"RRDInsDel":311,"RRDHead":312,"RRDChgCell":315,"RRTabId":317,"RRDRenSheet":318,"RRSort":319,"RRDMove":320,"RRFormat":330,"RRAutoFmt":331,"RRInsertSh":333,"RRDMoveBegin":334,"RRDMoveEnd":335,"RRDInsDelBegin":336,"RRDInsDelEnd":337,"RRDConflict":338,"RRDDefName":339,"RRDRstEtxp":340,"LRng":351,"UsesELFs":352,"DSF":353,"CUsr":401,"CbUsr":402,"UsrInfo":403,"UsrExcl":404,"FileLock":405,"RRDInfo":406,"BCUsrs":407,"UsrChk":408,"UserBView":425,"UserSViewBegin":426,"UserSViewBegin_Chart":426,"UserSViewEnd":427,"RRDUserView":428,"Qsi":429,"SupBook":430,"Prot4Rev":431,"CondFmt":432,"CF":433,"DVal":434,"DConBin":437,"TxO":438,"RefreshAll":439,"HLink":440,"Lel":441,"CodeName":442,"SXFDBType":443,"Prot4RevPass":444,"ObNoMacros":445,"Dv":446,"Excel9File":448,"RecalcId":449,"EntExU2":450,"Dimensions":512,"Blank":513,"Number":515,"Label":516,"BoolErr":517,"String":519,"Row":520,"Index":523,"Array":545,"DefaultRowHeight":549,"Table":566,"Window2":574,"RK":638,"Style":659,"BigName":1048,"Format":1054,"ContinueBigName":1084,"ShrFmla":1212,"HLinkTooltip":2048,"WebPub":2049,"QsiSXTag":2050,"DBQueryExt":2051,"ExtString":2052,"TxtQry":2053,"Qsir":2054,"Qsif":2055,"RRDTQSIF":2056,"BOF":2057,"OleDbConn":2058,"WOpt":2059,"SXViewEx":2060,"SXTH":2061,"SXPIEx":2062,"SXVDTEx":2063,"SXViewEx9":2064,"ContinueFrt":2066,"RealTimeData":2067,"ChartFrtInfo":2128,"FrtWrapper":2129,"StartBlock":2130,"EndBlock":2131,"StartObject":2132,"EndObject":2133,"CatLab":2134,"YMult":2135,"SXViewLink":2136,"PivotChartBits":2137,"FrtFontList":2138,"SheetExt":2146,"BookExt":2147,"SXAddl":2148,"CrErr":2149,"HFPicture":2150,"FeatHdr":2151,"Feat":2152,"DataLabExt":2154,"DataLabExtContents":2155,"CellWatch":2156,"FeatHdr11":2161,"Feature11":2162,"DropDownObjIds":2164,"ContinueFrt11":2165,"DConn":2166,"List12":2167,"Feature12":2168,"CondFmt12":2169,"CF12":2170,"CFEx":2171,"XFCRC":2172,"XFExt":2173,"AutoFilter12":2174,"ContinueFrt12":2175,"MDTInfo":2180,"MDXStr":2181,"MDXTuple":2182,"MDXSet":2183,"MDXProp":2184,"MDXKPI":2185,"MDB":2186,"PLV":2187,"Compat12":2188,"DXF":2189,"TableStyles":2190,"TableStyle":2191,"TableStyleElement":2192,"StyleExt":2194,"NamePublish":2195,"NameCmt":2196,"SortData":2197,"Theme":2198,"GUIDTypeLib":2199,"FnGrp12":2200,"NameFnGrp12":2201,"MTRSettings":2202,"CompressPictures":2203,"HeaderFooter":2204,"CrtLayout12":2205,"CrtMlFrt":2206,"CrtMlFrtContinue":2207,"ForceFullCalculation":2211,"ShapePropsStream":2212,"TextPropsStream":2213,"RichTextStream":2214,"CrtLayout12A":2215,"Units":4097,"Chart":4098,"Series":4099,"DataFormat":4102,"LineFormat":4103,"MarkerFormat":4105,"AreaFormat":4106,"PieFormat":4107,"AttachedLabel":4108,"SeriesText":4109,"ChartFormat":4116,"Legend":4117,"SeriesList":4118,"Bar":4119,"Line":4120,"Pie":4121,"Area":4122,"Scatter":4123,"CrtLine":4124,"Axis":4125,"Tick":4126,"ValueRange":4127,"CatSerRange":4128,"AxisLine":4129,"CrtLink":4130,"DefaultText":4132,"Text":4133,"FontX":4134,"ObjectLink":4135,"Frame":4146,"Begin":4147,"End":4148,"PlotArea":4149,"Chart3d":4154,"PicF":4156,"DropBar":4157,"Radar":4158,"Surf":4159,"RadarArea":4160,"AxisParent":4161,"LegendExceptionsection":4163,"ShtProps":4164,"SerToCrt":4165,"AxesUsed":4166,"SBaseRef":4168,"SerParent":4170,"SerAuxTrend":4171,"IFmtRecord":4174,"Pos":4175,"AlRuns":4176,"BRAI":4177,"SerAuxErrBar":4187,"ClrtClient":4188,"SerFmt":4189,"Chart3DBarShape":4191,"Fbi":4192,"BopPop":4193,"AxcExt":4194,"Dat":4195,"PlotGrowth":4196,"SIIndex":4197,"GelFrame":4198,"BopPopCustom":4199,"Fbi2":4200}
ctab_functions = {0x0000:"BEEP",0x0001:"OPEN",0x0002:"OPEN.LINKS",0x0003:"CLOSE.ALL",0x0004:"SAVE",0x0005:"SAVE.AS",0x0006:"FILE.DELETE",0x0007:"PAGE.SETUP",0x0008:"PRINT",0x0009:"PRINTER.SETUP",0x000A:"QUIT",0x000B:"NEW.WINDOW",0x000C:"ARRANGE.ALL",0x000D:"WINDOW.SIZE",0x000E:"WINDOW.MOVE",0x000F:"FULL",0x0010:"CLOSE",0x0011:"RUN",0x0016:"SET.PRINT.AREA",0x0017:"SET.PRINT.TITLES",0x0018:"SET.PAGE.BREAK",0x0019:"REMOVE.PAGE.BREAK",0x001A:"FONT",0x001B:"DISPLAY",0x001C:"PROTECT.DOCUMENT",0x001D:"PRECISION",0x001E:"A1.R1C1",0x001F:"CALCULATE.NOW",0x0020:"CALCULATION",0x0022:"DATA.FIND",0x0023:"EXTRACT",0x0024:"DATA.DELETE",0x0025:"SET.DATABASE",0x0026:"SET.CRITERIA",0x0027:"SORT",0x0028:"DATA.SERIES",0x0029:"TABLE",0x002A:"FORMAT.NUMBER",0x002B:"ALIGNMENT",0x002C:"STYLE",0x002D:"BORDER",0x002E:"CELL.PROTECTION",0x002F:"COLUMN.WIDTH",0x0030:"UNDO",0x0031:"CUT",0x0032:"COPY",0x0033:"PASTE",0x0034:"CLEAR",0x0035:"PASTE.SPECIAL",0x0036:"EDIT.DELETE",0x0037:"INSERT",0x0038:"FILL.RIGHT",0x0039:"FILL.DOWN",0x003D:"DEFINE.NAME",0x003E:"CREATE.NAMES",0x003F:"FORMULA.GOTO",0x0040:"FORMULA.FIND",0x0041:"SELECT.LAST.CELL",0x0042:"SHOW.ACTIVE.CELL",0x0043:"GALLERY.AREA",0x0044:"GALLERY.BAR",0x0045:"GALLERY.COLUMN",0x0046:"GALLERY.LINE",0x0047:"GALLERY.PIE",0x0048:"GALLERY.SCATTER",0x0049:"COMBINATION",0x004A:"PREFERRED",0x004B:"ADD.OVERLAY",0x004C:"GRIDLINES",0x004D:"SET.PREFERRED",0x004E:"AXES",0x004F:"LEGEND",0x0050:"ATTACH.TEXT",0x0051:"ADD.ARROW",0x0052:"SELECT.CHART",0x0053:"SELECT.PLOT.AREA",0x0054:"PATTERNS",0x0055:"MAIN.CHART",0x0056:"OVERLAY",0x0057:"SCALE",0x0058:"FORMAT.LEGEND",0x0059:"FORMAT.TEXT",0x005A:"EDIT.REPEAT",0x005B:"PARSE",0x005C:"JUSTIFY",0x005D:"HIDE",0x005E:"UNHIDE",0x005F:"WORKSPACE",0x0060:"FORMULA",0x0061:"FORMULA.FILL",0x0062:"FORMULA.ARRAY",0x0063:"DATA.FIND.NEXT",0x0064:"DATA.FIND.PREV",0x0065:"FORMULA.FIND.NEXT",0x0066:"FORMULA.FIND.PREV",0x0067:"ACTIVATE",0x0068:"ACTIVATE.NEXT",0x0069:"ACTIVATE.PREV",0x006A:"UNLOCKED.NEXT",0x006B:"UNLOCKED.PREV",0x006C:"COPY.PICTURE",0x006D:"SELECT",0x006E:"DELETE.NAME",0x006F:"DELETE.FORMAT",0x0070:"VLINE",0x0071:"HLINE",0x0072:"VPAGE",0x0073:"HPAGE",0x0074:"VSCROLL",0x0075:"HSCROLL",0x0076:"ALERT",0x0077:"NEW",0x0078:"CANCEL.COPY",0x0079:"SHOW.CLIPBOARD",0x007A:"MESSAGE",0x007C:"PASTE.LINK",0x007D:"APP.ACTIVATE",0x007E:"DELETE.ARROW",0x007F:"ROW.HEIGHT",0x0080:"FORMAT.MOVE",0x0081:"FORMAT.SIZE",0x0082:"FORMULA.REPLACE",0x0083:"SEND.KEYS",0x0084:"SELECT.SPECIAL",0x0085:"APPLY.NAMES",0x0086:"REPLACE.FONT",0x0087:"FREEZE.PANES",0x0088:"SHOW.INFO",0x0089:"SPLIT",0x008A:"ON.WINDOW",0x008B:"ON.DATA",0x008C:"DISABLE.INPUT",0x008E:"OUTLINE",0x008F:"LIST.NAMES",0x0090:"FILE.CLOSE",0x0091:"SAVE.WORKBOOK",0x0092:"DATA.FORM",0x0093:"COPY.CHART",0x0094:"ON.TIME",0x0095:"WAIT",0x0096:"FORMAT.FONT",0x0097:"FILL.UP",0x0098:"FILL.LEFT",0x0099:"DELETE.OVERLAY",0x009B:"SHORT.MENUS",0x009F:"SET.UPDATE.STATUS",0x00A1:"COLOR.PALETTE",0x00A2:"DELETE.STYLE",0x00A3:"WINDOW.RESTORE",0x00A4:"WINDOW.MAXIMIZE",0x00A6:"CHANGE.LINK",0x00A7:"CALCULATE.DOCUMENT",0x00A8:"ON.KEY",0x00A9:"APP.RESTORE",0x00AA:"APP.MOVE",0x00AB:"APP.SIZE",0x00AC:"APP.MINIMIZE",0x00AD:"APP.MAXIMIZE",0x00AE:"BRING.TO.FRONT",0x00AF:"SEND.TO.BACK",0x00B9:"MAIN.CHART.TYPE",0x00BA:"OVERLAY.CHART.TYPE",0x00BB:"SELECT.END",0x00BC:"OPEN.MAIL",0x00BD:"SEND.MAIL",0x00BE:"STANDARD.FONT",0x00BF:"CONSOLIDATE",0x00C0:"SORT.SPECIAL",0x00C1:"GALLERY.3D.AREA",0x00C2:"GALLERY.3D.COLUMN",0x00C3:"GALLERY.3D.LINE",0x00C4:"GALLERY.3D.PIE",0x00C5:"VIEW.3D",0x00C6:"GOAL.SEEK",0x00C7:"WORKGROUP",0x00C8:"FILL.GROUP",0x00C9:"UPDATE.LINK",0x00CA:"PROMOTE",0x00CB:"DEMOTE",0x00CC:"SHOW.DETAIL",0x00CE:"UNGROUP",0x00CF:"OBJECT.PROPERTIES",0x00D0:"SAVE.NEW.OBJECT",0x00D1:"SHARE",0x00D2:"SHARE.NAME",0x00D3:"DUPLICATE",0x00D4:"APPLY.STYLE",0x00D5:"ASSIGN.TO.OBJECT",0x00D6:"OBJECT.PROTECTION",0x00D7:"HIDE.OBJECT",0x00D8:"SET.EXTRACT",0x00D9:"CREATE.PUBLISHER",0x00DA:"SUBSCRIBE.TO",0x00DB:"ATTRIBUTES",0x00DC:"SHOW.TOOLBAR",0x00DE:"PRINT.PREVIEW",0x00DF:"EDIT.COLOR",0x00E0:"SHOW.LEVELS",0x00E1:"FORMAT.MAIN",0x00E2:"FORMAT.OVERLAY",0x00E3:"ON.RECALC",0x00E4:"EDIT.SERIES",0x00E5:"DEFINE.STYLE",0x00F0:"LINE.PRINT",0x00F3:"ENTER.DATA",0x00F9:"GALLERY.RADAR",0x00FA:"MERGE.STYLES",0x00FB:"EDITION.OPTIONS",0x00FC:"PASTE.PICTURE",0x00FD:"PASTE.PICTURE.LINK",0x00FE:"SPELLING",0x0100:"ZOOM",0x0103:"INSERT.OBJECT",0x0104:"WINDOW.MINIMIZE",0x0109:"SOUND.NOTE",0x010A:"SOUND.PLAY",0x010B:"FORMAT.SHAPE",0x010C:"EXTEND.POLYGON",0x010D:"FORMAT.AUTO",0x0110:"GALLERY.3D.BAR",0x0111:"GALLERY.3D.SURFACE",0x0112:"FILL.AUTO",0x0114:"CUSTOMIZE.TOOLBAR",0x0115:"ADD.TOOL",0x0116:"EDIT.OBJECT",0x0117:"ON.DOUBLECLICK",0x0118:"ON.ENTRY",0x0119:"WORKBOOK.ADD",0x011A:"WORKBOOK.MOVE",0x011B:"WORKBOOK.COPY",0x011C:"WORKBOOK.OPTIONS",0x011D:"SAVE.WORKSPACE",0x0120:"CHART.WIZARD",0x0121:"DELETE.TOOL",0x0122:"MOVE.TOOL",0x0123:"WORKBOOK.SELECT",0x0124:"WORKBOOK.ACTIVATE",0x0125:"ASSIGN.TO.TOOL",0x0127:"COPY.TOOL",0x0128:"RESET.TOOL",0x0129:"CONSTRAIN.NUMERIC",0x012A:"PASTE.TOOL",0x012E:"WORKBOOK.NEW",0x0131:"SCENARIO.CELLS",0x0132:"SCENARIO.DELETE",0x0133:"SCENARIO.ADD",0x0134:"SCENARIO.EDIT",0x0135:"SCENARIO.SHOW",0x0136:"SCENARIO.SHOW.NEXT",0x0137:"SCENARIO.SUMMARY",0x0138:"PIVOT.TABLE.WIZARD",0x0139:"PIVOT.FIELD.PROPERTIES",0x013A:"PIVOT.FIELD",0x013B:"PIVOT.ITEM",0x013C:"PIVOT.ADD.FIELDS",0x013E:"OPTIONS.CALCULATION",0x013F:"OPTIONS.EDIT",0x0140:"OPTIONS.VIEW",0x0141:"ADDIN.MANAGER",0x0142:"MENU.EDITOR",0x0143:"ATTACH.TOOLBARS",0x0144:"VBAActivate",0x0145:"OPTIONS.CHART",0x0148:"VBA.INSERT.FILE",0x014A:"VBA.PROCEDURE.DEFINITION",0x0150:"ROUTING.SLIP",0x0152:"ROUTE.DOCUMENT",0x0153:"MAIL.LOGON",0x0156:"INSERT.PICTURE",0x0157:"EDIT.TOOL",0x0158:"GALLERY.DOUGHNUT",0x015E:"CHART.TREND",0x0160:"PIVOT.ITEM.PROPERTIES",0x0162:"WORKBOOK.INSERT",0x0163:"OPTIONS.TRANSITION",0x0164:"OPTIONS.GENERAL",0x0172:"FILTER.ADVANCED",0x0175:"MAIL.ADD.MAILER",0x0176:"MAIL.DELETE.MAILER",0x0177:"MAIL.REPLY",0x0178:"MAIL.REPLY.ALL",0x0179:"MAIL.FORWARD",0x017A:"MAIL.NEXT.LETTER",0x017B:"DATA.LABEL",0x017C:"INSERT.TITLE",0x017D:"FONT.PROPERTIES",0x017E:"MACRO.OPTIONS",0x017F:"WORKBOOK.HIDE",0x0180:"WORKBOOK.UNHIDE",0x0181:"WORKBOOK.DELETE",0x0182:"WORKBOOK.NAME",0x0184:"GALLERY.CUSTOM",0x0186:"ADD.CHART.AUTOFORMAT",0x0187:"DELETE.CHART.AUTOFORMAT",0x0188:"CHART.ADD.DATA",0x0189:"AUTO.OUTLINE",0x018A:"TAB.ORDER",0x018B:"SHOW.DIALOG",0x018C:"SELECT.ALL",0x018D:"UNGROUP.SHEETS",0x018E:"SUBTOTAL.CREATE",0x018F:"SUBTOTAL.REMOVE",0x0190:"RENAME.OBJECT",0x019C:"WORKBOOK.SCROLL",0x019D:"WORKBOOK.NEXT",0x019E:"WORKBOOK.PREV",0x019F:"WORKBOOK.TAB.SPLIT",0x01A0:"FULL.SCREEN",0x01A1:"WORKBOOK.PROTECT",0x01A4:"SCROLLBAR.PROPERTIES",0x01A5:"PIVOT.SHOW.PAGES",0x01A6:"TEXT.TO.COLUMNS",0x01A7:"FORMAT.CHARTTYPE",0x01A8:"LINK.FORMAT",0x01A9:"TRACER.DISPLAY",0x01AE:"TRACER.NAVIGATE",0x01AF:"TRACER.CLEAR",0x01B0:"TRACER.ERROR",0x01B1:"PIVOT.FIELD.GROUP",0x01B2:"PIVOT.FIELD.UNGROUP",0x01B3:"CHECKBOX.PROPERTIES",0x01B4:"LABEL.PROPERTIES",0x01B5:"LISTBOX.PROPERTIES",0x01B6:"EDITBOX.PROPERTIES",0x01B7:"PIVOT.REFRESH",0x01B8:"LINK.COMBO",0x01B9:"OPEN.TEXT",0x01BA:"HIDE.DIALOG",0x01BB:"SET.DIALOG.FOCUS",0x01BC:"ENABLE.OBJECT",0x01BD:"PUSHBUTTON.PROPERTIES",0x01BE:"SET.DIALOG.DEFAULT",0x01BF:"FILTER",0x01C0:"FILTER.SHOW.ALL",0x01C1:"CLEAR.OUTLINE",0x01C2:"FUNCTION.WIZARD",0x01C3:"ADD.LIST.ITEM",0x01C4:"SET.LIST.ITEM",0x01C5:"REMOVE.LIST.ITEM",0x01C6:"SELECT.LIST.ITEM",0x01C7:"SET.CONTROL.VALUE",0x01C8:"SAVE.COPY.AS",0x01CA:"OPTIONS.LISTS.ADD",0x01CB:"OPTIONS.LISTS.DELETE",0x01CC:"SERIES.AXES",0x01CD:"SERIES.X",0x01CE:"SERIES.Y",0x01CF:"ERRORBAR.X",0x01D0:"ERRORBAR.Y",0x01D1:"FORMAT.CHART",0x01D2:"SERIES.ORDER",0x01D3:"MAIL.LOGOFF",0x01D4:"CLEAR.ROUTING.SLIP",0x01D5:"APP.ACTIVATE.MICROSOFT",0x01D6:"MAIL.EDIT.MAILER",0x01D7:"ON.SHEET",0x01D8:"STANDARD.WIDTH",0x01D9:"SCENARIO.MERGE",0x01DA:"SUMMARY.INFO",0x01DB:"FIND.FILE",0x01DC:"ACTIVE.CELL.FONT",0x01DD:"ENABLE.TIPWIZARD",0x01DE:"VBA.MAKE.ADDIN",0x01E0:"INSERTDATATABLE",0x01E1:"WORKGROUP.OPTIONS",0x01E2:"MAIL.SEND.MAILER",0x01E5:"AUTOCORRECT",0x01E9:"POST.DOCUMENT",0x01EB:"PICKLIST",0x01ED:"VIEW.SHOW",0x01EE:"VIEW.DEFINE",0x01EF:"VIEW.DELETE",0x01FD:"SHEET.BACKGROUND",0x01FE:"INSERT.MAP.OBJECT",0x01FF:"OPTIONS.MENONO",0x0205:"MSOCHECKS",0x0206:"NORMAL",0x0207:"LAYOUT",0x0208:"RM.PRINT.AREA",0x0209:"CLEAR.PRINT.AREA",0x020A:"ADD.PRINT.AREA",0x020B:"MOVE.BRK",0x0221:"HIDECURR.NOTE",0x0222:"HIDEALL.NOTES",0x0223:"DELETE.NOTE",0x0224:"TRAVERSE.NOTES",0x0225:"ACTIVATE.NOTES",0x026C:"PROTECT.REVISIONS",0x026D:"UNPROTECT.REVISIONS",0x0287:"OPTIONS.ME",0x028D:"WEB.PUBLISH",0x029B:"NEWWEBQUERY",0x02A1:"PIVOT.TABLE.CHART",0x02F1:"OPTIONS.SAVE",0x02F3:"OPTIONS.SPELL",0x0328:"HIDEALL.INKANNOTS"}
ftab_functions = {0x0000:"COUNT",0x0001:"IF",0x0002:"ISNA",0x0003:"ISERROR",0x0004:"SUM",0x0005:"AVERAGE",0x0006:"MIN",0x0007:"MAX",0x0008:"ROW",0x0009:"COLUMN",0x000A:"NA",0x000B:"NPV",0x000C:"STDEV",0x000D:"DOLLAR",0x000E:"FIXED",0x000F:"SIN",0x0010:"COS",0x0011:"TAN",0x0012:"ATAN",0x0013:"PI",0x0014:"SQRT",0x0015:"EXP",0x0016:"LN",0x0017:"LOG10",0x0018:"ABS",0x0019:"INT",0x001A:"SIGN",0x001B:"ROUND",0x001C:"LOOKUP",0x001D:"INDEX",0x001E:"REPT",0x001F:"MID",0x0020:"LEN",0x0021:"VALUE",0x0022:"TRUE",0x0023:"FALSE",0x0024:"AND",0x0025:"OR",0x0026:"NOT",0x0027:"MOD",0x0028:"DCOUNT",0x0029:"DSUM",0x002A:"DAVERAGE",0x002B:"DMIN",0x002C:"DMAX",0x002D:"DSTDEV",0x002E:"VAR",0x002F:"DVAR",0x0030:"TEXT",0x0031:"LINEST",0x0032:"TREND",0x0033:"LOGEST",0x0034:"GROWTH",0x0035:"GOTO",0x0036:"HALT",0x0037:"RETURN",0x0038:"PV",0x0039:"FV",0x003A:"NPER",0x003B:"PMT",0x003C:"RATE",0x003D:"MIRR",0x003E:"IRR",0x003F:"RAND",0x0040:"MATCH",0x0041:"DATE",0x0042:"TIME",0x0043:"DAY",0x0044:"MONTH",0x0045:"YEAR",0x0046:"WEEKDAY",0x0047:"HOUR",0x0048:"MINUTE",0x0049:"SECOND",0x004A:"NOW",0x004B:"AREAS",0x004C:"ROWS",0x004D:"COLUMNS",0x004E:"OFFSET",0x004F:"ABSREF",0x0050:"RELREF",0x0051:"ARGUMENT",0x0052:"SEARCH",0x0053:"TRANSPOSE",0x0054:"ERROR",0x0055:"STEP",0x0056:"TYPE",0x0057:"ECHO",0x0058:"SET.NAME",0x0059:"CALLER",0x005A:"DEREF",0x005B:"WINDOWS",0x005C:"SERIES",0x005D:"DOCUMENTS",0x005E:"ACTIVE.CELL",0x005F:"SELECTION",0x0060:"RESULT",0x0061:"ATAN2",0x0062:"ASIN",0x0063:"ACOS",0x0064:"CHOOSE",0x0065:"HLOOKUP",0x0066:"VLOOKUP",0x0067:"LINKS",0x0068:"INPUT",0x0069:"ISREF",0x006A:"GET.FORMULA",0x006B:"GET.NAME",0x006C:"SET.VALUE",0x006D:"LOG",0x006E:"EXEC",0x006F:"CHAR",0x0070:"LOWER",0x0071:"UPPER",0x0072:"PROPER",0x0073:"LEFT",0x0074:"RIGHT",0x0075:"EXACT",0x0076:"TRIM",0x0077:"REPLACE",0x0078:"SUBSTITUTE",0x0079:"CODE",0x007A:"NAMES",0x007B:"DIRECTORY",0x007C:"FIND",0x007D:"CELL",0x007E:"ISERR",0x007F:"ISTEXT",0x0080:"ISNUMBER",0x0081:"ISBLANK",0x0082:"T",0x0083:"N",0x0084:"FOPEN",0x0085:"FCLOSE",0x0086:"FSIZE",0x0087:"FREADLN",0x0088:"FREAD",0x0089:"FWRITELN",0x008A:"FWRITE",0x008B:"FPOS",0x008C:"DATEVALUE",0x008D:"TIMEVALUE",0x008E:"SLN",0x008F:"SYD",0x0090:"DDB",0x0091:"GET.DEF",0x0092:"REFTEXT",0x0093:"TEXTREF",0x0094:"INDIRECT",0x0095:"REGISTER",0x0096:"CALL",0x0097:"ADD.BAR",0x0098:"ADD.MENU",0x0099:"ADD.COMMAND",0x009A:"ENABLE.COMMAND",0x009B:"CHECK.COMMAND",0x009C:"RENAME.COMMAND",0x009D:"SHOW.BAR",0x009E:"DELETE.MENU",0x009F:"DELETE.COMMAND",0x00A0:"GET.CHART.ITEM",0x00A1:"DIALOG.BOX",0x00A2:"CLEAN",0x00A3:"MDETERM",0x00A4:"MINVERSE",0x00A5:"MMULT",0x00A6:"FILES",0x00A7:"IPMT",0x00A8:"PPMT",0x00A9:"COUNTA",0x00AA:"CANCEL.KEY",0x00AB:"FOR",0x00AC:"WHILE",0x00AD:"BREAK",0x00AE:"NEXT",0x00AF:"INITIATE",0x00B0:"REQUEST",0x00B1:"POKE",0x00B2:"EXECUTE",0x00B3:"TERMINATE",0x00B4:"RESTART",0x00B5:"HELP",0x00B6:"GET.BAR",0x00B7:"PRODUCT",0x00B8:"FACT",0x00B9:"GET.CELL",0x00BA:"GET.WORKSPACE",0x00BB:"GET.WINDOW",0x00BC:"GET.DOCUMENT",0x00BD:"DPRODUCT",0x00BE:"ISNONTEXT",0x00BF:"GET.NOTE",0x00C0:"NOTE",0x00C1:"STDEVP",0x00C2:"VARP",0x00C3:"DSTDEVP",0x00C4:"DVARP",0x00C5:"TRUNC",0x00C6:"ISLOGICAL",0x00C7:"DCOUNTA",0x00C8:"DELETE.BAR",0x00C9:"UNREGISTER",0x00CC:"USDOLLAR",0x00CD:"FINDB",0x00CE:"SEARCHB",0x00CF:"REPLACEB",0x00D0:"LEFTB",0x00D1:"RIGHTB",0x00D2:"MIDB",0x00D3:"LENB",0x00D4:"ROUNDUP",0x00D5:"ROUNDDOWN",0x00D6:"ASC",0x00D7:"DBCS",0x00D8:"RANK",0x00DB:"ADDRESS",0x00DC:"DAYS360",0x00DD:"TODAY",0x00DE:"VDB",0x00DF:"ELSE",0x00E0:"ELSE.IF",0x00E1:"END.IF",0x00E2:"FOR.CELL",0x00E3:"MEDIAN",0x00E4:"SUMPRODUCT",0x00E5:"SINH",0x00E6:"COSH",0x00E7:"TANH",0x00E8:"ASINH",0x00E9:"ACOSH",0x00EA:"ATANH",0x00EB:"DGET",0x00EC:"CREATE.OBJECT",0x00ED:"VOLATILE",0x00EE:"LAST.ERROR",0x00EF:"CUSTOM.UNDO",0x00F0:"CUSTOM.REPEAT",0x00F1:"FORMULA.CONVERT",0x00F2:"GET.LINK.INFO",0x00F3:"TEXT.BOX",0x00F4:"INFO",0x00F5:"GROUP",0x00F6:"GET.OBJECT",0x00F7:"DB",0x00F8:"PAUSE",0x00FB:"RESUME",0x00FC:"FREQUENCY",0x00FD:"ADD.TOOLBAR",0x00FE:"DELETE.TOOLBAR",0x00FF:"User Defined Function",0x0100:"RESET.TOOLBAR",0x0101:"EVALUATE",0x0102:"GET.TOOLBAR",0x0103:"GET.TOOL",0x0104:"SPELLING.CHECK",0x0105:"ERROR.TYPE",0x0106:"APP.TITLE",0x0107:"WINDOW.TITLE",0x0108:"SAVE.TOOLBAR",0x0109:"ENABLE.TOOL",0x010A:"PRESS.TOOL",0x010B:"REGISTER.ID",0x010C:"GET.WORKBOOK",0x010D:"AVEDEV",0x010E:"BETADIST",0x010F:"GAMMALN",0x0110:"BETAINV",0x0111:"BINOMDIST",0x0112:"CHIDIST",0x0113:"CHIINV",0x0114:"COMBIN",0x0115:"CONFIDENCE",0x0116:"CRITBINOM",0x0117:"EVEN",0x0118:"EXPONDIST",0x0119:"FDIST",0x011A:"FINV",0x011B:"FISHER",0x011C:"FISHERINV",0x011D:"FLOOR",0x011E:"GAMMADIST",0x011F:"GAMMAINV",0x0120:"CEILING",0x0121:"HYPGEOMDIST",0x0122:"LOGNORMDIST",0x0123:"LOGINV",0x0124:"NEGBINOMDIST",0x0125:"NORMDIST",0x0126:"NORMSDIST",0x0127:"NORMINV",0x0128:"NORMSINV",0x0129:"STANDARDIZE",0x012A:"ODD",0x012B:"PERMUT",0x012C:"POISSON",0x012D:"TDIST",0x012E:"WEIBULL",0x012F:"SUMXMY2",0x0130:"SUMX2MY2",0x0131:"SUMX2PY2",0x0132:"CHITEST",0x0133:"CORREL",0x0134:"COVAR",0x0135:"FORECAST",0x0136:"FTEST",0x0137:"INTERCEPT",0x0138:"PEARSON",0x0139:"RSQ",0x013A:"STEYX",0x013B:"SLOPE",0x013C:"TTEST",0x013D:"PROB",0x013E:"DEVSQ",0x013F:"GEOMEAN",0x0140:"HARMEAN",0x0141:"SUMSQ",0x0142:"KURT",0x0143:"SKEW",0x0144:"ZTEST",0x0145:"LARGE",0x0146:"SMALL",0x0147:"QUARTILE",0x0148:"PERCENTILE",0x0149:"PERCENTRANK",0x014A:"MODE",0x014B:"TRIMMEAN",0x014C:"TINV",0x014E:"MOVIE.COMMAND",0x014F:"GET.MOVIE",0x0150:"CONCATENATE",0x0151:"POWER",0x0152:"PIVOT.ADD.DATA",0x0153:"GET.PIVOT.TABLE",0x0154:"GET.PIVOT.FIELD",0x0155:"GET.PIVOT.ITEM",0x0156:"RADIANS",0x0157:"DEGREES",0x0158:"SUBTOTAL",0x0159:"SUMIF",0x015A:"COUNTIF",0x015B:"COUNTBLANK",0x015C:"SCENARIO.GET",0x015D:"OPTIONS.LISTS.GET",0x015E:"ISPMT",0x015F:"DATEDIF",0x0160:"DATESTRING",0x0161:"NUMBERSTRING",0x0162:"ROMAN",0x0163:"OPEN.DIALOG",0x0164:"SAVE.DIALOG",0x0165:"VIEW.GET",0x0166:"GETPIVOTDATA",0x0167:"HYPERLINK",0x0168:"PHONETIC",0x0169:"AVERAGEA",0x016A:"MAXA",0x016B:"MINA",0x016C:"STDEVPA",0x016D:"VARPA",0x016E:"STDEVA",0x016F:"VARA",0x0170:"BAHTTEXT",0x0171:"THAIDAYOFWEEK",0x0172:"THAIDIGIT",0x0173:"THAIMONTHOFYEAR",0x0174:"THAINUMSOUND",0x0175:"THAINUMSTRING",0x0176:"THAISTRINGLENGTH",0x0177:"ISTHAIDIGIT",0x0178:"ROUNDBAHTDOWN",0x0179:"ROUNDBAHTUP",0x017A:"THAIYEAR",0x017B:"RTD"}
XF_Array = []
Cells = {}

Ptg = {
0x01:"PtgExp",
0x02:"PtgTbl",
0x03:"PtgAdd",
0x04:"PtgSub",
0x05:"PtgMul",
0x06:"PtgDiv",
0x07:"PtgPower",
0x08:"PtgConcat",
0x09:"PtgLt",
0x0A:"PtgLe",
0x0B:"PtgEq",
0x0C:"PtgGe",
0x0D:"PtgGt",
0x0E:"PtgNe",
0x0F:"PtgIsect",
0x10:"PtgUnion",
0x11:"PtgRange",
0x12:"PtgUplus",
0x13:"PtgUminus",
0x14:"PtgPercent",
0x15:"PtgParen",
0x16:"PtgMissArg",
0x17:"PtgStr",
0x18: {
	0x01:"PtgElfLel",
	0x02:"PtgElfRw",
	0x03:"PtgElfCol",
	0x06:"PtgElfRwV",
	0x07:"PtgElfColV",
	0x0A:"PtgElfRadical",
	0x0B:"PtgElfRadicalS",
	0x0D:"PtgElfColS",
	0x0F:"PtgElfColSV",
	0x10:"PtgElfRadicalLel",
	0x1D:"PtgSxName",
	},
0x19:{
	0x01:"PtgAttrSemi",
	0x02:"PtgAttrIf",
	0x04:"PtgAttrChoose",
	0x08:"PtgAttrGoto",
	0x10:"PtgAttrSum",
	0x20:"PtgAttrBaxcel",
	0x21:"PtgAttrBaxcel",
	0x40:"PtgAttrSpace",
	0x41:"PtgAttrSpaceSemi",
	},

0x1C:"PtgErr",
0x1D:"PtgBool",
0x1E:"PtgInt",
0x1F:"PtgNum",
0x20:"PtgArray",
0x21:"PtgFunc",
0x22:"PtgFuncVar",
0x23:"PtgName",
0x24:"PtgRef",
0x25:"PtgArea",
0x26:"PtgMemArea",
0x27:"PtgMemErr",
0x28:"PtgMemNoMem",
0x29:"PtgMemFunc",
0x2A:"PtgRefErr",
0x2B:"PtgAreaErr",
0x2C:"PtgRefN",
0x2D:"PtgAreaN",
0x39:"PtgNameX",
0x3A:"PtgRef3d",
0x3B:"PtgArea3d",
0x3C:"PtgRefErr3d",
0x3D:"PtgAreaErr3d",
0x40:"PtgArray",
0x41:"PtgFunc",
0x42:"PtgFuncVar",
0x43:"PtgName",
0x44:"PtgRef",
0x45:"PtgArea",
0x46:"PtgMemArea",
0x47:"PtgMemErr",
0x48:"PtgMemNoMem",
0x49:"PtgMemFunc",
0x4A:"PtgRefErr",
0x4B:"PtgAreaErr",
0x4C:"PtgRefN",
0x4D:"PtgAreaN",
0x59:"PtgNameX",
0x5A:"PtgRef3d",
0x5B:"PtgArea3d",
0x5C:"PtgRefErr3d",
0x5D:"PtgAreaErr3d",
0x60:"PtgArray",
0x61:"PtgFunc",
0x62:"PtgFuncVar",
0x63:"PtgName",
0x64:"PtgRef",
0x65:"PtgArea",
0x66:"PtgMemArea",
0x67:"PtgMemErr",
0x68:"PtgMemNoMem",
0x69:"PtgMemFunc",
0x6A:"PtgRefErr",
0x6B:"PtgAreaErr",
0x6C:"PtgRefN",
0x6D:"PtgAreaN",
0x79:"PtgNameX",
0x7A:"PtgRef3d",
0x7B:"PtgArea3d",
0x7C:"PtgRefErr3d",
0x7D:"PtgAreaErr3d"
}

Lbl_Builtin_names = {
    0x00:"Consolidate_Area",
    0x01:"Auto_Open",
    0x02:"Auto_Close",
    0x03:"Extract",
    0x04:"Database",
    0x05:"Criteria",
    0x06:"Print_Area",
    0x07:"Print_Titles",
    0x08:"Recorder",
    0x09:"Data_Form",
    0x0A:"Auto_Activate",
    0x0B:"Auto_Deactivate",
    0x0C:"Sheet_Title",
    0x0D:"_FilterDatabase"
}

#For demo purpose; these values will change for cell to cell.
# Ideal way is to implement get cell by reading row properties
# To do in future
Get_Cell = {
    50 : 3, #printable pages
    17 : 14.4, # height of the cell
    19 : 11, # size of font in points
    38 : 0 , # Shade foreground color as a number in the range 1 to 56. If color is automatic, returns 0
    24 : 0, #Font color of the first character in the cell, as a number in the range 1 to 56. If font color is automatic, returns 0.
}

def int2xlscolumn(x):
    base = len(XLS_columns)
    digits = []
    while x:
        digits.insert(0,XLS_columns[int((x % base)-1)%base])
        x = int(x / base)
    return ''.join(digits)

def col2int(col):
    out = 0
    for i in range(0,len(col)):
        out += (ord(col[i])-64)*pow(26,len(col)-1-i)

    return hex(out)

def getcell(formula):
    match_getcell = None
    getcell_typenum = None
    match_getcell = re.match('GET.CELL\((\d+),[A-Z]+\d+\)',formula)
    if match_getcell:
        getcell_typenum = int(match_getcell.groups()[0])

    if getcell_typenum and getcell_typenum in Get_Cell.keys():
        return Get_Cell[getcell_typenum]

    return 0

def calculate_values(setvalues):
    operators = ['+','-','*','/']
    for each_set_cell in setvalues.keys():
        set_value = setvalues[each_set_cell]
        if not type(set_value) == str:
            continue
        if set_value.find("GET.CELL")>-1:
            i = 0
            args = []
            arg1 = ""
            while i < len(set_value):
                if not set_value[i] in operators:
                    arg1 += set_value[i]
                else:
                    if len(arg1.strip())>0:
                        args.append(float(arg1.strip()))
                    else:
                        args.append(float('0.0'))
                    arg1 = ""
                    args.append(set_value[i])
                i +=1
            if len(arg1.strip())>0:
                if arg1.strip().startswith('GET.CELL'):
                    args.append(getcell(arg1))
                else:
                    args.append(float(arg1.strip()))

            result = 0

            while len(args)>0:
                element = args.pop(0)
                if element == '-':
                    result = result - args.pop(0)
                elif element == '+':
                    result = result + args.pop(0)
                elif element == '*':
                    result = result * args.pop(0)
                elif element == '/':
                    result = result / args.pop(0)
                elif element == '^':
                    result = result / args.pop(0)
                else:
                    result = element
        else:
            result = set_value
        setvalues[each_set_cell] = result
    return setvalues

def Emulate_setvalues(setvalues):
    operators = ['+','-','*','/']
    for each_set_cell in setvalues.keys():
        set_value = setvalues[each_set_cell]
        if (set_value).find("GET.CELL")>-1:
            i = 0
            args = []
            arg1 = ""
            while i < len(set_value):
                if not set_value[i] in operators:
                    arg1 += set_value[i]
                else:
                    if len(arg1.strip())>0:
                        if arg1.strip().startswith('GET.CELL'):
                            args.append(getcell(arg1))
                        else:
                            args.append(float(arg1.strip()))
                        arg1 = ""
                    args.append(set_value[i])
                i +=1
            if len(arg1.strip())>0:
                if arg1.strip().startswith('GET.CELL'):
                    args.append(getcell(arg1))
                else:
                    args.append(float(arg1.strip()))
            result = 0
            print args
            while len(args)>0:
                element = args.pop(0)
                if element == '-':
                    result = result - args.pop(0)
                elif element == '+':
                    result = result + args.pop(0)
                elif element == '*':
                    result = result * args.pop(0)
                elif element == '/':
                    result = result / args.pop(0)
                elif element == '^':
                    result = result / args.pop(0)
                else:
                    result = element
            setvalues[each_set_cell] = result
    return setvalues


def decyrpt_values(values_dict,formula_list):
    decrypted_formulas = []
    for each_formula in formula_list:
        #last , is the cell that it be added to
        m = re.match('FORMULA\((.*),\s*[A-Z]+\d+\s*\)',each_formula)
        output = ""
        if m:
            for each_sub_formula in m.groups()[0].split('&'):
                temp_formula = each_sub_formula.split('(')[0]
                if temp_formula == 'MID':
                    m = re.match('([A-Z]+)\(([A-Z]+\d+),(\d+),(\d+)\)',each_sub_formula.strip())
                    if m:
##                    if m.groups()[0] == 'MID':
                        if (m.groups()[1]) in values_dict.keys():
                            output += values_dict[m.groups()[1]][int(m.groups()[2]):int(m.groups()[2])+int(m.groups()[3])]
                elif temp_formula == 'CHAR':
                    m = re.match('[A-Z]+\(([A-Z]+\d+)([\+\-\*\/\^])(([A-Z]+\d+))\)',each_sub_formula.strip())
                    if m:
                        if m.groups()[0] in values_dict.keys() and m.groups()[2] in values_dict.keys():
                            if m.groups()[1]=='+':
                                output +=chr(iint(values_dict[m.groups()[0]]) + int(values_dict[m.groups()[2]]) )
                            elif m.groups()[1]=='-':
                                output +=chr(int(values_dict[m.groups()[0]]) - int(values_dict[m.groups()[2]]) )
                            elif m.groups()[1]=='*':
                                output +=chr(int(values_dict[m.groups()[0]]) * int(values_dict[m.groups()[2]]) )
                            elif m.groups()[1]=='/':
                                output +=chr(int(int(values_dict[m.groups()[0]]) / int(values_dict[m.groups()[2]]))%256 )
                            elif m.groups()[1]=='^':
                                output +=chr(int(values_dict[m.groups()[0]]) ^ int(values_dict[m.groups()[2]]) )


        decrypted_formulas.append(output)
    return decrypted_formulas



def parse_ptg_records(hex_str):

    if hex_str =="":
        return "",""
    Binary_operators = { "PtgAdd" : [0x3,"+"], "PtgSub":[0x4,"-"], "PtgMul":[0x5,"*"], "PtgDiv" : [0x6,"/"], "PtgPower" :[0x7,"^"], "PtgConcat" :[0x8,"&"], "PtgLt" :[0x9,"<"], "PtgEq":[0xa,'='],
                          "PtgLe" : [0xa,"<="], "PtgGe" : [0xc,">="], "PtgGt":[0xd,">"], "PtgNe":[0xe,"!="], "PtgIsect":[0xf," "], "PtgUnion":[0x10,","], "PtgRange":[0x11,":"]
    }

    Unary_operators = { "PtgUplus":[0x12,'+'],"PtgUminus":[0x13,'-'],"PtgPercent":[0x14,"%"]}
    Ignore_operators = {"PtgParen":[0x15],"PtgMissArg":[0x16]}
    Formula_Error_Codes = { 0x00:"#NULL!", 0x07:"#DIV/0!", 0x0F:"#VALUE!", 0x17:"#REF!", 0x1D:"#NAME?",0x24:"#NUM!",0x2A:"#N/A"}

    output = []
    output_1 = []
    main_formula = ""

    while hex_str!="":
        ptg_ordinal = ord(hex_str[0])
        if  ptg_ordinal in Ptg.keys():
            #print Ptg[ptg_ordinal]
            if Ptg[ptg_ordinal]=='PtgExp':
                assert (ptg_ordinal & 0x7F) == 0x01
                assert (ptg_ordinal & 0x80) == 0x00
                row = struct.unpack("<H",hex_str[1:3])[0]
                column = struct.unpack("<H",hex_str[3:5])[0]&0x3FFF
                hex_str = hex_str[5:]

            elif Ptg[ptg_ordinal]=='PtgTbl':
                assert (ptg_ordinal & 0x7F) == 0x02
                assert (ptg_ordinal & 0x80) == 0x00
                # following should be ignored if this is part of ObjectParsedFormula
                row = struct.unpack("<H",hex_str[1:3])[0]
                column = struct.unpack("<H",hex_str[3:5])[0]&0x3FFF
                hex_str = hex_str[5:]

            elif Ptg[ptg_ordinal] in Binary_operators.keys():
                ptg_operator = Binary_operators[Ptg[ptg_ordinal]]
                assert (ptg_ordinal & 0x7F) == ptg_operator[0]
                assert (ptg_ordinal & 0x80) == 0x00
                if len(output)<2:
                    print('insufficient arguments for %s'%(Ptg[ptg_ordinal]))
                    hex_str = hex_str[1:]
                    continue
                arg2 = output.pop()
                arg1 = output.pop()
                output.append("%s%s%s"%(arg1,ptg_operator[1],arg2))

                if len(output_1)<2:
                    print('insufficient arguments for %s'%(Ptg[ptg_ordinal]))
                    hex_str = hex_str[1:]
                    continue
                arg2 = output_1.pop()
                arg1 = output_1.pop()
                output_1.append("%s%s%s"%(arg1,ptg_operator[1],arg2))

                hex_str = hex_str[1:]

            elif Ptg[ptg_ordinal] in Unary_operators.keys():
                ptg_operator = Unary_operators[Ptg[ptg_ordinal]]
                assert (ptg_ordinal & 0x7F) == ptg_operator[0]
                assert (ptg_ordinal & 0x80) == 0x00
                if len(output)<1:
                    print('insufficient arguments for %s'%(Ptg[ptg_ordinal]))
                    hex_str = hex_str[1:]
                    continue
                arg1 = output.pop()
                if ptg_operator[1] == '-':
                    output.append("%s%s"%('-',arg1))
                elif ptg_operator[1] == '%':
                    output.append("%s%s"%(arg1,"%"))
                else:
                    output.append(arg1)


                if len(output_1)<1:
                    print('insufficient arguments for %s'%(Ptg[ptg_ordinal]))
                    hex_str = hex_str[1:]
                    continue
                arg1 = output_1.pop()
                if ptg_operator[1] == '-':
                    output_1.append("%s%s"%('-',arg1))
                elif ptg_operator[1] == '%':
                    output_1.append("%s%s"%(arg1,"%"))
                else:
                    output_1.append(arg1)


                hex_str = hex_str[1:]

            elif Ptg[ptg_ordinal] in Ignore_operators.keys():
                ptg_operator = Ignore_operators[Ptg[ptg_ordinal]]
                assert (ptg_ordinal & 0x7F) == ptg_operator[0]
                assert (ptg_ordinal & 0x80) == 0x00
                hex_str = hex_str[1:]

            elif Ptg[ptg_ordinal]=='PtgStr':
                assert (ptg_ordinal & 0x7F) == 0x17
                cch = ord(hex_str[1])  #length of String with higher bit indicating 0 means all higher bytes of the string are zero hence not added in length
                if ord(hex_str[2]) & 0x80:
                    output.append('"'+hex_str[3:3+(2*cch)]+'"')
                    output_1.append('"'+hex_str[3:3+(2*cch)]+'"')
                    hex_str = hex_str[3+(2*cch):]
                else:
                    output.append('"'+hex_str[3:3+(cch)]+'"')
                    output_1.append('"'+hex_str[3:3+(cch)]+'"')
                    hex_str = hex_str[3+(cch):]

            elif ptg_ordinal == 0x18:
                eptg = ord(hex_str[1])
                if eptg in Ptg[ptg_ordinal].keys():
                    hex_str = hex_str[6:]
                else:
                    print 'unknown sub ordinal'
                    output.append('Parsing incomplete')
                    output_1.append('Parsing incomplete')
                    hex = ""

            elif ptg_ordinal == 0x19:
                ptg_sub_ordinal = ord(hex_str[1])
                if ptg_sub_ordinal in Ptg[ptg_ordinal].keys():
                    if Ptg[ptg_ordinal][ptg_sub_ordinal] == 'PtgAttrSpace':
                        PtgAttrSpaceType = ord(hex_str[2])
                        assert PtgAttrSpaceType <= 0x06
                        cch = ord(hex_str[3])
                        hex_str = hex_str[4:]
                    elif Ptg[ptg_ordinal][ptg_sub_ordinal] == 'PtgAttrChoose':
                        assert (ptg_ordinal & 0x80) == 0x0
                        assert (ptg_sub_ordinal & 0x03) == 0x0
                        assert (ptg_sub_ordinal & 0x04) == 0x1
                        assert (ptg_sub_ordinal & 0xF1) == 0x0
                        coffset = struct.unpack("<H",hex_str[2:4])[0]
                        hex_str = hex_str[4+((coffset+1)*2):]
                    else:
                        hex_str = hex_str[4:]
                else:
                    print 'unknown sub ordinal'
                    output.append('Parsing incomplete')
                    output_1.append('Parsing incomplete')
                    hex = ""

            elif Ptg[ptg_ordinal]=="PtgErr":
                assert (ptg_ordinal & 0x7F) == 0x1c
                err_code = ord(hex_str[1])
                output.append(Formula_Error_Codes[err_code])
                output_1.append(Formula_Error_Codes[err_code])
                hex_str = hex_str[2:]

            elif Ptg[ptg_ordinal]=="PtgBool":
                assert (ptg_ordinal & 0x7F) == 0x1d
                if ord(hex_str[1]):
                    output.append("True")
                    output_1.append("True")
                else:
                    output.append("False")
                    output_1.append("False")
                hex_str = hex_str[2:]

            elif Ptg[ptg_ordinal]=="PtgInt":
                assert (ptg_ordinal & 0x7F) == 0x1e
                ptg_int_value = struct.unpack('<H',hex_str[1:3])[0]
                output.append("%s"%(str(ptg_int_value)))
                output_1.append("%s"%(str(ptg_int_value)))
                hex_str = hex_str[3:]

            elif Ptg[ptg_ordinal]=="PtgNum":
                assert (ptg_ordinal & 0x7F) == 0x1f
                ptg_float_value = struct.unpack('<d',hex_str[1:9])[0]
                output.append(ptg_float_value)
                output_1.append(ptg_float_value)
                hex_str = hex_str[9:]

            elif Ptg[ptg_ordinal]=="PtgArray":
                assert (ptg_ordinal & 0x1F) == 0x00  # 5 bis of ptg_ordinal
                ptg_data_type = (0x60 & ptg_ordinal)>>6
                assert ( (ptg_data_type==2) or (ptg_data_type == 3))
                hex_str = hex_str[8:]

            elif Ptg[ptg_ordinal]=="PtgFunc":
                assert (ptg_ordinal & 0x1F) == 0x01  # 5 bis of ptg_ordinal
                ptg_data_type = (0x60 & ptg_ordinal)>>6
                ptg_iftab = struct.unpack('<H',hex_str[1:3])[0]
                #Need to implement based on arguments
                argument_count = 0
                if ftab_functions[ptg_iftab] == 'MID':
                    argument_count = 3
                elif ftab_functions[ptg_iftab] == 'GET.CELL' or ftab_functions[ptg_iftab] == 'SET.VALUE' :
                    argument_count = 2
                else:
                    argument_count = 1
                if argument_count>0:
                    temp = []
                    temp_1 = []
                    while argument_count and len(output)>0:
                        temp.insert(0,output.pop())
                        temp_1.insert(0,output_1.pop())
                        argument_count -= 1
                output.append("%s(%s)"%(ftab_functions[ptg_iftab],','.join(temp)))

                if ftab_functions[ptg_iftab] == 'GET.CELL':
                    output_1.append(getcell("%s(%s)"%(ftab_functions[ptg_iftab],','.join(temp_1))))
                else:
                    output_1.append("%s(%s)"%(ftab_functions[ptg_iftab],','.join(temp_1)))
                hex_str = hex_str[3:]

            elif Ptg[ptg_ordinal]=="PtgFuncVar":
                assert (ptg_ordinal & 0x1F) == 0x02  # 5 bis of ptg_ordinal
                ptg_data_type = (0x60 & ptg_ordinal)>>6
                cparams = ord(hex_str[1])
                fCeFunc = (struct.unpack("<H",hex_str[2:4])[0] & 0x8000)>>15
                tab = struct.unpack("<H",hex_str[2:4])[0] & 0x7FFF
                if fCeFunc:
                     main_formula = ctab_functions[tab]
                else:
                    main_formula = ftab_functions[tab]

                temp = []
                temp_1 = []
                argument_count = cparams
                while argument_count and len(output)>0:
                    temp.insert(0,output.pop())
                    temp_1.insert(0,output_1.pop())
                    argument_count -= 1
                output.append("%s(%s)"%(main_formula,','.join(temp)))
                if main_formula=='GET.CELL':
                    output_1.append(getcell("%s(%s)"%(main_formula,','.join(temp_1))))
                else:
                    output_1.append("%s(%s)"%(main_formula,','.join(temp_1)))

                hex_str = hex_str[4:]


            elif Ptg[ptg_ordinal]=="PtgName":
                assert (ptg_ordinal & 0x1F) == 0x03  # 5 bis of ptg_ordinal
                ptg_data_type = (0x60 & ptg_ordinal)>>6
                name_index = struct.unpack("<I",hex_str[1:5])[0]
                output.append("Name_index_impelemenatation required : %d"%(name_index))
                output_1.append("Name_index_impelemenatation required : %d"%(name_index))
                hex_str = hex_str[5:]

            elif Ptg[ptg_ordinal]=="PtgRef":
                assert (ptg_ordinal & 0x1F) == 0x04  # 5 bis of ptg_ordinal
                ptg_data_type = (0x60 & ptg_ordinal)>>6
                row = struct.unpack("<H",hex_str[1:3])[0]
                column = struct.unpack("<H",hex_str[3:5])[0]&0x3FFF
                col_relative = struct.unpack("<H",hex_str[3:5])[0]&0x80
                row_relative = struct.unpack("<H",hex_str[3:5])[0]&0xC0
                output.append("%s%s"%(int2xlscolumn(column+1),str(row+1)))
                output_1.append("%s%s"%(int2xlscolumn(column+1),str(row+1)))
                hex_str = hex_str[5:]


            elif Ptg[ptg_ordinal]=="PtgArea":
                assert (ptg_ordinal & 0x1F) == 0x05  # 5 bis of ptg_ordinal
                ptg_data_type = (0x60 & ptg_ordinal)>>6
                row_first = struct.unpack("<H",hex_str[1:3])[0]
                row_last =  struct.unpack("<H",hex_str[3:5])[0]
                column_first = struct.unpack("<H",hex_str[5:7])[0]&0x3FFF
                column_last = struct.unpack("<H",hex_str[7:9])[0]&0x3FFF
                output.append("%s%s:%s%s"%(int2xlscolumn(column_first+1),str(row_first+1),int2xlscolumn(column_last+1),str(row_last+1)))
                output_1.append("%s%s:%s%s"%(int2xlscolumn(column_first+1),str(row_first+1),int2xlscolumn(column_last+1),str(row_last+1)))
                hex_str = hex_str[9:]

            elif Ptg[ptg_ordinal]=="PtgRef3d":
                assert (ptg_ordinal & 0x1F) == 0x1A  # 5 bis of ptg_ordinal
                ptg_data_type = (0x60 & ptg_ordinal)>>6
                ixti = struct.unpack("<H",hex_str[1:3])[0]
                row = struct.unpack("<H",hex_str[3:5])[0]
                column =  struct.unpack("<H",hex_str[5:7])[0]
                output.append("%s%s"%(int2xlscolumn(column+1),str(row+1)))
                output_1.append("%s%s"%(int2xlscolumn(column+1),str(row+1)))
                hex_str = hex_str[7:]

            else:
                hex_str = hex_str[-4:]

        else:
            print 'check value'
            hex_str = hex_str[1:]

    return (''.join([str(each) for each in output]),''.join([str(each) for each in output_1]))

def workbook(ole):
    global output,output_temp
    set_values = {}
    Sheets = []
    other_forumla = []
    formulas_used = {}
    defined_name = ""
    starting_point = []


    with ole.openstream('Workbook') as fh:
        record_id = 0
        while fh.tell() < ole.get_size('Workbook'):
            record_id = struct.unpack('<H',fh.read(2))[0]
            if record_id in XLS_Record.values():
                #print("%s at position %X"%(XLS_Record.keys()[XLS_Record.values().index(record_id)],fh.tell()-2))
                if (XLS_Record.keys()[XLS_Record.values().index(record_id)] == 'BoundSheet8'):
                    visibility = ""
                    sheet_type = ""
                    struct_length = struct.unpack('<H',fh.read(2))[0]
                    hex_str = fh.read(struct_length)
                    offset = struct.unpack('<I',hex_str[0:4])[0]
                    hsState = ord(hex_str[4])&0x03
                    if hsState == 0:
                        visibility = 'Visible'
                    elif hsState == 1:
                        visibility = "hidden"
                    elif hsState == 2:
                        visibility = "SuperHidden"

                    dt = ord(hex_str[5])
                    if dt == 0:
                        sheet_type = "Work/dialog sheet"
                    elif dt == 1:
                        sheet_type = "Macro sheet"
                    elif dt ==2 :
                        sheet_type = "Chart sheet"
                    elif dt == 6:
                        sheet_type = "VBA module"
                    string_length = ord(hex_str[6])
                    #hex_str[7] indicates whether high byte is \x00 or not
                    sheet_name = hex_str[8:]
                    Sheets.append("%s :: %s :: %s :: %X"%(sheet_name,sheet_type,visibility,offset))
                    continue
                elif (XLS_Record.keys()[XLS_Record.values().index(record_id)] == 'Number'):
                    struct_length = struct.unpack('<H',fh.read(2))[0]
                    hex_str = fh.read(struct_length)
                    row = struct.unpack('<H',hex_str[0:2])[0]
                    column = struct.unpack('<H',hex_str[2:4])[0]
                    ixfe = struct.unpack('<H',hex_str[4:6])[0]
                    value = struct.unpack('<d',hex_str[6:])[0]
                    set_values["%s%d"%(int2xlscolumn(column+1),row+1)]=value
                    continue
                elif (XLS_Record.keys()[XLS_Record.values().index(record_id)] == 'Lbl'):
                    #print 'inside lbl record'
                    defined_name = ""
                    struct_length = struct.unpack('<H',fh.read(2))[0]
                    hex_str = fh.read(struct_length)
                    temp_word = struct.unpack('<H',hex_str[0:2])[0]
                    fHidden = temp_word&0x0001
                    fFunc = temp_word & 0x0002
                    fOB = temp_word & 0x0004
                    fProc = temp_word & 0x0008
                    fCalcExp = temp_word & 0x0010
                    fBuiltin = temp_word & 0x0020
                    fGrp = (temp_word & 0x0FC00)>>6
                    reserved1 = temp_word & 0x1000
                    fPublished = temp_word & 0x2000
                    fWorkbookParam = temp_word & 0x4000
                    reserved2 = temp_word & 0x8000
                    chkey = ord(hex_str[2])
                    cch = ord(hex_str[3])
                    cce = struct.unpack('<H',hex_str[4:6])[0]
                    reserved3 = struct.unpack('<H',hex_str[6:8])[0]
                    itab = struct.unpack('<H',hex_str[0x8:0xa])[0]
                    reserved4 = ord(hex_str[0xb])
                    reserved5 = ord(hex_str[0xc])
                    reserved6 = ord(hex_str[0xd])
                    reserved7 = ord(hex_str[0xe])
                    name = hex_str[0xF:0xF+(cch)]
                    if fBuiltin:
                        if ord(name[0]) in Lbl_Builtin_names.keys():
                            defined_name = "%s%s"%(Lbl_Builtin_names[ord(name[0])],name[1:])
                    #print
                    #if hex_str
                    output,output_1 = parse_ptg_records(hex_str[0xF+(cch):])
                    if fHidden:
                        state ="hidden"
                    else:
                        state = "visible"
                    if defined_name:
                        temp_lbl = ("%s at cell : %s :: and state is :%s "%(defined_name,output,state))
                        starting_point.append(temp_lbl)
                    continue
                elif (XLS_Record.keys()[XLS_Record.values().index(record_id)] == 'XF'):
                    XF_Array.append(hex(fh.tell()-2))
                elif (XLS_Record.keys()[XLS_Record.values().index(record_id)] == 'Formula'):
                    pos = fh.tell()-2
                    struct_length = struct.unpack('<h',fh.read(2))[0]
                    row = struct.unpack('<H',fh.read(2))[0]
                    column = struct.unpack('<H',fh.read(2))[0]
                    ixfe = struct.unpack('<H',fh.read(2))[0]
                    fh.seek(14,1)
                    cell_parsed_formula = fh.read(struct.unpack('<H',fh.read(2))[0])

                    output = ""
                    output_1 = ""
                    output,output_1 = parse_ptg_records(cell_parsed_formula)
                    #print ("Cell :: %s%d :: %s"%(int2xlscolumn(column+1),row+1,output))
                    Cells["%s%d"%(int2xlscolumn(column+1),row+1)] = output
                    match = None
                    match = re.match("SET.VALUE\(([A-Z]+\d+),(.*)\)$",output_1)
                    if match:
                         set_values[match.groups()[0]] = match.groups()[1]
                    else:
                        other_forumla.append(output)

                    formula = output.split('(')[0]
                    if formula in formulas_used.keys():
                        formulas_used[formula] = formulas_used[formula]+ 1
                    else:
                        formulas_used[formula] = 1

                    continue
            else:
                print('unknown record at %X'%(fh.tell()-2))
            struct_length = struct.unpack('<h',fh.read(2))[0]
            fh.seek(struct_length,1)

    print set_values
    new_set_Values = ""
    d = ""
    try:
        new_set_Values = calculate_values(set_values)
        d = decyrpt_values(new_set_Values,other_forumla)
    except Exception as e:
        print("Exception occured :: %s "%str(e))

    #print new_set_Values
    print '*'*100
    print d
    print formulas_used
    print("total sheets : %d"%(len(Sheets)))
    print '\n'.join(Sheets)
    print '\n'.join(starting_point)
    return

def main():
    global Debug
    if len(sys.argv)<2:
        print(" Usage :: %s <input_xls_file_to_extract_macro4>"%(sys.argv[0]))
        sys.exit()
    if len(sys.argv)==3:
        Debug = True
    try:
        assert olefile.OleFileIO(sys.argv[1])
    except Exception as e:
        print str(e)
        sys.exit()

    ole = olefile.OleFileIO(sys.argv[1])
    if not ole.exists('Workbook'):
        print 'Expected xls 97 file as an input'
        return
    workbook(ole)
    return


if __name__ == '__main__':
    main()
