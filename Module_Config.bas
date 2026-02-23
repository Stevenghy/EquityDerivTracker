'==============================================================================
' Module_Config — Constants, Settings & Sheet References
'==============================================================================
Option Explicit

' ── Sheet Names ──
Public Const SHT_SCAN As String = "ScanInput"
Public Const SHT_LOG As String = "DataLog"
Public Const SHT_CLOSE As String = "PrevClose"
Public Const SHT_DASH As String = "Dashboard"

' ── ScanInput Column Layout ──
' A=Timestamp, B=Product, C=Lots, D=Notional, E=Condition, F=CrossLevel, G=TradeTime
Public Const COL_SCAN_TIMESTAMP As Long = 1
Public Const COL_SCAN_PRODUCT As Long = 2
Public Const COL_SCAN_LOTS As Long = 3
Public Const COL_SCAN_NOTIONAL As Long = 4
Public Const COL_SCAN_CONDITION As Long = 5
Public Const COL_SCAN_CROSSLEVEL As Long = 6
Public Const COL_SCAN_TRADETIME As Long = 7

' ── DataLog Column Layout ──
' A=DedupeKey, B=ScanTimestamp, C=Product, D=Lots, E=Notional, F=Condition,
' G=CrossLevel, H=TradeTime, I=PrevClose, J=PremiumPct, K=PremiumRounded,
' L=IDB_Flag, M=ManualOverride, N=Notes, O=TradeDate
Public Const COL_LOG_KEY As Long = 1
Public Const COL_LOG_SCANTS As Long = 2
Public Const COL_LOG_PRODUCT As Long = 3
Public Const COL_LOG_LOTS As Long = 4
Public Const COL_LOG_NOTIONAL As Long = 5
Public Const COL_LOG_CONDITION As Long = 6
Public Const COL_LOG_CROSSLEVEL As Long = 7
Public Const COL_LOG_TRADETIME As Long = 8
Public Const COL_LOG_PREVCLOSE As Long = 9
Public Const COL_LOG_PREMIUM As Long = 10
Public Const COL_LOG_PREMROUND As Long = 11
Public Const COL_LOG_IDBFLAG As Long = 12
Public Const COL_LOG_OVERRIDE As Long = 13
Public Const COL_LOG_NOTES As Long = 14
Public Const COL_LOG_TRADEDATE As Long = 15

' ── PrevClose Column Layout ──
' A=Product, B=CloseLevel, C=CloseDate
Public Const COL_PC_PRODUCT As Long = 1
Public Const COL_PC_CLOSE As Long = 2
Public Const COL_PC_DATE As Long = 3

' ── IDB Detection ──
Public Const IDB_DP_THRESHOLD As Long = 2  ' Max decimal places for IDB trade
Public Const PREMIUM_FORMAT As String = "0.00%"

' ── Auto Scan Settings ──
Public Const AUTO_SCAN_INTERVAL_MIN As Long = 30  ' Minutes between scans
Public Const MARKET_OPEN_HOUR As Long = 9          ' Local market open
Public Const MARKET_CLOSE_HOUR As Long = 17        ' Local market close

' ── Bloomberg ──
Public Const BBG_SCAN_FORMULA As String = "BSRCH"  ' Bloomberg search function
