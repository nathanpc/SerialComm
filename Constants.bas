Attribute VB_Name = "Constants"
Option Explicit

' CommandBar CommandBoxes
Const keyCbxBaud = "ComboBaud"
Const keyCbxParity = "ComboParity"
Const keyCbxData = "ComboData"
Const keyCbxStop = "ComboStop"

' Serial parameters
Const serialBaud = 0
Const serialParity = 1
Const serialDataBits = 2
Const serialStopBits = 3

Const baud300 = "300"
Const baud1200 = "1200"
Const baud2400 = "2400"
Const baud4800 = "4800"
Const baud9600 = "9600"
Const baud14400 = "14400"
Const baud19200 = "19200"
Const baud28800 = "28800"
Const baud38400 = "38400"
Const baud57600 = "57600"

Const parityNone = "None"
Const parityOdd = "Odd"
Const parityEven = "Even"
Const parityMark = "Mark"
Const paritySpace = "Space"

Const dataBits8 = "8 bits"
Const dataBits7 = "7 bits"
Const dataBits6 = "6 bits"
Const dataBits5 = "5 bits"

Const stopBits1 = "1 bit"
Const stopBits2 = "2 bits"

' Settings
Const settingsEnterBehaviour = 0 '"EnterBehaviour"
Const settEnterBehaviourSend = "Send"
Const settEnterBehaviourSendCRLF = "SendCRLF"

Const settingsMonitorType = 1 '"MonitorType"
Const settMonitorTypeRecv = "Received"
Const settMonitorTypeLogs = "Logs"
Const settMonitorTypeMix = "Mix"
