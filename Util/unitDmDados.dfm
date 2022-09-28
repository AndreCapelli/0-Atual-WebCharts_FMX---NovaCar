object dmDados: TdmDados
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  Height = 433
  Width = 573
  object Conn: TFDConnection
    Params.Strings = (
      'Server=CT-SUPORTE\CALLTECHSUPORTE'
      'User_Name=sa'
      'Password=123456'
      'Database=Zero'
      'DriverID=MSSQL')
    ResourceOptions.AssignedValues = [rvMacroCreate]
    ResourceOptions.MacroCreate = False
    LoginPrompt = False
    Left = 104
    Top = 64
  end
  object FDPhysMSSQLDriverLink1: TFDPhysMSSQLDriverLink
    Left = 200
    Top = 56
  end
  object qryGeral: TFDQuery
    Connection = Conn
    FetchOptions.AssignedValues = [evMode, evRowsetSize, evRecordCountMode, evDetailOptimize]
    FetchOptions.Mode = fmAll
    FetchOptions.RowsetSize = -1
    FetchOptions.RecordCountMode = cmFetched
    Left = 312
    Top = 272
  end
  object qryGeral2: TFDQuery
    Connection = Conn
    FetchOptions.AssignedValues = [evMode, evRowsetSize, evRecordCountMode, evDetailOptimize]
    FetchOptions.Mode = fmAll
    FetchOptions.RowsetSize = -1
    FetchOptions.RecordCountMode = cmFetched
    Left = 360
    Top = 272
  end
  object qryGeral3: TFDQuery
    Connection = Conn
    FetchOptions.AssignedValues = [evMode, evRowsetSize, evRecordCountMode, evDetailOptimize]
    FetchOptions.Mode = fmAll
    FetchOptions.RowsetSize = -1
    FetchOptions.RecordCountMode = cmFetched
    Left = 408
    Top = 272
  end
  object qryGeral4: TFDQuery
    Connection = Conn
    FetchOptions.AssignedValues = [evMode, evRowsetSize, evRecordCountMode, evDetailOptimize]
    FetchOptions.Mode = fmAll
    FetchOptions.RowsetSize = -1
    FetchOptions.RecordCountMode = cmFetched
    Left = 456
    Top = 272
  end
  object qryIdentCurrent: TFDQuery
    Connection = Conn
    Left = 320
    Top = 152
  end
  object FDErrorDialog: TFDGUIxErrorDialog
    Provider = 'Forms'
    Left = 440
    Top = 152
  end
end
