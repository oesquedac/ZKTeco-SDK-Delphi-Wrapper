(**********************************************************************************************************************
**                                                                                                                   **
**                                                                                                                   **
**    Libreria para accesar Relojes ZKTeco                                                                           **
**                                                                                                                   **
**    Desarrollado para Standalone SDK 6.3.1.37                                                                      **
**    Segun Changelog con fecha del 2018-05-24                                                                       **
**                                                                                                                   **
**    Se requiere la importacion de la libreria zkemkeeper.dll (TLB)                                                 **
**    Funciona en 32 y 64 bits                                                                                       **
**                                                                                                                   **
**    Creada por Oscar Adrian Esqueda Cortes (www.itcmx.com)                                                         **
**    Con la colaboración de MJLB - malobo@arainfor.com                                                              **
**    Última modificacion 20 de abril 2022                                                                           **
**                                                                                                                   **
**                                                                                                                   **
***********************************************************************************************************************)

unit uRelojes;

interface
uses

  System.Classes, Vcl.OleCtrls, zkemkeeper_TLB, Data.DB, Vcl.Grids, System.Generics.Collections;

type
  TReloj = class;

  TRelojes = class
  protected
  private
    FItems:TDictionary<string, TReloj>;

    function GetItem(AIndex: string): TReloj;
    function getTotal: Integer;
  public
    constructor Create; virtual;
    destructor Destroy; override;

    property Item[AIndex:string]:TReloj read GetItem;

    property Total:Integer read getTotal;

    function Buscar(AIP:string):Boolean; overload;
    function Buscar(AReloj:TReloj):Boolean; overload;

    function Agregar(AIP:string):TReloj; overload;
    function Agregar(AReloj:TReloj):Boolean; overload;

    function Indice(AIP:string):Integer; overload;
    function Indice(AReloj:TReloj):Integer; overload;

    procedure Quitar(AIP:string); overload;
    procedure Quitar(AReloj:TReloj); overload;
    procedure QuitarTodos(AEliminarObjetos:Boolean = True);

    class function Version:String;
    class function Fecha:TDateTime;

    class function SDKVersion:string;
  end; {TRelojes}


  TDayLight = class;
  TPersonal = class;
  TLog = class;
  TAsistencia = class;

  (***********************************************
  * Opciones para eliminar datos del dispositivo *
  ************************************************)
  TLimpiarDatos = (ldInvalido, ldHuellas, ldNinguno, ldOperaciones, ldUsuarios);

  (********************************************
  * Controlar un reloj por medio de la SDK    *
  *********************************************)
  TReloj = class
  protected
    procedure ActualizarError(AMsg:string); overload;
    procedure ActualizarError(AMsg:string; Args: array of const); overload;
  private
    FCZKEM: TCZKEM;
    FDevID: Integer;

    FIP: string;
    FPuerto: Integer;
    FDescripcion: String;
    FDeviceStatus: TStringList;
    FConectado: Boolean;
    FErrorCode: Integer;
    FFechaHora: TDateTime;
    FDeviceInfo: TStringList;
    FDayLight: TDayLight;
    FOnDeviceStatusChange: TNotifyEvent;
    FPersonal: TPersonal;
    FLog: TLog;
    FAsistencia: TAsistencia;

    procedure DeviceStatusChange(Sender:TObject);
    Procedure InitDeviceInfo;
  public
    Modified: string;

    constructor Create(AOwner:TRelojes);
    destructor Destroy; override;

    property Descripcion:String read FDescripcion write FDescripcion;
    property IP:string read FIP write FIP;
    property Puerto:Integer read FPuerto write FPuerto;

    function BNetSpeedEncode(speed: Integer): string;
    function BTimeDecode(const TimeStr: string): Integer;
    function BTimeEncode(MinuteSecond: Integer): string;
    function BNetSpeedDecode(const speed: string): Integer;

    function Conectar:Boolean;
    procedure Desconectar;
    property Conectado:Boolean read FConectado;

    property DevID:Integer read FDevID;
    function GetDevID:Integer;

    function Respaldar(AArchivo:string):Boolean;
    function Restaurar(AArchivo:string):Boolean;

    function LimpiarDatos(ACualDatos:TLimpiarDatos):Boolean;

    property ErrorCode:Integer read FErrorCode;
    function ErrorMsg:string; overload;
    function ErrorMsg(AErrorCode:Integer):string; overload;
    procedure MostrarError; overload;
    procedure MostrarError(AMsg:string); overload;
    procedure MostrarError(AMsg:string; Args: array of const); overload;

    procedure ReadDeviceStatus;
    property DeviceStatus:TStringList read FDeviceStatus;
    procedure AddDeviceStatus(AMsg:String); overload;
    procedure AddDeviceStatus(AMsg:String; Args: array of const); overload;
    property OnDeviceStatusChange:TNotifyEvent read FOnDeviceStatusChange write FOnDeviceStatusChange;

    procedure LoadDeviceInfo;
    procedure SaveDeviceInfo;
    procedure InitModified;
    property DeviceInfo:TStringList read FDeviceInfo;

    property FechaHora:TDateTime read FFechaHora write FFechaHora;
    function FechaHoraLeer:Boolean;

    function FechaHoraCambiar:Boolean; overload;
    function FechaHoraCambiar(AFechaHora:TDateTime):Boolean; overload;
    function FechaHoraCambiarPC:Boolean;

    property DayLight:TDayLight read FDayLight;
    function LoadDayLight:Boolean;
    function SaveDayLight:Boolean; overload;
    function SaveDayLight(ADayLight:TDayLight):Boolean; overload;
    function SaveDayLight(ASoporte:Integer; AInicio, AFin:String):Boolean; overload;

    function TotalUsuarios:Integer;
    function TotalRegistros:Integer;

    function TomarImagen:Boolean;

    procedure Beep(ADelayMS:Integer);
    function SubirArchivo(AArchivo:string):Boolean;

    property Personal:TPersonal read FPersonal;
    property Log:TLog read FLog;

    property Asistencia:TAsistencia read FAsistencia;
  end; {TReloj}

  (********************************************
  * Clase para controlar el Horario de verano *
  *********************************************)
  TDayLight = class
  private
    FOwner:TReloj;
    FFin: string;
    FInicio: string;
    FSoporte: Integer;
    FCargado: Boolean;
  public
    constructor Create(AOwner:TReloj);

    property Soporte:Integer read FSoporte write FSoporte;
    property Inicio:string read FInicio write FInicio;
    property Fin:string read FFin write FFin;

    function Actualizar:Boolean;
    function Guardar(ASoporte:Integer; AInicio, AFin:String):Boolean; overload;

    property Cargado:Boolean read FCargado;
  end; {TDayLight}

  (********************************************
  * Administrar usuarios / personal           *
  *                                           *
  * Record para llenado de Dataset            *
  *********************************************)
  TDataSetCamposPersonal = class
  private
    FCodigo: string;
    FTarjeta: string;
    FNombre: string;
    FID: string;
  public
    constructor Create;

    property ID:string read FID write FID;
    property Codigo:string read FCodigo write FCodigo;
    property Nombre:string read FNombre write FNombre;
    property Tarjeta:string read FTarjeta write FTarjeta;
  end; {TDataSetCamposPersonal}

  TPersonal = class
  private
    FOwner:TReloj;
    FEnrollNumber: Integer;
    FPrivilegios: Integer;
    FActivo: Boolean;
    FPassword: String;
    FNombre: String;
    FTarjeta: string;
    FDataSetCampos: TDataSetCamposPersonal;

    procedure Init;
  public
    constructor Create(AOwner:TReloj);
    destructor Destroy; override;

    property EnrollNumber:Integer read FEnrollNumber write FEnrollNumber;
    property Nombre:String read FNombre write FNombre;
    property Password:String read FPassword write FPassword;
    property Privilegios:Integer read FPrivilegios write FPrivilegios;
    property Activo:Boolean read FActivo write FActivo;
    property Tarjeta:string read FTarjeta write FTarjeta;

    function Cargar:Boolean; overload;
    function Cargar(AEnrollNumber:Integer):Boolean; overload;
    function Guardar:Boolean;

    function SubirFoto(AImagen:String):Boolean;
    function EliminarFoto:Boolean; overload;
    function EliminarFoto(AEnroll:Integer):Boolean; overload;
    function EliminarTodasFotos:Boolean;

    function Eliminar(AEnrollNumber:Integer):Boolean;

    function CapturarHuella(AIDUsuario, ADedo:Integer):Boolean;

    function QuitarAdministradores:Boolean;
    function LimpiarPersonal:Boolean;

    property DataSetCampos:TDataSetCamposPersonal read FDataSetCampos;
    procedure ListaEn(ADataset:tDataset);

    function Respaldar(AArchivo:string):Boolean;
    function Restaurar(AArchivo:string):Boolean;
  end; {TPersonal}

  (********************************************
  * Administrar la bitacora                   *
  *********************************************)
  TLog = class
  private
    FOwner:TReloj;
    FBitacora: TStringList;
  public
    constructor Create(AOwner:TReloj);
    destructor Destroy; override;

    property Bitacora:TStringList read FBitacora;

    procedure Actualizar;
    procedure GuardarCSV(AArchivo:string);

    procedure LimpiarBitacora;
  end; {TLog}

  (********************************************
  * Procesar registros de asistencia          *
  *                                           *
  * Record para llenado de Dataset            *
  *********************************************)
  TDataSetCamposAsistencia = class
  private
    FCodigo: string;
    FYear: string;
    FMinutos: string;
    FInOut: string;
    FHora: string;
    FMes: string;
    FFechaHora: string;
    FDia: string;
    FSegundos: string;
    FWorkCode: string;
    FVerify: string;
    FDevID: string;
  public
    constructor Create;

    property DevID:string read FDevID write FDevID;
    property Codigo:string read FCodigo write FCodigo;
    property Verify:string read FVerify write FVerify;
    property InOut:string read FInOut write FInOut;
    property Year:string read FYear write FYear;
    property Mes:string read FMes write FMes;
    property Dia:string read FDia write FDia;
    property Hora:string read FHora write FHora;
    property Minutos:string read FMinutos write FMinutos;
    property Segundos:string read FSegundos write FSegundos;
    property FechaHora:string read FFechaHora write FFechaHora;
    property WorkCode:string read FWorkCode write FWorkCode;
  end; {TDataSetCamposPersonal}

  TAsistencia = class
  private
    FOwner:TReloj;
    FDataSetCamposAsistencia: TDataSetCamposAsistencia;
    FAsistencia: TStringList;
    function getTotal: Integer;
  public
    constructor Create(AOwner:TReloj);
    destructor Destroy; override;

    property DataSetCamposAsistencia:TDataSetCamposAsistencia read FDataSetCamposAsistencia;

    property Asistencia:TStringList read FAsistencia write FAsistencia;
    function Descargar:Boolean;
    property Total:Integer read getTotal;

    procedure GuardarCSV(AArchivo:string);
    procedure GuardarDataset(ADataset:TDataSet);
    procedure EnviarStringGrid(AStringGrid:TStringGrid);

    function Limpiar:Boolean;
  end; {TAsistencia}

const
  StatusNames: array [1 .. 14] of string = ('Total administradores', 'Total usuarios', 'Total FP', // 2
  'Total Password', 'Total manage record', 'Total registros (E/S)', 'Nominal FP number', // 6
  'Nominal user number', 'Nominal In and out record number', 'Remain FP number',
  'Remain user number', 'Remain In and out record number', 'Total FC', 'Nominal FC number');
  Languages: array [0 .. 2] of string = ('English', 'Simplified Chinese', 'Traditional Chinese');
  BaudRates: array [0 .. 6] of string = ('1200 bps', '2400 bps', '4800 bps', '9600 bps',
    '19200 bps', '38400 bps', '115200 bps');
  CRCs: array [0 .. 2] of string = ('Nothing', 'Even', 'Odd');
  StopBits: array [0 .. 1] of string = ('One', 'Two');
  DateSps: array [0 .. 1] of string = ('"/"', '"-"');
  MSpeeds: array [0 .. 2] of string = ('Low speed', 'High speed', 'Auto');
  OnOffs: array [0 .. 1] of string = ('Off', 'On');
  NetSpeeds: array [0 .. 4] of string = ('10M_H', '100M_H', '10M_F', '100M_F', 'AUTO');
  NetSpeedValues: array [0 .. 4] of Integer = (0, 1, 4, 5, 8);

implementation

uses
  System.SysUtils, System.DateUtils;

{ TRelojes }

function TRelojes.Agregar(AIP: string): TReloj;
begin
if not Buscar(AIP) then begin
  Result := TReloj.Create(Self);
  Result.IP := AIP;
  Result.Descripcion := AIP;

  FItems.Add(AIP, Result);
end else
  Result := nil;
end;

function TRelojes.Agregar(AReloj: TReloj): Boolean;
begin
Result := Buscar(AReloj);
if not Result then
  FItems.Add(AReloj.IP, AReloj);
end;

function TRelojes.Buscar(AReloj: TReloj): Boolean;
begin
Result := FItems.ContainsValue(AReloj);
end;

function TRelojes.Buscar(AIP: string): Boolean;
begin
Result := FItems.ContainsKey(AIP);
end;

constructor TRelojes.Create;
begin
FItems := TDictionary<string, TReloj>.Create;
end;

destructor TRelojes.Destroy;
begin
QuitarTodos(True);
FItems.Free;

inherited Destroy;
end;

class function TRelojes.Fecha: TDateTime;
var
  FormatSettings:TFormatSettings;
begin
{$WARNINGS OFF}
FormatSettings := TFormatSettings.Create($080A);
{$WARNINGS ON}
Result := StrToDate('13/may/2022', FormatSettings);
end;

function TRelojes.GetItem(AIndex: string): TReloj;
begin
if Buscar(AIndex) then
  Result := FItems.Items[AIndex]
else
  Result := nil;
end;

function TRelojes.getTotal: Integer;
begin
Result := FItems.Count;
end;

function TRelojes.Indice(AReloj: TReloj): Integer;
var
  i:Integer;
  sIP:string;
begin
sIP := '';
var aKeys := FItems.Keys.ToArray;

for i := low(aKeys) to High(aKeys) do begin
  sIP := aKeys[i];

  if AReloj.IP = sIP then
    Break;
end; {for}

if AReloj.IP <> sIP then
  i := -1;

Result := i;
end;

procedure TRelojes.Quitar(AReloj: TReloj);
begin
Quitar(AReloj.IP);
end;

procedure TRelojes.Quitar(AIP: string);
begin
if Buscar(AIP) then begin
  Item[AIP].Free;
  Fitems.Remove(AIP);
end; {if}

FItems.TrimExcess;
end;

procedure TRelojes.QuitarTodos(AEliminarObjetos:Boolean = True);
var
  sIP:string;
begin
for sIP in FItems.Keys do begin
  FItems.Items[sIP].Free;
  FItems.Items[sIP] := nil;
end; {for}

FItems.Clear;

FItems.TrimExcess;
end;

function TRelojes.Indice(AIP: string): Integer;
var
  i:Integer;
  sIP:string;
begin
sIP := '';
var aKeys := FItems.Keys.ToArray;

for i := low(aKeys) to High(aKeys) do begin
  sIP := aKeys[i];

  if AIP = sIP then
    Break;
end; {for}

if AIP <> sIP then
  i := -1;

Result := i;
end;

class function TRelojes.SDKVersion: string;
var
  CZKEM: TCZKEM;
  s:WideString;
begin
CZKEM := TCZKEM.Create(nil);
try
  CZKEM.GetSDKVersion(s);
  Result := s;
finally
  CZKEM.Free;
end; {try}
end;


class function TRelojes.Version: String;
begin
Result := 'v0.0.33.20220513';
end;

{ TReloj }

procedure TReloj.Beep(ADelayMS: Integer);
begin
FCZKEM.Beep(ADelayMS);
end;

function TReloj.BNetSpeedDecode(const speed: string): Integer;
var
  i: Integer;
begin
result := NetSpeedValues[0];
for i := 0 to length(NetSpeeds) - 1 do
  if speed = NetSpeeds[i] then begin
    result := NetSpeedValues[i];
    exit;
  end; {if}
end;

function TReloj.BNetSpeedEncode(speed: Integer): string;
var
  i: Integer;
begin
  result := NetSpeeds[0];
  for i := 0 to length(NetSpeeds) - 1 do
    if speed = NetSpeedValues[i] then
      begin
        result := NetSpeeds[i];
        exit;
      end;
end;

function TReloj.BTimeDecode(const TimeStr: string): Integer;
var
  p, m, s: Integer;
begin
  p := pos(':', TimeStr);
  m := strtointdef(copy(TimeStr, 1, p - 1), -1);
  s := strtointdef(copy(TimeStr, p + 1, 100), -1);

  {$HINTS OFF}
  if (m < 0) or (s < 0) or (m > 255) or (s > 255) then

    result := 255;
  {$HINTS ON}

  result := m * 256 + s;
end;

function TReloj.BTimeEncode(MinuteSecond: Integer): string;
var
  m, s: Integer;
begin
  m := MinuteSecond div 256;
  s := MinuteSecond mod 256;
  if (MinuteSecond < 0) or (m > 59) or (s > 59) then
    result := 'No'
  else
    result := format('%d:%d', [m, s]);
end;

function TReloj.Conectar: Boolean;
begin
FConectado := FCZKEM.Connect_net(IP, Puerto);
if FConectado then
  FDevID := GetDevID;
Result := FConectado;
end;

constructor TReloj.Create(AOwner: TRelojes);
begin
FCZKEM := TCZKEM.Create(nil);
FDevID := 1;

FDeviceStatus := TStringList.Create;
FDeviceStatus.OnChange := DeviceStatusChange;

FDeviceInfo := TStringList.Create;
InitDeviceInfo;

FDayLight := TDayLight.Create(Self);

FPersonal := TPersonal.Create(Self);

FLog := TLog.Create(Self);

FAsistencia := TAsistencia.Create(Self);
end;

procedure TReloj.Desconectar;
begin
FCZKEM.Disconnect;
FConectado := False;
end;

destructor TReloj.Destroy;
begin
FAsistencia.Free;
FLog.Free;
FPersonal.Free;
FDayLight.Free;
FCZKEM.Free;
FDeviceStatus.Free;
FDeviceInfo.Free;
inherited Destroy;
end;


procedure TReloj.DeviceStatusChange(Sender: TObject);
begin
if Assigned(FOnDeviceStatusChange) then
  FOnDeviceStatusChange(Sender);
end;

function TReloj.ErrorMsg(AErrorCode: Integer): string;
begin
case AErrorCode of
  0: Result := 'Conectado correctamente';
  -1: Result := 'Error al invocar la interfaz o el SDK no se inicializa y debe volver a conectarse';
  -2: Result := 'Error al inicializar o Error de lectura/escritura de archivos';
  -3: Result := 'Error al inicializar parametros o tamaoo incorrecto';
  -4: Result := 'Espacio insuficiente';
  -5: Result := 'Error de lectura del modo de datos o Los datos ya existen';
  -6: Result := 'Contraseoa incorrecta';
  -7: Result := 'Error de respuesta';
  -8: Result := 'Tiempo de espera de recepcion';
  -307: Result := 'Tiempo de espera de conexion';
  -201: Result := 'El dispositivo esto ocupado';
  -199: Result := 'Nuevo modo';
  -103: Result := 'dispositivo enviar de vuelta error de version de la cara error';
  -102: Result := '"Error de version de plantilla de cara, como plantilla de cara 8.0 enviar al dispositivo 7.0"';
  -101: Result := 'Error en la memoria malloc';
  -100: Result := 'No se admite o los datos no existe';
  -10: Result := 'La duracion de los datos transmitidos es incorrecta';
  //0: Result := 'Datos no encontrados o datos duplicados';
  1: Result := 'Funcionamiento correcto';
  4: Result := 'Error de parametro';
  101: Result := 'Error de asignacion de bofer';
  102: Result := 'invocar repetidamente';
  -12001: Result := 'Tiempo de espera de creacion de sockets (tiempo de espera de conexion)';
  -12002: Result := 'Memoria insuficiente';
  -12003: Result := 'Version incorrecta del socket';
  -12004: Result := 'No es protocolo TCP';
  -12005: Result := 'Tiempo de espera';
  -12006: Result := 'Tiempo de espera de la transmision de datos';
  -12007: Result := 'Tiempo de espera de lectura de datos';
  -12008: Result := 'Error al leer socket';
  -13009: Result := 'Error de evento en espera';
  -13010: Result := 'Intentos de reintento excedidos';
  -13011: Result := 'ID de respuesta incorrecto';
  -13012: Result := 'Error de suma de comprobacion';
  -13013: Result := 'Tiempo de espera del evento';
  -13014: Result := 'DIRTY_DATA';
  -13015: Result := 'Tamaoo del bofer demasiado pequeoo';
  -13016: Result := 'Longitud de datos incorrecta';
  -13017: Result := 'Lectura de datos no volida1';
  -13018: Result := 'Lectura de datos no volida2';
  -13019: Result := 'Lectura de datos no volida3';
  -13020: Result := 'Pordida de datos';
  -13021: Result := 'Error de inicializacion de memoria';
  -15001: Result := 'Invocar repetidamente el valor devuelto de la clave de estado emitida por la interfaz SetShortkey';
  -15002: Result := 'Invocar repetidamente el valor de cambio de descripcion emitido por la interfaz SetShortkey';
  -15003: Result := 'El meno de dos niveles no se abre en el dispositivo y no es necesario emitir los datos';
  -15100: Result := 'Se produce un error al obtener table structure';
  -15101: Result := 'El campo de condicion no existe en la estructura de la tabla';
  -15102: Result := 'Inconsistencia en el nomero total de campos';
  -15103: Result := 'Inconsistencia en la ordenacion de campos';
  -15104: Result := 'Error de asignacion de memoria';
  -15105: Result := 'Error de anolisis de datos';
  -15106: Result := 'Desbordamiento de datos a medida que los datos transmitidos superan los 4M';
  -15108: Result := 'Opciones no volidas';
  -15113: Result := 'Error de anolisis de datos: no se encontro el identificador de tabla';
  -15114: Result := 'Se devuelve una excepcion de datos ya que el nomero de campos es menor o igual a 0';
  -15115: Result := '"Se devuelve una excepcion de datos, ya que el nomero total de campos de tabla es independiente del nomero total de campos de los datos"';
  2000: Result := 'Devolver ACEPTAR para ejecutar';
  -2001: Result := 'Comando Return Fail to execute';
  -2002: Result := 'Devolver datos';
  -2003: Result := 'Evento registrado ocurre';
  -2004: Result := 'Devolvio Comando REPEAT';
  -2005: Result := 'Devolvio Comando UNAUTH';
  -65535: Result := '0xffff Devolvio Comando Unknown';
  65535: Result := '0xffff Devolvio Comando Unknown';
  -4999: Result := 'Error de lectura del parametro del dispositivo';
  -4998: Result := 'Error de escritura de parametros de dispositivo';
  -4997: Result := 'La longitud de los datos enviados por el software al dispositivo es incorrecta';
  -4996: Result := 'Existe un error de parametro en los datos enviados por el software al dispositivo';
  -4995: Result := 'Error al agregar datos a la base de datos';
  -4994: Result := 'Error al actualizar la base de datos';
  -4993: Result := 'Error al leer los datos de la base de datos';
  -4992: Result := 'Error al eliminar datos de la base de datos';
  -4991: Result := 'Datos no encontrados en la base de datos';
  -4990: Result := 'La cantidad de datos en la base de datos alcanza el lomite';
  -4989: Result := 'Error al asignar memoria a una sesion';
  -4988: Result := 'Espacio insuficiente en la memoria asignada a una sesion';
  -4987: Result := 'La memoria asignada a una sesion se desborda';
  -4986: Result := 'El archivo no existe';
  -4985: Result := 'Error de lectura de archivos';
  -4984: Result := 'Error de escritura de archivos';
  -4983: Result := 'Error al calcular el valor de hash';
  -4982: Result := 'Error al asignar memoria';
end; {case}
end;

function TReloj.ErrorMsg: string;
begin
Result := ErrorMsg(FErrorCode);
end;

function TReloj.FechaHoraCambiar: Boolean;
begin
Result := FechaHoraCambiar(FechaHora);
end;

function TReloj.FechaHoraCambiar(AFechaHora: TDateTime): Boolean;
var
  dwYear, dwMonth, dwDay, dwHour, dwMinute, dwSecond, dwMil: Word;
begin
DecodeDateTime(AFechaHora, dwYear, dwMonth, dwDay, dwHour, dwMinute, dwSecond, dwMil);

Result := FCZKEM.SetDeviceTime2(FDevID, dwYear, dwMonth, dwDay, dwHour, dwMinute, dwSecond);
if Result then begin
  AddDeviceStatus('SetDeviceTime OK.');
  if FechaHoraLeer then
    AddDeviceStatus('DeviceTime=%s', [FormatDateTime('dd-mm-yyyy hh:MM:ss', FechaHora)])
  else begin
      ActualizarError('Read Device Time');
  end;
end else begin
  ActualizarError('Set Device Time');;
end;
end;

function TReloj.FechaHoraCambiarPC: Boolean;
begin
Result := FCZKEM.SetDeviceTime(FDevID);
end;

function TReloj.FechaHoraLeer: Boolean;
var
  dwYear, dwMonth, dwDay, dwHour, dwMinute, dwSecond: Integer;
begin
Result := FCZKEM.GetDeviceTime(FDevID, dwYear, dwMonth, dwDay, dwHour, dwMinute, dwSecond);

if Result then
  FechaHora := EncodeDateTime(dwYear, dwMonth, dwDay, dwHour, dwMinute, dwSecond, 0)
else
  ActualizarError('Get Device Time');
end;

function TReloj.GetDevID: Integer;
var
  Value:Integer;
begin
if FCZKEM.GetDeviceInfo(FDevID, 2, Value) then
  Result := Value
else
  Result := 1;
end;

procedure TReloj.InitDeviceInfo;
begin
FDeviceInfo.Clear;
FDeviceInfo.Add('Registrable administrators number=');
FDeviceInfo.Add('Device ID=');
FDeviceInfo.Add('Languages=');
FDeviceInfo.Add('Auto power off time=');
FDeviceInfo.Add('Lock control delay(20ms)=');
FDeviceInfo.Add('In and out record warning=');
FDeviceInfo.Add('manage record warning=');
FDeviceInfo.Add('Confirm interval time=');
FDeviceInfo.Add('Baud Rate=');
FDeviceInfo.Add('Even and Odd=');

FDeviceInfo.Add('Stop bit=');
FDeviceInfo.Add('Date list separator=');
FDeviceInfo.Add('Network=');
FDeviceInfo.Add('RS232=');
FDeviceInfo.Add('RS485=');
FDeviceInfo.Add('Voice=');
FDeviceInfo.Add('Identification speed=');
FDeviceInfo.Add('idle=');
FDeviceInfo.Add('Shutdown time=');
FDeviceInfo.Add('PowerOn time=');

FDeviceInfo.Add('Sleep time=');
FDeviceInfo.Add('Auto Bell=');
FDeviceInfo.Add('Match threshold=');
FDeviceInfo.Add('Register threshold=');
FDeviceInfo.Add('1:1 threshold=');
FDeviceInfo.Add('Show score=');
FDeviceInfo.Add('Unlock person count=');
FDeviceInfo.Add('Only verify number card=');
FDeviceInfo.Add('Net Speed=');
FDeviceInfo.Add('Must registe card=');

FDeviceInfo.Add('Time out of temp state keep=');
FDeviceInfo.Add('Time out of input number=');
FDeviceInfo.Add('Time out of menu keep=');
FDeviceInfo.Add('Date formate=');
FDeviceInfo.Add('Only 1:1=');
end;

procedure TReloj.InitModified;
begin
Modified := stringofchar(#0, 100);
end;

function TReloj.LimpiarDatos(ACualDatos: TLimpiarDatos): Boolean;
var
  iCual:Integer;
begin
iCual := ord(ACualDatos);

Result := iCual > 0;
if Result then begin
  Result := FCZKEM.ClearData(DevID, iCual);
  if Result then begin
    FCZKEM.RefreshData(DevID);
    AddDeviceStatus('Limpiar datos');
  end else
    ActualizarError('Limpiar datos');
end else
  ActualizarError('Limpiar datos invalido');
end;

function TReloj.LoadDayLight: Boolean;
begin
Result := DayLight.Actualizar;
end;

procedure TReloj.LoadDeviceInfo;
var
  Value, i: Integer;
begin
for i := 1 to pred(FDeviceInfo.Count) do
  if FCZKEM.GetDeviceInfo(FDevID, i, Value) then begin
    if i = 29 then
      FDeviceInfo.ValueFromIndex[i - 1] := BNetSpeedEncode(Value)
    else if (i >= 19) and (i <= 22) then
      FDeviceInfo.ValueFromIndex[i - 1] := BTimeEncode(Value)
    else
      FDeviceInfo.ValueFromIndex[i - 1] := inttostr(Value);
  end else begin
    ActualizarError('Load Device Info');
  end;
end;

procedure TReloj.ReadDeviceStatus;
var
  s: widestring;
  Value, i: Integer;

begin
FDeviceStatus.Clear;

if FCZKEM.GetFirmwareVersion(FDevID, s) then
  AddDeviceStatus('Firmware Version: ' + s)
else
  ActualizarError('Device Status GetFirmwareVersion');

if FCZKEM.GetSerialNumber(FDevID, s) then
  AddDeviceStatus('Serial Number: ' + s)
else
  ActualizarError('Device Status GetSerialNumber');

if FCZKEM.GetProductCode(FDevID, s) then
  AddDeviceStatus('ProductCode: ' + s)
else
  ActualizarError('Device Status GetProductCode');

if FechaHoraLeer then
  AddDeviceStatus('DeviceTime=%s', [FormatDateTime('dd-mm-yyyy hh:MM:ss', FechaHora)])
else
  ActualizarError('Device Status GetDeviceTime');

for i := 1 to length(StatusNames) - 2 do begin
  if i = 4 then
    if FCZKEM.GetDeviceStatus(FDevID, 21, Value) then
      AddDeviceStatus('%s: %d', [StatusNames[13], Value])
    else
      ActualizarError(format('GetDeviceStatus(%d)', [21]));

  if i = 8 then
    if FCZKEM.GetDeviceStatus(FDevID, 22, Value) then
      AddDeviceStatus('%s: %d', [StatusNames[14], Value])
    else
      ActualizarError(format('GetDeviceStatus(%d)', [22]));

  if FCZKEM.GetDeviceStatus(FDevID, i, Value) then
    AddDeviceStatus('%s: %d', [StatusNames[i], Value])
  else
    ActualizarError(format('GetDeviceStatus(%d)', [i]));
end; {for}
end;

function TReloj.Respaldar(AArchivo: string): Boolean;
begin
Result := FCZKEM.BackupData(AArchivo);
if Result then
  AddDeviceStatus('Backup Data al archivo: "%s" OK.', [AArchivo])
else
  ActualizarError('CancelOperation Backup');
end;

function TReloj.Restaurar(AArchivo: string): Boolean;
begin
Result := FCZKEM.RestoreData(AArchivo);
if Result then
  AddDeviceStatus('Restore Data from File: "%s" OK.', [AArchivo])
else
  ActualizarError('RestoreData Error');
end;

function TReloj.SaveDayLight: Boolean;
begin
Result := FDayLight.Guardar(DayLight.Soporte, DayLight.Inicio, DayLight.Fin);
end;

function TReloj.SaveDayLight(ADayLight: TDayLight): Boolean;
begin
Result := FDayLight.Guardar(ADayLight.Soporte, ADayLight.Inicio, ADayLight.Fin);
end;

function TReloj.SaveDayLight(ASoporte: Integer; AInicio, AFin: String): Boolean;
begin
Result := FDayLight.Guardar(ASoporte, AInicio, AFin);
end;

procedure TReloj.SaveDeviceInfo;
var
  Value, i: Integer;
begin
for i := 1 to DeviceInfo.Count - 1 do
  if Modified[i] = '1' then begin
    Value := strtoint(DeviceInfo.ValueFromIndex[i - 1]);

    if not FCZKEM.SetDeviceInfo(FDevID, i, Value) then begin
        ActualizarError('Set Device Info');
        exit;
    end else
      Modified[i] := #0;

    if i = 2 then
      FDevID := GetDevID;
  end; {if}

AddDeviceStatus('SetDeviceInfo OK');
end;

procedure TReloj.MostrarError;
begin
MostrarError(Format('Ocurrio un detalle en el reloj %d-%s', [ErrorCode, ErrorMsg]));
end;

procedure TReloj.MostrarError(AMsg: string);
begin
MostrarError(Format('Ocurrio un detalle en el reloj %s: %d-%s', [AMsg, ErrorCode, ErrorMsg]));
end;

procedure TReloj.MostrarError(AMsg: string; Args: array of const);
begin
MostrarError(Format(AMsg, Args));
end;

function TReloj.SubirArchivo(AArchivo: string): Boolean;
begin
Result := FCZKEM.SendFile(DevID, AArchivo);
if Result then
  FCZKEM.RefreshData(FDevID)
else
  ActualizarError('Send File')
end;

procedure TReloj.ActualizarError(AMsg:string);
begin
FCZKEM.GetLastError(FErrorCode);
DeviceStatus.Add(format('! %s ErrorNo.=%d: %s', [AMsg, FErrorCode, ErrorMsg]));
end;

procedure TReloj.ActualizarError(AMsg: string; Args: array of const);
begin
ActualizarError(Format(AMsg, Args));
end;

procedure TReloj.AddDeviceStatus(AMsg: String; Args: array of const);
begin
AddDeviceStatus(Format(AMsg, Args));
end;

procedure TReloj.AddDeviceStatus(AMsg: String);
begin
FDeviceStatus.Add(AMsg);
end;

function TReloj.TomarImagen: Boolean;
var
  w, h: Integer;
  image: array of byte;
begin
  setlength(image, 1024 * 512);
  w := 80;

  Result := FCZKEM.CaptureImage(TRUE, w, h, image[0], 'D:\fp.bmp');
  if Result then
    begin
      DeviceStatus.Add('Captured an image.');
    end
  else
    ActualizarError('Capture image fail.');
end;

function TReloj.TotalRegistros: Integer;
begin
if FCZKEM.GetDeviceStatus(FDevID, 6, Result) then
  AddDeviceStatus('%s: %d', [StatusNames[6], Result])
else begin
  ActualizarError('Total de registros');
end;
end;

function TReloj.TotalUsuarios: Integer;
begin
if FCZKEM.GetDeviceStatus(FDevID, 1, Result) then
  AddDeviceStatus('%s: %d', [StatusNames[1], Result])
else begin
  ActualizarError('Total de usuarios');
end;
end;

{ TDayLight }

function TDayLight.Actualizar: Boolean;
var
  iSoporte:Integer;
  wInicio, wFin:WideString;
begin
FCargado := FOwner.FCZKEM.GetDayLight(FOwner.DevID, iSoporte, wInicio, wFin);

if FCargado then begin
  Soporte := iSoporte;
  Inicio := wInicio;
  Fin := wFin;
end else
  FOwner.ActualizarError('Get Day Light');

Result := FCargado;
end;

constructor TDayLight.Create(AOwner: TReloj);
begin
FOwner := AOwner;
FCargado := False;
end;

function TDayLight.Guardar(ASoporte: Integer; AInicio, AFin: String): Boolean;
begin
//Result := FOwner.FCZKEM.SetDaylight(FOwner.DevID, ASoporte, Ainicio, AFin);
Result := FOwner.FCZKEM.SetDaylight(FOwner.DevID, ASoporte, '', '');

if not Result then
  FOwner.ActualizarError('Set Day Light');
end;

{ TPersonal }

function TPersonal.CapturarHuella(AIDUsuario, ADedo: Integer): Boolean;
begin
FOwner.FCZKEM.CancelOperation;
FOwner.FCZKEM.DelUserTmp(FOwner.DevID, AIDUsuario, ADedo);
Result := FOwner.FCZKEM.StartEnroll(AIDUsuario, ADedo);

if Result then begin
  FOwner.AddDeviceStatus('Huella registrada');
  FOwner.FCZKEM.StartIdentify();
end else
  FOwner.ActualizarError('Capturar huella');
end;

function TPersonal.Cargar(AEnrollNumber: Integer): Boolean;
var
  bEnable: WordBool;
  dwEnroll, dwName, dwPwd, dwCard: widestring;
begin
Init;

dwEnroll := IntToStr(AEnrollNumber);

Result := FOwner.FCZKEM.SSR_GetUserInfo(FOwner.DevID, dwEnroll, dwName, dwPwd, FPrivilegios, bEnable);
if Result then begin
  FOwner.FCZKEM.GetStrCardNumber(dwCard);

  FEnrollNumber := AEnrollNumber;
  FActivo := bEnable;
  FPassword := dwPwd;
  FNombre := dwName;
  FTarjeta := dwCard;
  FOwner.AddDeviceStatus('Se cargo los datos del personal %d %s', [EnrollNumber, FNombre]);
end else
  FOwner.ActualizarError('Error al cargo los datos del personal');
end;

function TPersonal.Cargar: Boolean;
begin
Result := Cargar(EnrollNumber);
end;

constructor TPersonal.Create(AOwner: TReloj);
begin
FOwner := AOwner;

FDataSetCampos := TDataSetCamposPersonal.Create;
end;

destructor TPersonal.Destroy;
begin
DataSetCampos.Free;
inherited;
end;

function TPersonal.Eliminar(AEnrollNumber: Integer): Boolean;
var
  dwEnroll:WideString;
begin
dwEnroll := IntTostr(AEnrollNumber);
Result := FOwner.FCZKEM.SSR_DeleteEnrollData(FOwner.DevID, dwEnroll, 12);
if Result then
  FOwner.FCZKEM.RefreshData(FOwner.DevID);
end;

function TPersonal.EliminarFoto(AEnroll: Integer): Boolean;
var
  sImagen:string;
begin
sImagen := Format('%d.jpg', [AEnroll]);
Result := FOwner.FCZKEM.DeleteUserPhoto(FOwner.DevID, sImagen);
if Result then
  FOwner.AddDeviceStatus('Se elimino la foto de %d', [AEnroll])
else
  FOwner.ActualizarError('Eliminar foto');
end;

function TPersonal.EliminarFoto: Boolean;
begin
Result := FEnrollNumber > 0;
if Result then
  Result := EliminarFoto(FEnrollNumber);
end;

function TPersonal.Guardar: Boolean;
var
  bEnable: WordBool;
  dwEnroll, dwName, dwPwd, dwCard: widestring;
begin
dwEnroll := IntToStr(FEnrollNumber);
dwName := Nombre;
dwPwd := Password;
dwCard := Tarjeta;
bEnable := Activo;

FOwner.FCZKEM.SetStrCardNumber(dwCard);
Result := FOwner.FCZKEM.SSR_SetUserInfo(FOwner.DevID,  dwEnroll, dwName, dwPwd, FPrivilegios, bEnable);
if Result then
  FOwner.FCZKEM.RefreshData(FOwner.DevID);
end;

procedure TPersonal.Init;
begin
FEnrollNumber := 0;
FPrivilegios := 0;
FActivo := False;;
FPassword := '';
FNombre := '';
FTarjeta := '';
end;

function TPersonal.LimpiarPersonal: Boolean;
begin
Result := FOwner.LimpiarDatos(TLimpiarDatos.ldUsuarios);
if Result then
  FOwner.AddDeviceStatus('Limpiar usuarios')
else
  FOwner.ActualizarError('Limpiar usuarios');
end;

procedure TPersonal.ListaEn(ADataset: tDataset);
var
  sCardNumber: widestring;
  dwEnrollNumber: widestring;
  iMachinePrivilege: Integer;
  bEnable: WordBool;
  dwName, dwPwd: widestring;
  iRegistro: integer;
begin
if FOwner.FCZKEM.ReadAllUserID(FOwner.DevID) then begin
  iRegistro := 1;
  while FOwner.FCZKEM.SSR_GetAllUserInfo(FOwner.DevID, dwEnrollNumber, dwName, dwPwd, iMachinePrivilege, bEnable) do begin
    FOwner.FCZKEM.GetStrCardNumber(sCardNumber);

    ADataset.Append;
    ADataset.FieldValues[DataSetCampos.ID] := iRegistro;
    ADataset.FieldValues[DataSetCampos.FCodigo] :=  dwEnrollNumber;
    ADataset.FieldValues[DataSetCampos.FNombre] :=  dwName;
    ADataset.FieldValues[DataSetCampos.Tarjeta] :=  sCardNumber;
    ADataset.Post;

    inc(iRegistro);
  end; {while}

  ADataset.First;
end; {if}

FOwner.AddDeviceStatus('Lista de personal guardada en un Dataset');
end;

function TPersonal.QuitarAdministradores: Boolean;
begin
Result := FOwner.FCZKEM.ClearAdministrators(FOwner.DevID);
if Result then begin
  FOwner.FCZKEM.RefreshData(FOwner.DevID);
  FOwner.AddDeviceStatus('Se quitaron administradores');
end else
  FOwner.ActualizarError('Quitar administradores');
end;

function TPersonal.Respaldar(AArchivo: string): Boolean;
var
  sCardNumber: widestring;
  dwEnrollNumber: widestring;
  iMachinePrivilege: Integer; // dwBackupNumber ,Password, c ,
  bEnable: WordBool;
  dwName, dwPwd: widestring;

  iRegistro: integer;
  strArchivo, strLinea: TStringList;
begin
Result := False;

  strArchivo := TStringList.Create();
  strLinea := TStringList.Create();
  try
    if FOwner.FCZKEM.ReadAllUserID(FOwner.DevID) then
      begin
        iRegistro := 1;
        while FOwner.FCZKEM.SSR_GetAllUserInfo(FOwner.DevID, dwEnrollNumber, dwName, dwPwd, iMachinePrivilege, bEnable) do begin
          FOwner.FCZKEM.GetStrCardNumber(sCardNumber);
          inc(iRegistro);

          strLinea.Clear;
          strLinea.Add(IntToStr(iRegistro));
          strLinea.Add(dwEnrollNumber);
          strLinea.Add(dwName);
          strLinea.Add(dwPwd);
          strLinea.Add(IntToStr(iMachinePrivilege));
          strLinea.Add(sCardNumber);
          strLinea.Add(BoolToStr(bEnable));

          strArchivo.Add(strLinea.CommaText);
        end; {while}

        strArchivo.SaveToFile(AArchivo);
        Result := True;
      end;
  finally
    strArchivo.Free;
    strLinea.Free;
  end; {try}

  FOwner.AddDeviceStatus('Termino el resapldo de usuarios en %s', [AArchivo]);
end;

function TPersonal.Restaurar(AArchivo: string): Boolean;
var
  sCardNumber: widestring;
  dwEnrollNumber: Integer;
  iMachinePrivilege: Integer; // dwBackupNumber ,Password, c ,
  bEnable: WordBool;
  dwName, dwPwd: widestring;

  i, iRegistro: integer;
  strArchivo, strLinea: TStringList;
begin
Result := False;

strArchivo := TStringList.Create();
strLinea := TStringList.Create();
try
  strArchivo.LoadFromFile(AArchivo);

  for i := 0 to pred(strArchivo.Count) do begin
    strLinea.CommaText := strArchivo[i];

    iRegistro := StrToint(strLinea[0]);
    dwEnrollNumber := StrToInt(strLinea[1]);
    dwName := strLinea[2];
    dwPwd := strLinea[3];
    iMachinePrivilege := StrToInt(strLinea[4]);
    sCardNumber := strLinea[5];
    bEnable := StrToBool(strLinea[6]);

    FOwner.FCZKEM.SetStrCardNumber(sCardnumber);
    if FOwner.FCZKEM.SSR_SetUserInfo(FOwner.DevID, IntToStr(dwEnrollNumber), dwName, dwPwd, iMachinePrivilege, bEnable) then
      FOwner.AddDeviceStatus('Se restauro el usuario %d %d: %s', [iRegistro, dwEnrollNumber, dwName])
    else
      FOwner.ActualizarError('Restaurar usuario');
  end; {for}

  FOwner.FCZKEM.RefreshData(FOwner.DevID);
finally
  strArchivo.Free;
  strLinea.Free;
end; {try}

FOwner.AddDeviceStatus('Termino el resapldo de usuarios en %s', [AArchivo]);
end;

function TPersonal.SubirFoto(AImagen: String): Boolean;
begin
Result := FOwner.FCZKEM.SendFile(FOwner.DevID, AImagen);
if Result then begin
  FOwner.FCZKEM.RefreshData(FOwner.DevID);
  FOwner.AddDeviceStatus('Se subio la imagen %s', [AImagen]);
end else
  FOwner.ActualizarError('Subir imagen');
end;

function TPersonal.EliminarTodasFotos: Boolean;
begin
Result := FOwner.FCZKEM.DeleteUserPhoto(FOwner.DevID, 'ALL');
if Result then
  FOwner.AddDeviceStatus('Se eliminaron todas las fotos')
else
  FOwner.ActualizarError('Eliminar todas las fotos');
end;

{ TLog }

procedure TLog.Actualizar;
var
  iSuperLogCount, iDevID, iIDUser, iManipulation, iParams1, iParams2, iParams3, iParams4,
  iYear, iMes, iDia, iHora, iMinuto, iSegundo:Integer;
  strLinea:TStringList;
begin
FBitacora.Clear;
iSuperLogCount := 0;

strLinea := TStringList.Create;
try
  if (FOwner.FCZKEM.ReadSuperLogData(FOwner.DevID)) then begin
    while (FOwner.FCZKEM.GetSuperLogData2(FOwner.DevID, iDevID, iIDUser, iParams4, iParams1, iParams2, iManipulation, iParams3, iYear, iMes, iDia, iHora, iMinuto, iSegundo)) do begin
      iSuperLogCount := iSuperLogCount + 1;

      strLinea.Clear;
      strLinea.Add(IntToStr(iSuperLogCount));
      strLinea.Add(IntToStr(iDevID));
      strLinea.Add(IntToStr(iIDUser));
      strLinea.Add(IntToStr(iParams4));
      strLinea.Add(IntToStr(iParams1));
      strLinea.Add(IntToStr(iParams2));
      strLinea.Add(IntToStr(iManipulation));
      strLinea.Add(IntToStr(iParams3));
      strLinea.Add(Format('%d-%d-%d %d:%d:%s', [iYear, iYear, iMes, iDia, iHora, iMinuto, iSegundo]));

      FBitacora.Add(strLinea.CommaText);
    end; {while}

    FOwner.AddDeviceStatus('Bitacora actualizada');
  end else
    FOwner.ActualizarError('Leer bitacora');
finally
  strLinea.Free;
end; {try}
end;

constructor TLog.Create(AOwner: TReloj);
begin
FOwner := AOwner;
FBitacora := TStringList.Create;
end;

destructor TLog.Destroy;
begin

inherited Destroy;
end;

procedure TLog.GuardarCSV(AArchivo: string);
begin
FBitacora.SaveToFile(AArchivo);
FOwner.AddDeviceStatus('Se guardo bitacora en %s', [AArchivo]);
end;

procedure TLog.LimpiarBitacora;
begin

if (FOwner.FCZKEM.ClearSLog(FOwner.DevID)) then
begin
                FOwner.FCZKEM.RefreshData(FOwner.DevID);
                FOwner.AddDeviceStatus('Se limpio la bitacora');
end else
  FOwner.ActualizarError('Error al limpiar la bitacora');
end;

{ TDataSetCamposPersonal }

constructor TDataSetCamposPersonal.Create;
begin
FID := 'ID';
FCodigo := 'Codigo';
FNombre := 'Nombre';
FTarjeta := 'Tarjeta';
end;

{ TAsistencia }

constructor TAsistencia.Create(AOwner: TReloj);
begin
FOwner := AOwner;
FDataSetCamposAsistencia := TDataSetCamposAsistencia.Create;
FAsistencia := TStringList.Create;
end;

function TAsistencia.Descargar: Boolean;
var
  strLinea:TStringList;
  dwVerifyMode, dwInOutMode,
  dwYear, dwMonth, dwDay, dwHour, dwMinute, dwSecond,
  dwWorkcode: Integer;
  dwEnrollNumber: widestring;
  dFechaHora:TDateTime;
begin
FAsistencia.Clear;

strLinea := TStringList.Create;
try
  if FOwner.FCZKEM.ReadAllGLogData(FOwner.DevID) then
    while FOwner.FCZKEM.SSR_GetGeneralLogData(FOwner.DevID, dwEnrollNumber, dwVerifyMode, dwInOutMode,
    dwYear, dwMonth, dwDay, dwHour, dwMinute, dwSecond, dwWorkcode) do begin
      strLinea.Clear;

      strLinea.Add(IntToStr(FOwner.DevID)); {0}

      strLinea.Add(dwEnrollNumber);  {1}

      strLinea.Add(IntToStr(dwVerifyMode));  {2}
      strLinea.Add(IntToStr(dwInOutMode));  {3}
      strLinea.Add(IntToStr(dwYear)); {4}
      strLinea.Add(IntToStr(dwMonth)); {5}
      strLinea.Add(IntToStr(dwDay)); {6}
      strLinea.Add(IntToStr(dwHour)); {7}
      strLinea.Add(IntToStr(dwMinute)); {8}
      strLinea.Add(IntToStr(dwSecond)); {9}
      dFechaHora := EncodeDateTime(dwYear, dwMonth, dwDay, dwHour, dwMinute, dwSecond, 0);
      strLinea.Add(DateTimeToStr(dFechaHora)); {10}
      strLinea.Add(IntToStr(dwWorkcode)); {11}

      FAsistencia.Add(strLinea.CommaText);
    end; {while}

    Result := True;
finally
  strLinea.Free;
end;
end;

destructor TAsistencia.Destroy;
begin
FAsistencia.Free;
FDataSetCamposAsistencia.Free;
inherited Destroy;
end;

procedure TAsistencia.EnviarStringGrid(AStringGrid: TStringGrid);
var
  i:Integer;
  strLinea:TStringList;
begin
AStringGrid.RowCount := 1;
AStringGrid.ColCount := 1;
AStringGrid.Cells[0, 0] := '';

AStringGrid.RowCount := Self.Total;
AStringGrid.ColCount := 4;

AStringGrid.FixedCols := 1;
AStringGrid.FixedRows := 1;

AStringGrid.Cells[0, 0] := 'Reg';
AStringGrid.ColWidths[0] := 50;
AStringGrid.Cells[1, 0] := 'DevID';
AStringGrid.ColWidths[1] := 50;
AStringGrid.Cells[2, 0] := 'Código';
AStringGrid.ColWidths[2] := 50;
AStringGrid.Cells[3, 0] := 'Fecha / Hora';
AStringGrid.ColWidths[3] := 150;

strLinea := TStringList.Create;
try
  for i := 0 to pred(Self.Total) do begin
    strLinea.CommaText := Asistencia[i];

    AStringGrid.Cells[0, i + 1] := IntToStr(i + 1);
    AStringGrid.Cells[1, i + 1] := strLinea[0];
    AStringGrid.Cells[2, i + 1] := strLinea[1];
    AStringGrid.Cells[3, i + 1] := strLinea[10];
  end; {for}
finally
  strLinea.Free;
end; {try}
end;

function TAsistencia.getTotal: Integer;
begin
Result := FAsistencia.Count;
end;

procedure TAsistencia.GuardarCSV(AArchivo: string);
begin
Asistencia.SaveToFile(AArchivo);
end;

procedure TAsistencia.GuardarDataset(ADataset: TDataSet);
var
  i:Integer;
  strLinea:TStringList;
  dFechaHora:TDateTime;
begin
strLinea := TStringList.Create;
try
  for i := 0 to pred(FAsistencia.Count) do begin
    strLinea.Clear;
    strLinea.CommaText := FAsistencia.Strings[i];

    ADataset.Append;
    ADataset.FieldValues[DataSetCamposAsistencia.DevID] := strLinea[0];
    ADataset.FieldValues[DataSetCamposAsistencia.Codigo] := strLinea[1];
    ADataset.FieldValues[DataSetCamposAsistencia.Verify] := strLinea[2];
    ADataset.FieldValues[DataSetCamposAsistencia.InOut] := strLinea[3];
    ADataset.FieldValues[DataSetCamposAsistencia.Year] := strLinea[4];
    ADataset.FieldValues[DataSetCamposAsistencia.Mes] := strLinea[5];
    ADataset.FieldValues[DataSetCamposAsistencia.Dia] := strLinea[6];
    ADataset.FieldValues[DataSetCamposAsistencia.Hora] := strLinea[7];
    ADataset.FieldValues[DataSetCamposAsistencia.Minutos] := strLinea[8];
    ADataset.FieldValues[DataSetCamposAsistencia.Segundos] := strLinea[9];
    dFechaHora := StrToDateTime(strLinea[10]);
    ADataset.FieldValues[DataSetCamposAsistencia.FechaHora] := dFechaHora;
    ADataset.FieldValues[DataSetCamposAsistencia.WorkCode] := strLinea[11];

    ADataset.Post;
  end; {for}
finally
  strLinea.Free;
end; {try}
end;

function TAsistencia.Limpiar: Boolean;
begin
Result := FOwner.FCZKEM.ClearGLog(FOwner.FDevID);
end;

{ TDataSetCamposAsistencia }

constructor TDataSetCamposAsistencia.Create;
begin
FDevID := 'DevID';
FCodigo := 'Codigo';
FVerify := 'Verify';
FInOut := 'InOut';
FYear := 'Year';
FMinutos := 'Minutos';
FHora := 'Hora';
FMes := 'Mes';
FFechaHora := 'Fecha_Hora';
FDia := 'Dia';
FSegundos := 'Segundos';
FWorkCode := 'WorkCode';
end;

end.
