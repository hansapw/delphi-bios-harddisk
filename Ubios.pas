unit Ubios;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls;

type
  TForm1 = class(TForm)
    Button1: TButton;
    Edit1: TEdit;
    Edit2: TEdit;
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

uses
      ActiveX,
      ComObj;

    var
      FSWbemLocator : OLEVariant;
      FWMIService   : OLEVariant;

    function  GetWMIstring(const WMIClass, WMIProperty:string): string;
    const
      wbemFlagForwardOnly = $00000020;
    var
      FWbemObjectSet: OLEVariant;
      FWbemObject   : OLEVariant;
      oEnum         : IEnumvariant;
      iValue        : LongWord;
    begin;
      Result:='';
      FWbemObjectSet:= FWMIService.ExecQuery(Format('Select %s from %s',[WMIProperty, WMIClass]),'WQL',wbemFlagForwardOnly);
      oEnum         := IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
      if oEnum.Next(1, FWbemObject, iValue) = 0 then

  if not VarIsNull(FWbemObject.Properties_.Item(WMIProperty).Value) then

     Result:=FWbemObject.Properties_.Item(WMIProperty).Value;

    FWbemObject:=Unassigned;
    end;

    procedure TForm1.Button1Click(Sender: TObject);
    begin
      FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
      FWMIService   := FSWbemLocator.ConnectServer('localhost', 'root\CIMV2', '', '');

      edit1.Text:=GetWMIstring('Win32_BIOS','SerialNumber');
      edit2.Text:=GetWMIstring('Win32_PhysicalMedia','SerialNumber');

    end;


end.
