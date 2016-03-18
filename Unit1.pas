unit Unit1;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, VBIDE_TLB, Word_TLB, Office_TLB,
  Vcl.StdCtrls, Vcl.ComCtrls, Math, WordDoc;

type
  TForm1 = class(TForm)
    Button1: TButton;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Edit1: TEdit;
    Edit2: TEdit;
    Edit3: TEdit;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    DateTimePicker1: TDateTimePicker;
    Label8: TLabel;
    Label9: TLabel;
    RadioButton1: TRadioButton;
    RadioButton2: TRadioButton;
    Label10: TLabel;
    Edit4: TEdit;
    Edit5: TEdit;
    Label11: TLabel;
    Edit6: TEdit;
    Edit7: TEdit;
    Label12: TLabel;
    Label13: TLabel;
    Label15: TLabel;
    Label16: TLabel;
    Label17: TLabel;
    Edit8: TEdit;
    Label18: TLabel;
    Label19: TLabel;
    Label21: TLabel;
    Edit9: TEdit;
    Label20: TLabel;
    Label22: TLabel;
    Edit10: TEdit;
    Label23: TLabel;
    Label24: TLabel;
    DateTimePicker2: TDateTimePicker;
    Label25: TLabel;
    Label26: TLabel;
    Label27: TLabel;
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

procedure TForm1.Button1Click(Sender: TObject);
var
  masStrok: array of string;
begin
  if ((Sender as TEdit).Text = '') then ShowMessage('Заполните все поля ввода данных')
  else
  begin
    SetLength(masStrok, 13);
    masStrok[0] := Edit1.Text;
    masStrok[1] := Edit2.Text;
    masStrok[2] := Edit3.Text;
    masStrok[3] := Edit4.Text;
    masStrok[4] := Edit5.Text;
    masStrok[5] := Edit6.Text;
    masStrok[6] := Edit7.Text;
    masStrok[7] := Edit8.Text;
    masStrok[8] := Edit9.Text;
    masStrok[9] := Edit10.Text;
    masStrok[10] := DateToStr(DateTimePicker1.DateTime);
    masStrok[11] := DateToStr(DateTimePicker2.DateTime);
  //  masStrok[12] := ifthen(RadioButton1.Checked, 'Мужской', 'Женский');
    if (RadioButton1.Checked) then masStrok[12]:='Мужской'
    else if (RadioButton2.Checked) then masStrok[12]:='Женский';

    createDoc(masStrok);
  end;
end;

end.
