program Ырср1;

uses
  Vcl.Forms,
  Ырс1 in 'Ырс1.pas' {Form1};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
