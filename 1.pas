unit Лаб1;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, VBIDE_TLB, Word_TLB, Office_TLB,
  Vcl.StdCtrls, Vcl.ComCtrls;

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
    procedure RadioButton1Click(Sender: TObject);
    procedure RadioButton2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  Word: WordApplication;
  Pol: String = 'Мужской';

implementation

{$R *.dfm}

procedure TForm1.Button1Click(Sender: TObject);
var
  WordApp: WordApplication; // класс приложения ворда
  Docs: Documents;  // массив документов
  Doc: WordDocument;  // 1 документ
begin
  WordApp := CoWordApplication.Create;  // создаем экземпляр ворда
  WordApp.Visible := true;  // делаем его видимым

  Docs:=WordApp.Documents;
  Doc:=Docs.Add('Normal', False, EmptyParam, True);

  Doc.Paragraphs.Item(1).Alignment := wdAlignParagraphCenter;  // Выравнивание по центру
  Doc.Paragraphs.Item(1).Range.Font.Bold := 1;  // Стиль текста жирный
  Doc.Paragraphs.Item(1).Range.Font.Size := 16; // Размер текста
  Doc.Paragraphs.Item(1).Range.Text :=
  #13 + 'ЗАЯВЛЕНИЕ О ВЫБОРЕ ИНВЕСТИЦИОННОГО ПОРТФЕЛЯ (УПРАВЛЯЮЩЕЙ КОМПАНИИ)'
  + #13
  + #13 + 'Я, ' + Edit1.Text + ' ' + Edit2.Text + ' ' + Edit3.Text
  + #13 + 'Дата рождения:' + #09+#09+#09+#09+#09+#09+#09 + 'Пол: ' + Pol
  + #13 + DateToStr(DateTimePicker1.DateTime)
  + #13 + 'Номер страхового свидетельства обязательного пенсионного страхования:'
  + #13 + Edit4.Text + '-' + Edit5.Text + '-' + Edit6.Text + ' ' +Edit7.Text
  + #13 + 'поручаю пенсионному фонду Российской Федерации направить все средства, учтенные в специальной части моего индивидуального лецивого счета, в управляющую кампанию '
  + #13 + '------------------------------------------------------------------------------------------------------------- '
  + #13 + 'Реквизиты управляющей кампании'
  + #13 + 'ИНН управляющей кампании: ' + Edit8.Text
  + #13 + 'Наименование управляющей кампании: ' + Edit9.Text
  + #13 + '* Наименование инвестиционного портфеля: ' + Edit10.Text
  + #13 + '* (Обязательно для заполнения, если кампания предлагает более одного инвестиционного портфеля)'
  + #13 + '------------------------------------------------------------------------------------------------------------- '
  + #13 + DateToStr(DateTimePicker2.DateTime) + #09+#09+#09+#09+#09+#09+#09+#09+#09 + '___________'
  + #13 + 'Дата заполнения заявления' + #09+#09+#09+#09+#09+#09+#09 + '          Подпись'
  ;

  //Стили Имени/Фамилии/Отчества
  Doc.Paragraphs.Item(4).Alignment := wdAlignParagraphLeft;
  Doc.Paragraphs.Item(4).Range.Font.Size := 14;
  Doc.Paragraphs.Item(4).Range.Font.Bold := 0;

  //Стили Дата рождения/Пол
  Doc.Paragraphs.Item(5).Alignment := wdAlignParagraphLeft;
  Doc.Paragraphs.Item(5).Range.Font.Size := 14;
  Doc.Paragraphs.Item(5).Range.Font.Bold := 0;

  //Стили Выбранной даты/Выбранного пола
  Doc.Paragraphs.Item(6).Alignment := wdAlignParagraphLeft;
  Doc.Paragraphs.Item(6).Range.Font.Size := 14;
  Doc.Paragraphs.Item(6).Range.Font.Bold := 0;

  //Стили пенсионного страхования
  Doc.Paragraphs.Item(7).Alignment := wdAlignParagraphLeft;
  Doc.Paragraphs.Item(7).Range.Font.Size := 14;
  Doc.Paragraphs.Item(7).Range.Font.Bold := 0;
  Doc.Paragraphs.Item(8).Alignment := wdAlignParagraphLeft;
  Doc.Paragraphs.Item(8).Range.Font.Size := 14;
  Doc.Paragraphs.Item(8).Range.Font.Bold := 0;

  //Стили длинного текста
  Doc.Paragraphs.Item(9).Alignment := wdAlignParagraphLeft;
  Doc.Paragraphs.Item(9).Range.Font.Size := 14;
  Doc.Paragraphs.Item(9).Range.Font.Bold := 0;
  Doc.Paragraphs.Item(10).Alignment := wdAlignParagraphLeft;
  Doc.Paragraphs.Item(10).Range.Font.Size := 14;
  Doc.Paragraphs.Item(10).Range.Font.Bold := 0;

  //Стили (Реквизиты управляющей кампании)
  Doc.Paragraphs.Item(11).Alignment := wdAlignParagraphCenter;
  Doc.Paragraphs.Item(11).Range.Font.Size := 14;
  Doc.Paragraphs.Item(11).Range.Font.Bold := 0;
  Doc.Paragraphs.Item(12).Alignment := wdAlignParagraphLeft;
  Doc.Paragraphs.Item(12).Range.Font.Size := 14;
  Doc.Paragraphs.Item(12).Range.Font.Bold := 0;

  //Стили (Наименование управляющей кампании)
  Doc.Paragraphs.Item(13).Alignment := wdAlignParagraphLeft;
  Doc.Paragraphs.Item(13).Range.Font.Size := 14;
  Doc.Paragraphs.Item(13).Range.Font.Bold := 0;

  //Стили (Наименование инвестиционного порфтеля)
  Doc.Paragraphs.Item(14).Alignment := wdAlignParagraphLeft;
  Doc.Paragraphs.Item(14).Range.Font.Size := 14;
  Doc.Paragraphs.Item(14).Range.Font.Bold := 0;

  //Стили текста со звездочкой
  Doc.Paragraphs.Item(15).Alignment := wdAlignParagraphLeft;
  Doc.Paragraphs.Item(15).Range.Font.Size := 10;
  Doc.Paragraphs.Item(15).Range.Font.Bold := 0;
  Doc.Paragraphs.Item(15).Range.Font.Italic := 1;
  Doc.Paragraphs.Item(16).Alignment := wdAlignParagraphLeft;
  Doc.Paragraphs.Item(16).Range.Font.Size := 14;
  Doc.Paragraphs.Item(16).Range.Font.Bold := 0;

  //Стили Даты заполнения
  Doc.Paragraphs.Item(17).Alignment := wdAlignParagraphLeft;
  Doc.Paragraphs.Item(17).Range.Font.Size := 14;
  Doc.Paragraphs.Item(17).Range.Font.Bold := 0;
  Doc.Paragraphs.Item(17).Format.SpaceAfter:=0;
  Doc.Paragraphs.Item(18).Alignment := wdAlignParagraphLeft;
  Doc.Paragraphs.Item(18).Range.Font.Size := 10;
  Doc.Paragraphs.Item(18).Range.Font.Bold := 0;
end;

procedure TForm1.RadioButton1Click(Sender: TObject);
begin
  Pol:= 'Мужской';
end;

procedure TForm1.RadioButton2Click(Sender: TObject);
begin
  Pol:= 'Женский';
end;

end.
