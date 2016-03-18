UNIT WordDoc;

INTERFACE

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, VBIDE_TLB, Word_TLB, Office_TLB,
  Vcl.StdCtrls, Vcl.ComCtrls;

function createDoc(masStrok: array of string): WordDocument;

VAR
  WordApp: WordApplication; // ����� ���������� �����
  Docs: Documents; // ������ ����������
  Doc: WordDocument; // 1 ��������

IMPLEMENTATION

function createDoc(masStrok: array of string): WordDocument;
begin
  WordApp := CoWordApplication.Create; // ������� ��������� �����
  WordApp.Visible := true; // ������ ��� �������

  Docs := WordApp.Documents;
  Doc := Docs.Add('Normal', False, EmptyParam, true);

  Doc.Paragraphs.Item(1).Alignment := wdAlignParagraphCenter;
  // ������������ �� ������
  Doc.Paragraphs.Item(1).Range.Font.Bold := 1; // ����� ������ ������
  Doc.Paragraphs.Item(1).Range.Font.Size := 16; // ������ ������
  Doc.Paragraphs.Item(1).Range.Text := #13 +
    '��������� � ������ ��������������� �������� (����������� ��������)' + #13 +
    #13 + '�, ' + masStrok[0] + ' ' + masStrok[1] + ' ' + masStrok[2] + #13 +
    '���� ��������:' + #09 + #09 + #09 + #09 + #09 + #09 + #09 + '���: ' + masStrok[12] +
    #13 + masStrok[10] + #13 +
    '����� ���������� ������������� ������������� ����������� �����������:' +
    #13 + masStrok[3] + '-' + masStrok[4] + '-' + masStrok[5] + ' ' + masStrok[6] +
    #13 + '������� ����������� ����� ���������� ��������� ��������� ��� ��������, �������� � ����������� ����� ����� ��������������� �������� �����, � ����������� �������� '
    + #13 + '------------------------------------------------------------------------------------------------------------- '
    + #13 + '��������� ����������� ��������' + #13 +
    '��� ����������� ��������: ' + masStrok[7] + #13 +
    '������������ ����������� ��������: ' + masStrok[8] + #13 +
    '* ������������ ��������������� ��������: ' + masStrok[9] + #13 +
    '* (����������� ��� ����������, ���� �������� ���������� ����� ������ ��������������� ��������)'
    + #13 + '------------------------------------------------------------------------------------------------------------- '
    + #13 + masStrok[11] + #09 + #09 + #09 + #09 + #09 +
    #09 + #09 + #09 + #09 + '___________' + #13 + '���� ���������� ���������' +
    #09 + #09 + #09 + #09 + #09 + #09 + #09 + '          �������';

  // ����� �����/�������/��������
  Doc.Paragraphs.Item(4).Alignment := wdAlignParagraphLeft;
  Doc.Paragraphs.Item(4).Range.Font.Size := 14;
  Doc.Paragraphs.Item(4).Range.Font.Bold := 0;

  // ����� ���� ��������/���
  Doc.Paragraphs.Item(5).Alignment := wdAlignParagraphLeft;
  Doc.Paragraphs.Item(5).Range.Font.Size := 14;
  Doc.Paragraphs.Item(5).Range.Font.Bold := 0;

  // ����� ��������� ����/���������� ����
  Doc.Paragraphs.Item(6).Alignment := wdAlignParagraphLeft;
  Doc.Paragraphs.Item(6).Range.Font.Size := 14;
  Doc.Paragraphs.Item(6).Range.Font.Bold := 0;

  // ����� ����������� �����������
  Doc.Paragraphs.Item(7).Alignment := wdAlignParagraphLeft;
  Doc.Paragraphs.Item(7).Range.Font.Size := 14;
  Doc.Paragraphs.Item(7).Range.Font.Bold := 0;
  Doc.Paragraphs.Item(8).Alignment := wdAlignParagraphLeft;
  Doc.Paragraphs.Item(8).Range.Font.Size := 14;
  Doc.Paragraphs.Item(8).Range.Font.Bold := 0;

  // ����� �������� ������
  Doc.Paragraphs.Item(9).Alignment := wdAlignParagraphLeft;
  Doc.Paragraphs.Item(9).Range.Font.Size := 14;
  Doc.Paragraphs.Item(9).Range.Font.Bold := 0;
  Doc.Paragraphs.Item(10).Alignment := wdAlignParagraphLeft;
  Doc.Paragraphs.Item(10).Range.Font.Size := 14;
  Doc.Paragraphs.Item(10).Range.Font.Bold := 0;

  // ����� (��������� ����������� ��������)
  Doc.Paragraphs.Item(11).Alignment := wdAlignParagraphCenter;
  Doc.Paragraphs.Item(11).Range.Font.Size := 14;
  Doc.Paragraphs.Item(11).Range.Font.Bold := 0;
  Doc.Paragraphs.Item(12).Alignment := wdAlignParagraphLeft;
  Doc.Paragraphs.Item(12).Range.Font.Size := 14;
  Doc.Paragraphs.Item(12).Range.Font.Bold := 0;

  // ����� (������������ ����������� ��������)
  Doc.Paragraphs.Item(13).Alignment := wdAlignParagraphLeft;
  Doc.Paragraphs.Item(13).Range.Font.Size := 14;
  Doc.Paragraphs.Item(13).Range.Font.Bold := 0;

  // ����� (������������ ��������������� ��������)
  Doc.Paragraphs.Item(14).Alignment := wdAlignParagraphLeft;
  Doc.Paragraphs.Item(14).Range.Font.Size := 14;
  Doc.Paragraphs.Item(14).Range.Font.Bold := 0;

  // ����� ������ �� ����������
  Doc.Paragraphs.Item(15).Alignment := wdAlignParagraphLeft;
  Doc.Paragraphs.Item(15).Range.Font.Size := 10;
  Doc.Paragraphs.Item(15).Range.Font.Bold := 0;
  Doc.Paragraphs.Item(15).Range.Font.Italic := 1;
  Doc.Paragraphs.Item(16).Alignment := wdAlignParagraphLeft;
  Doc.Paragraphs.Item(16).Range.Font.Size := 14;
  Doc.Paragraphs.Item(16).Range.Font.Bold := 0;

  // ����� ���� ����������
  Doc.Paragraphs.Item(17).Alignment := wdAlignParagraphLeft;
  Doc.Paragraphs.Item(17).Range.Font.Size := 14;
  Doc.Paragraphs.Item(17).Range.Font.Bold := 0;
  Doc.Paragraphs.Item(17).Format.SpaceAfter := 0;
  Doc.Paragraphs.Item(18).Alignment := wdAlignParagraphLeft;
  Doc.Paragraphs.Item(18).Range.Font.Size := 10;
  Doc.Paragraphs.Item(18).Range.Font.Bold := 0;
end;

END.
