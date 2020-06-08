#ifndef _EXCEL_H
#define _EXCEL_H
#include "system.hpp"
#include "dbtables.hpp"
#include "utilcls.h"
//#include "excel_2k.h"

#define LEFT_HEADER 0
#define CENTERHEADER 1
#define RIGHTHEADER 2
#define LEFT_FOOTER 3
#define CENTER_FOOTER 4
#define RIGHT_FOOTER 5

#define mGREY 15
#define mDARKGREY 48

#define mWHITE xlNone

//----------------------------
class EXCEL_SHEET {
public:
    Variant Sheet;
    Variant Book; // �����, � ������� ���������
    void SetName(AnsiString Name);
    AnsiString GetName();
//    ~EXCEL_SHEET();
};

#define MAXSHEET 20
//----------------------------
typedef struct {
    int up,left,down,right;
} RANGE;

//----------------------------
class EXCEL_APP {
private:
    RANGE SelectedRange;
public:
    Variant App;
    Variant Book;  // ����� - ��� ����
    int CurSheet;
    int nSheet;
    EXCEL_SHEET Sheet[MAXSHEET];
    EXCEL_APP();
    int LangID;
    void AddSheet(AnsiString Name);

    void NextSheet(AnsiString NewSheetName);
    // ������� � ���������� ����� � ���� ��� ���

    void DeleteSheet(AnsiString Name);

    void DeleteSheets(int First,int Last);
    // ������� ����� � First �� Last

    void SelectSheet(AnsiString Name);
    void PutVal(int i,int j,AnsiString Val);
    Variant GetVal(int i,int j);
    ~EXCEL_APP();
    void Show();
    void Hide();
    void ClearSheet(AnsiString SheetName);
    void ClearAllSheets();
    Variant SelectRange(int Row1,int Col1, int Row2, int Col2);

    void FindTableBorders(int i,int j,int *i1, int *j1, int *i2, int *j2);
    // ����� ������� ������� excel �� �����,������� �� ����� i,j

    void ColumnsAutoFit();
    void HorAlign(long How);
    void RangeColor(int Color);
    void FontBold(bool b);
    void PutSetka();
    void NumberFormat(AnsiString Format);

    void Freeze(bool b);
    // ����� ������� �� �����
    // ���������� - �����������

    void RowsAutoFit();
    // ����� ������� �� �����
    // ���������� ������ �����

    void WrapText(bool b);
    // ����� ������� �� �����
    // ������� �� ������

    void FontSize(int n);
    // ����� ������� �� �����
    // ������ ������

    void FontNameAndSize(AnsiString nam, int siz);
    // ����� ������� �� �����
    // ��� � ������ font

    bool GetVisible();
    // excel visible

    void Quit();

    // ���������� �������
    void PivotTable(
      AnsiString DataSheet, // ��� ����� � ��� �������
      AnsiString PvtSheet, // ��� ����� �� �������
      AnsiString PvtTableName, // ��� ������� �������
      int i1,int j1,int i2, int j2, // �������� � ��� ������� �� ����� DataSheet
      int is, int js, // ��������� ������ �������
      char *RowFields[], // ����� �����, � ���������� �����
      char *ColFields[], // ����� �����-���������� ��������
      char *PageFields[], // ����� ����� � ��������
      char *DataFields[] // ���� ������
    );

    void TextOrientation(int Gradus);
    void SetPivotTableFieldPosition(AnsiString TableName,
      AnsiString FieldName,AnsiString FieldValue,int Position);

    void PageOrientation(int Orient);
    // �������, ������

    void PagePropertiesSet();

    void StringsToRepeatOnPage(int First, int Last);
    // ������, ������������� �� ������ ��������

    void ColonTitul(int Where, AnsiString Text);

    void MergeCells(int Row1,int Col1, int Row2, int Col2);
    // ���������� ������

    void VertAlign(long How);

    void SetColumnWidth(float Width);

    void SetColor(int Clr);
    // ������� �����

    float GetColumnWidth();
    void DeleteRows(int Start,int End);
    void InsertRow(int Row);
    void FontItalic(bool b);
    void MoneyFormat();
    void SetFormula(int Row,int Col,AnsiString Formula);
    void SetRowHeight(int Height);

    void MergeLabels(AnsiString PivotTableName,bool Merge);
    // ���������� ������ ����������

    void PivotAutoFormat(AnsiString PivotTableName, bool Auto);
    // ����� ���������� ��� �������

    void PivotSubTotals(
      AnsiString PivotTableName,AnsiString FieldName, bool Yes);
    // �����/���������� subtotals
    
    void ShowCommandBar(AnsiString BarName, bool Show);

    void CenterSheet();
    // ������������ �� ����������� � ���������

    void QuerySQLServer(TDatabase *DB,AnsiString Sql,
      int iRow, int iCol);
    // ���������� ��������� ������� �� �����

    void PutLabel(float x,float y,float Width,float Height,
      AnsiString Text);
      
    void PutLabelElement(float x,float y,
      float Width,float Height,AnsiString Text);

    void PutLabelDraw(float x,float y,
      float Width,float Height,AnsiString Text, int Align,
      bool Border);
    // ������� �� ������ ���������

    void RenameCurSheet(AnsiString Name);

    void ShowGridLines(bool Show);
    
    void PutRangeVal(int i1,int j1,int i2,int j2, Variant v);

    int LanguageInterfaceID();
};
AnsiString CellName(int Row,int Col);
#endif

