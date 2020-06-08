#include <vcl.h>
#pragma hdrstop
//#include "excel_2k.h"
#include "excel.h"
#include "tools.h"
#include "office_2k.h"
// константы взяты из excel_2k.h и office_2k.h

#define xlSolid 1
#define xlEdgeTop 8
#define xlEdgeBottom 9
#define xlEdgeLeft 7
#define xlEdgeRight 10
#define xlInsideVertical 11
#define xlInsideHorizontal 12
#define xlContinuous 1
#define xlThin 2
#define xlDatabase 1
#define xlDataField 4
#define xlSum 0xFFFFEFC3
#define xlPrintNoComments 0xFFFFEFD2
#define xlPaperA4  9
#define xlAutomatic 0xFFFFEFF7
#define xlDownThenOver 1
#define xlDown  0xFFFFEFE7
#define xlToLeft 0xFFFFEFC1
#define xlToRight 0xFFFFEFBF
#define xlUp 0xFFFFEFBE

//#define msoTextOrientationHorizontal 1
//#define  msoLanguageIDUI 2

//------------------------
EXCEL_APP::EXCEL_APP(){
Variant Books;
Variant Sheets;
int i;

SelectedRange.up=0;
SelectedRange.left=0;
SelectedRange.down=0;
SelectedRange.right=0;
App=Variant::CreateObject("Excel.Application");
App.OlePropertySet("DisplayAlerts",false);
App.OlePropertySet("IgnoreRemoteRequests",false);
this->LangID=LanguageInterfaceID();
try {
    Books=App.OlePropertyGet("Workbooks");
    Books.OleFunction("Add");
    Book=App.OlePropertyGet("ActiveWorkbook");
    Sheets=Book.OlePropertyGet("Worksheets");
    nSheet=Sheets.OlePropertyGet("Count");
    if(nSheet>=MAXSHEET){
        throw Exception("EXCEL_APP::EXCEL_APP():\n"
        "Превышено максимальное число листов");
    }
    for(i=0;i<nSheet;i++){
        Sheet[i].Book=Book;
        Sheet[i].Sheet=Book.OlePropertyGet("Worksheets",i+1);
    }
}
catch(...){
    App.OleFunction("Quit");
}
Sheet[0].Sheet.OleFunction("Select");
CurSheet=0;
}

//---------------------------------------
void EXCEL_APP::NextSheet(AnsiString NewSheetName){
// перейти к следующему листу и дать ему имя
if(CurSheet>=nSheet-1){
    this->AddSheet(NewSheetName);
} else {
    CurSheet++;
    Sheet[CurSheet].Sheet.OleFunction("Select");
    RenameCurSheet(NewSheetName);
}
}

//---------------------------------------
void EXCEL_APP::Show(){
App.OlePropertySet("Visible",true);
}

//----------------------------------------
bool EXCEL_APP::GetVisible(){
// excel visible
bool b;
b=(bool)App.OlePropertyGet("Visible");
return b;
}

//------------------------
void EXCEL_APP::Hide(){
App.OlePropertySet("Visible",false);
}

#define NLETTER 26
//---------------------------
AnsiString CellName(int Row,int Col){
AnsiString r;
char s[3];
static char Alf[]="ABCDEFGHIJKLMNOPQRSTUVWXYZ";
int First,Second;

if(Col<=NLETTER){
    s[0]=Alf[Col-1];
    s[1]='\0';
} else {
    First=(Col-1)/NLETTER;
    Second=Col-First*NLETTER-1;
    s[0]=Alf[First-1];
    s[1]=Alf[Second];
    s[2]='\0';
}
r=s;
r+=Row;
return r;
}

//---------------------------
Variant EXCEL_APP::SelectRange(int Row1,int Col1, int Row2, int Col2){
AnsiString RangeCells,s;
Variant Sh,Range,Result;

if(Row2<Row1 || Col2<Col1){
    s="EXCEL_APP::SelectRange()\n"
          "Правый нижний угол диапазона\n"
          "левее или выше левого верхнего угла\n";
    throw Exception(s+" ("+Row1+","+Col1+","+Row2+","+Col2+")");
}
if(Row2==Row1 && Col2==Col1){
    RangeCells=CellName(Row1,Col1);
} else {
    RangeCells=CellName(Row1,Col1)+":"+CellName(Row2,Col2);
}
Sh=Sheet[CurSheet].Sheet;
Range=Sh.OlePropertyGet("Range",RangeCells);
Result=Range.OleFunction("Select");
SelectedRange.up=Row1;
SelectedRange.left=Col1;
SelectedRange.down=Row2;
SelectedRange.right=Col2;
return Result;
}

//-------------------------------------------
EXCEL_APP::~EXCEL_APP(){
App=Unassigned;
}

//--------------------------------------------
void EXCEL_APP::AddSheet(AnsiString Name){
Variant Sheets, NewSheet;

if(nSheet==MAXSHEET){
    throw Exception("EXCEL_APP::AddSheet()\nЧисло листов максимально");
}
Sheets=Book.OlePropertyGet("WorkSheets");
Sheets.OleFunction("Add");
NewSheet=Book.OlePropertyGet("ActiveSheet");
NewSheet.OlePropertySet("Name",Name);
Sheet[nSheet].Sheet=NewSheet;
Sheet[nSheet].Book=Book;
CurSheet=nSheet;
nSheet++;
}

//---------------------------
void EXCEL_APP::DeleteSheet(AnsiString Name){
bool found=false;
int i,j;
AnsiString s;

for(i=0;i<nSheet;i++){
    s=Sheet[i].GetName();
    if(s.AnsiCompare(Name)==0){
        found=true;
        break;
    }
}
if(found){
    Sheet[i].Sheet.OleFunction("Delete");
    for(j=i;j<nSheet-1;j++){
        memcpy(Sheet+j,Sheet+j+1,sizeof(EXCEL_SHEET));
    }
    nSheet--;
}
}

//---------------------------
void EXCEL_APP::DeleteSheets(int First,int Last){
// удалить листы с First по Last
int i,j,N;
AnsiString s;

N=Last;
if(N>nSheet-1){
    N=nSheet-1;
}
for(i=N;i>=First && i>=0;i--){
    Sheet[i].Sheet.OleFunction("Delete");
    for(j=i;j<nSheet-1;j++){
        memcpy(Sheet+j,Sheet+j+1,sizeof(EXCEL_SHEET));
    }
    nSheet--;
}
}

//---------------------------
void EXCEL_APP::SelectSheet(AnsiString Name){
bool found=false;
int i;
AnsiString s;

for(i=0;i<nSheet;i++){
    s=Sheet[i].GetName();
    if(s.AnsiCompare(Name)==0){
        found=true;
        break;
    }
}
if(found){
    Sheet[i].Sheet.OleFunction("Select");
    CurSheet=i;
}
}

//---------------------------
void EXCEL_APP::ClearSheet(AnsiString SheetName){
MessageBox(0,"ClearSheet не написан","",MB_OK);
}

//---------------------------
void EXCEL_APP::ClearAllSheets(){
int i;
Variant sh;
AnsiString Name;

for(i=0;i<nSheet;i++){
    sh=Sheet[i].Sheet;
    Name=sh.OlePropertyGet("Name");
    ClearSheet(Name);
}
}

//---------------------------
void EXCEL_APP::PutVal(int i,int j,AnsiString Val){
// поместить значение Val в ячейку i,j текущего листа
Variant Sh;
Variant Cell;

Sh=Sheet[CurSheet].Sheet;
Cell=Sh.OlePropertyGet("Cells",i,j);
Cell.OlePropertySet("Value",Val);
}

//---------------------------
Variant EXCEL_APP::GetVal(int i, int j){
Variant Sh;
Variant Cell, Result;

Sh=Sheet[CurSheet].Sheet;
Cell=Sh.OlePropertyGet("Cells",i,j);
Result=Cell.OlePropertyGet("Value");
return Result;
}

//---------------------------
void EXCEL_APP::FindTableBorders(int i,int j,
  int *i1, int *j1, int *i2, int *j2){
// найти границы таблицы excel на листе,стартуя из точки i,j
int m;
Variant p;

for(m=i;m>0;m--){
    p=GetVal(m,j);
    if(p.IsNull() || p.IsEmpty()){
        break;
    }
}
*i1=m+1;

for(m=i;;m++){
    p=GetVal(m,j);
    if(p.IsNull() || p.IsEmpty()){
        break;
    }
}
*i2=m-1;

for(m=j;m>0;m--){
    p=GetVal(i,m);
    if(p.IsNull() || p.IsEmpty()){
        break;
    }
}
*j1=m+1;

for(m=j;;m++){
    p=GetVal(i,m);
    if(p.IsNull() || p.IsEmpty()){
        break;
    }
}
*j2=m-1;
}

//---------------------------
void EXCEL_APP::ColumnsAutoFit(){
Variant Selection;
Variant Columns;

Selection=App.OlePropertyGet("Selection");
Columns=Selection.OlePropertyGet("Columns");
Columns.OleFunction("AutoFit");
}

//---------------------------
void EXCEL_APP::HorAlign(long How){
// Выравнивание
Variant Selection;

Selection=App.OlePropertyGet("Selection");
Selection.OlePropertySet("HorizontalAlignment",How);
}

//--------------------------------------
void EXCEL_APP::VertAlign(long How){
// Выравнивание
Variant Selection;

Selection=App.OlePropertyGet("Selection");
Selection.OlePropertySet("VerticalAlignment",How);
}

//---------------------------
void EXCEL_APP::RangeColor(int Color){
Variant Selection,Interior;

Selection=App.OlePropertyGet("Selection");
Interior=Selection.OlePropertyGet("Interior");
Interior.OlePropertySet("ColorIndex",Color);
Interior.OlePropertySet("Pattern",xlSolid);
}

//---------------------------
void EXCEL_APP::FontBold(bool b){
Variant Selection;
Variant Font;

Selection=App.OlePropertyGet("Selection");
Font=Selection.OlePropertyGet("Font");
Font.OlePropertySet("Bold",b);
}

//---------------------------
void EXCEL_APP::FontItalic(bool b){
Variant Selection;
Variant Font;

Selection=App.OlePropertyGet("Selection");
Font=Selection.OlePropertyGet("Font");
Font.OlePropertySet("Italic",b);
}

//---------------------------
void EXCEL_APP::PutSetka(){
Variant Selection;
Variant Border;
int i,b[]={xlEdgeTop,xlEdgeBottom,xlEdgeLeft,xlEdgeRight,
  xlInsideVertical,xlInsideHorizontal};

Selection=App.OlePropertyGet("Selection");
for(i=0;i<sizeof(b)/sizeof(int);i++){
    if(SelectedRange.up==SelectedRange.down && b[i]==xlInsideHorizontal){
        continue;
    }
    if(SelectedRange.left==SelectedRange.right && b[i]==xlInsideVertical){
        continue;
    }
    Border=Selection.OlePropertyGet("Borders",b[i]);
    Border.OlePropertySet("LineStyle",xlContinuous);
    Border.OlePropertySet("Weight",xlThin);
}
}

//---------------------------
void EXCEL_APP::NumberFormat(AnsiString Format){
Variant Selection;

//this->Show();
Selection=App.OlePropertyGet("Selection");
Selection.OlePropertySet("NumberFormat",Format);
}

//---------------------------
void EXCEL_APP::Freeze(bool b){
MessageBox(0,"EXCEL_APP::Freeze не написан","",MB_OK);
}

//---------------------------
void EXCEL_APP::FontNameAndSize(AnsiString nam, int siz){
Variant Selection,Font;

Selection=App.OlePropertyGet("Selection");
Font=Selection.OlePropertyGet("Font");
Font.OlePropertySet("Name",nam);
Font.OlePropertySet("Size",siz);
}

//---------------------------
void EXCEL_APP::RowsAutoFit(){
Variant Selection;
Variant Rows;

Selection=App.OlePropertyGet("Selection");
Rows=Selection.OlePropertyGet("Rows");
Rows.OleFunction("AutoFit");
}

//---------------------------
void EXCEL_APP::WrapText(bool b){
Variant Selection;

Selection=App.OlePropertyGet("Selection");
Selection.OlePropertySet("WrapText",b);
}

//---------------------------
void EXCEL_APP::FontSize(int n){
Variant Selection,Font;

Selection=App.OlePropertyGet("Selection");
Font=Selection.OlePropertyGet("Font");
Font.OlePropertySet("Size",n);
}

//---------------------------
AnsiString EXCEL_SHEET::GetName(){
AnsiString s;
s=Sheet.OlePropertyGet("Name");
return s;
}

//----------------------------
void EXCEL_SHEET::SetName(AnsiString Name){
// имя листа
Sheet.OlePropertySet("Name",Name);
}

//-----------------------------------------
void EXCEL_APP::Quit(){
App.OleFunction("Quit");
}

//------------------------------------------
void EXCEL_APP::PivotTable(
  AnsiString DataSheet, // имя листа с исх данными
  AnsiString PvtSheet, // имя листа со сводной
  AnsiString PvtTableName, // имя сводной таблицы
  int i1,int j1,int i2, int j2, // диапазон с исх данными на листе DataSheet
  int is, int js, // начальная ячейка сводной
  char *RowFields[], // имена полей, в заголовках строк
  char *ColFields[], // имена полей-заголовков столбцов
  char *PageFields[], // имена полей в страницу
  char *DataFields[] // поля данных
  // все ...Fields заканчиваются NULL
){
Variant Sheet;
Variant Pvt,PvtField;
AnsiString DataRange;
int i,nRow,nCol,nPage;
Variant *vr=0,*vc=0,*vp=0,VR,VC,VP;

this->SelectSheet(DataSheet);
this->SelectRange(i1,j1,i2,j2);
DataRange=AnsiString("R")+i1+"C"+j1+":R"+i2+"C"+j2;
this->SelectSheet(PvtSheet);
Sheet=this->Sheet[CurSheet].Sheet;
this->SelectRange(is,js,is,js);
Pvt=Sheet.OleFunction("PivotTableWizard",xlDatabase,
  DataSheet+"!"+DataRange, Null, PvtTableName);

for(i=0;DataFields[i]!=NULL;i++){
    PvtField=Pvt.OlePropertyGet("PivotFields",
      AnsiString(DataFields[i]));
    PvtField.OlePropertySet("Orientation",xlDataField);
    PvtField.OlePropertySet("Function",xlSum);
}

nRow=0;
while(RowFields[nRow]!=NULL){
    nRow++;
}
if(nRow>0){
    vr=new Variant[nRow];
    for(i=0;i<nRow;i++){
        vr[i]=AnsiString(RowFields[i]);
    }
    VR=VarArrayOf(vr,nRow-1);
} else {
    VR=Null;
}

nCol=0;
while(ColFields[nCol]!=NULL){
    nCol++;
}
if(nCol>0){
    vc=new Variant[nCol];
    for(i=0;i<nCol;i++){
        vc[i]=AnsiString(ColFields[i]);
    }
    VC=VarArrayOf(vc,nCol-1);
} else {
    VC=Null;
}

nPage=0;
while(PageFields[nPage]!=NULL){
    nPage++;
}
if(nPage>0){
    vp=new Variant[nPage];
    for(i=0;i<nPage;i++){
        vp[i]=AnsiString(PageFields[i]);
    }
    VP=VarArrayOf(vp,nPage-1);
} else {
    VP=Null;
}
Pvt.OleFunction("AddFields",VR,VC,VP,false);//по горизонтали
if(vr!=0){
    delete [] vr;
}
if(vc!=0){
    delete [] vc;
}
if(vp!=0){
    delete [] vp;
}
Variant CommandBar=this->App.OlePropertyGet("CommandBars",
  AnsiString("PivotTable"));
CommandBar.OlePropertySet("Visible",false);
}

//---------------------------------------------
void EXCEL_APP::TextOrientation(int Gradus){
// Поворот текста на Gradus
Variant Selection=this->App.OlePropertyGet("Selection");
Selection.OlePropertySet("Orientation",Gradus);
}

//----------------------------------------------
void EXCEL_APP::SetPivotTableFieldPosition(
  AnsiString TableName,
  AnsiString FieldName,
  AnsiString FieldValue,
  int Position){
Variant Sheet,pvt,fn,fv;

Sheet=this->App.OlePropertyGet("ActiveSheet");
pvt=Sheet.OlePropertyGet("PivotTables",TableName);
fn=pvt.OlePropertyGet("PivotFields",FieldName);
fv=fn.OlePropertyGet("PivotItems",(WideString)FieldValue);
fv.OlePropertySet("Position",Position);
}

//-----------------------------------------
void EXCEL_APP::PageOrientation(int Orient){
// портрет, альбом (xlPortrait,xlLandscape)
Variant Sheet,PageSetup;

Sheet=this->App.OlePropertyGet("ActiveSheet");
PageSetup=Sheet.OlePropertyGet("PageSetup");
PageSetup.OlePropertySet("Orientation",Orient);
}

//-----------------------------------------
void EXCEL_APP::CenterSheet(){
// Центрировать по горизонтали и вертикали
Variant Sheet,PageSetup;

Sheet=this->App.OlePropertyGet("ActiveSheet");
PageSetup=Sheet.OlePropertyGet("PageSetup");
PageSetup.OlePropertySet("CenterHorizontally",true);
PageSetup.OlePropertySet("CenterVertically",true);
}

//-----------------------------------------
void EXCEL_APP::PagePropertiesSet(){
Variant Sheet,PageSetup;
AnsiString InchesToPoints="InchesToPoints";
Variant LeftMargin,RightMargin,TopMargin,
  BottomMargin,HeaderMargin,FooterMargin;


Sheet=this->App.OlePropertyGet("ActiveSheet");
PageSetup=Sheet.OlePropertyGet("PageSetup");

PageSetup.OlePropertySet("LeftMargin", 35);
PageSetup.OlePropertySet("RightMargin",35);
PageSetup.OlePropertySet("TopMargin",80);
PageSetup.OlePropertySet("BottomMargin",35);
PageSetup.OlePropertySet("HeaderMargin",30);
PageSetup.OlePropertySet("FooterMargin",20);

PageSetup.OlePropertySet("PrintHeadings",false);
PageSetup.OlePropertySet("PrintGridlines",true);
PageSetup.OlePropertySet("PrintComments",xlPrintNoComments);
PageSetup.OlePropertySet("CenterHorizontally",false);
PageSetup.OlePropertySet("CenterVertically",false);
PageSetup.OlePropertySet("Draft",false);
PageSetup.OlePropertySet("PaperSize",xlPaperA4);
PageSetup.OlePropertySet("FirstPageNumber",xlAutomatic);
PageSetup.OlePropertySet("Order",xlDownThenOver);
PageSetup.OlePropertySet("BlackAndWhite",false);
PageSetup.OlePropertySet("Zoom",false);

PageSetup.OlePropertySet("FitToPagesWide", 1);
// на одну строку по горизогтали

PageSetup.OlePropertySet("FitToPagesTall",100);
// на 100 строк по вертикали
}

//--------------------------------------
void EXCEL_APP::StringsToRepeatOnPage(int First, int Last){
// строки, повторяющиеся на каждой странице
Variant Sheet,PageSetup;
AnsiString s;

Sheet=this->App.OlePropertyGet("ActiveSheet");
PageSetup=Sheet.OlePropertyGet("PageSetup");
s=AnsiString("$")+First+":$"+Last;
PageSetup.OlePropertySet("PrintTitleRows",s);
}

//-----------------------------------------
void EXCEL_APP::ColonTitul(int Where, AnsiString Text){
AnsiString s;
Variant Sheet,PageSetup;

Sheet=this->App.OlePropertyGet("ActiveSheet");
PageSetup=Sheet.OlePropertyGet("PageSetup");

switch(Where){
    case LEFT_HEADER:
        s="LeftHeader";
        break;
    case CENTERHEADER:
        s="CenterHeader";
        break;
    case RIGHTHEADER:
        s="RightHeader";
        break;
    case LEFT_FOOTER:
        s="LeftFooter";
        break;
    case CENTER_FOOTER:
        s="CenterFooter";
        break;
    case RIGHT_FOOTER:
        s="RightFooter";
        break;
    default:
        MessageBox(0,
          "EXCEL_APP::ColonTitul():\n"
          "Непредвиденный вид колонтитула","",MB_OK);
          return;
} // switch
PageSetup.OlePropertySet(s,Text);
}

//-------------------------------------------------
void EXCEL_APP::MergeCells(int Row1,int Col1, int Row2, int Col2){
Variant Range,Selection;

if(Row1<1 || Col1<1 ||Row2<1 || Col2<1) return;
if(Row1>Row2 || Col1>Col2) return;
Range=SelectRange(Row1,Col1,Row2,Col2);
Selection=App.OlePropertyGet("Selection");
Selection.OleFunction("Merge");
}

//-------------------------------------------------
void EXCEL_APP::SetColumnWidth(float Width){
Variant Selection;

Selection=App.OlePropertyGet("Selection");
Selection.OlePropertySet("ColumnWidth",Width);
}

//-------------------------------------------------
void EXCEL_APP::SetRowHeight(int Height){
Variant Selection;

Selection=App.OlePropertyGet("Selection");
Selection.OlePropertySet("RowHeight",Height);
}

//-------------------------------------------------
float EXCEL_APP::GetColumnWidth(){
Variant Selection;
float w;

Selection=App.OlePropertyGet("Selection");
w=Selection.OlePropertyGet("ColumnWidth");
return w;
}

//--------------------------------------------------
void EXCEL_APP::SetColor(int Clr){
Variant Selection,Interior;

Selection=App.OlePropertyGet("Selection");
Interior=Selection.OlePropertyGet("Interior");
Interior.OlePropertySet("ColorIndex",Clr);
Interior.OlePropertySet("Pattern",xlSolid);
}

//----------------------------------------------------
void EXCEL_APP::DeleteRows(int Start,int End){
Variant Selection;

SelectRange(Start,1,End, 200);
Selection=App.OlePropertyGet("Selection");
Selection.OleFunction("Delete",xlUp);
}

//----------------------------------------------------
void EXCEL_APP::InsertRow(int Row){
Variant Selection;

SelectRange(Row,1,Row, 200);
Selection=App.OlePropertyGet("Selection");
Selection.OleFunction("Insert",xlDown);
}

//-----------------------------------------------------
void EXCEL_APP::MoneyFormat(){
Variant Selection;
Variant Style;

Selection=App.OlePropertyGet("Selection");
Style=Selection.OlePropertyGet("Style");
Style.OlePropertySet("NumberFormat","Currency");
}

//------------------------------------------------------
void EXCEL_APP::SetFormula(int Row,int Col,AnsiString Formula){
Variant ActiveCell;

SelectRange(Row,Col,Row,Col);
ActiveCell=App.OlePropertyGet("ActiveCell");
ActiveCell.OlePropertySet("FormulaR1C1",Formula);
}

//-------------------------------------------------------
void EXCEL_APP::MergeLabels(AnsiString PivotTableName,bool Merge){
// объединять ячейки заголовков
Variant ActiveSheet,PivotTable;

ActiveSheet=App.OlePropertyGet("ActiveSheet");
PivotTable=ActiveSheet.OlePropertyGet("PivotTables",PivotTableName);
PivotTable.OlePropertySet("MergeLabels",Merge);
}

//-------------------------------------------------------
void EXCEL_APP::PivotAutoFormat(AnsiString PivotTableName,
  bool Auto){
// иметь Автоформат для сводной
Variant ActiveSheet,PivotTable;

ActiveSheet=App.OlePropertyGet("ActiveSheet");
PivotTable=ActiveSheet.OlePropertyGet("PivotTables",PivotTableName);
PivotTable.OlePropertySet("HasAutoFormat",Auto);
}

#define NYES 12
//---------------------------------------------------------
void EXCEL_APP::PivotSubTotals(
  AnsiString PivotTableName,AnsiString FieldName, bool Yes){
Variant ActiveSheet,PivotTable,PivotField;
Variant VR,vr[NYES];
int i;

ActiveSheet=App.OlePropertyGet("ActiveSheet");
PivotTable=ActiveSheet.OlePropertyGet("PivotTables",PivotTableName);
PivotField=PivotTable.OlePropertyGet("PivotFields",FieldName);
for(i=0;i<NYES;i++){
    vr[i]=Yes;
}
VR=VarArrayOf(vr,NYES-1);
PivotField.OlePropertySet("Subtotals",VR);
}

//---------------------------------------------
void EXCEL_APP::ShowCommandBar(AnsiString BarName, bool Show){
Variant CommandBar;

CommandBar=App.OlePropertyGet("CommandBars","PivotTable");
CommandBar.OlePropertySet("Visible",Show);
}

#define COMPUTERNAMELEN 50
//-----------------------------------------------
void EXCEL_APP::QuerySQLServer(TDatabase *DB,AnsiString Sql,
  int iRow, int iCol){
// разместить результат запроса на листе с ячейки iRow,iCol
AnsiString Connection;
Variant ActiveSheet,QueryTables;
Variant Destination,Query;
AnsiString DatabaseName,ServerName;
char ComputerName[COMPUTERNAMELEN];
unsigned long lcn=COMPUTERNAMELEN-1;
AnsiString s;
AnsiString Password,UserName;

if(!GetComputerName(ComputerName,&lcn)){
    s="Не найдено имя компьютера\n";
    s=s+"Код ошибки="+GetLastError();
    MessageBox(0,s.c_str(),"",MB_OK);
}
if(!DB->AliasName.IsEmpty()){
    TSession *ts=DB->Session;
    TStringList *sl=new TStringList();
    ts->GetAliasParams(DB->AliasName,sl);
    ServerName=FindParamValue(sl, "SERVER NAME");
    DatabaseName=FindParamValue(sl, "DATABASE NAME");
} else {
    ServerName=DB->Params->Values["SERVER NAME"];
    DatabaseName=DB->Params->Values["DATABASE NAME"];
}
UserName=DB->Params->Values["USER NAME"];
Password=DB->Params->Values["PASSWORD"];
ActiveSheet=App.OlePropertyGet("ActiveSheet");
QueryTables=ActiveSheet.OlePropertyGet("QueryTables");
Connection="ODBC;DRIVER=SQL Server;Database="+DatabaseName+";"+
  "SERVER="+ServerName+";UID="+UserName+";"
  "PWD="+Password+";WSID="+ComputerName;
Destination=ActiveSheet.OlePropertyGet("Range",CellName(iRow,iCol));

Query=QueryTables.OleFunction("Add",Connection,Destination,Sql);
Query.OlePropertySet("BackgroundQuery",false);
Query.OleFunction("Refresh");
}

//------------------------------------------
void EXCEL_APP::PutLabel(float x,float y,
  float Width,float Height,AnsiString Text){
Variant ActiveSheet,Labels,Label;

ActiveSheet=App.OlePropertyGet("ActiveSheet");
Labels=ActiveSheet.OlePropertyGet("Labels");
Label=Labels.OleFunction("Add",x,y,Width,Height);
Label.OlePropertySet("Text",Text);
}

//----------------------------------------------
void EXCEL_APP::PutLabelElement(float x,float y,
  float Width,float Height,AnsiString Text){
Variant ActiveSheet,Label,OLEObjects,Obj,Font;

ActiveSheet=App.OlePropertyGet("ActiveSheet");
OLEObjects=ActiveSheet.OlePropertyGet("OLEObjects");
Label=OLEObjects.OleFunction("Add","Forms.Label.1");
Label.OlePropertySet("Left",x);
Label.OlePropertySet("Top",y);
Label.OlePropertySet("Width",Width);
Label.OlePropertySet("Height",Height);
Obj=Label.OlePropertyGet("Object");
Obj.OlePropertySet("AutoSize",false);
Obj.OlePropertySet("TextAlign",2);
Obj.OlePropertySet("Caption",Text);
Font=Obj.OlePropertyGet("Font");
Font.OlePropertySet("Size",10);
Font.OlePropertySet("Bold",false);
}

//---------------------------------------------------
void EXCEL_APP::PutLabelDraw(float x,float y,
  float Width,float Height,AnsiString Text, int Align,
  bool Border){
// Надпись из панели рисования
Variant ActiveSheet,Shapes,Label,Characters,Selection,TextFrame;
Variant Line;

ActiveSheet=App.OlePropertyGet("ActiveSheet");
Shapes=ActiveSheet.OlePropertyGet("Shapes");
Label=Shapes.OleFunction("AddTextBox",msoTextOrientationHorizontal,
  x,y,Width,Height);
TextFrame=Label.OlePropertyGet("TextFrame");
Characters=TextFrame.OleFunction("Characters");
Characters.OlePropertySet("Text",Text);
TextFrame.OlePropertySet("HorizontalAlignment",Align);
Line=Label.OlePropertyGet("Line");
Line.OlePropertySet("Visible",Border);
}

//------------------------------------------
void EXCEL_APP::RenameCurSheet(AnsiString Name){
Variant CurSheet;

CurSheet=App.OlePropertyGet("ActiveSheet");
CurSheet.OlePropertySet("Name",Name);
}

//--------------------------------------------
void EXCEL_APP::ShowGridLines(bool Show){
Variant ActiveWindow;
ActiveWindow=App.OlePropertyGet("ActiveWindow");
ActiveWindow.OlePropertySet("DisplayGridlines",Show);
}

//---------------------------
void EXCEL_APP::PutRangeVal(int i1,int j1,
  int i2,int j2, Variant v){
Variant Range;

Variant Sh;
Sh=Sheet[CurSheet].Sheet;
this->SelectRange(i1,j1,i2,j2);
Range=App.OlePropertyGet("Selection");
Range.OlePropertySet("Value",v);
}

//--------------------------------------------------------
int EXCEL_APP::LanguageInterfaceID(){
/*
MSDN:The following example returns the install language, user interface language, and Help language LCIDs in a message box.

MsgBox "The following locale IDs are registered " & _
    "for this application: Install Language - " & _
    Application.LanguageSettings.LanguageID(msoLanguageIDInstall) & _
    " User Interface Language - " & _
    Application.LanguageSettings.LanguageID(msoLanguageIDUI) & _
    " Help Language - " & _
    Application.LanguageSettings.LanguageID(msoLanguageIDHelp)
*/
// язык интерфейса пользователя
Variant LanguageSettings,LanguageID;
LanguageSettings=App.OlePropertyGet("LanguageSettings");
LanguageID=LanguageSettings.OlePropertyGet("LanguageID",msoLanguageIDUI);
return (int)LanguageID;
}
