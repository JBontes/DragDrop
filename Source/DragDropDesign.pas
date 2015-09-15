unit DragDropDesign;
// TODO : Default event for target components should be OnDrop.
// TODO : Add parent form to Target property editor list.
// -----------------------------------------------------------------------------
// Project:         Drag and Drop Component Suite
// Module:          DragDropDesign
// Description:     Contains design-time support for the drag and drop
//                  components.
// Version:         5.2
// Date:            17-AUG-2010
// Target:          Win32, Delphi 5-2010
// Authors:         Anders Melander, anders@melander.dk, http://melander.dk
// Copyright        © 1997-1999 Angus Johnson & Anders Melander
//                  © 2000-2010 Anders Melander
// -----------------------------------------------------------------------------

interface

{$include DragDrop.inc}

procedure Register;

implementation

uses
  DragDrop,
  DropSource,
  DropTarget,
  DragDropFile,
  DragDropGraphics,
  DragDropContext,
  DragDropHandler,
  DropHandler,
  DragDropInternet,
  DragDropPIDL,
  DragDropText,
  DropComboTarget,
{$ifndef VER14_PLUS}
  DsgnIntf,
{$else}
{$IFNDEF CPUX64}
  DesignIntf,
  DesignEditors,
{$ENDIF}
{$endif}
  Classes;

{$IFNDEF CPUX64}
type
  TDataFormatNameEditor = class(TStringProperty)
  public
    function GetAttributes: TPropertyAttributes; override;
    procedure GetValues(Proc: TGetStrProc); override;
  end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
//
//              Component and Design-time editor registration
//
////////////////////////////////////////////////////////////////////////////////
procedure Register;
begin
  {$IFNDEF CPUX64}
  RegisterPropertyEditor(TypeInfo(string), TDataFormatAdapter, 'DataFormatName',
    TDataFormatNameEditor);
  {$ENDIF}
  RegisterComponents(DragDropComponentPalettePage,
    [TDropEmptySource, TDropEmptyTarget, TDropDummy, TDataFormatAdapter,
    TDropFileTarget, TDropFileSource, TDropBMPTarget, TDropBMPSource,
    TDropMetaFileTarget, TDropImageTarget, TDropURLTarget, TDropURLSource,
    TDropPIDLTarget, TDropPIDLSource, TDropTextTarget, TDropTextSource,
    TDropComboTarget]);
  RegisterComponents(DragDropComponentPalettePage,
    [TDropHandler, TDragDropHandler, TDropContextMenu]);
end;

{$IFNDEF CPUX64}
{ TDataFormatNameEditor }

function TDataFormatNameEditor.GetAttributes: TPropertyAttributes;
begin
  Result := [paValueList, paSortList, paMultiSelect];
end;

procedure TDataFormatNameEditor.GetValues(Proc: TGetStrProc);
var
  i: Integer;
begin
  for i := 0 to TDataFormatClasses.Count-1 do
    Proc(TDataFormatClasses.Formats[i].ClassName);
end;
{$ENDIF}

end.
