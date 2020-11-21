unit TaskbarProgress;

interface

uses
  Classes, ShlObj;

type
  TTaskbarProgressState = (tbsNormal, tbsError, tbsPaused);

  TTaskbarProgressStyle = (tbstNormal, tbstMarquee);

  TTaskbarProgress = class(TComponent)
  private
    FPosition: Integer;
    FMin: Integer;
    FMax: Integer;
    FState: TTaskbarProgressState;
    FStyle: TTaskbarProgressStyle;
    FEnabled: Boolean;
    FTaskbarList: ITaskbarList3;
    FOwnerHandle: THandle;
    procedure SetEnabled(const AEnabled: Boolean);
    procedure SetState(const AState: TTaskbarProgressState);
    procedure SetStyle(const AStyle: TTaskbarProgressStyle);
    procedure SetPosition(const APosition: Integer);
    procedure SetMax(const AMax: Integer);
    procedure SetMin(const AMin: Integer);
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
  published
    property Position: Integer read FPosition write SetPosition;
    property Min: Integer read FMin write SetMin default 0;
    property Max: Integer read FMax write SetMax default 100;
    property State: TTaskbarProgressState read FState write SetState default tbsNormal;
    property Style: TTaskbarProgressStyle read FStyle write SetStyle default tbstNormal;
    property Enabled: Boolean read FEnabled write SetEnabled default False;
  end;

procedure Register;

implementation

uses
  SysUtils, Forms, Controls, ComObj;

procedure Register;
begin
  Classes.RegisterComponents('Taskbar', [TTaskbarProgress]);
end;

{TTaskbarProgress}

constructor TTaskbarProgress.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);

  FPosition := 0;
  FMin := 0;
  FMax := 100;
  FState := TTaskbarProgressState.tbsNormal;
  FStyle := TTaskbarProgressStyle.tbstNormal;
  FEnabled := False;

  if Assigned(AOwner) and (AOwner is TWinControl) then
    FOwnerHandle := TWinControl(AOwner).Handle
  else if not Application.MainFormOnTaskBar then
    FOwnerHandle := Application.Handle
  else
    FOwnerHandle := Application.MainForm.Handle;

  FTaskbarList := CreateComObject(CLSID_TaskbarList) as ITaskbarList3;
  FTaskbarList.HrInit;
end;

destructor TTaskbarProgress.Destroy;
begin
  FTaskbarList := nil;
  inherited;
end;

procedure TTaskbarProgress.SetEnabled(const AEnabled: Boolean);
begin
  FEnabled := AEnabled;

  if not FEnabled then
  begin
    FTaskbarList.SetProgressState(FOwnerHandle, TBPF_NOPROGRESS);
    Exit;
  end;

  SetState(FState);
  SetStyle(FStyle);
  SetPosition(FPosition);
end;

procedure TTaskbarProgress.SetMax(const AMax: Integer);
begin
  if (AMax < FMin) then
    raise Exception.Create('Invalid max value.');

  FMax := AMax;
end;

procedure TTaskbarProgress.SetMin(const AMin: Integer);
begin
  if (AMin > FMax) then
    raise Exception.Create('Invalid min value.');

  FMin := AMin;
end;

procedure TTaskbarProgress.SetPosition(const APosition: Integer);
begin
  if (APosition < FMin) or (APosition > FMax) then
    raise Exception.Create('Invalid position.');

  FPosition := APosition;

  if not FEnabled then
    Exit;

  if FStyle = tbstMarquee then
    Exit;

  FTaskbarList.SetProgressValue(FOwnerHandle, APosition - FMin, FMax - FMin);
end;

procedure TTaskbarProgress.SetState(const AState: TTaskbarProgressState);
begin
  FState := AState;

  if not FEnabled then
    Exit;

  if FStyle = tbstMarquee then
    Exit;

  case AState of
    tbsNormal:
      FTaskbarList.SetProgressState(FOwnerHandle, TBPF_NORMAL);
    tbsError:
      FTaskbarList.SetProgressState(FOwnerHandle, TBPF_ERROR);
    tbsPaused:
      FTaskbarList.SetProgressState(FOwnerHandle, TBPF_PAUSED);
  end;
end;

procedure TTaskbarProgress.SetStyle(const AStyle: TTaskbarProgressStyle);
begin
  FStyle := AStyle;

  if not FEnabled then
    Exit;

  if FStyle = tbstMarquee then
  begin
    FTaskbarList.SetProgressState(FOwnerHandle, TBPF_INDETERMINATE);
    Exit;
  end;

  SetState(FState);
  SetPosition(FPosition);
end;

end.
