unit uDadosHardware;

interface

uses
  sharemem,
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls, Spin, magwmi, magsubs1, ImgList;

const
   namespace = 'root\CIMV2' ;
   classSO   = 'Win32_OperatingSystem' ;
   classHD   = 'Win32_DiskDrive' ;
   classBios = 'Win32_BIOS' ;
   classMem  = 'Win32_PhysicalMemory';
   classProc = 'Win32_Processor' ;
   classRede = 'Win32_NetworkAdapterConfiguration' ;
   classDisp = 'Win32_VideoController' ;
   classPlac = 'Win32_ComputerSystem' ;

type
  TForm1 = class(TForm)
    Button1: TButton;
    dh: TTreeView;
    Button2: TButton;
    Label2: TLabel;
    im: TImageList;
    procedure Button1Click(Sender: TObject);
    procedure Inicial;
    procedure DadosSistemaOperacional;
    procedure DadosHD;
    procedure Bios;
    procedure Memoria;
    procedure Processador;
    procedure Rede;
    procedure Display;
    procedure Button2Click(Sender: TObject);

  private
    { Private declarations }
    stWorkGroup : string ;

  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.Button1Click(Sender: TObject);
var
  Raiz: TTreeNode;

begin

    dh.Items.Clear ;

    Inicial ;

    DadosSistemaOperacional ;

    Bios ;

    Processador;

    Display ;

    Memoria ;

    Rede ;

    DadosHD;

end;

procedure TForm1.Button2Click(Sender: TObject);
begin
  close;
end;

procedure TForm1.Inicial;
var
  Raiz            : TTreeNode;
  rows,instances  : integer ;
  WmiResults      : T2DimStrArray ;
  comp, user, pass, errstr : string ;

begin

    try
        rows := MagWmiGetInfoEx( comp, namespace, user, pass, classPlac, WmiResults, instances, errstr) ;

        with dh.Items do begin
                                      //C  L
            Raiz := add( nil, 'Dados Placa M�e' ) ;
            raiz.ImageIndex := 0 ;
            AddChild( Raiz, 'Fabricante..............: ' + WmiResults [ 1, 26] ).ImageIndex := -1;
            AddChild( Raiz, 'Modelo..................: ' + WmiResults [ 1, 27] ).ImageIndex := -1;
            AddChild( Raiz, 'Placa M�e...............: ' + MagWmiGetBaseBoard  ).ImageIndex := -1;

            stWorkGroup := WmiResults [ 1, 59] ;

        end;

    except
        on E:EConvertError do begin

            Raiz := dh.Items.add( nil, 'Dados Placa M�e' ) ;
            raiz.ImageIndex := 0 ;
            dh.Items.AddChild( Raiz, 'Sistem gerou erro na busca das informa��es.' );
            dh.Items.AddChild( Raiz, e.Message );

        end;

    end;

end;

procedure TForm1.DadosSistemaOperacional;
var
  Raiz            : TTreeNode;
  rows,instances  : integer ;
  WmiResults      : T2DimStrArray ;
  comp, user, pass, errstr : string ;

begin

    try
        rows := MagWmiGetInfoEx( comp, namespace, user, pass, classso, WmiResults, instances, errstr) ;

        with dh.Items do begin
                                      //C  L
            Raiz := add( nil, 'Dados Sistema Operacional' ) ;
            raiz.ImageIndex := 1 ;
            AddChild( Raiz, 'Empresa Fabricante do SO: ' + WmiResults [ 1, 29] ).ImageIndex := -1; //Empresa Fabricante do SO
            AddChild( Raiz, 'Sistema Operacional.....: ' + WmiResults [ 1,  4] ).ImageIndex := -1; //Sistema Operacional
            AddChild( Raiz, 'Vers�o do SO............: ' + WmiResults [ 1, 63] ).ImageIndex := -1; //vers�o do windows
            AddChild( Raiz, 'Build de Vers�o.........: ' + WmiResults [ 1,  2] ).ImageIndex := -1; //Build de Vers�o
            AddChild( Raiz, 'N�mero Serial...........: ' + WmiResults [ 1, 51] ).ImageIndex := -1; //N�mero Serial
            AddChild( Raiz, 'Sistema de Arquivo BOOT.: ' + WmiResults [ 1,  1] ).ImageIndex := -1; //Sistema de Arquivo BOOT
            AddChild( Raiz, 'Sistema de Arquivo......: ' + WmiResults [ 1, 57] ).ImageIndex := -1; //Sistema de Arquivo
            AddChild( Raiz, 'Drive do Sistema........: ' + WmiResults [ 1, 59] ).ImageIndex := -1; //Drive do Sistema
            AddChild( Raiz, 'Pasta do Sistema........: ' + WmiResults [ 1, 64] ).ImageIndex := -1; //Pasta do windows
            AddChild( Raiz, 'Diret�rio do Sistema....: ' + WmiResults [ 1, 58] ).ImageIndex := -1; //Diret�rio do Sistema
            AddChild( Raiz, 'Arquitetura.............: ' + WmiResults [ 1, 39] ).ImageIndex := -1; //Arquitetura
            AddChild( Raiz, 'N�vel de Criptografia...: ' + WmiResults [ 1, 19] ).ImageIndex := -1; //N�vel de Criptografia
            AddChild( Raiz, 'L�ngua..................: ' + WmiResults [ 1, 32] ).ImageIndex := -1; //L�ngua

        end;
    except
        on E:EConvertError do begin

            Raiz := dh.Items.add( nil, 'Dados Placa M�e' ) ;
            raiz.ImageIndex := 1 ;
            dh.Items.AddChild( Raiz, 'Sistem gerou erro na busca das informa��es.' );
            dh.Items.AddChild( Raiz, e.Message );

        end;

    end;

end;
procedure TForm1.DadosHD;
var
  Raiz            : TTreeNode;
  rows,instances, C, L  : integer ;
  WmiResults      : T2DimStrArray ;
  comp, user, pass, errstr : string ;

begin

    try
        rows := MagWmiGetInfoEx( comp, namespace, user, pass, classHD, WmiResults, instances, errstr) ;

        with dh.Items do begin

            if rows > 0 then begin

                for C := 1 to instances do begin

                    Raiz := add( nil, 'Dados Hard Disk Device ' + inttostr( C ) + ' - ' + WmiResults [ C, 05] ) ;
                    raiz.ImageIndex := 2 ;

                    try
                      AddChild( Raiz, 'Fabricante..............: ' + WmiResults [ C, 05] ).ImageIndex := -1; //Empresa Fabricante do SO
                      AddChild( Raiz, 'ID F�sico...............: ' + WmiResults [ C, 12] ).ImageIndex := -1; //ID F�sico
                      AddChild( Raiz, 'Firmware................: ' + WmiResults [ C, 16] ).ImageIndex := -1; //Firmware
                      AddChild( Raiz, 'Interface...............: ' + WmiResults [ C, 19] ).ImageIndex := -1; //Interface
                      AddChild( Raiz, 'Modelo..................: ' + WmiResults [ C, 27] ).ImageIndex := -1; //Modelo
                      AddChild( Raiz, 'N�mero de Parti��es.....: ' + WmiResults [ C, 31] ).ImageIndex := -1; //N�mero de Parti��es
                      AddChild( Raiz, 'Conex�o PnP.............: ' + WmiResults [ C, 32] ).ImageIndex := -1; //Conex�o PnP
                      AddChild( Raiz, 'N�mero Serial do HD.....: ' + WmiResults [ C, 40] ).ImageIndex := -1; //Serial F�sico HD
                    except
                      ShowMessage( 'Hardware n�o encontrado. Selecione outro!' );
                    end;

                end;

            end

        end;
    except
        on E:EConvertError do begin

            Raiz := dh.Items.add( nil, 'Dados Placa M�e' ) ;
            raiz.ImageIndex := 2 ;
            dh.Items.AddChild( Raiz, 'Sistem gerou erro na busca das informa��es.' );
            dh.Items.AddChild( Raiz, e.Message );

        end;

    end;

end;
procedure TForm1.Bios;
var
  Raiz            : TTreeNode;
  rows,instances  : integer ;
  WmiResults      : T2DimStrArray ;
  comp, user, pass, errstr : string ;

begin

    try
        rows := MagWmiGetInfoEx( comp, namespace, user, pass, classBios, WmiResults, instances, errstr) ;

        with dh.Items do begin

          try
                                      //C  L
            Raiz := add( nil, 'Dados da BIOS' ) ;
            raiz.ImageIndex := 3 ;
            AddChild( Raiz, 'Fabricante..............: ' + WmiResults [ 1, 13] ).ImageIndex := -1;
            AddChild( Raiz, 'Vers�o..................: ' + WmiResults [ 1, 27] ).ImageIndex := -1;
            AddChild( Raiz, 'Firmware................: ' + WmiResults [ 1, 19] + ' v' + WmiResults [ 1, 20] +'.'+WmiResults [ 1, 21] ).ImageIndex := -1;
            AddChild( Raiz, 'Descri��o...............: ' + WmiResults [ 1, 04] ).ImageIndex := -1;
            AddChild( Raiz, 'Data Release............: ' + WmiResults [ 1, 17] ).ImageIndex := -1;
            AddChild( Raiz, 'N�mero Serial...........: ' + WmiResults [ 1, 18] ).ImageIndex := -1;
          except
            ShowMessage( 'Hardware n�o encontrado. Selecione outro!' );
          end;
        end;
    except
        on E:EConvertError do begin

            Raiz := dh.Items.add( nil, 'Dados Placa M�e' ) ;
            raiz.ImageIndex := 3 ;
            dh.Items.AddChild( Raiz, 'Sistem gerou erro na busca das informa��es.' );
            dh.Items.AddChild( Raiz, e.Message );

        end;

    end;

end;
procedure TForm1.Memoria;
var
  Raiz            : TTreeNode;
  rows,instances  : integer ;
  WmiResults      : T2DimStrArray ;
  comp, user, pass, errstr : string ;

begin

    try
        rows := MagWmiGetInfoEx( comp, namespace, user, pass, classmem, WmiResults, instances, errstr) ;

        with dh.Items do begin

          try
                                      //C  L
            Raiz := add( nil, 'Dados da Mem�ria' ) ;
            raiz.ImageIndex := 4 ;
            AddChild( Raiz, 'Fabricante..............: ' + WmiResults [ 1, 13] ).ImageIndex := -1;
            AddChild( Raiz, 'Dados...................: ' + WmiResults [ 1, 03] ).ImageIndex := -1;
            AddChild( Raiz, 'Banco da Mem�ria........: ' + WmiResults [ 1, 01] ).ImageIndex := -1;
            AddChild( Raiz, 'Localiza��o.............: ' + WmiResults [ 1, 07] ).ImageIndex := -1;
            AddChild( Raiz, 'Capacidade..............: ' + WmiResults [ 1, 02] ).ImageIndex := -1;
            AddChild( Raiz, 'Velocidade..............: ' + WmiResults [ 1, 25] ).ImageIndex := -1;
            AddChild( Raiz, 'Part Number.............: ' + WmiResults [ 1, 18] ).ImageIndex := -1;
            AddChild( Raiz, 'Serial..................: ' + WmiResults [ 1, 23] ).ImageIndex := -1;
          except
            ShowMessage( 'Hardware n�o encontrado. Selecione outro!' );
          end;
        end;
    except
        on E:EConvertError do begin

            Raiz := dh.Items.add( nil, 'Dados Placa M�e' ) ;
            raiz.ImageIndex := 4 ;
            dh.Items.AddChild( Raiz, 'Sistem gerou erro na busca das informa��es.' );
            dh.Items.AddChild( Raiz, e.Message );

        end;

    end;

end;
procedure TForm1.Processador;
var
  Raiz            : TTreeNode;
  rows,instances  : integer ;
  WmiResults      : T2DimStrArray ;
  comp, user, pass, errstr : string ;

begin

    try
        rows := MagWmiGetInfoEx( comp, namespace, user, pass, classProc, WmiResults, instances, errstr) ;

        with dh.Items do begin

          try
                                      //C  L
            Raiz := add( nil, 'Dados da Processador' ) ;
            raiz.ImageIndex := 5 ;
            AddChild( Raiz, 'Fabricante..............: ' + WmiResults [ 1, 26] ).ImageIndex := -1;
            AddChild( Raiz, 'Descri��o...............: ' + WmiResults [ 1, 12] ).ImageIndex := -1;
            AddChild( Raiz, 'Processador.............: ' + WmiResults [ 1, 28] ).ImageIndex := -1;
            AddChild( Raiz, 'Clock (M�ximo)..........: ' + WmiResults [ 1, 27] ).ImageIndex := -1;
            AddChild( Raiz, 'Cache Mem�ria L2........: ' + WmiResults [ 1, 19] ).ImageIndex := -1;
            AddChild( Raiz, 'Cache Mem�ria L3........: ' + WmiResults [ 1, 21] ).ImageIndex := -1;
            AddChild( Raiz, 'N�cleos.................: ' + WmiResults [ 1, 29] ).ImageIndex := -1;
            AddChild( Raiz, 'Processadores L�gicos...: ' + WmiResults [ 1, 30] ).ImageIndex := -1;
            AddChild( Raiz, 'Serial..................: ' + WmiResults [ 1, 35] ).ImageIndex := -1;

          except
            ShowMessage( 'Hardware n�o encontrado. Selecione outro!' );
          end;
        end;
    except
        on E:EConvertError do begin

            Raiz := dh.Items.add( nil, 'Dados Placa M�e' ) ;
            raiz.ImageIndex := 5 ;
            dh.Items.AddChild( Raiz, 'Sistem gerou erro na busca das informa��es.' );
            dh.Items.AddChild( Raiz, e.Message );

        end;

    end;

end;
procedure TForm1.Rede;
var
  Raiz            : TTreeNode;
  rows,instances, C, L  : integer ;
  WmiResults      : T2DimStrArray ;
  comp, user, pass, errstr, desc : string ;

begin

    try
        rows := MagWmiGetInfoEx( comp, namespace, user, pass, classRede, WmiResults, instances, errstr) ;

        with dh.Items do begin

            if rows > 0 then begin

                for C := 1 to instances do begin

                    if WmiResults[C, 44] <> 'NULL' then begin

                        desc := 'Dados Rede - ' + WmiResults [ C, 44] ;

                        if UpperCaseAnsi( WmiResults[C, 28] ) = 'TRUE' then desc := desc + ' - Ativo' ;

                        Raiz := add( nil, desc ) ;
                        raiz.ImageIndex := 6 ;

                        try
                          AddChild( Raiz, 'Fabricante..............: ' + WmiResults [ C, 09] ).ImageIndex := -1;
                          AddChild( Raiz, 'HostName................: ' + WmiResults [ C, 17] ).ImageIndex := -1;
                          AddChild( Raiz, 'Grupo de Trabalho.......: ' + stWorkGroup ).ImageIndex := -1;
                          AddChild( Raiz, 'IP Endere�o.............: ' + WmiResults [ C, 26] ).ImageIndex := -1;
                          AddChild( Raiz, 'Mascara de Rede.........: ' + WmiResults [ C, 34] ).ImageIndex := -1;
                          AddChild( Raiz, 'Gateway.................: ' + WmiResults [ C, 06] ).ImageIndex := -1;
                          AddChild( Raiz, 'DNS Prim�rio............: ' + WmiResults [ C, 18] ).ImageIndex := -1;
                          AddChild( Raiz, 'MAC Adress..............: ' + WmiResults [ C, 44] ).ImageIndex := -1;
                          AddChild( Raiz, 'Ativa...................: ' + WmiResults [ C, 28] ).ImageIndex := -1;

                        except
                          ShowMessage( 'Hardware n�o encontrado. Selecione outro!' );
                        end;

                    end;

                end;

            end

        end;
    except
        on E:EConvertError do begin

            Raiz := dh.Items.add( nil, 'Dados Placa M�e' ) ;
            raiz.ImageIndex := 6 ;
            dh.Items.AddChild( Raiz, 'Sistem gerou erro na busca das informa��es.' );
            dh.Items.AddChild( Raiz, e.Message );

        end;

    end;

end;
procedure TForm1.Display;
var
  Raiz            : TTreeNode;
  rows,instances  : integer ;
  WmiResults      : T2DimStrArray ;
  comp, user, pass, errstr : string ;

begin

    try
        rows := MagWmiGetInfoEx( comp, namespace, user, pass, classDisp, WmiResults, instances, errstr) ;

        with dh.Items do begin

          try
                                      //C  L
            Raiz := add( nil, 'Dados da Placa de V�deo' ) ;
            raiz.ImageIndex := 7 ;
            AddChild( Raiz, 'Fabricante..............: ' + WmiResults [ 1, 02] ).ImageIndex := -1;
            AddChild( Raiz, 'RAM.....................: ' + WmiResults [ 1, 04] ).ImageIndex := -1;
            AddChild( Raiz, 'Descri��o...............: ' + WmiResults [ 1, 20] ).ImageIndex := -1;
            AddChild( Raiz, 'Processador.............: ' + WmiResults [ 1, 59] ).ImageIndex := -1;
            AddChild( Raiz, 'ID Hardware.............: ' + WmiResults [ 1, 43] ).ImageIndex := -1;

          except
            ShowMessage( 'Hardware n�o encontrado. Selecione outro!' );
          end;
        end;
    except
        on E:EConvertError do begin

            Raiz := dh.Items.add( nil, 'Dados Placa M�e' ) ;
            raiz.ImageIndex := 7 ;
            dh.Items.AddChild( Raiz, 'Sistem gerou erro na busca das informa��es.' );
            dh.Items.AddChild( Raiz, e.Message );

        end;

    end;

end;


end.
